const { google } = require("googleapis");
const https = require("https");

// "Premature close" (ERR_STREAM_PREMATURE_CLOSE) no fetch do oauth2/v4/token vem
// de socket keep-alive morto: o gaxios reusa conexao que o googleapis.com ja fechou.
// keepAlive:false abre socket novo a cada request -> nunca reusa socket zumbi.
// google.options({agent}) aplica em TODAS as chamadas, inclusive o fetch do token.
const agent = new https.Agent({ keepAlive: false });
google.options({ agent });

// Auth via service account JSON no env GOOGLE_CREDENTIALS.
// trim() remove BOM que o PowerShell injeta ao gravar o secret.
async function getAuthClient(scopes) {
  const raw = process.env.GOOGLE_CREDENTIALS;
  if (!raw) throw new Error("Faltou env GOOGLE_CREDENTIALS (JSON da service account).");
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(raw.trim()),
    scopes,
  });
  const client = await auth.getClient();
  // O fetch do token usa o transporter PROPRIO do client, nao o google.options.
  // Forca o mesmo agent (keepAlive:false) no gaxios do auth -> mata o premature close.
  if (client.transporter?.defaults) {
    client.transporter.defaults.agent = agent;
  }
  return client;
}

// Retenta fn em erro transiente (429/5xx, reset/timeout de rede) com backoff exponencial + jitter.
async function withRetry(fn, { label = "api", tries = 8, baseMs = 1000 } = {}) {
  let lastErr;
  for (let attempt = 1; attempt <= tries; attempt++) {
    try {
      return await fn();
    } catch (e) {
      lastErr = e;
      const status = Number(e.status ?? e.code ?? e.response?.status);
      // code pode estar no erro ou aninhado em e.cause (caso tipico do fetch do token OAuth).
      const codes = [e.code, e.cause?.code].filter(Boolean);
      const netCodes = ["ECONNRESET", "ETIMEDOUT", "ENOTFOUND", "EAI_AGAIN", "ECONNREFUSED", "ERR_STREAM_PREMATURE_CLOSE", "UND_ERR_SOCKET"];
      // Erros de rede transitorios costumam vir so com mensagem (sem status/code), ex:
      // "Premature close" / "socket hang up" / "fetch failed" no fetch do oauth2/v4/token.
      const msg = `${e.message || ""} ${e.cause?.message || ""}`;
      const transientMsg = /premature close|socket hang up|fetch failed|network socket|econnreset|etimedout|eai_again|terminated|tls|esockettimedout/i.test(msg);
      const transient =
        [429, 500, 502, 503, 504].includes(status) ||
        codes.some((c) => netCodes.includes(c)) ||
        transientMsg;
      if (!transient || attempt === tries) throw e;
      const wait = Math.round(baseMs * 2 ** (attempt - 1) * (1 + Math.random()));
      console.warn(`[retry] ${label}: tentativa ${attempt}/${tries} falhou (${status || e.code}); aguardando ${wait}ms`);
      await new Promise((r) => setTimeout(r, wait));
    }
  }
  throw lastErr;
}

module.exports = { getAuthClient, withRetry };
