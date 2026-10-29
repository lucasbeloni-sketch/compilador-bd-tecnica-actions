// ---- Colunas da Carteira_Planejador por NOME de cabecalho (linha headerRow), nao por letra. ----
// Assim inserir/mover colunas na planilha nao faz o robo gravar no lugar errado. Cabecalho
// esperado ausente ou duplicado = erro ANTES de qualquer escrita.
const { withRetry } = require("../lib/google");
const { normHeader } = require("./transform");
const { numToCol } = require("./carteira");

async function lerCabecalho(sheets, spreadsheetId, sheet, row) {
  const r = await withRetry(
    () =>
      sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `'${sheet.replace(/'/g, "''")}'!${row}:${row}`,
      }),
    { label: "cabecalho-get" }
  );
  return (r.data.values && r.data.values[0]) || [];
}

// nomes -> { nome: numeroDaColuna (1-based) }. normHeader colapsa quebra de linha/espacos
// e ignora maiusculas ("TOTAL \nPT" == "total pt").
function resolverColunas(header, nomes, ctx = "") {
  const idx = new Map();
  const dup = new Set();
  header.forEach((h, i) => {
    const k = normHeader(h);
    if (!k) return;
    if (idx.has(k)) dup.add(k);
    else idx.set(k, i + 1);
  });
  const faltando = [];
  const duplicados = [];
  const out = {};
  for (const n of nomes) {
    const k = normHeader(n);
    if (dup.has(k)) duplicados.push(n);
    else if (!idx.has(k)) faltando.push(n);
    else out[n] = idx.get(k);
  }
  if (faltando.length || duplicados.length)
    throw new Error(
      `[colunas${ctx ? " " + ctx : ""}] cabecalho nao bate. ` +
        (faltando.length ? `Faltando: ${faltando.join(" | ")}. ` : "") +
        (duplicados.length ? `Duplicado: ${duplicados.join(" | ")}. ` : "") +
        "Nada foi gravado. Ajuste o cabecalho da planilha ou o config.json."
    );
  return out;
}

// [18,19,20,22,37] -> ["R:T", "V", "AK"] (agrupa colunas vizinhas).
function agruparFaixas(nums) {
  const s = [...new Set(nums)].sort((a, b) => a - b);
  const out = [];
  for (let i = 0; i < s.length; ) {
    let j = i;
    while (j + 1 < s.length && s[j + 1] === s[j] + 1) j++;
    out.push(i === j ? numToCol(s[i]) : `${numToCol(s[i])}:${numToCol(s[j])}`);
    i = j + 1;
  }
  return out;
}

module.exports = { lerCabecalho, resolverColunas, agruparFaixas };
