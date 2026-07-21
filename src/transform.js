// ---- Funcoes puras (sem rede), pra dar pra testar isolado. ----

// Normaliza nome de cabecalho pra comparar: trim, upper, espacos colapsados.
function normHeader(v) {
  return String(v ?? "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}

// Extrai a data do NOME do arquivo. Aceita "2026.06.25 - BASE SIRTEC_TECNICA".
// Retorna numero AAAAMMDD (ordenavel) ou null se nao achar data.
function parseFileDate(name) {
  const m = String(name || "").match(/(\d{4})[._-](\d{2})[._-](\d{2})/);
  if (!m) return null;
  const [, y, mo, d] = m;
  const n = Number(y) * 10000 + Number(mo) * 100 + Number(d);
  // sanidade basica
  if (Number(mo) < 1 || Number(mo) > 12 || Number(d) < 1 || Number(d) > 31) return null;
  return n;
}

// Monta o mapa coluna-referencia -> indice no cabecalho do arquivo.
// refHeader: array de nomes (ordem alvo). fileHeader: array de nomes do arquivo.
// Retorna { indices: [...], missing: [...] }.
// - Para cada nome da referencia, acha a coluna correspondente no arquivo (por nome).
// - Colunas EXTRAS do arquivo (que nao estao na referencia) sao ignoradas.
// - Se faltar qualquer nome da referencia, vai pra `missing` (o chamador decide pular).
function buildColumnMap(refHeader, fileHeader) {
  const idxByName = new Map();
  fileHeader.forEach((h, i) => {
    const k = normHeader(h);
    if (k !== "" && !idxByName.has(k)) idxByName.set(k, i); // 1a ocorrencia vence
  });
  const indices = [];
  const missing = [];
  for (const ref of refHeader) {
    const k = normHeader(ref);
    if (idxByName.has(k)) indices.push(idxByName.get(k));
    else {
      indices.push(-1);
      missing.push(ref);
    }
  }
  return { indices, missing };
}

// Reordena uma linha do arquivo pra ordem da referencia, usando o mapa de indices.
// Indice -1 (coluna ausente) vira "". Largura final = refHeader.length.
function reorderRow(row, indices) {
  return indices.map((i) => (i < 0 ? "" : row[i] ?? ""));
}

// Compila varios arquivos ja reordenados (cada um na largura da referencia) num unico
// conjunto, deduplicando por TRABALHO (keyIndex). Vence o arquivo MAIS RECENTE.
//
// files: array de { date:Number|null, name:String, rows:Array<Array> } JA REORDENADAS.
//        Pode vir em qualquer ordem; ordenamos por date asc aqui (mais antigo -> mais novo).
// keyIndex: indice da coluna TRABALHO dentro da largura de referencia.
//
// Regra: processa do mais antigo pro mais novo, gravando num Map por chave (TRABALHO).
//   set() mais novo SOBRESCREVE o antigo (prevalece o recente) mantendo a posicao
//   da 1a aparicao (ordem estavel). Linhas com TRABALHO vazio nao deduplicam: sao
//   mantidas todas, anexadas no fim.
function compile(files, keyIndex) {
  const sorted = [...files].sort((a, b) => (a.date ?? -1) - (b.date ?? -1));
  // Data do arquivo mais recente = "verdade" do cadastro atual (snapshot completo).
  // TRABALHO cuja ultima aparicao NAO e nesse arquivo => saiu do cadastro.
  const maxDate = sorted.reduce((m, f) => (f.date != null && f.date > m ? f.date : m), -1);
  const byKey = new Map(); // key -> { row, date } (date da ultima aparicao)
  const emptyKey = []; // { row, date }
  let total = 0;
  for (const f of sorted) {
    for (const row of f.rows) {
      total++;
      const key = normHeader(row[keyIndex]);
      if (key === "") emptyKey.push({ row, date: f.date });
      else byKey.set(key, { row, date: f.date }); // mais novo sobrescreve
    }
  }
  const entries = [...byKey.values(), ...emptyKey];
  const out = entries.map((e) => e.row);
  // Status alinhado 1:1 com out. ativo = apareceu no arquivo mais recente.
  const status = entries.map((e) => ({
    ativo: e.date === maxDate ? "SIM" : "NAO",
    ultimaBase: e.date, // AAAAMMDD (ou null)
  }));
  const inactive = status.filter((s) => s.ativo === "NAO").length;
  return {
    rows: out,
    status,
    maxDate,
    inactive,
    totalRead: total,
    uniqueKeys: byKey.size,
    emptyKeyRows: emptyKey.length,
  };
}

// Formata AAAAMMDD (numero) -> "dd/mm/aaaa". null/invalido -> "".
function ymdToStr(n) {
  if (n == null || n < 0) return "";
  const s = String(n).padStart(8, "0");
  return `${s.slice(6, 8)}/${s.slice(4, 6)}/${s.slice(0, 4)}`;
}

// Formata uma data (Date do SheetJS) como string pt-BR "dd/mm/aaaa" (ou com hora se houver).
// IMPORTANTE: usa getters LOCAIS. O SheetJS monta a Date em horario local; os getters UTC
// trazem o offset do fuso (ex 03:00:28) e quebrariam o dia. Arredonda pro minuto pra matar
// o artefato de fracao de segundo do serial do Excel (ex 00:00:28 -> 00:00).
function excelDateToStr(d) {
  const r = new Date(Math.round(d.getTime() / 60000) * 60000);
  const dd = String(r.getDate()).padStart(2, "0");
  const mm = String(r.getMonth() + 1).padStart(2, "0");
  const yyyy = r.getFullYear();
  const hh = r.getHours();
  const mi = r.getMinutes();
  const date = `${dd}/${mm}/${yyyy}`;
  return hh === 0 && mi === 0 ? date : `${date} ${String(hh).padStart(2, "0")}:${String(mi).padStart(2, "0")}`;
}

// Converte o valor cru de uma celula do xlsx pro que vai pro Sheets:
// - Date  -> string pt-BR (acima);
// - numero/string/bool -> como esta (numero permanece numero -> Sheets guarda numero).
function cellToValue(v) {
  if (v instanceof Date) return excelDateToStr(v);
  return v;
}

// Garante retangulo de largura fixa (preenche faltantes com "").
function padToWidth(rows, width) {
  return rows.map((r) => {
    if (r.length === width) return r;
    const out = r.slice(0, width);
    while (out.length < width) out.push("");
    return out;
  });
}

module.exports = {
  normHeader,
  parseFileDate,
  buildColumnMap,
  reorderRow,
  compile,
  padToWidth,
  excelDateToStr,
  cellToValue,
  ymdToStr,
};
