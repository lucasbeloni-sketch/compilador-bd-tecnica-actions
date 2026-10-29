// ---- Etapa 2: upsert da Carteira_Planejador (substitui o Apps Script atualizarCarteiraPlanejador). ----
// Funcoes puras (sem rede), pra dar pra testar isolado.
//
// Regra (igual ao script):
//   - fonte = linhas da ultima Tecnica (TRABALHO presente no .xlsx mais recente);
//   - chave = TRABALHO -> coluna keyCol (E) do destino;
//   - TRABALHO ja existe no destino -> sobrescreve so as colunas mapeadas (nunca a chave,
//     nunca o carimbo nem a cidade);
//   - TRABALHO novo -> linha nova apos a ultima chave preenchida, com chave, colunas
//     mapeadas, carimbo "dd/mm/aaaa hh:mm - Entrada na Carteira-Planejador" e cidade;
//   - TRABALHO que sumiu da Tecnica fica intocado no destino.
const { normHeader } = require("./transform");

// "A" -> 1, "BN" -> 66.
function colToNum(col) {
  let n = 0;
  for (const ch of String(col).toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

// 66 -> "BN".
function numToCol(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// Resolve o mapa { from: <nome no cabecalho ref>, to: <coluna destino> } -> indices.
// Coluna da ref inexistente = erro (config quebrada, melhor falhar que gravar torto).
function resolveMap(refHeader, map) {
  const refNorm = refHeader.map(normHeader);
  return map.map((m) => {
    const srcIdx = refNorm.indexOf(normHeader(m.from));
    if (srcIdx < 0) throw new Error(`carteira.map: coluna "${m.from}" nao existe no cabecalho de referencia.`);
    return { from: m.from, srcIdx, destCol: colToNum(m.to) };
  });
}

// destKeys = valores da coluna chave do destino a partir de startRow (array 1-D, cru).
// Retorna:
//   updates: { [destCol]: array alinhado com destKeys (null = nao toca na celula) }
//   newRows: matriz (largura = width) pra colar em A{firstNewRow}
//   firstNewRow, updated, inserted
function planUpsert({ srcRows, refHeader, destKeys, cfg, stamp }) {
  const mapping = resolveMap(refHeader, cfg.map);
  const keyCol = colToNum(cfg.keyCol);
  const keyMap = mapping.find((m) => m.destCol === keyCol);
  if (!keyMap) throw new Error(`carteira.map: precisa mapear a chave pra coluna ${cfg.keyCol}.`);
  const stampCol = cfg.stampCol ? colToNum(cfg.stampCol) : 0;
  const cityCol = cfg.cityCol ? colToNum(cfg.cityCol) : 0;
  const width = Math.max(...mapping.map((m) => m.destCol), stampCol, cityCol);

  // Indexa chaves do destino -> posicao (0-based a partir de startRow).
  const keyPos = new Map();
  let lastFilled = -1;
  destKeys.forEach((v, i) => {
    const k = normHeader(v);
    if (k !== "") {
      keyPos.set(k, i);
      lastFilled = i;
    }
  });

  const updates = {};
  const updCols = mapping.filter((m) => m.destCol !== keyCol);
  for (const m of updCols) updates[m.destCol] = new Array(lastFilled + 1).fill(null);

  const newRows = [];
  const seenNew = new Set();
  const touched = new Set();
  for (const row of srcRows) {
    const k = normHeader(row[keyMap.srcIdx]);
    if (k === "") continue;
    if (keyPos.has(k)) {
      const pos = keyPos.get(k);
      touched.add(pos);
      for (const m of updCols) updates[m.destCol][pos] = row[m.srcIdx] ?? "";
    } else if (!seenNew.has(k)) {
      seenNew.add(k);
      const nr = new Array(width).fill("");
      for (const m of mapping) nr[m.destCol - 1] = row[m.srcIdx] ?? "";
      if (stampCol) nr[stampCol - 1] = stamp;
      if (cityCol) nr[cityCol - 1] = cfg.cityValue ?? "";
      newRows.push(nr);
    }
  }

  return {
    updates,
    newRows,
    width,
    lastFilledRow: cfg.startRow + lastFilled, // linha real (startRow-1 se destino vazio)
    firstNewRow: cfg.startRow + lastFilled + 1,
    updated: touched.size,
    inserted: newRows.length,
  };
}

module.exports = { colToNum, numToCol, resolveMap, planUpsert };
