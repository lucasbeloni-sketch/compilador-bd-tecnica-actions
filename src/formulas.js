// ---- Etapa 3: formulas da Carteira_Planejador (substitui os Apps Scripts copiarChecklistGPM,
// Formulas_Carteira_Plan_Nova_Part1/2/3 e classificarSoloRochoso). ----
//
// 1) Copia BD_Checklist_GPM (planilha Bases) pra BD_Checklist_GPM da planilha de Juazeiro,
//    ordenado DECRESCENTE pela coluna sortCol (vazios no fim).
// 2) statusCell = "Em Atualizacao".
// 3) Limpa as faixas da linha startRow ate a ultima chave em keyCol.
// 4) Cola as formulas-modelo da linha templateRow nas faixas (PASTE_FORMULA ajusta referencias
//    relativas, igual ao copyTo do script), le o resultado calculado e grava como VALOR.
// 5) Classifica o solo (AZ/BA -> BM) via src/geologia.js.
// 6) statusCell = data/hora do fim.
const { withRetry } = require("../lib/google");
const { colToNum } = require("./carteira");
const { classificarPar } = require("./geologia");

// Comparador DECRESCENTE do script: vazio sempre no fim; numero com numero; resto como texto.
function compararDecrescente(a, b) {
  const va = a === "" || a === null || a === undefined;
  const vb = b === "" || b === null || b === undefined;
  if (va && vb) return 0;
  if (va) return 1;
  if (vb) return -1;
  if (typeof a === "number" && typeof b === "number") return b - a;
  return String(b).localeCompare(String(a), "pt-BR", { numeric: true });
}

// Retangulo rows x width preenchido com "" (a API corta celulas/linhas vazias do fim).
function padRect(values, rows, width) {
  const out = [];
  for (let i = 0; i < rows; i++) {
    const r = (values[i] || []).slice(0, width);
    while (r.length < width) r.push("");
    out.push(r);
  }
  return out;
}

// "B:D" -> { c1: 2, c2: 4 }; "AK" -> { c1: 37, c2: 37 }.
function parseCols(spec) {
  const [a, b] = String(spec).split(":");
  return { c1: colToNum(a), c2: colToNum(b || a) };
}

// Ultima linha (1-based) com valor numa coluna (array 1-D a partir da linha 1). 0 se vazia.
function lastFilledRow(colValues) {
  for (let i = colValues.length - 1; i >= 0; i--) {
    const v = colValues[i];
    if (v !== "" && v !== null && v !== undefined) return i + 1;
  }
  return 0;
}

const q = (name) => `'${String(name).replace(/'/g, "''")}'!`;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function sheetProps(sheets, spreadsheetId) {
  const meta = await withRetry(
    () =>
      sheets.spreadsheets.get({
        spreadsheetId,
        fields: "sheets.properties(sheetId,title,gridProperties.rowCount)",
      }),
    { label: "formulas-meta" }
  );
  return meta.data.sheets.map((s) => s.properties);
}

async function ensureRows(sheets, spreadsheetId, props, needed) {
  const missing = needed - props.gridProperties.rowCount;
  if (missing <= 0) return;
  await withRetry(
    () =>
      sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [{ appendDimension: { sheetId: props.sheetId, dimension: "ROWS", length: missing } }],
        },
      }),
    { label: "formulas-append-rows" }
  );
  console.log(`[formulas] +${missing} linha(s) na grade de "${props.title}".`);
}

// ---- 1) Checklist ----
async function copiarChecklist(sheets, ck, destId, dryRun) {
  const resp = await withRetry(
    () =>
      sheets.spreadsheets.values.get({
        spreadsheetId: ck.srcSpreadsheetId,
        range: `${q(ck.srcSheet)}A${ck.srcStartRow}:${ck.lastCol}`,
        valueRenderOption: "UNFORMATTED_VALUE",
        dateTimeRenderOption: "SERIAL_NUMBER",
      }),
    { label: "checklist-get" }
  );
  const width = colToNum(ck.lastCol);
  const raw = resp.data.values || [];
  if (!raw.length) {
    console.warn("[checklist] origem vazia; destino mantido.");
    return;
  }
  const rows = padRect(raw, raw.length, width);
  const si = colToNum(ck.sortCol) - 1;
  rows.sort((a, b) => compararDecrescente(a[si], b[si]));
  console.log(`[checklist] ${rows.length} linha(s), ordenadas por ${ck.sortCol} (desc).`);
  if (dryRun) return;

  const props = (await sheetProps(sheets, destId)).find((p) => p.title === ck.destSheet);
  if (!props) throw new Error(`[checklist] aba "${ck.destSheet}" nao encontrada no destino.`);
  await ensureRows(sheets, destId, props, ck.destStartRow + rows.length - 1);

  await withRetry(
    () =>
      sheets.spreadsheets.values.clear({
        spreadsheetId: destId,
        range: `${q(ck.destSheet)}A${ck.destStartRow}:${ck.lastCol}`,
      }),
    { label: "checklist-clear" }
  );
  await withRetry(
    () =>
      sheets.spreadsheets.values.update({
        spreadsheetId: destId,
        range: `${q(ck.destSheet)}A${ck.destStartRow}`,
        valueInputOption: "RAW",
        requestBody: { values: rows },
      }),
    { label: "checklist-write" }
  );

  // RAW grava data como serial; copia o formato de data/hora das colunas da origem
  // (o setValues(Date) do script formatava sozinho).
  const fmt = await withRetry(
    () =>
      sheets.spreadsheets.get({
        spreadsheetId: ck.srcSpreadsheetId,
        ranges: [`${q(ck.srcSheet)}A${ck.srcStartRow}:${ck.lastCol}${ck.srcStartRow}`],
        fields: "sheets.data.rowData.values.userEnteredFormat.numberFormat",
      }),
    { label: "checklist-formats" }
  );
  const cells = fmt.data.sheets[0].data[0].rowData?.[0]?.values || [];
  const requests = [];
  cells.forEach((c, i) => {
    const nf = c.userEnteredFormat?.numberFormat;
    if (!nf || !["DATE", "TIME", "DATE_TIME"].includes(nf.type)) return;
    requests.push({
      repeatCell: {
        range: {
          sheetId: props.sheetId,
          startRowIndex: ck.destStartRow - 1,
          endRowIndex: ck.destStartRow - 1 + rows.length,
          startColumnIndex: i,
          endColumnIndex: i + 1,
        },
        cell: { userEnteredFormat: { numberFormat: nf } },
        fields: "userEnteredFormat.numberFormat",
      },
    });
  });
  if (requests.length)
    await withRetry(
      () => sheets.spreadsheets.batchUpdate({ spreadsheetId: destId, requestBody: { requests } }),
      { label: "checklist-format-write" }
    );
  console.log(`[checklist] gravado em ${ck.destSheet}!A${ck.destStartRow} (${requests.length} coluna(s) com formato de data).`);
}

// ---- 2..6) Formulas + geologia ----
async function atualizarFormulas(sheets, cfg, { dryRun, stampNow }) {
  const fc = cfg.formulas;
  if (!fc || !fc.enabled) return;
  const id = fc.spreadsheetId;
  const a1 = q(fc.sheet);

  if (fc.checklist?.enabled) await copiarChecklist(sheets, fc.checklist, id, dryRun);

  const keyResp = await withRetry(
    () =>
      sheets.spreadsheets.values.get({
        spreadsheetId: id,
        range: `${a1}${fc.keyCol}1:${fc.keyCol}`,
        valueRenderOption: "UNFORMATTED_VALUE",
      }),
    { label: "formulas-get-keys" }
  );
  const lastRow = lastFilledRow((keyResp.data.values || []).map((r) => r[0]));
  if (lastRow < fc.startRow) {
    console.warn(`[formulas] sem obras a partir da linha ${fc.startRow}; nada a fazer.`);
    return;
  }
  const n = lastRow - fc.startRow + 1;
  const ranges = fc.ranges.map((spec) => ({ spec, ...parseCols(spec) }));
  const rangeA1 = (r) => {
    const [a, b] = r.spec.split(":");
    return `${a1}${a}${fc.startRow}:${b || a}${lastRow}`;
  };
  console.log(`[formulas] linhas ${fc.startRow}..${lastRow} (${n} obra(s)), ${ranges.length} faixa(s): ${fc.ranges.join(", ")}.`);

  const geoCfg = fc.geologia;
  if (dryRun) {
    if (geoCfg?.enabled) {
      const g = await withRetry(
        () =>
          sheets.spreadsheets.values.get({
            spreadsheetId: id,
            range: `${a1}${geoCfg.colA}${fc.startRow}:${geoCfg.colB}${lastRow}`,
            valueRenderOption: "UNFORMATTED_VALUE",
          }),
        { label: "geo-get" }
      );
      const cnt = {};
      padRect(g.data.values || [], n, 2).forEach((r) => {
        const v = classificarPar(r[0], r[1], geoCfg.fusoUtm) || "(sem coord)";
        cnt[v] = (cnt[v] || 0) + 1;
      });
      console.log(`[dry-run] geologia: ${JSON.stringify(cnt)}`);
    }
    return;
  }

  const setStatus = (value, input) =>
    withRetry(
      () =>
        sheets.spreadsheets.values.update({
          spreadsheetId: id,
          range: `${a1}${fc.statusCell}`,
          valueInputOption: input,
          requestBody: { values: [[value]] },
        }),
      { label: "formulas-status" }
    );
  await setStatus(fc.statusRunning, "RAW");

  // 3) Limpa.
  await withRetry(
    () =>
      sheets.spreadsheets.values.batchClear({
        spreadsheetId: id,
        requestBody: { ranges: ranges.map(rangeA1) },
      }),
    { label: "formulas-clear" }
  );

  // 4a) Cola as formulas da linha-modelo.
  const props = (await sheetProps(sheets, id)).find((p) => p.title === fc.sheet);
  if (!props) throw new Error(`[formulas] aba "${fc.sheet}" nao encontrada.`);
  const requests = ranges.map((r) => ({
    copyPaste: {
      source: {
        sheetId: props.sheetId,
        startRowIndex: fc.templateRow - 1,
        endRowIndex: fc.templateRow,
        startColumnIndex: r.c1 - 1,
        endColumnIndex: r.c2,
      },
      destination: {
        sheetId: props.sheetId,
        startRowIndex: fc.startRow - 1,
        endRowIndex: lastRow,
        startColumnIndex: r.c1 - 1,
        endColumnIndex: r.c2,
      },
      pasteType: "PASTE_FORMULA",
    },
  }));
  await withRetry(
    () => sheets.spreadsheets.batchUpdate({ spreadsheetId: id, requestBody: { requests } }),
    { label: "formulas-paste" }
  );
  console.log(`[formulas] formulas da linha ${fc.templateRow} coladas.`);

  // 4b) Le o calculado. IMPORTRANGE pode responder "Loading..." logo apos colar: espera e rele.
  const loading = /^(loading|carregando)/i;
  let vrs;
  for (let tent = 1; ; tent++) {
    const got = await withRetry(
      () =>
        sheets.spreadsheets.values.batchGet({
          spreadsheetId: id,
          ranges: ranges.map(rangeA1),
          valueRenderOption: "UNFORMATTED_VALUE",
          dateTimeRenderOption: "SERIAL_NUMBER",
        }),
      { label: "formulas-read" }
    );
    vrs = got.data.valueRanges;
    const pend = vrs.reduce(
      (acc, vr) => acc + (vr.values || []).flat().filter((v) => typeof v === "string" && loading.test(v)).length,
      0
    );
    if (!pend) break;
    if (tent >= (fc.loadingRetries || 12))
      throw new Error(`[formulas] ${pend} celula(s) ainda "Loading..." apos ${tent} leituras. Nao congela.`);
    console.warn(`[formulas] ${pend} celula(s) "Loading..."; relendo (${tent}).`);
    await sleep(fc.loadingWaitMs || 5000);
  }

  // 4c) Congela (grava os valores por cima das formulas). RAW mantem numero/serial como numero;
  // a formatacao da celula fica como esta.
  const data = ranges.map((r, i) => ({
    range: rangeA1(r),
    values: padRect(vrs[i].values || [], n, r.c2 - r.c1 + 1),
  }));
  await withRetry(
    () =>
      sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: id,
        requestBody: { valueInputOption: "RAW", data },
      }),
    { label: "formulas-freeze" }
  );
  console.log(`[formulas] valores congelados (${data.length} faixa(s)).`);

  // 5) Geologia.
  if (geoCfg?.enabled) {
    const g = await withRetry(
      () =>
        sheets.spreadsheets.values.get({
          spreadsheetId: id,
          range: `${a1}${geoCfg.colA}${fc.startRow}:${geoCfg.colB}${lastRow}`,
          valueRenderOption: "UNFORMATTED_VALUE",
        }),
      { label: "geo-get" }
    );
    const out = padRect(g.data.values || [], n, 2).map((r) => [classificarPar(r[0], r[1], geoCfg.fusoUtm)]);
    await withRetry(
      () =>
        sheets.spreadsheets.values.update({
          spreadsheetId: id,
          range: `${a1}${geoCfg.outCol}${fc.startRow}:${geoCfg.outCol}${lastRow}`,
          valueInputOption: "RAW",
          requestBody: { values: out },
        }),
      { label: "geo-write" }
    );
    const semCoord = out.filter((r) => r[0] === "").length;
    console.log(`[geologia] ${n - semCoord} classificada(s), ${semCoord} sem coordenada.`);
  }

  // 6) Fim.
  const fim = stampNow(cfg.timezone);
  await setStatus(fim, cfg.valueInputOption);
  console.log(`[formulas] ${fc.statusCell}: ${fim}`);
}

module.exports = { atualizarFormulas, compararDecrescente, padRect, parseCols, lastFilledRow };
