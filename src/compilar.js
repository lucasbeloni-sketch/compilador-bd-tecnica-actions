const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
const XLSX = require("xlsx");
const { getAuthClient, withRetry } = require("../lib/google");
const { listXlsxInFolder, downloadFile } = require("../lib/drive");
const {
  normHeader,
  parseFileDate,
  buildColumnMap,
  reorderRow,
  compile,
  padToWidth,
  cellToValue,
  ymdToStr,
} = require("./transform");
const { planUpsert, numToCol } = require("./carteira");

// Drive leitura + Sheets escrita.
const SCOPES = [
  "https://www.googleapis.com/auth/drive.readonly",
  "https://www.googleapis.com/auth/spreadsheets",
];

function loadConfig() {
  const raw = fs.readFileSync(path.join(__dirname, "..", "config.json"), "utf8");
  return JSON.parse(raw);
}

// Le um .xlsx (Buffer) -> { header:[], rows:[[...]] } usando a 1a planilha.
// raw:true + cellDates:true devolve os valores CRUS: numero fica numero (Sheets guarda
// como numero), data vira Date. cellToValue normaliza Date -> "dd/mm/aaaa" (pt-BR), pra
// USER_ENTERED parsear certo. (raw:false traria "46,569.21" / "2/25/26" formato US, que
// numa planilha pt-BR viram TEXTO quebrado.)
function readXlsx(buf, headerRow1Based) {
  const wb = XLSX.read(buf, { type: "buffer", cellDates: true });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  if (!ws) return { header: [], rows: [] };
  const matrix = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    raw: true,
    defval: "",
    blankrows: false,
  });
  const hIdx = Math.max(0, (headerRow1Based || 1) - 1);
  const header = (matrix[hIdx] || []).map((h) => String(h ?? ""));
  const rows = matrix.slice(hIdx + 1).map((row) => row.map(cellToValue));
  return { header, rows };
}

// Etapa 2: upsert na Carteira_Planejador (substitui o Apps Script atualizarCarteiraPlanejador).
// srcRows = linhas da ultima Tecnica, ja na ordem do cabecalho de referencia.
async function atualizarCarteira(sheets, cfg, refHeader, srcRows, dryRun) {
  const cc = cfg.carteira;
  if (!cc || !cc.enabled) return;
  const minRows = Number(cc.minRows) || 0;
  if (srcRows.length < minRows)
    throw new Error(`[carteira] so ${srcRows.length} linha(s) na ultima Tecnica (< minRows=${minRows}). Nao mexe na carteira.`);

  const a1 = `'${cc.sheet.replace(/'/g, "''")}'!`;
  const keyResp = await withRetry(
    () =>
      sheets.spreadsheets.values.get({
        spreadsheetId: cc.spreadsheetId,
        range: `${a1}${cc.keyCol}${cc.startRow}:${cc.keyCol}`,
        valueRenderOption: "UNFORMATTED_VALUE",
      }),
    { label: "carteira-get-keys" }
  );
  const destKeys = (keyResp.data.values || []).map((r) => r[0] ?? "");

  // Mesmo formato do script: "dd/MM/yyyy HH:mm - Entrada na Carteira-Planejador".
  const stamp =
    new Intl.DateTimeFormat("pt-BR", {
      timeZone: cfg.timezone || "America/Sao_Paulo",
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false,
    })
      .format(new Date())
      .replace(", ", " ") + (cc.stampSuffix || "");

  const plan = planUpsert({ srcRows, refHeader, destKeys, cfg: cc, stamp });
  console.log(
    `[carteira] ${srcRows.length} linha(s) na ultima Tecnica; destino com chave ate a linha ${plan.lastFilledRow}. ` +
      `Atualizadas: ${plan.updated} | Inseridas: ${plan.inserted} (a partir da linha ${plan.firstNewRow}).`
  );

  if (dryRun) {
    if (plan.newRows.length)
      console.log(`[dry-run] carteira 1a linha nova: ${JSON.stringify(plan.newRows[0]).slice(0, 800)}`);
    return;
  }

  // Garante grade suficiente pras linhas novas (aba nao cresce sozinha).
  const lastNeeded = plan.firstNewRow + plan.newRows.length - 1;
  if (plan.newRows.length) {
    const meta = await withRetry(
      () =>
        sheets.spreadsheets.get({
          spreadsheetId: cc.spreadsheetId,
          fields: "sheets.properties(sheetId,title,gridProperties.rowCount)",
        }),
      { label: "carteira-meta" }
    );
    const props = meta.data.sheets.map((s) => s.properties).find((p) => p.title === cc.sheet);
    if (!props) throw new Error(`[carteira] aba "${cc.sheet}" nao encontrada.`);
    const missing = lastNeeded - props.gridProperties.rowCount;
    if (missing > 0) {
      await withRetry(
        () =>
          sheets.spreadsheets.batchUpdate({
            spreadsheetId: cc.spreadsheetId,
            requestBody: {
              requests: [{ appendDimension: { sheetId: props.sheetId, dimension: "ROWS", length: missing } }],
            },
          }),
        { label: "carteira-append-rows" }
      );
      console.log(`[carteira] +${missing} linha(s) na grade.`);
    }
  }

  // Existentes: 1 range por coluna mapeada; null = celula nao tocada (preserva o que o
  // planejador digitou nas linhas que nao estao na ultima Tecnica).
  const data = [];
  if (plan.updated) {
    for (const [col, vals] of Object.entries(plan.updates)) {
      const L = numToCol(Number(col));
      data.push({
        range: `${a1}${L}${cc.startRow}:${L}${plan.lastFilledRow}`,
        values: vals.map((v) => [v]),
      });
    }
  }
  // Novas: linha inteira A..ultima coluna, igual ao script.
  if (plan.newRows.length) {
    data.push({
      range: `${a1}A${plan.firstNewRow}:${numToCol(plan.width)}${lastNeeded}`,
      values: plan.newRows,
    });
  }
  if (!data.length) return;
  await withRetry(
    () =>
      sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: cc.spreadsheetId,
        requestBody: { valueInputOption: cfg.valueInputOption, data },
      }),
    { label: "carteira-write" }
  );
  console.log(`[carteira] gravado (${data.length} range(s)).`);
}

async function main() {
  const cfg = loadConfig();
  const dryRun = process.env.DRY_RUN === "1" || process.argv.includes("--dry-run");

  const authClient = await getAuthClient(SCOPES);
  const sheets = google.sheets({ version: "v4", auth: authClient });
  const drive = google.drive({ version: "v3", auth: authClient });

  // 1) Cabecalho de referencia (nomes + ordem alvo) lido do proprio destino.
  const refResp = await withRetry(
    () =>
      sheets.spreadsheets.values.get({
        spreadsheetId: cfg.destSpreadsheetId,
        range: `'${cfg.destSheet.replace(/'/g, "''")}'!${cfg.headerRefRange}`,
        valueRenderOption: "UNFORMATTED_VALUE",
        dateTimeRenderOption: "FORMATTED_STRING",
      }),
    { label: "get-header-ref" }
  );
  const refHeader = (refResp.data.values && refResp.data.values[0]) || [];
  if (!refHeader.length)
    throw new Error(`Cabecalho de referencia vazio em ${cfg.destSheet}!${cfg.headerRefRange}.`);
  const width = refHeader.length;
  const keyIndex = refHeader.findIndex((h) => normHeader(h) === normHeader(cfg.keyColumn));
  if (keyIndex < 0)
    throw new Error(
      `Coluna chave "${cfg.keyColumn}" nao achada no cabecalho de referencia (${cfg.headerRefRange}).`
    );
  console.log(`[ref] ${width} coluna(s); chave "${cfg.keyColumn}" no indice ${keyIndex}.`);

  // 2) Lista os .xlsx da pasta no Drive.
  const files = await listXlsxInFolder(drive, cfg.driveFolderId);
  const xlsxFiles = files.filter((f) =>
    f.name.toLowerCase().endsWith((cfg.fileExtension || ".xlsx").toLowerCase())
  );
  if (!xlsxFiles.length)
    throw new Error(
      `Nenhum ${cfg.fileExtension} na pasta ${cfg.driveFolderId}. ` +
        `Confira o ID e se a Service Account tem acesso ao Shared Drive.`
    );
  console.log(`[drive] ${xlsxFiles.length} arquivo(s) ${cfg.fileExtension}:`);

  // 3) Baixa, le, valida cabecalho e reordena cada arquivo.
  const prepared = [];
  const skipped = [];
  for (const f of xlsxFiles) {
    const date = parseFileDate(f.name);
    if (date === null) {
      skipped.push(`${f.name} -> sem data no nome (nao da pra ordenar por recencia)`);
      console.warn(`  [PULA] ${f.name}: sem data no nome.`);
      continue;
    }
    const buf = await downloadFile(drive, f.id);
    const { header, rows } = readXlsx(buf, cfg.sourceHeaderRow);
    const { indices, missing } = buildColumnMap(refHeader, header);
    // Faltar a coluna chave (TRABALHO) inviabiliza o dedup -> sempre pula.
    const keyMissing = missing.some((m) => normHeader(m) === normHeader(cfg.keyColumn));
    const maxMissing = Number(cfg.maxMissingCols) || 0;
    if (keyMissing || missing.length > maxMissing) {
      const why = keyMissing
        ? `falta a coluna chave "${cfg.keyColumn}"`
        : `faltam ${missing.length} coluna(s) (> maxMissingCols=${maxMissing})`;
      skipped.push(`${f.name} -> ${why}: ${missing.slice(0, 8).join(", ")}${missing.length > 8 ? "..." : ""}`);
      console.warn(`  [PULA] ${f.name}: ${why}.`);
      continue;
    }
    if (missing.length) {
      console.warn(`  [AVISO] ${f.name}: ${missing.length} coluna(s) faltando, entram VAZIAS: ${missing.join(", ")}`);
    }
    const reordered = rows.map((r) => reorderRow(r, indices));
    prepared.push({ date, name: f.name, rows: reordered });
    console.log(`  [OK]   ${f.name} (${date}): ${reordered.length} linha(s).`);
  }

  if (skipped.length) {
    console.warn(`[aviso] ${skipped.length} arquivo(s) pulado(s):`);
    skipped.forEach((s) => console.warn(`  - ${s}`));
  }
  if (!prepared.length)
    throw new Error("Nenhum arquivo valido pra compilar (todos pulados). Abortando sem escrever.");

  // 4) Compila + deduplica por TRABALHO (vence o arquivo mais recente).
  const { rows: compiled, status, maxDate, inactive, totalRead, uniqueKeys, emptyKeyRows } =
    compile(prepared, keyIndex);
  const rect = padToWidth(compiled, width);
  console.log(
    `[compila] ${totalRead} linha(s) lidas -> ${rect.length} final ` +
      `(${uniqueKeys} TRABALHO unico${emptyKeyRows ? ` + ${emptyKeyRows} sem TRABALHO` : ""}).`
  );
  if (cfg.statusEnabled) {
    console.log(
      `[status] arquivo mais recente = ${ymdToStr(maxDate)}; ` +
        `${rect.length - inactive} ATIVO(s), ${inactive} INATIVO(s) (sairam do cadastro).`
    );
  }

  // Guarda anti-apagao.
  const minRows = Number(cfg.minRows) || 0;
  if (rect.length < minRows)
    throw new Error(
      `So ${rect.length} linha(s), abaixo de minRows=${minRows}. Abortando SEM limpar o destino.`
    );

  // Valores das colunas de status (ATIVO, ULTIMA_BASE), alinhados 1:1 com rect.
  const statusValues = cfg.statusEnabled
    ? status.map((s) => [s.ativo, ymdToStr(s.ultimaBase)])
    : null;

  // Ultima Tecnica = TRABALHO presente no arquivo mais recente (ATIVO=SIM).
  const ultimaTecnica = rect.filter((_, i) => status[i].ativo === "SIM");

  if (dryRun) {
    console.log("[dry-run] nao escreve. Amostra (1a linha):");
    console.log(JSON.stringify(rect[0] || [], null, 0).slice(0, 800));
    if (cfg.statusEnabled) {
      const inativos = status.filter((s) => s.ativo === "NAO");
      console.log(`[dry-run] status 1a linha: ${JSON.stringify(statusValues[0])}`);
      if (inativos.length)
        console.log(
          `[dry-run] INATIVOS (${inativos.length}), ultima base: ` +
            inativos.map((s) => ymdToStr(s.ultimaBase)).join(", ")
        );
    }
    await atualizarCarteira(sheets, cfg, refHeader, ultimaTecnica, true);
    return;
  }

  // 5) Limpa o destino e escreve.
  const destA1 = `'${cfg.destSheet.replace(/'/g, "''")}'!`;
  await withRetry(
    () =>
      sheets.spreadsheets.values.clear({
        spreadsheetId: cfg.destSpreadsheetId,
        range: `${destA1}${cfg.destClearRange}`,
      }),
    { label: "clear-dest" }
  );
  console.log(`[destino] limpo ${cfg.destClearRange}`);

  await withRetry(
    () =>
      sheets.spreadsheets.values.update({
        spreadsheetId: cfg.destSpreadsheetId,
        range: `${destA1}${cfg.destStartCell}`,
        valueInputOption: cfg.valueInputOption,
        requestBody: { values: rect },
      }),
    { label: "write-dest" }
  );
  console.log(`[destino] ${rect.length} linha(s) x ${width} col a partir de ${cfg.destStartCell}.`);

  // 5b) Colunas de status (ATIVO, ULTIMA_BASE), fora do espelho D:DK.
  if (cfg.statusEnabled) {
    // Cabecalho (idempotente).
    await withRetry(
      () =>
        sheets.spreadsheets.values.update({
          spreadsheetId: cfg.destSpreadsheetId,
          range: `${destA1}${cfg.statusHeaderRange}`,
          valueInputOption: "RAW",
          requestBody: { values: [cfg.statusHeaders] },
        }),
      { label: "write-status-header" }
    );
    // Limpa a faixa de dados antiga e escreve a nova.
    await withRetry(
      () =>
        sheets.spreadsheets.values.clear({
          spreadsheetId: cfg.destSpreadsheetId,
          range: `${destA1}${cfg.statusClearRange}`,
        }),
      { label: "clear-status" }
    );
    await withRetry(
      () =>
        sheets.spreadsheets.values.update({
          spreadsheetId: cfg.destSpreadsheetId,
          range: `${destA1}${cfg.statusStartCell}`,
          valueInputOption: cfg.valueInputOption,
          requestBody: { values: statusValues },
        }),
      { label: "write-status" }
    );
    console.log(`[destino] status ${statusValues.length} linha(s) em ${cfg.statusStartCell} (${cfg.statusHeaders.join(", ")}).`);
  }

  // 6) Carimbo opcional.
  if (cfg.stampCell) {
    const stamp = new Intl.DateTimeFormat("pt-BR", {
      timeZone: cfg.timezone || "America/Sao_Paulo",
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hour12: false,
    })
      .format(new Date())
      .replace(", ", " "); // "dd/mm/aaaa, hh:mm:ss" -> "dd/mm/aaaa hh:mm:ss" pra virar datetime real (USER_ENTERED)
    await withRetry(
      () =>
        sheets.spreadsheets.values.update({
          spreadsheetId: cfg.destSpreadsheetId,
          range: `${destA1}${cfg.stampCell}`,
          valueInputOption: cfg.valueInputOption,
          requestBody: { values: [[stamp]] },
        }),
      { label: "write-stamp" }
    );
    console.log(`[destino] carimbo ${cfg.stampCell}: ${stamp}`);
  }

  // 7) Etapa 2: Carteira_Planejador.
  await atualizarCarteira(sheets, cfg, refHeader, ultimaTecnica, false);

  console.log("OK.");
}

main().catch((err) => {
  console.error("FALHOU:", err.message);
  if (err.response?.data) console.error(JSON.stringify(err.response.data, null, 2));
  process.exit(1);
});
