const { test } = require("node:test");
const assert = require("node:assert");
const {
  normHeader,
  parseFileDate,
  buildColumnMap,
  reorderRow,
  compile,
  padToWidth,
  excelDateToStr,
  cellToValue,
  ymdToStr,
} = require("../src/transform");

test("excelDateToStr formata data-only e arredonda artefato de segundos", () => {
  // 25/02/2026 00:00:28 (artefato de serial) -> data pura
  const d = new Date(2026, 1, 25, 0, 0, 28);
  assert.equal(excelDateToStr(d), "25/02/2026");
});

test("excelDateToStr mantem hora quando relevante", () => {
  const d = new Date(2026, 4, 15, 14, 30, 0);
  assert.equal(excelDateToStr(d), "15/05/2026 14:30");
});

test("cellToValue: numero e string passam intactos, Date vira pt-BR", () => {
  assert.equal(cellToValue(46569.21), 46569.21);
  assert.equal(cellToValue("CCIM8504003"), "CCIM8504003");
  assert.equal(cellToValue(new Date(2026, 1, 25, 0, 0, 0)), "25/02/2026");
});

test("parseFileDate extrai data do nome", () => {
  assert.equal(parseFileDate("2026.06.25 - BASE SIRTEC_TECNICA.xlsx"), 20260625);
  assert.equal(parseFileDate("2026.06.03 - BASE SIRTEC_TECNICA"), 20260603);
  assert.equal(parseFileDate("BASE SEM DATA.xlsx"), null);
  assert.equal(parseFileDate("2026.13.01 invalida"), null);
});

test("parseFileDate ordena recencia corretamente", () => {
  assert.ok(parseFileDate("2026.06.25 x") > parseFileDate("2026.06.03 x"));
});

test("normHeader colapsa espacos e caixa", () => {
  assert.equal(normHeader("  Trabalho  "), "TRABALHO");
  assert.equal(normHeader("Data   Inicio"), "DATA INICIO");
});

test("buildColumnMap reordena cabecalho fora de ordem", () => {
  const ref = ["TRABALHO", "STATUS", "DATA"];
  const file = ["DATA", "TRABALHO", "STATUS"];
  const { indices, missing } = buildColumnMap(ref, file);
  assert.deepEqual(indices, [1, 2, 0]);
  assert.equal(missing.length, 0);
});

test("buildColumnMap ignora colunas extras do arquivo", () => {
  const ref = ["TRABALHO", "STATUS"];
  const file = ["EXTRA1", "TRABALHO", "EXTRA2", "STATUS"];
  const { indices, missing } = buildColumnMap(ref, file);
  assert.deepEqual(indices, [1, 3]);
  assert.equal(missing.length, 0);
});

test("buildColumnMap sinaliza colunas faltando", () => {
  const ref = ["TRABALHO", "STATUS", "FALTANTE"];
  const file = ["TRABALHO", "STATUS"];
  const { missing } = buildColumnMap(ref, file);
  assert.deepEqual(missing, ["FALTANTE"]);
});

test("reorderRow aplica indices e preenche faltante com ''", () => {
  assert.deepEqual(reorderRow(["a", "b", "c"], [2, 0, -1]), ["c", "a", ""]);
});

test("compile: arquivo mais recente vence pro mesmo TRABALHO", () => {
  // ref: [TRABALHO(0), VALOR(1)]
  const antigo = { date: 20260603, name: "antigo", rows: [["T1", "velho"], ["T2", "x"]] };
  const novo = { date: 20260625, name: "novo", rows: [["T1", "novo"]] };
  const { rows, uniqueKeys } = compile([novo, antigo], 0); // ordem de entrada baguncada de proposito
  assert.equal(uniqueKeys, 2);
  const t1 = rows.find((r) => r[0] === "T1");
  assert.equal(t1[1], "novo"); // prevalece o de 25/06
});

test("compile: linha sem TRABALHO nao deduplica, fica no fim", () => {
  const f = { date: 20260601, name: "f", rows: [["T1", "a"], ["", "sem"], ["", "sem2"]] };
  const { rows, uniqueKeys, emptyKeyRows } = compile([f], 0);
  assert.equal(uniqueKeys, 1);
  assert.equal(emptyKeyRows, 2);
  assert.equal(rows.length, 3);
});

test("padToWidth normaliza largura", () => {
  assert.deepEqual(padToWidth([["a"], ["b", "c", "d"]], 2), [["a", ""], ["b", "c"]]);
});

test("compile: status marca ATIVO/INATIVO por presenca no arquivo mais recente", () => {
  // T1 e T3 estao no mais recente (25/06); T2 so no antigo (03/06) -> saiu do cadastro.
  const antigo = { date: 20260603, name: "antigo", rows: [["T1", "a"], ["T2", "b"]] };
  const novo = { date: 20260625, name: "novo", rows: [["T1", "a2"], ["T3", "c"]] };
  const { rows, status, maxDate, inactive } = compile([antigo, novo], 0);
  assert.equal(maxDate, 20260625);
  assert.equal(inactive, 1);
  const byKey = Object.fromEntries(rows.map((r, i) => [r[0], status[i]]));
  assert.equal(byKey.T1.ativo, "SIM");
  assert.equal(byKey.T1.ultimaBase, 20260625);
  assert.equal(byKey.T3.ativo, "SIM");
  assert.equal(byKey.T2.ativo, "NAO");
  assert.equal(byKey.T2.ultimaBase, 20260603); // ultima vez que apareceu
});

test("ymdToStr formata AAAAMMDD e trata vazio", () => {
  assert.equal(ymdToStr(20260715), "15/07/2026");
  assert.equal(ymdToStr(20260603), "03/06/2026");
  assert.equal(ymdToStr(null), "");
  assert.equal(ymdToStr(-1), "");
});
