const { test } = require("node:test");
const assert = require("node:assert");
const { compararDecrescente, padRect, parseCols, lastFilledRow } = require("../src/formulas");
const { classificarPar, ROTULOS } = require("../src/geologia");

test("compararDecrescente: numeros desc, vazios no fim", () => {
  const v = [46200, "", 46323, null, 46209];
  v.sort(compararDecrescente);
  assert.deepEqual(v.slice(0, 3), [46323, 46209, 46200]);
  assert.ok(v.slice(3).every((x) => x === "" || x === null));
});

test("padRect completa linhas e colunas com vazio", () => {
  assert.deepEqual(padRect([[1], []], 3, 2), [[1, ""], ["", ""], ["", ""]]);
});

test("parseCols e lastFilledRow", () => {
  assert.deepEqual(parseCols("B:D"), { c1: 2, c2: 4 });
  assert.deepEqual(parseCols("AK"), { c1: 37, c2: 37 });
  assert.equal(lastFilledRow(["", "x", "", "y", ""]), 4);
  assert.equal(lastFilledRow([]), 0);
});

test("geologia: par invalido -> vazio; UTM do Ceara -> rotulo conhecido", () => {
  assert.equal(classificarPar(1, 1), "");
  assert.equal(classificarPar("", ""), "");
  const r = classificarPar(517490, 9186546); // Mauriti (fuso 24S)
  assert.ok(Object.values(ROTULOS).includes(r), r);
  // E/N trocados dao o mesmo resultado.
  assert.equal(classificarPar(9186546, 517490), r);
});
