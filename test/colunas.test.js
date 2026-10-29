const { test } = require("node:test");
const assert = require("node:assert");
const { resolverColunas, agruparFaixas } = require("../src/colunas");

test("resolverColunas: acha por nome ignorando caixa/quebra de linha", () => {
  const h = ["", "APTA?", "TOTAL \nPT", "Log"];
  assert.deepEqual(resolverColunas(h, ["total pt", "LOG", "APTA?"]), { "total pt": 3, LOG: 4, "APTA?": 2 });
});

test("resolverColunas: faltando ou duplicado = erro", () => {
  assert.throws(() => resolverColunas(["A", "B"], ["C"]), /Faltando: C/);
  assert.throws(() => resolverColunas(["A", "B", "a"], ["A"]), /Duplicado: A/);
});

test("agruparFaixas junta colunas vizinhas", () => {
  assert.deepEqual(agruparFaixas([2, 3, 4, 15, 16, 37, 51]), ["B:D", "O:P", "AK", "AY"]);
  // coluna inserida no meio quebra a faixa em duas (nao e limpa)
  assert.deepEqual(agruparFaixas([18, 19, 21, 22]), ["R:S", "U:V"]);
});
