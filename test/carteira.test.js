const { test } = require("node:test");
const assert = require("node:assert");
const { colToNum, numToCol, planUpsert } = require("../src/carteira");

test("colToNum/numToCol ida e volta", () => {
  assert.equal(colToNum("A"), 1);
  assert.equal(colToNum("BN"), 66);
  assert.equal(colToNum("DH"), 112);
  assert.equal(numToCol(66), "BN");
  assert.equal(numToCol(54), "BB");
});

const ref = ["NUM_SOB", "TRABALHO", "MUNICIPIO"];
const cfg = {
  startRow: 6,
  keyCol: "E",
  map: [
    { from: "TRABALHO", to: "E" },
    { from: "MUNICIPIO", to: "I" },
    { from: "NUM_SOB", to: "BL" },
  ],
  stampCol: "BB",
  cityCol: "BI",
  cityValue: "JUAZEIRO DO NORTE",
};

test("planUpsert: atualiza existente sem tocar chave/carimbo/cidade e insere novo", () => {
  const plan = planUpsert({
    srcRows: [
      [111, 8480490, "BREJO SANTO NOVO"], // existe (destino guarda numero)
      [222, "999", "CRATO"], // novo
      [333, "999", "DUPLICADO"], // mesma chave nova -> ignora
      [444, "", "SEM CHAVE"], // sem chave -> ignora
    ],
    refHeader: ref,
    destKeys: ["8480490", "", "777"],
    cfg,
    stamp: "29/10/2026 08:00 - Entrada na Carteira-Planejador",
  });
  assert.equal(plan.updated, 1);
  assert.equal(plan.inserted, 1);
  assert.equal(plan.lastFilledRow, 8);
  assert.equal(plan.firstNewRow, 9);
  assert.equal(plan.width, 64); // BL
  assert.deepEqual(plan.updates[colToNum("I")], ["BREJO SANTO NOVO", null, null]);
  assert.deepEqual(plan.updates[colToNum("BL")], [111, null, null]);
  assert.equal(plan.updates[colToNum("E")], undefined); // chave nunca e sobrescrita
  const nr = plan.newRows[0];
  assert.equal(nr[colToNum("E") - 1], "999");
  assert.equal(nr[colToNum("I") - 1], "CRATO");
  assert.equal(nr[colToNum("BB") - 1], "29/10/2026 08:00 - Entrada na Carteira-Planejador");
  assert.equal(nr[colToNum("BI") - 1], "JUAZEIRO DO NORTE");
});

test("planUpsert: destino vazio insere a partir de startRow", () => {
  const plan = planUpsert({ srcRows: [[1, "5", "X"]], refHeader: ref, destKeys: [], cfg, stamp: "s" });
  assert.equal(plan.firstNewRow, 6);
  assert.equal(plan.inserted, 1);
});

test("planUpsert: coluna do mapa inexistente na ref = erro", () => {
  assert.throws(() =>
    planUpsert({
      srcRows: [],
      refHeader: ref,
      destKeys: [],
      cfg: { ...cfg, map: [...cfg.map, { from: "NAO_EXISTE", to: "Z" }] },
      stamp: "",
    })
  );
});
