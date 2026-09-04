const assert = require('assert');
const data = require('../data.js');
const Engine = require('../engine.js');

const check = Engine.runGoldenTests(data);
assert.strictEqual(check.pass, true, JSON.stringify(check.failed.slice(0, 10), null, 2));
assert.strictEqual(check.passed, 186);
assert.strictEqual(check.total, 186);

const allConstructions = Object.values(data.constructions).flat();
assert.strictEqual(allConstructions.some(x => /полистирол/i.test(x.name)), false);
assert.strictEqual(data.constructions.part.some(x => /^не проходит/i.test(String(x.reference))), false);
assert.strictEqual(allConstructions.flatMap(x => x.components).some(x => /доборн/i.test(x.name) && Math.abs(Number(x.sheetContribution)||0) > 1e-9), false);

for (const code of ['M-048','M-050','M-053','M-055']) {
  const m = data.materials.find(x => x.code === code);
  assert(m, `missing ${code}`);
  assert.strictEqual(m.unit, 'м³', `${code} must be m3`);
}

console.log(`Golden tests: ${check.passed}/${check.total}`);
