const fs = require('fs');
const path = require('path');
const assert = require('assert');
const root = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const app = fs.readFileSync(path.join(root, 'app.js'), 'utf8');
const patch = fs.readFileSync(path.join(root, 'v384-patch.js'), 'utf8');
const config = fs.readFileSync(path.join(root, 'config.js'), 'utf8');
const backend = fs.readFileSync(path.join(root, '..', 'Apps_Script', 'Code.gs'), 'utf8');

for (const id of ['syncState','syncText','priceUpdatedText','syncBtn','pricingTable','adminHeaderBtn']) {
  assert(html.includes(`id="${id}"`), `missing #${id}`);
}
assert(html.includes('<script src="config.js"></script>'), 'config.js must load before app');
assert(config.includes('API_URL'), 'API_URL config missing');
assert(app.includes("apiGetJsonp('bootstrap'"), 'main bootstrap must use JSONP');
assert(app.includes('data-v384-price-date-head'), 'pricing date column missing');
assert(app.includes('lastPriceUpdate'), 'header price update field missing');
assert(app.includes("startsWith('3.8.4')"), 'backend version guard missing');
assert(patch.includes("apiGetJsonp('bootstrap'"), 'patch bootstrap JSONP fallback missing');

assert.strictEqual((backend.match(/function doGet\(/g) || []).length, 1, 'must have one doGet');
assert.strictEqual((backend.match(/function doPost\(/g) || []).length, 1, 'must have one doPost');
for (const feature of ['3.8.4-web','webJsonOrJsonpV384_','lastPriceUpdate','webReadLatestPriceLogV384_','priceRequestCreate','chatWrite']) {
  assert(backend.includes(feature), `backend feature missing: ${feature}`);
}
assert(backend.includes('window.top.postMessage'), 'bridge must post to top window');
console.log('Static sync/date checks: OK');
