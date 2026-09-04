const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync(require('path').join(__dirname,'..','index.html'),'utf8');
const app = fs.readFileSync(require('path').join(__dirname,'..','app.js'),'utf8');
const backend = fs.readFileSync(require('path').join(__dirname,'..','Code.gs'),'utf8');

for (const id of ['adminHeaderBtn','pdfBtn','historyMaterialPicker','certificateTable','newMaterialForm','certificateUploadForm','instructionDialog']) {
  assert(html.includes(`id="${id}"`), `missing #${id}`);
}
for (const action of ['priceHistory','buildPdfReport','downloadCertificate','addMaterial','uploadCertificate','deleteCertificate']) {
  assert(backend.includes(`'${action}'`), `backend action missing: ${action}`);
}
assert(app.includes("Backend Apps Script ещё не обновлён до v3.8"));
console.log('Static feature checks: OK');
