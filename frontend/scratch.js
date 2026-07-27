const XLSX = require('xlsx');
const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet([
  [{v: 46160, t: 'n', z: 'm/d/yyyy', s: { font: { bold: true } }}]
]);
XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
const out = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', cellStyles: true });
const parsed = XLSX.read(out, { cellStyles: true, cellNF: true });
console.log(JSON.stringify(parsed.Sheets.Sheet1.A1, null, 2));
