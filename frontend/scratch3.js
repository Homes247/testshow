const XLSX = require('xlsx');
const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet([
  ['Test']
]);
ws['A1'].s = { fill: { fgColor: { rgb: "FFFF0000" } } };
ws['!cols'] = [{ wch: 20 }];
XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
const out = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
const parsed = XLSX.read(out, { cellStyles: true });
console.log("Cols:", parsed.Sheets.Sheet1['!cols']);
console.log("A1 Style:", parsed.Sheets.Sheet1.A1.s);
