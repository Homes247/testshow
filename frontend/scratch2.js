const XLSX = require('xlsx');
const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet([
  [{v: 46160, t: 'n', z: 'm/d/yyyy', w: '5/18/2026'}]
]);
XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
console.log("Raw true:", XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }));
console.log("Raw false:", XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false }));
