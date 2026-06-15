const xlsx = require('xlsx');
const path = require('path');

const baseDir = path.join(__dirname, '..');
const file = 'F2 - CONSELHO 1TRI 2026.xlsx';
const filePath = path.join(baseDir, file);

console.log(`Analyzing file: ${file}`);
const workbook = xlsx.readFile(filePath);
console.log('Sheets:', workbook.SheetNames);

workbook.SheetNames.forEach(sheetName => {
  console.log(`\n--- Sheet: ${sheetName} ---`);
  const sheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
  console.log(`Total rows: ${data.length}`);
  if (data.length > 0) {
    console.log('Sample rows (first 10):');
    data.slice(0, 10).forEach((row, idx) => {
      console.log(`  Row ${idx}:`, JSON.stringify(row).substring(0, 350));
    });
  }
});
