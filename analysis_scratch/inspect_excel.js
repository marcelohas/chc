const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const baseDir = path.join(__dirname, '..');
const files = [
  'F2 - CONSELHO 1TRI 2026.xlsx',
  'FTD_SIMU_2025_LÍNGUA PORTUGUESA.xlsx',
  'FTD_SIMU_2025_MATEMÁTICA.xlsx'
];

files.forEach(file => {
  const filePath = path.join(baseDir, file);
  if (!fs.existsSync(filePath)) {
    console.log(`File not found: ${file}`);
    return;
  }
  
  console.log(`\n=========================================`);
  console.log(`FILE: ${file}`);
  console.log(`=========================================`);
  
  const workbook = xlsx.readFile(filePath);
  console.log('Sheets:', workbook.SheetNames);
  
  workbook.SheetNames.forEach(sheetName => {
    console.log(`\n--- Sheet: ${sheetName} ---`);
    const sheet = workbook.Sheets[sheetName];
    // Convert to JSON
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    console.log(`Total rows: ${data.length}`);
    if (data.length > 0) {
      console.log('Sample rows (first 5):');
      data.slice(0, 5).forEach((row, idx) => {
        console.log(`  Row ${idx}:`, JSON.stringify(row).substring(0, 300));
      });
    }
  });
});
