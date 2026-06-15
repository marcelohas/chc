const xlsx = require('xlsx');
const path = require('path');

const baseDir = path.join(__dirname, '..');

const files = [
  'FTD_SIMU_2025_LÍNGUA PORTUGUESA.xlsx',
  'FTD_SIMU_2025_MATEMÁTICA.xlsx'
];

files.forEach(file => {
  console.log(`\n=========================================`);
  console.log(`FILE: ${file}`);
  console.log(`=========================================`);
  const filePath = path.join(baseDir, file);
  const workbook = xlsx.readFile(filePath);
  console.log('Sheets:', workbook.SheetNames);
  
  workbook.SheetNames.forEach(sheetName => {
    console.log(`\n--- Sheet: ${sheetName} ---`);
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    console.log(`Total rows: ${data.length}`);
    if (data.length > 0) {
      console.log('Columns:', JSON.stringify(data[0]));
      console.log('Sample rows (first 3):');
      data.slice(1, 4).forEach((row, idx) => {
        console.log(`  Row ${idx + 1}:`, JSON.stringify(row).substring(0, 300));
      });
    }
  });
});
