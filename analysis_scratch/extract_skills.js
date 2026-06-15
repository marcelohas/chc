const xlsx = require('xlsx');
const path = require('path');

const lpPath = path.join(__dirname, '..', 'FTD_SIMU_2025_LÍNGUA PORTUGUESA.xlsx');
const matPath = path.join(__dirname, '..', 'FTD_SIMU_2025_MATEMÁTICA.xlsx');

const lpWb = xlsx.readFile(lpPath);
const matWb = xlsx.readFile(matPath);

const targetSheets = ['6ANO', '7ANO', '8ANO'];

console.log('=== LÍNGUA PORTUGUESA SKILLS ===');
targetSheets.forEach(grade => {
  console.log(`\n--- ${grade} ---`);
  lpWb.SheetNames.forEach(sheetName => {
    if (!sheetName.includes(grade)) return;
    const sheet = lpWb.Sheets[sheetName];
    const rows = xlsx.utils.sheet_to_json(sheet);
    console.log(`Sheet: ${sheetName} (${rows.length} skills)`);
    rows.forEach(r => {
      console.log(`  [${r['BNCC']}] - ${String(r['Descritores / Objetos de Conhecimento'] || '').substring(0, 120)}... (Escola: ${(parseFloat(r['% acerto escola']) * 100).toFixed(1)}%, Geral: ${(parseFloat(r['% acerto geral']) * 100).toFixed(1)}%, Diff: ${(parseFloat(r['Comparativo']) * 100).toFixed(1)}%)`);
    });
  });
});

console.log('\n=== MATEMÁTICA SKILLS ===');
targetSheets.forEach(grade => {
  console.log(`\n--- ${grade} ---`);
  matWb.SheetNames.forEach(sheetName => {
    if (!sheetName.includes(grade)) return;
    const sheet = matWb.Sheets[sheetName];
    const rows = xlsx.utils.sheet_to_json(sheet);
    console.log(`Sheet: ${sheetName} (${rows.length} skills)`);
    rows.forEach(r => {
      console.log(`  [${r['BNCC']}] - ${String(r['Descritores / Objetos de Conhecimento'] || '').substring(0, 120)}... (Escola: ${(parseFloat(r['% acerto escola']) * 100).toFixed(1)}%, Geral: ${(parseFloat(r['% acerto geral']) * 100).toFixed(1)}%, Diff: ${(parseFloat(r['Comparativo']) * 100).toFixed(1)}%)`);
    });
  });
});
