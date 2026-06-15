const xlsx = require('xlsx');
const path = require('path');

const lpPath = path.join(__dirname, '..', 'FTD_SIMU_2025_LÍNGUA PORTUGUESA.xlsx');
const matPath = path.join(__dirname, '..', 'FTD_SIMU_2025_MATEMÁTICA.xlsx');

const lpWb = xlsx.readFile(lpPath);
const matWb = xlsx.readFile(matPath);

const targetSheets = ['6ANO', '7ANO', '8ANO'];

console.log('=== LÍNGUA PORTUGUESA HIGHLIGHTS ===');
targetSheets.forEach(grade => {
  console.log(`\n================== ${grade} ==================`);
  let allSkills = [];
  
  lpWb.SheetNames.forEach(sheetName => {
    if (!sheetName.includes(grade)) return;
    const sheet = lpWb.Sheets[sheetName];
    const rows = xlsx.utils.sheet_to_json(sheet);
    rows.forEach(r => {
      const school = parseFloat(r['% acerto escola']);
      const general = parseFloat(r['% acerto geral']);
      const comp = parseFloat(r['Comparativo']);
      if (!isNaN(school)) {
        allSkills.push({
          bncc: String(r['BNCC']).trim(),
          desc: String(r['Descritores / Objetos de Conhecimento'] || '').trim(),
          school: school * 100,
          general: general * 100,
          comp: (school - general) * 100,
          sheet: sheetName
        });
      }
    });
  });
  
  // Sort by school performance ascending (lowest first)
  allSkills.sort((a, b) => a.school - b.school);
  console.log('Lowest Performers (School % < 50% or bottom 3):');
  allSkills.slice(0, 3).forEach(s => {
    console.log(`  - [${s.bncc}] (${s.sheet}): School: ${s.school.toFixed(1)}% | Geral: ${s.general.toFixed(1)}% | Diff: ${s.comp.toFixed(1)}%`);
    console.log(`    Desc: ${s.desc}`);
  });
  
  console.log('Highest Performers (top 3):');
  allSkills.slice(-3).reverse().forEach(s => {
    console.log(`  - [${s.bncc}] (${s.sheet}): School: ${s.school.toFixed(1)}% | Geral: ${s.general.toFixed(1)}% | Diff: ${s.comp.toFixed(1)}%`);
    console.log(`    Desc: ${s.desc}`);
  });
});

console.log('\n=== MATEMÁTICA HIGHLIGHTS ===');
targetSheets.forEach(grade => {
  console.log(`\n================== ${grade} ==================`);
  let allSkills = [];
  
  matWb.SheetNames.forEach(sheetName => {
    if (!sheetName.includes(grade)) return;
    const sheet = matWb.Sheets[sheetName];
    const rows = xlsx.utils.sheet_to_json(sheet);
    rows.forEach(r => {
      const school = parseFloat(r['% acerto escola']);
      const general = parseFloat(r['% acerto geral']);
      const comp = parseFloat(r['Comparativo']);
      if (!isNaN(school)) {
        allSkills.push({
          bncc: String(r['BNCC']).trim(),
          desc: String(r['Descritores / Objetos de Conhecimento'] || '').trim(),
          school: school * 100,
          general: general * 100,
          comp: (school - general) * 100,
          sheet: sheetName
        });
      }
    });
  });
  
  // Sort by school performance ascending (lowest first)
  allSkills.sort((a, b) => a.school - b.school);
  console.log('Lowest Performers (School % < 50% or bottom 3):');
  allSkills.slice(0, 3).forEach(s => {
    console.log(`  - [${s.bncc}] (${s.sheet}): School: ${s.school.toFixed(1)}% | Geral: ${s.general.toFixed(1)}% | Diff: ${s.comp.toFixed(1)}%`);
    console.log(`    Desc: ${s.desc}`);
  });
  
  console.log('Highest Performers (top 3):');
  allSkills.slice(-3).reverse().forEach(s => {
    console.log(`  - [${s.bncc}] (${s.sheet}): School: ${s.school.toFixed(1)}% | Geral: ${s.general.toFixed(1)}% | Diff: ${s.comp.toFixed(1)}%`);
    console.log(`    Desc: ${s.desc}`);
  });
});
