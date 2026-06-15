const xlsx = require('xlsx');
const path = require('path');

const file = 'F2 - CONSELHO 1TRI 2026.xlsx';
const filePath = path.join(__dirname, '..', file);
const workbook = xlsx.readFile(filePath);

const targetSheets = ['6ANO', '7ANO', '8ANO'];

targetSheets.forEach(sheetName => {
  console.log(`\n=========================================`);
  console.log(`SHEET: ${sheetName}`);
  console.log(`=========================================`);
  
  const sheet = workbook.Sheets[sheetName];
  const rawData = xlsx.utils.sheet_to_json(sheet, { header: 1 });
  
  const studentKeys = {};
  const turmas = {};
  
  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    if (!row || row.length === 0) continue;
    
    const turma = String(row[0] || '').trim();
    const studentRef = row[1];
    
    if (studentRef === undefined || studentRef === null || studentRef === '') continue;
    
    const studentKey = `${turma}_${String(studentRef).trim().toUpperCase()}`;
    
    studentKeys[studentKey] = {
      turma,
      ref: studentRef
    };
    
    if (!turmas[turma]) {
      turmas[turma] = new Set();
    }
    turmas[turma].add(String(studentRef).trim().toUpperCase());
  }
  
  console.log(`Total unique student keys (turma + ref): ${Object.keys(studentKeys).length}`);
  Object.entries(turmas).forEach(([turma, set]) => {
    console.log(`  Turma: ${turma} | Unique students: ${set.size}`);
  });
});
