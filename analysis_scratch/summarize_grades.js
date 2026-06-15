const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

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
  
  // Headers are in row 0
  const headers = rawData[0];
  console.log('Headers:', headers);
  
  // Group rows by student
  const students = {};
  
  // Iterate from row 1 to end
  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    if (!row || row.length === 0) continue;
    
    const turma = row[0];
    const studentRef = row[1]; // number in 6ANO, name in 7/8ANO
    const disciplina = row[2];
    const tri1 = parseFloat(row[3]); // 1st trimester grade
    
    if (studentRef === undefined || studentRef === null || studentRef === '') continue;
    if (!disciplina) continue;
    
    const studentId = String(studentRef).trim();
    
    if (!students[studentId]) {
      students[studentId] = {
        turma: String(turma).trim(),
        disciplinas: {}
      };
    }
    
    if (!isNaN(tri1)) {
      students[studentId].disciplinas[disciplina] = tri1;
    }
  }
  
  const studentList = Object.keys(students);
  console.log(`Total unique students identified: ${studentList.length}`);
  
  // Count how many have CHC and Matemática grades
  let countBoth = 0;
  let chcVals = [];
  let matVals = [];
  let lpVals = [];
  let redVals = [];
  
  studentList.forEach(id => {
    const s = students[id];
    const chc = s.disciplinas['Computação Humanidades e Criatividade'];
    const mat = s.disciplinas['Matemática'];
    const lp = s.disciplinas['Língua Portuguesa'];
    const red = s.disciplinas['Redação'];
    
    if (chc !== undefined) chcVals.push(chc);
    if (mat !== undefined) matVals.push(mat);
    if (lp !== undefined) lpVals.push(lp);
    if (red !== undefined) redVals.push(red);
    
    if (chc !== undefined && mat !== undefined) {
      countBoth++;
    }
  });
  
  console.log(`Students with CHC grades: ${chcVals.length}`);
  console.log(`Students with Matemática grades: ${matVals.length}`);
  console.log(`Students with both: ${countBoth}`);
  
  const avg = arr => (arr.reduce((a, b) => a + b, 0) / arr.length).toFixed(2);
  if (chcVals.length > 0) console.log(`CHC Average: ${avg(chcVals)}`);
  if (matVals.length > 0) console.log(`Matemática Average: ${avg(matVals)}`);
  if (lpVals.length > 0) console.log(`Língua Portuguesa Average: ${avg(lpVals)}`);
  if (redVals.length > 0) console.log(`Redação Average: ${avg(redVals)}`);
});
