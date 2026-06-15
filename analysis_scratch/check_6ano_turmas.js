const xlsx = require('xlsx');
const path = require('path');

const file = 'F2 - CONSELHO 1TRI 2026.xlsx';
const filePath = path.join(__dirname, '..', file);
const workbook = xlsx.readFile(filePath);

const sheet = workbook.Sheets['6ANO'];
const rawData = xlsx.utils.sheet_to_json(sheet, { header: 1 });

console.log('Total rows in 6ANO:', rawData.length);

const turmaCount = {};
const studentTurma = {};

for (let i = 1; i < rawData.length; i++) {
  const row = rawData[i];
  if (!row || row.length === 0) continue;
  
  const turma = String(row[0] || '').trim();
  const studentRef = row[1];
  const disciplina = String(row[2] || '').trim();
  const tri1 = parseFloat(row[3]);
  
  if (studentRef === undefined || studentRef === null || studentRef === '') continue;
  
  const studentId = String(studentRef).trim();
  
  if (!turmaCount[turma]) {
    turmaCount[turma] = 0;
  }
  turmaCount[turma]++;
  
  studentTurma[studentId] = turma;
}

console.log('Row counts by Turma field in sheet:', turmaCount);

const turmaStudents = {};
Object.entries(studentTurma).forEach(([sId, turma]) => {
  if (!turmaStudents[turma]) {
    turmaStudents[turma] = [];
  }
  turmaStudents[turma].push(sId);
});

Object.entries(turmaStudents).forEach(([turma, list]) => {
  console.log(`Turma: ${turma} | Unique Student IDs: ${list.length} | IDs:`, list.sort((a,b) => parseInt(a) - parseInt(b)).join(', '));
});
