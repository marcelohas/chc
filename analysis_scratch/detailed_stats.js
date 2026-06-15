const xlsx = require('xlsx');
const path = require('path');

const file = 'F2 - CONSELHO 1TRI 2026.xlsx';
const filePath = path.join(__dirname, '..', file);
const workbook = xlsx.readFile(filePath);

const targetSheets = ['6ANO', '7ANO', '8ANO'];

function pearsonCorrelation(x, y) {
  const n = x.length;
  if (n === 0) return 0;
  
  const sumX = x.reduce((a, b) => a + b, 0);
  const sumY = y.reduce((a, b) => a + b, 0);
  
  const meanX = sumX / n;
  const meanY = sumY / n;
  
  let num = 0;
  let denX = 0;
  let denY = 0;
  
  for (let i = 0; i < n; i++) {
    const diffX = x[i] - meanX;
    const diffY = y[i] - meanY;
    num += diffX * diffY;
    denX += diffX * diffX;
    denY += diffY * diffY;
  }
  
  if (denX === 0 || denY === 0) return 0;
  return num / Math.sqrt(denX * denY);
}

targetSheets.forEach(sheetName => {
  console.log(`\n=========================================`);
  console.log(`SHEET: ${sheetName}`);
  console.log(`=========================================`);
  
  const sheet = workbook.Sheets[sheetName];
  const rawData = xlsx.utils.sheet_to_json(sheet, { header: 1 });
  
  const students = {};
  const turmas = new Set();
  
  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    if (!row || row.length === 0) continue;
    
    const turma = String(row[0]).trim();
    const studentRef = row[1];
    const disciplina = String(row[2]).trim();
    const tri1 = parseFloat(row[3]);
    
    if (studentRef === undefined || studentRef === null || studentRef === '') continue;
    if (!disciplina) continue;
    
    const studentId = String(studentRef).trim();
    turmas.add(turma);
    
    if (!students[studentId]) {
      students[studentId] = {
        turma: turma,
        disciplinas: {}
      };
    }
    
    if (!isNaN(tri1)) {
      students[studentId].disciplinas[disciplina] = tri1;
    }
  }
  
  console.log('Turmas found:', Array.from(turmas));
  
  const chcList = [];
  const matList = [];
  let chcHigherCount = 0;
  let matHigherCount = 0;
  let equalCount = 0;
  
  Object.keys(students).forEach(id => {
    const s = students[id];
    const chc = s.disciplinas['Computação Humanidades e Criatividade'];
    const mat = s.disciplinas['Matemática'];
    
    if (chc !== undefined && mat !== undefined) {
      chcList.push(chc);
      matList.push(mat);
      
      if (chc > mat) {
        chcHigherCount++;
      } else if (mat > chc) {
        matHigherCount++;
      } else {
        equalCount++;
      }
    }
  });
  
  const correlation = pearsonCorrelation(chcList, matList);
  console.log(`Students with both grades: ${chcList.length}`);
  console.log(`Pearson Correlation (CHC vs Mat): ${correlation.toFixed(4)}`);
  console.log(`Students with CHC > Matemática: ${chcHigherCount} (${((chcHigherCount / chcList.length) * 100).toFixed(1)}%)`);
  console.log(`Students with Matemática > CHC: ${matHigherCount} (${((matHigherCount / chcList.length) * 100).toFixed(1)}%)`);
  console.log(`Students with CHC == Matemática: ${equalCount} (${((equalCount / chcList.length) * 100).toFixed(1)}%)`);
});
