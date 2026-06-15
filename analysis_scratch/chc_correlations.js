const fs = require('fs');
const path = require('path');

const dataPath = path.join(__dirname, 'data.json');
if (!fs.existsSync(dataPath)) {
  console.log('data.json not found!');
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(dataPath, 'utf8'));

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

const targetGrades = ['6ANO', '7ANO', '8ANO'];

targetGrades.forEach(grade => {
  const students = data.conselho[grade];
  
  // Filter students who have CHC and others
  const chc_mat = students.filter(s => s.chc !== null && s.mat !== null);
  const chc_lp = students.filter(s => s.chc !== null && s.lp !== null);
  const chc_red = students.filter(s => s.chc !== null && s.red !== null);
  
  const corr_mat = pearsonCorrelation(chc_mat.map(s => s.chc), chc_mat.map(s => s.mat));
  const corr_lp = pearsonCorrelation(chc_lp.map(s => s.chc), chc_lp.map(s => s.lp));
  const corr_red = pearsonCorrelation(chc_red.map(s => s.chc), chc_red.map(s => s.red));
  
  console.log(`\n================= ${grade} =================`);
  console.log(`Total Students: ${students.length}`);
  console.log(`Correlation CHC vs Matemática: ${corr_mat.toFixed(4)}`);
  console.log(`Correlation CHC vs Língua Portuguesa: ${corr_lp.toFixed(4)}`);
  console.log(`Correlation CHC vs Redação: ${corr_red.toFixed(4)}`);
  
  // Calculate average grades
  const chc_avg = chc_mat.reduce((sum, s) => sum + s.chc, 0) / chc_mat.length;
  const mat_avg = chc_mat.reduce((sum, s) => sum + s.mat, 0) / chc_mat.length;
  const lp_avg = chc_lp.reduce((sum, s) => sum + s.lp, 0) / chc_lp.length;
  const red_avg = chc_red.reduce((sum, s) => sum + s.red, 0) / chc_red.length;
  
  console.log(`Averages - CHC: ${chc_avg.toFixed(2)} | MAT: ${mat_avg.toFixed(2)} | LP: ${lp_avg.toFixed(2)} | RED: ${red_avg.toFixed(2)}`);
  
  // Let's count by performance range for CHC
  // Turma	< 2,0	2,0 a 4,9	5,0 a 5,9	6,0 a 6,9	7,0 a 10
  const turmas = {};
  students.forEach(s => {
    if (s.chc === null) return;
    const t = s.turma;
    if (!turmas[t]) {
      turmas[t] = { '<2.0': 0, '2.0-4.9': 0, '5.0-5.9': 0, '6.0-6.9': 0, '7.0-10': 0, total: 0 };
    }
    const val = s.chc;
    if (val < 2.0) turmas[t]['<2.0']++;
    else if (val < 5.0) turmas[t]['2.0-4.9']++;
    else if (val < 6.0) turmas[t]['5.0-5.9']++;
    else if (val < 7.0) turmas[t]['6.0-6.9']++;
    else turmas[t]['7.0-10']++;
    turmas[t].total++;
  });
  
  console.log('CHC Grade Distribution by Turma:');
  Object.keys(turmas).sort().forEach(t => {
    const d = turmas[t];
    console.log(`  Turma ${t}: <2.0: ${d['<2.0']} | 2.0-4.9: ${d['2.0-4.9']} | 5.0-5.9: ${d['5.0-5.9']} | 6.0-6.9: ${d['6.0-6.9']} | 7.0-10: ${d['7.0-10']} | Total: ${d.total}`);
  });
});
