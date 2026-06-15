const fs = require('fs');
const path = require('path');

const data = JSON.parse(fs.readFileSync(path.join(__dirname, 'data.json'), 'utf8'));

['6ANO', '7ANO', '8ANO'].forEach(grade => {
  console.log(`\n=== PREDICTIVE POWER FOR ${grade} ===`);
  const students = data.conselho[grade];
  
  // Group students by CHC grade: 6.0-6.9 (Low/Medium) vs 7.0-7.9 (Medium) vs 8.0-10 (High)
  const low = students.filter(s => s.chc !== null && s.chc >= 6.0 && s.chc < 7.0);
  const mid = students.filter(s => s.chc !== null && s.chc >= 7.0 && s.chc < 8.0);
  const high = students.filter(s => s.chc !== null && s.chc >= 8.0);
  
  const printAverages = (group, name) => {
    const validMat = group.filter(s => s.mat !== null).map(s => s.mat);
    const validLp = group.filter(s => s.lp !== null).map(s => s.lp);
    
    const matAvg = validMat.length > 0 ? (validMat.reduce((a,b)=>a+b,0)/validMat.length).toFixed(2) : 'N/A';
    const lpAvg = validLp.length > 0 ? (validLp.reduce((a,b)=>a+b,0)/validLp.length).toFixed(2) : 'N/A';
    
    console.log(`  CHC [${name}] (N=${group.length}) -> Avg Internal Matemática: ${matAvg} | Avg Internal Língua Portuguesa: ${lpAvg}`);
  };
  
  printAverages(low, '6.0 - 6.9');
  printAverages(mid, '7.0 - 7.9');
  printAverages(high, '8.0 - 10.0');
});
