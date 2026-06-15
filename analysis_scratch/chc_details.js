const fs = require('fs');
const path = require('path');

const data = JSON.parse(fs.readFileSync(path.join(__dirname, 'data.json'), 'utf8'));

['6ANO', '7ANO', '8ANO'].forEach(grade => {
  const students = data.conselho[grade];
  const chc = students.map(s => s.chc).filter(v => v !== null);
  const mat = students.map(s => s.mat).filter(v => v !== null);
  const lp = students.map(s => s.lp).filter(v => v !== null);
  const red = students.map(s => s.red).filter(v => v !== null);
  
  console.log(`\n=== GRADE DETAILS FOR ${grade} ===`);
  console.log(`CHC - Min: ${Math.min(...chc)}, Max: ${Math.max(...chc)}, Avg: ${(chc.reduce((a,b)=>a+b,0)/chc.length).toFixed(2)}, Count: ${chc.length}`);
  console.log(`MAT - Min: ${Math.min(...mat)}, Max: ${Math.max(...mat)}, Avg: ${(mat.reduce((a,b)=>a+b,0)/mat.length).toFixed(2)}, Count: ${mat.length}`);
  console.log(`LP  - Min: ${Math.min(...lp)},  Max: ${Math.max(...lp)},  Avg: ${(lp.reduce((a,b)=>a+b,0)/lp.length).toFixed(2)},  Count: ${lp.length}`);
  console.log(`RED - Min: ${Math.min(...red)}, Max: ${Math.max(...red)}, Avg: ${(red.reduce((a,b)=>a+b,0)/red.length).toFixed(2)}, Count: ${red.length}`);
  
  // Count specific values in CHC
  const counts = {};
  chc.forEach(v => {
    counts[v] = (counts[v] || 0) + 1;
  });
  console.log('CHC grade frequency:', counts);
});
