const fs = require('fs');
const path = require('path');

const dataPath = path.join(__dirname, 'data.json');
if (!fs.existsSync(dataPath)) {
  console.log('data.json not found!');
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(dataPath, 'utf8'));

console.log('--- STATS DATA ---');
console.log(JSON.stringify(data.stats, null, 2));

console.log('\n--- SIMULADOS LP DATA SAMPLE ---');
Object.keys(data.simulados.LP).forEach(grade => {
  console.log(`Grade: ${grade}`);
  Object.keys(data.simulados.LP[grade]).forEach(simu => {
    const list = data.simulados.LP[grade][simu];
    console.log(`  ${simu}: ${list.length} rows`);
    if (list.length > 0) {
      console.log(`    Sample BNCC: ${list[0].bncc} - ${list[0].desc.substring(0, 100)}...`);
    }
  });
});

console.log('\n--- SIMULADOS MAT DATA SAMPLE ---');
Object.keys(data.simulados.MAT).forEach(grade => {
  console.log(`Grade: ${grade}`);
  Object.keys(data.simulados.MAT[grade]).forEach(simu => {
    const list = data.simulados.MAT[grade][simu];
    console.log(`  ${simu}: ${list.length} rows`);
    if (list.length > 0) {
      console.log(`    Sample BNCC: ${list[0].bncc} - ${list[0].desc.substring(0, 100)}...`);
    }
  });
});
