const xlsx = require('xlsx');
const path = require('path');

const baseDir = path.join(__dirname, '..');
const files = {
  'Matemática': 'FTD_SIMU_2025_MATEMÁTICA.xlsx',
  'Língua Portuguesa': 'FTD_SIMU_2025_LÍNGUA PORTUGUESA.xlsx'
};

const grades = ['6ANO', '7ANO', '8ANO'];

Object.entries(files).forEach(([subject, filename]) => {
  console.log(`\n=========================================`);
  console.log(`Subject: ${subject} (${filename})`);
  console.log(`=========================================`);
  
  const filePath = path.join(baseDir, filename);
  const workbook = xlsx.readFile(filePath);
  
  grades.forEach(grade => {
    // Find sheets that match the grade. 
    // They might be named "FTD1 - 6ANO", "FTD2 - 6ANO", "FTD - 6ANO", etc.
    const matchingSheets = workbook.SheetNames.filter(name => name.includes(grade));
    
    matchingSheets.forEach(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      const data = xlsx.utils.sheet_to_json(sheet);
      
      let sumEscola = 0;
      let sumGeral = 0;
      let count = 0;
      
      data.forEach(row => {
        const escolaVal = parseFloat(row['% acerto escola']);
        const geralVal = parseFloat(row['% acerto geral']);
        
        if (!isNaN(escolaVal) && !isNaN(geralVal)) {
          sumEscola += escolaVal;
          sumGeral += geralVal;
          count++;
        }
      });
      
      if (count > 0) {
        const avgEscola = (sumEscola / count) * 100;
        const avgGeral = (sumGeral / count) * 100;
        console.log(`Sheet: ${sheetName} | Skills: ${count} | School Avg: ${avgEscola.toFixed(1)}% | General Avg: ${avgGeral.toFixed(1)}% | Diff: ${(avgEscola - avgGeral).toFixed(1)}%`);
      } else {
        console.log(`Sheet: ${sheetName} | No valid rows found`);
      }
    });
  });
});
