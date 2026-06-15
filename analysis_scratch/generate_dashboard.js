const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

// Path definition
const baseDir = path.join(__dirname, '..');
const conselhoFile = 'F2 - CONSELHO 1TRI 2026.xlsx';
const lpSimuFile = 'FTD_SIMU_2025_LÍNGUA PORTUGUESA.xlsx';
const matSimuFile = 'FTD_SIMU_2025_MATEMÁTICA.xlsx';

const conselhoPath = path.join(baseDir, conselhoFile);
const lpSimuPath = path.join(baseDir, lpSimuFile);
const matSimuPath = path.join(baseDir, matSimuFile);

// Helper to calculate Pearson Correlation
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

// 1. Process Conselho 2026 Data
console.log('Loading Conselho 2026 data...');
const conselhoWorkbook = xlsx.readFile(conselhoPath);
const targetSheets = ['6ANO', '7ANO', '8ANO'];
const conselhoData = {};
const statsData = {};

targetSheets.forEach(sheetName => {
  const sheet = conselhoWorkbook.Sheets[sheetName];
  const rawData = xlsx.utils.sheet_to_json(sheet, { header: 1 });
  
  const students = {};
  const turmaNameToId = {};
  const turmaNextId = {};
  
  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    if (!row || row.length === 0) continue;
    
    const turma = String(row[0]).trim();
    const studentRef = row[1];
    const disciplinaRaw = row[2];
    const tri1 = parseFloat(row[3]);
    
    if (studentRef === undefined || studentRef === null || studentRef === '') continue;
    if (!disciplinaRaw) continue;
    
    const disciplina = String(disciplinaRaw).trim();
    
    // Determine anonymized student ID local to the turma
    let studentId;
    if (sheetName === '6ANO') {
      studentId = parseInt(studentRef, 10);
      if (isNaN(studentId)) {
        const key = String(studentRef).trim().toUpperCase();
        if (!turmaNameToId[turma]) {
          turmaNameToId[turma] = {};
          turmaNextId[turma] = 1;
        }
        if (!turmaNameToId[turma][key]) {
          turmaNameToId[turma][key] = turmaNextId[turma]++;
        }
        studentId = turmaNameToId[turma][key];
      }
    } else {
      const key = String(studentRef).trim().toUpperCase();
      if (!turmaNameToId[turma]) {
        turmaNameToId[turma] = {};
        turmaNextId[turma] = 1;
      }
      if (!turmaNameToId[turma][key]) {
        turmaNameToId[turma][key] = turmaNextId[turma]++;
      }
      studentId = turmaNameToId[turma][key];
    }
    
    const studentKey = `${turma}_${studentId}`;
    
    if (!students[studentKey]) {
      students[studentKey] = {
        id: studentId,
        turma: turma,
        grades: {}
      };
    }
    
    if (!isNaN(tri1)) {
      let key = '';
      if (disciplina.includes('Computação') || disciplina.includes('CHC')) {
        key = 'chc';
      } else if (disciplina.includes('Matemática')) {
        key = 'mat';
      } else if (disciplina.includes('Língua Portuguesa') || disciplina.includes('LP')) {
        key = 'lp';
      } else if (disciplina.includes('Redação')) {
        key = 'red';
      }
      
      if (key) {
        students[studentKey].grades[key] = tri1;
      }
    }
  }
  
  const studentList = Object.values(students).map(s => {
    return {
      id: s.id,
      turma: s.turma,
      chc: s.grades.chc !== undefined ? s.grades.chc : null,
      mat: s.grades.mat !== undefined ? s.grades.mat : null,
      lp: s.grades.lp !== undefined ? s.grades.lp : null,
      red: s.grades.red !== undefined ? s.grades.red : null
    };
  });
  
  conselhoData[sheetName] = studentList;
  
  // Calculate Stats
  const chcVals = studentList.filter(s => s.chc !== null).map(s => s.chc);
  const matVals = studentList.filter(s => s.mat !== null).map(s => s.mat);
  const lpVals = studentList.filter(s => s.lp !== null).map(s => s.lp);
  const redVals = studentList.filter(s => s.red !== null).map(s => s.red);
  
  const chcMatBoth = studentList.filter(s => s.chc !== null && s.mat !== null);
  const chcBoth = chcMatBoth.map(s => s.chc);
  const matBoth = chcMatBoth.map(s => s.mat);
  
  let chcHigherCount = 0;
  let matHigherCount = 0;
  let equalCount = 0;
  
  chcMatBoth.forEach(s => {
    if (s.chc > s.mat) chcHigherCount++;
    else if (s.mat > s.chc) matHigherCount++;
    else equalCount++;
  });
  
  const correlation = pearsonCorrelation(chcBoth, matBoth);
  const mean = arr => arr.length > 0 ? parseFloat((arr.reduce((a, b) => a + b, 0) / arr.length).toFixed(2)) : null;
  
  statsData[sheetName] = {
    totalStudents: studentList.length,
    chcAvg: mean(chcVals),
    matAvg: mean(matVals),
    lpAvg: mean(lpVals),
    redAvg: mean(redVals),
    correlation: parseFloat(correlation.toFixed(4)),
    chcHigherCount,
    matHigherCount,
    equalCount,
    bothCount: chcMatBoth.length
  };
});

console.log('Conselho data processed!');

// 2. Process Simulator 2025 Data
console.log('Loading Simulator 2025 data...');
const simuData = {
  LP: {},
  MAT: {}
};

const lpSimuWorkbook = xlsx.readFile(lpSimuPath);
const matSimuWorkbook = xlsx.readFile(matSimuPath);

const simuGrades = ['6ANO', '7ANO', '8ANO'];

simuGrades.forEach(grade => {
  simuData.LP[grade] = {};
  simuData.MAT[grade] = {};
  
  // LP Simulator Sheets
  const lpSheets = lpSimuWorkbook.SheetNames.filter(name => name.includes(grade));
  lpSheets.forEach(sheetName => {
    const simuType = sheetName.includes('FTD1') ? 'FTD1' : sheetName.includes('FTD2') ? 'FTD2' : 'Outro';
    const sheet = lpSimuWorkbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet);
    
    simuData.LP[grade][simuType] = data.map(row => ({
      bncc: String(row['BNCC'] || '').trim(),
      desc: String(row['Descritores / Objetos de Conhecimento'] || '').trim(),
      acertoGeral: parseFloat(row['% acerto geral']) || null,
      acertoEscola: parseFloat(row['% acerto escola']) || null,
      comparativo: parseFloat(row['Comparativo']) || null,
      plano: String(row['Plano de Ação para execução em 2025'] || row['Plano de Ação para execução em 2026'] || '').trim(),
      justificativa: String(row['Justificativa docente para nossa taxa de acerto'] || '').trim()
    })).filter(row => row.bncc !== '');
  });
  
  // MAT Simulator Sheets
  const matSheets = matSimuWorkbook.SheetNames.filter(name => name.includes(grade));
  matSheets.forEach(sheetName => {
    const simuType = sheetName.includes('FTD1') ? 'FTD1' : sheetName.includes('FTD2') ? 'FTD2' : 'Outro';
    const sheet = matSimuWorkbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet);
    
    simuData.MAT[grade][simuType] = data.map(row => ({
      bncc: String(row['BNCC'] || '').trim(),
      desc: String(row['Descritores / Objetos de Conhecimento'] || '').trim(),
      acertoGeral: parseFloat(row['% acerto geral']) || null,
      acertoEscola: parseFloat(row['% acerto escola']) || null,
      comparativo: parseFloat(row['Comparativo']) || null,
      plano: String(row['Plano de Ação para execução em 2025'] || row['Plano de Ação para execução em 2026'] || '').trim(),
      justificativa: String(row['Justificativa docente para nossa taxa de acerto'] || '').trim()
    })).filter(row => row.bncc !== '');
  });
});

console.log('Simulators data processed!');

// Combine into one database
const fullData = {
  conselho: conselhoData,
  stats: statsData,
  simulados: simuData
};

// Write JSON for verification
fs.writeFileSync(path.join(baseDir, 'analysis_scratch', 'data.json'), JSON.stringify(fullData, null, 2));
console.log('data.json exported.');

// Generate HTML
const htmlTemplate = `<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Análise Acadêmica Integrada | Conselho & Simulados</title>
  <!-- Google Fonts -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
  <!-- FontAwesome Icons -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <!-- Chart.js -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  
  <style>
    :root {
      --bg-color: #0b0f19;
      --card-bg: #151c2c;
      --card-bg-hover: #1e293b;
      --text-main: #f3f4f6;
      --text-muted: #9ca3af;
      --border-color: #2d3748;
      --primary: #3b82f6;
      --primary-hover: #2563eb;
      --secondary: #10b981;
      --secondary-hover: #059669;
      --warning: #f59e0b;
      --danger: #ef4444;
      --purple: #8b5cf6;
      --pink: #ec4899;
      --font-family: 'Outfit', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }
    
    .light-theme {
      --bg-color: #f8fafc;
      --card-bg: #ffffff;
      --card-bg-hover: #f1f5f9;
      --text-main: #0f172a;
      --text-muted: #64748b;
      --border-color: #e2e8f0;
      --primary: #2563eb;
      --primary-hover: #1d4ed8;
      --secondary: #059669;
      --secondary-hover: #047857;
      --warning: #d97706;
      --danger: #dc2626;
      --purple: #7c3aed;
      --pink: #db2777;
    }
    
    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
      transition: background-color 0.3s ease, border-color 0.3s ease, color 0.3s ease;
    }
    
    body {
      background-color: var(--bg-color);
      color: var(--text-main);
      font-family: var(--font-family);
      min-height: 100vh;
      display: flex;
      flex-direction: column;
      overflow-x: hidden;
    }
    
    /* Layout Header */
    header {
      background-color: var(--card-bg);
      border-bottom: 1px solid var(--border-color);
      padding: 1rem 2rem;
      display: flex;
      align-items: center;
      justify-content: space-between;
      position: sticky;
      top: 0;
      z-index: 100;
      flex-wrap: wrap;
      gap: 1rem;
    }
    
    .brand {
      display: flex;
      align-items: center;
      gap: 0.75rem;
    }
    
    .brand i {
      font-size: 1.8rem;
      color: var(--primary);
    }
    
    .brand h1 {
      font-size: 1.3rem;
      font-weight: 700;
      letter-spacing: -0.5px;
    }
    
    .nav-tabs {
      display: flex;
      gap: 0.5rem;
      background-color: var(--bg-color);
      padding: 0.25rem;
      border-radius: 8px;
      border: 1px solid var(--border-color);
    }
    
    .nav-btn {
      background: none;
      border: none;
      color: var(--text-muted);
      padding: 0.5rem 1.25rem;
      border-radius: 6px;
      cursor: pointer;
      font-weight: 500;
      font-size: 0.9rem;
      display: flex;
      align-items: center;
      gap: 0.5rem;
    }
    
    .nav-btn.active {
      background-color: var(--card-bg);
      color: var(--primary);
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
    }
    
    .header-actions {
      display: flex;
      align-items: center;
      gap: 1rem;
    }
    
    .theme-toggle, .btn-refresh {
      background-color: var(--bg-color);
      border: 1px solid var(--border-color);
      color: var(--text-main);
      width: 40px;
      height: 40px;
      border-radius: 8px;
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 1rem;
    }
    
    .theme-toggle:hover, .btn-refresh:hover {
      background-color: var(--border-color);
    }
    
    /* Main Layout */
    .container {
      max-width: 1600px;
      width: 100%;
      margin: 0 auto;
      padding: 2rem;
      flex-grow: 1;
      display: flex;
      flex-direction: column;
      gap: 2rem;
    }
    
    /* Filters Area */
    .filters-bar {
      background-color: var(--card-bg);
      border: 1px solid var(--border-color);
      border-radius: 12px;
      padding: 1.25rem;
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 1.5rem;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    
    .filters-left {
      display: flex;
      align-items: center;
      gap: 1rem;
      flex-wrap: wrap;
    }
    
    .filter-label {
      font-size: 0.85rem;
      font-weight: 600;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    
    .select-style {
      background-color: var(--bg-color);
      border: 1px solid var(--border-color);
      color: var(--text-main);
      padding: 0.5rem 2rem 0.5rem 1rem;
      border-radius: 8px;
      outline: none;
      cursor: pointer;
      font-family: var(--font-family);
      font-size: 0.9rem;
      font-weight: 500;
      appearance: none;
      background-image: url("data:image/svg+xml;utf8,<svg fill='%239ca3af' height='24' viewBox='0 0 24 24' width='24' xmlns='http://www.w3.org/2000/svg'><path d='M7 10l5 5 5-5z'/><path d='M0 0h24v24H0z' fill='none'/></svg>");
      background-repeat: no-repeat;
      background-position-x: calc(100% - 8px);
      background-position-y: 50%;
    }
    
    .select-style:focus {
      border-color: var(--primary);
    }
    
    .axis-selector-box {
      border-left: 2px solid var(--primary);
      padding-left: 0.75rem;
      display: flex;
      align-items: center;
      gap: 0.75rem;
    }
    
    /* Tab View Management */
    .tab-content {
      display: none;
      flex-direction: column;
      gap: 2rem;
    }
    
    .tab-content.active {
      display: flex;
    }
    
    /* KPI Grid */
    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
      gap: 1.5rem;
    }
    
    .kpi-card {
      background-color: var(--card-bg);
      border: 1px solid var(--border-color);
      border-radius: 12px;
      padding: 1.5rem;
      display: flex;
      flex-direction: column;
      gap: 0.5rem;
      position: relative;
      overflow: hidden;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    
    .kpi-card::before {
      content: '';
      position: absolute;
      top: 0;
      left: 0;
      width: 4px;
      height: 100%;
      background-color: var(--primary);
    }
    
    .kpi-card.secondary::before {
      background-color: var(--secondary);
    }
    
    .kpi-card.warning::before {
      background-color: var(--warning);
    }
    
    .kpi-card.danger::before {
      background-color: var(--danger);
    }
    
    .kpi-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      color: var(--text-muted);
      font-size: 0.85rem;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    
    .kpi-header i {
      font-size: 1.2rem;
    }
    
    .kpi-value {
      font-size: 1.8rem;
      font-weight: 700;
      color: var(--text-main);
    }
    
    .kpi-footer {
      font-size: 0.8rem;
      color: var(--text-muted);
      margin-top: 0.25rem;
    }
    
    /* Grid Charts & Content */
    .dashboard-grid {
      display: grid;
      grid-template-columns: 3fr 2fr;
      gap: 1.5rem;
    }
    
    @media (max-width: 1024px) {
      .dashboard-grid {
        grid-template-columns: 1fr;
      }
    }
    
    .chart-container {
      background-color: var(--card-bg);
      border: 1px solid var(--border-color);
      border-radius: 12px;
      padding: 1.5rem;
      display: flex;
      flex-direction: column;
      gap: 1rem;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
      min-height: 400px;
    }
    
    .chart-title {
      font-size: 1.05rem;
      font-weight: 600;
      display: flex;
      align-items: center;
      justify-content: space-between;
    }
    
    .chart-title-left {
      display: flex;
      align-items: center;
      gap: 0.5rem;
    }
    
    .chart-title-left i {
      color: var(--primary);
    }
    
    .canvas-wrapper {
      position: relative;
      flex-grow: 1;
      width: 100%;
      height: 100%;
    }
    
    /* Table & Details Area */
    .section-card {
      background-color: var(--card-bg);
      border: 1px solid var(--border-color);
      border-radius: 12px;
      padding: 1.5rem;
      display: flex;
      flex-direction: column;
      gap: 1.25rem;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    
    .section-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 1rem;
    }
    
    .section-title {
      font-size: 1.1rem;
      font-weight: 600;
      display: flex;
      align-items: center;
      gap: 0.5rem;
    }
    
    .section-title i {
      color: var(--primary);
    }
    
    .table-actions {
      display: flex;
      gap: 1rem;
      align-items: center;
      flex-wrap: wrap;
    }
    
    .search-input-wrapper {
      position: relative;
      width: 200px;
    }
    
    .search-input-wrapper i {
      position: absolute;
      left: 10px;
      top: 50%;
      transform: translateY(-50%);
      color: var(--text-muted);
      font-size: 0.9rem;
    }
    
    .search-input {
      background-color: var(--bg-color);
      border: 1px solid var(--border-color);
      color: var(--text-main);
      padding: 0.5rem 1rem 0.5rem 2.2rem;
      border-radius: 8px;
      font-family: var(--font-family);
      font-size: 0.9rem;
      width: 100%;
      outline: none;
    }
    
    .search-input:focus {
      border-color: var(--primary);
    }
    
    /* Table Design */
    .table-responsive {
      overflow-x: auto;
      border-radius: 8px;
      border: 1px solid var(--border-color);
    }
    
    table {
      width: 100%;
      border-collapse: collapse;
      text-align: left;
      font-size: 0.9rem;
    }
    
    thead {
      background-color: var(--bg-color);
      color: var(--text-muted);
      font-weight: 600;
    }
    
    th {
      padding: 0.75rem 1rem;
      border-bottom: 1px solid var(--border-color);
      cursor: pointer;
      user-select: none;
    }
    
    th:hover {
      color: var(--text-main);
    }
    
    th i {
      margin-left: 0.25rem;
      font-size: 0.8rem;
    }
    
    td {
      padding: 0.85rem 1rem;
      border-bottom: 1px solid var(--border-color);
    }
    
    tr:last-child td {
      border-bottom: none;
    }
    
    tr:hover {
      background-color: var(--card-bg-hover);
    }
    
    .badge-turma {
      background-color: rgba(59, 130, 246, 0.15);
      color: var(--primary);
      padding: 0.2rem 0.5rem;
      border-radius: 4px;
      font-weight: 600;
      font-size: 0.75rem;
    }
    
    .grade-cell {
      font-weight: 600;
    }
    
    .grade-danger {
      color: var(--danger);
    }
    
    .grade-success {
      color: var(--secondary);
    }
    
    .grade-warning {
      color: var(--warning);
    }
    
    .diff-badge {
      display: inline-flex;
      align-items: center;
      gap: 0.25rem;
      font-weight: 700;
      padding: 0.15rem 0.4rem;
      border-radius: 4px;
      font-size: 0.8rem;
    }
    
    .diff-badge.pos {
      background-color: rgba(16, 185, 129, 0.15);
      color: var(--secondary);
    }
    
    .diff-badge.neg {
      background-color: rgba(59, 130, 246, 0.15);
      color: var(--primary);
    }
    
    .diff-badge.zero {
      background-color: rgba(156, 163, 175, 0.15);
      color: var(--text-muted);
    }
    
    /* Pedagogy Section */
    .pedagogy-container {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 1.5rem;
    }
    
    @media (max-width: 768px) {
      .pedagogy-container {
        grid-template-columns: 1fr;
      }
    }
    
    .pedagogy-card {
      background-color: var(--card-bg);
      border: 1px solid var(--border-color);
      border-radius: 12px;
      padding: 1.5rem;
      display: flex;
      flex-direction: column;
      gap: 1rem;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    
    .pedagogy-title {
      font-size: 1.05rem;
      font-weight: 600;
      display: flex;
      align-items: center;
      gap: 0.5rem;
    }
    
    .pedagogy-title i {
      color: var(--warning);
    }
    
    .pedagogy-text {
      color: var(--text-muted);
      font-size: 0.95rem;
      line-height: 1.6;
    }
    
    .pedagogy-text strong {
      color: var(--text-main);
    }
    
    .insight-list {
      list-style-type: none;
      display: flex;
      flex-direction: column;
      gap: 0.75rem;
    }
    
    .insight-item {
      display: flex;
      gap: 0.75rem;
      align-items: flex-start;
    }
    
    .insight-item i {
      color: var(--primary);
      margin-top: 0.25rem;
      font-size: 0.9rem;
    }
    
    .insight-desc {
      color: var(--text-muted);
      font-size: 0.9rem;
      line-height: 1.5;
    }
    
    .insight-desc strong {
      color: var(--text-main);
    }
    
    /* Simulator View Styling */
    .simu-top-kpis {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
      gap: 1.5rem;
    }
    
    .simu-skill-row {
      border: 1px solid var(--border-color);
      border-radius: 8px;
      padding: 1rem;
      margin-bottom: 1rem;
      background-color: var(--card-bg);
    }
    
    .simu-skill-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 0.5rem;
      flex-wrap: wrap;
      gap: 0.5rem;
    }
    
    .simu-skill-code {
      font-weight: 700;
      background-color: rgba(59, 130, 246, 0.15);
      color: var(--primary);
      padding: 0.2rem 0.5rem;
      border-radius: 4px;
      font-size: 0.8rem;
    }
    
    .simu-skill-comparison {
      display: flex;
      gap: 1rem;
      align-items: center;
      font-size: 0.85rem;
    }
    
    .bar-comparison-outer {
      background-color: var(--bg-color);
      height: 12px;
      border-radius: 6px;
      width: 100%;
      position: relative;
      margin: 0.5rem 0;
      overflow: hidden;
      border: 1px solid var(--border-color);
    }
    
    .bar-comparison-inner {
      background-color: var(--primary);
      height: 100%;
      border-radius: 5px;
    }
    
    .bar-comparison-national {
      position: absolute;
      top: 0;
      bottom: 0;
      width: 2px;
      background-color: var(--warning);
      z-index: 10;
    }
    
    .simu-card-details {
      font-size: 0.85rem;
      color: var(--text-muted);
      margin-top: 0.5rem;
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 1rem;
      border-top: 1px solid var(--border-color);
      padding-top: 0.5rem;
    }
    
    @media (max-width: 600px) {
      .simu-card-details {
        grid-template-columns: 1fr;
      }
    }
    
    .simu-card-action {
      background-color: rgba(245, 158, 11, 0.05);
      border-left: 2px solid var(--warning);
      padding: 0.4rem;
      border-radius: 0 4px 4px 0;
    }
    
    .simu-card-just {
      background-color: rgba(16, 185, 129, 0.05);
      border-left: 2px solid var(--secondary);
      padding: 0.4rem;
      border-radius: 0 4px 4px 0;
    }
    
    .empty-state {
      text-align: center;
      padding: 3rem;
      color: var(--text-muted);
      font-size: 1.1rem;
    }
    
    .empty-state i {
      font-size: 3rem;
      margin-bottom: 1rem;
      color: var(--border-color);
    }
    
    /* Report Layout Styling */
    .report-layout {
      display: flex;
      gap: 1.5rem;
      align-items: flex-start;
      margin-top: 1rem;
    }
    
    @media (max-width: 1024px) {
      .report-layout {
        flex-direction: column;
      }
      .report-sidebar {
        width: 100% !important;
      }
    }
    
    .report-sidebar {
      width: 300px;
      background-color: var(--card-bg);
      border: 1px solid var(--border-color);
      border-radius: 12px;
      padding: 1.25rem;
      flex-shrink: 0;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    
    .report-menu {
      list-style-type: none;
      display: flex;
      flex-direction: column;
      gap: 0.5rem;
    }
    
    .report-menu-item {
      padding: 0.85rem 1rem;
      border-radius: 8px;
      cursor: pointer;
      font-weight: 500;
      font-size: 0.95rem;
      display: flex;
      align-items: center;
      gap: 0.75rem;
      color: var(--text-muted);
      border: 1px solid transparent;
      transition: all 0.2s ease;
    }
    
    .report-menu-item:hover {
      background-color: var(--card-bg-hover);
      color: var(--text-main);
    }
    
    .report-menu-item.active {
      background-color: rgba(59, 130, 246, 0.1);
      border-color: var(--primary);
      color: var(--primary);
    }
    
    .report-menu-item .num {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 24px;
      height: 24px;
      border-radius: 50%;
      background-color: var(--bg-color);
      font-size: 0.8rem;
      font-weight: 700;
      color: var(--text-muted);
    }
    
    .report-menu-item.active .num {
      background-color: var(--primary);
      color: #ffffff;
    }
    
    .report-main-content {
      flex-grow: 1;
      background-color: var(--card-bg);
      border: 1px solid var(--border-color);
      border-radius: 12px;
      padding: 2.25rem;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
      min-height: 650px;
    }
    
    .report-section {
      display: none;
      flex-direction: column;
      gap: 1.5rem;
    }
    
    .report-section.active {
      display: flex;
      animation: fadeIn 0.3s ease;
    }
    
    .report-title {
      font-size: 1.4rem;
      font-weight: 700;
      color: var(--text-main);
      border-bottom: 2px solid var(--border-color);
      padding-bottom: 0.75rem;
      margin-bottom: 0.5rem;
      display: flex;
      align-items: center;
      gap: 0.75rem;
    }
    
    .report-title i {
      color: var(--primary);
    }
    
    .report-alert {
      background-color: rgba(245, 158, 11, 0.05);
      border-left: 4px solid var(--warning);
      padding: 1rem 1.25rem;
      border-radius: 0 8px 8px 0;
      color: var(--text-muted);
      font-size: 0.95rem;
      line-height: 1.6;
    }
    
    .report-alert strong {
      color: var(--text-main);
    }
    
    .report-alert.danger {
      background-color: rgba(239, 68, 68, 0.05);
      border-left-color: var(--danger);
    }
    
    .report-alert.info {
      background-color: rgba(59, 130, 246, 0.05);
      border-left-color: var(--primary);
    }
    
    .report-text {
      font-size: 1rem;
      color: var(--text-muted);
      line-height: 1.7;
    }
    
    .report-text strong {
      color: var(--text-main);
    }
    
    .report-text ul {
      margin-left: 1.5rem;
      margin-top: 0.5rem;
      display: flex;
      flex-direction: column;
      gap: 0.5rem;
    }
    
    .report-text li strong {
      color: var(--text-main);
    }
    
    @keyframes fadeIn {
      from { opacity: 0; transform: translateY(8px); }
      to { opacity: 1; transform: translateY(0); }
    }
  </style>
</head>
<body>

  <header>
    <div class="brand">
      <i class="fa-solid fa-graduation-cap"></i>
      <div>
        <h1>Dashboard Acadêmico</h1>
        <p style="font-size: 0.75rem; color: var(--text-muted);">Conselho & Simulados • 6º ao 8º Ano</p>
      </div>
    </div>
    
    <div class="nav-tabs">
      <button class="nav-btn active" id="btnTabConselho" onclick="switchTab('conselho')">
        <i class="fa-solid fa-users-rectangle"></i> Conselho 2026
      </button>
      <button class="nav-btn" id="btnTabSimulados" onclick="switchTab('simulados')">
        <i class="fa-solid fa-square-poll-vertical"></i> Simulados 2025
      </button>
      <button class="nav-btn" id="btnTabJuntos" onclick="switchTab('juntos')">
        <i class="fa-solid fa-code-compare"></i> Análise Integrada (Juntos)
      </button>
      <button class="nav-btn" id="btnTabRelatorio" onclick="switchTab('relatorio')">
        <i class="fa-solid fa-file-invoice"></i> Relatório CHC
      </button>
    </div>
    
    <div class="header-actions">
      <button class="btn-refresh" onclick="resetFilters()" title="Resetar Filtros">
        <i class="fa-solid fa-arrows-rotate"></i>
      </button>
      <button class="theme-toggle" onclick="toggleTheme()" title="Alternar Tema">
        <i class="fa-solid fa-moon"></i>
      </button>
    </div>
  </header>

  <div class="container">
    
    <!-- Filters Bar (Shared for tabs, behaves appropriately based on active tab) -->
    <div class="filters-bar" id="mainFiltersBar">
      <div class="filters-left">
        <div>
          <span class="filter-label">Ano Escolar:</span>
          <select id="seriesFilter" class="select-style" onchange="onSeriesChange()">
            <option value="6ANO">6º Ano</option>
            <option value="7ANO" selected>7º Ano</option>
            <option value="8ANO">8º Ano</option>
          </select>
        </div>
        
        <div id="conselhoSpecificFilters" class="filters-left">
          <div>
            <span class="filter-label">Turma:</span>
            <select id="turmaFilter" class="select-style" onchange="applyFilters()">
              <option value="ALL">Todas</option>
            </select>
          </div>
          
          <div class="axis-selector-box">
            <div>
              <span class="filter-label">Eixo X:</span>
              <select id="axisXFilter" class="select-style" onchange="onAxisChange()">
                <option value="chc" selected>CHC</option>
                <option value="mat">Matemática</option>
                <option value="lp">Língua Portuguesa</option>
                <option value="red">Redação</option>
              </select>
            </div>
            <div>
              <span class="filter-label">Eixo Y:</span>
              <select id="axisYFilter" class="select-style" onchange="onAxisChange()">
                <option value="chc">CHC</option>
                <option value="mat" selected>Matemática</option>
                <option value="lp">Língua Portuguesa</option>
                <option value="red">Redação</option>
              </select>
            </div>
          </div>
          
          <div>
            <span class="filter-label">Performance:</span>
            <select id="performanceFilter" class="select-style" onchange="applyFilters()">
              <!-- Dynamically Populated by JavaScript -->
            </select>
          </div>
        </div>
        
        <div id="simuSpecificFilters" class="filters-left" style="display: none;">
          <div>
            <span class="filter-label">Disciplina:</span>
            <select id="simuSubjectFilter" class="select-style" onchange="renderSimulados()">
              <option value="MAT" selected>Matemática</option>
              <option value="LP">Língua Portuguesa</option>
            </select>
          </div>
          <div>
            <span class="filter-label">Simulado:</span>
            <select id="simuTypeFilter" class="select-style" onchange="renderSimulados()">
              <option value="FTD1">Simulado 1 (FTD1)</option>
              <option value="FTD2" selected>Simulado 2 (FTD2)</option>
            </select>
          </div>
        </div>
        
        <div id="juntosSpecificFilters" class="filters-left" style="display: none;">
          <div>
            <span class="filter-label">Disciplina Comparativa:</span>
            <select id="juntosSubjectFilter" class="select-style" onchange="renderJuntos()">
              <option value="MAT" selected>Matemática</option>
              <option value="LP">Língua Portuguesa</option>
              <option value="CHC">CHC</option>
            </select>
          </div>
        </div>
      </div>
      
      <div class="table-actions" id="searchWrapper">
        <div class="search-input-wrapper">
          <i class="fa-solid fa-magnifying-glass"></i>
          <input type="text" id="tableSearch" class="search-input" placeholder="Buscar..." onkeyup="applyFilters()">
        </div>
      </div>
    </div>

    <!-- TAB 1: CONSELHO 2026 -->
    <div id="conselhoTab" class="tab-content active">
      
      <!-- KPI Grid -->
      <div class="kpi-grid">
        <div class="kpi-card">
          <div class="kpi-header">
            <span>Total de Alunos</span>
            <i class="fa-solid fa-users" style="color: var(--primary)"></i>
          </div>
          <div class="kpi-value" id="kpiTotalStudents">0</div>
          <div class="kpi-footer">Matriculados na série e turma selecionadas</div>
        </div>
        
        <div class="kpi-card secondary">
          <div class="kpi-header">
            <span id="kpiTitleX">Média CHC</span>
            <i class="fa-solid fa-laptop-code" style="color: var(--secondary)"></i>
          </div>
          <div class="kpi-value" id="kpiAvgX">0.0</div>
          <div class="kpi-footer">Nota média (1º Tri) no Eixo X</div>
        </div>
        
        <div class="kpi-card secondary">
          <div class="kpi-header">
            <span id="kpiTitleY">Média Matemática</span>
            <i class="fa-solid fa-calculator" style="color: var(--secondary)"></i>
          </div>
          <div class="kpi-value" id="kpiAvgY">0.0</div>
          <div class="kpi-footer">Nota média (1º Tri) no Eixo Y</div>
        </div>
        
        <div class="kpi-card warning" id="kpiCorrelationCard">
          <div class="kpi-header">
            <span id="kpiTitleCorr">Correlação</span>
            <i class="fa-solid fa-chart-line" style="color: var(--warning)"></i>
          </div>
          <div class="kpi-value" id="kpiCorrelation">0.000</div>
          <div class="kpi-footer" id="kpiCorrelationDesc">Calculando...</div>
        </div>
      </div>

      <!-- Dashboard Grid: Charts -->
      <div class="dashboard-grid">
        <!-- Scatter Plot CHC vs Math -->
        <div class="chart-container">
          <div class="chart-title">
            <div class="chart-title-left">
              <i class="fa-solid fa-circle-nodes"></i>
              <span id="scatterChartTitle">Correlação: Eixo X vs Eixo Y</span>
            </div>
            <span style="font-size: 0.8rem; color: var(--text-muted); font-weight: normal;">Ponto = 1 Aluno • Linha Tracejada = Igualdade</span>
          </div>
          <div class="canvas-wrapper">
            <canvas id="scatterChart"></canvas>
          </div>
        </div>
        
        <!-- Bar Chart: Average Grades comparison -->
        <div class="chart-container">
          <div class="chart-title">
            <div class="chart-title-left">
              <i class="fa-solid fa-chart-bar"></i>
              <span>Comparativo de Médias das Disciplinas</span>
            </div>
          </div>
          <div class="canvas-wrapper">
            <canvas id="barsChart"></canvas>
          </div>
        </div>
      </div>

      <!-- Data Table Section -->
      <div class="section-card">
        <div class="section-header">
          <div class="section-title">
            <i class="fa-solid fa-list-ol"></i>
            <span>Listagem Geral de Alunos (Notas Individuais)</span>
          </div>
          <div style="font-size: 0.85rem; color: var(--text-muted);" id="tableLegendText">
            Legenda: <span style="color: var(--secondary); font-weight: bold;">▲ Eixo X > Eixo Y</span> | <span style="color: var(--primary); font-weight: bold;">▼ Eixo Y > Eixo X</span>
          </div>
        </div>
        
        <div class="table-responsive">
          <table id="studentTable">
            <thead>
              <tr>
                <th onclick="sortStudentTable('id')">Nº Aluno <i class="fa-solid fa-sort"></i></th>
                <th onclick="sortStudentTable('turma')">Turma <i class="fa-solid fa-sort"></i></th>
                <th onclick="sortStudentTable('chc')">CHC (1º TRI) <i class="fa-solid fa-sort"></i></th>
                <th onclick="sortStudentTable('mat')">Matemática <i class="fa-solid fa-sort"></i></th>
                <th onclick="sortStudentTable('diff')">Diferença (X - Y) <i class="fa-solid fa-sort"></i></th>
                <th onclick="sortStudentTable('lp')">Língua Portuguesa <i class="fa-solid fa-sort"></i></th>
                <th onclick="sortStudentTable('red')">Redação <i class="fa-solid fa-sort"></i></th>
              </tr>
            </thead>
            <tbody id="studentTableBody">
              <!-- Dynamically populated -->
            </tbody>
          </table>
        </div>
      </div>
      
      <!-- Pedagogy & Statistics Analysis -->
      <div class="pedagogy-container">
        <!-- Core Analysis Card -->
        <div class="pedagogy-card">
          <div class="pedagogy-title">
            <i class="fa-solid fa-circle-info"></i>
            <span>Análise Pedagógica da Série Selecionada</span>
          </div>
          <div class="pedagogy-text" id="seriesPedagogicText">
            <!-- Dynamically populated -->
          </div>
        </div>
        
        <!-- Insights / Action Plan Card -->
        <div class="pedagogy-card">
          <div class="pedagogy-title">
            <i class="fa-solid fa-lightbulb" style="color: var(--secondary)"></i>
            <span>Recomendações e Plano de Ação Pedagógico</span>
          </div>
          <ul class="insight-list" id="seriesInsightsList">
            <!-- Dynamically populated -->
          </ul>
        </div>
      </div>
    </div>

    <!-- TAB 2: SIMULADOS 2025 -->
    <div id="simuladosTab" class="tab-content">
      
      <!-- Simulados Top KPI metrics -->
      <div class="simu-top-kpis">
        <div class="kpi-card" style="grid-column: span 1;">
          <div class="kpi-header">
            <span>Média da Escola no Simulado</span>
            <i class="fa-solid fa-school" style="color: var(--primary)"></i>
          </div>
          <div class="kpi-value" id="simuKpiEscola">0.0%</div>
          <div class="kpi-footer">Taxa de acertos média nos itens</div>
        </div>
        
        <div class="kpi-card secondary" style="grid-column: span 1;">
          <div class="kpi-header">
            <span>Média Geral (Nacional)</span>
            <i class="fa-solid fa-globe" style="color: var(--secondary)"></i>
          </div>
          <div class="kpi-value" id="simuKpiGeral">0.0%</div>
          <div class="kpi-footer">Referência geral da plataforma FTD</div>
        </div>
        
        <div class="kpi-card" id="simuKpiDiffCard" style="grid-column: span 1;">
          <div class="kpi-header">
            <span>Diferencial da Escola</span>
            <i class="fa-solid fa-scale-balanced" id="simuKpiDiffIcon"></i>
          </div>
          <div class="kpi-value" id="simuKpiDiff">0.0%</div>
          <div class="kpi-footer">Comparativo Escola vs Geral</div>
        </div>
      </div>
      
      <!-- Performance Summary Chart and Skills Detailed View -->
      <div class="dashboard-grid">
        <div class="chart-container" style="flex-grow: 1;">
          <div class="chart-title">
            <div class="chart-title-left">
              <i class="fa-solid fa-chart-column"></i>
              <span>Acertos por Habilidade da BNCC (Escola vs Geral)</span>
            </div>
          </div>
          <div class="canvas-wrapper">
            <canvas id="simuCompareChart"></canvas>
          </div>
        </div>
        
        <div class="section-card">
          <div class="section-title">
            <i class="fa-solid fa-bullseye"></i>
            <span>Resumo de Habilidades Avaliadas</span>
          </div>
          <p class="pedagogy-text" style="font-size: 0.85rem;">
            Os simulados FTD avaliam competências específicas da Base Nacional Comum Curricular (BNCC). 
            Abaixo estão os percentuais de acerto da escola em relação ao geral nacional, com plano de ação proposto pelo docente.
          </p>
          <div id="simuAveragesList" style="display: flex; flex-direction: column; gap: 0.5rem; max-height: 380px; overflow-y: auto; padding-right: 0.5rem;">
            <!-- Dynamically populated -->
          </div>
        </div>
      </div>
      
      <!-- Detailed Skill List Card -->
      <div class="section-card">
        <div class="section-title">
          <i class="fa-solid fa-microscope"></i>
          <span>Detalhamento por Habilidade BNCC & Justificativas Docentes</span>
        </div>
        
        <div id="simuSkillsContainer">
          <!-- Dynamically populated cards representing skill details -->
        </div>
      </div>
      
    </div>

    <!-- TAB 3: ANALISE INTEGRADA (JUNTOS) -->
    <div id="juntosTab" class="tab-content">
      
      <!-- Integrated KPIs -->
      <div class="kpi-grid">
        <div class="kpi-card">
          <div class="kpi-header">
            <span>Avaliação Interna (Conselho 2026)</span>
            <i class="fa-solid fa-graduation-cap" style="color: var(--primary)"></i>
          </div>
          <div class="kpi-value" id="juntosKpiInterna">0.00</div>
          <div class="kpi-footer" id="juntosKpiInternaLabel">Média de notas (escala 0-10)</div>
        </div>
        
        <div class="kpi-card secondary">
          <div class="kpi-header">
            <span>Avaliação Externa (Simulados 2025)</span>
            <i class="fa-solid fa-square-poll-vertical" style="color: var(--secondary)"></i>
          </div>
          <div class="kpi-value" id="juntosKpiExterna">0.0%</div>
          <div class="kpi-footer" id="juntosKpiExternaLabel">Taxa média de acerto (escola)</div>
        </div>
        
        <div class="kpi-card danger" id="juntosKpiGargaloCard">
          <div class="kpi-header">
            <span>Gargalo de Avaliação (Gap)</span>
            <i class="fa-solid fa-arrows-left-right" style="color: var(--danger)"></i>
          </div>
          <div class="kpi-value" id="juntosKpiGap">0.0%</div>
          <div class="kpi-footer">Diferença: Interno % - Externo %</div>
        </div>
      </div>
      
      <!-- Charts for Integrated Comparison -->
      <div class="dashboard-grid">
        <div class="chart-container" style="min-height: 420px;">
          <div class="chart-title">
            <div class="chart-title-left">
              <i class="fa-solid fa-chart-gantt"></i>
              <span id="juntosChartTitle">Comparativo: Avaliação Interna (Conselho) vs Externa (Simulados)</span>
            </div>
            <span style="font-size: 0.8rem; color: var(--text-muted); font-weight: normal;">Notas do Conselho multiplicadas por 10 para escala em %</span>
          </div>
          <div class="canvas-wrapper">
            <canvas id="juntosCompareChart"></canvas>
          </div>
        </div>
        
        <!-- Summary Table -->
        <div class="section-card">
          <div class="section-title">
            <i class="fa-solid fa-table-list"></i>
            <span>Tabela Comparativa Geral por Série</span>
          </div>
          <div class="table-responsive">
            <table>
              <thead>
                <tr style="background-color: var(--bg-color);">
                  <th>Série</th>
                  <th>Conselho (Interno)</th>
                  <th>Simulado 1 (FTD1)</th>
                  <th>Simulado 2 (FTD2)</th>
                  <th>Diferença (Gap)</th>
                </tr>
              </thead>
              <tbody id="juntosTableBody">
                <!-- Dynamically populated -->
              </tbody>
            </table>
          </div>
          <p style="font-size: 0.75rem; color: var(--text-muted); line-height: 1.4;">
            * A diferença (Gap) é calculada comparando a nota média do Conselho convertida para porcentagem (ex: nota 7.5 = 75%) contra a média obtida no Simulado 2 (FTD2).
          </p>
        </div>
      </div>
      
      <!-- Pedagogical Cross-Analysis -->
      <div class="pedagogy-container" style="grid-template-columns: 1fr;">
        <div class="pedagogy-card">
          <div class="pedagogy-title">
            <i class="fa-solid fa-code-merge"></i>
            <span>Análise Cruzada e Diagnóstico de Alinhamento Avaliativo</span>
          </div>
          <div class="pedagogy-text" id="juntosAnalysisText" style="font-size: 0.95rem;">
            <!-- Dynamically populated -->
          </div>
        </div>
      </div>
      
    </div>

    <!-- TAB 4: RELATORIO CHC (DYNAMIC MENU) -->
    <div id="relatorioTab" class="tab-content">
      <div class="report-layout">
        <!-- Sidebar Navigation Menu -->
        <aside class="report-sidebar">
          <ul class="report-menu">
            <li class="report-menu-item active" onclick="switchReportSection(1)">
              <span class="num">1</span> Análise Comparativa
            </li>
            <li class="report-menu-item" onclick="switchReportSection(2)">
              <span class="num">2</span> Distribuição por Faixas
            </li>
            <li class="report-menu-item" onclick="switchReportSection(3)">
              <span class="num">3</span> Relação Interno/Externo
            </li>
            <li class="report-menu-item" onclick="switchReportSection(4)">
              <span class="num">4</span> Atenção & Potencialidades
            </li>
            <li class="report-menu-item" onclick="switchReportSection(5)">
              <span class="num">5</span> Justificativas Observadas
            </li>
            <li class="report-menu-item" onclick="switchReportSection(6)">
              <span class="num">6</span> Estratégias de Intervenção
            </li>
            <li class="report-menu-item" onclick="switchReportSection(7)">
              <span class="num">7</span> Retomada Compromissos 2025
            </li>
            <li class="report-menu-item" onclick="switchReportSection(8)">
              <span class="num">8</span> Média Global & Rubricas
            </li>
          </ul>
        </aside>

        <!-- Main Content Panel -->
        <main class="report-main-content">
          
          <!-- Section 1 -->
          <div id="reportSection1" class="report-section active">
            <h2 class="report-title"><i class="fa-solid fa-chart-line"></i> 1. Análise Comparativa dos Resultados</h2>
            <div class="report-text">
              Como a disciplina de <strong>Computação, Humanidades e Criatividade (CHC)</strong> não conta com um simulado externo exclusivo, esta análise adota como referência comparativa (*proxy*) os indicadores externos das áreas que compõem seus eixos estruturantes:
              <ul>
                <li><strong>Matemática (FTD)</strong> como indicador para o eixo de <strong>Pensamento Computacional</strong>.</li>
                <li><strong>Língua Portuguesa (FTD)</strong> como indicador para os eixos de <strong>Cultura Digital e Mundo Digital</strong>.</li>
              </ul>
              <br>
              Abaixo estão os resultados das médias gerais das notas internas do Conselho de Classe (1º Trimestre de 2026 - antes da REC) comparadas entre as disciplinas:
            </div>
            
            <div class="table-responsive" style="margin-top: 0.5rem;">
              <table style="width: 100%; border-collapse: collapse;">
                <thead>
                  <tr style="background-color: var(--bg-color);">
                    <th style="padding: 10px;">Ano Escolar</th>
                    <th style="padding: 10px; color: var(--secondary)">Média Interna CHC</th>
                    <th style="padding: 10px;">Média Interna Matemática</th>
                    <th style="padding: 10px;">Média Interna Língua Portuguesa</th>
                    <th style="padding: 10px;">Média Interna Redação</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td style="padding: 10px; font-weight: bold;">6º Ano</td>
                    <td style="padding: 10px; font-weight: 600; color: var(--secondary)">7.01</td>
                    <td style="padding: 10px;">7.85</td>
                    <td style="padding: 10px;">6.82</td>
                    <td style="padding: 10px;">7.07</td>
                  </tr>
                  <tr>
                    <td style="padding: 10px; font-weight: bold;">7º Ano</td>
                    <td style="padding: 10px; font-weight: 600; color: var(--secondary)">7.03</td>
                    <td style="padding: 10px;">7.80</td>
                    <td style="padding: 10px;">7.48</td>
                    <td style="padding: 10px;">7.22</td>
                  </tr>
                  <tr>
                    <td style="padding: 10px; font-weight: bold;">8º Ano</td>
                    <td style="padding: 10px; font-weight: 600; color: var(--secondary)">6.97</td>
                    <td style="padding: 10px;">6.76</td>
                    <td style="padding: 10px;">6.81</td>
                    <td style="padding: 10px;">7.40</td>
                  </tr>
                </tbody>
              </table>
            </div>

            <div class="report-alert info" style="margin-top: 0.5rem;">
              <i class="fa-solid fa-circle-info"></i> &nbsp;<strong>Nota de Análise:</strong> 
              As médias gerais de CHC mantêm uma notável estabilidade em torno de <strong>7.0</strong> em todas as séries. Esse comportamento decorre da estrutura de avaliação formativa por rubricas de projeto, reduzindo a dispersão e eliminando notas extremas inferiores.
            </div>
            
            <div class="report-text">
              <strong>Comparativo de Simulados FTD (Desempenho da Escola vs. Média Nacional):</strong>
              <ul>
                <li><strong>Matemática:</strong> O 8º ano demonstra a maior margem de vantagem em relação ao país (+16.2% de acerto no FTD2). O 6º ano também supera a média nacional por +10.2% no FTD2, enquanto o 7º ano apresenta um pequeno déficit de -1.8% no FTD2.</li>
                <li><strong>Língua Portuguesa:</strong> Todas as séries apresentam margem positiva expressiva no FTD2, liderados pelo 6º ano (+18.8% de vantagem sobre a média nacional) e 8º ano (+16.2%).</li>
              </ul>
            </div>
          </div>
          
          <!-- Section 2 -->
          <div id="reportSection2" class="report-section">
            <h2 class="report-title"><i class="fa-solid fa-users-rectangle"></i> 2. Distribuição dos Estudantes por Faixa de Desempenho</h2>
            <div class="report-text">
              Distribuição absoluta e percentual de alunos no componente CHC pelas turmas e quadrantes de notas estabelecidos pela coordenação pedagógica (1º Trimestre 2026):
            </div>
            
            <div class="table-responsive">
              <table style="width: 100%; border-collapse: collapse;">
                <thead>
                  <tr style="background-color: var(--bg-color);">
                    <th style="padding: 10px;">Série e Turma</th>
                    <th style="padding: 10px; text-align: center;">&lt; 2,0</th>
                    <th style="padding: 10px; text-align: center;">2,0 a 4,9</th>
                    <th style="padding: 10px; text-align: center;">5,0 a 5,9</th>
                    <th style="padding: 10px; text-align: center;">6,0 a 6,9</th>
                    <th style="padding: 10px; text-align: center; color: var(--secondary);">7,0 a 10</th>
                    <th style="padding: 10px; text-align: center;">Total Alunos</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td style="padding: 10px; font-weight: bold;">6º Ano A</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">16 (59.3%)</td>
                    <td style="padding: 10px; text-align: center; font-weight: 600; color: var(--secondary);">11 (40.7%)</td>
                    <td style="padding: 10px; text-align: center;">27</td>
                  </tr>
                  <tr>
                    <td style="padding: 10px; font-weight: bold;">6º Ano B</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">10 (35.7%)</td>
                    <td style="padding: 10px; text-align: center; font-weight: 600; color: var(--secondary);">18 (64.3%)</td>
                    <td style="padding: 10px; text-align: center;">28</td>
                  </tr>
                  <tr>
                    <td style="padding: 10px; font-weight: bold;">6º Ano C</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">18 (64.3%)</td>
                    <td style="padding: 10px; text-align: center; font-weight: 600; color: var(--secondary);">10 (35.7%)</td>
                    <td style="padding: 10px; text-align: center;">28</td>
                  </tr>
                  <tr style="border-top: 2px solid var(--border-color);">
                    <td style="padding: 10px; font-weight: bold;">7º Ano A</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">12 (37.5%)</td>
                    <td style="padding: 10px; text-align: center; font-weight: 600; color: var(--secondary);">20 (62.5%)</td>
                    <td style="padding: 10px; text-align: center;">32</td>
                  </tr>
                  <tr>
                    <td style="padding: 10px; font-weight: bold;">7º Ano B</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">15 (50.0%)</td>
                    <td style="padding: 10px; text-align: center; font-weight: 600; color: var(--secondary);">15 (50.0%)</td>
                    <td style="padding: 10px; text-align: center;">30</td>
                  </tr>
                  <tr style="border-top: 2px solid var(--border-color);">
                    <td style="padding: 10px; font-weight: bold;">8º Ano A</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">10 (38.5%)</td>
                    <td style="padding: 10px; text-align: center; font-weight: 600; color: var(--secondary);">16 (61.5%)</td>
                    <td style="padding: 10px; text-align: center;">26</td>
                  </tr>
                  <tr>
                    <td style="padding: 10px; font-weight: bold;">8º Ano B</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">0 (0.0%)</td>
                    <td style="padding: 10px; text-align: center;">13 (54.2%)</td>
                    <td style="padding: 10px; text-align: center; font-weight: 600; color: var(--secondary);">11 (45.8%)</td>
                    <td style="padding: 10px; text-align: center;">24</td>
                  </tr>
                </tbody>
              </table>
            </div>

            <div class="report-alert info">
              <i class="fa-solid fa-circle-check"></i> &nbsp;<strong>Concentração de Excelência:</strong> 
              As turmas <strong>6ºB (64.3%)</strong>, <strong>7ºA (62.5%)</strong> e <strong>8ºA (61.5%)</strong> concentram a maior proporção de alunos na faixa de 7.0 a 10.
            </div>

            <div class="report-alert warning">
              <i class="fa-solid fa-triangle-exclamation"></i> &nbsp;<strong>Turmas Limitantes:</strong> 
              As turmas <strong>6ºC (64.3%)</strong>, <strong>7ºB (50.0%)</strong> e <strong>8ºB (54.2%)</strong> apresentam a maioria de seus alunos na faixa de 6.0 a 6.9, sugerindo menor nível de engajamento nos portfólios ou dificuldades na complexidade lógica dos projetos.
            </div>
          </div>
          
          <!-- Section 3 -->
          <div id="reportSection3" class="report-section">
            <h2 class="report-title"><i class="fa-solid fa-square-poll-vertical"></i> 3. Relação entre Desempenho Interno e Externo</h2>
            <div class="report-text">
              <strong>Correlações Estatísticas de Pearson (CHC vs. Outros Componentes):</strong>
              <ul>
                <li><strong>6º Ano:</strong> Correlação de <strong>0.2868</strong> com Matemática e <strong>0.3444</strong> com Língua Portuguesa.</li>
                <li><strong>7º Ano:</strong> Correlação moderada a forte de <strong>0.4349</strong> com Matemática e <strong>0.3194</strong> com Língua Portuguesa.</li>
                <li><strong>8º Ano:</strong> Correlação de <strong>0.2601</strong> com Matemática e <strong>0.3145</strong> com Língua Portuguesa.</li>
              </ul>
              <br>
              <strong>Interpretação Cognitiva dos Indicadores:</strong>
              <br>
              No 6º e no 8º ano, a nota em CHC possui maior dependência das habilidades verbal/leitoras (leitura, escrita reflexiva do portfólio), o que explica a maior proximidade com a média de Língua Portuguesa. 
              Por outro lado, no <strong>7º Ano</strong>, há uma clara inflexão prática com Matemática (0.4349), indicando que os projetos da série envolveram uma exigência lógica (Pensamento Computacional) de alta correspondência cognitiva com a álgebra escolar.
            </div>
            
            <div class="report-alert info">
              <i class="fa-solid fa-arrow-trend-up"></i> &nbsp;<strong>Capacidade Preditiva do Risco Acadêmico (Exemplo no 7º Ano):</strong>
              <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1.5rem; margin-top: 0.75rem;">
                <div style="background: rgba(239, 68, 68, 0.08); padding: 0.75rem; border-radius: 6px; border-left: 3px solid var(--danger);">
                  <strong>Alunos na faixa limite CHC [6.0 - 6.9] (N=27)</strong>
                  <br>• Média Interna de Matemática: <strong>7.23</strong>
                  <br>• Média Interna de Língua Port.: <strong>7.23</strong>
                </div>
                <div style="background: rgba(16, 185, 129, 0.08); padding: 0.75rem; border-radius: 6px; border-left: 3px solid var(--secondary);">
                  <strong>Alunos na faixa de excelência CHC [8.0 - 10] (N=11)</strong>
                  <br>• Média Interna de Matemática: <strong>8.75</strong>
                  <br>• Média Interna de Língua Port.: <strong>8.23</strong>
                </div>
              </div>
              <br>
              <strong>Conclusão:</strong> A queda de rendimento do aluno em CHC para a faixa limite (6.0 - 6.9) é um <strong>preditor precoce</strong> de dificuldades gerais em raciocínio abstrato ou atitude de estudo, manifestando-se antes mesmo das notas formais de provas tradicionais declinarem.
            </div>
          </div>
          
          <!-- Section 4 -->
          <div id="reportSection4" class="report-section">
            <h2 class="report-title"><i class="fa-solid fa-circle-exclamation"></i> 4. Pontos de Atenção e Potencialidades (Matriz BNCC)</h2>
            <div class="report-text">
              Cruzando as fragilidades apontadas nos simulados externos de Matemática e Língua Portuguesa com as exigências cognitivas nos eixos de CHC, identificamos:
            </div>
            
            <div class="report-alert danger">
              <strong style="color: var(--danger);"><i class="fa-solid fa-triangle-exclamation"></i> Ponto de Atenção Crítico: Álgebra e Variáveis (7º Ano - Habilidade EF07MA13)</strong>
              <br>
              • <em>Resultado no Simulado:</em> Escola obteve apenas <strong>8.5% de acertos</strong> (média nacional: 11.3%).
              <br>
              • <em>Impacto em CHC (Pensamento Computacional):</em> O conceito de "variável" (armazenar estados como pontuação, vidas e tempos) é a espinha dorsal de qualquer lógica de programação. A dificuldade matemática dos alunos com álgebra impacta diretamente sua capacidade de autonomia técnica nos projetos de computação.
            </div>

            <div class="report-alert danger">
              <strong style="color: var(--danger);"><i class="fa-solid fa-triangle-exclamation"></i> Ponto de Atenção: Frações e Proporções (6º Ano - Habilidade EF06MA10)</strong>
              <br>
              • <em>Resultado no Simulado:</em> Apenas <strong>7.7% de acerto</strong> em soma/subtração de frações.
              <br>
              • <em>Impacto em CHC:</em> Dificulta a compreensão de divisão de telas, escala de imagens e proporcionalidade de velocidades em programação em blocos.
            </div>

            <div class="report-alert info">
              <strong style="color: var(--primary);"><i class="fa-solid fa-circle-check"></i> Potencialidade Lógica: Plano Cartesiano (8º Ano - Habilidade EF08MA08)</strong>
              <br>
              • <em>Resultado no Simulado:</em> Escola atingiu <strong>69.7% de acertos</strong> (acima do geral do país de 60.1%).
              <br>
              • <em>Impacto em CHC:</em> Alunos exibem facilidade com coordenadas bidimensionais (X, Y) e movimentação espacial de objetos na tela de programação.
            </div>

            <div class="report-alert info">
              <strong style="color: var(--primary);"><i class="fa-solid fa-circle-check"></i> Potencialidade de Leitura: Análise Multissemiótica (8º Ano - Habilidade EF69LP02)</strong>
              <br>
              • <em>Resultado no Simulado:</em> **83.3% de acerto** em análise de publicidade e multimídia digital.
              <br>
              • <em>Impacto em CHC:</em> Senso crítico e facilidade na estruturação visual, diagramação de layouts e expressão criativa na web.
            </div>

            <div class="report-alert danger">
              <strong style="color: var(--danger);"><i class="fa-solid fa-triangle-exclamation"></i> Ponto de Atenção: Análise Crítica de Argumentos (8º Ano - Habilidade EF89LP04 RED)</strong>
              <br>
              • <em>Resultado no Simulado:</em> Apenas <strong>36.5% de acertos</strong> escolar.
              <br>
              • <em>Impacto em CHC (Cultura Digital):</em> Fragilidade em atividades de cidadania digital e letramento informacional (combate a fake news, discernimento de fatos vs. opiniões em mídias sociais).
            </div>
          </div>
          
          <!-- Section 5 -->
          <div id="reportSection5" class="report-section">
            <h2 class="report-title"><i class="fa-solid fa-lightbulb"></i> 5. Justificativas para os Resultados Observados</h2>
            <div class="report-text">
              Para compreender as notas observadas em CHC e as correlações com as áreas de Matemática e Linguagens, propomos as seguintes hipóteses:
              <br><br>
              <ol style="margin-left: 1.5rem; display: flex; flex-direction: column; gap: 1rem;">
                <li>
                  <strong>Impacto da Avaliação Centrada no Processo:</strong> 
                  Como CHC avalia por rubricas formativas, a nota final engloba entrega, colaboração, comportamento e esforço de pesquisa (soft skills). Isso explica o piso estável de notas $\ge 6.0$. Contudo, esse piso tende a mascarar deficiências técnicas e de raciocínio abstrato em programação de alunos que apenas cumpriram burocraticamente os prazos.
                </li>
                <li>
                  <strong>Abstracionismo Precoce:</strong> 
                  A alarmante taxa de acerto de 8.5% em variáveis em matemática (7º ano) sinaliza a dificuldade na transição do pensamento lógico concreto para o abstrato. No Scratch, muitos estudantes se limitam a arrastar blocos visuais por "tentativa e erro" ou copiar soluções prontas, sem dominar a estrutura da lógica de programação subjacente.
                </li>
                <li>
                  <strong>Disparidades de Engajamento entre as Turmas:</strong> 
                  A maior incidência de notas na faixa de 6.0 a 6.9 nas turmas <strong>6ºC</strong>, <strong>7ºB</strong> e <strong>8ºB</strong> está associada a ruídos de trabalho em grupo, dificuldades de foco e organização na utilização de Chromebooks e resistência a registrar as evidências de autoria nos diários de bordo/portfólios individuais.
                </li>
              </ol>
            </div>
          </div>
          
          <!-- Section 6 -->
          <div id="reportSection6" class="report-section">
            <h2 class="report-title"><i class="fa-solid fa-route"></i> 6. Estratégias de Intervenção e Acompanhamento</h2>
            <div class="report-text">
              Em busca de consolidar a aprendizagem lógica e apoiar a meta institucional de alcançar o <strong>Top 100</strong>, estruturamos as seguintes frentes de ação:
              <br><br>
              <strong>1. Recuperação Lógica Prática (Foco na faixa 6.0 - 6.9):</strong>
              <ul>
                <li><strong>Computação Desplugada (Unplugged Computing):</strong> Utilização de dinâmicas físicas, jogos de lógica em cartões e tabuleiros para ensinar conceitos de sequenciamento e loops antes da digitação do código.</li>
                <li><strong>Apoio Orientado aos Diários de Portfólio:</strong> Criação de roteiros padronizados e checklists simples para apoiar os alunos na redação reflexiva dos portfólios, garantindo melhor pontuação atitudinal.</li>
              </ul>
              <br>
              <strong>2. Aceleração de Alto Desempenho (Foco na faixa 8.0 - 10.0):</strong>
              <ul>
                <li><strong>Desafios de Autoria Textual:</strong> Para alunos proficientes em blocos visuais, oferecer a transição para pequenos desafios de escrita de código (como HTML/CSS básico ou linguagem Python).</li>
                <li><strong>Monitoria Cooperativa:</strong> Incentivar os estudantes de alto desempenho a atuarem como mentores de programação para apoiar colegas de grupos com notas limites.</li>
              </ul>
              <br>
              <strong>3. Integração Curricular Direta:</strong>
              <ul>
                <li><strong>Reinvenção de Variáveis (7º Ano - CHC + Matemática):</strong> Desenvolvimento de um projeto integrado no Scratch onde o aluno construa uma calculadora interativa ou simulador de equações algébricas, materializando visualmente o conceito matemático de variável.</li>
                <li><strong>Letramento Crítico de Mídias (8º Ano - CHC + Português):</strong> Oficina de fact-checking e cidadania digital, integrando a habilidade de analisar opiniões em textos (EF89LP04) com o estudo de algoritmos de recomendação de redes sociais.</li>
              </ul>
            </div>
          </div>
          
          <!-- Section 7 -->
          <div id="reportSection7" class="report-section">
            <h2 class="report-title"><i class="fa-solid fa-clock-rotate-left"></i> 7. Retomada dos Compromissos de 2025</h2>
            <div class="report-text">
              Balanço das metas planejadas no encerramento de 2025 para a qualificação pedagógica em 2026:
              <br><br>
              <strong>Avanços Alcançados:</strong>
              <ul>
                <li>Consolidação do uso de rubricas em todas as avaliações de projetos práticos.</li>
                <li>Organização regular do portfólio digital individual no Ensino Fundamental II.</li>
              </ul>
              <br>
              <strong>Ações Implementadas:</strong>
              <ul>
                <li>Uso sistemático dos Chromebooks e plataformas de computação visual criativa.</li>
                <li>Definição clara e antecipada de metas atitudinais para as atividades em grupo.</li>
              </ul>
              <br>
              <strong>Pontos Pendentes & Necessidades de Ajustes para o 2º Semestre/2026:</strong>
              <ul>
                <li>Calibrar a avaliação atitudinal para evitar que o esforço mascare completamente a incompreensão dos conceitos de lógica.</li>
                <li>Implementar agendas conjuntas de planejamento entre os docentes de CHC, Matemática e Língua Portuguesa para sincronização conceitual.</li>
              </ul>
            </div>
          </div>
          
          <!-- Section 8 -->
          <div id="reportSection8" class="report-section">
            <h2 class="report-title"><i class="fa-solid fa-list-check"></i> 8. Média Global e Uso de Rubricas</h2>
            <div class="report-text">
              <strong>Reflexão sobre as Rubricas de Avaliação em CHC:</strong>
              <br><br>
              As rubricas desempenham um papel pedagógico fundamental ao deslocar a nota de uma verificação mecânica de conteúdos para a avaliação formativa do percurso de aprendizagem. No entanto, os dados revelam impactos estruturais importantes:
              <br><br>
              <strong>1. O Mascaramento de Gaps Lógicos:</strong>
              <br>
              Como as rubricas englobam o engajamento atitudinal, o diário de bordo e o trabalho coletivo, o aluno "esforçado", mesmo com dificuldades profundas de abstração algorítmica, atinge com facilidade notas de aprovação no Conselho. Isso cria o "piso de 6.0" observado nos dados, mas impede a detecção prévia de sérias dificuldades de raciocínio.
              <br><br>
              <strong>2. Encaminhamentos Propostos para Ajuste:</strong>
              <ul>
                <li><strong>Desacoplamento de Critérios:</strong> Dividir o peso da avaliação em dois eixos explícitos de 50%: <em>"Processo e Atitude"</em> (colaboração, portfólio) e <em>"Rigor Cognitivo e Autoria"</em> (lógica do script, complexidade algorítmica).</li>
                <li><strong>Implementação de Validações Individuais Rápidas:</strong> Realizar pequenas autoavaliações ou questionários lógicos individuais (ex: explicar em 1 parágrafo como funciona a condicional do seu código) para atestar a autoria individual, mesmo em projetos coletivos.</li>
                <li><strong>Rubrica de Progressão Lógica:</strong> Mapear nas rubricas de autoria a complexidade técnica alcançada (Ex: Nível 1 - Script linear simples; Nível 2 - Uso de condicionais se/então; Nível 3 - Uso de loops e variáveis controladas).</li>
              </ul>
            </div>
          </div>

        </main>
      </div>
    </div>

  </div>

  <footer style="background-color: var(--card-bg); border-top: 1px solid var(--border-color); padding: 1.5rem; text-align: center; font-size: 0.8rem; color: var(--text-muted); margin-top: 3rem;">
    <p>Dashboard Desenvolvido para Análise do Conselho de Classe e Simulados FTD 2025/2026.</p>
    <p style="margin-top: 0.25rem;">Dados Anonimizados em Conformidade com as Diretrizes da Escola.</p>
  </footer>

  <script>
    // Embedded Data
    const DB = ${JSON.stringify(fullData)};
    
    // State variables
    let currentTab = 'conselho';
    let currentSeries = '7ANO';
    let currentSort = { column: 'id', asc: true };
    
    // Chart References
    let scatterChart = null;
    let barsChart = null;
    let simuCompareChart = null;
    let juntosChart = null;
    
    // On Document Load
    document.addEventListener("DOMContentLoaded", () => {
      // Sync Filters with state
      document.getElementById('seriesFilter').value = currentSeries;
      
      // Setup dynamic performance filter labels first
      updatePerformanceDropdown();
      
      // Load current series
      onSeriesChange();
      
      // Setup Chart defaults for Theme
      setupChartDefaults();
    });
    
    function setupChartDefaults() {
      Chart.defaults.color = getComputedStyle(document.documentElement).getPropertyValue('--text-muted').trim();
      Chart.defaults.borderColor = getComputedStyle(document.documentElement).getPropertyValue('--border-color').trim();
      Chart.defaults.font.family = "'Outfit', sans-serif";
    }
    
    function toggleTheme() {
      const isLight = document.body.classList.toggle('light-theme');
      const icon = document.querySelector('.theme-toggle i');
      if (isLight) {
        icon.className = 'fa-solid fa-sun';
      } else {
        icon.className = 'fa-solid fa-moon';
      }
      
      // Redraw charts with new colors
      setupChartDefaults();
      renderConselhoCharts();
      if (currentTab === 'simulados') {
        renderSimulados();
      } else if (currentTab === 'juntos') {
        renderJuntos();
      }
    }
    
    function switchTab(tabName) {
      currentTab = tabName;
      document.querySelectorAll('.nav-btn').forEach(btn => btn.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(tc => tc.classList.remove('active'));
      
      // Reset filter bar display based on active tab
      document.getElementById('conselhoSpecificFilters').style.display = 'none';
      document.getElementById('simuSpecificFilters').style.display = 'none';
      document.getElementById('juntosSpecificFilters').style.display = 'none';
      document.getElementById('searchWrapper').style.display = 'flex';
      document.getElementById('mainFiltersBar').style.display = 'flex';
      
      if (tabName === 'conselho') {
        document.getElementById('btnTabConselho').classList.add('active');
        document.getElementById('conselhoTab').classList.add('active');
        document.getElementById('conselhoSpecificFilters').style.display = 'flex';
        applyFilters();
      } else if (tabName === 'simulados') {
        document.getElementById('btnTabSimulados').classList.add('active');
        document.getElementById('simuladosTab').classList.add('active');
        document.getElementById('simuSpecificFilters').style.display = 'flex';
        renderSimulados();
      } else if (tabName === 'juntos') {
        document.getElementById('btnTabJuntos').classList.add('active');
        document.getElementById('juntosTab').classList.add('active');
        document.getElementById('juntosSpecificFilters').style.display = 'flex';
        document.getElementById('searchWrapper').style.display = 'none'; // No search table on integrated view
        renderJuntos();
      } else if (tabName === 'relatorio') {
        document.getElementById('btnTabRelatorio').classList.add('active');
        document.getElementById('relatorioTab').classList.add('active');
        document.getElementById('mainFiltersBar').style.display = 'none';
        document.getElementById('searchWrapper').style.display = 'none';
      }
    }
    
    function switchReportSection(sectionNum) {
      // Deactivate all menu items
      document.querySelectorAll('.report-menu-item').forEach(item => {
        item.classList.remove('active');
      });
      // Activate selected menu item
      const menuItems = document.querySelectorAll('.report-menu-item');
      if (menuItems[sectionNum - 1]) {
        menuItems[sectionNum - 1].classList.add('active');
      }
      
      // Hide all report sections
      document.querySelectorAll('.report-section').forEach(sec => {
        sec.classList.remove('active');
      });
      // Show selected report section
      const selectedSection = document.getElementById(\`reportSection\${sectionNum}\`);
      if (selectedSection) {
        selectedSection.classList.add(\`active\`);
      }
    }
    
    function resetFilters() {
      document.getElementById('seriesFilter').value = '7ANO';
      document.getElementById('turmaFilter').value = 'ALL';
      document.getElementById('axisXFilter').value = 'chc';
      document.getElementById('axisYFilter').value = 'mat';
      document.getElementById('tableSearch').value = '';
      
      updatePerformanceDropdown();
      document.getElementById('performanceFilter').value = 'ALL';
      
      if (currentTab === 'simulados') {
        document.getElementById('simuSubjectFilter').value = 'MAT';
        document.getElementById('simuTypeFilter').value = 'FTD2';
      } else if (currentTab === 'juntos') {
        document.getElementById('juntosSubjectFilter').value = 'MAT';
      }
      
      onSeriesChange();
    }
    
    function onSeriesChange() {
      currentSeries = document.getElementById('seriesFilter').value;
      
      // Update Turma Dropdown for Consiglio
      const turmaSelect = document.getElementById('turmaFilter');
      turmaSelect.innerHTML = '<option value="ALL">Todas</option>';
      
      const students = DB.conselho[currentSeries] || [];
      const turmas = [...new Set(students.map(s => s.turma))].sort();
      
      turmas.forEach(t => {
        const opt = document.createElement('option');
        opt.value = t;
        opt.textContent = t;
        turmaSelect.appendChild(opt);
      });
      
      // Apply filters and updates
      if (currentTab === 'conselho') {
        applyFilters();
      } else if (currentTab === 'simulados') {
        renderSimulados();
      } else if (currentTab === 'juntos') {
        renderJuntos();
      }
    }
    
    function onAxisChange() {
      updatePerformanceDropdown();
      applyFilters();
    }
    
    function updatePerformanceDropdown() {
      const axisX = document.getElementById('axisXFilter').value;
      const axisY = document.getElementById('axisYFilter').value;
      
      const names = {
        chc: 'CHC',
        mat: 'Matemática',
        lp: 'Língua Portuguesa',
        red: 'Redação'
      };
      
      const nameX = names[axisX];
      const nameY = names[axisY];
      
      const perfSelect = document.getElementById('performanceFilter');
      const currentVal = perfSelect.value;
      
      perfSelect.innerHTML = \`
        <option value="ALL">Todos os Alunos</option>
        <option value="X_GT_Y">\${nameX} > \${nameY}</option>
        <option value="Y_GT_X">\${nameY} > \${nameX}</option>
        <option value="X_LT_6">\${nameX} < 6.0 (Atenção)</option>
        <option value="Y_LT_6">\${nameY} < 6.0 (Atenção)</option>
      \`;
      
      // Fix labels dynamically to avoid template bugs
      perfSelect.options[1].text = \`\${nameX} > \${nameY}\`;
      perfSelect.options[2].text = \`\${nameY} > \${nameX}\`;
      
      // Restore previous filter if still valid
      if (['ALL', 'X_GT_Y', 'Y_GT_X', 'X_LT_6', 'Y_LT_6'].includes(currentVal)) {
        perfSelect.value = currentVal;
      } else {
        perfSelect.value = 'ALL';
      }
    }
    
    // Pearson Correlation function for browser
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
    
    function calculateCorrelationForSubset(students, axisX, axisY) {
      const xVals = [];
      const yVals = [];
      
      students.forEach(s => {
        if (s[axisX] !== null && s[axisY] !== null) {
          xVals.push(s[axisX]);
          yVals.push(s[axisY]);
        }
      });
      
      return pearsonCorrelation(xVals, yVals);
    }
    
    function applyFilters() {
      const turmaVal = document.getElementById('turmaFilter').value;
      const perfVal = document.getElementById('performanceFilter').value;
      const searchVal = document.getElementById('tableSearch').value.toLowerCase();
      
      const axisX = document.getElementById('axisXFilter').value;
      const axisY = document.getElementById('axisYFilter').value;
      
      const allStudents = DB.conselho[currentSeries] || [];
      
      // 1. Filter by Turma only to establish class baseline (total size)
      const classStudents = allStudents.filter(s => {
        if (turmaVal !== 'ALL' && s.turma !== turmaVal) return false;
        return true;
      });
      
      // 2. Filter by Performance and Search for the subset
      let filtered = classStudents.filter(s => {
        // Search Filter
        if (searchVal) {
          const idStr = String(s.id);
          const turmaStr = s.turma.toLowerCase();
          if (!idStr.includes(searchVal) && !turmaStr.includes(searchVal)) return false;
        }
        
        // Dynamic Performance Filter based on selected axes
        if (perfVal === 'X_GT_Y') {
          return s[axisX] !== null && s[axisY] !== null && s[axisX] > s[axisY];
        } else if (perfVal === 'Y_GT_X') {
          return s[axisX] !== null && s[axisY] !== null && s[axisY] > s[axisX];
        } else if (perfVal === 'X_LT_6') {
          return s[axisX] !== null && s[axisX] < 6.0;
        } else if (perfVal === 'Y_LT_6') {
          return s[axisY] !== null && s[axisY] < 6.0;
        }
        
        return true;
      });
      
      // Update dynamic KPIs (passing classStudents baseline)
      updateKPIs(filtered, classStudents);
      
      // Populate student table
      populateTable(filtered);
      
      // Render Charts
      renderConselhoCharts(filtered);
      
      // Update Pedagogy text
      updatePedagogy();
    }
    
    function updateKPIs(filteredStudents, classStudents) {
      const filteredCount = filteredStudents.length;
      const totalCount = classStudents.length;
      
      const perfVal = document.getElementById('performanceFilter').value;
      const searchVal = document.getElementById('tableSearch').value.trim();
      
      const kpiTotal = document.getElementById('kpiTotalStudents');
      const kpiTotalDesc = kpiTotal.parentElement.querySelector('.kpi-footer');
      
      if (perfVal !== 'ALL' || searchVal !== '') {
        kpiTotal.textContent = \`\${filteredCount} de \${totalCount}\`;
        kpiTotalDesc.textContent = "Alunos que atendem ao critério selecionado";
      } else {
        kpiTotal.textContent = totalCount;
        kpiTotalDesc.textContent = "Matriculados na série e turma selecionadas";
      }
      
      const axisX = document.getElementById('axisXFilter').value;
      const axisY = document.getElementById('axisYFilter').value;
      
      const names = {
        chc: 'CHC',
        mat: 'Matemática',
        lp: 'Língua Portuguesa',
        red: 'Redação'
      };
      
      // Update KPI Headers
      document.getElementById('kpiTitleX').textContent = \`Média \${names[axisX]}\`;
      document.getElementById('kpiTitleY').textContent = \`Média \${names[axisY]}\`;
      document.getElementById('kpiTitleCorr').textContent = \`Correlação (\${names[axisX]} vs \${names[axisY]})\`;
      
      const xVals = filteredStudents.filter(s => s[axisX] !== null).map(s => s[axisX]);
      const yVals = filteredStudents.filter(s => s[axisY] !== null).map(s => s[axisY]);
      
      const mean = arr => arr.length > 0 ? (arr.reduce((a, b) => a + b, 0) / arr.length).toFixed(2) : 'N/A';
      
      document.getElementById('kpiAvgX').textContent = mean(xVals);
      document.getElementById('kpiAvgY').textContent = mean(yVals);
      
      // Calculate dynamic correlation for the subset
      const bothValStudents = filteredStudents.filter(s => s[axisX] !== null && s[axisY] !== null);
      const corrVal = calculateCorrelationForSubset(filteredStudents, axisX, axisY);
      
      document.getElementById('kpiCorrelation').textContent = corrVal.toFixed(3);
      
      const corrCard = document.getElementById('kpiCorrelationCard');
      const corrDesc = document.getElementById('kpiCorrelationDesc');
      
      corrCard.className = 'kpi-card warning';
      if (bothValStudents.length < 3) {
        corrDesc.textContent = 'Dados insuficientes';
        corrCard.classList.add('neutral');
      } else if (Math.abs(corrVal) < 0.1) {
        corrDesc.textContent = 'Correlação Inexistente';
        corrCard.classList.add('neutral');
      } else if (Math.abs(corrVal) < 0.3) {
        corrDesc.textContent = 'Correlação Fraca';
      } else if (Math.abs(corrVal) < 0.5) {
        corrDesc.textContent = 'Correlação Moderada';
        corrCard.className = 'kpi-card secondary'; // green
      } else {
        corrDesc.textContent = 'Correlação Forte';
        corrCard.className = 'kpi-card'; // blue
      }
    }
    
    function populateTable(students) {
      const col = currentSort.column;
      const asc = currentSort.asc;
      const axisX = document.getElementById('axisXFilter').value;
      const axisY = document.getElementById('axisYFilter').value;
      
      // Update table headers to show which one is the current X and Y axis difference
      const headers = document.querySelectorAll('#studentTable th');
      headers[4].innerHTML = \`Diferença (X - Y) <i class="fa-solid fa-sort"></i>\`;
      
      // Update Legend Text
      const names = {
        chc: 'CHC',
        mat: 'Matemática',
        lp: 'Língua Portuguesa',
        red: 'Redação'
      };
      document.getElementById('tableLegendText').innerHTML = \`Legenda: <span style="color: var(--secondary); font-weight: bold;">▲ \${names[axisX]} > \${names[axisY]}</span> | <span style="color: var(--primary); font-weight: bold;">▼ \${names[axisY]} > \${names[axisX]}</span>\`;
      
      students.sort((a, b) => {
        let valA = a[col];
        let valB = b[col];
        
        // Calculate difference on the fly based on current selected disciplines
        if (col === 'diff') {
          valA = (a[axisX] !== null && a[axisY] !== null) ? (a[axisX] - a[axisY]) : -9999;
          valB = (b[axisX] !== null && b[axisY] !== null) ? (b[axisX] - b[axisY]) : -9999;
        }
        
        if (valA === null || valA === undefined) return asc ? 1 : -1;
        if (valB === null || valB === undefined) return asc ? -1 : 1;
        
        if (typeof valA === 'string') {
          return asc ? valA.localeCompare(valB) : valB.localeCompare(valA);
        }
        
        return asc ? valA - valB : valB - valA;
      });
      
      const tbody = document.getElementById('studentTableBody');
      tbody.innerHTML = '';
      
      if (students.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="empty-state"><i class="fa-solid fa-face-frown"></i><br>Nenhum aluno encontrado para os filtros aplicados.</td></tr>';
        return;
      }
      
      students.forEach(s => {
        const tr = document.createElement('tr');
        const formatGrade = val => val !== null ? val.toFixed(1) : '<span style="color: var(--text-muted);">-</span>';
        
        const getGradeClass = val => {
          if (val === null) return '';
          if (val < 6.0) return 'grade-cell grade-danger';
          if (val >= 8.0) return 'grade-cell grade-success';
          return 'grade-cell';
        };
        
        // Difference Axis X - Axis Y
        let diffHtml = '<span style="color: var(--text-muted);">-</span>';
        if (s[axisX] !== null && s[axisY] !== null) {
          const diff = s[axisX] - s[axisY];
          const diffStr = (diff >= 0 ? '+' : '') + diff.toFixed(1);
          if (diff > 0) {
            diffHtml = \`<span class="diff-badge pos" title="\${names[axisX]} é maior em \${diffStr} pts"><i class="fa-solid fa-caret-up"></i> \${diffStr}</span>\`;
          } else if (diff < 0) {
            diffHtml = \`<span class="diff-badge neg" title="\${names[axisY]} é maior em \text{\${Math.abs(diff).toFixed(1)}} pts"><i class="fa-solid fa-caret-down"></i> \${diffStr}</span>\`;
          } else {
            diffHtml = '<span class="diff-badge zero" title="Notas iguais"><i class="fa-solid fa-minus"></i> 0.0</span>';
          }
        }
        
        diffHtml = diffHtml.replace('text{', '').replace('}', '');
        
        tr.innerHTML = \`
          <td style="font-weight: 500;">Aluno #\${s.id}</td>
          <td><span class="badge-turma">\${s.turma}</span></td>
          <td class="\${getGradeClass(s.chc)}">\${formatGrade(s.chc)}</td>
          <td class="\${getGradeClass(s.mat)}">\${formatGrade(s.mat)}</td>
          <td>\${diffHtml}</td>
          <td class="\${getGradeClass(s.lp)}">\${formatGrade(s.lp)}</td>
          <td class="\${getGradeClass(s.red)}">\${formatGrade(s.red)}</td>
        \`;
        
        tbody.appendChild(tr);
      });
    }
    
    function sortStudentTable(col) {
      if (currentSort.column === col) {
        currentSort.asc = !currentSort.asc;
      } else {
        currentSort.column = col;
        currentSort.asc = true;
      }
      
      const headers = document.querySelectorAll('#studentTable th');
      const cols = ['id', 'turma', 'chc', 'mat', 'diff', 'lp', 'red'];
      const index = cols.indexOf(col);
      
      headers.forEach((h, idx) => {
        let text = h.textContent.replace(/[▼▲]/g, '').trim();
        if (idx === index) {
          h.innerHTML = \`\${h.childNodes[0].textContent} <i class="fa-solid fa-sort-\${currentSort.asc ? 'up' : 'down'}"></i>\`;
        } else {
          h.innerHTML = \`\${h.childNodes[0].textContent} <i class="fa-solid fa-sort"></i>\`;
        }
      });
      
      applyFilters();
    }
    
    function renderConselhoCharts(filteredStudents) {
      const students = filteredStudents || DB.conselho[currentSeries] || [];
      const stats = DB.stats[currentSeries] || {};
      
      const axisX = document.getElementById('axisXFilter').value;
      const axisY = document.getElementById('axisYFilter').value;
      
      const names = {
        chc: 'CHC',
        mat: 'Matemática',
        lp: 'Língua Portuguesa',
        red: 'Redação'
      };
      
      document.getElementById('scatterChartTitle').textContent = \`Correlação: \${names[axisX]} vs \${names[axisY]}\`;
      
      if (scatterChart) scatterChart.destroy();
      if (barsChart) barsChart.destroy();
      
      const themeColors = {
        primary: getComputedStyle(document.documentElement).getPropertyValue('--primary').trim(),
        secondary: getComputedStyle(document.documentElement).getPropertyValue('--secondary').trim(),
        textMain: getComputedStyle(document.documentElement).getPropertyValue('--text-main').trim(),
        border: getComputedStyle(document.documentElement).getPropertyValue('--border-color').trim(),
        danger: getComputedStyle(document.documentElement).getPropertyValue('--danger').trim(),
        warning: getComputedStyle(document.documentElement).getPropertyValue('--warning').trim()
      };
      
      const scatterPoints = students
        .filter(s => s[axisX] !== null && s[axisY] !== null)
        .map(s => ({ x: s[axisX], y: s[axisY], studentId: s.id, turma: s.turma }));
        
      const ctxScatter = document.getElementById('scatterChart').getContext('2d');
      scatterChart = new Chart(ctxScatter, {
        type: 'scatter',
        data: {
          datasets: [
            {
              label: 'Alunos',
              data: scatterPoints,
              backgroundColor: themeColors.primary,
              borderColor: 'rgba(255, 255, 255, 0.4)',
              borderWidth: 1.5,
              pointRadius: 6,
              pointHoverRadius: 8,
            },
            {
              label: 'Igualdade',
              data: [{x: 0, y: 0}, {x: 10, y: 10}],
              type: 'line',
              borderColor: themeColors.warning,
              borderWidth: 1.5,
              borderDash: [5, 5],
              fill: false,
              pointRadius: 0,
              pointHoverRadius: 0,
              showLine: true
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                label: function(context) {
                  if (context.datasetIndex === 1) return 'Linha de Notas Iguais';
                  const p = context.raw;
                  return \`Aluno #\${p.studentId} (\${p.turma}) | \text{\${names[axisX]}}: \${p.x.toFixed(1)} • \text{\${names[axisY]}}: \${p.y.toFixed(1)}\`;
                }
              }
            }
          },
          scales: {
            x: {
              title: {
                display: true,
                text: names[axisX],
                color: themeColors.textMain,
                font: { weight: 'bold' }
              },
              min: 2,
              max: 10,
              grid: { color: themeColors.border }
            },
            y: {
              title: {
                display: true,
                text: names[axisY],
                color: themeColors.textMain,
                font: { weight: 'bold' }
              },
              min: 2,
              max: 10,
              grid: { color: themeColors.border }
            }
          }
        }
      });
      
      scatterChart.options.plugins.tooltip.callbacks.label = function(context) {
        if (context.datasetIndex === 1) return 'Linha de Notas Iguais';
        const p = context.raw;
        return 'Aluno #' + p.studentId + ' (' + p.turma + ') | ' + names[axisX] + ': ' + p.x.toFixed(1) + ' • ' + names[axisY] + ': ' + p.y.toFixed(1);
      };
      
      const ctxBars = document.getElementById('barsChart').getContext('2d');
      barsChart = new Chart(ctxBars, {
        type: 'bar',
        data: {
          labels: ['CHC', 'Matemática', 'Língua Port.', 'Redação'],
          datasets: [{
            data: [stats.chcAvg, stats.matAvg, stats.lpAvg, stats.redAvg],
            backgroundColor: [
              themeColors.secondary,
              themeColors.primary,
              '#8b5cf6', // purple
              '#ec4899'  // pink
            ],
            borderRadius: 6,
            borderWidth: 0,
            barThickness: 40
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                label: function(context) {
                  return \` Média Geral: \${context.raw.toFixed(2)}\`;
                }
              }
            }
          },
          scales: {
            y: {
              min: 0,
              max: 10,
              grid: { color: themeColors.border },
              ticks: { stepSize: 2 }
            },
            x: {
              grid: { display: false }
            }
          }
        }
      });
    }
    
    function updatePedagogy() {
      const axisX = document.getElementById('axisXFilter').value;
      const axisY = document.getElementById('axisYFilter').value;
      const stats = DB.stats[currentSeries] || {};
      
      const names = {
        chc: 'CHC',
        mat: 'Matemática',
        lp: 'Língua Portuguesa',
        red: 'Redação'
      };
      
      const pedTextDiv = document.getElementById('seriesPedagogicText');
      const insightsList = document.getElementById('seriesInsightsList');
      
      if (axisX === 'chc' && axisY === 'mat') {
        let text = '';
        let insights = [];
        
        if (currentSeries === '6ANO') {
          text = \`
            Na análise do <strong>6º Ano</strong>, observamos que a média em Matemática (<strong>\${stats.matAvg.toFixed(2)}</strong>) é significativamente superior à média de CHC (<strong>\${stats.chcAvg.toFixed(2)}</strong>). 
            A correlação de Pearson é de <strong>\${stats.correlation.toFixed(4)}</strong>, o que pedagogicamente é interpretado como uma <strong>correlação praticamente nula (inexistente)</strong>.
            <br><br>
            Isso indica que o desempenho individual de um aluno em CHC não serve para estimar seu desempenho em Matemática nesta série. Além disso, <strong>\${stats.matHigherCount} alunos (\${((stats.matHigherCount/stats.bothCount)*100).toFixed(1)}%)</strong> obtiveram notas superiores em Matemática comparado ao CHC. 
            Esta disparidade sugere que a transição para o Ensino Fundamental II traz desafios distintos para essas duas disciplinas nos anos iniciais do ciclo.
          \`;
          insights = [
            { icon: 'fa-triangle-exclamation', text: '<strong>Atenção ao CHC no 6º ano</strong>: A média (\${stats.chcAvg}) está próxima da média geral do conselho. Recomenda-se revisar se os critérios de avaliação estão alinhados com o grau de maturidade da série.' },
            { icon: 'fa-arrows-split-up-and-left', text: '<strong>Independência de habilidades</strong>: Como a correlação é nula, planos de reforço escolar devem ser totalmente isolados para ambas as disciplinas.' },
            { icon: 'fa-lightbulb', text: '<strong>Explorar conexões práticas</strong>: Desenvolver projetos de CHC que incorporem raciocínio lógico-matemático para auxiliar alunos com dificuldades em abstração.' }
          ];
        } else if (currentSeries === '7ANO') {
          text = \`
            Para o <strong>7º Ano</strong>, identificamos o maior coeficiente de correlação do ciclo: <strong>\${stats.correlation.toFixed(4)}</strong>, o que representa uma <strong>correlação moderada positiva</strong>. 
            Este dado aponta que os alunos que demonstram alta performance em CHC também tendem a se destacar em Matemática, e vice-versa.
            <br><br>
            Apesar disso, a tendência de notas maiores em Matemática se mantém, com <strong>\${stats.matHigherCount} alunos (\${((stats.matHigherCount/stats.bothCount)*100).toFixed(1)}%)</strong> superando o rendimento em CHC. 
            Esse alinhamento no desempenho geral sugere um ganho na maturidade cognitiva e uma possível integração orgânica das habilidades de lógica e pensamento criativo nesta faixa etária.
          \`;
          insights = [
            { icon: 'fa-circle-check', text: '<strong>Aproveitamento da Correlação Moderada</strong>: Alunos sob risco acadêmico em Matemática podem ser apoiados através de ferramentas lúdico-tecnológicas do currículo de CHC.' },
            { icon: 'fa-users-gear', text: '<strong>Ação conjunta docente</strong>: Os professores de Matemática e CHC podem planejar metodologias de ensino conjuntas, visto que a performance dos alunos caminha de forma paralela.' },
            { icon: 'fa-brain', text: '<strong>Fortalecer Lógica Algorítmica</strong>: O raciocínio lógico compartilhado indica que projetos de programação escolar aplicados em CHC terão forte eco na proficiência matemática desta série.' }
          ];
        } else if (currentSeries === '8ANO') {
          text = \`
            No <strong>8º Ano</strong>, ocorre um fenômeno muito interessante: <strong>inversão do padrão das notas</strong>. A média em CHC (<strong>\${stats.chcAvg.toFixed(2)}</strong>) supera a de Matemática (<strong>\${stats.matAvg.toFixed(2)}</strong>). 
            <strong>\${stats.chcHigherCount} alunos (\${((stats.chcHigherCount/stats.bothCount)*100).toFixed(1)}%)</strong> obtiveram notas superiores em CHC.
            <br><br>
            A correlação de Pearson é <strong>\${stats.correlation.toFixed(4)}</strong>, indicando uma <strong>correlação fraca positiva</strong>. 
            Os dados sugerem que no 8º ano as exigências de abstração matemática se tornam consideravelmente mais rigorosas para os alunos, ao passo que as disciplinas voltadas à computação e criatividade oferecem canais onde os alunos conseguem exprimir melhor seu aprendizado.
          \`;
          insights = [
            { icon: 'fa-chart-line-down', text: '<strong>Alerta de Queda em Matemática</strong>: A média de Matemática recuou para \${stats.matAvg}. Há necessidade de intervenção pedagógica focada nos conteúdos estruturantes do 1º trimestre.' },
            { icon: 'fa-hands-holding-child', text: '<strong>Alavancagem de Talentos</strong>: Aproveitar o engajamento e as notas excelentes em CHC (\${stats.chcAvg}) para contextualizar problemas complexos de matemática.' },
            { icon: 'fa-clipboard-question', text: '<strong>Revisão de Conteúdos BNCC</strong>: Cruzar as habilidades do simulado de Matemática do 8º ano com os resultados do conselho para identificar tópicos exatos de defasagem.' }
          ];
        }
        
        pedTextDiv.innerHTML = text;
        insightsList.innerHTML = '';
        insights.forEach(ins => {
          const li = document.createElement('li');
          li.className = 'insight-item';
          li.innerHTML = \`
            <i class="fa-solid \${ins.icon}"></i>
            <span class="insight-desc">\${ins.text}</span>
          \`;
          insightsList.appendChild(li);
        });
      } else {
        // Dynamic analysis for other combinations
        const allStudents = DB.conselho[currentSeries] || [];
        const xVals = allStudents.filter(s => s[axisX] !== null).map(s => s[axisX]);
        const yVals = allStudents.filter(s => s[axisY] !== null).map(s => s[axisY]);
        
        const meanX = xVals.length > 0 ? (xVals.reduce((a,b) => a+b, 0) / xVals.length) : 0;
        const meanY = yVals.length > 0 ? (yVals.reduce((a,b) => a+b, 0) / yVals.length) : 0;
        
        const corrVal = calculateCorrelationForSubset(allStudents, axisX, axisY);
        
        let relationDesc = '';
        if (Math.abs(corrVal) < 0.1) relationDesc = 'inexistente';
        else if (Math.abs(corrVal) < 0.3) relationDesc = 'fraca';
        else if (Math.abs(corrVal) < 0.5) relationDesc = 'moderada';
        else relationDesc = 'forte';
        
        const signDesc = corrVal >= 0 ? 'positiva' : 'negativa';
        
        let compareText = '';
        const bothCount = allStudents.filter(s => s[axisX] !== null && s[axisY] !== null).length;
        let xGtYCount = 0;
        let yGtXCount = 0;
        allStudents.forEach(s => {
          if (s[axisX] !== null && s[axisY] !== null) {
            if (s[axisX] > s[axisY]) xGtYCount++;
            else if (s[axisY] > s[axisX]) yGtXCount++;
          }
        });
        
        if (meanX > meanY) {
          compareText = \`A média de <strong>\${names[axisX]}</strong> (<strong>\${meanX.toFixed(2)}</strong>) é superior à de <strong>\${names[axisY]}</strong> (<strong>\${meanY.toFixed(2)}</strong>) nesta série, com <strong>\${xGtYCount} alunos (\${((xGtYCount/bothCount)*100).toFixed(1)}%)</strong> obtendo melhor rendimento em \${names[axisX]}.\`;
        } else if (meanY > meanX) {
          compareText = \`A média de <strong>\${names[axisY]}</strong> (<strong>\${meanY.toFixed(2)}</strong>) é superior à de <strong>\${names[axisX]}</strong> (<strong>\${meanX.toFixed(2)}</strong>) nesta série, com <strong>\${yGtXCount} alunos (\text{\${((yGtXCount/bothCount)*100).toFixed(1)}}%)</strong> obtendo melhor rendimento em \${names[axisY]}.\`;
        } else {
          compareText = \`Ambas as disciplinas possuem médias iguais (<strong>\${meanX.toFixed(2)}</strong>) nesta série.\`;
        }
        
        compareText = compareText.replace('text{', '').replace('}', '');
        
        pedTextDiv.innerHTML = \`
          Na análise do <strong>\${currentSeries === '6ANO' ? '6º Ano' : currentSeries === '7ANO' ? '7º Ano' : '8º Ano'}</strong> relacionando <strong>\${names[axisX]}</strong> e <strong>\${names[axisY]}</strong>, encontramos um coeficiente de correlação de Pearson de <strong>\${corrVal.toFixed(4)}</strong>. 
          Isso indica uma <strong>correlação \${relationDesc} \${signDesc}</strong> entre as duas matérias.
          <br><br>
          \${compareText}
          <br><br>
          Estes cruzamentos dinâmicos permitem identificar se as competências cognitivas de uma matéria (como a leitura/interpretação em Língua Portuguesa ou escrita em Redação) auxiliam diretamente no desempenho de outras áreas, ou se as avaliações seguem trajetórias pedagógicas independentes.
        \`;
        
        // Dynamic Insights
        insightsList.innerHTML = '';
        
        let insight1 = '';
        if (relationDesc === 'inexistente' || relationDesc === 'fraca') {
          insight1 = \`<strong>Habilidades Distintas</strong>: A correlação \${relationDesc} indica que o desempenho em \${names[axisX]} não explica o desempenho em \${names[axisY]}. Intervenções pedagógicas devem ser focadas em dificuldades específicas de cada matéria.\`;
        } else {
          insight1 = \`<strong>Sinergia Cognitiva</strong>: A correlação \${relationDesc} indica uma forte ligação no rendimento. Alunos com dificuldades em \${names[axisY]} podem se beneficiar de abordagens que explorem a lógica de \${names[axisX]}.\`;
        }
        
        let insight2 = \`<strong>Comparativo de Médias</strong>: A diferença entre as médias das duas matérias é de <strong>\${Math.abs(meanX - meanY).toFixed(2)}</strong> pontos. Equipes docentes podem discutir se essa distância decorre de critérios avaliativos ou complexidade de conteúdo.\`;
        let insight3 = \`<strong>Alunos sob Atenção</strong>: Filtre a tabela acima por <em>"\${names[axisX]} < 6.0"</em> ou <em>"\${names[axisY]} < 6.0"</em> para mapear os estudantes que necessitam de apoio pedagógico imediato nestas áreas.\`;
        
        [insight1, insight2, insight3].forEach((text, i) => {
          const li = document.createElement('li');
          li.className = 'insight-item';
          const icons = ['fa-chart-line', 'fa-scale-balanced', 'fa-hand-holding-child'];
          li.innerHTML = \`
            <i class="fa-solid \${icons[i]}"></i>
            <span class="insight-desc">\${text}</span>
          \`;
          insightsList.appendChild(li);
        });
      }
    }
    
    // TAB 2: SIMULADOS RENDERING
    function renderSimulados() {
      const subject = document.getElementById('simuSubjectFilter').value;
      const type = document.getElementById('simuTypeFilter').value;
      
      const skills = (DB.simulados[subject] && DB.simulados[subject][currentSeries] && DB.simulados[subject][currentSeries][type]) || [];
      
      const themeColors = {
        primary: getComputedStyle(document.documentElement).getPropertyValue('--primary').trim(),
        secondary: getComputedStyle(document.documentElement).getPropertyValue('--secondary').trim(),
        textMain: getComputedStyle(document.documentElement).getPropertyValue('--text-main').trim(),
        border: getComputedStyle(document.documentElement).getPropertyValue('--border-color').trim(),
        danger: getComputedStyle(document.documentElement).getPropertyValue('--danger').trim(),
        warning: getComputedStyle(document.documentElement).getPropertyValue('--warning').trim()
      };
      
      // Calculate general stats
      let sumEscola = 0;
      let sumGeral = 0;
      let count = 0;
      
      skills.forEach(s => {
        if (s.acertoEscola !== null && s.acertoGeral !== null) {
          sumEscola += s.acertoEscola;
          sumGeral += s.acertoGeral;
          count++;
        }
      });
      
      const schoolPct = count > 0 ? (sumEscola / count) * 100 : 0;
      const generalPct = count > 0 ? (sumGeral / count) * 100 : 0;
      const diffPct = schoolPct - generalPct;
      
      // Update KPIs
      document.getElementById('simuKpiEscola').textContent = schoolPct.toFixed(1) + '%';
      document.getElementById('simuKpiGeral').textContent = generalPct.toFixed(1) + '%';
      
      const diffValText = (diffPct >= 0 ? '+' : '') + diffPct.toFixed(1) + '%';
      document.getElementById('simuKpiDiff').textContent = diffValText;
      
      const diffCard = document.getElementById('simuKpiDiffCard');
      const diffIcon = document.getElementById('simuKpiDiffIcon');
      
      if (diffPct > 0) {
        diffCard.className = 'kpi-card secondary'; // green
        diffIcon.className = 'fa-solid fa-arrow-trend-up';
        diffIcon.style.color = 'var(--secondary)';
      } else if (diffPct < 0) {
        diffCard.className = 'kpi-card danger'; // red
        diffIcon.className = 'fa-solid fa-arrow-trend-down';
        diffIcon.style.color = 'var(--danger)';
      } else {
        diffCard.className = 'kpi-card'; // neutral
        diffIcon.className = 'fa-solid fa-scale-balanced';
        diffIcon.style.color = 'var(--text-muted)';
      }
      
      // Populate Averages list
      const avgList = document.getElementById('simuAveragesList');
      avgList.innerHTML = '';
      
      const skillsContainer = document.getElementById('simuSkillsContainer');
      skillsContainer.innerHTML = '';
      
      if (skills.length === 0) {
        avgList.innerHTML = '<div class="empty-state">Sem dados para este simulado.</div>';
        skillsContainer.innerHTML = '<div class="empty-state"><i class="fa-solid fa-triangle-exclamation"></i><br>Nenhum dado encontrado para os filtros selecionados.</div>';
        if (simuCompareChart) simuCompareChart.destroy();
        return;
      }
      
      skills.forEach(s => {
        const item = document.createElement('div');
        item.style.display = 'flex';
        item.style.justifyContent = 'space-between';
        item.style.alignItems = 'center';
        item.style.fontSize = '0.85rem';
        item.style.padding = '0.4rem 0';
        item.style.borderBottom = '1px solid var(--border-color)';
        
        const badgeClass = s.comparativo >= 0 ? 'grade-success' : 'grade-danger';
        const diffStr = (s.comparativo >= 0 ? '+' : '') + (s.comparativo * 100).toFixed(1) + '%';
        
        item.innerHTML = \`
          <span style="font-weight: 500;">\${s.bncc}</span>
          <span>Escola: <strong>\${(s.acertoEscola * 100).toFixed(0)}%</strong> vs Geral: \${(s.acertoGeral * 100).toFixed(0)}% 
            <span class="\text{\${badgeClass}}" style="margin-left: 0.5rem; font-weight: bold;">(\${diffStr})</span>
          </span>
        \`;
        
        item.innerHTML = item.innerHTML.replace('text{', '').replace('}', '');
        
        avgList.appendChild(item);
        
        // Detailed skill cards
        const card = document.createElement('div');
        card.className = 'simu-skill-row';
        
        const barEscolaVal = s.acertoEscola * 100;
        const barGeralVal = s.acertoGeral * 100;
        
        let actionsHtml = '';
        if (s.plano) {
          actionsHtml += \`<div class="simu-card-action"><strong><i class="fa-solid fa-route"></i> Plano de Ação:</strong> \${s.plano}</div>\`;
        }
        if (s.justificativa) {
          actionsHtml += \`<div class="simu-card-just"><strong><i class="fa-solid fa-comment-dots"></i> Justificativa Docente:</strong> \${s.justificativa}</div>\`;
        }
        
        card.innerHTML = \`
          <div class="simu-skill-header">
            <div>
              <span class="simu-skill-code">\${s.bncc}</span>
              <strong style="margin-left: 0.5rem; font-size: 0.95rem;">\${s.desc.substring(0, 80)}\${s.desc.length > 80 ? '...' : ''}</strong>
            </div>
            <div class="simu-skill-comparison">
              <span>Escola: <strong class="\${s.comparativo >= 0 ? 'grade-success' : 'grade-danger'}">\${barEscolaVal.toFixed(1)}%</strong></span>
              <span>Geral: <strong>\${barGeralVal.toFixed(1)}%</strong></span>
            </div>
          </div>
          <p style="font-size: 0.85rem; color: var(--text-muted); margin-bottom: 0.5rem;" title="\${s.desc}">\${s.desc}</p>
          <div class="bar-comparison-outer">
            <div class="bar-comparison-inner" style="width: \${barEscolaVal}%; background-color: \text{\${s.comparativo >= 0 ? 'var(--secondary)' : 'var(--primary)'}}"></div>
            <div class="bar-comparison-national" style="left: \${barGeralVal}%" title="Média Nacional: \${barGeralVal.toFixed(1)}%"></div>
          </div>
          \${actionsHtml ? \`<div class="simu-card-details">\${actionsHtml}</div>\` : ''}
        \`;
        
        card.innerHTML = card.innerHTML.replace('text{', '').replace('}', '');
        
        skillsContainer.appendChild(card);
      });
      
      // Render Simulator Comparison Chart
      const labels = skills.map(s => s.bncc);
      const escolaData = skills.map(s => s.acertoEscola * 100);
      const geralData = skills.map(s => s.acertoGeral * 100);
      
      if (simuCompareChart) simuCompareChart.destroy();
      
      const ctxSimu = document.getElementById('simuCompareChart').getContext('2d');
      simuCompareChart = new Chart(ctxSimu, {
        type: 'bar',
        data: {
          labels: labels,
          datasets: [
            {
              label: 'Nossa Escola',
              data: escolaData,
              backgroundColor: themeColors.primary,
              borderRadius: 4,
              barPercentage: 0.6,
              categoryPercentage: 0.8
            },
            {
              label: 'Média Geral (Nacional)',
              data: geralData,
              backgroundColor: themeColors.border,
              borderRadius: 4,
              barPercentage: 0.6,
              categoryPercentage: 0.8
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: {
              position: 'top',
              labels: { boxWidth: 12, padding: 15 }
            },
            tooltip: {
              callbacks: {
                label: function(context) {
                  return \` \${context.dataset.label}: \${context.raw.toFixed(1)}%\`;
                }
              }
            }
          },
          scales: {
            y: {
              min: 0,
              max: 100,
              grid: { color: themeColors.border },
              ticks: { callback: value => value + '%' }
            },
            x: {
              grid: { display: false }
            }
          }
        }
      });
    }
    
    // TAB 3: INTEGRATED ANALYSIS (JUNTOS)
    function renderJuntos() {
      const subject = document.getElementById('juntosSubjectFilter').value;
      const years = ['6ANO', '7ANO', '8ANO'];
      
      const themeColors = {
        primary: getComputedStyle(document.documentElement).getPropertyValue('--primary').trim(),
        secondary: getComputedStyle(document.documentElement).getPropertyValue('--secondary').trim(),
        purple: getComputedStyle(document.documentElement).getPropertyValue('--purple').trim(),
        textMain: getComputedStyle(document.documentElement).getPropertyValue('--text-main').trim(),
        border: getComputedStyle(document.documentElement).getPropertyValue('--border-color').trim(),
        danger: getComputedStyle(document.documentElement).getPropertyValue('--danger').trim(),
        warning: getComputedStyle(document.documentElement).getPropertyValue('--warning').trim()
      };
      
      const dataSet = {
        conselho: [],
        simu1: [],
        simu2: []
      };
      
      // Update table headers based on subject
      const thead = document.querySelector('#juntosTab table thead tr');
      if (subject === 'CHC') {
        thead.innerHTML = \`
          <th>Série</th>
          <th>Conselho (CHC)</th>
          <th>Simulado Matemática (FTD2)</th>
          <th>Simulado LP (FTD2)</th>
          <th>Gap Médio</th>
        \`;
      } else {
        thead.innerHTML = \`
          <th>Série</th>
          <th>Conselho (Interno)</th>
          <th>Simulado 1 (FTD1)</th>
          <th>Simulado 2 (FTD2)</th>
          <th>Diferença (Gap)</th>
        \`;
      }
      
      if (subject === 'CHC') {
        years.forEach(yr => {
          // Conselho CHC
          const students = DB.conselho[yr] || [];
          const grades = students.filter(s => s.chc !== null).map(s => s.chc);
          const avgConselho = grades.length > 0 ? (grades.reduce((a,b) => a+b, 0) / grades.length) * 10 : 0;
          dataSet.conselho.push(parseFloat(avgConselho.toFixed(1)));
          
          // Simulado Mat FTD2
          const matAvg = getSimuAverageForYear('MAT', yr, 'FTD2');
          dataSet.simu1.push(matAvg ? parseFloat(matAvg.toFixed(1)) : 0);
          
          // Simulado LP FTD2
          const lpAvg = getSimuAverageForYear('LP', yr, 'FTD2');
          dataSet.simu2.push(lpAvg ? parseFloat(lpAvg.toFixed(1)) : 0);
        });
      } else {
        const key = subject === 'MAT' ? 'mat' : 'lp';
        years.forEach(yr => {
          const students = DB.conselho[yr] || [];
          const grades = students.filter(s => s[key] !== null).map(s => s[key]);
          const avgConselho = grades.length > 0 ? (grades.reduce((a,b) => a+b, 0) / grades.length) * 10 : 0;
          dataSet.conselho.push(parseFloat(avgConselho.toFixed(1)));
          
          const simu1Avg = getSimuAverageForYear(subject, yr, 'FTD1');
          dataSet.simu1.push(simu1Avg ? parseFloat(simu1Avg.toFixed(1)) : 0);
          
          const simu2Avg = getSimuAverageForYear(subject, yr, 'FTD2');
          dataSet.simu2.push(simu2Avg ? parseFloat(simu2Avg.toFixed(1)) : 0);
        });
      }
      
      const seriesIndex = years.indexOf(currentSeries);
      const internalVal = dataSet.conselho[seriesIndex];
      let externalVal, gapVal;
      
      if (subject === 'CHC') {
        // Average of Mat and LP simulators
        externalVal = (dataSet.simu1[seriesIndex] + dataSet.simu2[seriesIndex]) / 2;
        gapVal = internalVal - externalVal;
        
        document.getElementById('juntosKpiInterna').textContent = (internalVal / 10).toFixed(2);
        document.getElementById('juntosKpiExterna').textContent = externalVal.toFixed(1) + '%';
        document.getElementById('juntosKpiGap').textContent = (gapVal >= 0 ? '+' : '') + gapVal.toFixed(1) + '%';
        
        document.getElementById('juntosKpiInternaLabel').textContent = \`Média Conselho (CHC) - \${currentSeries === '6ANO' ? '6º' : '7º'} Ano\`;
        document.getElementById('juntosKpiExternaLabel').textContent = \`Média Simulados (Mat + LP) - \${currentSeries === '6ANO' ? '6º' : '7º'} Ano\`;
        
        // Fix series label dynamic text
        const serStr = currentSeries === '6ANO' ? '6º' : currentSeries === '7ANO' ? '7º' : '8º';
        document.getElementById('juntosKpiInternaLabel').textContent = \`Média Conselho (CHC) - \${serStr} Ano\`;
        document.getElementById('juntosKpiExternaLabel').textContent = \`Média Simulados (Mat + LP) - \${serStr} Ano\`;
      } else {
        const names = { 'MAT': 'Matemática', 'LP': 'Língua Portuguesa' };
        externalVal = dataSet.simu2[seriesIndex];
        gapVal = internalVal - externalVal;
        
        document.getElementById('juntosKpiInterna').textContent = (internalVal / 10).toFixed(2);
        document.getElementById('juntosKpiExterna').textContent = externalVal.toFixed(1) + '%';
        document.getElementById('juntosKpiGap').textContent = (gapVal >= 0 ? '+' : '') + gapVal.toFixed(1) + '%';
        
        const serStr = currentSeries === '6ANO' ? '6º' : currentSeries === '7ANO' ? '7º' : '8º';
        document.getElementById('juntosKpiInternaLabel').textContent = \`Média Conselho (\${names[subject]}) - \${serStr} Ano\`;
        document.getElementById('juntosKpiExternaLabel').textContent = \`Simulado FTD2 (\${names[subject]}) - \${serStr} Ano\`;
      }
      
      const gapCard = document.getElementById('juntosKpiGargaloCard');
      if (gapVal > 30) {
        gapCard.className = 'kpi-card danger';
      } else if (gapVal > 15) {
        gapCard.className = 'kpi-card warning';
      } else {
        gapCard.className = 'kpi-card secondary';
      }
      
      // Populate Comparison Table
      const tbody = document.getElementById('juntosTableBody');
      tbody.innerHTML = '';
      
      years.forEach((yr, idx) => {
        const tr = document.createElement('tr');
        const consGrade = (dataSet.conselho[idx] / 10).toFixed(2);
        
        let s1Text, s2Text, gapText, gap;
        if (subject === 'CHC') {
          s1Text = dataSet.simu1[idx] > 0 ? dataSet.simu1[idx].toFixed(1) + '%' : '-';
          s2Text = dataSet.simu2[idx] > 0 ? dataSet.simu2[idx].toFixed(1) + '%' : '-';
          gap = dataSet.conselho[idx] - ((dataSet.simu1[idx] + dataSet.simu2[idx]) / 2);
          gapText = (gap >= 0 ? '+' : '') + gap.toFixed(1) + '%';
        } else {
          s1Text = dataSet.simu1[idx] > 0 ? dataSet.simu1[idx].toFixed(1) + '%' : '-';
          s2Text = dataSet.simu2[idx] > 0 ? dataSet.simu2[idx].toFixed(1) + '%' : '-';
          gap = dataSet.conselho[idx] - dataSet.simu2[idx];
          gapText = (gap >= 0 ? '+' : '') + gap.toFixed(1) + '%';
        }
        
        const gapClass = gap > 30 ? 'grade-danger' : gap > 15 ? 'grade-warning' : 'grade-success';
        
        tr.className = yr === currentSeries ? 'card-bg-hover' : '';
        tr.style.cursor = 'pointer';
        tr.onclick = () => {
          document.getElementById('seriesFilter').value = yr;
          onSeriesChange();
        };
        
        tr.innerHTML = \`
          <td style="font-weight: bold;">\${yr === '6ANO' ? '6º Ano' : yr === '7ANO' ? '7º Ano' : '8º Ano'} \${yr === currentSeries ? '<span style="color:var(--primary); font-size:0.75rem;">(Ativo)</span>' : ''}</td>
          <td style="font-weight: 600;">\${consGrade} <span style="font-size:0.75rem; color:var(--text-muted);">(\${dataSet.conselho[idx].toFixed(0)}%)</span></td>
          <td>\${s1Text}</td>
          <td>\${s2Text}</td>
          <td class="\${gapClass}" style="font-weight: bold;">\${gapText}</td>
        \`;
        tbody.appendChild(tr);
      });
      
      // Render Chart
      if (juntosChart) juntosChart.destroy();
      
      let datasets = [];
      if (subject === 'CHC') {
        document.getElementById('juntosChartTitle').textContent = \`Comparativo: Conselho CHC vs Simulados FTD2 (Matemática e LP)\`;
        datasets = [
          {
            label: 'Conselho (CHC %)',
            data: dataSet.conselho,
            backgroundColor: themeColors.secondary,
            borderRadius: 4,
            barPercentage: 0.6,
            categoryPercentage: 0.7
          },
          {
            label: 'Simulado Matemática (FTD2 %)',
            data: dataSet.simu1,
            backgroundColor: themeColors.primary,
            borderRadius: 4,
            barPercentage: 0.6,
            categoryPercentage: 0.7
          },
          {
            label: 'Simulado LP (FTD2 %)',
            data: dataSet.simu2,
            backgroundColor: themeColors.purple,
            borderRadius: 4,
            barPercentage: 0.6,
            categoryPercentage: 0.7
          }
        ];
      } else {
        const names = { 'MAT': 'Matemática', 'LP': 'Língua Portuguesa' };
        document.getElementById('juntosChartTitle').textContent = \`Comparativo (\${names[subject]}): Conselho 2026 vs Simulados 2025\`;
        datasets = [
          {
            label: 'Conselho (Interno %)',
            data: dataSet.conselho,
            backgroundColor: themeColors.primary,
            borderRadius: 4,
            barPercentage: 0.6,
            categoryPercentage: 0.7
          },
          {
            label: 'Simulado 1 (FTD1 %)',
            data: dataSet.simu1,
            backgroundColor: themeColors.purple,
            borderRadius: 4,
            barPercentage: 0.6,
            categoryPercentage: 0.7
          },
          {
            label: 'Simulado 2 (FTD2 %)',
            data: dataSet.simu2,
            backgroundColor: themeColors.secondary,
            borderRadius: 4,
            barPercentage: 0.6,
            categoryPercentage: 0.7
          }
        ];
      }
      
      const ctxJuntos = document.getElementById('juntosCompareChart').getContext('2d');
      juntosChart = new Chart(ctxJuntos, {
        type: 'bar',
        data: {
          labels: ['6º Ano', '7º Ano', '8º Ano'],
          datasets: datasets
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { position: 'top', labels: { boxWidth: 12 } },
            tooltip: {
              callbacks: {
                label: function(context) {
                  return \` \${context.dataset.label}: \${context.raw.toFixed(1)}%\`;
                }
              }
            }
          },
          scales: {
            y: {
              min: 0,
              max: 100,
              grid: { color: themeColors.border },
              ticks: { callback: value => value + '%' }
            },
            x: {
              grid: { display: false }
            }
          }
        }
      });
      
      updateJuntosAnalysisText(subject, dataSet);
    }
    
    function getSimuAverageForYear(subject, grade, type) {
      const skills = (DB.simulados[subject] && DB.simulados[subject][grade] && DB.simulados[subject][grade][type]) || [];
      let sum = 0;
      let count = 0;
      skills.forEach(s => {
        if (s.acertoEscola !== null) {
          sum += s.acertoEscola;
          count++;
        }
      });
      return count > 0 ? (sum / count) * 100 : null;
    }
    
    function updateJuntosAnalysisText(subject, dataSet) {
      const textDiv = document.getElementById('juntosAnalysisText');
      const names = {
        'MAT': 'Matemática',
        'LP': 'Língua Portuguesa',
        'CHC': 'CHC'
      };
      const name = names[subject];
      
      const gap6 = dataSet.conselho[0] - (subject === 'CHC' ? ((dataSet.simu1[0] + dataSet.simu2[0]) / 2) : dataSet.simu2[0]);
      const gap7 = dataSet.conselho[1] - (subject === 'CHC' ? ((dataSet.simu1[1] + dataSet.simu2[1]) / 2) : dataSet.simu2[1]);
      const gap8 = dataSet.conselho[2] - (subject === 'CHC' ? ((dataSet.simu1[2] + dataSet.simu2[2]) / 2) : dataSet.simu2[2]);
      
      let analysis = '';
      
      if (subject === 'MAT') {
        analysis = \`
          O comparativo integrado de <strong>Matemática</strong> revela uma <strong>discrepância (Gap) muito acentuada</strong> entre a avaliação interna da escola (Conselho) e o teste padronizado externo (FTD). 
          <br><br>
          No <strong>6º Ano</strong>, a nota média do conselho é <strong>7.67 (76.7%)</strong>, enquanto no simulado FTD2 os alunos acertaram apenas <strong>\${dataSet.simu2[0].toFixed(1)}%</strong> dos itens (Gap de <strong>+\${gap6.toFixed(1)}%</strong>). 
          No <strong>7º Ano</strong>, esta diferença se alarga ainda mais, atingindo <strong>+\${gap7.toFixed(1)}%</strong> (média de <strong>7.80</strong> no Conselho vs <strong>\${dataSet.simu2[1].toFixed(1)}%</strong> no FTD2). 
          No <strong>8º Ano</strong>, embora as notas do Conselho tenham sido mais rígidas (média <strong>6.76</strong> / 67.6%), o rendimento no simulado foi de <strong>\${dataSet.simu2[2].toFixed(1)}%</strong>, resultando em um gap de <strong>+\${gap8.toFixed(1)}%</strong>.
          <br><br>
          <strong>Interpretação Pedagógica:</strong>
          Esse distanciamento sistemático (Gaps superiores a 30%) é comum em sistemas escolares e sugere que:
          <ul style="margin: 0.75rem 0 0.75rem 1.5rem; display:flex; flex-direction:column; gap:0.5rem; color: var(--text-muted);">
            <li>A estrutura das questões dos simulados FTD exige um nível de <strong>leitura instrumental e raciocínio lógico dedutivo</strong> mais complexo do que as avaliações internas aplicadas pela escola.</li>
            <li>Critérios internos de atribuição de notas (como participação, atividades, pesquisas) podem estar elevando a nota final do conselho, mas não necessariamente refletem a proficiência em resolver problemas inéditos sob pressão de tempo.</li>
            <li>Habilidades cruciais da BNCC, como as descritas nos planos de ação da FTD (Ex: dízimas periódicas no 8º ano e múltiplos no 7º ano), apresentam gargalos graves que precisam ser trabalhados com materiais didáticos complementares.</li>
          </ul>
        \`;
      } else if (subject === 'LP') {
        analysis = \`
          O comparativo integrado de <strong>Língua Portuguesa</strong> apresenta um cenário de <strong>forte alinhamento avaliativo</strong>, com gaps consideravelmente menores do que em Matemática, demonstrando consistência entre as avaliações internas e externas.
          <br><br>
          No <strong>6º Ano</strong>, os resultados são praticamente equivalentes: a média do Conselho foi de <strong>6.59 (65.9%)</strong> contra <strong>\${dataSet.simu2[0].toFixed(1)}%</strong> de acertos no simulado FTD2 (um gap mínimo de <strong>+\${gap6.toFixed(1)}%</strong>). 
          No <strong>7º Ano</strong>, a média interna de <strong>7.48 (74.8%)</strong> superou o FTD2 que registrou <strong>\${dataSet.simu2[1].toFixed(1)}%</strong> (gap de <strong>+\${gap7.toFixed(1)}%</strong>). 
          No <strong>8º Ano</strong>, a média do conselho de <strong>6.81 (68.1%)</strong> comparada aos <strong>\${dataSet.simu2[2].toFixed(1)}%</strong> no FTD2 gerou um gap de <strong>+\${gap8.toFixed(1)}%</strong>.
          <br><br>
          <strong>Interpretação Pedagógica:</strong>
          A forte proximidade dos números em Língua Portuguesa indica que:
          <ul style="margin: 0.75rem 0 0.75rem 1.5rem; display:flex; flex-direction:column; gap:0.5rem; color: var(--text-muted);">
            <li>Os critérios internos e o formato das avaliações escolares estão <strong>bem alinhados</strong> à matriz de competências de leitura, interpretação e gramática da FTD e da BNCC.</li>
            <li>O rendimento demonstrado em sala de aula é um refletor fidedigno da capacidade de leitura instrumental que os alunos exibem em provas externas padronizadas.</li>
            <li>A proficiência em LP é mais estável. A escola pode focar os esforços de planejamento pedagógico em Matemática (onde há disparidade de critérios) e CHC.</li>
          </ul>
        \`;
      } else if (subject === 'CHC') {
        analysis = \`
          O comparativo integrado de <strong>CHC (Computação Humanidades e Criatividade)</strong> cruza a nota média interna da disciplina com os resultados dos simulados externos de <strong>Matemática e Língua Portuguesa (FTD2)</strong>.
          <br><br>
          No <strong>6º Ano</strong>, a média de CHC é de <strong>6.85 (68.5%)</strong>, o que apresenta um gap de <strong>+21.1%</strong> em relação ao simulado de Matemática (47.4%) e um gap de <strong>+5.4%</strong> em relação ao simulado de LP (63.1%).
          No <strong>7º Ano</strong>, a média de CHC é <strong>7.03 (70.3%)</strong>, com um gap de <strong>+35.3%</strong> contra o simulado de Matemática (35.0%) e um gap de <strong>+7.9%</strong> contra o simulado de LP (62.4%).
          No <strong>8º Ano</strong>, a média de CHC é <strong>6.97 (69.7%)</strong>, com um gap de <strong>+21.1%</strong> contra o simulado de Matemática (48.6%) e um gap de <strong>+9.1%</strong> contra o simulado de LP (60.6%).
          <br><br>
          <strong>Interpretação Pedagógica:</strong>
          CHC é uma disciplina interdisciplinar. Ao analisar o rendimento em relação aos testes externos FTD, observamos que:
          <ul style="margin: 0.75rem 0 0.75rem 1.5rem; display:flex; flex-direction:column; gap:0.5rem; color: var(--text-muted);">
            <li>A performance em CHC está <strong>muito mais próxima dos resultados de Língua Portuguesa (Leitura/Interpretação)</strong> do que dos resultados de Matemática (Lógica Instrumental). Isso sugere que o currículo avaliado em CHC prioriza habilidades de reflexão, leitura e argumentação criativa em detrimento de uma abordagem puramente técnica ou exata.</li>
            <li>Existe um descompasso claro em relação ao simulado de Matemática (com gaps de +21% a +35%). Se CHC pretende fortalecer o pensamento computacional e raciocínio algorítmico, os professores podem introduzir mais lógica formal, álgebra lógica ou atividades matemáticas no escopo dos projetos criativos.</li>
            <li>O rendimento em CHC é estável no ciclo (média próxima de 7.00). Essa constância demonstra que as turmas se adaptam bem à metodologia de trabalho por projetos da disciplina, servindo como uma área de sucesso acadêmico que pode ser alavancada para motivar os alunos em outras disciplinas sob atenção.</li>
          </ul>
        \`;
      }
      
      textDiv.innerHTML = analysis;
    }
  </script>
</body>
</html>
`;

// Inject and write file
const finalHtml = htmlTemplate.replace('${JSON.stringify(fullData)}', JSON.stringify(fullData));
const dashboardHtmlPath = path.join(baseDir, 'dashboard.html');
fs.writeFileSync(dashboardHtmlPath, finalHtml);

console.log(`\n=========================================`);
console.log(`SUCCESS: Dashboard generated at:`);
console.log(`file:///${dashboardHtmlPath.replace(/\\/g, '/')}`);
console.log(`=========================================`);
