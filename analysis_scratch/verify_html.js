const fs = require('fs');
const path = require('path');

const htmlPath = path.join(__dirname, '..', 'dashboard.html');
if (!fs.existsSync(htmlPath)) {
  console.error('dashboard.html not found!');
  process.exit(1);
}

const html = fs.readFileSync(htmlPath, 'utf8');

console.log('File size:', html.length, 'bytes');

const hasButton = html.includes('btnTabRelatorio');
console.log('Has report tab button:', hasButton);

const hasDiv = html.includes('id="relatorioTab"');
console.log('Has report tab div:', hasDiv);

const hasJS = html.includes('switchReportSection');
console.log('Has report section JS function:', hasJS);

if (hasButton && hasDiv && hasJS) {
  console.log('VERIFICATION SUCCESSFUL: Report integrated into HTML.');
} else {
  console.error('VERIFICATION FAILED: Some parts are missing!');
  process.exit(1);
}
