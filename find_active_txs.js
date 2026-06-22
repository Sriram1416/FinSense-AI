import fs from 'fs';

const appContent = fs.readFileSync('src/App.jsx', 'utf-8');
const lines = appContent.split('\n');

console.log("Searching for activeAnalysisTxs in App.jsx...");
lines.forEach((line, index) => {
  if (line.includes('activeAnalysisTxs') && (line.includes('const') || line.includes('let') || line.includes('useMemo'))) {
    console.log(`${index + 1}: ${line.trim()}`);
  }
});
