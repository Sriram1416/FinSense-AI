import fs from 'fs';

const appContent = fs.readFileSync('src/App.jsx', 'utf-8');
const lines = appContent.split('\n');

console.log("Searching for split displays (is_shared and tx.amount) in App.jsx...");
lines.forEach((line, index) => {
  if (line.includes('is_shared') && (line.includes('amount') || line.includes('roommates.length'))) {
    console.log(`${index + 1}: ${line.trim()}`);
  }
});
