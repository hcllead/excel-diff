import XLSX from 'xlsx';
import fs from 'fs';

// Read Excel file
const workbook = XLSX.readFile('example.xlsx');
let mdContent = '';

workbook.SheetNames.forEach(sheetName => {
  const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
  if (sheet.length === 0) return;

  // Add sheet title
  mdContent += `## ${sheetName}\n\n`;

  // Create table headers
  const headers = Object.keys(sheet[0]);
  mdContent += `| ${headers.join(' | ')} |\n`;
  mdContent += `| ${headers.map(() => '---').join(' | ')} |\n`;

  // Add rows
  sheet.forEach(row => {
    mdContent += `| ${headers.map(h => row[h] || '').join(' | ')} |\n`;
  });

  mdContent += '\n\n';
});

// Save to .md file
fs.writeFileSync('output.md', mdContent);
console.log('Markdown file created: output.md');