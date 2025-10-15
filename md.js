import XLSX from 'xlsx';
import fs from 'fs';



// Read Excel file
const workbook = XLSX.readFile('example.xlsx');
const sheetName = workbook.SheetNames[0];
const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

// Convert to Markdown table
const headers = Object.keys(sheet[0]);
let mdTable = `| ${headers.join(' | ')} |\n`;
mdTable += `| ${headers.map(() => '---').join(' | ')} |\n`;


sheet.forEach(row => {
  mdTable += `| ${headers.map(h => row[h] || '').join(' | ')} |\n`;
});

// Save to .md file
fs.writeFileSync('output.md', mdTable);
console.log('Markdown file created: output.md');

