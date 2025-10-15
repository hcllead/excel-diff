import XLSX from 'xlsx';
import fs from 'fs';

// Read Excel file
const workbook = XLSX.readFile('example.xlsx');
let mdContent = '';

const colors = ['red', 'yellow', 'green'];

workbook.SheetNames.forEach(sheetName => {
  const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
  if (sheet.length === 0) return;

  mdContent += `## ${sheetName}\n\n`;

  const headers = Object.keys(sheet[0]);
  mdContent += `| ${headers.join(' | ')} |\n`;
  mdContent += `| ${headers.map(() => '---').join(' | ')} |\n`;

  sheet.forEach(row => {
    mdContent += '| ';
    headers.forEach((h, i) => {
      const color = colors[Math.floor(Math.random() * colors.length)];
      mdContent += `<span style="color:${color}">${row[h] || ''}</span> | `;
    });
    mdContent += '\n';
  });

  mdContent += '\n\n';
});

fs.writeFileSync('output.md', mdContent);
