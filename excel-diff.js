import { readFileSync, existsSync } from 'fs';
import { read, utils } from 'xlsx';

// ✅ Read Excel File
export function readExcelFile(filePath) {
    if (!existsSync(filePath)) {
        console.error(`File not found: ${filePath}`);
        return [];
    }
    const workbook = read(readFileSync(filePath));
    const sheetName = workbook.SheetNames[0]; // First sheet only
    return utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
}

// ✅ Compare Sheets
export function compareSheets(oldSheet, newSheet) {
    const oldRows = new Set(oldSheet.map(row => JSON.stringify(row)));
    const newRows = new Set(newSheet.map(row => JSON.stringify(row)));

    const added = [...newRows].filter(row => !oldRows.has(row)).map(JSON.parse);
    const removed = [...oldRows].filter(row => !newRows.has(row)).map(JSON.parse);
    const unchanged = [...newRows].filter(row => oldRows.has(row)).map(JSON.parse);

    return { added, removed, unchanged };
}

// ✅ Print Differences
export function printDiff(diffs) {
    console.log("➕ Added Rows:");
    diffs.added.forEach(row => console.log(row));
    console.log("➖ Removed Rows:");
    diffs.removed.forEach(row => console.log(row));
    console.log("✅ Unchanged Rows:");
    diffs.unchanged.forEach(row => console.log(row));
}

// ✅ Entry Point
if (process.argv.length === 4) {
    const oldFile = process.argv[2];
    const newFile = process.argv[3];

    const oldSheet = readExcelFile(oldFile);
    const newSheet = readExcelFile(newFile);

    const diffs = compareSheets(oldSheet, newSheet);
    printDiff(diffs);
} else {
    console.log('Usage: node excel-diff.js <old_file.xlsx> <new_file.xlsx>');
}
