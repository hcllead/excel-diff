// scripts/custom-diff-visual.js
import { execSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import * as XLSX from "xlsx";

const BASE = process.env.BASE_SHA;
const HEAD = process.env.HEAD_SHA;
const LIST = (process.env.XLSX_LIST || "").split("\n").filter(Boolean);

function writeBlobToTmp(sha, filePath) {
  const tmp = path.join(os.tmpdir(), `${sha}-${filePath.replace(/[\\/]/g, "__")}`);
  try {
    const buf = execSync(`git show ${sha}:${filePath}`, { encoding: "buffer" });
    fs.writeFileSync(tmp, buf);
    return tmp;
  } catch {
    return null;
  }
}

function workbookToCellMap(tmpXlsxPath) {
  if (!tmpXlsxPath || !fs.existsSync(tmpXlsxPath)) return {};
  const wb = XLSX.read(fs.readFileSync(tmpXlsxPath));
  const map = {};
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const cells = {};
    const range = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        if (cell && cell.v !== undefined) {
          cells[addr] = String(cell.v);
        }
      }
    }
    map[sheetName] = cells;
  }
  return map;
}

function diffCellMaps(aMap, bMap) {
  const diffs = [];
  const sheets = new Set([...Object.keys(aMap), ...Object.keys(bMap)]);
  for (const sheet of sheets) {
    const A = aMap[sheet] || {};
    const B = bMap[sheet] || {};
    const addrs = new Set([...Object.keys(A), ...Object.keys(B)]);
    for (const addr of addrs) {
      const av = A[addr];
      const bv = B[addr];
      if (av === undefined && bv !== undefined) {
        diffs.push({ sheet, addr, type: "added", to: bv });
      } else if (av !== undefined && bv === undefined) {
        diffs.push({ sheet, addr, type: "removed", from: av });
      } else if (av !== bv) {
        diffs.push({ sheet, addr, type: "changed", from: av, to: bv });
      }
    }
  }
  return diffs;
}

function buildVisualTable(sheetName, diffs, aMap, bMap) {
  const allAddrs = new Set([...Object.keys(aMap[sheetName] || {}), ...Object.keys(bMap[sheetName] || {})]);
  if (allAddrs.size === 0) return "<p>No data</p>";

  const rows = [];
  const cols = [];

  for (const addr of allAddrs) {
    const match = addr.match(/^([A-Z]+)(\d+)$/i);
    if (match) {
      cols.push(match[1]);
      rows.push(parseInt(match[2], 10));
    }
  }

  const uniqueCols = [...new Set(cols)].sort((a, b) => a.localeCompare(b));
  const uniqueRows = [...new Set(rows)].sort((a, b) => a - b);

  const diffMap = new Map();
  for (const d of diffs) {
    diffMap.set(d.addr, d);
  }

  let html = `<table border="1" cellspacing="0" cellpadding="4" style="border-collapse:collapse;">`;
  html += `<tr><th></th>`;
  for (const col of uniqueCols) html += `<th>${col}</th>`;
  html += `</tr>`;

  for (const r of uniqueRows) {
    html += `<tr><th>${r}</th>`;
    for (const col of uniqueCols) {
      const addr = `${col}${r}`;
      const diff = diffMap.get(addr);
      let cellHtml = "";
      if (diff) {
        if (diff.type === "changed") {
          cellHtml = `<td style="background-color:yellow"><del>${diff.from}<del> → $$\\color{orange}{${diff.to}}$$</td>`;
        } else if (diff.type === "added") {
          cellHtml = `<td style="background-color:green">$$\\color{green}{${diff.to}}$$</td>`;
        } else {
          cellHtml = `<td style="background-color:red"><del>$$\\color{red}{${diff.from}}$$<del></td>`;
        }
      } else {
        const val = bMap[sheetName]?.[addr] || aMap[sheetName]?.[addr] || "";
        cellHtml = `<td>${val}</td>`;
      }
      html += cellHtml;
    }
    html += `</tr>`;
  }
  html += `</table>`;
  return html;
}

// Main report
let md = [];
md.push(`# Custom Diff Report (Excel)`);
md.push(`Changed Excel files: **${LIST.length}**`);
md.push("");
md.push(`**Legend:**  
- Yellow = Modified (~~old~~ → $$\\color{orange}{new}$$)  
- Red = Deleted (~~$$\\color{red}{old}$$~~ ) 
- Green = Added ($$\\color{green}{new}$$)`);

for (const file of LIST) {
  md.push(`## ${file}`);
  const aTmp = writeBlobToTmp(BASE, file);
  const bTmp = writeBlobToTmp(HEAD, file);

  if (!aTmp && bTmp) {
    md.push(`_Added file_`);
    md.push("");
    continue;
  }
  if (aTmp && !bTmp) {
    md.push(`_Removed file_`);
    md.push("");
    continue;
  }

  const aMap = workbookToCellMap(aTmp);
  const bMap = workbookToCellMap(bTmp);
  const cellDiffs = diffCellMaps(aMap, bMap);
  const total = cellDiffs.length;

  if (total === 0) {
    md.push(`No cell changes.`);
    md.push("");
    continue;
  }

  const bySheet = new Map();
  for (const d of cellDiffs) {
    if (!bySheet.has(d.sheet)) bySheet.set(d.sheet, []);
    bySheet.get(d.sheet).push(d);
  }

  md.push(`**Total cell changes:** ${total}`);
  md.push("");

  for (const [sheet, diffs] of bySheet) {
    md.push(`### Sheet: ${sheet}`); // Strike sheet name for consistency
    md.push(buildVisualTable(sheet, diffs, aMap, bMap));
    md.push("");
  }
}

fs.writeFileSync("custom-diff.md", md.join("\n"), "utf8");
console.log("Wrote custom-diff.md with visual tables and strike-through");