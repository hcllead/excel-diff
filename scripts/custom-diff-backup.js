// scripts/custom-diff.js
import { execSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import * as XLSX from "xlsx";

const BASE = process.env.BASE_SHA;
const HEAD = process.env.HEAD_SHA;
const LIST = (process.env.XLSX_LIST || "").split("\n").filter(Boolean);

// helpers
const sh = (cmd) => execSync(cmd, { encoding: "utf8" }).trim();

function writeBlobToTmp(sha, filePath) {
  const tmp = path.join(
    os.tmpdir(),
    `${sha}-${filePath.replace(/[\\/]/g, "__")}`
  );
  try {
    const buf = execSync(`git show ${sha}:${filePath}`, { encoding: "buffer" });
    fs.writeFileSync(tmp, buf);
    return tmp;
  } catch {
    return null; // added/removed
  }
}

function workbookToCellMap(tmpXlsxPath) {
  if (!tmpXlsxPath || !fs.existsSync(tmpXlsxPath)) return {};
  const wb = XLSX.read(fs.readFileSync(tmpXlsxPath));
  const map = {}; // { "Sheet1": { "A1": "val", ... } }
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const cells = {};
    const range = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        if (cell && cell.v !== undefined) {
          // Normalize values to strings for stable compare
          cells[addr] = String(cell.v);
        }
      }
    }
    map[sheetName] = cells;
  }
  return map;
}

function diffCellMaps(aMap, bMap) {
  // returns array of {sheet, addr, type, from?, to?}
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
    // Optional: row/col summaries
  }
  return diffs;
}

function summarizeByRowCol(diffs) {
  // returns { "<sheet>": { rows: Map<row, count>, cols: Map<col, count> } }
  const sum = {};
  for (const d of diffs) {
    const sheet = d.sheet;
    sum[sheet] ||= { rows: new Map(), cols: new Map() };
    // Parse "A1" -> col "A", row 1
    const match = d.addr.match(/^([A-Z]+)(\d+)$/i);
    if (match) {
      const col = match[1].toUpperCase();
      const row = parseInt(match[2], 10);
      sum[sheet].rows.set(row, (sum[sheet].rows.get(row) || 0) + 1);
      sum[sheet].cols.set(col, (sum[sheet].cols.get(col) || 0) + 1);
    }
  }
  return sum;
}

let md = [];
md.push(`# Custom Diff Report (Excel)`);
md.push(`Base: \`${BASE}\` → Head: \`${HEAD}\``);
md.push(`Changed Excel files: **${LIST.length}**`);
md.push("");

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
    // Summaries by row/col
    const summary = summarizeByRowCol(diffs);
    const s = summary[sheet];

    md.push(`### Sheet: \`${sheet}\``);
    if (s) {
      // Row summary (top 10)
      const topRows = [...s.rows.entries()]
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);
      const topCols = [...s.cols.entries()]
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);
      md.push(
        `**Rows touched (top 10):** ${
          topRows.map(([r, c]) => `${r}(${c})`).join(", ") || "—"
        }`
      );
      md.push(
        `**Cols touched (top 10):** ${
          topCols.map(([c, k]) => `${c}(${k})`).join(", ") || "—"
        }`
      );
    }
    md.push("");
    md.push(`| Cell | Change |`);
    md.push(`|---|---|`);
    // Limit table size for comment readability
    const MAX = 200;
    for (const d of diffs.slice(0, MAX)) {
      if (d.type === "changed") {
        md.push(`| ${sheet}!${d.addr} | \`${d.from}\` → \`${d.to}\` |`);
      } else if (d.type === "added") {
        md.push(`| ${sheet}!${d.addr} | ⊕ \`${d.to}\` |`);
      } else {
        md.push(`| ${sheet}!${d.addr} | ⊖ \`${d.from}\` |`);
      }
    }
    if (diffs.length > MAX) {
      md.push(
        `_…and ${diffs.length - MAX} more cells (see artifact for full list)._`
      );
    }
    md.push("");
  }
}

fs.writeFileSync("custom-diff.md", md.join("\n"), "utf8");
console.log("Wrote custom-diff.md");