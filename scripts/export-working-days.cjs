#!/usr/bin/env node
'use strict';

// Generates a monthly "Working Days" roster workbook styled like the reference template.
// Working day = a roster assignment that represents actual work (status Planned or Confirmed),
// i.e. excludes Day OFF, Paid Leave, Sick and Emergency Leave.

const path = require('node:path');
const fs = require('node:fs');
const ExcelJS = require('exceljs');
const { createStore } = require('../server/db.cjs');

const YEAR = 2026;
const MONTH = 7; // July

function isWorkedStatus(status) {
  const value = String(status || '').toLowerCase();
  return value === 'planned' || value === 'confirmed';
}

function weekdaysInMonth(year, month) {
  const days = new Date(year, month, 0).getDate();
  let count = 0;
  for (let d = 1; d <= days; d += 1) {
    const day = new Date(year, month - 1, d).getDay();
    if (day !== 0 && day !== 6) count += 1;
  }
  return count;
}

async function main() {
  const store = createStore({ dbPath: path.join(__dirname, '..', 'data', 'soc-cockpit.sqlite'), env: process.env });
  const snapshot = store.getSnapshot();
  const prefix = `${YEAR}-${String(MONTH).padStart(2, '0')}`;

  const worked = new Map();
  for (const a of snapshot.rosterAssignments) {
    if (!String(a.rosterDate || '').startsWith(prefix)) continue;
    if (!isWorkedStatus(a.status)) continue;
    const name = String(a.analyst || '').trim();
    if (!name) continue;
    worked.set(name, (worked.get(name) || 0) + 1);
  }

  const rosterStaff = (snapshot.settings && snapshot.settings.rosterStaff) || [];
  const names = rosterStaff.filter((n) => worked.has(n));
  for (const n of worked.keys()) {
    if (!names.includes(n)) names.push(n);
  }

  const baseline = weekdaysInMonth(YEAR, MONTH);
  const rows = names.map((name) => {
    const wd = worked.get(name) || 0;
    return { name, wd, extra: Math.max(0, wd - baseline) };
  });

  const totalWorking = rows.reduce((sum, r) => sum + r.wd, 0);
  const totalExtra = rows.reduce((sum, r) => sum + r.extra, 0);

  const monthLabel = new Date(YEAR, MONTH - 1, 1).toLocaleString('en-US', { month: 'short' }) + `-${String(YEAR).slice(-2)}`;

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Working Days');

  const headerBlue = 'FF4472C4';
  const monthBlue = 'FF2E5496';
  const rowAlt = 'FFD9E1F2';
  const white = 'FFFFFFFF';
  const border = { style: 'thin', color: { argb: 'FFAABBD6' } };

  ws.columns = [
    { width: 6 },
    { width: 40 },
    { width: 14 },
    { width: 10 },
  ];

  // Row 1: month banner over the WD column.
  const r1 = ws.getRow(1);
  r1.height = 20;
  ws.mergeCells('A1:C1');
  const monthCell = ws.getCell('D1');
  monthCell.value = monthLabel;
  monthCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: monthBlue } };
  monthCell.font = { bold: true, color: { argb: white } };
  monthCell.alignment = { horizontal: 'center', vertical: 'middle' };
  monthCell.border = { top: border, bottom: border, left: border, right: border };

  // Row 2: column headers.
  const headerRow = ws.addRow(['SN', 'Description', 'Partner', 'WD' + MONTH]);
  headerRow.height = 20;
  headerRow.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: headerBlue } };
    cell.font = { bold: true, color: { argb: white } };
    cell.alignment = { horizontal: cell.col === 2 ? 'left' : 'center', vertical: 'middle' };
    cell.border = { top: border, bottom: border, left: border, right: border };
  });

  rows.forEach((r, index) => {
    const row = ws.addRow([index + 1, r.name, '', r.wd]);
    row.height = 18;
    const alt = index % 2 === 1;
    row.eachCell((cell, col) => {
      if (alt) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowAlt } };
      cell.border = { top: border, bottom: border, left: border, right: border };
      cell.alignment = { horizontal: col === 2 ? 'left' : 'center', vertical: 'middle' };
    });
  });

  const startCol = 1;
  const totalRow = ws.addRow(['', 'Actual Working days', '', totalWorking]);
  ws.mergeCells(`A${totalRow.number}:C${totalRow.number}`);
  totalRow.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
  totalRow.getCell(2).font = { bold: false };
  totalRow.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
  totalRow.getCell(4).font = { bold: true };
  [1, 2, 3, 4].forEach((c) => { totalRow.getCell(c).border = { top: border, bottom: border, left: border, right: border }; });

  const extraRow = ws.addRow(['', 'Extra days', '', totalExtra]);
  ws.mergeCells(`A${extraRow.number}:C${extraRow.number}`);
  extraRow.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
  extraRow.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
  extraRow.getCell(4).font = { bold: true };
  [1, 2, 3, 4].forEach((c) => { extraRow.getCell(c).border = { top: border, bottom: border, left: border, right: border }; });

  const outDir = path.join(__dirname, '..', 'artifacts');
  fs.mkdirSync(outDir, { recursive: true });
  const stamp = new Date().toISOString().replace(/[:T]/g, '-').slice(0, 15);
  const outFile = path.join(outDir, `working-days-${prefix}-${stamp}.xlsx`);
  await wb.xlsx.writeFile(outFile);

  console.log(`Staff counted: ${rows.length}`);
  console.log(`Weekday baseline (Mon-Fri) for ${monthLabel}: ${baseline}`);
  console.log(`Actual Working days: ${totalWorking}`);
  console.log(`Extra days (worked beyond ${baseline} weekdays): ${totalExtra}`);
  console.log(`Saved: ${outFile}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
