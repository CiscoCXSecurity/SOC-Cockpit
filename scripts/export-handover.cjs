// SPDX-FileCopyrightText: 2026 Cisco Systems, Inc.
// SPDX-License-Identifier: Apache-2.0

const fs = require('node:fs');
const path = require('node:path');
const dotenv = require('dotenv');
const ExcelJS = require('exceljs');
const { createStore } = require('../server/db.cjs');

const workspaceRoot = path.join(__dirname, '..');
const dbPath = path.join(workspaceRoot, 'data', 'soc-cockpit.sqlite');
const outputDir = path.join(workspaceRoot, 'artifacts', 'handover');

dotenv.config({ path: path.join(workspaceRoot, '.env') });

function pad(value) {
  return String(value).padStart(2, '0');
}

function timestamp(date) {
  return [
    date.getFullYear(),
    pad(date.getMonth() + 1),
    pad(date.getDate()),
    '-',
    pad(date.getHours()),
    pad(date.getMinutes()),
    pad(date.getSeconds()),
  ].join('');
}

function toText(value) {
  if (value == null) return '';
  if (typeof value === 'boolean') return value ? 'Yes' : 'No';
  return String(value);
}

function addSheet(workbook, name, columns, rows) {
  const sheet = workbook.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
  sheet.columns = columns.map((col) => ({ header: col.header, key: col.key, width: col.width || 18 }));

  const header = sheet.getRow(1);
  header.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0D2138' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFC7D2DD' } },
      bottom: { style: 'thin', color: { argb: 'FFC7D2DD' } },
      left: { style: 'thin', color: { argb: 'FFC7D2DD' } },
      right: { style: 'thin', color: { argb: 'FFC7D2DD' } },
    };
  });

  rows.forEach((row) => {
    const normalized = {};
    columns.forEach((col) => {
      normalized[col.key] = toText(row[col.key]);
    });
    sheet.addRow(normalized);
  });

  sheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: columns.length },
  };
}

async function main() {
  fs.mkdirSync(outputDir, { recursive: true });

  const store = createStore({ dbPath, env: process.env });
  const snapshot = store.getSnapshot();
  const now = new Date();
  const stamp = timestamp(now);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'SOC Cockpit';
  workbook.created = now;

  addSheet(
    workbook,
    'Overview',
    [
      { header: 'Section', key: 'section', width: 32 },
      { header: 'Count', key: 'count', width: 14 },
      { header: 'Notes', key: 'notes', width: 64 },
    ],
    [
      { section: 'Generated At', count: now.toISOString(), notes: 'UTC timestamp for this handover package.' },
      { section: 'Workforce Assignments', count: snapshot.rosterAssignments.length, notes: 'Roster planning rows.' },
      { section: 'Workforce Staff', count: (snapshot.settings?.rosterStaff || []).length, notes: 'Configured roster staff list.' },
      { section: 'Workforce Labels', count: (snapshot.settings?.rosterLabels || []).length, notes: 'Shift/label configuration used in planner.' },
      { section: 'Directory', count: snapshot.contacts.length, notes: 'People and contact records.' },
      { section: 'Issues', count: snapshot.issues.length, notes: 'Issue register.' },
      { section: 'Standups', count: snapshot.matrixTasks.length, notes: 'Standup/stand-down tasks.' },
      { section: 'Projects', count: snapshot.projects.length, notes: 'Project register.' },
      { section: 'Objectives', count: snapshot.objectives.length, notes: 'Objective register.' },
      { section: 'Notes', count: snapshot.notes.length, notes: 'Operational notes and pins.' },
      { section: 'Service KPIs', count: snapshot.kpis.length, notes: 'KPI status register.' },
    ]
  );

  addSheet(
    workbook,
    'Workforce_Roster',
    [
      { header: 'ID', key: 'id', width: 10 },
      { header: 'Roster Date', key: 'rosterDate', width: 14 },
      { header: 'Shift', key: 'shiftName', width: 12 },
      { header: 'Role', key: 'role', width: 20 },
      { header: 'Analyst', key: 'analyst', width: 24 },
      { header: 'Status', key: 'status', width: 14 },
      { header: 'Health', key: 'health', width: 12 },
      { header: 'Notes/Hours', key: 'notes', width: 24 },
    ],
    snapshot.rosterAssignments
  );

  addSheet(
    workbook,
    'Workforce_Staff',
    [
      { header: 'Staff Name', key: 'name', width: 28 },
    ],
    (snapshot.settings?.rosterStaff || []).map((name) => ({ name }))
  );

  addSheet(
    workbook,
    'Workforce_Labels',
    [
      { header: 'Code', key: 'code', width: 10 },
      { header: 'Label', key: 'label', width: 20 },
      { header: 'Hours/Notes', key: 'hours', width: 16 },
      { header: 'Status', key: 'status', width: 14 },
      { header: 'Health', key: 'health', width: 12 },
      { header: 'Fill Color', key: 'fill', width: 14 },
      { header: 'Text Color', key: 'textColor', width: 14 },
    ],
    snapshot.settings?.rosterLabels || []
  );

  addSheet(
    workbook,
    'Directory',
    [
      { header: 'ID', key: 'id', width: 8 },
      { header: 'Name', key: 'name', width: 24 },
      { header: 'Role', key: 'role', width: 26 },
      { header: 'Sub Team', key: 'channel', width: 20 },
      { header: 'Email', key: 'email', width: 30 },
      { header: 'Phone', key: 'phone', width: 18 },
      { header: 'Scope', key: 'scope', width: 12 },
      { header: 'Is Lead', key: 'isLead', width: 10 },
    ],
    snapshot.contacts
  );

  addSheet(
    workbook,
    'Issues',
    [
      { header: 'ID', key: 'id', width: 8 },
      { header: 'Scope', key: 'scope', width: 8 },
      { header: 'Status', key: 'status', width: 16 },
      { header: 'Title', key: 'title', width: 44 },
      { header: 'Summary', key: 'summary', width: 52 },
      { header: 'Owner', key: 'owner', width: 20 },
      { header: 'Assigned To', key: 'assignedTo', width: 20 },
      { header: 'Linked Type', key: 'linkedType', width: 14 },
      { header: 'Linked ID', key: 'linkedId', width: 10 },
      { header: 'Due', key: 'due', width: 12 },
      { header: 'Progress', key: 'progress', width: 10 },
      { header: 'Health', key: 'health', width: 12 },
    ],
    snapshot.issues
  );

  addSheet(
    workbook,
    'Standups',
    [
      { header: 'ID', key: 'id', width: 8 },
      { header: 'Task Date', key: 'taskDate', width: 14 },
      { header: 'Title', key: 'title', width: 44 },
      { header: 'Lead', key: 'lead', width: 20 },
      { header: 'Assignee', key: 'assignee', width: 20 },
      { header: 'Minutes', key: 'durationMinutes', width: 10 },
      { header: 'Urgent', key: 'urgency', width: 10 },
      { header: 'Important', key: 'importance', width: 12 },
      { header: 'Quadrant', key: 'quadrant', width: 14 },
      { header: 'Disposition', key: 'disposition', width: 14 },
      { header: 'Due Time', key: 'dueTime', width: 12 },
      { header: 'Status', key: 'status', width: 16 },
      { header: 'Notes', key: 'notes', width: 46 },
    ],
    snapshot.matrixTasks
  );

  addSheet(
    workbook,
    'Projects',
    [
      { header: 'ID', key: 'id', width: 8 },
      { header: 'Code', key: 'code', width: 12 },
      { header: 'Name', key: 'name', width: 42 },
      { header: 'Owner', key: 'owner', width: 20 },
      { header: 'Phase', key: 'phase', width: 14 },
      { header: 'Due', key: 'due', width: 12 },
      { header: 'Progress', key: 'progress', width: 10 },
      { header: 'Health', key: 'health', width: 12 },
    ],
    snapshot.projects
  );

  addSheet(
    workbook,
    'Objectives',
    [
      { header: 'ID', key: 'id', width: 8 },
      { header: 'Scope', key: 'scope', width: 8 },
      { header: 'Title', key: 'title', width: 42 },
      { header: 'Summary', key: 'summary', width: 50 },
      { header: 'Owner', key: 'owner', width: 20 },
      { header: 'Assigned To', key: 'assignedTo', width: 20 },
      { header: 'Due', key: 'due', width: 12 },
      { header: 'Progress', key: 'progress', width: 10 },
      { header: 'Health', key: 'health', width: 12 },
      { header: 'Attachment Name', key: 'attachmentName', width: 24 },
      { header: 'Attachment Path', key: 'attachmentPath', width: 36 },
    ],
    snapshot.objectives
  );

  addSheet(
    workbook,
    'Notes',
    [
      { header: 'ID', key: 'id', width: 8 },
      { header: 'Created At', key: 'createdAt', width: 24 },
      { header: 'Note Date', key: 'noteDate', width: 14 },
      { header: 'Pinned', key: 'pinned', width: 10 },
      { header: 'Text', key: 'text', width: 80 },
    ],
    snapshot.notes
  );

  addSheet(
    workbook,
    'Service_KPIs',
    [
      { header: 'ID', key: 'id', width: 8 },
      { header: 'Code', key: 'code', width: 12 },
      { header: 'Name', key: 'name', width: 34 },
      { header: 'Actual', key: 'actual', width: 14 },
      { header: 'Target', key: 'target', width: 14 },
      { header: 'Health', key: 'health', width: 12 },
      { header: 'Trend', key: 'trend', width: 10 },
      { header: 'Delta', key: 'delta', width: 12 },
    ],
    snapshot.kpis
  );

  const workbookPath = path.join(outputDir, `soc-handover-${stamp}.xlsx`);
  const snapshotPath = path.join(outputDir, `soc-handover-${stamp}.snapshot.json`);

  await workbook.xlsx.writeFile(workbookPath);
  fs.writeFileSync(snapshotPath, JSON.stringify(snapshot, null, 2));

  console.log(`Handover workbook created: ${workbookPath}`);
  console.log(`Snapshot JSON created: ${snapshotPath}`);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
});
