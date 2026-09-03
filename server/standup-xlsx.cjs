// SPDX-FileCopyrightText: 2026 Cisco Systems, Inc.
// SPDX-License-Identifier: Apache-2.0

const ExcelJS = require('exceljs');

const SHEET_NAME = 'Standup';
const INSTRUCTIONS_SHEET = 'Instructions';
const DATA_ROWS = 200;

// Canonical columns in display order. `key` is the internal field the importer maps to.
const COLUMNS = [
  { key: 'date', header: 'Date (YYYY-MM-DD)', width: 16 },
  { key: 'objective', header: 'Objective', width: 30 },
  { key: 'title', header: 'Task Description', width: 52 },
  { key: 'lead', header: 'Lead', width: 20 },
  { key: 'assignee', header: 'Assignee', width: 20 },
  { key: 'minutes', header: 'Minutes', width: 10 },
  { key: 'urgent', header: 'Urgent (Y/N)', width: 12 },
  { key: 'important', header: 'Important (Y/N)', width: 14 },
  { key: 'dueTime', header: 'Due Time (HH:MM)', width: 16 },
  { key: 'status', header: 'Status', width: 18 },
  { key: 'notes', header: 'Stand-down Notes', width: 52 },
];

const INSTRUCTIONS = [
  'SOC Daily Standup / Stand-down — Offline Capture Sheet',
  '',
  'Use this workbook when the SOC Cockpit app is unavailable. Record each standup task on its own row in the "Standup" tab, then upload this file back into the app when it is available again.',
  '',
  'How to fill it in:',
  '• Date — the standup day for the task, formatted YYYY-MM-DD (e.g. 2026-07-31). One file can contain multiple dates.',
  '• Objective — the objective this task contributes to (optional). On upload, a matching objective is reused or a new one is created on the Objectives page, and the task is linked to it. Leave blank if the task is not tied to an objective.',
  '• Task Description — what needs to be done. Required. Rows without a task are ignored.',
  '• Lead — the standup lead who owns the task. Required for import.',
  '• Assignee — the person doing the work (optional).',
  '• Minutes — estimated/actual effort in minutes (optional, whole number).',
  '• Urgent (Y/N) and Important (Y/N) — drive the Eisenhower quadrant (Do / Delegate / Delay / Discard).',
  '• Due Time (HH:MM) — optional target time, e.g. 17:00.',
  '• Status — "Completed" or "Not Completed" (this is the stand-down outcome). Blank is treated as Not Completed.',
  '• Stand-down Notes — end-of-day comment, blockers, or follow-up captured at stand-down (optional).',
  '',
  'On upload the app matches rows by Date + Task Description + Lead. New rows are added; matching rows are updated in place (e.g. stand-down status and notes), so re-sending the same sheet day after day will not create duplicates.',
  '',
  'Do not rename the "Standup" tab or the header row — the importer matches columns by their header names.',
];

function buildStandupTemplate() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'SOC Cockpit';
  workbook.created = new Date();

  const guide = workbook.addWorksheet(INSTRUCTIONS_SHEET, { properties: { defaultColWidth: 20 } });
  guide.getColumn(1).width = 118;
  INSTRUCTIONS.forEach((line, index) => {
    const row = guide.getRow(index + 1);
    const cell = row.getCell(1);
    cell.value = line;
    cell.alignment = { vertical: 'top', wrapText: true };
    if (index === 0) {
      cell.font = { bold: true, size: 14, color: { argb: 'FF0D2138' } };
    } else if (line.endsWith(':')) {
      cell.font = { bold: true };
    }
  });

  const sheet = workbook.addWorksheet(SHEET_NAME, {
    views: [{ state: 'frozen', ySplit: 1 }],
  });

  const headerRow = sheet.getRow(1);
  COLUMNS.forEach((column, index) => {
    const cell = headerRow.getCell(index + 1);
    cell.value = column.header;
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A60FF' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFCFD9E4' } },
      bottom: { style: 'thin', color: { argb: 'FFCFD9E4' } },
      left: { style: 'thin', color: { argb: 'FFCFD9E4' } },
      right: { style: 'thin', color: { argb: 'FFCFD9E4' } },
    };
    sheet.getColumn(index + 1).width = column.width;
  });
  headerRow.height = 26;
  headerRow.commit();

  const columnLetter = (key) => sheet.getColumn(COLUMNS.findIndex((column) => column.key === key) + 1).letter;
  const lastRow = DATA_ROWS + 1;
  const listValidation = (values) => ({
    type: 'list',
    allowBlank: true,
    formulae: [`"${values.join(',')}"`],
    showErrorMessage: true,
    errorStyle: 'warning',
    errorTitle: 'Unexpected value',
    error: `Please choose one of: ${values.join(', ')}`,
  });

  for (const key of ['urgent', 'important']) {
    const letter = columnLetter(key);
    sheet.dataValidations.add(`${letter}2:${letter}${lastRow}`, listValidation(['Y', 'N']));
  }
  const statusLetter = columnLetter('status');
  sheet.dataValidations.add(`${statusLetter}2:${statusLetter}${lastRow}`, listValidation(['Completed', 'Not Completed']));

  // Keep the Date column as text so YYYY-MM-DD is preserved exactly across locales.
  const dateColumn = sheet.getColumn(COLUMNS.findIndex((column) => column.key === 'date') + 1);
  dateColumn.numFmt = '@';

  return workbook.xlsx.writeBuffer().then((data) => Buffer.from(data));
}

function detectColumnKey(headerText) {
  const value = String(headerText || '').trim().toLowerCase();
  if (!value) return null;
  if (value.includes('objective')) return 'objective';
  if (value.includes('date')) return 'date';
  if (value.includes('task') || value.includes('description')) return 'title';
  if (value.includes('lead')) return 'lead';
  if (value.includes('assign')) return 'assignee';
  if (value.includes('minute')) return 'minutes';
  if (value.includes('urgent')) return 'urgent';
  if (value.includes('important')) return 'important';
  if (value.includes('due')) return 'dueTime';
  if (value.includes('status')) return 'status';
  if (value.includes('note') || value.includes('stand-down') || value.includes('stand down')) return 'notes';
  return null;
}

function cellToText(value) {
  if (value == null) return '';
  if (typeof value === 'string') return value.trim();
  if (typeof value === 'number' || typeof value === 'boolean') return String(value).trim();
  if (value instanceof Date) return isoFromDate(value);
  if (typeof value === 'object') {
    if (Array.isArray(value.richText)) return value.richText.map((part) => part.text).join('').trim();
    if (value.text != null) return String(value.text).trim();
    if (value.result != null) return String(value.result).trim();
    if (value.hyperlink != null) return String(value.text || value.hyperlink).trim();
  }
  return String(value).trim();
}

function isoFromDate(date) {
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const day = String(date.getUTCDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function normalizeDate(raw) {
  const value = String(raw || '').trim();
  if (!value) return '';
  const isoMatch = value.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (isoMatch) {
    return `${isoMatch[1]}-${isoMatch[2].padStart(2, '0')}-${isoMatch[3].padStart(2, '0')}`;
  }
  const slashMatch = value.match(/^(\d{1,2})[/](\d{1,2})[/](\d{4})$/);
  if (slashMatch) {
    // Assume day/month/year (British) ordering to match the app's locale.
    return `${slashMatch[3]}-${slashMatch[2].padStart(2, '0')}-${slashMatch[1].padStart(2, '0')}`;
  }
  return '';
}

function parseBoolean(raw) {
  const value = String(raw || '').trim().toLowerCase();
  return value === 'y' || value === 'yes' || value === 'true' || value === '1' || value === 'x';
}

function normalizeStatus(raw) {
  const value = String(raw || '').trim().toLowerCase();
  return value === 'completed' || value === 'complete' || value === 'done' || value === 'yes' ? 'Completed' : 'Not Completed';
}

function parseMinutes(raw) {
  const number = Number(String(raw || '').replace(/[^0-9.]/g, ''));
  return Number.isFinite(number) && number > 0 ? Math.round(number) : 0;
}

async function parseStandupWorkbook(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const sheet = workbook.getWorksheet(SHEET_NAME) || workbook.worksheets.find((ws) => ws.name !== INSTRUCTIONS_SHEET) || workbook.worksheets[0];
  if (!sheet) {
    throw new Error('The workbook has no worksheets to import.');
  }

  const columnMap = {};
  sheet.getRow(1).eachCell((cell, colNumber) => {
    const key = detectColumnKey(cellToText(cell.value));
    if (key && columnMap[key] == null) {
      columnMap[key] = colNumber;
    }
  });

  if (columnMap.date == null || columnMap.title == null || columnMap.lead == null) {
    throw new Error('This file does not look like a SOC standup template. It must have Date, Task Description, and Lead columns.');
  }

  const rows = [];
  const errors = [];
  const valueAt = (row, key) => (columnMap[key] != null ? cellToText(row.getCell(columnMap[key]).value) : '');

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const title = valueAt(row, 'title');
    const dateRaw = valueAt(row, 'date');
    const lead = valueAt(row, 'lead');
    if (!title && !dateRaw && !lead) {
      return; // fully blank row
    }
    if (!title) {
      errors.push(`Row ${rowNumber}: skipped (no task description).`);
      return;
    }
    const taskDate = normalizeDate(dateRaw);
    if (!taskDate) {
      errors.push(`Row ${rowNumber}: skipped (invalid or missing date "${dateRaw}").`);
      return;
    }
    if (!lead) {
      errors.push(`Row ${rowNumber}: skipped (no lead).`);
      return;
    }
    rows.push({
      taskDate,
      objective: valueAt(row, 'objective'),
      title,
      lead,
      assignee: valueAt(row, 'assignee'),
      durationMinutes: parseMinutes(valueAt(row, 'minutes')),
      urgency: parseBoolean(valueAt(row, 'urgent')),
      importance: parseBoolean(valueAt(row, 'important')),
      dueTime: valueAt(row, 'dueTime'),
      status: normalizeStatus(valueAt(row, 'status')),
      notes: valueAt(row, 'notes'),
    });
  });

  return { rows, errors };
}

module.exports = { buildStandupTemplate, parseStandupWorkbook, SHEET_NAME };
