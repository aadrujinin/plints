const ExcelJS = require('exceljs');
const path = require('path');

const base = __dirname;

function getCellText(cell) {
    if (!cell || cell.value === undefined || cell.value === null) return '';
    if (typeof cell.value === 'string') return cell.value;
    if (typeof cell.value === 'object') {
        if (cell.value.richText && Array.isArray(cell.value.richText)) {
            return cell.value.richText.map(rt => rt.text).join('');
        }
        if (cell.value.text) return cell.value.text;
        if (cell.value.result !== undefined && cell.value.result !== null) {
            return String(cell.value.result);
        }
        if (cell.value.formula !== undefined) return 'formula:'+cell.value.formula;
        return JSON.stringify(cell.value);
    }
    return String(cell.value);
}

async function dumpSheet(file, sheetName, label) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  if (sheetName === '*') {
    for (const ws of wb.worksheets) {
      await dumpSheetContent(ws, ws.name, label);
    }
  } else {
    const ws = wb.getWorksheet(sheetName);
    if (ws) await dumpSheetContent(ws, sheetName, label);
  }
}

async function dumpSheetContent(ws, sheetName, label) {
  console.log('=== ' + label + ' / ' + sheetName + ' === rows=' + ws.rowCount + ' cols=' + ws.columnCount);
  let lastRowWithContent = 0;
  for (let r = 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    let hasContent = false;
    for (let c = 1; c <= (ws.columnCount || 25); c++) {
      if (row.getCell(c).value) { hasContent = true; break; }
    }
    if (hasContent) lastRowWithContent = r;
  }
  console.log('  Last row with content: ' + lastRowWithContent);
  
  for (let r = 1; r <= lastRowWithContent; r++) {
    const row = ws.getRow(r);
    const vals = [];
    for (let c = 1; c <= Math.min(ws.columnCount || 25, 25); c++) {
      const cell = row.getCell(c);
      let v = getCellText(cell);
      if (v) vals.push('C' + c + '=' + v.substring(0, 80));
    }
    if (vals.length) {
      console.log('  R' + r + ': ' + vals.join(' | '));
    } else {
      console.log('  R' + r + ': [empty]');
    }
  }
  console.log('');
}

async function main() {
  // Dump ALL sheets from etalon SV005
  await dumpSheet(
    path.join(base, '(г.Курган_ ул.Карла Маркса_ 149)_2026-07-03_SV005_fixed (1).xlsx'),
    '*',
    'ETALON SV005'
  );
}
main();
