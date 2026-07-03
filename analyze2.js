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
        return JSON.stringify(cell.value);
    }
    return String(cell.value);
}

async function dumpSheet(file, sheetName, label) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  const ws = wb.getWorksheet(sheetName);
  if (!ws) {
    console.log('Sheet ' + sheetName + ' not found in ' + label);
    return;
  }
  console.log('=== ' + label + ' / ' + sheetName + ' === rows=' + ws.rowCount);
  
  // Find all rows that have any content
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
  
  // Print all rows that have content
  for (let r = 1; r <= lastRowWithContent; r++) {
    const row = ws.getRow(r);
    const vals = [];
    // Focus on columns 1-22 (left block) and important columns
    for (const c of [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44]) {
      if (c > (ws.columnCount || 44)) continue;
      const cell = row.getCell(c);
      let v = getCellText(cell);
      if (v) vals.push('C' + c + (c===14?'(N)':c===15?'(O)':c===13?'(M)':'') + '=' + v.substring(0, 60));
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
  // Dump etalon SV005 Лист1 all rows
  await dumpSheet(
    path.join(base, '(г.Курган_ ул.Карла Маркса_ 149)_2026-07-03_SV005_fixed (1).xlsx'),
    'Лист1',
    'ETALON SV005'
  );
  
  // Dump etalon SV004 SV777-1 (SV004) all rows
  await dumpSheet(
    path.join(base, '(г.Курган_ ул.Карла Маркса_ 149)_2026-07-03_SV004_fixed (1).xlsx'),
    'SV777-1 (SV004)',
    'ETALON SV004'
  );
  
  // Dump template SV005 Лист1 - first 50 rows to see block structure
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(base, 'templateSV005_test.xlsx'));
  const ws = wb.getWorksheet('Лист1');
  console.log('=== TEMPLATE SV005 / Лист1 === rows=' + ws.rowCount);
  for (let r = 1; r <= 50; r++) {
    const row = ws.getRow(r);
    const vals = [];
    for (const c of [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22]) {
      if (c > (ws.columnCount || 22)) continue;
      const cell = row.getCell(c);
      let v = getCellText(cell);
      if (v) vals.push('C' + c + (c===14?'(N)':c===15?'(O)':c===13?'(M)':'') + '=' + v.substring(0, 60));
    }
    if (vals.length) console.log('  R' + r + ': ' + vals.join(' | '));
  }
  
  // Dump template SV004 SV777-1(SV004) first 50 rows
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.readFile(path.join(base, 'templateSV004_test.xlsx'));
  const ws2 = wb2.getWorksheet('SV777-1 (SV004)');
  console.log('=== TEMPLATE SV004 / SV777-1 (SV004) === rows=' + ws2.rowCount);
  for (let r = 1; r <= 50; r++) {
    const row = ws2.getRow(r);
    const vals = [];
    for (const c of [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40]) {
      if (c > (ws2.columnCount || 40)) continue;
      const cell = row.getCell(c);
      let v = getCellText(cell);
      if (v) vals.push('C' + c + (c===14?'(N)':c===15?'(O)':c===13?'(M)':'') + '=' + v.substring(0, 60));
    }
    if (vals.length) console.log('  R' + r + ': ' + vals.join(' | '));
  }
}
main();
