const ExcelJS = require('exceljs');
const path = require('path');

const ETALON = path.join(__dirname, '(г.Курган_ ул.Карла Маркса_ 149)_2026-07-03_SV004_fixed (1).xlsx');

function getCellText(cell) {
  if (!cell || cell.value === undefined || cell.value === null) return '';
  if (typeof cell.value === 'string') return cell.value;
  if (typeof cell.value === 'object') {
    if (cell.value.richText) return cell.value.richText.map(rt => rt.text).join('');
    if (cell.value.text) return cell.value.text;
    if (cell.value.result !== undefined && cell.value.result !== null) return String(cell.value.result);
    if (cell.value.formula !== undefined) return 'formula:' + cell.value.formula;
    return JSON.stringify(cell.value);
  }
  return String(cell.value);
}

async function main() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(ETALON);
  
  const cheat = wb.getWorksheet('шпоргалка-1');
  console.log('=== ETALON шпоргалка-1 (cols A-G, rows 4+) ===');
  for (let r = 4; r <= cheat.rowCount; r++) {
    const cols = [];
    for (let c = 1; c <= 7; c++) {
      const v = getCellText(cheat.getCell(r, c));
      if (v) cols.push(`${String.fromCharCode(64+c)}=${v}`);
    }
    if (cols.length) console.log(`R${r}: ${cols.join(', ')}`);
  }
  
  console.log('\n=== ETALON Disp-1 (cols A-L, rows 2+) ===');
  const disp = wb.getWorksheet('Disp-1');
  for (let r = 2; r <= disp.rowCount; r++) {
    const cols = [];
    for (let c = 1; c <= 12; c++) {
      const v = getCellText(disp.getCell(r, c));
      if (v) cols.push(`${String.fromCharCode(64+c)}=${v}`);
    }
    if (cols.length) console.log(`R${r}: ${cols.join(', ')}`);
  }
}

main().catch(console.error);
