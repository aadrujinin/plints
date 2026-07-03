const ExcelJS = require('exceljs');
const path = require('path');

const base = __dirname;
const files = {
  tpl005: path.join(base, 'templateSV005_test.xlsx'),
  tpl004: path.join(base, 'templateSV004_test.xlsx'),
  etal005: path.join(base, '(г.Курган_ ул.Карла Маркса_ 149)_2026-07-03_SV005_fixed (1).xlsx'),
  etal004: path.join(base, '(г.Курган_ ул.Карла Маркса_ 149)_2026-07-03_SV004_fixed (1).xlsx')
};

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

async function analyze(file, label) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  console.log('=== ' + label + ' ===');
  console.log('Sheets:', wb.worksheets.map(s => s.name));
  wb.worksheets.forEach(ws => {
    console.log('  Sheet "' + ws.name + '": rows=' + ws.rowCount + ', cols=' + (ws.columnCount || '?'));
    for (let r = 1; r <= Math.min(ws.rowCount, 12); r++) {
      const row = ws.getRow(r);
      const vals = [];
      for (let c = 1; c <= (ws.columnCount || 25); c++) {
        const cell = row.getCell(c);
        let v = getCellText(cell);
        if (v) vals.push('[' + c + ']' + v.substring(0, 120));
      }
      if (vals.length) console.log('    R' + r + ': ' + vals.join(' | '));
    }
  });
  console.log('');
}

async function main() {
  for (const [k, v] of Object.entries(files)) {
    try {
      await analyze(v, k);
    } catch(e) {
      console.log('Error ' + k + ': ' + e.message);
    }
  }
}
main();
