const ExcelJS = require('exceljs');
const path = require('path');

async function main() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, '(г.Курган_ ул.Карла Маркса_ 149)_2026-07-03_SV005_fixed (1).xlsx'));
  const ws = wb.getWorksheet('Лист1');
  
  // Check R2 (Стойка), columns N(14), O(15), M(13)
  for (const r of [1,2,3,4,5,6,7,19,20,21,22,23,24,36,37,38,39,40,41]) {
    const n = ws.getCell(r, 14);
    const o = ws.getCell(r, 15);
    const m = ws.getCell(r, 13);
    console.log(`R${r}: M=${JSON.stringify(m.value)} N=${JSON.stringify(n.value)} O=${JSON.stringify(o.value)}`);
  }
  
  // Also check for any formulas in column O
  let formulaCount = 0;
  for (let r = 1; r <= ws.rowCount; r++) {
    const cell = ws.getCell(r, 15);
    if (cell.value && typeof cell.value === 'object' && cell.value.formula) {
      formulaCount++;
      console.log(`R${r} O: formula=${cell.value.formula}, result=${cell.value.result}`);
    }
  }
  console.log(`Formula count in column O: ${formulaCount}`);
}
main().catch(console.error);
