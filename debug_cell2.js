const ExcelJS = require('exceljs');
const path = require('path');

async function main() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, '(г.Курган_ ул.Карла Маркса_ 149)_2026-07-03_SV005_fixed (1).xlsx'));
  const ws = wb.getWorksheet('Лист1');
  
  // Find ALL rows with non-null column O
  console.log('=== Rows with non-null column O(15) in etalon SV005 ===');
  for (let r = 1; r <= ws.rowCount; r++) {
    const cell = ws.getCell(r, 15);
    if (cell.value !== null && cell.value !== undefined) {
      const e = ws.getCell(r, 5);
      console.log(`R${r}: E="${e.value}", O=${JSON.stringify(cell.value)}`);
    }
  }
  
  // Also check for formula vs value in column O for SV004 etalon
  console.log('\n=== SV004: Check RS-485 right side formulas ===');
  const wb4 = new ExcelJS.Workbook();
  await wb4.xlsx.readFile(path.join(__dirname, '(г.Курган_ ул.Карла Маркса_ 149)_2026-07-03_SV004_fixed (1).xlsx'));
  const ws4 = wb4.getWorksheet('SV777-1 (SV004)');
  
  // Check if C38/C39 have formulas or values
  for (const r of [2,3,4,5,6]) {
    const c38 = ws4.getCell(r, 38);
    const c39 = ws4.getCell(r, 39);
    console.log(`R${r}: C38=${JSON.stringify(c38.value)}, C39=${JSON.stringify(c39.value)}`);
  }
  
  // Check template SV004 for the same
  console.log('\n=== SV004 Template: Check RS-485 right side ===');
  const wb4t = new ExcelJS.Workbook();
  await wb4t.xlsx.readFile(path.join(__dirname, 'templateSV004_test.xlsx'));
  const ws4t = wb4t.getWorksheet('SV777-1 (SV004)');
  
  for (const r of [2,3,4,5,6]) {
    const c38 = ws4t.getCell(r, 38);
    const c39 = ws4t.getCell(r, 39);
    console.log(`R${r}: C38=${JSON.stringify(c38.value)}, C39=${JSON.stringify(c39.value)}`);
  }
  
  // Check template SV004 for column O data
  console.log('\n=== SV004 Template: Column O values ===');
  for (let r = 1; r <= 50; r++) {
    const o = ws4t.getCell(r, 15);
    if (o.value !== null && o.value !== undefined) {
      console.log(`R${r}: O=${JSON.stringify(o.value)}`);
    }
  }
}
main().catch(console.error);
