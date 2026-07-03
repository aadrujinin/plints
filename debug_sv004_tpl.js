const ExcelJS = require('exceljs');
const path = require('path');

async function main() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, 'templateSV004_test.xlsx'));
  const ws = wb.getWorksheet('SV777-1 (SV004)');
  
  // Check row 2 in detail - all cells with values
  console.log('=== Row 2 cells (template) ===');
  const row2 = ws.getRow(2);
  let cellCount = 0;
  row2.eachCell((cell, col) => {
    cellCount++;
    console.log(`  C${col}: value=${JSON.stringify(cell.value)}, type=${typeof cell.value}`);
  });
  console.log(`Total cells in row 2: ${cellCount}`);
  
  // Also check specific cells
  for (const c of [24, 25, 29, 37, 38, 39, 40]) {
    const cell = ws.getCell(2, c);
    console.log(`C${c}: value=${JSON.stringify(cell.value)}, type=${typeof cell.value}, formula=${cell.formula}`);
  }
  
  // Check row 1 separator
  console.log('\n=== Row 1 cells ===');
  const row1 = ws.getRow(1);
  row1.eachCell((cell, col) => {
    console.log(`  C${col}: value=${JSON.stringify(cell.value)}`);
  });
  
  // Check if columns 24-44 have any content at all
  console.log('\n=== Checking right side content ===');
  for (let r = 1; r <= 30; r++) {
    for (const c of [24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44]) {
      const cell = ws.getCell(r, c);
      if (cell.value !== null && cell.value !== undefined) {
        console.log(`R${r} C${c}: ${JSON.stringify(cell.value).substring(0, 80)}`);
      }
    }
  }
}
main().catch(console.error);
