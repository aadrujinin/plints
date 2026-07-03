const ExcelJS = require('exceljs');
const path = require('path');

async function main() {
  const genWb = new ExcelJS.Workbook();
  await genWb.xlsx.readFile(path.join(__dirname, 'test_output_sv005.xlsx'));
  const genWs = genWb.getWorksheet('Лист1');
  
  const etaWb = new ExcelJS.Workbook();
  await etaWb.xlsx.readFile(path.join(__dirname, '(г.Курган_ ул.Карла Маркса_ 149)_2026-07-03_SV005_fixed (1).xlsx'));
  const etaWs = etaWb.getWorksheet('Лист1');
  
  // Compare rows with differences
  for (const r of [2,19,22,24,36,56,58,70,90,92,104,124,126,138,158,160,172,189,192]) {
    const go = genWs.getCell(r, 15).value;
    const eo = etaWs.getCell(r, 15).value;
    const ge = genWs.getCell(r, 5).value;
    const ee = etaWs.getCell(r, 5).value;
    console.log(`R${r}: GEN E="${ge}", O=${JSON.stringify(go)} | ETA E="${ee}", O=${JSON.stringify(eo)}`);
  }
}
main().catch(console.error);
