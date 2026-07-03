const ExcelJS = require('exceljs');
const path = require('path');

const base = __dirname;
const TEMPLATE_SV004 = path.join(base, 'templateSV004_test.xlsx');

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

function getNumericCellValue(cell) {
  const v = cell.value;
  if (typeof v === 'number') return v;
  if (v && v.result !== undefined && v.result !== null) return parseFloat(v.result) || NaN;
  return NaN;
}

function getBlocksSV004(worksheet) {
  if (!worksheet) return [];
  const blocks = [];
  const lastRow = worksheet.rowCount;
  for (let row = 1; row <= lastRow; row++) {
    const cellE = worksheet.getCell(`E${row}`);
    const text = getCellText(cellE);
    if (text.trim() === 'Стойка') {
      let plinthNum = NaN;
      for (let offset = 1; offset <= 10; offset++) {
        const checkRow = row + offset;
        if (checkRow > lastRow) break;
        const cellE2 = worksheet.getCell(`E${checkRow}`);
        if (getCellText(cellE2).trim() === 'Номер плинта') {
          const cellN2 = worksheet.getCell(`N${checkRow}`);
          plinthNum = getNumericCellValue(cellN2);
          break;
        }
      }
      if (!isNaN(plinthNum) && plinthNum > 0) {
        const type = (plinthNum % 2 === 1) ? 'input' : 'output';
        let endRow = row;
        let found = false;
        for (let r = row + 1; r <= Math.min(row + 30, lastRow); r++) {
          const cellEnext = worksheet.getCell(`E${r}`);
          if (getCellText(cellEnext).trim() === 'Стойка') {
            endRow = r - 1;
            found = true;
            break;
          }
        }
        if (!found) endRow = row + 19;
        blocks.push({ startRow: row, endRow, type });
      }
    }
  }
  return blocks;
}

function copyBlock(worksheet, srcStart, srcEnd, targetStart) {
  const rowCount = srcEnd - srcStart + 1;
  for (let r = 0; r < rowCount; r++) {
    worksheet.spliceRows(targetStart + r, 0, []);
    const srcRow = worksheet.getRow(srcStart + r);
    const dstRow = worksheet.getRow(targetStart + r);
    srcRow.eachCell((cell, colNumber) => {
      dstRow.getCell(colNumber).value = cell.value;
    });
    dstRow.height = srcRow.height;
  }
}

async function fillPlinthBlockSV004(worksheet, startRow, plinthData, globalModel) {
  const groups = [
    { offsetDev: 14, offsetRoom: 15, pins: [0,1,2,3] }
  ];

  const rackClean = plinthData.rack.replace(/^ХК\s*/i, '');
  worksheet.getCell(`N${startRow}`).value = rackClean;
  worksheet.getCell(`O${startRow}`).value = rackClean;

  const controllerRow = startRow + 1;
  const controllerCell = worksheet.getCell(`E${controllerRow}`);
  let controllerText = getCellText(controllerCell);
  if (controllerText.includes('SV 777')) {
    controllerCell.value = controllerText.replace('SV 777', globalModel);
  }
  const controllerVal = `ХК ${rackClean}.1`;
  worksheet.getCell(`N${controllerRow}`).value = controllerVal;
  worksheet.getCell(`O${controllerRow}`).value = controllerVal;

  const boardRow = startRow + 2;
  worksheet.getCell(`N${boardRow}`).value = plinthData.boardNumber;

  const holderRow = startRow + 3;
  const absNum = plinthData.plinthNumber;
  const group = Math.floor((absNum - 1) / 15) + 1;
  const holderValue = `ХВ-${rackClean}.${group}`;
  worksheet.getCell(`N${holderRow}`).value = holderValue;

  const plinthNumRow = startRow + 4;
  worksheet.getCell(`N${plinthNumRow}`).value = plinthData.plinthNumber;

  const tm = plinthData.terminalMap || {};
  const cm = plinthData.cableMap || {};
  const rm = plinthData.roomMap || {};

  for (const group of groups) {
    const devRow = startRow + group.offsetDev;
    const roomRow = startRow + group.offsetRoom;
    if (devRow > worksheet.rowCount || roomRow > worksheet.rowCount) continue;
    for (let i = 0; i < group.pins.length; i++) {
      const pin = group.pins[i];
      const colDevice = ['B','F','J','N'][i];
      const colCable = ['D','H','L','P'][i];
      const device = tm[pin] || '';
      const cable = cm[pin] || '';
      let room = (rm[pin] && rm[pin].trim()) ? rm[pin] : '';

      if (!device) room = 'Резерв';

      worksheet.getCell(`${colDevice}${devRow}`).value = device;
      worksheet.getCell(`${colCable}${devRow}`).value = cable;

      if (room === 'Резерв') {
        worksheet.getCell(`${colDevice}${roomRow}`).value = 'Резерв';
        worksheet.getCell(`${colCable}${roomRow}`).value = '';
      } else if (room) {
        worksheet.getCell(`${colDevice}${roomRow}`).value = 'пом.';
        worksheet.getCell(`${colCable}${roomRow}`).value = room;
      } else {
        worksheet.getCell(`${colDevice}${roomRow}`).value = '';
        worksheet.getCell(`${colCable}${roomRow}`).value = '';
      }
    }
  }
}

// Build test data matching SV004 etalon (3 boards, 6 plinths)
const testBoards = [
  {
    rack: 'ХК 2.1',
    plinth1: {
      number: 13,
      terminalMap: {0:'КАВ (1)',1:'КАВ (1)',2:'КАВ (1)',3:'КАВ (1)'},
      cableMap: {0:'10',1:'43',2:'55',3:'130'},
      roomMap: {0:'14',1:'17',2:'17',3:'20'},
      skud: '2.1',
      room: '14'
    },
    plinth2: {
      number: 14,
      terminalMap: {0:'КАВ (1)',1:'КАВ (1)',2:'ИК (1)',3:'ИК (1)'},
      cableMap: {0:'109',1:'163',2:'16',3:'22'},
      roomMap: {0:'19',1:'23',2:'14',3:'15'},
      skud: '2.1',
      room: '19'
    }
  },
  {
    rack: 'ХК 2.1',
    plinth1: {
      number: 15,
      terminalMap: {0:'ИК (1)',1:'ИК (6)',2:'ИК (1)',3:'СМК (1), ИК (1)'},
      cableMap: {0:'28',1:'64',2:'94',3:'170'},
      roomMap: {0:'16',1:'17',2:'18',3:'ДГУ'},
      skud: '2.1',
      room: '16'
    },
    plinth2: {
      number: 16,
      terminalMap: {0:'ИК (1)',1:'ИК (4)',2:'ИК (1)',3:'ДРС (1)'},
      cableMap: {0:'115',1:'136',2:'169',3:'19'},
      roomMap: {0:'19',1:'20',2:'24',3:'14'},
      skud: '2.1',
      room: '19'
    }
  },
  {
    rack: 'ХК 2.1',
    plinth1: {
      number: 17,
      terminalMap: {0:'ДРС (1)',1:'ДРС (1)',2:'ДРС (2)',3:'ДРС (1)'},
      cableMap: {0:'25',1:'31',2:'64',3:'97'},
      roomMap: {0:'15',1:'16',2:'17',3:'18'},
      skud: '2.1',
      room: '15'
    },
    plinth2: {
      number: 18,
      terminalMap: {0:'ДРС (1)',1:'',2:'',3:''},
      cableMap: {0:'118',1:'',2:'',3:''},
      roomMap: {0:'19',1:'',2:'',3:''},
      skud: '2.1',
      room: '19'
    }
  }
];

async function main() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(TEMPLATE_SV004);
  const worksheet = workbook.getWorksheet('SV777-1 (SV004)');

  let allBlocks = getBlocksSV004(worksheet);
  allBlocks.sort((a, b) => a.startRow - b.startRow);

  const neededPairs = testBoards.length;
  let inputBlocks = allBlocks.filter(b => b.type === 'input');
  let outputBlocks = allBlocks.filter(b => b.type === 'output');

  while (inputBlocks.length < neededPairs) {
    const lastInput = inputBlocks[inputBlocks.length - 1];
    const lastBlockEnd = Math.max(...allBlocks.map(b => b.endRow));
    const targetStart = lastBlockEnd + 1;
    copyBlock(worksheet, lastInput.startRow, lastInput.endRow, targetStart);
    allBlocks = getBlocksSV004(worksheet);
    allBlocks.sort((a, b) => a.startRow - b.startRow);
    inputBlocks = allBlocks.filter(b => b.type === 'input');
    outputBlocks = allBlocks.filter(b => b.type === 'output');
  }

  while (outputBlocks.length < neededPairs) {
    const lastOutput = outputBlocks[outputBlocks.length - 1];
    const lastBlockEnd = Math.max(...allBlocks.map(b => b.endRow));
    const targetStart = lastBlockEnd + 1;
    copyBlock(worksheet, lastOutput.startRow, lastOutput.endRow, targetStart);
    allBlocks = getBlocksSV004(worksheet);
    allBlocks.sort((a, b) => a.startRow - b.startRow);
    inputBlocks = allBlocks.filter(b => b.type === 'input');
    outputBlocks = allBlocks.filter(b => b.type === 'output');
  }

  const usedInputBlocks = inputBlocks.slice(0, neededPairs);
  const usedOutputBlocks = outputBlocks.slice(0, neededPairs);

  for (let i = 0; i < testBoards.length; i++) {
    const board = testBoards[i];
    const boardNumber = i + 1;

    await fillPlinthBlockSV004(worksheet, usedInputBlocks[i].startRow, {
      rack: board.rack,
      boardNumber,
      plinthNumber: board.plinth1.number,
      skud: board.plinth1.skud,
      room: board.plinth1.room,
      terminalMap: board.plinth1.terminalMap,
      cableMap: board.plinth1.cableMap,
      roomMap: board.plinth1.roomMap || {}
    }, 'SV 777');

    await fillPlinthBlockSV004(worksheet, usedOutputBlocks[i].startRow, {
      rack: board.rack,
      boardNumber,
      plinthNumber: board.plinth2.number,
      skud: board.plinth2.skud,
      room: board.plinth2.room,
      terminalMap: board.plinth2.terminalMap,
      cableMap: board.plinth2.cableMap,
      roomMap: board.plinth2.roomMap || {}
    }, 'SV 777');
  }

  // Delete unused blocks
  const allBlocksFinal = getBlocksSV004(worksheet);
  const usedStartRows = new Set();
  usedInputBlocks.forEach(b => usedStartRows.add(b.startRow));
  usedOutputBlocks.forEach(b => usedStartRows.add(b.startRow));
  const blocksToDelete = allBlocksFinal.filter(b => !usedStartRows.has(b.startRow));
  blocksToDelete.sort((a, b) => b.startRow - a.startRow);
  for (const block of blocksToDelete) {
    const rowCount = block.endRow - block.startRow + 1;
    worksheet.spliceRows(block.startRow, rowCount);
  }

  const lastUsedEndRow = Math.max(
    ...usedInputBlocks.map(b => b.endRow),
    ...usedOutputBlocks.map(b => b.endRow)
  );
  if (lastUsedEndRow > 0) {
    const totalRows = worksheet.rowCount;
    if (totalRows > lastUsedEndRow) {
      worksheet.spliceRows(lastUsedEndRow + 1, totalRows - lastUsedEndRow);
    }
  }

  // Save generated file
  const outPath = path.join(base, 'test_output_sv004.xlsx');
  await workbook.xlsx.writeFile(outPath);
  console.log('Generated SV004 saved to ' + outPath);

  // Compare with etalon
  const etalonWb = new ExcelJS.Workbook();
  await etalonWb.xlsx.readFile(path.join(base, '(г.Курган_ ул.Карла Маркса_ 149)_2026-07-03_SV004_fixed (1).xlsx'));
  const etalonSheet = etalonWb.getWorksheet('SV777-1 (SV004)');
  const genSheet = workbook.getWorksheet('SV777-1 (SV004)');

  console.log('\n=== COMPARISON SV004 SV777-1 ===');
  console.log('Generated rows: ' + genSheet.rowCount + ', Etalon rows: ' + etalonSheet.rowCount);

  let differences = 0;
  const maxRows = Math.max(genSheet.rowCount, etalonSheet.rowCount);
  for (let r = 1; r <= maxRows; r++) {
    const diffs = [];
    for (const c of [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40]) {
      const gv = getCellText(genSheet.getCell(r, c));
      const ev = getCellText(etalonSheet.getCell(r, c));
      if (gv !== ev) {
        diffs.push({c, gv, ev});
      }
    }
    if (diffs.length) {
      differences++;
      const diffStr = diffs.slice(0, 5).map(d => `C${d.c}(${d.gv} vs ${d.ev})`).join(', ');
      console.log(`R${r}: ${diffStr}${diffs.length > 5 ? `... (+${diffs.length-5})` : ''}`);
    }
  }
  console.log(`\nTotal differences: ${differences}`);
}

main().catch(console.error);
