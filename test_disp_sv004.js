const ExcelJS = require('exceljs');
const path = require('path');

const base = __dirname;
const TEMPLATE_SV004 = path.join(base, 'templateSV004_test.xlsx');
const ETALON_SV004 = path.join(base, '(г.Курган_ ул.Карла Маркса_ 149)_2026-07-03_SV004_fixed (1).xlsx');

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
        for (let r = row + 1; r <= Math.min(row + 30, lastRow); r++) {
          const cellEnext = worksheet.getCell(`E${r}`);
          if (getCellText(cellEnext).trim() === 'Стойка') {
            endRow = r - 1;
            break;
          }
        }
        if (endRow === row) endRow = row + 19;
        blocks.push({ startRow: row, endRow, type });
      }
    }
  }
  return blocks;
}

function getNumericCellValue(cell) {
  const v = cell.value;
  if (typeof v === 'number') return v;
  if (v && v.result !== undefined && v.result !== null) return parseFloat(v.result) || NaN;
  return NaN;
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
  const groups = [{ offsetDev: 14, offsetRoom: 15, pins: [0,1,2,3] }];
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

function fillSheetsSV004(workbook, boards, globalModel) {
  const maxPlinthsPerController = 16;
  const allPlinths = [];
  boards.forEach((board, idx) => {
    allPlinths.push({ boardIndex: idx, plinth: board.plinth1, type: 'input' });
    allPlinths.push({ boardIndex: idx, plinth: board.plinth2, type: 'output' });
  });
  const groups = [];
  for (let i = 0; i < allPlinths.length; i += maxPlinthsPerController) {
    groups.push(allPlinths.slice(i, i + maxPlinthsPerController));
  }
  const sheetNames = workbook.worksheets.map(s => s.name);
  for (let g = 0; g < groups.length; g++) {
    const groupPlinths = groups[g];
    const cheatSheetName = `шпоргалка-${g+1}`;
    const dispSheetName = `Disp-${g+1}`;
    if (!sheetNames.includes(cheatSheetName) || !sheetNames.includes(dispSheetName)) {
      console.warn(`Листы ${cheatSheetName} или ${dispSheetName} не найдены, пропускаем`);
      continue;
    }
    const cheatSheet = workbook.getWorksheet(cheatSheetName);
    const dispSheet = workbook.getWorksheet(dispSheetName);
    const cheatLastRow = cheatSheet.rowCount;
    for (let r = 4; r <= cheatLastRow; r++) {
      for (let col = 1; col <= 7; col++) {
        cheatSheet.getCell(r, col).value = null;
      }
    }
    const dispLastRow = dispSheet.rowCount;
    for (let r = 2; r <= dispLastRow; r++) {
      for (let col = 2; col <= 9; col++) {
        dispSheet.getCell(r, col).value = null;
      }
    }
    let rowIdx = 4;
    let counter = 1;
    for (let i = 0; i < groupPlinths.length; i++) {
      const plinthData = groupPlinths[i].plinth;
      const plinthNumber = plinthData.number;
      const tm = plinthData.terminalMap || {};
      const cm = plinthData.cableMap || {};
      const rm = plinthData.roomMap || {};
      for (let pin = 0; pin <= 3; pin++) {
        const device = tm[pin] || '';
        const cable = cm[pin] || '';
        let room = (rm[pin] && rm[pin].trim()) ? rm[pin] : '';
        if (!device) {
          room = 'Резерв';
        }
        cheatSheet.getCell(`A${rowIdx}`).value = '';
        cheatSheet.getCell(`B${rowIdx}`).value = plinthNumber;
        cheatSheet.getCell(`C${rowIdx}`).value = `Пин${pin}`;
        cheatSheet.getCell(`D${rowIdx}`).value = device;
        cheatSheet.getCell(`E${rowIdx}`).value = room ? (room === 'Резерв' ? 'Резерв' : `пом. ${room}`) : '';
        cheatSheet.getCell(`F${rowIdx}`).value = cable ? `ШЛ.${cable}` : '';
        cheatSheet.getCell(`G${rowIdx}`).value = counter++;
        rowIdx++;
      }
      rowIdx++;
    }
    const cheatData = [];
    for (let r = 4; r < rowIdx; r++) {
      const b = cheatSheet.getCell(`B${r}`).value;
      const c = cheatSheet.getCell(`C${r}`).value;
      const d = cheatSheet.getCell(`D${r}`).value;
      const e = cheatSheet.getCell(`E${r}`).value;
      const f = cheatSheet.getCell(`F${r}`).value;
      const g = cheatSheet.getCell(`G${r}`).value;
      if (b || c || d || e || f || g) {
        cheatData.push({ b, c, d, e, f, g });
      }
    }
    const firstBoard = groupPlinths.length > 0 ? boards[groupPlinths[0].boardIndex] : null;
    const rackDisplay = firstBoard ? String(firstBoard.rack || '').trim() : '';
    function formatDispText(item) {
      if (!item) return 'Резерв';
      let roomNum = '';
      if (item.e && typeof item.e === 'string') {
        if (item.e === 'Резерв') {
          return 'Резерв';
        } else if (item.e.startsWith('пом. ')) {
          roomNum = item.e.substring(5).trim();
        }
      }
      if (roomNum) {
        const deviceText = item.d ? String(item.d) : '';
        return `п.${roomNum} - ${deviceText}`;
      }
      return 'Резерв';
    }
    function fillPanel(dispSheet, startRow, startNum) {
      for (let i = 0; i < 16; i++) {
        const row = startRow + i;
        const leftIdx = startNum - 1 + i;
        const leftItem = cheatData[leftIdx];
        const leftText = leftItem ? formatDispText(leftItem) : 'Резерв';
        const leftNum = leftItem && leftItem.g ? String(leftItem.g) : String(startNum + i);
        dispSheet.getCell(`B${row}`).value = leftNum;
        dispSheet.getCell(`C${row}`).value = leftText;
        dispSheet.getCell(`D${row}`).value = leftNum;
        dispSheet.getCell(`E${row}`).value = leftText;
        const rightItemNum = startNum + 16 + i;
        if (i === 15) {
          dispSheet.getCell(`F${row}`).value = String(rightItemNum);
          dispSheet.getCell(`G${row}`).value = rackDisplay;
          dispSheet.getCell(`H${row}`).value = String(rightItemNum);
          dispSheet.getCell(`I${row}`).value = 'IP';
        } else {
          const rightIdx = startNum - 1 + 16 + i;
          const rightItem = cheatData[rightIdx];
          const rightText = rightItem ? formatDispText(rightItem) : 'Резерв';
          const rightNum = rightItem && rightItem.g ? String(rightItem.g) : String(rightItemNum);
          dispSheet.getCell(`F${row}`).value = rightNum;
          dispSheet.getCell(`G${row}`).value = rightText;
          dispSheet.getCell(`H${row}`).value = rightNum;
          dispSheet.getCell(`I${row}`).value = rightText;
        }
      }
    }
    fillPanel(dispSheet, 2, 1);
    fillPanel(dispSheet, 21, 33);
    dispSheet.getCell(`K2`).value = `Display panel №${g+1}`;
  }
}

// Test boards matching etalon exactly (3 boards, 6 plinths)
const testBoards = [
  {
    rack: 'ХК 2.1',
    plinth1: {
      number: 13,
      terminalMap: {0:'КАВ (1)',1:'КАВ (1)',2:'КАВ (1)',3:'КАВ (1)'},
      cableMap: {0:'10',1:'43',2:'55',3:'130'},
      roomMap: {0:'14',1:'17',2:'17',3:'20'}
    },
    plinth2: {
      number: 14,
      terminalMap: {0:'КАВ (1)',1:'КАВ (1)',2:'ИК (1)',3:'ИК (1)'},
      cableMap: {0:'109',1:'163',2:'16',3:'22'},
      roomMap: {0:'19',1:'23',2:'14',3:'15'}
    }
  },
  {
    rack: 'ХК 2.1',
    plinth1: {
      number: 15,
      terminalMap: {0:'ИК (1)',1:'ИК (6)',2:'ИК (1)',3:'СМК (1), ИК (1)'},
      cableMap: {0:'28',1:'64',2:'94',3:'170'},
      roomMap: {0:'16',1:'17',2:'18',3:'ДГУ'}
    },
    plinth2: {
      number: 16,
      terminalMap: {0:'ИК (1)',1:'ИК (4)',2:'ИК (1)',3:'ДРС (1)'},
      cableMap: {0:'115',1:'136',2:'169',3:'19'},
      roomMap: {0:'19',1:'20',2:'24',3:'14'}
    }
  },
  {
    rack: 'ХК 2.1',
    plinth1: {
      number: 17,
      terminalMap: {0:'ДРС (1)',1:'ДРС (1)',2:'ДРС (2)',3:'ДРС (1)'},
      cableMap: {0:'25',1:'31',2:'64',3:'97'},
      roomMap: {0:'15',1:'16',2:'17',3:'18'}
    },
    plinth2: {
      number: 18,
      terminalMap: {0:'ДРС (1)',1:'',2:'',3:'ОПС'},
      cableMap: {0:'118',1:'',2:'',3:''},
      roomMap: {0:'19',1:'',2:'',3:''}
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
    copyBlock(worksheet, lastInput.startRow, lastInput.endRow, lastBlockEnd + 1);
    allBlocks = getBlocksSV004(worksheet);
    allBlocks.sort((a, b) => a.startRow - b.startRow);
    inputBlocks = allBlocks.filter(b => b.type === 'input');
    outputBlocks = allBlocks.filter(b => b.type === 'output');
  }

  while (outputBlocks.length < neededPairs) {
    const lastOutput = outputBlocks[outputBlocks.length - 1];
    const lastBlockEnd = Math.max(...allBlocks.map(b => b.endRow));
    copyBlock(worksheet, lastOutput.startRow, lastOutput.endRow, lastBlockEnd + 1);
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
      rack: board.rack, boardNumber, plinthNumber: board.plinth1.number,
      terminalMap: board.plinth1.terminalMap, cableMap: board.plinth1.cableMap,
      roomMap: board.plinth1.roomMap || {}
    }, 'SV 777');
    await fillPlinthBlockSV004(worksheet, usedOutputBlocks[i].startRow, {
      rack: board.rack, boardNumber, plinthNumber: board.plinth2.number,
      terminalMap: board.plinth2.terminalMap, cableMap: board.plinth2.cableMap,
      roomMap: board.plinth2.roomMap || {}
    }, 'SV 777');
  }

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

  // Now fill the sheets (шпоргалка, Disp)
  fillSheetsSV004(workbook, testBoards, 'SV 777');

  const outPath = path.join(base, 'test_output_sv004.xlsx');
  await workbook.xlsx.writeFile(outPath);
  console.log('Generated SV004 saved');

  // Compare Disp-1 with etalon
  const etalonWb = new ExcelJS.Workbook();
  await etalonWb.xlsx.readFile(ETALON_SV004);
  const etalonDisp = etalonWb.getWorksheet('Disp-1');
  const genDisp = workbook.getWorksheet('Disp-1');

  console.log('\n=== COMPARISON Disp-1 ===');
  const maxRows = Math.max(genDisp.rowCount, etalonDisp.rowCount);
  let totalDiffs = 0;
  for (let r = 1; r <= maxRows; r++) {
    const diffs = [];
    for (let c = 1; c <= 12; c++) {
      const gv = getCellText(genDisp.getCell(r, c));
      const ev = getCellText(etalonDisp.getCell(r, c));
      if (gv !== ev) {
        diffs.push({ c, gv, ev });
      }
    }
    if (diffs.length) {
      totalDiffs++;
      const diffStr = diffs.slice(0, 6).map(d => `C${d.c}(${d.gv}|${d.ev})`).join(', ');
      console.log(`R${r}: ${diffStr}${diffs.length > 6 ? `...(+${diffs.length-6})` : ''}`);
    }
  }
  console.log(`\nTotal differing rows: ${totalDiffs}`);

  // Print Disp-1 from generated file
  console.log('\n=== GENERATED Disp-1 ===');
  for (let r = 1; r <= genDisp.rowCount; r++) {
    const cols = [];
    for (let c = 1; c <= 12; c++) {
      const v = getCellText(genDisp.getCell(r, c));
      if (v) cols.push(`C${c}=${v}`);
    }
    if (cols.length) console.log(`R${r}: ${cols.join(', ')}`);
  }
}

main().catch(console.error);
