const ExcelJS = require('exceljs');
const path = require('path');

const base = __dirname;
const TEMPLATE_SV005 = path.join(base, 'templateSV005_test.xlsx');
const TEMPLATE_SV004 = path.join(base, 'templateSV004_test.xlsx');

function getCellText(cell) {
    if (!cell || cell.value === undefined || cell.value === null) return '';
    if (typeof cell.value === 'string') return cell.value;
    if (typeof cell.value === 'object') {
        if (cell.value.richText) return cell.value.richText.map(rt => rt.text).join('');
        if (cell.value.text) return cell.value.text;
        if (cell.value.result !== undefined) return String(cell.value.result);
        if (cell.value.formula) return 'formula:' + cell.value.formula;
        return JSON.stringify(cell.value);
    }
    return String(cell.value);
}

async function cloneSheet(srcSheet) {
    const newSheet = srcSheet.workbook.addWorksheet(srcSheet.name + '_copy');
    for (let r = 1; r <= srcSheet.rowCount; r++) {
        const srcRow = srcSheet.getRow(r);
        const dstRow = newSheet.getRow(r);
        srcRow.eachCell((cell, col) => {
            dstRow.getCell(col).value = cell.value;
        });
    }
    return newSheet;
}

function findRowBySubstring(worksheet, startRow, substring) {
    for (let r = startRow; r <= startRow + 20; r++) {
        const cell = worksheet.getCell(`E${r}`);
        const text = getCellText(cell);
        if (text.includes(substring)) return r;
    }
    return null;
}

function getBlockType(worksheet, headerRow) {
    const cell = worksheet.getCell(`E${headerRow}`);
    const text = getCellText(cell);
    if (text.includes('R1 и D1')) return 'input';
    if (text.includes('R2 и D2')) return 'output';
    return 'R1/D1';
}

async function getBlocksSV005(worksheet) {
    if (!worksheet) return [];
    const blocks = [];
    const lastRow = worksheet.rowCount;
    for (let row = 1; row <= lastRow; row++) {
        const cellE = worksheet.getCell(`E${row}`);
        const cellEText = getCellText(cellE);
        if (cellEText.trim() === 'Стойка') {
            let type = null;
            let foundHeader = false;
            for (let offset = 1; offset <= 20 && row + offset <= lastRow; offset++) {
                for (const col of ['A', 'B', 'C', 'D', 'E']) {
                    const headerCell = worksheet.getCell(`${col}${row + offset}`);
                    const headerText = getCellText(headerCell);
                    if (headerText.includes('к SV 005 разъемы')) {
                        foundHeader = true;
                        if (headerText.includes('R1 и D1')) type = 'input';
                        else if (headerText.includes('R2 и D2')) type = 'output';
                        break;
                    }
                }
                if (foundHeader) break;
            }
            if (type) {
                let endRow = row;
                for (let r = row + 1; r <= lastRow; r++) {
                    let hasRoom = false;
                    for (const col of ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R']) {
                        const cell = worksheet.getCell(`${col}${r}`);
                        const text = getCellText(cell);
                        if (text.includes('*(помещение№)')) {
                            hasRoom = true;
                            break;
                        }
                    }
                    if (hasRoom) {
                        endRow = r;
                        break;
                    }
                    const nextCellE = worksheet.getCell(`E${r}`);
                    if (getCellText(nextCellE).trim() === 'Стойка') {
                        endRow = r - 1;
                        break;
                    }
                }
                blocks.push({ startRow: row, endRow: endRow || row + 15, type });
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

async function fillPlinthBlockSV005(worksheet, startRow, blockType, plinthData, globalModel) {
    function findRowWithLabel(labelPart) {
        for (let r = startRow; r <= startRow + 15; r++) {
            const cell = worksheet.getCell(`E${r}`);
            const text = getCellText(cell);
            if (text.includes(labelPart)) return r;
        }
        return null;
    }

    const rowRack = startRow;
    const rowController = findRowWithLabel('Контроллер');
    const rowBoard = findRowWithLabel('Плата SV 005');
    const rowHolder = findRowWithLabel('Наименование плинтодержателя');
    const rowPlinthNum = findRowWithLabel('Номер плинта');
    const rowSkud = findRowWithLabel('СКУД');

    if (!rowController || !rowBoard || !rowHolder || !rowPlinthNum || !rowSkud) {
        console.error(`❌ Не найдены все необходимые строки в блоке, startRow=${startRow}`);
        return;
    }

    const rackClean = plinthData.rack.replace(/^ХК\s*/i, '');
    worksheet.getCell(`N${rowRack}`).value = rackClean;

    const controllerCell = worksheet.getCell(`E${rowController}`);
    let controllerText = getCellText(controllerCell);
    if (controllerText.includes('SV 777')) {
        controllerCell.value = controllerText.replace('SV 777', globalModel);
    }
    worksheet.getCell(`N${rowController}`).value = `ХК ${rackClean}.1`;

    worksheet.getCell(`N${rowBoard}`).value = plinthData.boardNumber;

    const absNum = plinthData.plinthNumber;
    const group = Math.floor((absNum - 1) / 15) + 1;
    const holderValue = `ХВ-${rackClean}.${group}`;
    worksheet.getCell(`N${rowHolder}`).value = holderValue;
    if (blockType === 'output') {
        worksheet.getCell(`O${rowHolder}`).value = holderValue;
    }

    worksheet.getCell(`N${rowPlinthNum}`).value = plinthData.plinthNumber;
    worksheet.getCell(`N${rowSkud}`).value = plinthData.skud;
    if (blockType === 'output') {
        worksheet.getCell(`O${rowSkud}`).value = plinthData.skud;
    }

    // Find devices row
    let devicesRow = null;
    for (let r = startRow; r <= startRow + 50; r++) {
        let found = false;
        for (const col of ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R']) {
            const cell = worksheet.getCell(`${col}${r}`);
            let cellValue = '';
            if (cell.formula) {
                cellValue = cell.formula;
            } else if (cell.value) {
                cellValue = getCellText(cell);
            }
            if ((typeof cellValue === 'string' && cellValue.includes('IF(Q')) ||
                (typeof cellValue === 'string' && cellValue.includes('Считыватель'))) {
                devicesRow = r;
                found = true;
                break;
            }
        }
        if (found) break;
    }
    if (!devicesRow) {
        console.warn(`⚠️ Не найдена строка устройств в блоке startRow=${startRow}`);
        return;
    }

    const tm = plinthData.terminalMap || {};
    const cn = plinthData.cableNumbers || {};

    const hasDevices = Object.values(tm).some(v => v && v !== '');
    const roomText = hasDevices ? (plinthData.room || '') : 'Резерв';

    const hasReader = Object.values(tm).includes('reader');
    let deviceLabel = '';
    if (hasReader) {
        if (blockType === 'input') {
            deviceLabel = 'Считыватель вх';
        } else {
            const hasExit = Object.values(tm).includes('exit_btn');
            if (hasExit) {
                deviceLabel = 'Считыватель вх';
            } else {
                deviceLabel = 'Считыватель вых';
            }
        }
        worksheet.getCell(`D${devicesRow}`).value = deviceLabel;
        worksheet.getCell(`F${devicesRow}`).value = cn.reader || '';
    } else {
        worksheet.getCell(`D${devicesRow}`).value = '';
        worksheet.getCell(`F${devicesRow}`).value = '';
    }

    plinthData.deviceLabel = deviceLabel;

    const hasLock = Object.values(tm).includes('lock');
    const hasFireLock = Object.values(tm).includes('fire_lock');
    if (hasLock) {
        worksheet.getCell(`J${devicesRow}`).value = 'замок';
        worksheet.getCell(`K${devicesRow}`).value = cn.lock || '';
    } else if (hasFireLock) {
        worksheet.getCell(`J${devicesRow}`).value = 'Замок пож.дв.';
        worksheet.getCell(`K${devicesRow}`).value = cn.fire_lock || '';
    } else {
        worksheet.getCell(`J${devicesRow}`).value = '';
        worksheet.getCell(`K${devicesRow}`).value = '';
    }

    const hasContact = Object.values(tm).includes('contact');
    worksheet.getCell(`N${devicesRow}`).value = hasContact ? 'геркон' : '';
    worksheet.getCell(`O${devicesRow}`).value = hasContact ? (cn.contact || '') : '';

    const hasExit = Object.values(tm).includes('exit_btn');
    worksheet.getCell(`P${devicesRow}`).value = hasExit ? 'кн.Вых' : '';
    worksheet.getCell(`Q${devicesRow}`).value = hasExit ? (cn.exit_btn || '') : '';

    const hasSiren = Object.values(tm).includes('siren');
    worksheet.getCell(`L${devicesRow}`).value = hasSiren ? 'сирена' : '';
    worksheet.getCell(`M${devicesRow}`).value = hasSiren ? (cn.siren || '') : '';

    let roomRow = null;
    for (let r = devicesRow; r <= devicesRow + 5; r++) {
        const cell = worksheet.getCell(`B${r}`);
        const text = getCellText(cell);
        if (text.includes('*(помещение№)')) {
            roomRow = r;
            break;
        }
    }
    if (roomRow) {
        worksheet.getCell(`B${roomRow}`).value = roomText ? `пом. ${roomText}` : '';
    } else {
        worksheet.getCell(`B${devicesRow + 1}`).value = roomText ? `пом. ${roomText}` : '';
    }
}

// Build test data matching the SV005 etalon (6 boards)
const testBoards = [
    {
        rack: 'ХК 2.1',
        plinth1: { number: 1, room: '14', terminalMap: {0:'reader',1:'lock',2:'contact'}, cableNumbers: {reader:'1',lock:'13',contact:'7'} },
        plinth2: { number: 2, room: '14', terminalMap: {0:'reader'}, cableNumbers: {reader:'4'} },
        skud1: '2.1',
        skud2: '2.1'
    },
    {
        rack: 'ХК 2.1',
        plinth1: { number: 3, room: '17', terminalMap: {0:'reader',1:'lock',2:'contact'}, cableNumbers: {reader:'34',lock:'46',contact:'40'} },
        plinth2: { number: 4, room: '17', terminalMap: {0:'reader'}, cableNumbers: {reader:'37'} },
        skud1: '2.2',
        skud2: '2.2'
    },
    {
        rack: 'ХК 2.1',
        plinth1: { number: 5, room: '17', terminalMap: {0:'reader',1:'lock',2:'contact'}, cableNumbers: {reader:'49',lock:'61',contact:'55'} },
        plinth2: { number: 6, room: '17', terminalMap: {0:'reader'}, cableNumbers: {reader:'52'} },
        skud1: '2.6',
        skud2: '2.6'
    },
    {
        rack: 'ХК 2.1',
        plinth1: { number: 7, room: '20', terminalMap: {0:'reader',1:'lock',2:'contact'}, cableNumbers: {reader:'121',lock:'133',contact:'127'} },
        plinth2: { number: 8, room: '20', terminalMap: {0:'reader'}, cableNumbers: {reader:'124'} },
        skud1: '2.4',
        skud2: '2.4'
    },
    {
        rack: 'ХК 2.1',
        plinth1: { number: 9, room: '20', terminalMap: {0:'reader',1:'lock',2:'contact',3:'exit_btn'}, cableNumbers: {reader:'100',lock:'112',contact:'106',exit_btn:'103'} },
        plinth2: { number: 10, room: '23', terminalMap: {0:'reader',1:'lock',2:'contact',3:'exit_btn'}, cableNumbers: {reader:'154',lock:'166',contact:'160',exit_btn:'157'} },
        skud1: '2.3',
        skud2: '2.5'
    },
    {
        rack: 'ХК 2.1',
        plinth1: { number: 11, room: '', terminalMap: {}, cableNumbers: {} },
        plinth2: { number: 12, room: '', terminalMap: {}, cableNumbers: {} },
        skud1: '',
        skud2: ''
    }
];

async function main() {
    // Generate SV005
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(TEMPLATE_SV005);
    const worksheet = workbook.getWorksheet('Лист1');

    let blocks = await getBlocksSV005(worksheet);
    blocks.sort((a, b) => a.startRow - b.startRow);

    const neededPairs = testBoards.length;
    let inputBlocks = blocks.filter(b => b.type === 'input');
    let outputBlocks = blocks.filter(b => b.type === 'output');

    while (inputBlocks.length < neededPairs) {
        const lastInput = inputBlocks[inputBlocks.length - 1];
        const lastBlockEnd = Math.max(...blocks.map(b => b.endRow));
        const targetStart = lastBlockEnd + 1;
        copyBlock(worksheet, lastInput.startRow, lastInput.endRow, targetStart);
        blocks = await getBlocksSV005(worksheet);
        blocks.sort((a, b) => a.startRow - b.startRow);
        inputBlocks = blocks.filter(b => b.type === 'input');
        outputBlocks = blocks.filter(b => b.type === 'output');
    }

    while (outputBlocks.length < neededPairs) {
        const lastOutput = outputBlocks[outputBlocks.length - 1];
        const lastBlockEnd = Math.max(...blocks.map(b => b.endRow));
        const targetStart = lastBlockEnd + 1;
        copyBlock(worksheet, lastOutput.startRow, lastOutput.endRow, targetStart);
        blocks = await getBlocksSV005(worksheet);
        blocks.sort((a, b) => a.startRow - b.startRow);
        inputBlocks = blocks.filter(b => b.type === 'input');
        outputBlocks = blocks.filter(b => b.type === 'output');
    }

    const usedInputBlocks = inputBlocks.slice(0, neededPairs);
    const usedOutputBlocks = outputBlocks.slice(0, neededPairs);

    for (let i = 0; i < testBoards.length; i++) {
        const board = testBoards[i];
        const boardNumber = i + 1;

        await fillPlinthBlockSV005(worksheet, usedInputBlocks[i].startRow, 'input', {
            rack: board.rack,
            boardNumber: boardNumber,
            plinthNumber: board.plinth1.number,
            skud: board.skud1,
            room: board.plinth1.room,
            terminalMap: board.plinth1.terminalMap,
            cableNumbers: board.plinth1.cableNumbers
        }, 'SV 777');

        await fillPlinthBlockSV005(worksheet, usedOutputBlocks[i].startRow, 'output', {
            rack: board.rack,
            boardNumber: boardNumber,
            plinthNumber: board.plinth2.number,
            skud: board.skud2,
            room: board.plinth2.room,
            terminalMap: board.plinth2.terminalMap,
            cableNumbers: board.plinth2.cableNumbers
        }, 'SV 777');
    }

    // Delete unused blocks
    const allBlocks = await getBlocksSV005(worksheet);
    const usedStartRows = new Set();
    usedInputBlocks.forEach(b => usedStartRows.add(b.startRow));
    usedOutputBlocks.forEach(b => usedStartRows.add(b.startRow));
    const blocksToDelete = allBlocks.filter(b => !usedStartRows.has(b.startRow));
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
    const outPath = path.join(base, 'test_output_sv005.xlsx');
    await workbook.xlsx.writeFile(outPath);
    console.log('Generated SV005 saved to ' + outPath);
    
    // Now compare with etalon
    const etalonWb = new ExcelJS.Workbook();
    await etalonWb.xlsx.readFile(path.join(base, '(г.Курган_ ул.Карла Маркса_ 149)_2026-07-03_SV005_fixed (1).xlsx'));
    const etalonSheet = etalonWb.getWorksheet('Лист1');
    
    const genSheet = workbook.getWorksheet('Лист1');
    
    console.log('\n=== COMPARISON SV005 ===');
    console.log('Generated rows: ' + genSheet.rowCount + ', Etalon rows: ' + etalonSheet.rowCount);
    
    let differences = 0;
    const maxRows = Math.max(genSheet.rowCount, etalonSheet.rowCount);
    for (let r = 1; r <= maxRows; r++) {
        const genVals = [];
        const etaVals = [];
        for (const c of [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22]) {
            const genCell = genSheet.getCell(r, c);
            const etaCell = etalonSheet.getCell(r, c);
            const gv = getCellText(genCell);
            const ev = getCellText(etaCell);
            if (gv !== '' && gv !== ev) {
                genVals.push('C' + c + '=' + gv);
                etaVals.push('C' + c + '=' + ev);
            } else if (ev !== '' && ev !== gv) {
                genVals.push('C' + c + '=' + gv);
                etaVals.push('C' + c + '=' + ev);
            }
        }
        if (genVals.length || etaVals.length) {
            differences++;
            console.log(`R${r}: GEN: ${genVals.join(' | ') || '(empty)'}`);
            console.log(`     ETA: ${etaVals.join(' | ') || '(empty)'}`);
        }
    }
    console.log(`\nTotal differences: ${differences}`);
}

main().catch(console.error);
