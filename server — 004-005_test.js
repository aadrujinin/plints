const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 3000;

app.use(express.json());
app.use(express.static('public'));

const TEMPLATE_SV005 = path.join(__dirname, 'templateSV005_test.xlsx');
const TEMPLATE_SV004 = path.join(__dirname, '03. Расшивки на плинты SV004 образец v3 .xlsx');

// ---------- Общие утилиты ----------
function getCellText(cell) {
    if (!cell || cell.value === undefined || cell.value === null) return '';
    if (typeof cell.value === 'string') return cell.value;
    if (typeof cell.value === 'object') {
        if (cell.value.richText && Array.isArray(cell.value.richText)) {
            return cell.value.richText.map(rt => rt.text).join('');
        }
        if (cell.value.text) return cell.value.text;
        return JSON.stringify(cell.value);
    }
    return String(cell.value);
}

function buildHolder(rack, plinthNumber) {
    const rackClean = rack.replace(/^ХК\s*/i, '');
    if (!rackClean) return '';
    const absNum = parseInt(plinthNumber, 10);
    const group = Math.floor((absNum - 1) / 15) + 1;
    return `ХВ-${rackClean}.${group}`;
}

function copyBlock(worksheet, sourceStart, sourceEnd, targetStart) {
    const rowCount = sourceEnd - sourceStart + 1;
    worksheet.spliceRows(targetStart, 0, ...Array(rowCount).fill([]));
    for (let i = 0; i < rowCount; i++) {
        const srcRow = worksheet.getRow(sourceStart + i);
        const dstRow = worksheet.getRow(targetStart + i);
        srcRow.eachCell((cell, colNumber) => {
            const dstCell = dstRow.getCell(colNumber);
            dstCell.value = cell.value;
            if (cell.style) dstCell.style = { ...cell.style };
        });
        dstRow.height = srcRow.height;
    }
    return targetStart + rowCount - 1;
}

// ---------- Блоки для SV005 ----------
function getBlocksSV005(worksheet) {
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

// ---------- Блоки для SV004 (только входные) ----------
function getBlocksSV004(worksheet) {
    if (!worksheet) return [];
    const blocks = [];
    const lastRow = worksheet.rowCount;
    for (let row = 1; row <= lastRow; row++) {
        const cellE = worksheet.getCell(`E${row}`);
        const text = getCellText(cellE);
        if (text.trim() === 'Стойка') {
            const cellN = worksheet.getCell(`N${row}`);
            const plinthNum = parseInt(getCellText(cellN), 10);
            let type = null;
            if (!isNaN(plinthNum)) {
                type = (plinthNum % 2 === 1) ? 'input' : 'output';
            } else {
                for (let offset = 1; offset <= 10; offset++) {
                    const checkRow = row + offset;
                    if (checkRow > lastRow) break;
                    const cellE2 = worksheet.getCell(`E${checkRow}`);
                    if (getCellText(cellE2).trim() === 'Номер плинта') {
                        const cellN2 = worksheet.getCell(`N${checkRow}`);
                        const num2 = parseInt(getCellText(cellN2), 10);
                        if (!isNaN(num2)) {
                            type = (num2 % 2 === 1) ? 'input' : 'output';
                            break;
                        }
                    }
                }
            }
            if (type === 'input') {
                let endRow = row;
                for (let r = row + 1; r <= lastRow; r++) {
                    const cellEnext = worksheet.getCell(`E${r}`);
                    if (getCellText(cellEnext).trim() === 'Стойка') {
                        endRow = r - 1;
                        break;
                    }
                    if (r - row > 20) break;
                }
                blocks.push({ startRow: row, endRow: endRow || row + 15, type });
            }
        }
    }
    return blocks;
}

// ---------- Заполнение блока SV005 ----------
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

    worksheet.getCell(`N${rowPlinthNum}`).value = plinthData.plinthNumber;
    worksheet.getCell(`N${rowSkud}`).value = plinthData.skud;

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

    const hasReader = Object.values(tm).includes('reader');
    if (hasReader) {
        let deviceLabel = '';
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

    const hasLock = Object.values(tm).includes('lock');
    const hasFireLock = Object.values(tm).includes('fire_lock');
    if (hasLock) {
        worksheet.getCell(`J${devicesRow}`).value = 'Замок';
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
        worksheet.getCell(`B${roomRow}`).value = `пом. ${plinthData.room}`;
    } else {
        worksheet.getCell(`B${devicesRow + 1}`).value = `пом. ${plinthData.room}`;
    }
}

// ---------- Заполнение блока SV004 (один плинт) ----------
async function fillPlinthBlockSV004(worksheet, startRow, plinthData, globalModel) {
    const groups = [
        { offsetDev: 14, offsetRoom: 15, pins: [0,1,2,3] },
        { offsetDev: 34, offsetRoom: 35, pins: [4,5,6,7] },
        { offsetDev: 56, offsetRoom: 57, pins: [8,9] }
    ];

    const rackClean = plinthData.rack.replace(/^ХК\s*/i, '');
    worksheet.getCell(`N${startRow}`).value = rackClean;

    const controllerRow = startRow + 1;
    const controllerCell = worksheet.getCell(`E${controllerRow}`);
    let controllerText = getCellText(controllerCell);
    if (controllerText.includes('SV 777')) {
        controllerCell.value = controllerText.replace('SV 777', globalModel);
    }
    worksheet.getCell(`N${controllerRow}`).value = `ХК ${rackClean}.1`;

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

    for (const group of groups) {
        const devRow = startRow + group.offsetDev;
        const roomRow = startRow + group.offsetRoom;
        if (devRow > worksheet.rowCount || roomRow > worksheet.rowCount) continue;
        for (let i = 0; i < group.pins.length; i++) {
            const pin = group.pins[i];
            const colDevice = ['B','F','J','N','B','F','J','N','B','F'][pin];
            const colCable = ['D','H','L','P','D','H','L','P','D','H'][pin];
            const device = tm[pin] || '';
            const cable = cm[pin] || '';
            worksheet.getCell(`${colDevice}${devRow}`).value = device;
            worksheet.getCell(`${colCable}${devRow}`).value = cable;
            worksheet.getCell(`${colDevice}${roomRow}`).value = `пом. ${plinthData.room}`;
        }
    }
}

// ---------- Создание листов "шпоргалка" и "Disp" для SV004 ----------
function createSheetsSV004(workbook, boards, globalModel) {
    const sheetNames = workbook.worksheets.map(s => s.name);
    if (sheetNames.includes('шпоргалка')) workbook.removeWorksheet('шпоргалка');
    if (sheetNames.includes('Disp')) workbook.removeWorksheet('Disp');

    const cheatSheet = workbook.addWorksheet('шпоргалка', { properties: { tabColor: { argb: 'FFE0E0E0' } } });
    cheatSheet.getCell('C2').value = `${globalModel}`;
    cheatSheet.getCell('C2').font = { bold: true, size: 12 };

    let rowIdx = 4;
    let counter = 1;
    for (let i = 0; i < boards.length; i++) {
        const board = boards[i];
        const boardNumber = i + 1;
        const room = board.plinth1.room || '';
        const tm = board.plinth1.terminalMap || {};
        const cm = board.plinth1.cableMap || {};

        for (let pin = 0; pin <= 9; pin++) {
            const device = tm[pin] || '';
            const cable = cm[pin] || '';
            if (device) {
                cheatSheet.getCell(`A${rowIdx}`).value = '';
                cheatSheet.getCell(`B${rowIdx}`).value = boardNumber;
                cheatSheet.getCell(`C${rowIdx}`).value = `Пин${pin}`;
                cheatSheet.getCell(`D${rowIdx}`).value = device;
                cheatSheet.getCell(`E${rowIdx}`).value = `пом. ${room}`;
                cheatSheet.getCell(`F${rowIdx}`).value = cable ? `ШЛ.${cable}` : '';
                cheatSheet.getCell(`G${rowIdx}`).value = counter++;
                const color = (device === 'ИК') ? 'FFFF0000' : (device === 'КАВ') ? 'FF00FF00' : (device === 'СМК-ДРС') ? 'FF0000FF' : 'FFFFFF00';
                cheatSheet.getCell(`G${rowIdx}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
                rowIdx++;
            }
        }
        rowIdx++;
    }

    const dispSheet = workbook.addWorksheet('Disp', { properties: { tabColor: { argb: 'FFE0E0FF' } } });
    dispSheet.getCell('K2').value = 'Display panel';
    const cheatRows = cheatSheet.rowCount;
    for (let r = 4; r <= cheatRows; r++) {
        const srcRow = cheatSheet.getRow(r);
        if (srcRow.getCell(1).value === '' && srcRow.getCell(2).value === '') continue;
        const device = srcRow.getCell(4).value || '';
        const room = srcRow.getCell(5).value || '';
        const roomNum = room.replace('пом. ', '').trim();
        const text = `п.${roomNum} - ${device}`;
        const dispRow = dispSheet.getRow(r - 2);
        dispRow.getCell(2).value = srcRow.getCell(2).value;
        dispRow.getCell(3).value = text;
        dispRow.getCell(4).value = srcRow.getCell(7).value;
    }
}

// ---------- Маршрут для SV005 ----------
app.post('/generate-sv005', async (req, res) => {
    try {
        const { address, globalModel, boards } = req.body;
        if (!address || !boards || !boards.length) {
            return res.status(400).json({ error: 'Не указан адрес или список плат' });
        }
        if (!fs.existsSync(TEMPLATE_SV005)) {
            return res.status(500).json({ error: 'Файл шаблона SV005 не найден.' });
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(TEMPLATE_SV005);

        let worksheet = workbook.getWorksheet('Лист1');
        if (!worksheet) {
            worksheet = workbook.worksheets[0];
            if (!worksheet) {
                return res.status(500).json({ error: 'В файле шаблона нет ни одного листа.' });
            }
        }

        let blocks = getBlocksSV005(worksheet);
        if (blocks.length === 0) {
            return res.status(500).json({ error: 'В шаблоне SV005 не найдено блоков "Стойка".' });
        }
        blocks.sort((a, b) => a.startRow - b.startRow);

        // Для SV005 каждая плата даёт 2 блока (вход и выход)
        const neededInput = boards.length;
        const neededOutput = boards.length;

        let inputBlocks = blocks.filter(b => b.type === 'input');
        let outputBlocks = blocks.filter(b => b.type === 'output');

        while (inputBlocks.length < neededInput) {
            const lastInput = inputBlocks[inputBlocks.length - 1];
            if (!lastInput) {
                return res.status(500).json({ error: 'Нет ни одного входного блока для копирования.' });
            }
            const lastBlockEnd = Math.max(...blocks.map(b => b.endRow));
            const targetStart = lastBlockEnd + 1;
            copyBlock(worksheet, lastInput.startRow, lastInput.endRow, targetStart);
            blocks = getBlocksSV005(worksheet);
            blocks.sort((a, b) => a.startRow - b.startRow);
            inputBlocks = blocks.filter(b => b.type === 'input');
            outputBlocks = blocks.filter(b => b.type === 'output');
        }

        while (outputBlocks.length < neededOutput) {
            const lastOutput = outputBlocks[outputBlocks.length - 1];
            if (!lastOutput) {
                return res.status(500).json({ error: 'Нет ни одного выходного блока для копирования.' });
            }
            const lastBlockEnd = Math.max(...blocks.map(b => b.endRow));
            const targetStart = lastBlockEnd + 1;
            copyBlock(worksheet, lastOutput.startRow, lastOutput.endRow, targetStart);
            blocks = getBlocksSV005(worksheet);
            blocks.sort((a, b) => a.startRow - b.startRow);
            inputBlocks = blocks.filter(b => b.type === 'input');
            outputBlocks = blocks.filter(b => b.type === 'output');
        }

        const usedInputBlocks = inputBlocks.slice(0, neededInput);
        const usedOutputBlocks = outputBlocks.slice(0, neededOutput);

        for (let i = 0; i < boards.length; i++) {
            const board = boards[i];
            const boardNumber = i + 1;

            await fillPlinthBlockSV005(
                worksheet,
                usedInputBlocks[i].startRow,
                'input',
                {
                    rack: board.rack,
                    boardNumber: boardNumber,
                    plinthNumber: board.plinth1.number,
                    skud: board.skud1,
                    room: board.plinth1.room,
                    terminalMap: board.plinth1.terminalMap,
                    cableNumbers: board.plinth1.cableNumbers
                },
                globalModel
            );

            await fillPlinthBlockSV005(
                worksheet,
                usedOutputBlocks[i].startRow,
                'output',
                {
                    rack: board.rack,
                    boardNumber: boardNumber,
                    plinthNumber: board.plinth2.number,
                    skud: board.skud2,
                    room: board.plinth2.room,
                    terminalMap: board.plinth2.terminalMap,
                    cableNumbers: board.plinth2.cableNumbers
                },
                globalModel
            );
        }

        const allBlocks = getBlocksSV005(worksheet);
        const usedStartRows = new Set();
        usedInputBlocks.forEach(b => usedStartRows.add(b.startRow));
        usedOutputBlocks.forEach(b => usedStartRows.add(b.startRow));
        const blocksToDelete = allBlocks.filter(b => !usedStartRows.has(b.startRow));
        blocksToDelete.sort((a, b) => b.startRow - a.startRow);
        for (const block of blocksToDelete) {
            const rowCount = block.endRow - block.startRow + 1;
            worksheet.spliceRows(block.startRow, rowCount);
        }

        const remainingBlocks = getBlocksSV005(worksheet);
        if (remainingBlocks.length > 0) {
            const lastEndRow = Math.max(...remainingBlocks.map(b => b.endRow));
            const totalRows = worksheet.rowCount;
            if (totalRows > lastEndRow) {
                worksheet.spliceRows(lastEndRow + 1, totalRows - lastEndRow);
            }
        }

        const buffer = await workbook.xlsx.writeBuffer();
        const dateStr = new Date().toISOString().slice(0, 10);
        const safeAddress = address.replace(/[\\/:*?"<>|]/g, '_');
        const filename = `${safeAddress}_SV005_${dateStr}.xlsx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
        res.send(buffer);
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Ошибка генерации SV005: ' + err.message });
    }
});

// ---------- Маршрут для SV004 ----------
app.post('/generate-sv004', async (req, res) => {
    try {
        const { address, globalModel, boards } = req.body;
        if (!address || !boards || !boards.length) {
            return res.status(400).json({ error: 'Не указан адрес или список плат' });
        }
        if (!fs.existsSync(TEMPLATE_SV004)) {
            return res.status(500).json({ error: 'Файл шаблона SV004 не найден.' });
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(TEMPLATE_SV004);

        let worksheet = workbook.getWorksheet('SV777-1 (SV004)');
        if (!worksheet) {
            worksheet = workbook.worksheets[0];
            if (!worksheet) {
                return res.status(500).json({ error: 'В файле шаблона нет ни одного листа.' });
            }
        }

        let blocks = getBlocksSV004(worksheet);
        if (blocks.length === 0) {
            return res.status(500).json({ error: 'В шаблоне SV004 не найдено входных блоков "Стойка".' });
        }
        blocks.sort((a, b) => a.startRow - b.startRow);

        const needed = boards.length;
        let inputBlocks = blocks.filter(b => b.type === 'input');

        while (inputBlocks.length < needed) {
            const lastInput = inputBlocks[inputBlocks.length - 1];
            if (!lastInput) {
                return res.status(500).json({ error: 'Нет ни одного входного блока для копирования.' });
            }
            const lastBlockEnd = Math.max(...blocks.map(b => b.endRow));
            const targetStart = lastBlockEnd + 1;
            copyBlock(worksheet, lastInput.startRow, lastInput.endRow, targetStart);
            blocks = getBlocksSV004(worksheet);
            blocks.sort((a, b) => a.startRow - b.startRow);
            inputBlocks = blocks.filter(b => b.type === 'input');
        }

        const usedInputBlocks = inputBlocks.slice(0, needed);

        for (let i = 0; i < boards.length; i++) {
            const board = boards[i];
            const boardNumber = i + 1;

            await fillPlinthBlockSV004(
                worksheet,
                usedInputBlocks[i].startRow,
                {
                    rack: board.rack,
                    boardNumber: boardNumber,
                    plinthNumber: board.plinth1.number,
                    skud: board.skud1,
                    room: board.plinth1.room,
                    terminalMap: board.plinth1.terminalMap,
                    cableMap: board.plinth1.cableMap
                },
                globalModel
            );
        }

        const allBlocks = getBlocksSV004(worksheet);
        const usedStartRows = new Set();
        usedInputBlocks.forEach(b => usedStartRows.add(b.startRow));
        const blocksToDelete = allBlocks.filter(b => !usedStartRows.has(b.startRow));
        blocksToDelete.sort((a, b) => b.startRow - a.startRow);
        for (const block of blocksToDelete) {
            const rowCount = block.endRow - block.startRow + 1;
            worksheet.spliceRows(block.startRow, rowCount);
        }

        const remainingBlocks = getBlocksSV004(worksheet);
        if (remainingBlocks.length > 0) {
            const lastEndRow = Math.max(...remainingBlocks.map(b => b.endRow));
            const totalRows = worksheet.rowCount;
            if (totalRows > lastEndRow) {
                worksheet.spliceRows(lastEndRow + 1, totalRows - lastEndRow);
            }
        }

        createSheetsSV004(workbook, boards, globalModel);

        const buffer = await workbook.xlsx.writeBuffer();
        const dateStr = new Date().toISOString().slice(0, 10);
        const safeAddress = address.replace(/[\\/:*?"<>|]/g, '_');
        const filename = `${safeAddress}_SV004_${dateStr}.xlsx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
        res.send(buffer);
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Ошибка генерации SV004: ' + err.message });
    }
});

// Запуск сервера
app.listen(PORT, () => {
    console.log(`Сервер запущен: http://localhost:${PORT}`);
});