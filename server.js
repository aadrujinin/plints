require('dotenv').config();
const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const { Pool } = require('pg');

const app = express();
const PORT = 3000;

app.use(express.json());
app.use(express.static('public'));

const TEMPLATE_SV005 = path.join(__dirname, 'templateSV005_test.xlsx');
const TEMPLATE_SV004 = path.join(__dirname, 'templateSV004_test.xlsx');

// ---------- Проверка переменных окружения ----------
if (!process.env.API_KEY || !process.env.API_BASE_URL) {
    console.warn('⚠️ Внимание: API_KEY или API_BASE_URL не заданы в .env. Синхронизация и поиск проектов не будут работать.');
}
if (!process.env.DB_HOST || !process.env.DB_NAME) {
    console.warn('⚠️ Внимание: параметры PostgreSQL не заданы. База данных не будет работать.');
}

// ---------- Подключение к PostgreSQL ----------
const pool = new Pool({
    host: process.env.DB_HOST,
    port: process.env.DB_PORT,
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    database: process.env.DB_NAME,
});

// ---------- Инициализация таблицы проектов ----------
async function initDatabase() {
    const createQuery = `
        CREATE TABLE IF NOT EXISTS projects (
            id BIGINT PRIMARY KEY,
            name TEXT NOT NULL,
            is_archive BOOLEAN DEFAULT FALSE,
            expanded_downloaded BOOLEAN DEFAULT FALSE,
            file_name TEXT,
            file_path TEXT,
            cf_92 TEXT,
            cf_217 TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    `;
    await pool.query(createQuery);
    console.log('✅ Таблица projects проверена/создана');

    try {
        await pool.query('ALTER TABLE projects ALTER COLUMN id TYPE BIGINT;');
        console.log('✅ Тип колонки id изменён на BIGINT (если был INTEGER)');
    } catch (err) {
        if (err.code !== '42704' && !err.message.includes('already exists')) {
            console.warn('⚠️ Не удалось изменить тип id (возможно, уже BIGINT):', err.message);
        }
    }

    try {
        await pool.query('ALTER TABLE projects ADD COLUMN IF NOT EXISTS cf_92 TEXT;');
        console.log('✅ Поле cf_92 добавлено');
    } catch (err) {
        console.warn('⚠️ Не удалось добавить cf_92:', err.message);
    }

    try {
        await pool.query('ALTER TABLE projects ADD COLUMN IF NOT EXISTS cf_217 TEXT;');
        console.log('✅ Поле cf_217 добавлено');
    } catch (err) {
        console.warn('⚠️ Не удалось добавить cf_217:', err.message);
    }
}
initDatabase();

// ---------- ОБЩИЕ УТИЛИТЫ ----------
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

function getNumericCellValue(cell) {
    if (!cell) return NaN;
    let val = cell.value;
    if (val === undefined || val === null) return NaN;
    if (typeof val === 'number') return val;
    if (typeof val === 'string') {
        const num = parseFloat(val);
        return isNaN(num) ? NaN : num;
    }
    if (typeof val === 'object') {
        if (val.result !== undefined) {
            const num = parseFloat(val.result);
            return isNaN(num) ? NaN : num;
        }
        if (val.text) {
            const num = parseFloat(val.text);
            return isNaN(num) ? NaN : num;
        }
        if (val.richText && Array.isArray(val.richText)) {
            const text = val.richText.map(rt => rt.text).join('');
            const num = parseFloat(text);
            return isNaN(num) ? NaN : num;
        }
    }
    return NaN;
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

// ---------- АВТОПОДГОН ШИРИНЫ СТОЛБЦОВ ----------
function isPlinthSheet(sheetName) {
    const plinthSheetNames = ['Лист1', 'SV777-1 (SV004)'];
    if (plinthSheetNames.includes(sheetName)) return true;
    if (sheetName.startsWith('Disp-')) return true;
    return false;
}

function applyAutoFit(workbook) {
    workbook.worksheets.forEach(worksheet => {
        if (isPlinthSheet(worksheet.name)) return;

        const colMaxLength = {};
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                let text = '';
                try {
                    text = cell.text || '';
                } catch(e) {
                    text = '';
                }
                if (text && text.length > 0) {
                    const len = text.length;
                    if (!colMaxLength[colNumber] || len > colMaxLength[colNumber]) {
                        colMaxLength[colNumber] = len;
                    }
                }
            });
        });
        for (const colNumber in colMaxLength) {
            const maxLen = colMaxLength[colNumber];
            let width = Math.max(10, maxLen * 1.2 + 2);
            worksheet.getColumn(parseInt(colNumber, 10)).width = width;
        }
    });
}

// ---------- БЛОКИ ДЛЯ SV005 ----------
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

// ---------- БЛОКИ ДЛЯ SV004 ----------
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
                if (!found) {
                    endRow = row + 19;
                }
                blocks.push({ startRow: row, endRow: endRow, type });
            }
        }
    }
    return blocks;
}

// ---------- ЗАПОЛНЕНИЕ БЛОКА SV005 ----------
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

// ---------- ЗАПОЛНЕНИЕ БЛОКА SV004 ----------
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

            if (!device) {
                room = 'Резерв';
            }

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

// ---------- ЗАПОЛНЕНИЕ ЛИСТОВ ШПОРГАЛКА И DISP ДЛЯ SV004 ----------
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
            console.warn(`⚠️ Листы ${cheatSheetName} или ${dispSheetName} не найдены, пропускаем`);
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

// ---------- СОЗДАНИЕ ЛИСТА "Шпоргалка общая" ДЛЯ SV004 ----------
function createCommonCheatSheetSV004(workbook) {
    const commonSheetName = 'Шпоргалка общая';
    let commonSheet = workbook.getWorksheet(commonSheetName);
    if (!commonSheet) {
        commonSheet = workbook.addWorksheet(commonSheetName);
    } else {
        // Очищаем все строки, кроме заголовка (строки 1-2)
        if (commonSheet.rowCount >= 2) {
            commonSheet.spliceRows(2, commonSheet.rowCount - 1);
        }
    }

    // Копируем заголовок (строка 2) из первой шпоргалки
    const firstCheat = workbook.getWorksheet('шпоргалка-1');
    if (firstCheat) {
        const srcRow = firstCheat.getRow(2);
        const dstRow = commonSheet.getRow(2);
        srcRow.eachCell((cell, colNumber) => {
            dstRow.getCell(colNumber).value = cell.value;
            if (cell.style) dstRow.getCell(colNumber).style = { ...cell.style };
        });
        dstRow.height = srcRow.height;
    }

    let targetRowIndex = 4; // Данные начинаются с 4-й строки

    // Последовательно копируем данные из шпоргалка-1, -2, -3
    for (let i = 1; i <= 3; i++) {
        const cheatName = `шпоргалка-${i}`;
        const cheatSheet = workbook.getWorksheet(cheatName);
        if (!cheatSheet) continue;

        const lastRow = cheatSheet.rowCount;
        for (let r = 4; r <= lastRow; r++) {
            const srcRow = cheatSheet.getRow(r);
            let hasData = false;
            srcRow.eachCell((cell) => {
                if (cell.value !== undefined && cell.value !== null && cell.value !== '') {
                    hasData = true;
                }
            });
            if (!hasData) continue;

            const dstRow = commonSheet.getRow(targetRowIndex);
            srcRow.eachCell((cell, colNumber) => {
                dstRow.getCell(colNumber).value = cell.value;
                if (cell.style) dstRow.getCell(colNumber).style = { ...cell.style };
            });
            dstRow.height = srcRow.height;
            targetRowIndex++;
        }
    }
}

// ---------- СОЗДАНИЕ ЛИСТОВ "ШПОРГАЛКА" И "ШПОРА ОБЩАЯ" ДЛЯ SV005 ----------
function createSheetsSV005(workbook, boards, globalModel) {
    const sheetNames = workbook.worksheets.map(s => s.name);
    if (sheetNames.includes('шпоргалка')) workbook.removeWorksheet('шпоргалка');
    if (sheetNames.includes('Шпора общая')) workbook.removeWorksheet('Шпора общая');

    const cheatSheet = workbook.addWorksheet('шпоргалка', { properties: { tabColor: { argb: 'FFE0E0E0' } } });
    cheatSheet.getCell('C2').value = `${globalModel} - 1`;
    cheatSheet.getCell('C2').font = { bold: true, size: 12 };

    let rowIdx = 4;
    let counter = 1;

    for (let i = 0; i < boards.length; i++) {
        const board = boards[i];
        const boardNumber = i + 1;
        const commonRoom = board.plinth1.room || '';

        const num1 = i * 2 + 1;
        const num2 = i * 2 + 2;

        const tm1 = board.plinth1.terminalMap || {};
        const cn1 = board.plinth1.cableNumbers || {};
        const hasReader1 = Object.values(tm1).includes('reader');
        const hasLock1 = Object.values(tm1).includes('lock');
        const hasFireLock1 = Object.values(tm1).includes('fire_lock');

        const rowR1 = rowIdx;
        cheatSheet.getCell(`A${rowR1}`).value = '';
        cheatSheet.getCell(`B${rowR1}`).value = boardNumber;
        cheatSheet.getCell(`C${rowR1}`).value = `R${num1}`;
        if (hasReader1) {
            cheatSheet.getCell(`D${rowR1}`).value = 'Считыватель вх';
            cheatSheet.getCell(`F${rowR1}`).value = `ШЛ.${cn1.reader || ''}`;
            cheatSheet.getCell(`G${rowR1}`).value = counter++;
        } else {
            cheatSheet.getCell(`D${rowR1}`).value = '';
            cheatSheet.getCell(`F${rowR1}`).value = '';
            cheatSheet.getCell(`G${rowR1}`).value = '';
        }
        if (commonRoom) {
            cheatSheet.getCell(`E${rowR1}`).value = `пом. ${commonRoom}`;
        } else {
            const hasAnyDevice1 = Object.values(tm1).some(v => v && v !== '');
            cheatSheet.getCell(`E${rowR1}`).value = hasAnyDevice1 ? '' : 'Резерв';
        }
        rowIdx++;

        const rowD1 = rowIdx;
        cheatSheet.getCell(`A${rowD1}`).value = '';
        cheatSheet.getCell(`B${rowD1}`).value = '';
        cheatSheet.getCell(`C${rowD1}`).value = `D${num1}`;
        if (hasLock1) {
            cheatSheet.getCell(`D${rowD1}`).value = 'Замок';
        } else if (hasFireLock1) {
            cheatSheet.getCell(`D${rowD1}`).value = 'Замок пож.дв.';
        } else {
            cheatSheet.getCell(`D${rowD1}`).value = '';
        }
        if (commonRoom) {
            cheatSheet.getCell(`E${rowD1}`).value = `пом. ${commonRoom}`;
        } else {
            const hasAnyDevice1 = Object.values(tm1).some(v => v && v !== '');
            cheatSheet.getCell(`E${rowD1}`).value = hasAnyDevice1 ? '' : 'Резерв';
        }
        cheatSheet.getCell(`F${rowD1}`).value = board.skud1 ? `СКД.${board.skud1}` : '';
        rowIdx++;

        if (hasReader1 || hasLock1 || hasFireLock1) {
            cheatSheet.mergeCells(`G${rowR1}:G${rowD1}`);
            cheatSheet.getCell(`G${rowR1}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0000' } };
        }

        const tm2 = board.plinth2.terminalMap || {};
        const cn2 = board.plinth2.cableNumbers || {};
        const hasReader2 = Object.values(tm2).includes('reader');
        const hasLock2 = Object.values(tm2).includes('lock');
        const hasFireLock2 = Object.values(tm2).includes('fire_lock');

        const rowR2 = rowIdx;
        cheatSheet.getCell(`A${rowR2}`).value = '';
        cheatSheet.getCell(`B${rowR2}`).value = '';
        cheatSheet.getCell(`C${rowR2}`).value = `R${num2}`;
        if (hasReader2) {
            cheatSheet.getCell(`D${rowR2}`).value = board.plinth2.deviceLabel || 'Считыватель вых';
            cheatSheet.getCell(`F${rowR2}`).value = `ШЛ.${cn2.reader || ''}`;
            cheatSheet.getCell(`G${rowR2}`).value = counter++;
        } else {
            cheatSheet.getCell(`D${rowR2}`).value = '';
            cheatSheet.getCell(`F${rowR2}`).value = '';
            cheatSheet.getCell(`G${rowR2}`).value = '';
        }
        const room2 = board.plinth2.room || '';
        if (room2) {
            cheatSheet.getCell(`E${rowR2}`).value = `пом. ${room2}`;
        } else {
            const hasAnyDevice2 = Object.values(tm2).some(v => v && v !== '');
            cheatSheet.getCell(`E${rowR2}`).value = hasAnyDevice2 ? '' : 'Резерв';
        }
        rowIdx++;

        const rowD2 = rowIdx;
        cheatSheet.getCell(`A${rowD2}`).value = '';
        cheatSheet.getCell(`B${rowD2}`).value = '';
        cheatSheet.getCell(`C${rowD2}`).value = `D${num2}`;
        if (hasLock2) {
            cheatSheet.getCell(`D${rowD2}`).value = 'Замок';
        } else if (hasFireLock2) {
            cheatSheet.getCell(`D${rowD2}`).value = 'Замок пож.дв.';
        } else {
            cheatSheet.getCell(`D${rowD2}`).value = '';
        }
        if (room2) {
            cheatSheet.getCell(`E${rowD2}`).value = `пом. ${room2}`;
        } else {
            const hasAnyDevice2 = Object.values(tm2).some(v => v && v !== '');
            cheatSheet.getCell(`E${rowD2}`).value = hasAnyDevice2 ? '' : 'Резерв';
        }
        cheatSheet.getCell(`F${rowD2}`).value = board.skud2 ? `СКД.${board.skud2}` : '';
        rowIdx++;

        if (hasReader2) {
            const fillColorR2 = (board.plinth2.deviceLabel === 'Считыватель вх') ? 'FFA500' : 'FF0000';
            cheatSheet.getCell(`G${rowR2}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColorR2 } };
            cheatSheet.getCell(`G${rowD2}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0000' } };
        }

        rowIdx++;
    }

    const commonSheet = workbook.addWorksheet('Шпора общая', { properties: { tabColor: { argb: 'FFE0E0FF' } } });
    const cheatRows = cheatSheet.rowCount;
    const cheatCols = cheatSheet.columnCount || 7;
    for (let r = 1; r <= cheatRows; r++) {
        for (let c = 1; c <= cheatCols; c++) {
            const srcCell = cheatSheet.getCell(r, c);
            const dstCell = commonSheet.getCell(r, c);
            if (srcCell.value) {
                dstCell.value = srcCell.value;
            }
            if (r === 2 && c === 3) {
                dstCell.font = srcCell.font;
            }
        }
    }
}

// ---------- ОЧИСТКА АДРЕСА ----------
function extractAddressPart(fullName) {
    if (!fullName) return '';
    let address = fullName;
    const eqIndex = address.indexOf('=');
    if (eqIndex !== -1) {
        address = address.substring(eqIndex + 1).trim();
    }
    address = address.replace(/^ПАО\s*_?\s*["']?ВЫМПЕЛКОМ["']?\s*_?\s*/i, '');
    const prefixes = [
        'ПАО _ВЫМПЕЛКОМ_',
        'ПАО ВЫМПЕЛКОМ',
        'ПАО_ВЫМПЕЛКОМ_',
        'ПАО "ВЫМПЕЛКОМ"',
        "ПАО 'ВЫМПЕЛКОМ'",
        'ПАО "ВЫМПЕЛКОМ" ',
        'ПАО "ВЫМПЕЛКОМ"',
        'ПАО "ВЫМПЕЛКОМ" '
    ];
    for (const prefix of prefixes) {
        if (address.startsWith(prefix)) {
            address = address.substring(prefix.length).trim();
            break;
        }
    }
    address = address.trim();
    address = address.replace(/^["']+|["']+$/g, '');
    return address;
}

// ---------- СОХРАНЕНИЕ НА СЕТЕВОЙ ДИСК ----------
function saveFileToNetwork(buffer, fileName, projectName) {
    const basePath = process.env.BASE_NETWORK_PATH || '//fileserver/!_Work/for Druzhinin Anton/vhd';
    const safeProjectName = projectName.replace(/[\\/:*?"<>|]/g, '_');
    const projectFolder = path.join(
        basePath,
        '1.Техническая документация',
        '1.2.Проект',
        '1.2.3.Расшивки',
        safeProjectName
    );
    const fullPath = path.join(projectFolder, fileName);

    if (!fs.existsSync(projectFolder)) {
        fs.mkdirSync(projectFolder, { recursive: true });
    }

    fs.writeFileSync(fullPath, buffer);
    return fullPath;
}

// ---------- МАРШРУТЫ ГЕНЕРАЦИИ ----------
app.post('/generate-sv005', async (req, res) => {
    try {
        const { address, globalModel, boards } = req.body;
        console.log(`[SV005] Генерация: адрес="${address}", плат=${boards?.length || 0}`);
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

        const neededInput = boards.length;
        const neededOutput = boards.length;

        let inputBlocks = blocks.filter(b => b.type === 'input');
        let outputBlocks = blocks.filter(b => b.type === 'output');

        console.log(`[SV005] Блоков: входных=${inputBlocks.length}, выходных=${outputBlocks.length}, нужно пар=${neededInput}`);

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

            const plinth1Data = {
                rack: board.rack,
                boardNumber: boardNumber,
                plinthNumber: board.plinth1.number,
                skud: board.skud1,
                room: board.plinth1.room,
                terminalMap: board.plinth1.terminalMap,
                cableNumbers: board.plinth1.cableNumbers
            };
            await fillPlinthBlockSV005(
                worksheet,
                usedInputBlocks[i].startRow,
                'input',
                plinth1Data,
                globalModel
            );
            board.plinth1.deviceLabel = plinth1Data.deviceLabel;

            const plinth2Data = {
                rack: board.rack,
                boardNumber: boardNumber,
                plinthNumber: board.plinth2.number,
                skud: board.skud2,
                room: board.plinth2.room,
                terminalMap: board.plinth2.terminalMap,
                cableNumbers: board.plinth2.cableNumbers
            };
            await fillPlinthBlockSV005(
                worksheet,
                usedOutputBlocks[i].startRow,
                'output',
                plinth2Data,
                globalModel
            );
            board.plinth2.deviceLabel = plinth2Data.deviceLabel;
        }
        console.log(`[SV005] Данные плат заполнены`);

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

        // Удаляем строки после последнего использованного блока
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

        createSheetsSV005(workbook, boards, globalModel);
        console.log(`[SV005] Шпоргалки/Disp заполнены`);

        applyAutoFit(workbook);

        const buffer = await workbook.xlsx.writeBuffer();

        const addressPart = extractAddressPart(address);
        const safeAddress = addressPart.replace(/[\/\\:*?"<>|]/g, '_');
        const dateStr = new Date().toISOString().slice(0, 10);
        const filename = `${safeAddress}_${dateStr}_SV005.xlsx`;

        const filePath = saveFileToNetwork(buffer, filename, address);
        console.log(`[SV005] Файл сохранён: ${filePath}`);

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
        res.setHeader('X-File-Path', encodeURIComponent(filePath));
        res.send(buffer);
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Ошибка генерации SV005: ' + err.message });
    }
});

app.post('/generate-sv004', async (req, res) => {
    try {
        const { address, globalModel, boards } = req.body;
        if (!address || !boards || !boards.length) {
            return res.status(400).json({ error: 'Не указан адрес или список плат' });
        }
        if (!fs.existsSync(TEMPLATE_SV004)) {
            return res.status(500).json({ error: 'Файл шаблона SV004 не найден.' });
        }
        console.log(`[SV004] Генерация: адрес="${address}", плат=${boards.length}`);

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(TEMPLATE_SV004);

        let worksheet = workbook.getWorksheet('SV777-1 (SV004)');
        if (!worksheet) {
            worksheet = workbook.worksheets[0];
            if (!worksheet) {
                return res.status(500).json({ error: 'В файле шаблона нет ни одного листа.' });
            }
        }

        let allBlocks = getBlocksSV004(worksheet);
        if (allBlocks.length === 0) {
            return res.status(500).json({ error: 'В шаблоне SV004 не найдено блоков "Стойка".' });
        }
        allBlocks.sort((a, b) => a.startRow - b.startRow);

        const neededPairs = boards.length;
        let inputBlocks = allBlocks.filter(b => b.type === 'input');
        let outputBlocks = allBlocks.filter(b => b.type === 'output');

        while (inputBlocks.length < neededPairs) {
            const lastInput = inputBlocks[inputBlocks.length - 1];
            if (!lastInput) {
                return res.status(500).json({ error: 'Нет входных блоков для копирования.' });
            }
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
            if (!lastOutput) {
                return res.status(500).json({ error: 'Нет выходных блоков для копирования.' });
            }
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
                    cableMap: board.plinth1.cableMap,
                    roomMap: board.plinth1.roomMap || {}
                },
                globalModel
            );

            await fillPlinthBlockSV004(
                worksheet,
                usedOutputBlocks[i].startRow,
                {
                    rack: board.rack,
                    boardNumber: boardNumber,
                    plinthNumber: board.plinth2.number,
                    skud: board.skud2,
                    room: board.plinth2.room,
                    terminalMap: board.plinth2.terminalMap,
                    cableMap: board.plinth2.cableMap,
                    roomMap: board.plinth2.roomMap || {}
                },
                globalModel
            );
        }
        console.log(`[SV004] Данные плат заполнены`);

        const usedStartRows = new Set();
        usedInputBlocks.forEach(b => usedStartRows.add(b.startRow));
        usedOutputBlocks.forEach(b => usedStartRows.add(b.startRow));
        const blocksToDelete = allBlocks.filter(b => !usedStartRows.has(b.startRow));
        blocksToDelete.sort((a, b) => b.startRow - a.startRow);
        for (const block of blocksToDelete) {
            const rowCount = block.endRow - block.startRow + 1;
            worksheet.spliceRows(block.startRow, rowCount);
        }

        // Удаляем строки после последнего использованного блока
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

        fillSheetsSV004(workbook, boards, globalModel);
        console.log(`[SV004] Шпоргалки/Disp заполнены`);

        // Создаём общую шпоргалку
        createCommonCheatSheetSV004(workbook);

        applyAutoFit(workbook);
        console.log(`[SV004] Автоподгон ширины завершён`);

        const buffer = await workbook.xlsx.writeBuffer();

        const addressPart = extractAddressPart(address);
        const safeAddress = addressPart.replace(/[\\/:*?"<>|]/g, '_');
        const dateStr = new Date().toISOString().slice(0, 10);
        const filename = `${safeAddress}_${dateStr}_SV004.xlsx`;

        const filePath = saveFileToNetwork(buffer, filename, address);
        console.log(`[SV004] Файл сохранён: ${filePath}`);

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
        res.setHeader('X-File-Path', encodeURIComponent(filePath));
        res.send(buffer);
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Ошибка генерации SV004: ' + err.message });
    }
});

// ---------- СИНХРОНИЗАЦИЯ С SSE ----------
app.get('/api/projects/sync-stream', async (req, res) => {
    const apiKey = process.env.API_KEY;
    const baseUrl = process.env.API_BASE_URL;

    if (!apiKey || !baseUrl) {
        res.writeHead(500, { 'Content-Type': 'text/plain' });
        res.end('Ошибка: API_KEY или API_BASE_URL не заданы в .env');
        return;
    }

    res.writeHead(200, {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Access-Control-Allow-Origin': '*'
    });

    const sendEvent = (data) => {
        try {
            res.write(`data: ${JSON.stringify(data)}\n\n`);
        } catch (err) {
            console.error('Ошибка отправки SSE:', err);
        }
    };

    try {
        sendEvent({ type: 'status', message: 'Начинаем загрузку списка проектов...' });

        let allProjects = [];
        let page = 1;
        const perPage = 50;
        let total = 0;

        while (true) {
            const url = `${baseUrl}?api_key=${apiKey}&page=${page}&per_page=${perPage}`;
            sendEvent({ type: 'status', message: `Загружаем страницу ${page}...` });
            const response = await fetch(url);
            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`API вернул ${response.status}: ${errorText}`);
            }
            const data = await response.json();
            const items = data.response?.items || [];
            total = data.response?.total || 0;
            allProjects = allProjects.concat(items);
            sendEvent({ type: 'progress', loaded: allProjects.length, total: total });

            if (allProjects.length >= total) break;
            if (items.length < perPage) break;
            page++;
            if (page > 30) break;
        }

        sendEvent({ type: 'status', message: `Загружено ${allProjects.length} проектов. Обрабатываем детали...` });

        const client = await pool.connect();
        let processed = 0;
        try {
            await client.query('BEGIN');
            for (const project of allProjects) {
                const { id, name, is_archive } = project;
                const isArchiveBool = (is_archive === 1 || is_archive === true);
                let cf_92 = null;
                let cf_217 = null;
                try {
                    const detailUrl = `https://inteko.aspro.cloud/api/v1/module/st/projects/get/${id}?api_key=${apiKey}`;
                    const detailResponse = await fetch(detailUrl);
                    if (detailResponse.ok) {
                        const detailData = await detailResponse.json();
                        cf_92 = detailData.response?.cf_92 || null;
                        cf_217 = detailData.response?.cf_217 || null;
                    }
                } catch (err) {
                    // игнорируем ошибки получения деталей
                }

                const check = await client.query('SELECT id FROM projects WHERE id = $1', [id]);
                if (check.rows.length === 0) {
                    await client.query(
                        'INSERT INTO projects (id, name, is_archive, expanded_downloaded, cf_92, cf_217) VALUES ($1, $2, $3, $4, $5, $6)',
                        [id, name, isArchiveBool, false, cf_92, cf_217]
                    );
                    console.log(`[БД] INSERT project id=${id}, name="${name}"`);
                } else {
                    await client.query(
                        'UPDATE projects SET name = $1, is_archive = $2, cf_92 = $3, cf_217 = $4, updated_at = CURRENT_TIMESTAMP WHERE id = $5',
                        [name, isArchiveBool, cf_92, cf_217, id]
                    );
                    console.log(`[БД] UPDATE project id=${id}, name="${name}"`);
                }
                processed++;
                if (processed % 10 === 0 || processed === allProjects.length) {
                    sendEvent({ type: 'progress', processed: processed, total: allProjects.length });
                }
            }
            await client.query('COMMIT');
            sendEvent({ type: 'done', total: allProjects.length });
        } catch (err) {
            await client.query('ROLLBACK');
            throw err;
        } finally {
            client.release();
        }
    } catch (error) {
        console.error('Ошибка синхронизации:', error);
        sendEvent({ type: 'error', message: error.message });
    } finally {
        res.end();
    }
});

// ---------- ОБЫЧНЫЙ POST ДЛЯ ОБРАТНОЙ СОВМЕСТИМОСТИ ----------
app.post('/api/projects/sync', async (req, res) => {
    const apiKey = process.env.API_KEY;
    const baseUrl = process.env.API_BASE_URL;

    if (!apiKey || !baseUrl) {
        return res.status(500).json({ error: 'Не заданы API_KEY или API_BASE_URL в .env' });
    }

    try {
        let allProjects = [];
        let page = 1;
        const perPage = 50;
        let total = 0;

        while (true) {
            const url = `${baseUrl}?api_key=${apiKey}&page=${page}&per_page=${perPage}`;
            const response = await fetch(url);
            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`API вернул ${response.status}: ${errorText}`);
            }
            const data = await response.json();
            const items = data.response?.items || [];
            total = data.response?.total || 0;
            allProjects = allProjects.concat(items);

            if (allProjects.length >= total) break;
            if (items.length < perPage) break;
            page++;
            if (page > 30) break;
        }

        const client = await pool.connect();
        try {
            await client.query('BEGIN');
            for (const project of allProjects) {
                const { id, name, is_archive } = project;
                const isArchiveBool = (is_archive === 1 || is_archive === true);
                let cf_92 = null;
                let cf_217 = null;
                try {
                    const detailUrl = `https://inteko.aspro.cloud/api/v1/module/st/projects/get/${id}?api_key=${apiKey}`;
                    const detailResponse = await fetch(detailUrl);
                    if (detailResponse.ok) {
                        const detailData = await detailResponse.json();
                        cf_92 = detailData.response?.cf_92 || null;
                        cf_217 = detailData.response?.cf_217 || null;
                    }
                } catch (err) {
                    console.warn(`Не удалось получить cf_92/cf_217 для проекта ${id}:`, err.message);
                }

                const check = await client.query('SELECT id FROM projects WHERE id = $1', [id]);
                if (check.rows.length === 0) {
                    await client.query(
                        'INSERT INTO projects (id, name, is_archive, expanded_downloaded, cf_92, cf_217) VALUES ($1, $2, $3, $4, $5, $6)',
                        [id, name, isArchiveBool, false, cf_92, cf_217]
                    );
                    console.log(`[БД] INSERT project id=${id}, name="${name}"`);
                } else {
                    await client.query(
                        'UPDATE projects SET name = $1, is_archive = $2, cf_92 = $3, cf_217 = $4, updated_at = CURRENT_TIMESTAMP WHERE id = $5',
                        [name, isArchiveBool, cf_92, cf_217, id]
                    );
                    console.log(`[БД] UPDATE project id=${id}, name="${name}"`);
                }
            }
            await client.query('COMMIT');
        } catch (err) {
            await client.query('ROLLBACK');
            throw err;
        } finally {
            client.release();
        }

        res.json({ success: true, total: allProjects.length });
    } catch (error) {
        console.error('Ошибка синхронизации:', error);
        res.status(500).json({ error: error.message });
    }
});

// ---------- ОБНОВЛЕНИЕ СТАТУСА ----------
app.post('/api/projects/update-status', async (req, res) => {
    const { projects } = req.body;

    if (!projects || !Array.isArray(projects) || projects.length === 0) {
        return res.status(400).json({ error: 'Не передан список проектов' });
    }

    try {
        const client = await pool.connect();
        let updatedCount = 0;
        let createdCount = 0;
        try {
            await client.query('BEGIN');
            for (const item of projects) {
                const { name, fileName, filePath } = item;
                if (!fileName) continue;

                const existing = await client.query(
                    'SELECT id FROM projects WHERE file_name = $1',
                    [fileName]
                );
                if (existing.rows.length > 0) {
                    await client.query(
                        `UPDATE projects 
                         SET name = $1, file_path = $2, expanded_downloaded = true, updated_at = CURRENT_TIMESTAMP 
                         WHERE file_name = $3`,
                        [name, filePath, fileName]
                    );
                    updatedCount++;
                    console.log(`[БД] UPDATE status fileName="${fileName}"`);
                } else {
                    const tempId = -Date.now() - Math.floor(Math.random() * 1000);
                    await client.query(
                        `INSERT INTO projects 
                         (id, name, is_archive, expanded_downloaded, file_name, file_path) 
                         VALUES ($1, $2, $3, $4, $5, $6)`,
                        [tempId, name, false, true, fileName, filePath]
                    );
                    createdCount++;
                    console.log(`[БД] INSERT status fileName="${fileName}"`);
                }
            }
            await client.query('COMMIT');
        } catch (err) {
            await client.query('ROLLBACK');
            throw err;
        } finally {
            client.release();
        }

        res.json({ success: true, updated: updatedCount, created: createdCount });
    } catch (error) {
        console.error('Ошибка обновления статуса:', error);
        res.status(500).json({ error: error.message });
    }
});

// ---------- ПОЛУЧЕНИЕ ПРОЕКТОВ ----------
app.get('/api/projects', async (req, res) => {
    const { onlyNotDownloaded } = req.query;
    try {
        let query = 'SELECT id, name, is_archive, expanded_downloaded, file_name, file_path, cf_92, cf_217, updated_at FROM projects';
        const params = [];
        if (onlyNotDownloaded === 'true') {
            query += ' WHERE expanded_downloaded = false';
        }
        query += ' ORDER BY name';
        const result = await pool.query(query, params);
        res.json({ items: result.rows });
    } catch (error) {
        console.error('Ошибка получения проектов:', error);
        res.status(500).json({ error: error.message });
    }
});

// ---------- СКАЧИВАНИЕ ФАЙЛА ----------
app.get('/api/projects/download/:fileName', async (req, res) => {
    const { fileName } = req.params;
    const basePath = process.env.BASE_NETWORK_PATH || '//fileserver/!_Work/for Druzhinin Anton/vhd';

    try {
        const result = await pool.query(
            'SELECT file_path FROM projects WHERE file_name = $1 AND file_path IS NOT NULL',
            [fileName]
        );
        if (result.rows.length > 0) {
            const filePath = result.rows[0].file_path;
            if (fs.existsSync(filePath)) {
                return res.download(filePath, fileName, (err) => {
                    if (err) {
                        console.error('Ошибка при скачивании файла:', err);
                        res.status(500).json({ error: 'Ошибка при скачивании файла' });
                    }
                });
            }
        }
    } catch (err) {
        console.warn('Ошибка поиска пути в БД:', err.message);
    }

    let foundPath = null;
    const walk = (dir) => {
        try {
            const files = fs.readdirSync(dir);
            for (const file of files) {
                const full = path.join(dir, file);
                try {
                    const stat = fs.statSync(full);
                    if (stat.isDirectory()) {
                        walk(full);
                    } else if (file === fileName) {
                        foundPath = full;
                        return;
                    }
                } catch (e) {}
            }
        } catch (e) {}
    };
    try {
        walk(basePath);
    } catch (err) {
        return res.status(404).json({ error: 'Файл не найден' });
    }

    if (!foundPath) {
        return res.status(404).json({ error: 'Файл не найден' });
    }

    res.download(foundPath, fileName, (err) => {
        if (err) {
            console.error('Ошибка при скачивании файла:', err);
            res.status(500).json({ error: 'Ошибка при скачивании файла' });
        }
    });
});

// ---------- ПОИСК ПРОЕКТОВ ----------
app.get('/api/projects/search', async (req, res) => {
    const { search, showArchived } = req.query;
    const apiKey = process.env.API_KEY;
    const baseUrl = process.env.API_BASE_URL;

    if (!apiKey || !baseUrl) {
        return res.status(500).json({ error: 'Не заданы API_KEY или API_BASE_URL в .env' });
    }

    if (!search || search.length < 2) {
        return res.json({ items: [] });
    }

    try {
        let allProjects = [];
        let page = 1;
        const perPage = 50;
        let total = 0;

        while (true) {
            const url = `${baseUrl}?api_key=${apiKey}&page=${page}&per_page=${perPage}`;
            console.log(`🔍 Поиск: запрос к ${url}`);
            const response = await fetch(url);
            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`API вернул ${response.status}: ${errorText}`);
            }
            const data = await response.json();
            const items = data.response?.items || [];
            total = data.response?.total || 0;
            allProjects = allProjects.concat(items);

            if (allProjects.length >= total) break;
            if (items.length < perPage) break;
            page++;
            if (page > 30) break;
        }

        const lowerSearch = search.toLowerCase();
        let filteredBySearch = allProjects.filter(p => p.name && p.name.toLowerCase().includes(lowerSearch));
        const totalBeforeArchive = filteredBySearch.length;

        let filteredByArchive = filteredBySearch;
        let archivedCount = 0;
        if (showArchived !== 'true') {
            filteredByArchive = filteredBySearch.filter(p => !p.is_archive);
            archivedCount = totalBeforeArchive - filteredByArchive.length;
        }

        const limit = 100;
        if (filteredByArchive.length > limit) {
            filteredByArchive = filteredByArchive.slice(0, limit);
        }

        res.json({
            items: filteredByArchive,
            total_before_filter: totalBeforeArchive,
            total_after_filter: filteredByArchive.length,
            archived_count: archivedCount
        });
    } catch (error) {
        console.error('Ошибка при поиске проектов:', error);
        res.status(500).json({ error: error.message });
    }
});

// ---------- ЗАПУСК ----------
app.listen(PORT, () => {
    console.log(`✅ Сервер запущен: http://localhost:${PORT}`);
});