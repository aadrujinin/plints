const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 3000;

app.use(express.json());
app.use(express.static('public'));

const TEMPLATE_PATH = path.join(__dirname, 'templateSV005_test.xlsx');

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

function getBlocks(worksheet) {
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
                blocks.push({ startRow: row, type });
            } else {
                console.log(`⚠️ Блок в строке ${row} не имеет заголовка R1/D1 или R2/D2`);
            }
        }
    }
    console.log(`✅ Найдено блоков: input=${blocks.filter(b => b.type === 'input').length}, output=${blocks.filter(b => b.type === 'output').length}`);
    return blocks;
}

async function fillPlinthBlock(worksheet, startRow, blockType, plinthData, globalModel) {
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
    const holderValue = `ХВ1 -${rackClean}.${plinthData.plinthNumber}`;
    worksheet.getCell(`N${rowHolder}`).value = holderValue;
    worksheet.getCell(`N${rowPlinthNum}`).value = plinthData.plinthNumber;
    worksheet.getCell(`N${rowSkud}`).value = plinthData.skud;

    // Поиск строки устройств
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
    console.log(`Блок ${blockType}, startRow=${startRow}, devicesRow=${devicesRow}`);

    const tm = plinthData.terminalMap || {};
    const cn = plinthData.cableNumbers || {};

    const hasReader = Object.values(tm).includes('reader');
    if (hasReader) {
        worksheet.getCell(`D${devicesRow}`).value = (blockType === 'input') ? 'Считыватель вх' : 'Считыватель вых';
        worksheet.getCell(`F${devicesRow}`).value = cn.reader || '';
    } else {
        worksheet.getCell(`D${devicesRow}`).value = '';
        worksheet.getCell(`F${devicesRow}`).value = '';
    }

    const hasLock = Object.values(tm).includes('lock');
    const hasFireLock = Object.values(tm).includes('fire_lock');
    if (hasLock) {
        const cellJ = worksheet.getCell(`J${devicesRow}`);
        cellJ.value = 'Замок';
        //cellJ.alignment = { textRotation: 255 };
        worksheet.getCell(`K${devicesRow}`).value = cn.lock || '';
        console.log(`  записан замок, кабель=${cn.lock}`);
    } else if (hasFireLock) {
        const cellJ = worksheet.getCell(`J${devicesRow}`);
        cellJ.value = 'Замок пож.дв.';
        cellJ.alignment = { textRotation: 255 };
        worksheet.getCell(`K${devicesRow}`).value = cn.fire_lock || '';
        console.log(`  записан замок пож.дв., кабель=${cn.fire_lock}`);
    } else {
        worksheet.getCell(`J${devicesRow}`).value = '';
        worksheet.getCell(`K${devicesRow}`).value = '';
    }

    const hasContact = Object.values(tm).includes('contact');
    worksheet.getCell(`N${devicesRow}`).value = hasContact ? 'геркон' : '';
    worksheet.getCell(`J${devicesRow}`).value = hasContact ? 'замок' : '';
    worksheet.getCell(`O${devicesRow}`).value = hasContact ? (cn.contact || '') : '';

    const hasExit = Object.values(tm).includes('exit_btn');
    worksheet.getCell(`P${devicesRow}`).value = hasExit ? 'кн.Вых' : '';
    worksheet.getCell(`Q${devicesRow}`).value = hasExit ? (cn.exit_btn || '') : '';

    const hasSiren = Object.values(tm).includes('siren');
    worksheet.getCell(`L${devicesRow}`).value = hasSiren ? 'сирена' : '';
    worksheet.getCell(`M${devicesRow}`).value = hasSiren ? (cn.siren || '') : '';

    if (blockType === 'input') {
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
}

app.post('/generate', async (req, res) => {
    try {
        const { address, globalModel, boards } = req.body;
        if (!address || !boards || !boards.length) {
            return res.status(400).json({ error: 'Не указан адрес или список плат' });
        }
        if (!fs.existsSync(TEMPLATE_PATH)) {
            return res.status(500).json({ error: 'Файл шаблона не найден.' });
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(TEMPLATE_PATH);

        let worksheet = workbook.getWorksheet('Лист1');
        if (!worksheet) {
            worksheet = workbook.worksheets[0];
            if (!worksheet) {
                return res.status(500).json({ error: 'В файле шаблона нет ни одного листа.' });
            }
            console.log(`Используется лист: ${worksheet.name}`);
        }

        // Получаем все блоки
        let blocks = getBlocks(worksheet);
        if (blocks.length === 0) {
            return res.status(500).json({ error: 'В шаблоне не найдено блоков "Стойка".' });
        }
        // Сортируем по startRow
        blocks.sort((a, b) => a.startRow - b.startRow);
        // Вычисляем endRow для каждого блока
        for (let i = 0; i < blocks.length; i++) {
            const nextStart = (i < blocks.length - 1) ? blocks[i+1].startRow : worksheet.rowCount + 1;
            blocks[i].endRow = nextStart - 1;
        }

        const inputBlocks = blocks.filter(b => b.type === 'input');
        const outputBlocks = blocks.filter(b => b.type === 'output');

        const neededInput = boards.length;
        const neededOutput = boards.length;

        if (inputBlocks.length < neededInput || outputBlocks.length < neededOutput) {
            return res.status(400).json({
                error: `Не хватает блоков в шаблоне. Нужно входных: ${neededInput}, доступно: ${inputBlocks.length}; выходных: ${neededOutput}, доступно: ${outputBlocks.length}.`
            });
        }

        // Берем первые N блоков каждого типа
        const usedInputBlocks = inputBlocks.slice(0, neededInput);
        const usedOutputBlocks = outputBlocks.slice(0, neededOutput);
        const usedStartRows = new Set([...usedInputBlocks.map(b => b.startRow), ...usedOutputBlocks.map(b => b.startRow)]);

        // Заполняем используемые блоки
        for (let i = 0; i < boards.length; i++) {
            const board = boards[i];
            const boardNumber = i + 1;

            await fillPlinthBlock(
                worksheet,
                usedInputBlocks[i].startRow,
                'input',
                {
                    rack: board.rack,
                    boardNumber: boardNumber,
                    plinthNumber: 2 * boardNumber - 1,
                    skud: board.skud1,
                    room: board.plinth1.room,
                    terminalMap: board.plinth1.terminalMap,
                    cableNumbers: board.plinth1.cableNumbers
                },
                globalModel
            );

            await fillPlinthBlock(
                worksheet,
                usedOutputBlocks[i].startRow,
                'output',
                {
                    rack: board.rack,
                    boardNumber: boardNumber,
                    plinthNumber: 2 * boardNumber,
                    skud: board.skud2,
                    room: board.plinth2.room,
                    terminalMap: board.plinth2.terminalMap,
                    cableNumbers: board.plinth2.cableNumbers
                },
                globalModel
            );
        }

        // Удаляем неиспользуемые блоки (в обратном порядке)
        const unusedBlocks = blocks.filter(b => !usedStartRows.has(b.startRow));
        // Сортируем по убыванию startRow
        unusedBlocks.sort((a, b) => b.startRow - a.startRow);
        for (const block of unusedBlocks) {
            const rowCount = block.endRow - block.startRow + 1;
            if (rowCount > 0) {
                worksheet.spliceRows(block.startRow, rowCount);
                console.log(`🗑 Удалены строки ${block.startRow}-${block.endRow} (блок ${block.type})`);
            }
        }

        const buffer = await workbook.xlsx.writeBuffer();
        const dateStr = new Date().toISOString().slice(0, 10);
        const safeAddress = address.replace(/[\\/:*?"<>|]/g, '_');
        const filename = `${safeAddress}_${dateStr}.xlsx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
        res.send(buffer);
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Ошибка генерации: ' + err.message });
    }
});

app.listen(PORT, () => {
    console.log(`Сервер запущен: http://localhost:${PORT}`);
});