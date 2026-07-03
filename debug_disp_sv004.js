const ExcelJS = require('exceljs');
const path = require('path');

const base = __dirname;
const templateFile = path.join(base, 'templateSV004_test.xlsx');
const etalonFile = path.join(base, '(г.Курган_ ул.Карла Маркса_ 149)_2026-07-03_SV004_fixed (1).xlsx');

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
        if (cell.value.formula !== undefined) return 'formula:' + cell.value.formula;
        if (cell.value.sharedFormula) return 'sharedFormula:' + JSON.stringify(cell.value.sharedFormula);
        return JSON.stringify(cell.value);
    }
    return String(cell.value);
}

function getCellType(cell) {
    if (!cell || cell.value === undefined || cell.value === null) return 'empty';
    if (typeof cell.value === 'string') return 'str';
    if (typeof cell.value === 'object') {
        if (cell.value.formula) return 'formula';
        if (cell.value.sharedFormula) return 'sharedFormula';
        if (cell.value.result !== undefined) return 'formulaResult';
        if (cell.value.richText) return 'richText';
        if (cell.value.text) return 'obj(text)';
        return 'obj';
    }
    return typeof cell.value;
}

function padRight(s, len) {
    s = String(s);
    while (s.length < len) s += ' ';
    return s;
}

async function dumpDispSheet(file, label) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(file);
    
    const sheetNames = wb.worksheets.map(s => s.name);
    console.log('\\n' + label + ' — Листы: ' + sheetNames.join(', '));
    
    for (const sn of ['Disp-1', 'Disp-2', 'Disp-3']) {
        const ws = wb.getWorksheet(sn);
        if (!ws) {
            console.log('\\n=== ' + label + ' / ' + sn + ': НЕТ ТАКОГО ЛИСТА ===');
            continue;
        }
        console.log('\\n=== ' + label + ' / ' + sn + ' === rows=' + ws.rowCount + ', cols=' + ws.columnCount);
        
        let lastRowWithContent = 0;
        for (let r = 1; r <= ws.rowCount; r++) {
            const row = ws.getRow(r);
            let hasContent = false;
            for (let c = 1; c <= (ws.columnCount || 12); c++) {
                if (row.getCell(c).value) { hasContent = true; break; }
            }
            if (hasContent) lastRowWithContent = r;
        }
        console.log('  Last row with content: ' + lastRowWithContent);
        
        for (let r = 1; r <= Math.max(lastRowWithContent, 1); r++) {
            const row = ws.getRow(r);
            const vals = [];
            for (let c = 1; c <= 18; c++) {
                const cell = row.getCell(c);
                let v = getCellText(cell);
                let t = getCellType(cell);
                if (v) {
                    const colLetter = String.fromCharCode(64 + c);
                    vals.push(colLetter + c + '(' + t + ')="' + v.substring(0, 80) + '"');
                }
            }
            if (vals.length) {
                console.log('  R' + r + ': ' + vals.join(' | '));
            } else {
                console.log('  R' + r + ': [empty]');
            }
        }
    }
    console.log('');
}

async function compareDispSheet(sheetName) {
    console.log('\\n========== СРАВНЕНИЕ ' + sheetName + ': ШАБЛОН vs ЭТАЛОН ==========');
    
    const wbTpl = new ExcelJS.Workbook();
    await wbTpl.xlsx.readFile(templateFile);
    const wsTpl = wbTpl.getWorksheet(sheetName);
    
    const wbEtal = new ExcelJS.Workbook();
    await wbEtal.xlsx.readFile(etalonFile);
    const wsEtal = wbEtal.getWorksheet(sheetName);
    
    if (!wsTpl && !wsEtal) {
        console.log('  Лист не найден ни в одном файле');
        return;
    }
    if (!wsTpl) { console.log('  Лист есть только в эталоне'); return; }
    if (!wsEtal) { console.log('  Лист есть только в шаблоне'); return; }
    
    function lastContentRow(ws) {
        let last = 0;
        for (let r = 1; r <= ws.rowCount; r++) {
            const row = ws.getRow(r);
            for (let c = 1; c <= 18; c++) {
                if (row.getCell(c).value) { last = r; break; }
            }
        }
        return last;
    }
    
    const lastTpl = lastContentRow(wsTpl);
    const lastEtal = lastContentRow(wsEtal);
    
    console.log('  Шаблон последняя строка: ' + lastTpl + ', Эталон последняя строка: ' + lastEtal);
    
    console.log('');
    console.log(padRight('Строка', 6) + ' | ' + padRight('ШАБЛОН (template)', 90) + ' | ' + padRight('ЭТАЛОН (fixed)', 90));
    console.log('-'.repeat(6) + '-+-' + '-'.repeat(90) + '-+-' + '-'.repeat(90));
    
    const maxShow = Math.max(lastTpl, lastEtal);
    let diffCount = 0;
    
    for (let r = 1; r <= maxShow; r++) {
        const tplCells = {};
        const etalCells = {};
        
        for (let c = 1; c <= 18; c++) {
            const colLetter = String.fromCharCode(64 + c);
            tplCells[colLetter] = getCellText(wsTpl.getCell(r, c));
            etalCells[colLetter] = getCellText(wsEtal.getCell(r, c));
        }
        
        function buildRowStr(cells) {
            const parts = [];
            for (const cl of ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R']) {
                const v = cells[cl];
                if (v) parts.push(cl + '=' + v.substring(0, 60));
            }
            return parts.join(' | ');
        }
        
        const tplStr = buildRowStr(tplCells);
        const etalStr = buildRowStr(etalCells);
        
        if (tplStr || etalStr) {
            const hasDiff = tplStr !== etalStr;
            if (hasDiff) diffCount++;
            const marker = hasDiff ? ' <-- РАЗЛИЧИЕ' : '';
            
            console.log('\\nR' + r + marker + ':');
            
            for (const cl of ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R']) {
                const tVal = tplCells[cl];
                const eVal = etalCells[cl];
                if (tVal || eVal) {
                    const idx = cl.charCodeAt(0) - 64;
                    const tType = getCellType(wsTpl.getCell(r, idx));
                    const eType = getCellType(wsEtal.getCell(r, idx));
                    const diff = tVal !== eVal ? ' <<<' : '';
                    
                    const tDisplay = tVal ? tVal.substring(0, 70) : '(empty)';
                    const eDisplay = eVal ? eVal.substring(0, 70) : '(empty)';
                    
                    if (tVal !== eVal) {
                        console.log('  ' + cl + ': шаблон(' + tType + ')="' + tDisplay + '" | эталон(' + eType + ')="' + eDisplay + '"' + diff);
                    } else {
                        console.log('  ' + cl + ': "' + tDisplay + '" (ok)');
                    }
                }
            }
        }
    }
    
    console.log('\\nИтого различий: ' + diffCount);
}

async function main() {
    console.log('========== АНАЛИЗ ЛИСТОВ Disp ДЛЯ SV004 ==========');
    console.log('Шаблон: ' + templateFile);
    console.log('Эталон: ' + etalonFile);
    
    await dumpDispSheet(templateFile, 'ШАБЛОН');
    await dumpDispSheet(etalonFile, 'ЭТАЛОН');
    
    await compareDispSheet('Disp-1');
    await compareDispSheet('Disp-2');
    await compareDispSheet('Disp-3');
    
    console.log('\\n========== ПРОВЕРКА ИСТОЧНИКА: шпоргалка-1 ==========');
    const wbTpl = new ExcelJS.Workbook();
    await wbTpl.xlsx.readFile(templateFile);
    const wsCheatTpl = wbTpl.getWorksheet('шпоргалка-1');
    console.log('\\nШАБЛОН шпоргалка-1: rows=' + wsCheatTpl.rowCount);
    for (let r = 1; r <= Math.min(wsCheatTpl.rowCount, 20); r++) {
        const cells = {};
        for (let c = 1; c <= 7; c++) {
            cells[c] = getCellText(wsCheatTpl.getCell(r, c));
        }
        const parts = [];
        for (let c = 1; c <= 7; c++) {
            if (cells[c]) parts.push('C' + c + '="' + cells[c].substring(0, 60) + '"');
        }
        if (parts.length) console.log('  R' + r + ': ' + parts.join(' | '));
    }
    
    const wbEtal = new ExcelJS.Workbook();
    await wbEtal.xlsx.readFile(etalonFile);
    const wsCheatEtal = wbEtal.getWorksheet('шпоргалка-1');
    console.log('\\nЭТАЛОН шпоргалка-1: rows=' + wsCheatEtal.rowCount);
    for (let r = 1; r <= Math.min(wsCheatEtal.rowCount, 20); r++) {
        const cells = {};
        for (let c = 1; c <= 7; c++) {
            cells[c] = getCellText(wsCheatEtal.getCell(r, c));
        }
        const parts = [];
        for (let c = 1; c <= 7; c++) {
            if (cells[c]) parts.push('C' + c + '="' + cells[c].substring(0, 60) + '"');
        }
        if (parts.length) console.log('  R' + r + ': ' + parts.join(' | '));
    }
}

main().catch(console.error);
