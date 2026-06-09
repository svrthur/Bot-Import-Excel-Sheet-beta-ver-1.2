/**
 * ==========================
 * КАЛЕНДАРЬ РОЛИКОВ — ПОЛНЫЙ СКРИПТ
 * ==========================
 * Часть 1 (ваш оригинал):
 *   — onEdit: пересчёт строки 2 при изменении N7
 *   — recolorExpiredByToday: перекрашивает просроченные зелёные → красные
 *   — recalcTotalsByN7: ручной пересчёт
 *
 * Часть 2 (Telegram-бот):
 *   — doPost / highlightCampaigns: подсветка ячеек зелёным по команде из бота
 *   — doGet: проверка порога 160 с (для алертов бота); ?action=calculate → пересчёт строки 2
 *   — hideFinishedCampaigns: скрытие завершённых кампаний (запускать по триггеру)
 *   — checkTKThreshold: возвращает список ТК, превысивших 160 с
 */

// ─────────────────────────────────────────────
// КОНФИГУРАЦИЯ
// ─────────────────────────────────────────────
const CFG = {
  SHEET_NAME: 'Лист1',

  // ТК: R..GN
  TK_START_COL: 18,
  TK_END_COL: 196,

  TOTALS_ROW: 2,
  DATA_START_ROW: 3,

  DATE_CELL_A1: 'N7',

  // A Название, B Тип, C Длит-ть, D Владелец, E Статус, F Дата старта, G Дата окончания
  DURATION_COL: 3, // C
  STATUS_COL: 5,   // E
  START_COL: 6,    // F
  END_COL: 7,      // G

  // Цвета
  GREEN_HEXES: ['#00ff00'],
  RED_HEX: '#ff0000',

  // Тексты статусов (как в таблице; регистр не важен)
  STATUS_PLANNED:   'запланировано',
  STATUS_PUBLISHED: 'опубликовано',
  STATUS_DONE:      'завершено',

  SPREADSHEET_ID: '17VeQQWTGotofrpNbUHDhUFhCc3qjLdwoesTxDDfJ7h4',
};

// ─────────────────────────────────────────────
// ЧАСТЬ 1 — ВАШИ ФУНКЦИИ (без изменений)
// ─────────────────────────────────────────────

/**
 * onEdit — пересчёт итогов при изменении N7
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== CFG.SHEET_NAME) return;

    if (e.range.getA1Notation() !== CFG.DATE_CELL_A1) return;

    const selectedDate = toDate00_(sh.getRange(CFG.DATE_CELL_A1).getValue());
    if (!selectedDate) return;

    recalcTotalsRow_(sh, selectedDate);
  } catch (err) {
    console.error('onEdit error:', err);
  }
}

/**
 * Задача 1 — перекраска по "сегодня"
 */
function recolorExpiredByToday() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_NAME);
  if (!sh) throw new Error(`Лист "${CFG.SHEET_NAME}" не найден`);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  recolorExpiredGreenToRed_(sh, today);
}

function createDailyRecolorTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'recolorExpiredByToday') ScriptApp.deleteTrigger(t);
  }

  ScriptApp.newTrigger('recolorExpiredByToday')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
}

/**
 * Ручной пересчёт по N7 (для теста)
 */
function recalcTotalsByN7() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_NAME);
  if (!sh) throw new Error(`Лист "${CFG.SHEET_NAME}" не найден`);

  const selectedDate = toDate00_(sh.getRange(CFG.DATE_CELL_A1).getValue());
  if (!selectedDate) throw new Error('В N7 не дата (или формат не распознан)');

  recalcTotalsRow_(sh, selectedDate);
}

/**
 * ВНУТРЕННЯЯ: перекраска просроченных зелёных → красный
 */
function recolorExpiredGreenToRed_(sh, compareDate) {
  const lastRow = sh.getLastRow();
  if (lastRow < CFG.DATA_START_ROW) return;

  const numRows = lastRow - CFG.DATA_START_ROW + 1;
  const numCols = CFG.TK_END_COL - CFG.TK_START_COL + 1;

  const endVals = sh.getRange(CFG.DATA_START_ROW, CFG.END_COL, numRows, 1).getValues();

  const tkRange = sh.getRange(CFG.DATA_START_ROW, CFG.TK_START_COL, numRows, numCols);
  const bgs = tkRange.getBackgrounds();

  let changed = false;

  for (let r = 0; r < numRows; r++) {
    const endDate = toDate00_(endVals[r][0]);
    if (!endDate) continue;

    if (compareDate.getTime() <= endDate.getTime()) continue;

    for (let c = 0; c < numCols; c++) {
      if (isGreen_(bgs[r][c])) {
        bgs[r][c] = CFG.RED_HEX;
        changed = true;
      }
    }
  }

  if (changed) tkRange.setBackgrounds(bgs);
}

/**
 * ВНУТРЕННЯЯ: итоги хронометража для строки 2
 */
function recalcTotalsRow_(sh, selectedDate) {
  const lastRow = sh.getLastRow();
  if (lastRow < CFG.DATA_START_ROW) {
    clearTotalsRow_(sh);
    return;
  }

  const numRows = lastRow - CFG.DATA_START_ROW + 1;
  const numCols = CFG.TK_END_COL - CFG.TK_START_COL + 1;

  const durations = sh.getRange(CFG.DATA_START_ROW, CFG.DURATION_COL, numRows, 1).getValues();
  const statuses  = sh.getRange(CFG.DATA_START_ROW, CFG.STATUS_COL,   numRows, 1).getDisplayValues();
  const starts    = sh.getRange(CFG.DATA_START_ROW, CFG.START_COL,    numRows, 1).getValues();
  const ends      = sh.getRange(CFG.DATA_START_ROW, CFG.END_COL,      numRows, 1).getValues();

  const tkRange = sh.getRange(CFG.DATA_START_ROW, CFG.TK_START_COL, numRows, numCols);
  const bgs = tkRange.getBackgrounds();

  const totals = new Array(numCols).fill(0);

  for (let r = 0; r < numRows; r++) {
    const startDate = toDate00_(starts[r][0]);
    const endDate   = toDate00_(ends[r][0]);
    if (!startDate || !endDate) continue;

    if (selectedDate.getTime() < startDate.getTime()) continue;
    if (selectedDate.getTime() > endDate.getTime()) continue;

    const effStatus = effectiveStatusOnDate_(statuses[r][0], startDate, endDate, selectedDate);
    if (effStatus === CFG.STATUS_DONE) continue;

    const sec = toNumber_(durations[r][0]);
    if (!(sec > 0)) continue;

    for (let c = 0; c < numCols; c++) {
      if (isGreen_(bgs[r][c])) totals[c] += sec;
    }
  }

  sh.getRange(CFG.TOTALS_ROW, CFG.TK_START_COL, 1, numCols).setValues([totals]);
}

function clearTotalsRow_(sh) {
  const numCols = CFG.TK_END_COL - CFG.TK_START_COL + 1;
  sh.getRange(CFG.TOTALS_ROW, CFG.TK_START_COL, 1, numCols).clearContent();
}

/**
 * Эффективный статус на дату
 */
function effectiveStatusOnDate_(rawStatus, startDate, endDate, date) {
  const st = normalizeStatus_(rawStatus);

  if (st === CFG.STATUS_PLANNED || st === CFG.STATUS_PUBLISHED || st === CFG.STATUS_DONE) {
    return st;
  }

  if (date.getTime() < startDate.getTime()) return CFG.STATUS_PLANNED;
  if (date.getTime() > endDate.getTime()) return CFG.STATUS_DONE;
  return CFG.STATUS_PUBLISHED;
}

// ─────────────────────────────────────────────
// ВСПОМОГАТЕЛЬНЫЕ
// ─────────────────────────────────────────────

function normalizeStatus_(s) {
  return (s || '').toString().trim().toLowerCase();
}

function isGreen_(hex) {
  if (!hex) return false;
  const h = String(hex).trim().toLowerCase();
  return CFG.GREEN_HEXES.some(g => h === String(g).trim().toLowerCase());
}

function toNumber_(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined) return 0;
  const n = Number(String(v).replace(',', '.').trim());
  return isNaN(n) ? 0 : n;
}

function toDate00_(v) {
  if (!v && v !== 0) return null;

  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    const d = new Date(v);
    d.setHours(0, 0, 0, 0);
    return d;
  }

  const s = String(v).trim();
  if (!s) return null;

  const m = s.match(/^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{2}|\d{4})$/);
  if (m) {
    const dd = parseInt(m[1], 10);
    const mm = parseInt(m[2], 10);
    let yy = parseInt(m[3], 10);
    if (yy < 100) yy += 2000;
    const d = new Date(yy, mm - 1, dd);
    if (!isNaN(d.getTime())) {
      d.setHours(0, 0, 0, 0);
      return d;
    }
  }

  const d2 = new Date(s);
  if (!isNaN(d2.getTime())) {
    d2.setHours(0, 0, 0, 0);
    return d2;
  }

  return null;
}

// ─────────────────────────────────────────────
// ЧАСТЬ 2 — TELEGRAM-БОТ
// ─────────────────────────────────────────────

/**
 * Открывает таблицу по ID (нужно для вызовов из бота/doGet/doPost,
 * где SpreadsheetApp.getActive() недоступен).
 */
function getSheet_() {
  const ss = SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
  return ss.getSheetByName(CFG.SHEET_NAME) || ss.getSheets()[0];
}

/**
 * doPost — точка входа для POST-запросов от Telegram-бота.
 * Тело запроса: { action: 'highlight', data: { '<rowNum>': ['ТК1','ТК2', ...] } }
 */
function doPost(e) {
  try {
    const requestData = JSON.parse(e.postData.contents);

    if (requestData.action === 'highlight') {
      highlightCampaigns(requestData.data);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Неизвестный action' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Закрашивает ячейки в зелёный (#00ff00) для переданных строк и ТК-кодов,
 * затем пересчитывает строку 2.
 * campaignData = { '234': ['203','227'], ... }
 */
function highlightCampaigns(campaignData) {
  const sh = getSheet_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  // Карта: заголовок ТК → номер столбца (1-based)
  const tkColMap = {};
  for (let col = 0; col < headers.length; col++) {
    const val = String(headers[col]).trim();
    if (val) tkColMap[val] = col + 1;
  }

  for (const rowKey in campaignData) {
    const tks = campaignData[rowKey];
    const rowIndex = parseInt(rowKey.replace(/\D/g, ''), 10);
    if (isNaN(rowIndex) || rowIndex <= 0) continue;

    for (let j = 0; j < tks.length; j++) {
      const tkNum = String(tks[j]).trim();
      const colIndex = tkColMap[tkNum];
      if (colIndex) {
        sh.getRange(rowIndex, colIndex).setBackground(CFG.GREEN_HEXES[0]);
      }
    }
  }

  // После подсветки пересчитываем строку 2 по дате из N7
  const selectedDate = toDate00_(sh.getRange(CFG.DATE_CELL_A1).getValue());
  if (selectedDate) recalcTotalsRow_(sh, selectedDate);
}

/**
 * Скрывает строки с кампаниями, у которых статус "завершено"
 * и дата окончания более 24 ч назад. Запускать по time-driven триггеру.
 */
function hideFinishedCampaigns() {
  const sh = getSheet_();
  const startRow = 120;
  const lastRow  = sh.getLastRow();
  if (lastRow < startRow) return;

  const statusRange = sh.getRange(startRow, CFG.STATUS_COL, lastRow - startRow + 1, 1);
  const values = statusRange.getValues();
  const now = Date.now();
  const DAY_MS = 24 * 60 * 60 * 1000;

  for (let i = 0; i < values.length; i++) {
    const rowNum = startRow + i;
    if (normalizeStatus_(values[i][0]) !== CFG.STATUS_DONE) continue;

    const endDateVal = sh.getRange(rowNum, CFG.END_COL).getValue();
    if (endDateVal instanceof Date && (now - endDateVal.getTime()) > DAY_MS) {
      sh.hideRows(rowNum);
    }
  }
}

/**
 * Проверяет строку 2 (R2:GN2): возвращает массив ТК, где значение > 160 с.
 * Используется ботом для предупреждений.
 */
function checkTKThreshold() {
  const sh = getSheet_();
  const numCols = CFG.TK_END_COL - CFG.TK_START_COL + 1;

  const values  = sh.getRange(CFG.TOTALS_ROW, CFG.TK_START_COL, 1, numCols).getValues()[0];
  const headers = sh.getRange(1,              CFG.TK_START_COL, 1, numCols).getValues()[0];

  const exceeded = [];
  for (let i = 0; i < values.length; i++) {
    const sec = parseFloat(values[i]);
    if (!isNaN(sec) && sec > 160) {
      exceeded.push(String(headers[i]).trim());
    }
  }

  return exceeded.length > 0 ? exceeded : null;
}

/**
 * doGet — точка входа для GET-запросов.
 *   ?action=calculate  → пересчитать строку 2 по N7 и вернуть результат
 *   (без параметра)    → проверить порог 160 с (для бота)
 */
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'calculate') {
    try {
      recalcTotalsByN7();
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // По умолчанию — проверка порогов
  const exceeded = checkTKThreshold();
  return ContentService.createTextOutput(JSON.stringify({ status: 'success', exceeded: exceeded }))
    .setMimeType(ContentService.MimeType.JSON);
}
