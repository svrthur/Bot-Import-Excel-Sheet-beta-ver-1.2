 n;
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

// ─────────────────────────────────────────────
// ДИАГНОСТИКА — запустите вручную при проблемах с нулями в строке 2
// ─────────────────────────────────────────────

/**
 * Запустите эту функцию вручную (▶ Run → debugTKRow2).
 * Результат смотрите в View → Logs (Ctrl+Enter).
 * Покажет: имя листа, дату N7, уникальные цвета зелёных ячеек,
 * статусы и строки, которые участвуют в подсчёте.
 */
function debugTKRow2() {
  // 1. Имя активного листа
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets().map(s => s.getName());
  Logger.log('Все листы: ' + JSON.stringify(sheets));

  const sh = ss.getSheets()[0]; // берём первый лист независимо от имени
  Logger.log('Активный лист (используется): ' + sh.getName());

  // 2. Дата в N7
  const rawN7 = sh.getRange(CFG.DATE_CELL_A1).getValue();
  const selectedDate = toDate00_(rawN7);
  Logger.log('N7 raw: ' + rawN7 + ' | parsed: ' + selectedDate);
  if (!selectedDate) {
    Logger.log('❌ N7 не распознана как дата — пересчёт невозможен');
    return;
  }

  const lastRow = sh.getLastRow();
  const numRows = lastRow - CFG.DATA_START_ROW + 1;
  const numCols = CFG.TK_END_COL - CFG.TK_START_COL + 1;

  if (numRows <= 0) {
    Logger.log('❌ Нет строк данных (DATA_START_ROW=' + CFG.DATA_START_ROW + ', lastRow=' + lastRow + ')');
    return;
  }

  // 3. Все уникальные цвета в блоке ТК (первые 50 строк, чтобы не тормозить)
  const sampleRows = Math.min(numRows, 50);
  const bgs = sh.getRange(CFG.DATA_START_ROW, CFG.TK_START_COL, sampleRows, numCols).getBackgrounds();
  const colorSet = {};
  for (let r = 0; r < sampleRows; r++) {
    for (let c = 0; c < numCols; c++) {
      const bg = (bgs[r][c] || '').toLowerCase();
      if (bg && bg !== '#ffffff' && bg !== 'white' && bg !== '') {
        colorSet[bg] = (colorSet[bg] || 0) + 1;
      }
    }
  }
  Logger.log('Уникальные цвета в R..GN (первые 50 строк): ' + JSON.stringify(colorSet));
  Logger.log('Ожидаемый зелёный: ' + JSON.stringify(CFG.GREEN_HEXES));

  // 4. Строки, которые проходят все фильтры (покажем первые 10)
  const durations = sh.getRange(CFG.DATA_START_ROW, CFG.DURATION_COL, numRows, 1).getValues();
  const statuses  = sh.getRange(CFG.DATA_START_ROW, CFG.STATUS_COL,   numRows, 1).getDisplayValues();
  const starts    = sh.getRange(CFG.DATA_START_ROW, CFG.START_COL,    numRows, 1).getValues();
  const ends      = sh.getRange(CFG.DATA_START_ROW, CFG.END_COL,      numRows, 1).getValues();
  const allBgs    = sh.getRange(CFG.DATA_START_ROW, CFG.TK_START_COL, numRows, numCols).getBackgrounds();

  let passDate = 0, passStatus = 0, passGreen = 0, shown = 0;

  for (let r = 0; r < numRows; r++) {
    const startDate = toDate00_(starts[r][0]);
    const endDate   = toDate00_(ends[r][0]);
    if (!startDate || !endDate) continue;

    if (selectedDate < startDate || selectedDate > endDate) continue;
    passDate++;

    const effStatus = effectiveStatusOnDate_(statuses[r][0], startDate, endDate, selectedDate);
    if (effStatus === CFG.STATUS_DONE) continue;
    passStatus++;

    // Есть ли хоть одна зелёная ячейка в строке?
    const hasGreen = allBgs[r].some(bg => isGreen_(bg));
    if (hasGreen) {
      passGreen++;
      if (shown < 10) {
        Logger.log(
          '✅ Строка ' + (CFG.DATA_START_ROW + r) +
          ' | статус: "' + statuses[r][0] + '"' +
          ' | эфф.статус: "' + effStatus + '"' +
          ' | секунды: ' + durations[r][0] +
          ' | F=' + starts[r][0] + ' G=' + ends[r][0]
        );
        shown++;
      }
    }
  }

  Logger.log('─────────────────────────────');
  Logger.log('Строк прошли фильтр по дате:   ' + passDate);
  Logger.log('Строк прошли фильтр по статусу: ' + passStatus);
  Logger.log('Строк с зелёными ячейками:      ' + passGreen);

  if (passGreen === 0) {
    if (passDate === 0) {
      Logger.log('❌ Никакая строка не попадает в диапазон дат N7=' + selectedDate);
    } else if (passStatus === 0) {
      Logger.log('❌ Все строки отфильтрованы по статусу — проверьте значения столбца E');
      Logger.log('   Примеры статусов из таблицы:');
      const sample = new Set();
      for (let r = 0; r < Math.min(numRows, 100); r++) {
        sample.add('"' + String(statuses[r][0]).trim() + '"');
        if (sample.size >= 10) break;
      }
      Logger.log('   ' + Array.from(sample).join(', '));
    } else {
      Logger.log('❌ Нет зелёных ячеек — или цвет не совпадает с GREEN_HEXES');
      Logger.log('   Добавьте в CFG.GREEN_HEXES нужный цвет из списка выше');
    }
  }
}
