function doPost(e) {
  try {
    var requestData = JSON.parse(e.postData.contents);
    
    if (requestData.action === 'highlight') {
      highlightCampaigns(requestData.data);
      return ContentService.createTextOutput(JSON.stringify({status: 'success'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function highlightCampaigns(campaignData) {
  var spreadsheetId = "17VeQQWTGotofrpNbUHDhUFhCc3qjLdwoesTxDDfJ7h4";
  var ss = SpreadsheetApp.openById(spreadsheetId);
  
  if (!ss) {
    throw new Error("Не удалось получить доступ к таблице по ID: " + spreadsheetId);
  }

  var sheet = ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  
  var headers = data[0]; 
  var tkColMap = {};
  for (var col = 0; col < headers.length; col++) {
    var val = headers[col];
    if (val !== null && val !== undefined) {
      var tkValue = String(val).trim();
      if (tkValue) {
        tkColMap[tkValue] = col + 1;
      }
    }
  }

  for (var rowKey in campaignData) {
    var tks = campaignData[rowKey];
    var rowIndex = parseInt(rowKey.replace(/\D/g, ''));
    
    if (!isNaN(rowIndex) && rowIndex > 0) {
      for (var j = 0; j < tks.length; j++) {
        var tkNum = String(tks[j]).trim();
        var colIndex = tkColMap[tkNum];
        
        if (colIndex) {
          sheet.getRange(rowIndex, colIndex).setBackground("#00ff00");
        }
      }
    }
  }
}

/**
 * Подсчёт секунд по строке 2 для каждого ТК (R–GN).
 * Суммирует столбец C для строк, где:
 *   — ячейка ТК закрашена зелёным (#00ff00)
 *   — дата из N7 попадает в диапазон [Дата старта (F) .. Дата окончания (G)]
 *   — Статус (E) = "опубликовано" или "запланировано"
 * Записывает результат в строку 2 (R2:GN2).
 * Можно запускать вручную или по триггеру (например, каждый час).
 */
function calculateTKSeconds() {
  var spreadsheetId = "17VeQQWTGotofrpNbUHDhUFhCc3qjLdwoesTxDDfJ7h4";
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheets()[0];

  // Целевая дата из N7
  var targetDateRaw = sheet.getRange("N7").getValue();
  if (!(targetDateRaw instanceof Date) || isNaN(targetDateRaw.getTime())) {
    Logger.log("Ошибка: ячейка N7 не содержит корректную дату.");
    return { status: 'error', message: 'N7 не содержит дату' };
  }
  // Нормализуем до полуночи, чтобы сравнивать только по дням
  var target = new Date(targetDateRaw.getFullYear(), targetDateRaw.getMonth(), targetDateRaw.getDate());

  var lastRow = sheet.getLastRow();
  var startRow = 3; // Данные кампаний начинаются с 3-й строки

  if (lastRow < startRow) {
    Logger.log("Нет данных для обработки.");
    return { status: 'ok', message: 'Нет строк' };
  }

  var numRows = lastRow - startRow + 1;

  // Колонки ТК: R=18 .. GN=195 (178 колонок)
  var tkStartCol = 18;
  var numTKCols = 178;

  // Читаем все нужные данные одним блоком для скорости
  // Секунды: столбец C (3)
  var secondsData = sheet.getRange(startRow, 3, numRows, 1).getValues();

  // Статус (E=5), Дата старта (F=6), Дата окончания (G=7) — 3 столбца начиная с E
  var campaignData = sheet.getRange(startRow, 5, numRows, 3).getValues();

  // Цвета фона для всего блока ТК-колонок
  var backgrounds = sheet.getRange(startRow, tkStartCol, numRows, numTKCols).getBackgrounds();

  var GREEN = "#00ff00";
  // Допустимые статусы (в нижнем регистре)
  var validStatuses = ["опубликовано", "запланировано"];

  var totals = new Array(numTKCols).fill(0);

  for (var row = 0; row < numRows; row++) {
    var seconds = parseFloat(secondsData[row][0]);
    if (isNaN(seconds) || seconds <= 0) continue;

    var status = String(campaignData[row][0]).toLowerCase().trim();
    if (validStatuses.indexOf(status) === -1) continue;

    var rawStart = campaignData[row][1];
    var rawEnd   = campaignData[row][2];
    if (!(rawStart instanceof Date) || !(rawEnd instanceof Date)) continue;

    var startDate = new Date(rawStart.getFullYear(), rawStart.getMonth(), rawStart.getDate());
    var endDate   = new Date(rawEnd.getFullYear(),   rawEnd.getMonth(),   rawEnd.getDate());

    // Дата N7 должна попадать в [startDate, endDate]
    if (target < startDate || target > endDate) continue;

    // Проверяем каждую ТК-колонку на зелёный цвет
    for (var col = 0; col < numTKCols; col++) {
      if (backgrounds[row][col].toLowerCase() === GREEN) {
        totals[col] += seconds;
      }
    }
  }

  // Записываем результаты в строку 2 (R2:GN2)
  sheet.getRange(2, tkStartCol, 1, numTKCols).setValues([totals]);

  Logger.log("calculateTKSeconds завершён. Дата: " + target + " | Итого колонок: " + numTKCols);
  return { status: 'success', date: target.toString() };
}

function hideFinishedCampaigns() {
  var spreadsheetId = "17VeQQWTGotofrpNbUHDhUFhCc3qjLdwoesTxDDfJ7h4";
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheets()[0];
  
  var startRow = 120;
  var statusCol = 5; 
  
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;
  
  var range = sheet.getRange(startRow, statusCol, lastRow - startRow + 1, 1);
  var values = range.getValues();
  
  var now = new Date().getTime();
  var twentyFourHours = 24 * 60 * 60 * 1000;
  
  for (var i = 0; i < values.length; i++) {
    var rowNum = startRow + i;
    var status = String(values[i][0]).toLowerCase().trim();
    
    if (status === "завершено") {
      var endDateVal = sheet.getRange(rowNum, 9).getValue();
      if (endDateVal instanceof Date) {
        if (now - endDateVal.getTime() > twentyFourHours) {
          sheet.hideRows(rowNum);
        }
      }
    }
  }
}

/**
 * Проверка превышения порога 160 секунд в строке 2.
 */
function checkTKThreshold() {
  var spreadsheetId = "17VeQQWTGotofrpNbUHDhUFhCc3qjLdwoesTxDDfJ7h4";
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheets()[0];
  
  var range = sheet.getRange(2, 18, 1, 178);
  var values = range.getValues()[0];
  var headers = sheet.getRange(1, 18, 1, 178).getValues()[0];
  
  var exceeded = [];
  for (var i = 0; i < values.length; i++) {
    var seconds = parseFloat(values[i]);
    if (!isNaN(seconds) && seconds > 160) {
      exceeded.push(String(headers[i]).trim());
    }
  }
  
  return exceeded.length > 0 ? exceeded : null;
}

/**
 * GET-обработчик:
 *   ?action=calculate  — запускает подсчёт секунд и возвращает результат
 *   (по умолчанию)     — проверяет превышение порога 160 с
 */
function doGet(e) {
  var action = e && e.parameter && e.parameter.action;

  if (action === 'calculate') {
    var result = calculateTKSeconds();
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // По умолчанию — проверка порогов
  var exceeded = checkTKThreshold();
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    exceeded: exceeded
  })).setMimeType(ContentService.MimeType.JSON);
}
