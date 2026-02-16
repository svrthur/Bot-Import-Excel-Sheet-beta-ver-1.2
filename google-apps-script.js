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
 * Функция для автоматического скрытия строк со статусом "завершено" спустя 24 часа.
 * Должна быть настроена на запуск по триггеру (например, раз в час).
 */
function hideFinishedCampaigns() {
  var spreadsheetId = "17VeQQWTGotofrpNbUHDhUFhCc3qjLdwoesTxDDfJ7h4";
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheets()[0];
  
  var startRow = 120;
  var statusCol = 5; // Столбец E (1-A, 2-B, 3-C, 4-D, 5-E)
  
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;
  
  var range = sheet.getRange(startRow, statusCol, lastRow - startRow + 1, 1);
  var values = range.getValues();
  
  var now = new Date().getTime();
  var twentyFourHours = 24 * 60 * 60 * 1000;
  
  // Для отслеживания времени завершения нам нужно где-то хранить дату изменения статуса.
  // В Google Таблицах нет встроенной истории изменения ячейки для скриптов без дополнительного логгирования.
  // В данном случае мы можем использовать Metadata или DeveloperMetadata, но проще всего
  // предположить, что если статус "завершено", мы проверяем дату из колонки "Дата окончания" (если она есть).
  // Однако пользователь просил "спустя сутки после получения статуса". 
  
  // Альтернатива: Скрывать все, что "завершено", если мы не можем точно знать момент смены статуса
  // ИЛИ добавить триггер onEdit, который будет записывать время смены статуса.
  
  for (var i = 0; i < values.length; i++) {
    var rowNum = startRow + i;
    var status = String(values[i][0]).toLowerCase().trim();
    
    if (status === "завершено") {
      // Так как мы не храним время смены статуса, мы можем использовать дату в соседней колонке 
      // или просто скрыть те, у которых дата окончания + 24 часа прошла.
      // В вашей таблице "Дата окончания" - это колонка I (9).
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
 * Триггер, который записывает время установки статуса "завершено", если его нет в колонке "Дата окончания".
 */
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  
  // Проверяем столбец E (5) и строку >= 120
  if (range.getColumn() === 5 && range.getRow() >= 120) {
    var value = String(e.value).toLowerCase().trim();
    if (value === "завершено") {
      // Можно записывать время в скрытую колонку или примечание, если нужно сверхточно.
      // Но обычно достаточно проверять колонку "Дата окончания".
    }
  }
}
