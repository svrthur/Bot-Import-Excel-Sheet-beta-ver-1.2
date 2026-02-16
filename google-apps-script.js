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
 * Функция для проверки превышения порога в 160 секунд.
 * Запускается триггером или вручную.
 */
function checkTKThreshold() {
  var spreadsheetId = "17VeQQWTGotofrpNbUHDhUFhCc3qjLdwoesTxDDfJ7h4";
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheets()[0];
  
  // Диапазон R2:GN2 (Строка 2, Колонки 18-195)
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
  
  if (exceeded.length > 0) {
    return exceeded;
  }
  return null;
}

/**
 * Функция для обработки запросов от бота для проверки порогов
 */
function doGet(e) {
  var exceeded = checkTKThreshold();
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    exceeded: exceeded
  })).setMimeType(ContentService.MimeType.JSON);
}
