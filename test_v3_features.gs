/**
 * 見積→注文のAI紐付けテスト
 */
function testAiLinkQuote() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('案件管理');
  const data = sheet.getDataRange().getValues();
  
  // 見積書のIDを適当に取得
  let testQuoteId = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).startsWith('MQ')) {
      testQuoteId = data[i][0];
      break;
    }
  }
  
  if (!testQuoteId) {
    Logger.log('テスト用の見積データが見つかりません');
    return;
  }
  
  Logger.log('Testing AI Link for Quote ID: ' + testQuoteId);
  const result = aiLinkQuoteToOrder(testQuoteId);
  Logger.log('Result: ' + JSON.stringify(result, null, 2));
}

/**
 * チャット通知のテスト
 */
function testChatNotification() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('案件管理');
  const data = sheet.getDataRange().getValues();
  
  if (data.length > 1) {
    const testId = data[1][0];
    Logger.log('Sending test notification for ID: ' + testId);
    _sendChatNotification(testId, testId.startsWith('MQ') ? 'quote' : 'order');
  }
}
function listAvailableModels() {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  var url = 'https://generativelanguage.googleapis.com/v1beta/models?key=' + key;
  var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var data = JSON.parse(res.getContentText());
  
  if (data.models) {
    data.models.forEach(function(m) {
      // generateContentに対応しているモデルのみ表示
      if (m.supportedGenerationMethods && 
          m.supportedGenerationMethods.indexOf('generateContent') >= 0) {
        Logger.log(m.name + ' | ' + (m.description || ''));
      }
    });
  } else {
    Logger.log(res.getContentText());
  }
}