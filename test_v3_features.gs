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
/**
 * この関数を実行して、夜間用APIキーを強制的に保存する
 */
function forceSaveBatchKey() {
  // ★ ここを " " で囲むのを忘れずに！
  var key = "AIzaSyARY7ly2Mt6li6gcqR68GWAXHtsY0V"; 
  
  // 保存実行
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY_MATCHING', key);
  
  // 確認
  var savedKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY_MATCHING');
  if (savedKey === key) {
    Logger.log("成功！プロパティが正常に保存されました: " + savedKey.substring(0, 5) + "...");
  } else {
    Logger.log("エラー：保存に失敗しました。");
  }
}