// ============================================================
// 見積・注文 管理システム
// ファイル 10/10: 自動アラート・Cron
// ============================================================

const CHAT_WEBHOOK_URL = 'https://chat.googleapis.com/v1/spaces/XXXX/messages?key=YYYY&token=ZZZZ';

function checkAndSendAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 適宜、実際の管理シート名に変更してください（例: CONFIG.SHEET_MANAGEMENT）
  const sheet = ss.getSheetByName('見積提出管理');
  if(!sheet) return;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 過去仕様に合わせて「最新フラグ」等の列が存在する前提
  let isLatestIdx = headers.indexOf('最新フラグ');
  let statusIdx = headers.indexOf('ステータス');
  let dateIdx = headers.indexOf('提出日');
  let numIdx = headers.indexOf('見積No');
  let boardIdx = headers.indexOf('基板名');

  if (isLatestIdx === -1 || dateIdx === -1) {
    Logger.log('必要なカラムが見つからないためアラート機能はスキップしました。');
    return;
  }

  const today = new Date();
  let alertMessages = [];
  
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let isLatest = row[isLatestIdx];
    let status = statusIdx > -1 ? row[statusIdx] : '';
    let quoteDateStr = row[dateIdx];
    
    if (!quoteDateStr) continue;
    let quoteDate = new Date(quoteDateStr);

    // 最新フラグがTRUEで、かつ未発注のもの
    if (isLatest === true && status !== '発注済' && status !== 'キャンセル') {
      let diffDays = Math.floor((today - quoteDate) / (1000 * 60 * 60 * 24));
      
      if (diffDays >= 7) {
        let quoteNumber = row[numIdx] || '';
        let subject = row[boardIdx] || '';
        alertMessages.push(`⚠️ 見積[${quoteNumber}]: ${subject} は発行後 ${diffDays}日 経過していますが、注文未着（または未対応）です。`);
      }
    }
  }
  
  if (alertMessages.length > 0) {
    sendGoogleChatMessage("【注文書未着アラート・期限超過】\n" + alertMessages.join('\n'));
  }
}

function sendGoogleChatMessage(message) {
  if (CHAT_WEBHOOK_URL.indexOf('XXXX') !== -1) return; // Hookが未設定ならスキップ

  const payload = { 'text': message };
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload)
  };
  UrlFetchApp.fetch(CHAT_WEBHOOK_URL, options);
}

function setupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkAndSendAlerts') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 毎日朝9時〜10時の間に実行されるトリガーを登録
  ScriptApp.newTrigger('checkAndSendAlerts')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
}
