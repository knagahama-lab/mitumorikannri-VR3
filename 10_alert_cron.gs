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
  
  let isLatestIdx = headers.indexOf('最新フラグ');
  let statusIdx = headers.indexOf('ステータス');
  let dateIdx = headers.indexOf('提出日');
  let numIdx = headers.indexOf('見積No');
  let boardIdx = headers.indexOf('基板名');
  let deliveryIdx = headers.indexOf('納期'); // 新規追加
  
  const today = new Date();
  let alertMessages = [];
  
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let isLatest = isLatestIdx > -1 ? row[isLatestIdx] : true;
    if (!isLatest) continue; // 最新版以外はスキップ
    
    let status = statusIdx > -1 ? row[statusIdx] : '';
    let quoteNumber = numIdx > -1 ? row[numIdx] : '不明';
    let subject = boardIdx > -1 ? row[boardIdx] : '';
    
    // 1. 注文未着アラート / 停滞アラート
    if (status !== '発注済' && status !== '納品済み' && status !== 'キャンセル') {
      let quoteDateStr = dateIdx > -1 ? row[dateIdx] : null;
      if (quoteDateStr) {
        let quoteDate = new Date(quoteDateStr);
        let diffDays = Math.floor((today - quoteDate) / (1000 * 60 * 60 * 24));
        if (diffDays >= 7 && diffDays < 14) {
          alertMessages.push(`⚠️ 未着: 見積[${quoteNumber}] ${subject} は発行後 ${diffDays}日 経過しています。`);
        } else if (diffDays >= 14) {
          alertMessages.push(`🚨 停滞警告: 見積[${quoteNumber}] ${subject} は ${diffDays}日 放置されています！至急確認を。`);
        }
      }
    }

    // 2. 納期リマインド・遅延アラート
    if (status === '発注済' || status === '作成中') {
      let deliveryDateStr = deliveryIdx > -1 ? row[deliveryIdx] : null;
      if (deliveryDateStr) {
        let deliveryDate = new Date(deliveryDateStr);
        let daysToDelivery = Math.floor((deliveryDate - today) / (1000 * 60 * 60 * 24));
        
        if (daysToDelivery === 7) {
          alertMessages.push(`📅 納期リマインド: 注文[${quoteNumber}] ${subject} の納期まであと7日です。`);
        } else if (daysToDelivery === 1) {
          alertMessages.push(`🔥 明日納期: 注文[${quoteNumber}] ${subject} の納期が迫っています。`);
        } else if (daysToDelivery < 0) {
          alertMessages.push(`🧨 納期超過: 注文[${quoteNumber}] ${subject} は納期を ${Math.abs(daysToDelivery)}日 超過しています！`);
        }
      }
    }
  }
  
  if (alertMessages.length > 0) {
    sendGoogleChatMessage("【システム自動通知・運用アラート】\n" + alertMessages.join('\n'));
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
