// ============================================================
// 見積書・注文書管理システム
// ファイル 10: アラート・監視（統合版）
// ============================================================
//
// 【統合元】
//   10_alert_cron.gs   — 旧シンプルアラート（チェックと送信のみ）
//   14_alert_monitor.gs — システム全体監視（OCR失敗、未紐づけ、停滞、納期）
//
// 【メイン実行関数】
//   runAllMonitoring()      — 毎日9時に実行（全監視）
//   checkTriggerHealth()    — 毎時実行（トリガー停止検知）
//   checkAndSendAlerts()    — 旧来の簡易アラート（互換用）
// ============================================================

// ============================================================
// 設定
// ============================================================

var MONITOR_CONFIG = {
  OCR_FAIL_THRESHOLD    : 3,
  UNLINKED_ORDER_DAYS   : 2,
  STAGNANT_WARN_DAYS    : 7,
  STAGNANT_ALERT_DAYS   : 14,
  DELIVERY_REMIND_DAYS  : 7,
  DELIVERY_URGENT_DAYS  : 1,
  TRIGGER_HEALTH_KEY    : 'TRIGGER_LAST_RUN',
  TRIGGER_TIMEOUT_HOURS : 2,
};

// SYS_SETTINGS からユーザー設定を読み込んで MONITOR_CONFIG に反映
// 各監視関数の先頭で呼ぶことで最新設定を使用する
function _loadMonitorConfig() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('SYS_SETTINGS');
    if (!raw) return;
    var s = JSON.parse(raw);
    if (s.deliveryRemindDays) MONITOR_CONFIG.DELIVERY_REMIND_DAYS  = Number(s.deliveryRemindDays);
    if (s.deliveryUrgentDays) MONITOR_CONFIG.DELIVERY_URGENT_DAYS  = Number(s.deliveryUrgentDays);
    if (s.stagnantWarnDays)   MONITOR_CONFIG.STAGNANT_WARN_DAYS    = Number(s.stagnantWarnDays);
    if (s.stagnantAlertDays)  MONITOR_CONFIG.STAGNANT_ALERT_DAYS   = Number(s.stagnantAlertDays);
    if (s.unlinkedOrderDays)  MONITOR_CONFIG.UNLINKED_ORDER_DAYS   = Number(s.unlinkedOrderDays);
    if (s.ocrFailThreshold)   MONITOR_CONFIG.OCR_FAIL_THRESHOLD    = Number(s.ocrFailThreshold);
    // アラート有効フラグ
    MONITOR_CONFIG.ALERT_DELIVERY = s.alertDelivery !== false;
    MONITOR_CONFIG.ALERT_DEADLINE = s.alertDeadline !== false;
    MONITOR_CONFIG.ALERT_OVERDUE  = s.alertOverdue  !== false;
    MONITOR_CONFIG.ALERT_UNLINKED = s.alertUnlinked !== false;
  } catch(e) {
    Logger.log('[_loadMonitorConfig ERROR] ' + e.message);
  }
}

// Gmailへの通知送信（Webhookに加えてメールでも通知）
function _sendAlertEmail(message) {
  try {
    var notifyEmails = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAILS') || '';
    if (!notifyEmails) return;
    var emails = notifyEmails.split(',').map(function(e){ return e.trim(); }).filter(Boolean);
    var subject = '【見積管理システム】アラート通知 ' +
                  Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    emails.forEach(function(email) {
      try {
        GmailApp.sendEmail(email, subject, message);
      } catch(e2) {
        Logger.log('[_sendAlertEmail] 送信失敗: ' + email + ' / ' + e2.message);
      }
    });
    Logger.log('[_sendAlertEmail] 送信完了: ' + emails.length + '件');
  } catch(e) {
    Logger.log('[_sendAlertEmail ERROR] ' + e.message);
  }
}

// ============================================================
// メインエントリーポイント（毎日9時に実行）
// ============================================================
function runAllMonitoring() {
  Logger.log('[MONITOR] 監視開始: ' + nowJST());
  _loadMonitorConfig(); // SYS_SETTINGSからユーザー設定を反映

  var alerts = [];
  alerts = alerts.concat(_checkOcrFailures());
  if (MONITOR_CONFIG.ALERT_UNLINKED !== false) alerts = alerts.concat(_checkUnlinkedOrders());
  alerts = alerts.concat(_checkStagnantCases());
  if (MONITOR_CONFIG.ALERT_DELIVERY !== false) alerts = alerts.concat(_checkDeliveryDates());
  if (MONITOR_CONFIG.ALERT_DEADLINE !== false) alerts = alerts.concat(checkOrderDeadlines());
  alerts = alerts.concat(_checkUnprocessedDriveFiles());
  _updateTriggerHealth();

  if (alerts.length > 0) {
    var msg = '【📊 見積管理システム 自動監視レポート】\n' +
              Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + '\n\n' +
              alerts.join('\n\n');
    _sendMonitorAlert(msg);  // Chat Webhook
    _sendAlertEmail(msg);    // メール通知
    Logger.log('[MONITOR] アラート送信: ' + alerts.length + '件');
  } else {
    Logger.log('[MONITOR] 異常なし');
  }

  return { alertCount: alerts.length, alerts: alerts };
}

// ============================================================
// ① OCR失敗検知
// ============================================================
function _checkOcrFailures() {
  var alerts = [];
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName('OCR処理ログ');
    if (!sheet || sheet.getLastRow() <= 1) return alerts;

    var data      = sheet.getRange(2, 1, sheet.getLastRow()-1, 5).getValues();
    var since     = new Date(new Date().getTime() - 24*60*60*1000);
    var failCount = 0;
    var failFiles = [];

    data.forEach(function(row) {
      var sts = String(row[2] || '');
      try {
        var d = new Date(String(row[0]||'').replace(/\//g,'-'));
        if (d >= since && (sts === 'ocr_failed' || sts === 'error')) {
          failCount++;
          failFiles.push(String(row[1]||'').substring(0, 40));
        }
      } catch(e) {}
    });

    if (failCount >= MONITOR_CONFIG.OCR_FAIL_THRESHOLD) {
      alerts.push(
        '⚠️ *OCR失敗が多発しています*\n' +
        '直近24時間の失敗件数: ' + failCount + '件\n' +
        'ファイル例: ' + failFiles.slice(0,3).join(', ')
      );
    }
  } catch(e) {
    Logger.log('[MONITOR] _checkOcrFailures: ' + e.message);
  }
  return alerts;
}

// ============================================================
// ② 未紐づけ注文書の検知
// ============================================================
function _checkUnlinkedOrders() {
  var alerts = [];
  try {
    var data     = getAllMgmtData();
    var today    = new Date();
    var unlinked = [];

    data.forEach(function(row) {
      var orderNo   = String(row[MGMT_COLS.ORDER_NO - 1] || '').trim();
      var linked    = _isLinkedVal(row[MGMT_COLS.LINKED - 1]);
      var quoteNo   = String(row[MGMT_COLS.QUOTE_NO - 1] || '').trim();
      var orderDate = row[MGMT_COLS.ORDER_DATE - 1];
      var client    = String(row[MGMT_COLS.CLIENT - 1] || '');

      if (!orderNo || linked || quoteNo) return;

      var days = 0;
      try {
        var d = orderDate instanceof Date ? orderDate : new Date(String(orderDate).replace(/\//g,'-'));
        if (!isNaN(d.getTime())) days = Math.floor((today - d) / (1000*60*60*24));
      } catch(e) {}

      if (days >= MONITOR_CONFIG.UNLINKED_ORDER_DAYS) {
        unlinked.push(client + '（' + orderNo + '、' + days + '日経過）');
      }
    });

    if (unlinked.length > 0) {
      alerts.push(
        '🔗 *見積書未紐づけの注文書があります*\n' +
        unlinked.slice(0,5).join('\n') +
        (unlinked.length > 5 ? '\n...他' + (unlinked.length-5) + '件' : '')
      );
    }
  } catch(e) {
    Logger.log('[MONITOR] _checkUnlinkedOrders: ' + e.message);
  }
  return alerts;
}

// ============================================================
// ③ 案件の停滞アラート
// ============================================================
function _checkStagnantCases() {
  var alerts = [];
  try {
    var data  = getAllMgmtData();
    var today = new Date();
    var warns = [];
    var crits = [];

    var ignoreStatuses = [
      CONFIG.STATUS.ORDERED, CONFIG.STATUS.DELIVERED,
      CONFIG.STATUS.CANCELLED, 'キャンセル', '失注',
    ];

    data.forEach(function(row) {
      var status  = String(row[MGMT_COLS.STATUS    - 1] || '');
      var quoteNo = String(row[MGMT_COLS.QUOTE_NO  - 1] || '').trim();
      var client  = String(row[MGMT_COLS.CLIENT    - 1] || '');
      var updated = row[MGMT_COLS.UPDATED_AT - 1];

      if (!quoteNo) return;
      if (ignoreStatuses.indexOf(status) >= 0) return;

      var days = 0;
      try {
        var d = updated instanceof Date ? updated : new Date(String(updated).replace(/\//g,'-'));
        if (!isNaN(d.getTime())) days = Math.floor((today - d) / (1000*60*60*24));
      } catch(e) {}

      if (days >= MONITOR_CONFIG.STAGNANT_ALERT_DAYS) {
        crits.push(client + '「' + quoteNo + '」' + status + '（' + days + '日更新なし）');
      } else if (days >= MONITOR_CONFIG.STAGNANT_WARN_DAYS) {
        warns.push(client + '「' + quoteNo + '」（' + days + '日更新なし）');
      }
    });

    if (crits.length > 0) alerts.push('🚨 *停滞警告（' + MONITOR_CONFIG.STAGNANT_ALERT_DAYS + '日以上）*\n' + crits.slice(0,5).join('\n'));
    if (warns.length > 0) alerts.push('⚠️ *停滞注意（' + MONITOR_CONFIG.STAGNANT_WARN_DAYS + '日以上）*\n' + warns.slice(0,5).join('\n'));
  } catch(e) {
    Logger.log('[MONITOR] _checkStagnantCases: ' + e.message);
  }
  return alerts;
}

// ============================================================
// ④ 納期アラート
// ============================================================
function _checkDeliveryDates() {
  var alerts  = [];
  try {
    var data    = getAllMgmtData();
    var today   = new Date();
    var urgents = [];
    var reminds = [];
    var overdue = [];

    data.forEach(function(row) {
      var status       = String(row[MGMT_COLS.STATUS       - 1] || '');
      var deliveryDate = row[MGMT_COLS.DELIVERY_DATE - 1];
      var client       = String(row[MGMT_COLS.CLIENT  - 1] || '');
      var orderNo      = String(row[MGMT_COLS.ORDER_NO - 1] || '');

      if (!deliveryDate || !orderNo) return;
      if (status === CONFIG.STATUS.DELIVERED || status === CONFIG.STATUS.CANCELLED) return;

      var days = 9999;
      try {
        var d = deliveryDate instanceof Date ? deliveryDate : new Date(String(deliveryDate).replace(/\//g,'-'));
        if (!isNaN(d.getTime())) days = Math.floor((d - today) / (1000*60*60*24));
      } catch(e) {}

      var label = client + '「' + orderNo + '」';
      if (days < 0) {
        overdue.push(label + ' (' + Math.abs(days) + '日超過)');
      } else if (days <= MONITOR_CONFIG.DELIVERY_URGENT_DAYS) {
        urgents.push(label + ' (明日納期)');
      } else if (days <= MONITOR_CONFIG.DELIVERY_REMIND_DAYS) {
        reminds.push(label + ' (あと' + days + '日)');
      }
    });

    if (overdue.length  > 0) alerts.push('🧨 *納期超過*\n'       + overdue.slice(0,5).join('\n'));
    if (urgents.length  > 0) alerts.push('🔥 *明日納期*\n'       + urgents.slice(0,5).join('\n'));
    if (reminds.length  > 0) alerts.push('📅 *納期リマインド*\n' + reminds.slice(0,5).join('\n'));
  } catch(e) {
    Logger.log('[MONITOR] _checkDeliveryDates: ' + e.message);
  }
  return alerts;
}

// ============================================================
// ① 注文書の発注期限アラート（ORDER_DEADLINE = col 29）
//    DEADLINE_NOTIFIED（col 32）で重複通知を防止
//    runAllMonitoring() から呼び出されるほか、
//    setupDeadlineAlertTrigger() で6時間ごとにも実行される
// ============================================================
function checkOrderDeadlines() {
  var alerts = [];
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var last  = sheet.getLastRow();
    if (last <= 1) return alerts;

    // col 32 まで読む（DEADLINE_NOTIFIED が col 32）
    var readCols = Math.max(sheet.getLastColumn(), 32);
    var data     = sheet.getRange(2, 1, last - 1, readCols).getValues();
    var today    = new Date(); today.setHours(0, 0, 0, 0);
    var urgents  = [], reminds = [], overdue = [];

    data.forEach(function(row, i) {
      var id        = String(row[MGMT_COLS.ID             - 1] || '');
      var orderNo   = String(row[MGMT_COLS.ORDER_NO       - 1] || '').trim();
      var status    = String(row[MGMT_COLS.STATUS         - 1] || '');
      var deadline  = row[MGMT_COLS.ORDER_DEADLINE        - 1]; // col 29
      var notified  = String(row[MGMT_COLS.DEADLINE_NOTIFIED - 1] || ''); // col 32
      var client    = String(row[MGMT_COLS.CLIENT         - 1] || '');
      var slipNo    = String(row[MGMT_COLS.ORDER_SLIP_NO  - 1] || '');

      if (!id || !deadline) return;
      if (status === CONFIG.STATUS.DELIVERED || status === CONFIG.STATUS.CANCELLED) return;

      var deadlineDate;
      try {
        deadlineDate = deadline instanceof Date
          ? deadline
          : new Date(String(deadline).replace(/\//g, '-'));
        if (isNaN(deadlineDate.getTime())) return;
        deadlineDate.setHours(0, 0, 0, 0);
      } catch(e) { return; }

      var diffDays = Math.round((deadlineDate - today) / (1000 * 60 * 60 * 24));
      var label    = client
        + (orderNo ? '「' + orderNo + '」' : '')
        + (slipNo  ? '（伝票:' + slipNo + '）' : '')
        + ' 期限: ' + _toDateStr(deadlineDate);

      var todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
      var rowRef   = { rowNum: i + 2 };

      if (diffDays < 0) {
        // 超過: 毎回通知（通知済みフラグのリセットはしない）
        overdue.push(label + '（' + Math.abs(diffDays) + '日超過）');
      } else if (diffDays <= 1) {
        if (!_hasNotifiedToday(notified, 'urgent', todayStr)) {
          urgents.push(label + '（明日が期限）');
          rowRef.level = 'urgent';
          _writeDeadlineNotified(sheet, i + 2, 'urgent:' + todayStr);
        }
      } else if (diffDays <= 3) {
        if (!_hasNotifiedToday(notified, '3d', todayStr)) {
          reminds.push(label + '（あと' + diffDays + '日）');
          rowRef.level = '3d';
          _writeDeadlineNotified(sheet, i + 2, '3d:' + todayStr);
        }
      } else if (diffDays <= 7) {
        if (!_hasNotifiedToday(notified, '7d', todayStr)) {
          reminds.push(label + '（あと' + diffDays + '日）');
          rowRef.level = '7d';
          _writeDeadlineNotified(sheet, i + 2, '7d:' + todayStr);
        }
      }
    });

    if (overdue.length  > 0) alerts.push('🧨 *[発注期限超過]*\n'   + overdue.join('\n'));
    if (urgents.length  > 0) alerts.push('🔥 *[発注期限 明日]*\n'  + urgents.join('\n'));
    if (reminds.length  > 0) alerts.push('📅 *[発注期限 事前通知]*\n' + reminds.join('\n'));
  } catch(e) {
    Logger.log('[checkOrderDeadlines ERROR] ' + e.message);
  }
  return alerts;
}

function _hasNotifiedToday(notifiedStr, level, todayStr) {
  // "urgent:20260517" のような形式で判定
  return notifiedStr.indexOf(level + ':' + todayStr) >= 0;
}

function _writeDeadlineNotified(sheet, rowNum, value) {
  try {
    sheet.getRange(rowNum, MGMT_COLS.DEADLINE_NOTIFIED).setValue(value);
  } catch(e) {
    Logger.log('[_writeDeadlineNotified ERROR] ' + e.message);
  }
}

// ============================================================
// ⑤ 未処理Driveファイルの滞留検知
// ============================================================
function _checkUnprocessedDriveFiles() {
  var alerts = [];
  try {
    var folders = [
      { id: CONFIG.IMPORT_QUOTE_FOLDER_ID,       label: '見積書インポート' },
      { id: CONFIG.IMPORT_ORDER_TRIAL_FOLDER_ID, label: '注文書（試作）インポート' },
      { id: CONFIG.IMPORT_ORDER_MASS_FOLDER_ID,  label: '注文書（量産）インポート' },
    ];
    var threshold = new Date(new Date().getTime() - 4*60*60*1000);

    folders.forEach(function(f) {
      if (!f.id) return;
      try {
        var folder   = DriveApp.getFolderById(f.id);
        var files    = folder.getFiles();
        var oldFiles = [];
        while (files.hasNext()) {
          var file = files.next();
          if (file.getDateCreated() < threshold) {
            oldFiles.push(file.getName().substring(0, 40));
          }
        }
        if (oldFiles.length > 0) {
          alerts.push(
            '📂 *未処理ファイルが滞留しています*\n' +
            'フォルダ: ' + f.label + '\n' +
            'ファイル数: ' + oldFiles.length + '件\n' +
            '例: ' + oldFiles.slice(0,2).join(', ')
          );
        }
      } catch(e) {
        Logger.log('[MONITOR] folder check error: ' + f.label + ' ' + e.message);
      }
    });
  } catch(e) {
    Logger.log('[MONITOR] _checkUnprocessedDriveFiles: ' + e.message);
  }
  return alerts;
}

// ============================================================
// トリガー稼働監視
// ============================================================

function _updateTriggerHealth() {
  try {
    PropertiesService.getScriptProperties()
      .setProperty(MONITOR_CONFIG.TRIGGER_HEALTH_KEY, new Date().getTime().toString());
  } catch(e) {}
}

function checkTriggerHealth() {
  try {
    _updateTriggerHealth();

    var triggers          = ScriptApp.getProjectTriggers();
    var criticalFunctions = ['processNewEmails', 'processDriveImports', 'autoMatchNewOrders'];
    var missingTriggers   = [];

    criticalFunctions.forEach(function(fn) {
      var found = triggers.some(function(t){ return t.getHandlerFunction() === fn; });
      if (!found) missingTriggers.push(fn);
    });

    if (missingTriggers.length > 0) {
      _sendMonitorAlert(
        '🔴 *トリガーが停止しています*\n' +
        '停止中の関数:\n' + missingTriggers.map(function(f){ return '  - ' + f; }).join('\n') + '\n\n' +
        '対処: GASエディタで _registerTriggers() を再実行してください'
      );
    }
  } catch(e) {
    Logger.log('[MONITOR] checkTriggerHealth: ' + e.message);
  }
}

// ============================================================
// 通知送信
// ============================================================
function _sendMonitorAlert(message) {
  var webhookUrl = CONFIG.GOOGLE_CHAT_WEBHOOK_URL ||
    PropertiesService.getScriptProperties().getProperty('GOOGLE_CHAT_WEBHOOK_URL') || '';

  if (!webhookUrl) {
    Logger.log('[MONITOR] Webhook未設定。メッセージ:\n' + message);
    return;
  }

  try {
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ text: message }), muteHttpExceptions: true,
    });
    Logger.log('[MONITOR] Chat通知送信完了');
  } catch(e) {
    Logger.log('[MONITOR] Chat通知失敗: ' + e.message);
    try {
      GmailApp.sendEmail(
        Session.getActiveUser().getEmail(),
        '【見積管理システム監視アラート】',
        message
      );
    } catch(e2) {}
  }
}

// ============================================================
// トリガー登録
// ============================================================
function setupMonitoringTriggers() {
  var existing = ScriptApp.getProjectTriggers().map(function(t){ return t.getHandlerFunction(); });

  if (existing.indexOf('runAllMonitoring') < 0) {
    ScriptApp.newTrigger('runAllMonitoring').timeBased().atHour(9).everyDays(1).create();
    Logger.log('[MONITOR] runAllMonitoring トリガー登録（毎日9時）');
  }
  if (existing.indexOf('checkTriggerHealth') < 0) {
    ScriptApp.newTrigger('checkTriggerHealth').timeBased().everyHours(1).create();
    Logger.log('[MONITOR] checkTriggerHealth トリガー登録（毎時）');
  }
  // ① 発注期限アラートは6時間ごとに独立実行
  if (existing.indexOf('runDeadlineAlerts') < 0) {
    ScriptApp.newTrigger('runDeadlineAlerts').timeBased().everyHours(6).create();
    Logger.log('[MONITOR] runDeadlineAlerts トリガー登録（6時間ごと）');
  }
}

// 発注期限アラートのスタンドアロン実行（トリガーから呼ぶ）
function runDeadlineAlerts() {
  Logger.log('[DEADLINE] 発注期限チェック開始: ' + nowJST());
  _loadMonitorConfig(); // SYS_SETTINGSからユーザー設定を反映

  if (MONITOR_CONFIG.ALERT_DEADLINE === false) {
    Logger.log('[DEADLINE] アラート無効（設定で無効化されています）');
    return;
  }

  var alerts = [];
  if (MONITOR_CONFIG.ALERT_DELIVERY !== false) alerts = alerts.concat(_checkDeliveryDates());
  if (MONITOR_CONFIG.ALERT_DEADLINE !== false) alerts = alerts.concat(checkOrderDeadlines());

  if (alerts.length > 0) {
    var msg = '【📦 納期・発注期限アラート】\n' +
              Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + '\n\n' +
              alerts.join('\n\n');
    _sendMonitorAlert(msg);  // Chat Webhook
    _sendAlertEmail(msg);    // メール通知
    Logger.log('[DEADLINE] アラート送信: ' + alerts.length + '件');
  } else {
    Logger.log('[DEADLINE] 発注期限・納期超過なし');
  }
}

// ============================================================
// テスト
// ============================================================
function testMonitoring() {
  Logger.log('=== 監視テスト開始 ===');
  var result = runAllMonitoring();
  Logger.log('アラート件数: ' + result.alertCount);
  result.alerts.forEach(function(a){ Logger.log('---\n' + a); });
  Logger.log('=== 監視テスト完了 ===');
}

function testSendAlert() {
  _sendMonitorAlert('✅ テスト通知\n見積管理システムの監視が正常に動作しています。\n送信時刻: ' + nowJST());
}

// ============================================================
// 旧来の簡易アラート（10_alert_cron.gs 互換）
// シート名「見積提出管理」を使う場合に利用
// ============================================================
function checkAndSendAlerts() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT || '見積提出管理');
  if (!sheet) return;

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];

  var isLatestIdx  = headers.indexOf('最新フラグ');
  var statusIdx    = headers.indexOf('ステータス');
  var dateIdx      = headers.indexOf('提出日');
  var numIdx       = headers.indexOf('見積No');
  var boardIdx     = headers.indexOf('基板名');
  var deliveryIdx  = headers.indexOf('納期');
  var today        = new Date();
  var alertMessages = [];

  for (var i = 1; i < data.length; i++) {
    var row      = data[i];
    var isLatest = isLatestIdx > -1 ? row[isLatestIdx] : true;
    if (!isLatest) continue;

    var status      = statusIdx > -1 ? row[statusIdx] : '';
    var quoteNumber = numIdx > -1 ? row[numIdx] : '不明';
    var subject     = boardIdx > -1 ? row[boardIdx] : '';

    if (status !== '発注済' && status !== '納品済み' && status !== 'キャンセル') {
      var quoteDateStr = dateIdx > -1 ? row[dateIdx] : null;
      if (quoteDateStr) {
        var quoteDate = new Date(quoteDateStr);
        var diffDays  = Math.floor((today - quoteDate) / (1000 * 60 * 60 * 24));
        if (diffDays >= 7 && diffDays < 14) {
          alertMessages.push('⚠️ 未着: 見積[' + quoteNumber + '] ' + subject + ' は発行後 ' + diffDays + '日 経過しています。');
        } else if (diffDays >= 14) {
          alertMessages.push('🚨 停滞警告: 見積[' + quoteNumber + '] ' + subject + ' は ' + diffDays + '日 放置されています！至急確認を。');
        }
      }
    }

    if (status === '発注済' || status === '作成中') {
      var deliveryDateStr = deliveryIdx > -1 ? row[deliveryIdx] : null;
      if (deliveryDateStr) {
        var deliveryDate   = new Date(deliveryDateStr);
        var daysToDelivery = Math.floor((deliveryDate - today) / (1000 * 60 * 60 * 24));
        if (daysToDelivery === 7) {
          alertMessages.push('📅 納期リマインド: 注文[' + quoteNumber + '] ' + subject + ' の納期まであと7日です。');
        } else if (daysToDelivery === 1) {
          alertMessages.push('🔥 明日納期: 注文[' + quoteNumber + '] ' + subject + ' の納期が迫っています。');
        } else if (daysToDelivery < 0) {
          alertMessages.push('🧨 納期超過: 注文[' + quoteNumber + '] ' + subject + ' は納期を ' + Math.abs(daysToDelivery) + '日 超過しています！');
        }
      }
    }
  }

  if (alertMessages.length > 0) {
    _sendMonitorAlert('【システム自動通知・運用アラート】\n' + alertMessages.join('\n'));
  }
}
