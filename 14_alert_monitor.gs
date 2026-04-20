// ============================================================
// 見積書・注文書管理システム
// ファイル 14: エラー監視・通知強化【Phase2 信頼性向上】
// ============================================================
//
// 【10_alert_cron.gs との違い】
//   旧: 見積提出管理シートの停滞チェックのみ
//   新: システム全体の健全性を監視
//       ① OCR失敗の検知と通知
//       ② マッチング失敗（未紐づけ注文書）の検知
//       ③ トリガーの停止検知
//       ④ 案件の納期アラート（既存機能を強化）
//       ⑤ 未処理ファイルの滞留検知
//
// 【Google Chat Webhook の設定方法】
//   スクリプトプロパティ「GOOGLE_CHAT_WEBHOOK_URL」に
//   Google Chat の Incoming Webhook URL を設定してください
//
// 【トリガー設定】
//   runAllMonitoring() を毎日9時に実行
//   checkTriggerHealth() を毎時実行（トリガー停止検知）
// ============================================================

var MONITOR_CONFIG = {
  // OCRログの失敗が何件以上で通知するか
  OCR_FAIL_THRESHOLD       : 3,
  // 未紐づけ注文書が何日以上で通知するか
  UNLINKED_ORDER_DAYS      : 2,
  // 案件の停滞アラート（日数）
  STAGNANT_WARN_DAYS       : 7,
  STAGNANT_ALERT_DAYS      : 14,
  // 納期アラート（残り日数）
  DELIVERY_REMIND_DAYS     : 7,
  DELIVERY_URGENT_DAYS     : 1,
  // トリガー監視キー
  TRIGGER_HEALTH_KEY       : 'TRIGGER_LAST_RUN',
  TRIGGER_TIMEOUT_HOURS    : 2,  // 2時間以上実行されなければ異常
};

// ============================================================
// メインエントリーポイント（毎日9時に実行）
// ============================================================
function runAllMonitoring() {
  Logger.log('[MONITOR] 監視開始: ' + nowJST());

  var alerts = [];

  // ① OCR失敗検知
  var ocrAlerts = _checkOcrFailures();
  alerts = alerts.concat(ocrAlerts);

  // ② 未紐づけ注文書
  var unlinkAlerts = _checkUnlinkedOrders();
  alerts = alerts.concat(unlinkAlerts);

  // ③ 案件の停滞アラート
  var stagnantAlerts = _checkStagnantCases();
  alerts = alerts.concat(stagnantAlerts);

  // ④ 納期アラート
  var deliveryAlerts = _checkDeliveryDates();
  alerts = alerts.concat(deliveryAlerts);

  // ⑤ 未処理Drive ファイル
  var driveAlerts = _checkUnprocessedDriveFiles();
  alerts = alerts.concat(driveAlerts);

  // ⑥ トリガー稼働記録を更新
  _updateTriggerHealth();

  // 通知送信
  if (alerts.length > 0) {
    var msg = '【📊 見積管理システム 自動監視レポート】\n' +
              Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + '\n\n' +
              alerts.join('\n\n');
    _sendMonitorAlert(msg);
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
    var today     = new Date();
    var since     = new Date(today.getTime() - 24*60*60*1000); // 直近24時間
    var failCount = 0;
    var failFiles = [];

    data.forEach(function(row) {
      var dateStr = String(row[0] || '');
      var status  = String(row[1] || '');  // ← ファイル名は2列目
      var sts     = String(row[2] || '');  // ← ステータスは3列目
      // 直近24時間の失敗
      try {
        var d = new Date(dateStr.replace(/\//g,'-'));
        if (d >= since && (sts === 'ocr_failed' || sts === 'error')) {
          failCount++;
          failFiles.push(String(row[1]||'').substring(0,40));
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

      if (!orderNo || linked || quoteNo) return; // 紐づけ済み or 見積番号あり はスキップ

      // 注文日からの経過日数
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
    var data   = getAllMgmtData();
    var today  = new Date();
    var warns  = [];
    var crits  = [];

    var ignoreStatuses = [
      CONFIG.STATUS.ORDERED, CONFIG.STATUS.DELIVERED,
      CONFIG.STATUS.CANCELLED, 'キャンセル', '失注',
    ];

    data.forEach(function(row) {
      var status   = String(row[MGMT_COLS.STATUS    - 1] || '');
      var quoteNo  = String(row[MGMT_COLS.QUOTE_NO  - 1] || '').trim();
      var client   = String(row[MGMT_COLS.CLIENT    - 1] || '');
      var updated  = row[MGMT_COLS.UPDATED_AT - 1];

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

    if (crits.length > 0) {
      alerts.push('🚨 *停滞警告（' + MONITOR_CONFIG.STAGNANT_ALERT_DAYS + '日以上）*\n' + crits.slice(0,5).join('\n'));
    }
    if (warns.length > 0) {
      alerts.push('⚠️ *停滞注意（' + MONITOR_CONFIG.STAGNANT_WARN_DAYS + '日以上）*\n' + warns.slice(0,5).join('\n'));
    }
  } catch(e) {
    Logger.log('[MONITOR] _checkStagnantCases: ' + e.message);
  }
  return alerts;
}

// ============================================================
// ④ 納期アラート
// ============================================================
function _checkDeliveryDates() {
  var alerts = [];
  try {
    var data  = getAllMgmtData();
    var today = new Date();
    var urgents  = [];
    var reminds  = [];
    var overdue  = [];

    data.forEach(function(row) {
      var status       = String(row[MGMT_COLS.STATUS       - 1] || '');
      var deliveryDate = row[MGMT_COLS.DELIVERY_DATE - 1];
      var client       = String(row[MGMT_COLS.CLIENT - 1] || '');
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

    if (overdue.length  > 0) alerts.push('🧨 *納期超過*\n'         + overdue.slice(0,5).join('\n'));
    if (urgents.length  > 0) alerts.push('🔥 *明日納期*\n'         + urgents.slice(0,5).join('\n'));
    if (reminds.length  > 0) alerts.push('📅 *納期リマインド*\n'   + reminds.slice(0,5).join('\n'));
  } catch(e) {
    Logger.log('[MONITOR] _checkDeliveryDates: ' + e.message);
  }
  return alerts;
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

    var threshold = new Date(new Date().getTime() - 4*60*60*1000); // 4時間以上前

    folders.forEach(function(f) {
      if (!f.id) return;
      try {
        var folder = DriveApp.getFolderById(f.id);
        var files  = folder.getFiles();
        var oldFiles = [];
        while (files.hasNext()) {
          var file = files.next();
          if (file.getDateCreated() < threshold) {
            oldFiles.push(file.getName().substring(0,40));
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
// トリガー稼働監視（毎時実行）
// ============================================================

/**
 * 定期実行されるごとに「最終実行時刻」を記録する
 * processNewEmails や processDriveImports に呼ばせる
 */
function _updateTriggerHealth() {
  try {
    PropertiesService.getScriptProperties()
      .setProperty(MONITOR_CONFIG.TRIGGER_HEALTH_KEY, new Date().getTime().toString());
  } catch(e) {}
}

/**
 * トリガーが止まっていないか確認（毎時実行推奨）
 */
function checkTriggerHealth() {
  try {
    _updateTriggerHealth(); // 自身の実行を記録

    // GASのトリガー一覧を確認
    var triggers = ScriptApp.getProjectTriggers();
    var criticalFunctions = ['processNewEmails', 'processDriveImports', 'autoMatchNewOrders'];
    var missingTriggers   = [];

    criticalFunctions.forEach(function(fn) {
      var found = triggers.some(function(t) { return t.getHandlerFunction() === fn; });
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
    Logger.log('[MONITOR] Webhook未設定のためスキップ。メッセージ:\n' + message);
    return;
  }

  try {
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ text: message }),
      muteHttpExceptions: true,
    });
    Logger.log('[MONITOR] Chat通知送信完了');
  } catch(e) {
    Logger.log('[MONITOR] Chat通知失敗: ' + e.message);
    // Chatが失敗してもメールにフォールバック
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

  // 毎日9時: 全監視実行
  if (existing.indexOf('runAllMonitoring') < 0) {
    ScriptApp.newTrigger('runAllMonitoring').timeBased().atHour(9).everyDays(1).create();
    Logger.log('[MONITOR] runAllMonitoring トリガー登録（毎日9時）');
  }
  // 毎時: トリガー稼働確認
  if (existing.indexOf('checkTriggerHealth') < 0) {
    ScriptApp.newTrigger('checkTriggerHealth').timeBased().everyHours(1).create();
    Logger.log('[MONITOR] checkTriggerHealth トリガー登録（毎時）');
  }
}

// ============================================================
// テスト用（手動実行で動作確認）
// ============================================================
function testMonitoring() {
  Logger.log('=== 監視テスト開始 ===');
  var result = runAllMonitoring();
  Logger.log('アラート件数: ' + result.alertCount);
  result.alerts.forEach(function(a) { Logger.log('---\n' + a); });
  Logger.log('=== 監視テスト完了 ===');
}

function testSendAlert() {
  _sendMonitorAlert(
    '✅ テスト通知\n' +
    '見積管理システムの監視が正常に動作しています。\n' +
    '送信時刻: ' + nowJST()
  );
}
