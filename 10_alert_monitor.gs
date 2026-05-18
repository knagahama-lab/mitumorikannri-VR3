// ============================================================
// 見積書・注文書管理システム
// ファイル 10: アラート・監視（統合版）
// ============================================================
//
// 【変更点】
//   ・メール通知をHTMLカード形式に全面改修（URLを埋め込みリンクに）
//   ・各アラートはテキスト文字列ではなく構造化オブジェクト { level, title, items[] } を返す
//   ・_buildAlertHtml() で見やすいHTMLメールを生成
//   ・Chat Webhook 用のプレーンテキストは引き続き生成
//
// 【アラートオブジェクト構造】
//   {
//     level : 'critical'|'warning'|'info',   // 緊急度（カード色に反映）
//     title : '🔥 明日納期',                  // セクション見出し
//     items : [                               // 案件ごとの配列
//       {
//         heading  : '株式会社○○「注文番号」',  // 案件名（太字）
//         lines    : ['項目名: …', '納期: …'], // 詳細テキスト行
//         links    : [                         // クリッカブルリンク
//           { label: '注文書PDF', url: 'https://…' },
//           { label: '管理シート', url: 'https://…' },
//         ],
//       },
//     ],
//   }
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
    MONITOR_CONFIG.ALERT_DELIVERY = s.alertDelivery !== false;
    MONITOR_CONFIG.ALERT_DEADLINE = s.alertDeadline !== false;
    MONITOR_CONFIG.ALERT_OVERDUE  = s.alertOverdue  !== false;
    MONITOR_CONFIG.ALERT_UNLINKED = s.alertUnlinked !== false;
  } catch(e) {
    Logger.log('[_loadMonitorConfig ERROR] ' + e.message);
  }
}

// ============================================================
// ★ リンク生成ユーティリティ
// ============================================================

/**
 * 管理シートの指定行への直接URLを返す
 */
function _getMgmtRowUrl(rowNum) {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    if (!sheet) return '';
    return 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
           '/edit#gid=' + sheet.getSheetId() + '&range=A' + rowNum;
  } catch(e) { return ''; }
}

/**
 * 行データからリンク配列を生成
 * @returns {Array} [ { label, url }, … ]
 */
function _buildLinkItems(row, rowNum) {
  var links = [];
  var quotePdf = String(row[MGMT_COLS.QUOTE_PDF_URL    - 1] || '').trim();
  var orderPdf = String(row[MGMT_COLS.ORDER_PDF_URL    - 1] || '').trim();
  var folder   = String(row[MGMT_COLS.DRIVE_FOLDER_URL - 1] || '').trim();
  if (quotePdf) links.push({ label: '見積書PDF',   url: quotePdf });
  if (orderPdf) links.push({ label: '注文書PDF',   url: orderPdf });
  if (folder)   links.push({ label: '保存フォルダ', url: folder  });
  if (rowNum) {
    var mgmt = _getMgmtRowUrl(rowNum);
    if (mgmt) links.push({ label: '管理シートで確認', url: mgmt });
  }
  return links;
}

// ============================================================
// ★ HTML メール生成
// ============================================================

/** HTML 特殊文字エスケープ */
function _escHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * アラートオブジェクト配列 → HTML メール本文を生成
 * @param {Array}  alertGroups  - { level, title, items[] } の配列
 * @param {string} reportTitle  - メールの大見出し
 * @returns {string} HTML文字列
 */
function _buildAlertHtml(alertGroups, reportTitle) {
  var LEVEL_STYLE = {
    critical: { border: '#d93025', bg: '#fce8e6', headerBg: '#d93025', headerText: '#fff' },
    warning:  { border: '#e37400', bg: '#fef7e0', headerBg: '#e37400', headerText: '#fff' },
    info:     { border: '#1a73e8', bg: '#e8f0fe', headerBg: '#1a73e8', headerText: '#fff' },
  };

  var LINK_ICON = {
    '見積書PDF':      '📄',
    '注文書PDF':      '📋',
    '保存フォルダ':    '📁',
    '管理シートで確認': '🔗',
    'フォルダを開く':  '📂',
  };

  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');

  var html = [];
  html.push(
    '<div style="font-family:\'Helvetica Neue\',Arial,\'Hiragino Kaku Gothic ProN\',sans-serif;' +
    'max-width:680px;margin:0 auto;background:#f1f3f4;padding:16px 12px;">',

    // ── ヘッダー ──
    '<div style="background:#1a73e8;border-radius:8px 8px 0 0;padding:18px 24px 16px;">',
    '<div style="color:#fff;font-size:18px;font-weight:bold;line-height:1.3;">',
    '📊 ' + _escHtml(reportTitle),
    '</div>',
    '<div style="color:#b8d4ff;font-size:12px;margin-top:4px;">' + _escHtml(now) + '　自動生成</div>',
    '</div>',

    // ── サマリーバー ──
    '<div style="background:#fff;border:1px solid #dadce0;border-top:none;' +
    'padding:10px 24px;font-size:13px;color:#5f6368;">',
    'アラートグループ数: <strong style="color:#202124;">' + alertGroups.length + '</strong>',
    '</div>'
  );

  // ── アラートカードごと ──
  alertGroups.forEach(function(group, gi) {
    var st = LEVEL_STYLE[group.level] || LEVEL_STYLE.info;

    html.push(
      '<div style="margin-top:12px;border-radius:6px;overflow:hidden;' +
      'border:1px solid ' + st.border + '33;">',

      // カード見出し
      '<div style="background:' + st.headerBg + ';padding:10px 18px;">',
      '<span style="color:' + st.headerText + ';font-size:14px;font-weight:bold;">' +
      _escHtml(group.title) + '</span>',
      '</div>'
    );

    // 案件ごとの行
    group.items.forEach(function(item, ii) {
      var isLast = ii === group.items.length - 1;
      var dividerStyle = isLast ? '' : 'border-bottom:1px solid ' + st.border + '44;';

      html.push(
        '<div style="background:' + st.bg + ';padding:12px 18px;' + dividerStyle + '">',

        // 案件名
        '<div style="font-size:14px;font-weight:bold;color:#202124;margin-bottom:5px;">',
        _escHtml(item.heading),
        '</div>'
      );

      // 詳細テキスト行
      if (item.lines && item.lines.length > 0) {
        html.push('<div style="font-size:13px;color:#444;line-height:1.9;margin-bottom:8px;">');
        item.lines.forEach(function(line) {
          // ⚠️ で始まる行は強調
          if (String(line).indexOf('⚠️') === 0) {
            html.push('<span style="color:' + st.border + ';font-weight:bold;">' +
                      _escHtml(line) + '</span><br>');
          } else {
            html.push(_escHtml(line) + '<br>');
          }
        });
        html.push('</div>');
      }

      // リンクボタン群
      if (item.links && item.links.length > 0) {
        html.push('<div style="display:flex;flex-wrap:wrap;gap:6px;margin-top:4px;">');
        item.links.forEach(function(lnk) {
          if (!lnk.url) return;
          var icon = LINK_ICON[lnk.label] || '🔗';
          html.push(
            '<a href="' + _escHtml(lnk.url) + '" ' +
            'style="display:inline-block;padding:5px 12px;background:#fff;' +
            'border:1.5px solid ' + st.border + ';border-radius:4px;' +
            'color:' + st.border + ';font-size:12px;font-weight:bold;' +
            'text-decoration:none;white-space:nowrap;">' +
            icon + '&nbsp;' + _escHtml(lnk.label) +
            '</a>'
          );
        });
        html.push('</div>');
      }

      html.push('</div>'); // /item
    });

    html.push('</div>'); // /card
  });

  // ── フッター ──
  html.push(
    '<div style="margin-top:16px;padding:10px 24px;font-size:11px;color:#80868b;text-align:center;">',
    'このメールは見積書・注文書管理システムにより自動送信されています',
    '</div>',
    '</div>'
  );

  return html.join('\n');
}

// ============================================================
// ★ メール送信（HTML カード版）
// ============================================================

/**
 * アラートオブジェクト配列をHTMLメールで送信
 */
function _sendAlertEmailHtml(alertGroups, reportTitle) {
  try {
    var notifyEmails = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAILS') || '';
    if (!notifyEmails) return;
    var emails = notifyEmails.split(',').map(function(e){ return e.trim(); }).filter(Boolean);

    var subject = '【見積管理システム】' + reportTitle + ' ' +
                  Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    var htmlBody  = _buildAlertHtml(alertGroups, reportTitle);
    var plainText = _alertGroupsToPlainText(alertGroups, reportTitle);

    emails.forEach(function(email) {
      try {
        GmailApp.sendEmail(email, subject, plainText, { htmlBody: htmlBody });
      } catch(e2) {
        Logger.log('[_sendAlertEmailHtml] 送信失敗: ' + email + ' / ' + e2.message);
      }
    });
    Logger.log('[_sendAlertEmailHtml] 送信完了: ' + emails.length + '件');
  } catch(e) {
    Logger.log('[_sendAlertEmailHtml ERROR] ' + e.message);
  }
}

// ============================================================
// テキスト変換（Webhook / プレーンテキストフォールバック用）
// ============================================================

function _alertGroupsToText(alertGroups) {
  return alertGroups.map(function(g) {
    return g.title + '\n' + g.items.slice(0, 5).map(function(it) {
      var lines = [it.heading].concat(it.lines || []);
      var linkLines = (it.links || []).map(function(l) {
        return '  [' + l.label + '] ' + l.url;
      });
      return lines.concat(linkLines).join('\n');
    }).join('\n\n');
  }).join('\n\n');
}

function _alertGroupsToPlainText(alertGroups, reportTitle) {
  return '【' + reportTitle + '】\n' +
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + '\n\n' +
    _alertGroupsToText(alertGroups);
}

// ============================================================
// メインエントリーポイント（毎日9時に実行）
// ============================================================
function runAllMonitoring() {
  Logger.log('[MONITOR] 監視開始: ' + nowJST());
  _loadMonitorConfig();

  var alertGroups = [];
  alertGroups = alertGroups.concat(_checkOcrFailures());
  if (MONITOR_CONFIG.ALERT_UNLINKED !== false) alertGroups = alertGroups.concat(_checkUnlinkedOrders());
  alertGroups = alertGroups.concat(_checkStagnantCases());
  if (MONITOR_CONFIG.ALERT_DELIVERY !== false) alertGroups = alertGroups.concat(_checkDeliveryDates());
  if (MONITOR_CONFIG.ALERT_DEADLINE !== false) alertGroups = alertGroups.concat(checkOrderDeadlines());
  alertGroups = alertGroups.concat(_checkUnprocessedDriveFiles());
  _updateTriggerHealth();

  if (alertGroups.length > 0) {
    var reportTitle = '自動監視レポート';
    var plainText   = '【📊 見積管理システム ' + reportTitle + '】\n' +
                      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + '\n\n' +
                      _alertGroupsToText(alertGroups);

    _sendMonitorAlert(plainText);
    _sendAlertEmailHtml(alertGroups, reportTitle);
    Logger.log('[MONITOR] アラート送信: ' + alertGroups.length + 'グループ');
  } else {
    Logger.log('[MONITOR] 異常なし');
  }

  return { alertCount: alertGroups.length, alerts: alertGroups };
}

// ============================================================
// ① OCR失敗検知
// ============================================================
function _checkOcrFailures() {
  var groups = [];
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName('OCR処理ログ');
    if (!sheet || sheet.getLastRow() <= 1) return groups;

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
          failFiles.push(String(row[1]||'').substring(0, 50));
        }
      } catch(e) {}
    });

    if (failCount >= MONITOR_CONFIG.OCR_FAIL_THRESHOLD) {
      groups.push({
        level: 'critical',
        title: '⚠️ OCR失敗が多発しています',
        items: [{
          heading: '直近24時間の失敗件数: ' + failCount + '件',
          lines:   ['ファイル例: ' + failFiles.slice(0,3).join(' / ')],
          links:   [],
        }],
      });
    }
  } catch(e) {
    Logger.log('[MONITOR] _checkOcrFailures: ' + e.message);
  }
  return groups;
}

// ============================================================
// ② 未紐づけ注文書の検知
// ============================================================
function _checkUnlinkedOrders() {
  var groups = [];
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var last  = sheet.getLastRow();
    if (last <= 1) return groups;

    var rawData = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
    var today   = new Date();
    var items   = [];

    rawData.forEach(function(row, i) {
      var orderNo   = String(row[MGMT_COLS.ORDER_NO  - 1] || '').trim();
      var linked    = _isLinkedVal(row[MGMT_COLS.LINKED - 1]);
      var quoteNo   = String(row[MGMT_COLS.QUOTE_NO  - 1] || '').trim();
      var orderDate = row[MGMT_COLS.ORDER_DATE - 1];
      var client    = String(row[MGMT_COLS.CLIENT   - 1] || '');

      if (!orderNo || linked || quoteNo) return;

      var days = 0;
      try {
        var d = orderDate instanceof Date ? orderDate : new Date(String(orderDate).replace(/\//g,'-'));
        if (!isNaN(d.getTime())) days = Math.floor((today - d) / (1000*60*60*24));
      } catch(e) {}

      if (days >= MONITOR_CONFIG.UNLINKED_ORDER_DAYS) {
        items.push({
          heading: client + '「' + orderNo + '」',
          lines:   ['注文日からの経過日数: ' + days + '日'],
          links:   _buildLinkItems(row, i + 2),
        });
      }
    });

    if (items.length > 0) {
      groups.push({ level: 'warning', title: '🔗 見積書未紐づけの注文書', items: items.slice(0, 5) });
    }
  } catch(e) {
    Logger.log('[MONITOR] _checkUnlinkedOrders: ' + e.message);
  }
  return groups;
}

// ============================================================
// ③ 案件の停滞アラート
// ============================================================
function _checkStagnantCases() {
  var groups = [];
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var last  = sheet.getLastRow();
    if (last <= 1) return groups;

    var rawData   = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
    var today     = new Date();
    var critItems = [];
    var warnItems = [];

    var ignoreStatuses = [
      CONFIG.STATUS.ORDERED, CONFIG.STATUS.DELIVERED,
      CONFIG.STATUS.CANCELLED, 'キャンセル', '失注',
    ];

    rawData.forEach(function(row, i) {
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

      var item = {
        heading: client + '「' + quoteNo + '」',
        lines:   ['ステータス: ' + status, '最終更新からの経過: ' + days + '日'],
        links:   _buildLinkItems(row, i + 2),
      };

      if (days >= MONITOR_CONFIG.STAGNANT_ALERT_DAYS) {
        critItems.push(item);
      } else if (days >= MONITOR_CONFIG.STAGNANT_WARN_DAYS) {
        warnItems.push(item);
      }
    });

    if (critItems.length > 0) {
      groups.push({ level: 'critical', title: '🚨 停滞警告（' + MONITOR_CONFIG.STAGNANT_ALERT_DAYS + '日以上）', items: critItems.slice(0,5) });
    }
    if (warnItems.length > 0) {
      groups.push({ level: 'warning',  title: '⚠️ 停滞注意（' + MONITOR_CONFIG.STAGNANT_WARN_DAYS  + '日以上）', items: warnItems.slice(0,5) });
    }
  } catch(e) {
    Logger.log('[MONITOR] _checkStagnantCases: ' + e.message);
  }
  return groups;
}

// ============================================================
// ④ 納期アラート
// ============================================================
function _checkDeliveryDates() {
  var groups = [];
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var last  = sheet.getLastRow();
    if (last <= 1) return groups;

    var rawData      = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
    var today        = new Date();
    var overdueItems = [], urgentItems = [], remindItems = [];

    rawData.forEach(function(row, i) {
      var status       = String(row[MGMT_COLS.STATUS       - 1] || '');
      var deliveryDate = row[MGMT_COLS.DELIVERY_DATE       - 1];
      var client       = String(row[MGMT_COLS.CLIENT       - 1] || '');
      var orderNo      = String(row[MGMT_COLS.ORDER_NO     - 1] || '');
      var subject      = String(row[MGMT_COLS.SUBJECT      - 1] || '');

      if (!deliveryDate || !orderNo) return;
      if (status === CONFIG.STATUS.DELIVERED || status === CONFIG.STATUS.CANCELLED) return;

      var days = 9999;
      var deliveryStr = '';
      try {
        var d = deliveryDate instanceof Date ? deliveryDate : new Date(String(deliveryDate).replace(/\//g,'-'));
        if (!isNaN(d.getTime())) {
          days = Math.floor((d - today) / (1000*60*60*24));
          deliveryStr = _toDateStr(d);
        }
      } catch(e) {}

      var lines = [];
      if (subject)     lines.push('項目名: ' + subject);
      if (deliveryStr) lines.push('納期: '   + deliveryStr);

      var item = {
        heading: client + (orderNo ? '「' + orderNo + '」' : ''),
        lines:   lines,
        links:   _buildLinkItems(row, i + 2),
      };

      if (days < 0) {
        item.lines.push('⚠️ ' + Math.abs(days) + '日超過');
        overdueItems.push(item);
      } else if (days <= MONITOR_CONFIG.DELIVERY_URGENT_DAYS) {
        item.lines.push('⚠️ 明日が納期');
        urgentItems.push(item);
      } else if (days <= MONITOR_CONFIG.DELIVERY_REMIND_DAYS) {
        item.lines.push('⚠️ あと' + days + '日');
        remindItems.push(item);
      }
    });

    if (overdueItems.length > 0) groups.push({ level: 'critical', title: '🧨 納期超過',       items: overdueItems.slice(0,5) });
    if (urgentItems.length  > 0) groups.push({ level: 'critical', title: '🔥 明日納期',       items: urgentItems.slice(0,5)  });
    if (remindItems.length  > 0) groups.push({ level: 'warning',  title: '📅 納期リマインド', items: remindItems.slice(0,5)  });
  } catch(e) {
    Logger.log('[MONITOR] _checkDeliveryDates: ' + e.message);
  }
  return groups;
}

// ============================================================
// ⑤ 注文書の発注期限アラート（DEADLINE_NOTIFIED で重複防止）
// ============================================================
function checkOrderDeadlines() {
  var groups = [];
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var last  = sheet.getLastRow();
    if (last <= 1) return groups;

    var readCols     = Math.max(sheet.getLastColumn(), 32);
    var data         = sheet.getRange(2, 1, last - 1, readCols).getValues();
    var today        = new Date(); today.setHours(0, 0, 0, 0);
    var overdueItems = [], urgentItems = [], remindItems = [];

    data.forEach(function(row, i) {
      var id       = String(row[MGMT_COLS.ID                - 1] || '');
      var orderNo  = String(row[MGMT_COLS.ORDER_NO          - 1] || '').trim();
      var status   = String(row[MGMT_COLS.STATUS            - 1] || '');
      var deadline = row[MGMT_COLS.ORDER_DEADLINE           - 1];
      var notified = String(row[MGMT_COLS.DEADLINE_NOTIFIED - 1] || '');
      var client   = String(row[MGMT_COLS.CLIENT            - 1] || '');
      var slipNo   = String(row[MGMT_COLS.ORDER_SLIP_NO     - 1] || '');
      var subject  = String(row[MGMT_COLS.SUBJECT           - 1] || '');
      var delivery = row[MGMT_COLS.DELIVERY_DATE            - 1];

      if (!id || !deadline) return;
      if (status === CONFIG.STATUS.DELIVERED || status === CONFIG.STATUS.CANCELLED) return;

      var deadlineDate;
      try {
        deadlineDate = deadline instanceof Date ? deadline : new Date(String(deadline).replace(/\//g, '-'));
        if (isNaN(deadlineDate.getTime())) return;
        deadlineDate.setHours(0, 0, 0, 0);
      } catch(e) { return; }

      var deliveryStr = '';
      try {
        var dv = delivery instanceof Date ? delivery : new Date(String(delivery).replace(/\//g,'-'));
        if (!isNaN(dv.getTime())) deliveryStr = _toDateStr(dv);
      } catch(e) {}

      var diffDays = Math.round((deadlineDate - today) / (1000 * 60 * 60 * 24));
      var rowNum   = i + 2;
      var todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');

      var lines = ['発注期限: ' + _toDateStr(deadlineDate)];
      if (deliveryStr) lines.push('納期: ' + deliveryStr);
      if (subject)     lines.push('項目名: ' + subject);
      if (slipNo)      lines.push('伝票番号: ' + slipNo);

      var item = {
        heading: client + (orderNo ? '「' + orderNo + '」' : ''),
        lines:   lines,
        links:   _buildLinkItems(row, rowNum),
      };

      if (diffDays < 0) {
        item.lines.push('⚠️ ' + Math.abs(diffDays) + '日超過');
        overdueItems.push(item);
      } else if (diffDays <= 1 && !_hasNotifiedToday(notified, 'urgent', todayStr)) {
        item.lines.push('⚠️ 明日が期限');
        urgentItems.push(item);
        _writeDeadlineNotified(sheet, rowNum, 'urgent:' + todayStr);
      } else if (diffDays <= 3 && !_hasNotifiedToday(notified, '3d', todayStr)) {
        item.lines.push('⚠️ あと' + diffDays + '日');
        remindItems.push(item);
        _writeDeadlineNotified(sheet, rowNum, '3d:' + todayStr);
      } else if (diffDays <= 7 && !_hasNotifiedToday(notified, '7d', todayStr)) {
        item.lines.push('⚠️ あと' + diffDays + '日');
        remindItems.push(item);
        _writeDeadlineNotified(sheet, rowNum, '7d:' + todayStr);
      }
    });

    if (overdueItems.length > 0) groups.push({ level: 'critical', title: '🧨 [発注期限超過]',     items: overdueItems });
    if (urgentItems.length  > 0) groups.push({ level: 'critical', title: '🔥 [発注期限 明日]',    items: urgentItems  });
    if (remindItems.length  > 0) groups.push({ level: 'warning',  title: '📅 [発注期限 事前通知]', items: remindItems  });
  } catch(e) {
    Logger.log('[checkOrderDeadlines ERROR] ' + e.message);
  }
  return groups;
}

function _hasNotifiedToday(notifiedStr, level, todayStr) {
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
// ⑥ 未処理Driveファイルの滞留検知
// ============================================================
function _checkUnprocessedDriveFiles() {
  var groups = [];
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
            oldFiles.push(file.getName().substring(0, 50));
          }
        }
        if (oldFiles.length > 0) {
          groups.push({
            level: 'warning',
            title: '📂 未処理ファイルが滞留しています',
            items: [{
              heading: f.label + '（' + oldFiles.length + '件）',
              lines:   ['例: ' + oldFiles.slice(0,2).join(' / ')],
              links:   [{ label: 'フォルダを開く', url: 'https://drive.google.com/drive/folders/' + f.id }],
            }],
          });
        }
      } catch(e) {
        Logger.log('[MONITOR] folder check error: ' + f.label + ' ' + e.message);
      }
    });
  } catch(e) {
    Logger.log('[MONITOR] _checkUnprocessedDriveFiles: ' + e.message);
  }
  return groups;
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
// 通知送信（Chat Webhook）
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
      GmailApp.sendEmail(Session.getActiveUser().getEmail(), '【見積管理システム監視アラート】', message);
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
  if (existing.indexOf('runDeadlineAlerts') < 0) {
    ScriptApp.newTrigger('runDeadlineAlerts').timeBased().everyHours(6).create();
    Logger.log('[MONITOR] runDeadlineAlerts トリガー登録（6時間ごと）');
  }
}

// 発注期限アラートのスタンドアロン実行
function runDeadlineAlerts() {
  Logger.log('[DEADLINE] 発注期限チェック開始: ' + nowJST());
  _loadMonitorConfig();
  if (MONITOR_CONFIG.ALERT_DEADLINE === false) {
    Logger.log('[DEADLINE] アラート無効');
    return;
  }
  var groups = [];
  if (MONITOR_CONFIG.ALERT_DELIVERY !== false) groups = groups.concat(_checkDeliveryDates());
  if (MONITOR_CONFIG.ALERT_DEADLINE !== false) groups = groups.concat(checkOrderDeadlines());

  if (groups.length > 0) {
    var reportTitle = '納期・発注期限アラート';
    var plainText   = '【📦 ' + reportTitle + '】\n' +
                      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + '\n\n' +
                      _alertGroupsToText(groups);
    _sendMonitorAlert(plainText);
    _sendAlertEmailHtml(groups, reportTitle);
    Logger.log('[DEADLINE] アラート送信: ' + groups.length + 'グループ');
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
  Logger.log('アラートグループ数: ' + result.alertCount);
  Logger.log('=== 監視テスト完了 ===');
}

function testSendAlert() {
  _sendMonitorAlert('✅ テスト通知\n見積管理システムの監視が正常に動作しています。\n送信時刻: ' + nowJST());
}

/**
 * HTMLメールのダミーデータでプレビューをメール送信するテスト関数
 * GASエディタから手動実行してメールの見た目を確認できます
 */
function testSendAlertEmailHtml() {
  var dummyGroups = [
    {
      level: 'critical',
      title: '🔥 明日納期',
      items: [
        {
          heading: '株式会社サンプル「ORD-2026-001」',
          lines:   ['項目名: テスト基板 Rev.2', '納期: 2026/05/19', '⚠️ 明日が納期'],
          links:   [
            { label: '注文書PDF',     url: 'https://drive.google.com/file/d/dummy1/view' },
            { label: '保存フォルダ',   url: 'https://drive.google.com/drive/folders/dummy2' },
            { label: '管理シートで確認', url: 'https://docs.google.com/spreadsheets/d/dummy3/edit#gid=0&range=A5' },
          ],
        },
        {
          heading: '○○工業「ORD-2026-002」',
          lines:   ['項目名: メイン基板量産分', '納期: 2026/05/19', '⚠️ 明日が納期'],
          links:   [
            { label: '注文書PDF',     url: 'https://drive.google.com/file/d/dummy4/view' },
            { label: '管理シートで確認', url: 'https://docs.google.com/spreadsheets/d/dummy3/edit#gid=0&range=A8' },
          ],
        },
      ],
    },
    {
      level: 'warning',
      title: '⚠️ 停滞注意（7日以上）',
      items: [
        {
          heading: '株式会社テスト「QM-20260501-1234」',
          lines:   ['ステータス: 送信済み', '最終更新からの経過: 10日'],
          links:   [
            { label: '見積書PDF',     url: 'https://drive.google.com/file/d/dummy5/view' },
            { label: '管理シートで確認', url: 'https://docs.google.com/spreadsheets/d/dummy3/edit#gid=0&range=A12' },
          ],
        },
      ],
    },
    {
      level: 'info',
      title: '📂 未処理ファイルが滞留しています',
      items: [{
        heading: '注文書（試作）インポート（3件）',
        lines:   ['例: 注文書_A社_20260518.pdf / 発注書_B商事.pdf'],
        links:   [{ label: 'フォルダを開く', url: 'https://drive.google.com/drive/folders/dummy6' }],
      }],
    },
  ];

  _sendAlertEmailHtml(dummyGroups, '自動監視レポート（テスト）');
  Logger.log('[TEST] テストメール送信完了');
}

// ============================================================
// 旧来の簡易アラート（互換用）
// ============================================================
function checkAndSendAlerts() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT || '見積提出管理');
  if (!sheet) return;

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];

  var isLatestIdx = headers.indexOf('最新フラグ');
  var statusIdx   = headers.indexOf('ステータス');
  var dateIdx     = headers.indexOf('提出日');
  var numIdx      = headers.indexOf('見積No');
  var boardIdx    = headers.indexOf('基板名');
  var deliveryIdx = headers.indexOf('納期');
  var today       = new Date();
  var groups      = [];

  for (var i = 1; i < data.length; i++) {
    var row      = data[i];
    var isLatest = isLatestIdx > -1 ? row[isLatestIdx] : true;
    if (!isLatest) continue;

    var status      = statusIdx > -1 ? row[statusIdx] : '';
    var quoteNumber = numIdx > -1 ? row[numIdx] : '不明';
    var subject     = boardIdx > -1 ? row[boardIdx] : '';
    var mgmtUrl     = _getMgmtRowUrl(i + 1);
    var baseLinks   = mgmtUrl ? [{ label: '管理シートで確認', url: mgmtUrl }] : [];

    if (status !== '発注済' && status !== '納品済み' && status !== 'キャンセル') {
      var quoteDateStr = dateIdx > -1 ? row[dateIdx] : null;
      if (quoteDateStr) {
        var quoteDate = new Date(quoteDateStr);
        var diffDays  = Math.floor((today - quoteDate) / (1000 * 60 * 60 * 24));
        if (diffDays >= 14) {
          groups.push({ level: 'critical', title: '🚨 停滞警告', items: [{
            heading: '見積[' + quoteNumber + '] ' + subject,
            lines:   [diffDays + '日放置されています。至急確認を。'],
            links:   baseLinks,
          }]});
        } else if (diffDays >= 7) {
          groups.push({ level: 'warning', title: '⚠️ 未着', items: [{
            heading: '見積[' + quoteNumber + '] ' + subject,
            lines:   ['発行後 ' + diffDays + '日 経過しています。'],
            links:   baseLinks,
          }]});
        }
      }
    }

    if (status === '発注済' || status === '作成中') {
      var deliveryDateStr = deliveryIdx > -1 ? row[deliveryIdx] : null;
      if (deliveryDateStr) {
        var deliveryDate   = new Date(deliveryDateStr);
        var daysToDelivery = Math.floor((deliveryDate - today) / (1000 * 60 * 60 * 24));
        if (daysToDelivery <= 7) {
          var level  = daysToDelivery < 0 || daysToDelivery === 1 ? 'critical' : 'warning';
          var title  = daysToDelivery < 0 ? '🧨 納期超過' : daysToDelivery === 1 ? '🔥 明日納期' : '📅 納期リマインド';
          var detail = daysToDelivery < 0
            ? '納期を ' + Math.abs(daysToDelivery) + '日 超過！'
            : daysToDelivery === 1 ? '明日が納期です。' : '納期まであと' + daysToDelivery + '日。';
          groups.push({ level: level, title: title, items: [{
            heading: '注文[' + quoteNumber + '] ' + subject,
            lines:   [detail],
            links:   baseLinks,
          }]});
        }
      }
    }
  }

  if (groups.length > 0) {
    _sendMonitorAlert('【システム自動通知・運用アラート】\n' + _alertGroupsToText(groups));
    _sendAlertEmailHtml(groups, 'システム自動通知・運用アラート');
  }
}