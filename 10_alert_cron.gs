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
//   ・【文字化け修正】絵文字を &#xXXXXX; HTML数値文字参照に変換 + GmailApp.sendEmail 使用
//
// 【アラートオブジェクト構造】
//   {
//     level : 'critical'|'warning'|'info',
//     title : '🔥 明日納期',
//     items : [
//       {
//         heading  : '株式会社○○「注文番号」',
//         lines    : ['項目名: …', '納期: …'],
//         links    : [
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
  CHECKLIST_STALE_DAYS  : 3,
  // アラート種別 ON/OFF（デフォルト: 全ON）
  ALERT_DELIVERY   : true,
  ALERT_DEADLINE   : true,
  ALERT_OVERDUE    : true,
  ALERT_UNLINKED   : true,
  ALERT_CHECKLIST  : true,
  ALERT_STAGNANT   : true,   // 停滞アラート
  ALERT_OCR        : true,   // OCR失敗アラート
  ALERT_DRIVE      : true,   // 未処理Driveファイルアラート
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
    if (s.checklistStaleDays) MONITOR_CONFIG.CHECKLIST_STALE_DAYS  = Number(s.checklistStaleDays);
    // フラグ: undefined の場合は true（未設定＝有効）、明示的に false にした場合のみ無効
    MONITOR_CONFIG.ALERT_DELIVERY  = s.alertDelivery  !== false;
    MONITOR_CONFIG.ALERT_DEADLINE  = s.alertDeadline  !== false;
    MONITOR_CONFIG.ALERT_OVERDUE   = s.alertOverdue   !== false;
    MONITOR_CONFIG.ALERT_UNLINKED  = s.alertUnlinked  !== false;
    MONITOR_CONFIG.ALERT_CHECKLIST = s.alertChecklist !== false;
    MONITOR_CONFIG.ALERT_STAGNANT  = s.alertStagnant  !== false;
    MONITOR_CONFIG.ALERT_OCR       = s.alertOcr       !== false;
    MONITOR_CONFIG.ALERT_DRIVE     = s.alertDrive     !== false;
  } catch(e) {
    Logger.log('[_loadMonitorConfig ERROR] ' + e.message);
  }
}

// ============================================================
// リンク生成ユーティリティ
// ============================================================

function _getMgmtRowUrl(rowNum) {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    if (!sheet) return '';
    return 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
           '/edit#gid=' + sheet.getSheetId() + '&range=A' + rowNum;
  } catch(e) { return ''; }
}

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
// HTML メール生成
// ============================================================

function _escHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ============================================================
// 絵文字 HTML エンコード（文字化け防止）
// ============================================================

// サロゲートペア → HTML数値文字参照 &#xXXXXX; に変換
function _emojiToHtml(str) {
  return String(str || '').replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g, function(pair) {
    var hi = pair.charCodeAt(0);
    var lo = pair.charCodeAt(1);
    var cp = ((hi - 0xD800) * 0x400) + (lo - 0xDC00) + 0x10000;
    return '&#x' + cp.toString(16).toUpperCase() + ';';
  });
}

// HTML特殊文字エスケープ ＋ 絵文字を数値文字参照に変換（HTML本文用）
function _escHtmlEmoji(str) {
  return _emojiToHtml(
    String(str || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
  );
}

// 絵文字を除去（件名・プレーンテキスト用）
function _stripEmoji(str) {
  return String(str || '')
    .replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g, '')
    .replace(/[☀-➿][️]?/g, '')
    .replace(/[︀-️]/g, '');
}

function _buildAlertHtml(alertGroups, reportTitle) {
  var LEVEL_STYLE = {
    critical: { border: '#d93025', bg: '#fce8e6', headerBg: '#d93025', headerText: '#fff' },
    warning:  { border: '#e37400', bg: '#fef7e0', headerBg: '#e37400', headerText: '#fff' },
    info:     { border: '#1a73e8', bg: '#e8f0fe', headerBg: '#1a73e8', headerText: '#fff' },
  };

  var LINK_ICON = {
    '\u898b\u7a4d\u66f8PDF':      '\ud83d\udcc4',
    '\u6ce8\u6587\u66f8PDF':      '\ud83d\udccb',
    '\u4fdd\u5b58\u30d5\u30a9\u30eb\u30c0':    '\ud83d\udcc1',
    '\u7ba1\u7406\u30b7\u30fc\u30c8\u3067\u78ba\u8a8d': '\ud83d\udd17',
    '\u30d5\u30a9\u30eb\u30c0\u3092\u958b\u304f':  '\ud83d\udcc2',
  };

  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy\u5e74MM\u6708dd\u65e5 HH:mm');

  var html = [];
  html.push(
    '<meta charset="UTF-8">',
    '<div style="font-family:\'Helvetica Neue\',Arial,\'Hiragino Kaku Gothic ProN\',sans-serif;' +
    'max-width:680px;margin:0 auto;background:#f1f3f4;padding:16px 12px;">',

    '<div style="background:#1a73e8;border-radius:8px 8px 0 0;padding:18px 24px 16px;">',
    '<div style="color:#fff;font-size:18px;font-weight:bold;line-height:1.3;">',
    _emojiToHtml('\ud83d\udcca ') + _escHtml(reportTitle),
    '</div>',
    '<div style="color:#b8d4ff;font-size:12px;margin-top:4px;">' + _escHtml(now) + '\u3000\u81ea\u52d5\u751f\u6210</div>',
    '</div>',

    '<div style="background:#fff;border:1px solid #dadce0;border-top:none;' +
    'padding:10px 24px;font-size:13px;color:#5f6368;">',
    '\u30a2\u30e9\u30fc\u30c8\u30b0\u30eb\u30fc\u30d7\u6570: <strong style="color:#202124;">' + alertGroups.length + '</strong>',
    '</div>'
  );

  alertGroups.forEach(function(group) {
    var st = LEVEL_STYLE[group.level] || LEVEL_STYLE.info;

    html.push(
      '<div style="margin-top:12px;border-radius:6px;overflow:hidden;' +
      'border:1px solid ' + st.border + '33;">',
      '<div style="background:' + st.headerBg + ';padding:10px 18px;">',
      '<span style="color:' + st.headerText + ';font-size:14px;font-weight:bold;">' +
      _escHtmlEmoji(group.title) + '</span>',
      '</div>'
    );

    group.items.forEach(function(item, ii) {
      var isLast = ii === group.items.length - 1;
      var dividerStyle = isLast ? '' : 'border-bottom:1px solid ' + st.border + '44;';

      html.push(
        '<div style="background:' + st.bg + ';padding:12px 18px;' + dividerStyle + '">',
        '<div style="font-size:14px;font-weight:bold;color:#202124;margin-bottom:5px;">',
        _escHtmlEmoji(item.heading),
        '</div>'
      );

      if (item.lines && item.lines.length > 0) {
        html.push('<div style="font-size:13px;color:#444;line-height:1.9;margin-bottom:8px;">');
        item.lines.forEach(function(line) {
          if (String(line).indexOf('\u26a0\ufe0f') === 0) {
            html.push('<span style="color:' + st.border + ';font-weight:bold;">' +
                      _escHtmlEmoji(line) + '</span><br>');
          } else {
            html.push(_escHtmlEmoji(line) + '<br>');
          }
        });
        html.push('</div>');
      }

      if (item.links && item.links.length > 0) {
        html.push('<div style="display:flex;flex-wrap:wrap;gap:6px;margin-top:4px;">');
        item.links.forEach(function(lnk) {
          if (!lnk.url) return;
          var icon = LINK_ICON[lnk.label] || '\ud83d\udd17';
          html.push(
            '<a href="' + _escHtml(lnk.url) + '" ' +
            'style="display:inline-block;padding:5px 12px;background:#fff;' +
            'border:1.5px solid ' + st.border + ';border-radius:4px;' +
            'color:' + st.border + ';font-size:12px;font-weight:bold;' +
            'text-decoration:none;white-space:nowrap;">' +
            _emojiToHtml(icon) + '&nbsp;' + _escHtml(lnk.label) +
            '</a>'
          );
        });
        html.push('</div>');
      }

      html.push('</div>');
    });

    html.push('</div>');
  });

  html.push(
    '<div style="margin-top:16px;padding:10px 24px;font-size:11px;color:#80868b;text-align:center;">',
    '\u3053\u306e\u30e1\u30fc\u30eb\u306f\u898b\u7a4d\u66f8\u30fb\u6ce8\u6587\u66f8\u7ba1\u7406\u30b7\u30b9\u30c6\u30e0\u306b\u3088\u308a\u81ea\u52d5\u9001\u4fe1\u3055\u308c\u3066\u3044\u307e\u3059',
    '</div>',
    '</div>'
  );

  return html.join('\n');
}

// ============================================================
// ★★★ メール送信（文字化け修正版）★★★
// 【対策1】GmailApp.sendEmail 使用（MailApp より UTF-8 信頼性高）
// 【対策2】HTML本文の絵文字は _escHtmlEmoji/_emojiToHtml で
//          &#xXXXXX; 数値文字参照に変換済み → メールクライアントで確実表示
// 【対策3】件名・プレーンテキストは _stripEmoji() で絵文字除去
// ============================================================

function _sendAlertEmailHtml(alertGroups, reportTitle) {
  try {
    var notifyEmails = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAILS') || '';
    if (!notifyEmails) return;
    var emails = notifyEmails.split(',').map(function(e){ return e.trim(); }).filter(Boolean);

    var subject  = '\u3010\u898b\u7a4d\u7ba1\u7406\u30b7\u30b9\u30c6\u30e0\u3011' + _stripEmoji(reportTitle).trim() + ' ' +
                   Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    var htmlBody  = _buildAlertHtml(alertGroups, reportTitle);
    var plainText = _alertGroupsToPlainText(alertGroups, reportTitle);

    emails.forEach(function(email) {
      try {
        // ★ GmailApp.sendEmail を使用（MailApp より UTF-8/絵文字の扱いが信頼性高）
        //    HTMLBody の絵文字は _escHtmlEmoji/_emojiToHtml で &#xXXXXX; 変換済み
        GmailApp.sendEmail(
          email,
          subject,
          plainText,
          {
            htmlBody: htmlBody,
            name    : '\u898b\u7a4d\u7ba1\u7406\u30b7\u30b9\u30c6\u30e0',
          }
        );
      } catch(e2) {
        Logger.log('[_sendAlertEmailHtml] \u9001\u4fe1\u5931\u6557: ' + email + ' / ' + e2.message);
      }
    });
    Logger.log('[_sendAlertEmailHtml] \u9001\u4fe1\u5b8c\u4e86: ' + emails.length + '\u4ef6');
  } catch(e) {
    Logger.log('[_sendAlertEmailHtml ERROR] ' + e.message);
  }
}

// ============================================================
// テキスト変換（Webhook / プレーンテキストフォールバック用）
// ============================================================

function _alertGroupsToText(alertGroups) {
  return alertGroups.map(function(g) {
    return _stripEmoji(g.title) + '\n' + g.items.slice(0, 5).map(function(it) {
      var lines = [it.heading].concat(it.lines || []);
      var linkLines = (it.links || []).map(function(l) {
        return '  [' + l.label + '] ' + l.url;
      });
      return lines.concat(linkLines).join('\n');
    }).join('\n\n');
  }).join('\n\n');
}

function _alertGroupsToPlainText(alertGroups, reportTitle) {
  return '\u3010' + reportTitle + '\u3011\n' +
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + '\n\n' +
    _alertGroupsToText(alertGroups);
}

// ============================================================
// Chat通知（Google Chat Webhook）
// ============================================================

function _sendMonitorAlert(message) {
  try {
    var webhookUrl = PropertiesService.getScriptProperties().getProperty('GOOGLE_CHAT_WEBHOOK_URL') || '';
    if (!webhookUrl || webhookUrl.indexOf('XXXX') !== -1) return;
    var payload = { text: message };
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
  } catch(e) {
    Logger.log('[_sendMonitorAlert ERROR] ' + e.message);
  }
}

// ============================================================
// トリガーヘルスチェック
// ============================================================

function _updateTriggerHealth() {
  try {
    PropertiesService.getScriptProperties().setProperty(
      MONITOR_CONFIG.TRIGGER_HEALTH_KEY,
      new Date().toISOString()
    );
  } catch(e) {}
}

function _isLinkedVal(v) {
  return v === true || String(v).toUpperCase() === 'TRUE';
}

function _toDateStr(v) {
  try {
    if (!v) return '';
    var d = v instanceof Date ? v : new Date(String(v).replace(/\//g, '-'));
    if (isNaN(d.getTime())) return String(v);
    return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
  } catch(e) { return String(v || ''); }
}

// ============================================================
// メインエントリーポイント（毎日9時に実行）
// ============================================================

function runAllMonitoring() {
  Logger.log('[MONITOR] \u76e3\u8996\u958b\u59cb: ' + nowJST());
  _loadMonitorConfig();

  var alertGroups = [];
  if (MONITOR_CONFIG.ALERT_OCR       !== false) alertGroups = alertGroups.concat(_checkOcrFailures());
  if (MONITOR_CONFIG.ALERT_UNLINKED  !== false) alertGroups = alertGroups.concat(_checkUnlinkedOrders());
  if (MONITOR_CONFIG.ALERT_STAGNANT  !== false) alertGroups = alertGroups.concat(_checkStagnantCases());
  if (MONITOR_CONFIG.ALERT_DELIVERY  !== false) alertGroups = alertGroups.concat(_checkDeliveryDates());
  if (MONITOR_CONFIG.ALERT_DEADLINE  !== false) alertGroups = alertGroups.concat(checkOrderDeadlines());
  if (MONITOR_CONFIG.ALERT_DRIVE     !== false) alertGroups = alertGroups.concat(_checkUnprocessedDriveFiles());
  if (MONITOR_CONFIG.ALERT_CHECKLIST !== false) alertGroups = alertGroups.concat(_checkChecklistAlerts());
  _updateTriggerHealth();

  if (alertGroups.length > 0) {
    var reportTitle = '\u81ea\u52d5\u76e3\u8996\u30ec\u30dd\u30fc\u30c8';
    var plainText   = '\u3010\ud83d\udcca \u898b\u7a4d\u7ba1\u7406\u30b7\u30b9\u30c6\u30e0 ' + reportTitle + '\u3011\n' +
                      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + '\n\n' +
                      _alertGroupsToText(alertGroups);

    _sendMonitorAlert(plainText);
    _sendAlertEmailHtml(alertGroups, reportTitle);
    Logger.log('[MONITOR] \u30a2\u30e9\u30fc\u30c8\u9001\u4fe1: ' + alertGroups.length + '\u30b0\u30eb\u30fc\u30d7');
  } else {
    Logger.log('[MONITOR] \u7570\u5e38\u306a\u3057');
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
    var sheet = ss.getSheetByName('OCR\u51e6\u7406\u30ed\u30b0');
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
        title: '\u26a0\ufe0f OCR\u5931\u6557\u304c\u591a\u767a\u3057\u3066\u3044\u307e\u3059',
        items: [{
          heading: '\u76f4\u8fd124\u6642\u9593\u306e\u5931\u6557\u4ef6\u6570: ' + failCount + '\u4ef6',
          lines:   ['\u30d5\u30a1\u30a4\u30eb\u4f8b: ' + failFiles.slice(0,3).join(' / ')],
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
          heading: client + '\u300c' + orderNo + '\u300d',
          lines:   ['\u6ce8\u6587\u65e5\u304b\u3089\u306e\u7d4c\u904e\u65e5\u6570: ' + days + '\u65e5'],
          links:   _buildLinkItems(row, i + 2),
        });
      }
    });

    if (items.length > 0) {
      groups.push({ level: 'warning', title: '\ud83d\udd17 \u898b\u7a4d\u66f8\u672a\u7d10\u3065\u3051\u306e\u6ce8\u6587\u66f8', items: items.slice(0, 5) });
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
      CONFIG.STATUS.CANCELLED, '\u30ad\u30e3\u30f3\u30bb\u30eb', '\u5931\u6ce8',
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
        heading: client + '\u300c' + quoteNo + '\u300d',
        lines:   ['\u30b9\u30c6\u30fc\u30bf\u30b9: ' + status, '\u6700\u7d42\u66f4\u65b0\u304b\u3089\u306e\u7d4c\u904e: ' + days + '\u65e5'],
        links:   _buildLinkItems(row, i + 2),
      };

      if (days >= MONITOR_CONFIG.STAGNANT_ALERT_DAYS) {
        critItems.push(item);
      } else if (days >= MONITOR_CONFIG.STAGNANT_WARN_DAYS) {
        warnItems.push(item);
      }
    });

    if (critItems.length > 0) {
      groups.push({ level: 'critical', title: '\ud83d\udea8 \u505c\u6ede\u8b66\u544a\uff08' + MONITOR_CONFIG.STAGNANT_ALERT_DAYS + '\u65e5\u4ee5\u4e0a\uff09', items: critItems.slice(0,5) });
    }
    if (warnItems.length > 0) {
      groups.push({ level: 'warning',  title: '\u26a0\ufe0f \u505c\u6ede\u6ce8\u610f\uff08' + MONITOR_CONFIG.STAGNANT_WARN_DAYS  + '\u65e5\u4ee5\u4e0a\uff09', items: warnItems.slice(0,5) });
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
      if (subject)     lines.push('\u9805\u76ee\u540d: ' + subject);
      if (deliveryStr) lines.push('\u7d0d\u671f: '   + deliveryStr);

      var item = {
        heading: client + (orderNo ? '\u300c' + orderNo + '\u300d' : ''),
        lines:   lines,
        links:   _buildLinkItems(row, i + 2),
      };

      if (days < 0) {
        item.lines.push('\u26a0\ufe0f ' + Math.abs(days) + '\u65e5\u8d85\u904e');
        overdueItems.push(item);
      } else if (days <= MONITOR_CONFIG.DELIVERY_URGENT_DAYS) {
        item.lines.push('\u26a0\ufe0f \u660e\u65e5\u304c\u7d0d\u671f');
        urgentItems.push(item);
      } else if (days <= MONITOR_CONFIG.DELIVERY_REMIND_DAYS) {
        item.lines.push('\u26a0\ufe0f \u3042\u3068' + days + '\u65e5');
        remindItems.push(item);
      }
    });

    if (overdueItems.length > 0 && MONITOR_CONFIG.ALERT_OVERDUE !== false) groups.push({ level: 'critical', title: '\ud83e\udde8 \u7d0d\u671f\u8d85\u904e',       items: overdueItems.slice(0,5) });
    if (urgentItems.length  > 0) groups.push({ level: 'critical', title: '\ud83d\udd25 \u660e\u65e5\u7d0d\u671f',       items: urgentItems.slice(0,5)  });
    if (remindItems.length  > 0) groups.push({ level: 'warning',  title: '\ud83d\udcc5 \u7d0d\u671f\u30ea\u30de\u30a4\u30f3\u30c9', items: remindItems.slice(0,5)  });
  } catch(e) {
    Logger.log('[MONITOR] _checkDeliveryDates: ' + e.message);
  }
  return groups;
}

// ============================================================
// ⑤ 注文書の発注期限アラート
// ============================================================
// ⑥ 注文書の発注期限アラート
//   【バグ修正】DELIVERY_DATE(col23) → ORDER_DEADLINE(col29) に修正
//   【追加】件名・伝票番号・納期をメール本文に追加
//   【追加】DEADLINE_NOTIFIED(col32)で当日重複通知を防止
// ============================================================

function checkOrderDeadlines() {
  var groups = [];
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var last  = sheet.getLastRow();
    if (last <= 1) return groups;

    // col32（DEADLINE_NOTIFIED）まで読む
    var readCols     = Math.max(sheet.getLastColumn(), 32);
    var data         = sheet.getRange(2, 1, last - 1, readCols).getValues();
    var today        = new Date(); today.setHours(0, 0, 0, 0);
    var todayStr     = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
    var overdueItems = [], urgentItems = [], remindItems = [];

    data.forEach(function(row, i) {
      var orderNo  = String(row[MGMT_COLS.ORDER_NO          - 1] || '').trim();
      var status   = String(row[MGMT_COLS.STATUS            - 1] || '');
      var client   = String(row[MGMT_COLS.CLIENT            - 1] || '');
      // ★ バグ修正: DELIVERY_DATE → ORDER_DEADLINE（発注期限 col29）
      var deadline = row[MGMT_COLS.ORDER_DEADLINE           - 1];
      var notified = String(row[MGMT_COLS.DEADLINE_NOTIFIED - 1] || '');
      var subject  = String(row[MGMT_COLS.SUBJECT           - 1] || '');
      var slipNo   = String(row[MGMT_COLS.ORDER_SLIP_NO     - 1] || '');
      var delivery = row[MGMT_COLS.DELIVERY_DATE            - 1];

      if (!orderNo || !deadline) return;
      if (status === CONFIG.STATUS.DELIVERED || status === CONFIG.STATUS.CANCELLED) return;

      var days = 9999;
      var deadlineStr = '';
      try {
        var d = deadline instanceof Date ? deadline : new Date(String(deadline).replace(/\//g,'-'));
        if (!isNaN(d.getTime())) {
          d.setHours(0, 0, 0, 0);
          days = Math.round((d - today) / (1000*60*60*24));
          deadlineStr = _toDateStr(d);
        }
      } catch(e) {}

      var deliveryStr = '';
      try {
        if (delivery) {
          var dv = delivery instanceof Date ? delivery : new Date(String(delivery).replace(/\//g,'-'));
          if (!isNaN(dv.getTime())) deliveryStr = _toDateStr(dv);
        }
      } catch(e) {}

      var lines = [];
      if (subject)     lines.push('\u9805\u76ee\u540d: '   + subject);
      if (slipNo)      lines.push('\u4f1d\u7968\u756a\u53f7: ' + slipNo);
      if (deadlineStr) lines.push('\u767a\u6ce8\u671f\u9650: ' + deadlineStr);
      if (deliveryStr) lines.push('\u7d0d\u671f: '         + deliveryStr);

      var item = {
        heading: client + '\u300c' + orderNo + '\u300d',
        lines:   lines,
        links:   _buildLinkItems(row, i + 2),
      };

      if (days < 0) {
        // 超過: 毎回通知（重複防止なし）
        item.lines.push('\u26a0\ufe0f ' + Math.abs(days) + '\u65e5 \u767a\u6ce8\u671f\u9650\u8d85\u904e');
        overdueItems.push(item);
      } else if (days <= 1) {
        if (notified.indexOf('urgent:' + todayStr) < 0) {
          item.lines.push('\u26a0\ufe0f \u660e\u65e5\u304c\u767a\u6ce8\u671f\u9650');
          urgentItems.push({ item: item, rowNum: i + 2, level: 'urgent' });
        }
      } else if (days <= 3) {
        if (notified.indexOf('3d:' + todayStr) < 0) {
          item.lines.push('\u26a0\ufe0f \u3042\u3068' + days + '\u65e5');
          remindItems.push({ item: item, rowNum: i + 2, level: '3d' });
        }
      } else if (days <= 7) {
        if (notified.indexOf('7d:' + todayStr) < 0) {
          item.lines.push('\u26a0\ufe0f \u3042\u3068' + days + '\u65e5');
          remindItems.push({ item: item, rowNum: i + 2, level: '7d' });
        }
      }
    });

    // DEADLINE_NOTIFIED に通知済フラグを書き込む（重複防止）
    function _writeNotified(entries) {
      entries.forEach(function(e) {
        try {
          var cell = sheet.getRange(e.rowNum, MGMT_COLS.DEADLINE_NOTIFIED);
          var cur  = String(cell.getValue() || '');
          cell.setValue(cur ? cur + ',' + e.level + ':' + todayStr : e.level + ':' + todayStr);
        } catch(ex) { Logger.log('[checkOrderDeadlines] write notified: ' + ex.message); }
      });
    }

    var urgentAlertItems = urgentItems.map(function(e){ return e.item; });
    var remindAlertItems = remindItems.map(function(e){ return e.item; });

    if (overdueItems.length     > 0 && MONITOR_CONFIG.ALERT_OVERDUE !== false) groups.push({ level: 'critical', title: '\ud83d\udea8 \u767a\u6ce8\u671f\u9650\u8d85\u904e',              items: overdueItems.slice(0,5)     });
    if (urgentAlertItems.length > 0) groups.push({ level: 'critical', title: '\ud83d\udd25 \u767a\u6ce8\u671f\u9650\uff08\u660e\u65e5\uff09', items: urgentAlertItems.slice(0,5) });
    if (remindAlertItems.length > 0) groups.push({ level: 'warning',  title: '\ud83d\udcc5 \u767a\u6ce8\u671f\u9650\u30ea\u30de\u30a4\u30f3\u30c9', items: remindAlertItems.slice(0,5) });

    if (urgentItems.length > 0) _writeNotified(urgentItems);
    if (remindItems.length > 0) _writeNotified(remindItems);

  } catch(e) {
    Logger.log('[MONITOR] checkOrderDeadlines: ' + e.message);
  }
  return groups;
}

// ============================================================
// 発注期限＋納期アラート スタンドアロン実行（6時間ごとトリガー用）
// ============================================================
function runDeadlineAlerts() {
  Logger.log('[DEADLINE] \u767a\u6ce8\u671f\u9650\u30c1\u30a7\u30c3\u30af\u958b\u59cb: ' + nowJST());
  _loadMonitorConfig();

  var alertGroups = [];
  if (MONITOR_CONFIG.ALERT_DELIVERY !== false) alertGroups = alertGroups.concat(_checkDeliveryDates());
  if (MONITOR_CONFIG.ALERT_DEADLINE !== false) alertGroups = alertGroups.concat(checkOrderDeadlines());

  if (alertGroups.length > 0) {
    var reportTitle = '\ud83d\udce6 \u7d0d\u671f\u30fb\u767a\u6ce8\u671f\u9650\u30a2\u30e9\u30fc\u30c8';
    _sendMonitorAlert(_alertGroupsToPlainText(alertGroups, reportTitle));
    _sendAlertEmailHtml(alertGroups, reportTitle);
    Logger.log('[DEADLINE] \u30a2\u30e9\u30fc\u30c8\u9001\u4fe1: ' + alertGroups.length + '\u30b0\u30eb\u30fc\u30d7');
  } else {
    Logger.log('[DEADLINE] \u767a\u6ce8\u671f\u9650\u30fb\u7d0d\u671f\u8d85\u904e\u306a\u3057');
  }
}

// ============================================================
// \u30c8\u30ea\u30ac\u30fc\u767b\u9332（\u5f8c\u65b9\u4e92\u6362\u30a8\u30a4\u30ea\u30a2\u30b9）
// 00_config.gs / 06 board api.gs \u304b\u3089\u547c\u3070\u308c\u308b
// ============================================================
function setupMonitoringTriggers() {
  var existing = ScriptApp.getProjectTriggers().map(function(t){ return t.getHandlerFunction(); });
  if (existing.indexOf('runAllMonitoring') < 0) {
    ScriptApp.newTrigger('runAllMonitoring').timeBased().atHour(9).everyDays(1).create();
    Logger.log('[MONITOR] runAllMonitoring \u30c8\u30ea\u30ac\u30fc\u767b\u9332\uff08\u6bce\u65e59\u6642\uff09');
  }
  if (existing.indexOf('runDeadlineAlerts') < 0) {
    ScriptApp.newTrigger('runDeadlineAlerts').timeBased().everyHours(6).create();
    Logger.log('[MONITOR] runDeadlineAlerts \u30c8\u30ea\u30ac\u30fc\u767b\u9332\uff086\u6642\u9593\u3054\u3068\uff09');
  }
}





// ============================================================
// ⑦ 見積セット チェックリスト 未提出・停滞アラート
//   CHECKLIST_STALE_DAYS 日以上 未提出 or 作成中 の項目を通知
//   機種コードごとにグルーピングして表示
// ============================================================

function _checkChecklistAlerts() {
  var groups = [];
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CHECKLIST_SHEET);
    if (!sheet || sheet.getLastRow() <= 1) return groups;

    var data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, CL.COL_COUNT).getValues()
                     .filter(function(r){ return String(r[CL.ITEM_ID - 1]).trim() !== ''; });
    var today     = new Date();
    var staleDays = MONITOR_CONFIG.CHECKLIST_STALE_DAYS || 3;

    // 未対応ステータス（提出済・承認済・却下はスキップ）
    var pendingStatuses = [CL_STATUS.PENDING, CL_STATUS.DRAFT];  // 未提出・作成中

    // 機種コード → stale アイテム一覧
    var modelMap = {};

    data.forEach(function(row) {
      var status    = String(row[CL.STATUS     - 1] || '').trim();
      var modelCode = String(row[CL.MODEL_CODE - 1] || '').trim();
      var itemName  = String(row[CL.ITEM_NAME  - 1] || '');
      var section   = String(row[CL.SECTION    - 1] || '');
      var submitTo  = String(row[CL.SUBMIT_TO  - 1] || '');
      var updatedAt = row[CL.UPDATED_AT        - 1];
      var pdfUrl    = String(row[CL.PDF_URL    - 1] || '');

      if (pendingStatuses.indexOf(status) < 0) return;
      if (!modelCode || !itemName) return;

      // 更新日からの経過日数
      var days = 0;
      try {
        var d = updatedAt instanceof Date
          ? updatedAt : new Date(String(updatedAt).replace(/\//g, '-'));
        if (!isNaN(d.getTime())) days = Math.floor((today - d) / (1000*60*60*24));
      } catch(e) {}

      if (days < staleDays) return;

      if (!modelMap[modelCode]) modelMap[modelCode] = [];
      modelMap[modelCode].push({ itemName: itemName, section: section, status: status, submitTo: submitTo, pdfUrl: pdfUrl, days: days });
    });

    var critItems = [], warnItems = [];

    Object.keys(modelMap).forEach(function(modelCode) {
      var staleItems = modelMap[modelCode];
      var maxDays    = staleItems.reduce(function(m, it){ return Math.max(m, it.days); }, 0);

      var lines = staleItems.slice(0, 8).map(function(it) {
        return '[' + it.section + '] ' + it.itemName
          + ' (' + it.status + ' • ' + it.days + '日未更新'
          + (it.submitTo ? ' • 提出先: ' + it.submitTo : '') + ')';
      });

      var links = staleItems
        .filter(function(it){ return !!it.pdfUrl; })
        .slice(0, 3)
        .map(function(it){ return { label: it.itemName, url: it.pdfUrl }; });

      var alertItem = {
        heading: '機種: ' + modelCode
          + '　(' + staleItems.length + '項目未対応 / 最大 ' + maxDays + '日経過)',
        lines:   lines,
        links:   links,
        _maxDays: maxDays,
      };

      if (maxDays >= staleDays * 3) {
        critItems.push(alertItem);
      } else {
        warnItems.push(alertItem);
      }
    });

    // 放置日数が長い順にソート
    critItems.sort(function(a, b){ return b._maxDays - a._maxDays; });
    warnItems.sort(function(a, b){ return b._maxDays - a._maxDays; });

    if (critItems.length > 0) {
      groups.push({
        level: 'critical',
        title: '📋 見積セット 放置警告（' + (staleDays * 3) + '日以上未対応）',
        items: critItems.slice(0, 5),
      });
    }
    if (warnItems.length > 0) {
      groups.push({
        level: 'warning',
        title: '📋 見積セット 未提出リマインド（' + staleDays + '日以上）',
        items: warnItems.slice(0, 5),
      });
    }
  } catch(e) {
    Logger.log('[MONITOR] _checkChecklistAlerts: ' + e.message);
  }
  return groups;
}
