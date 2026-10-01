// ============================================================
// 39_notify.gs
// 注文書の検出・通知（客先部品コードマスタを活用）
//
//   注文書の取込時（02 _saveOrderData から cpmNotifyOrder を呼ぶ）に
//     🆕 新規品番   … マスタに未登録の客先部品コードを含む注文 → Chat／メール通知＋画面の通知
//     📦 見積→受注 … 弊社見積に紐づいた注文（見積書の注文書が発行された） → 画面の通知
//     ⚠ 要確認     … 部品コードを読み取れなかった注文 → 画面の通知
//     ✏ 差し替え／✖ キャンセル → Chat／メール通知＋画面の通知
//   外部通知（Chat／メール）のモードはスクリプトプロパティ ORDER_NOTIFY_MODE
//     newcode（既定：新規品番・差し替え・キャンセルのみ）／ all（すべての注文）／ off（送らない）
// ============================================================

var NOTIFY_SHEET   = '通知ログ';
var NOTIFY_HEADERS = ['通知ID','日時','種別','管理ID','注文No','見積No','見積管理ID','取引先','内容','既読'];
var NT = { ID:0, AT:1, TYPE:2, MGMT:3, ORDER:4, QUOTE:5, QUOTE_ID:6, CLIENT:7, BODY:8, READ:9 };

function _ntSheet() {
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName(NOTIFY_SHEET);
  if (sh) return sh;
  sh = ss.insertSheet(NOTIFY_SHEET);
  sh.getRange(1, 1, 1, NOTIFY_HEADERS.length).setValues([NOTIFY_HEADERS]).setBackground('#FEF3C7').setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange(1, 1, sh.getMaxRows(), NOTIFY_HEADERS.length).setNumberFormat('@');
  return sh;
}
function _ntMode() {
  var m = PropertiesService.getScriptProperties().getProperty('ORDER_NOTIFY_MODE') || 'newcode';
  return ['newcode', 'all', 'off'].indexOf(m) >= 0 ? m : 'newcode';
}
function _ntLog(type, o) {
  var id = 'NT-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 900 + 100);
  _ntSheet().appendRow([id, nowJST(), type, o.mgmtId || '', o.orderNo || '', o.quoteNo || '', o.quoteMgmtId || '', o.client || '', o.body || '', '']);
}
function _ntAppUrl() { try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } }

/** Chat（Webhook）とメール（管理コンソールの通知先メール）へ送る */
function _ntSendExternal(title, lines) {
  var props = PropertiesService.getScriptProperties();
  var text = title + '\n' + lines.join('\n') + (_ntAppUrl() ? '\n\n▶ 見積・注文管理: ' + _ntAppUrl() : '');
  var hook = props.getProperty('CHAT_WEBHOOK_URL') || (typeof _getChatWebhookUrl === 'function' ? _getChatWebhookUrl() : '');
  if (hook) {
    try { UrlFetchApp.fetch(hook, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ text: text }), muteHttpExceptions: true }); }
    catch (e) { Logger.log('[notify chat] ' + e.message); }
  }
  var to = String(props.getProperty('NOTIFY_EMAILS') || '').split(',').map(function(s) { return s.trim(); }).filter(Boolean);
  if (to.length) {
    try { MailApp.sendEmail(to.join(','), '【見積・注文管理】' + title.replace(/^[^\s]+\s*/, ''), text); }
    catch (e) { Logger.log('[notify mail] ' + e.message); }
  }
}

/**
 * ★ 注文書取込後に呼ぶ（02_ocr_and_processing.gs _saveOrderData）
 * @param {string} mgmtId  注文の管理ID
 * @param {string} action  'new' | 'revision' | 'cancellation'
 * @param {Object} cpm     cpmOnOrderSaved の戻り値（results: [{type:'new'|'repeat', code, name, qty, price, prev}]）
 * @param {Object} ocr     OCR結果
 */
function cpmNotifyOrder(mgmtId, action, cpm, ocr) {
  var mg = getAllMgmtData().map(_rowToObject).filter(function(o) { return o.id === String(mgmtId); })[0] || {};
  var orderNo = mg.orderNo || (ocr && ocr.documentNo) || '';
  var client = mg.client || (ocr && ocr.clientName) || '';
  var head = ['取引先: ' + (client || '（不明）'), '注文No: ' + orderNo + (mg.orderDate ? '（発注日 ' + mg.orderDate + '）' : ''),
              mg.subject ? '件名: ' + mg.subject : '', mg.orderPdfUrl ? 'PDF: ' + mg.orderPdfUrl : ''].filter(String);
  var yen = function(n) { return '¥' + Number(n || 0).toLocaleString(); };
  var results = (cpm && cpm.results) || [];
  var newOnes = results.filter(function(r) { return r.type === 'new'; });
  var lineCount = ((ocr && ocr.lineItems) || []).length;
  var noCode = Math.max(0, lineCount - results.length);

  // 見積に紐づいた注文（見積書の注文書が発行された）
  var quotes = [], links = null;
  try {
    links = apiQuoteOrderLinks();
    var bo = links.success && links.byOrder[orderNo];
    if (bo) quotes = bo.quotes || [];
  } catch (e) { Logger.log('[notify links] ' + e.message); }

  // 🎉 初回注文：その見積に初めて届いた注文／事前登録した部品コードで初めての注文
  var firstQuotes = quotes.filter(function(q) {
    var s = links && links.byQuote && links.byQuote[q.quoteNo];
    return s && s.orders && s.orders.length && s.orders[0].orderNo === orderNo;
  });
  var firstCodes = results.filter(function(r) { return r.type === 'repeat' && r.firstOrder; });
  var firstLines = firstQuotes.map(function(q) {
    return '・見積 ' + q.quoteNo + '（' + (q.submitDate || '提出日不明') + ' 提出' + (q.leadDays != null ? '／提出から' + q.leadDays + '日' : '') + '）';
  }).concat(firstCodes.map(function(r) {
    return '・部品コード ' + r.code + '　' + r.name + '　' + r.qty + '個 ' + yen(r.price) + (r.quoteNo ? '（見積 ' + r.quoteNo + '）' : '');
  }));

  var external = null;
  if (action === 'cancellation' || action === 'revision') {
    var label = action === 'cancellation' ? '✖ 注文書キャンセル' : '✏ 注文書差し替え';
    _ntLog(action === 'cancellation' ? 'キャンセル' : '差し替え', { mgmtId: mgmtId, orderNo: orderNo, client: client, body: label + '：' + orderNo + ' ' + (mg.subject || '') });
    external = { title: label + 'を受領しました', lines: head };
  }
  if (newOnes.length) {
    var nl = newOnes.map(function(r) { return '・' + r.code + '　' + r.name + '　' + r.qty + '個 ' + yen(r.price) + (r.quoteNo ? '（見積 ' + r.quoteNo + '）' : '（弊社見積 未紐づけ）'); });
    _ntLog('新規品番', { mgmtId: mgmtId, orderNo: orderNo, client: client,
      body: '🆕 新規品番 ' + newOnes.length + '件を含む注文 ' + orderNo + '\n' + nl.join('\n') });
    if (!external) external = { title: '🆕 新規品番（客先部品コード未登録）の注文書を受領しました', lines: head.concat(['', '【新規品番 ' + newOnes.length + '件】']).concat(nl) };
    else external.lines = external.lines.concat(['', '【新規品番】']).concat(nl);
  }
  quotes.forEach(function(q) {
    var first = firstQuotes.indexOf(q) >= 0;
    _ntLog('見積→受注', { mgmtId: mgmtId, orderNo: orderNo, quoteNo: q.quoteNo, quoteMgmtId: q.quoteMgmtId, client: client,
      body: (first ? '🎉 初回注文：' : '📦 ') + '見積 ' + q.quoteNo + '（' + (q.submitDate || '提出日不明') + ' 提出）の注文書 ' + orderNo + ' を受領' + (q.leadDays != null ? '（提出から' + q.leadDays + '日）' : '') });
  });
  if (firstLines.length && action === 'new') {
    if (!external) external = { title: '🎉 初回注文書を受領しました', lines: head.concat(['', '【初回注文】']).concat(firstLines) };
    else {
      if (newOnes.length) external.title = '🎉 初回注文書を受領しました（新規品番あり）';
      external.lines = external.lines.concat(['', '【初回注文】']).concat(firstLines);
    }
  }
  if (external && quotes.length) external.lines.push('', '【紐づいた弊社見積】', quotes.map(function(q) { return '・' + q.quoteNo; }).join(' '));
  if (noCode && action === 'new') {
    _ntLog('要確認', { mgmtId: mgmtId, orderNo: orderNo, client: client, body: '⚠ 注文 ' + orderNo + ' の ' + noCode + '行で客先部品コードを読み取れませんでした（新規品番か確認してください）' });
  }

  var mode = _ntMode();
  if (mode === 'off') return;
  if (mode === 'all') {
    if (external) _ntSendExternal(external.title, external.lines);                     // 新規品番等は詳しい内容で
    else if (typeof _sendChatNotification === 'function') _sendChatNotification(mgmtId, 'order', action); // それ以外は従来の通知
    return;
  }
  if (external) _ntSendExternal(external.title, external.lines); // newcode：新規品番・初回注文・差し替え・キャンセル
}

// ============================================================
// 注文待ちリマインド（毎週：提出から一定日数、注文書が届いていない見積の一覧）
//   設定はスクリプトプロパティ REMIND_WAITING（JSON: {days, weekday, hour}）
// ============================================================
var REMIND_PROP = 'REMIND_WAITING';
var REMIND_WEEKDAYS = ['SUNDAY', 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY'];
var REMIND_WEEKDAY_JA = ['日', '月', '火', '水', '木', '金', '土'];

function _remindCfg() {
  var c = {};
  try { c = JSON.parse(PropertiesService.getScriptProperties().getProperty(REMIND_PROP) || '{}'); } catch (e) {}
  return { days: Number(c.days) || 30, weekday: c.weekday != null ? Number(c.weekday) : 1, hour: c.hour != null ? Number(c.hour) : 8,
           lastRun: c.lastRun || '', lastCount: c.lastCount != null ? c.lastCount : '' };
}

/** 注文待ちの見積（提出から cfg.days 日以上、受注・アーカイブ・失注を除く） */
function _remindWaitingQuotes(days) {
  var links = apiQuoteOrderLinks();
  if (!links.success) throw new Error(links.error);
  return Object.keys(links.byQuote).map(function(k) { return links.byQuote[k]; })
    .filter(function(s) { return s.state === 'waiting' && s.waitingDays != null && s.waitingDays >= days; })
    .sort(function(a, b) { return b.waitingDays - a.waitingDays; });
}

/** 時間主導トリガーから呼ぶ（手動の「今すぐ送信」からも使う） */
function weeklyQuoteWaitingReminder() {
  var cfg = _remindCfg();
  var list = _remindWaitingQuotes(cfg.days);
  var props = PropertiesService.getScriptProperties();
  var saved = {}; try { saved = JSON.parse(props.getProperty(REMIND_PROP) || '{}'); } catch (e) {}
  saved.lastRun = nowJST(); saved.lastCount = list.length;
  props.setProperty(REMIND_PROP, JSON.stringify(saved));
  if (!list.length) return { sent: false, count: 0 };
  var lines = list.slice(0, 60).map(function(s) {
    return '・' + s.waitingDays + '日　' + s.quoteNo + '　' + (s.client || '') + '　' + (s.subject || '').substring(0, 40) +
      '（' + (s.submitDate || '') + ' 提出）' + (s.candidates && s.candidates.length ? '　🔎候補あり' : '');
  });
  if (list.length > 60) lines.push('…ほか ' + (list.length - 60) + '件');
  var title = '⏳ 注文待ちの見積 ' + list.length + '件（提出から' + cfg.days + '日以上）';
  _ntSendExternal(title, ['注文書がまだ届いていない見積の一覧です（長く待っている順）。', '不要になった見積はステータスを「ボツ」「旧見積」にすると次回から除外されます。', ''].concat(lines));
  _ntLog('注文待ち', { body: title + '\n' + lines.slice(0, 15).join('\n') + (list.length > 15 ? '\n…ほか ' + (list.length - 15) + '件' : '') });
  return { sent: true, count: list.length };
}

function apiRemindGet() {
  try {
    var cfg = _remindCfg();
    cfg.enabled = ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === 'weeklyQuoteWaitingReminder'; });
    cfg.weekdayJa = REMIND_WEEKDAY_JA[cfg.weekday];
    cfg.preview = _remindWaitingQuotes(cfg.days).length;
    return Object.assign({ success: true }, cfg);
  } catch (e) { return { success: false, error: e.message }; }
}

/** p: { enabled, days, weekday(0-6), hour(0-23) } */
function apiRemindSet(p) {
  try {
    p = p || {};
    var props = PropertiesService.getScriptProperties();
    var saved = {}; try { saved = JSON.parse(props.getProperty(REMIND_PROP) || '{}'); } catch (e) {}
    if (p.days != null) saved.days = Math.max(1, Number(p.days) || 30);
    if (p.weekday != null) saved.weekday = Math.min(6, Math.max(0, Number(p.weekday)));
    if (p.hour != null) saved.hour = Math.min(23, Math.max(0, Number(p.hour)));
    props.setProperty(REMIND_PROP, JSON.stringify(saved));
    ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() === 'weeklyQuoteWaitingReminder') ScriptApp.deleteTrigger(t); });
    if (p.enabled) {
      var c = _remindCfg();
      ScriptApp.newTrigger('weeklyQuoteWaitingReminder').timeBased()
        .onWeekDay(ScriptApp.WeekDay[REMIND_WEEKDAYS[c.weekday]]).atHour(c.hour).create();
    }
    return apiRemindGet();
  } catch (e) { return { success: false, error: e.message }; }
}

function apiRemindRunNow() {
  try { var r = weeklyQuoteWaitingReminder(); return { success: true, sent: r.sent, count: r.count }; }
  catch (e) { return { success: false, error: e.message }; }
}

// ============================================================
// 画面の通知（ベル）
// ============================================================
function apiNotifyList(p) {
  try {
    _ensureArchiveStatusValidation();
    var me = (typeof _ofUser === 'function') ? _ofUser() : '';
    var sh = _ntSheet(); var last = sh.getLastRow();
    var rows = last > 1 ? sh.getRange(Math.max(2, last - 499), 1, Math.min(500, last - 1), NOTIFY_HEADERS.length).getValues() : [];
    var items = rows.reverse().map(function(r) {
      var readers = String(r[NT.READ] || '').split(',');
      return { id: String(r[NT.ID]), at: String(r[NT.AT]), type: String(r[NT.TYPE]), mgmtId: String(r[NT.MGMT]), orderNo: String(r[NT.ORDER]),
               quoteNo: String(r[NT.QUOTE]), quoteMgmtId: String(r[NT.QUOTE_ID]), client: String(r[NT.CLIENT]), body: String(r[NT.BODY]),
               read: readers.indexOf(me) >= 0 };
    }).slice(0, 200);
    return { success: true, items: items, unread: items.filter(function(i) { return !i.read; }).length, mode: _ntMode() };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function apiNotifyMarkRead(p) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var me = (typeof _ofUser === 'function') ? _ofUser() : '';
    var ids = (p && p.ids) || [];
    var sh = _ntSheet(); var last = sh.getLastRow();
    if (last <= 1) return { success: true };
    var rng = sh.getRange(2, 1, last - 1, NOTIFY_HEADERS.length);
    var vals = rng.getValues(); var changed = false;
    vals.forEach(function(r) {
      if (!(p && p.all) && ids.indexOf(String(r[NT.ID])) < 0) return;
      var readers = String(r[NT.READ] || '').split(',').filter(Boolean);
      if (readers.indexOf(me) < 0) { readers.push(me); r[NT.READ] = readers.join(','); changed = true; }
    });
    if (changed) sh.getRange(2, NT.READ + 1, vals.length, 1).setValues(vals.map(function(r) { return [r[NT.READ]]; }));
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function apiNotifySetMode(p) {
  var m = String((p && p.mode) || '');
  if (['newcode', 'all', 'off'].indexOf(m) < 0) return { success: false, error: '不正なモード' };
  PropertiesService.getScriptProperties().setProperty('ORDER_NOTIFY_MODE', m);
  return { success: true, mode: m };
}

/** 既存の見積台帳シートのステータス入力規則に ボツ／旧見積 を追加（1回だけ） */
function _ensureArchiveStatusValidation() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('LEDGER_STATUS_ARCHIVE_V1') === 'done') return;
  try {
    var sheet = getSpreadsheet().getSheetByName(CONFIG.SHEET_LEDGER);
    if (sheet) {
      var rows = Math.max(sheet.getMaxRows() - 1, 1);
      sheet.getRange(2, LEDGER_COLS.STATUS, rows, 1).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(['作成予定', '作成中', '送信済み', '受注済み', 'キャンセル'].concat(ARCHIVE_STATUSES), true)
          .setAllowInvalid(true).build());
    }
    props.setProperty('LEDGER_STATUS_ARCHIVE_V1', 'done');
  } catch (e) { Logger.log('[_ensureArchiveStatusValidation] ' + e.message); }
}
