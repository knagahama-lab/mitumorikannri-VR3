// ============================================================
// 51_quote_deadline.gs
// 単発見積（輸送費・金型費などのイニシャル費の請求）の「受注期限」と、期限切れアラームメール
//
//   ・件名がキーワード（輸送費・送料・金型・イニシャル・治具 など）に当たる見積を「単発」として自動登録し、
//     受注期限 = 提出日 + 既定日数（30日）を付ける。画面から期限変更・手動追加・対象外にできる。
//   ・毎朝（平日9時）チェック：受注期限を過ぎても注文書が紐づいていない見積があれば、
//     通知先メール（管理コンソールの通知先 NOTIFY_EMAILS、または本機能の宛先）の全員へアラームメール。
//     1件につき 期限翌日に1回 → 以後は「再通知の間隔」日ごと（既定3日）に、解決するまで送る。
//   ・注文書が紐づくと自動で「受注済み」になり監視を終了。「不要（失注）」にしても終了。
//   ・自動登録は運用開始日（初回実行日）以降に提出された見積だけ。それ以前の見積は画面の「候補」から手動で追加。
//
//   シート: 見積受注期限
// ============================================================

var QD_SHEET = '見積受注期限';
var QD_HEADERS = ['見積番号', '管理ID', '件名', '取引先', '提出日', '受注期限', '状態', '登録', '最終通知日', '通知回数', 'メモ', '更新日時'];
var QD_COL = { NO: 1, MGMT: 2, SUBJECT: 3, CLIENT: 4, SUBMIT: 5, DEADLINE: 6, STATE: 7, ORIGIN: 8, LAST: 9, COUNT: 10, MEMO: 11, UPDATED: 12 };
var QD_STATES = ['監視中', '受注済み', '不要（失注）'];
var QD_PROP = 'QUOTE_DEADLINE_CFG';
var QD_DEFAULT_KEYWORDS = ['輸送費', '運送', '送料', '運賃', '金型', 'イニシャル', '治具', '型代', '版代', '設計費', '修理'];

function _qdCfg() {
  var c = {};
  try { c = JSON.parse(PropertiesService.getScriptProperties().getProperty(QD_PROP) || '{}'); } catch (e) {}
  return {
    days: Number(c.days) || 30,
    repeatDays: c.repeatDays != null ? Math.max(0, Number(c.repeatDays)) : 3,
    keywords: Array.isArray(c.keywords) && c.keywords.length ? c.keywords : QD_DEFAULT_KEYWORDS.slice(),
    recipients: String(c.recipients || ''),
    hour: c.hour != null ? Number(c.hour) : 9,
    startDate: c.startDate || '',
    lastRun: c.lastRun || '', lastSent: c.lastSent || '',
  };
}
function _qdSaveCfg(c) { PropertiesService.getScriptProperties().setProperty(QD_PROP, JSON.stringify(c)); }

function _qdSheet() {
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName(QD_SHEET);
  if (!sh) {
    sh = ss.insertSheet(QD_SHEET);
    sh.getRange(1, 1, 1, QD_HEADERS.length).setValues([QD_HEADERS]).setBackground('#FFF3E0').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(2, 5, sh.getMaxRows() - 1, 2).setNumberFormat('@');   // 提出日・受注期限は文字列 yyyy/MM/dd
    sh.getRange(2, 9, sh.getMaxRows() - 1, 1).setNumberFormat('@');
    sh.getRange(2, 12, sh.getMaxRows() - 1, 1).setNumberFormat('@');
  }
  return sh;
}
function _qdRows() {
  var sh = _qdSheet(), last = sh.getLastRow();
  if (last <= 1) return [];
  return sh.getRange(2, 1, last - 1, QD_HEADERS.length).getValues().map(function (r, i) {
    return { row: i + 2, quoteNo: String(r[0] || '').trim(), mgmtId: String(r[1] || ''), subject: String(r[2] || ''), client: String(r[3] || ''),
             submitDate: _qdDate(r[4]), deadline: _qdDate(r[5]), state: String(r[6] || '監視中'), origin: String(r[7] || ''),
             lastNotified: _qdDate(r[8]), count: Number(r[9]) || 0, memo: String(r[10] || ''), updatedAt: String(r[11] || '') };
  }).filter(function (o) { return o.quoteNo; });
}
function _qdDate(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd');
  var m = String(v).replace(/-/g, '/').match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  return m ? m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2) : '';
}
function _qdAddDays(ymd, n) {
  var m = String(ymd || '').match(/^(\d{4})\/(\d{2})\/(\d{2})$/);
  var d = m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : new Date();
  d.setDate(d.getDate() + n);
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
}
function _qdDiffDays(a, b) { // b − a（日）
  var p = function (s) { var m = s.match(/^(\d{4})\/(\d{2})\/(\d{2})$/); return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])); };
  return Math.round((p(b) - p(a)) / 86400000);
}
function _qdToday() { return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd'); }
function _qdIsSingle(subject, keywords) {
  var s = String(subject || '').normalize('NFKC');
  return keywords.some(function (k) { return k && s.indexOf(String(k).normalize('NFKC')) >= 0; });
}

/** 見積の受注状況（32_quote_order_link.gs）: 見積番号 → summary */
function _qdLinks() {
  var r = apiQuoteOrderLinks();
  if (!r || !r.success) throw new Error((r && r.error) || '見積と注文の紐づけが取得できません');
  return r.byQuote || {};
}

/** 自動登録（運用開始日以降に提出された単発見積）＋受注済みの自動判定。変更件数を返す */
function _qdRefresh(links) {
  var cfg = _qdCfg();
  if (!cfg.startDate) { cfg.startDate = _qdToday(); _qdSaveCfg(cfg); }
  var sh = _qdSheet();
  var rows = _qdRows(), have = {};
  rows.forEach(function (o) { have[o.quoteNo] = o; });
  var now = nowJST(), added = 0, done = 0;

  // 受注済み・失注の自動判定
  rows.forEach(function (o) {
    if (o.state !== '監視中') return;
    var s = links[o.quoteNo];
    if (!s) return;
    if (s.state === 'ordered') {
      sh.getRange(o.row, QD_COL.STATE).setValue('受注済み');
      sh.getRange(o.row, QD_COL.MEMO).setValue((o.memo ? o.memo + ' / ' : '') + '注文 ' + (s.orders[0] ? s.orders[0].orderNo : '') + ' で受注（自動）');
      sh.getRange(o.row, QD_COL.UPDATED).setValue(now);
      done++;
    } else if (s.state === 'archived' || s.state === 'lost') {
      sh.getRange(o.row, QD_COL.STATE).setValue('不要（失注）');
      sh.getRange(o.row, QD_COL.MEMO).setValue((o.memo ? o.memo + ' / ' : '') + 'ステータス「' + s.status + '」のため終了（自動）');
      sh.getRange(o.row, QD_COL.UPDATED).setValue(now);
      done++;
    }
  });

  // 自動登録
  var add = [];
  Object.keys(links).forEach(function (no) {
    var s = links[no];
    if (have[no] || s.state !== 'waiting') return;
    var submit = _qdDate(s.submitDate);
    if (!submit || submit < cfg.startDate) return;
    if (!_qdIsSingle(s.subject, cfg.keywords)) return;
    add.push([no, s.quoteMgmtId || '', s.subject || '', s.client || '', submit, _qdAddDays(submit, cfg.days), '監視中', '自動', '', 0, '', now]);
  });
  if (add.length) { sh.getRange(sh.getLastRow() + 1, 1, add.length, QD_HEADERS.length).setValues(add); added = add.length; }
  return { added: added, done: done };
}

/** 期限切れ（受注期限 < 今日、監視中、未受注） */
function _qdOverdue(rows, today) {
  return rows.filter(function (o) { return o.state === '監視中' && o.deadline && o.deadline < today; })
    .map(function (o) { o.overDays = _qdDiffDays(o.deadline, today); return o; })
    .sort(function (a, b) { return b.overDays - a.overDays; });
}

function _qdRecipients(cfg) {
  var s = cfg.recipients || PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAILS') || '';
  return String(s).split(/[,\s、;]+/).map(function (x) { return x.trim(); }).filter(function (x) { return /@/.test(x); });
}

/**
 * 毎朝のチェック（時間主導トリガー）。force=true は「今すぐ送信」（再通知間隔を無視して期限切れ全件を送る）
 */
function checkQuoteOrderDeadlines(force) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { sent: false, error: '他の処理が実行中です' };
  try {
    var today = _qdToday();
    var dow = new Date().getDay();
    if (!force && (dow === 0 || dow === 6)) return { sent: false, skipped: '土日' };
    var links = _qdLinks();
    _qdRefresh(links);
    var cfg = _qdCfg();
    var rows = _qdRows();
    var over = _qdOverdue(rows, today);
    // 送る対象：初回（未通知）または前回通知から再通知間隔が過ぎたもの
    var due = force ? over : over.filter(function (o) {
      if (!o.lastNotified) return true;
      return cfg.repeatDays > 0 && _qdDiffDays(o.lastNotified, today) >= cfg.repeatDays;
    });
    cfg.lastRun = nowJST(); _qdSaveCfg(cfg);
    if (!due.length) return { sent: false, overdue: over.length, count: 0 };

    var to = _qdRecipients(cfg);
    if (!to.length) throw new Error('通知先メールが未設定です（受注期限の設定、または管理コンソールの通知先メール）');
    var url = _ntAppUrl ? _ntAppUrl() : '';
    var yen = function (n) { return n ? '¥' + Number(n).toLocaleString() : ''; };
    var lines = due.map(function (o) {
      var s = links[o.quoteNo] || {};
      return '・' + o.quoteNo + '　' + (o.client || s.client || '') + '　' + (o.subject || s.subject || '') +
        '\n　　受注期限 ' + o.deadline + '（' + o.overDays + '日超過）／提出 ' + (o.submitDate || '') + (s.quoteAmount ? '／' + yen(s.quoteAmount) : '') +
        (s.quotePdfUrl ? '\n　　見積書: ' + s.quotePdfUrl : '');
    });
    var subject = '【見積・注文管理】⚠ 受注期限切れの単発見積 ' + due.length + '件（注文書が未着）';
    var body = [
      '受注期限を過ぎても注文書が届いていない単発見積（輸送費・金型費などのイニシャル費の請求）があります。',
      '受注漏れがないか確認してください。',
      '',
    ].concat(lines).concat([
      '',
      '※ 注文書を取り込んで見積と紐づくと、自動で監視を終了します。',
      '※ 注文が来ない・不要になった見積は、画面の「⏰ 受注期限」で「不要（失注）」にすると通知が止まります。',
      '※ 解決するまで ' + (cfg.repeatDays || '-') + ' 日ごとに再通知します。',
    ]).concat(url ? ['', '▶ 見積・注文管理: ' + url] : []).join('\n');
    MailApp.sendEmail(to.join(','), subject, body, { name: '見積・注文管理システム' });

    var sh = _qdSheet(), now = nowJST();
    due.forEach(function (o) {
      sh.getRange(o.row, QD_COL.LAST).setValue(today);
      sh.getRange(o.row, QD_COL.COUNT).setValue(o.count + 1);
      sh.getRange(o.row, QD_COL.UPDATED).setValue(now);
    });
    cfg = _qdCfg(); cfg.lastSent = now + '（' + due.length + '件 → ' + to.length + '名）'; _qdSaveCfg(cfg);
    try { if (typeof _ntLog === 'function') _ntLog('受注期限切れ', { body: subject + '\n' + lines.slice(0, 10).join('\n') }); } catch (e) {}
    return { sent: true, count: due.length, overdue: over.length, recipients: to.length };
  } finally {
    lock.releaseLock();
  }
}
function checkQuoteOrderDeadlinesDaily() { return checkQuoteOrderDeadlines(false); }

/** トリガー（毎日 cfg.hour 時。土日は処理内でスキップ） */
function _qdEnsureTrigger(enabled) {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'checkQuoteOrderDeadlinesDaily') ScriptApp.deleteTrigger(t); });
  if (enabled) ScriptApp.newTrigger('checkQuoteOrderDeadlinesDaily').timeBased().everyDays(1).atHour(_qdCfg().hour).create();
}
function _qdEnabled() {
  return ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'checkQuoteOrderDeadlinesDaily'; });
}

// ------------------------------------------------------------
// API
// ------------------------------------------------------------
/** 一覧＋候補＋設定。初回は自動でトリガーを作る（毎朝のチェックを有効に） */
function apiQdList() {
  try {
    var cfg0 = _qdCfg();
    if (!cfg0.startDate) { cfg0.startDate = _qdToday(); _qdSaveCfg(cfg0); _qdEnsureTrigger(true); }
    var links = _qdLinks();
    _qdRefresh(links);
    var cfg = _qdCfg(), today = _qdToday();
    var rows = _qdRows().map(function (o) {
      var s = links[o.quoteNo] || {};
      o.quoteAmount = s.quoteAmount || 0; o.quotePdfUrl = s.quotePdfUrl || '';
      o.overDays = (o.state === '監視中' && o.deadline && o.deadline < today) ? _qdDiffDays(o.deadline, today) : 0;
      o.leftDays = (o.state === '監視中' && o.deadline && o.deadline >= today) ? _qdDiffDays(today, o.deadline) : null;
      return o;
    });
    var have = {}; rows.forEach(function (o) { have[o.quoteNo] = true; });
    // 候補：未登録の注文待ち見積（単発キーワードに当たるものを先に、提出日の新しい順）
    var cands = Object.keys(links).map(function (k) { return links[k]; })
      .filter(function (s) { return s.state === 'waiting' && !have[s.quoteNo]; })
      .map(function (s) {
        return { quoteNo: s.quoteNo, mgmtId: s.quoteMgmtId || '', subject: s.subject || '', client: s.client || '',
                 submitDate: _qdDate(s.submitDate), quoteAmount: s.quoteAmount || 0, single: _qdIsSingle(s.subject, cfg.keywords) };
      })
      .sort(function (a, b) { return (b.single - a.single) || String(b.submitDate).localeCompare(String(a.submitDate)); })
      .slice(0, 300);
    cfg.enabled = _qdEnabled();
    cfg.defaultRecipients = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAILS') || '';
    return { success: true, today: today, rows: rows, candidates: cands, cfg: cfg, states: QD_STATES };
  } catch (e) { return { success: false, error: e.message }; }
}

/** 追加・変更 p: { quoteNo, deadline?, state?, memo?, add?:{mgmtId,subject,client,submitDate} } */
function apiQdSave(p) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    p = p || {};
    var no = String(p.quoteNo || '').trim();
    if (!no) return { success: false, error: '見積番号がありません' };
    var dl = p.deadline != null ? _qdDate(p.deadline) : null;
    if (p.deadline != null && p.deadline !== '' && !dl) return { success: false, error: '受注期限の日付が正しくありません' };
    if (p.state != null && QD_STATES.indexOf(p.state) < 0) return { success: false, error: '状態が正しくありません' };
    var sh = _qdSheet(), now = nowJST();
    var cur = _qdRows().filter(function (o) { return o.quoteNo === no; })[0];
    if (!cur) {
      var a = p.add || {}, cfg = _qdCfg();
      var submit = _qdDate(a.submitDate) || _qdToday();
      sh.getRange(sh.getLastRow() + 1, 1, 1, QD_HEADERS.length).setValues([[no, a.mgmtId || '', a.subject || '', a.client || '', submit,
        dl || _qdAddDays(submit, cfg.days), p.state || '監視中', '手動', '', 0, p.memo || '', now]]);
      return { success: true, added: true };
    }
    if (dl !== null) {
      sh.getRange(cur.row, QD_COL.DEADLINE).setValue(dl);
      if (dl !== cur.deadline) { sh.getRange(cur.row, QD_COL.LAST).setValue(''); sh.getRange(cur.row, QD_COL.COUNT).setValue(0); } // 期限を変えたら通知をやり直す
    }
    if (p.state != null) sh.getRange(cur.row, QD_COL.STATE).setValue(p.state);
    if (p.memo != null) sh.getRange(cur.row, QD_COL.MEMO).setValue(p.memo);
    sh.getRange(cur.row, QD_COL.UPDATED).setValue(now);
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
  finally { try { lock.releaseLock(); } catch (e2) {} }
}

/** 設定 p: { enabled, days, repeatDays, keywords[], recipients, hour } */
function apiQdCfgSave(p) {
  try {
    p = p || {};
    var c = {}; try { c = JSON.parse(PropertiesService.getScriptProperties().getProperty(QD_PROP) || '{}'); } catch (e) {}
    if (p.days != null) c.days = Math.max(1, Math.min(365, Number(p.days) || 30));
    if (p.repeatDays != null) c.repeatDays = Math.max(0, Math.min(30, Number(p.repeatDays) || 0));
    if (p.keywords != null) c.keywords = (Array.isArray(p.keywords) ? p.keywords : String(p.keywords).split(/[,、\n]+/)).map(function (s) { return String(s).trim(); }).filter(Boolean);
    if (p.recipients != null) {
      var bad = String(p.recipients).split(/[,\s、;]+/).filter(function (x) { return x && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x); });
      if (bad.length) return { success: false, error: 'メールアドレスが正しくありません: ' + bad.join(', ') };
      c.recipients = String(p.recipients).trim();
    }
    if (p.hour != null) c.hour = Math.max(0, Math.min(23, Number(p.hour)));
    if (!c.startDate) c.startDate = _qdToday();
    _qdSaveCfg(c);
    if (p.enabled != null || p.hour != null) _qdEnsureTrigger(p.enabled != null ? !!p.enabled : _qdEnabled());
    return apiQdList();
  } catch (e) { return { success: false, error: e.message }; }
}

/** 今すぐチェックして送信（期限切れ全件） */
function apiQdRunNow() {
  try { var r = checkQuoteOrderDeadlines(true); return { success: true, sent: !!r.sent, count: r.count || 0, overdue: r.overdue || 0, recipients: r.recipients || 0 }; }
  catch (e) { return { success: false, error: e.message }; }
}

/** processDriveImports（5分ごと）から呼ぶ：初回だけ運用開始日を記録し、毎朝のチェックを有効にする */
function _qdAutoInit() {
  try {
    var c = _qdCfg();
    if (c.startDate) return;
    c.startDate = _qdToday();
    _qdSaveCfg(c);
    _qdEnsureTrigger(true);
    Logger.log('[受注期限] 毎朝のチェックを有効にしました（運用開始日 ' + c.startDate + '）');
  } catch (e) { Logger.log('[_qdAutoInit] ' + e.message); }
}
