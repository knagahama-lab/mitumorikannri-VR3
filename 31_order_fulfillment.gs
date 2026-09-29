// ============================================================
// 31_order_fulfillment.gs
// 受注進捗ボード（CO-NECT 受注機能相当）
//
//   ① 受注ステータス管理 … 未確認→対応中→生産中→出荷準備→出荷済→納品完了
//   ② 納期回答・出荷通知 … 納期回答日/出荷日/送り状を登録し、取引先宛メールを Gmail 下書き作成
//   ③ 社内メモ共有       … 案件ごとの社内コメント（投稿者・日時つき履歴）
//   ④ 最低注文数・ロット … 品名キーワードごとに MOQ / ロット単位を持ち、注文明細をチェック
//   ⑤ CSV一括更新・出力  … 案件＋受注進捗 / 発注条件マスタを CSV で一括更新（差分プレビュー→適用）
//
// 管理シートの列構成は変更しない。受注後の進捗は「受注進捗」シートに管理IDキーで保持する。
// 納期回答日は管理シートの「納期」列へ、納品完了は管理ステータス「納品済み」へ同期する。
// ============================================================

var OF_SHEET = {
  PROGRESS: '受注進捗',
  MEMO:     '社内メモ',
  LOT:      '発注条件マスタ',
};

var OF_PROGRESS_HEADERS = ['管理ID','進捗ステータス','納期回答日','出荷予定日','出荷日','運送会社','送り状番号',
                           '通知先メール','最終通知日時','最終通知種別','更新者','更新日時'];
var OF_P = { ID:0, STATUS:1, ANSWER_DATE:2, SHIP_PLAN:3, SHIP_DATE:4, CARRIER:5, TRACKING:6,
             NOTIFY_TO:7, NOTIFIED_AT:8, NOTIFIED_TYPE:9, UPDATED_BY:10, UPDATED_AT:11 };

var OF_MEMO_HEADERS = ['メモID','管理ID','投稿者','本文','投稿日時'];
var OF_LOT_HEADERS  = ['品名キーワード','最低注文数','ロット単位','備考','更新日時'];

var OF_STATUSES = ['未確認','対応中','生産中','出荷準備','出荷済','納品完了'];

// 受注進捗 → 管理シートのステータス同期
var OF_TO_MGMT_STATUS = { '出荷済': '受注済み', '納品完了': '納品済み' };

var OF_CLIENT_EMAIL_PROP = 'OF_CLIENT_EMAILS';

// ============================================================
// シート準備
// ============================================================
function _ofSheet(name) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  var headers = name === OF_SHEET.PROGRESS ? OF_PROGRESS_HEADERS
              : name === OF_SHEET.MEMO     ? OF_MEMO_HEADERS
              : OF_LOT_HEADERS;
  sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setBackground('#E0F2FE').setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);
  // 日付文字列を勝手に Date 変換させない
  sheet.getRange(1, 1, sheet.getMaxRows(), headers.length).setNumberFormat('@');
  return sheet;
}

function _ofValues(sheet, width) {
  var last = sheet.getLastRow();
  if (last <= 1) return [];
  return sheet.getRange(2, 1, last - 1, width).getValues();
}

function _ofDate(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd');
  var s = String(v).trim().replace(/-/g, '/');
  var m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (!m) return s;
  return m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2);
}

function _ofUser() {
  try { return Session.getActiveUser().getEmail() || '(不明)'; } catch (e) { return '(不明)'; }
}

function _ofProgressToObj(r) {
  return {
    mgmtId:       String(r[OF_P.ID] || ''),
    progress:     String(r[OF_P.STATUS] || '') || OF_STATUSES[0],
    answerDate:   _ofDate(r[OF_P.ANSWER_DATE]),
    shipPlanDate: _ofDate(r[OF_P.SHIP_PLAN]),
    shipDate:     _ofDate(r[OF_P.SHIP_DATE]),
    carrier:      String(r[OF_P.CARRIER] || ''),
    trackingNo:   String(r[OF_P.TRACKING] || ''),
    notifyTo:     String(r[OF_P.NOTIFY_TO] || ''),
    notifiedAt:   String(r[OF_P.NOTIFIED_AT] || ''),
    notifiedType: String(r[OF_P.NOTIFIED_TYPE] || ''),
    updatedBy:    String(r[OF_P.UPDATED_BY] || ''),
    updatedAt:    String(r[OF_P.UPDATED_AT] || ''),
  };
}

function _ofProgressMap() {
  var map = {};
  _ofValues(_ofSheet(OF_SHEET.PROGRESS), OF_PROGRESS_HEADERS.length).forEach(function(r) {
    var id = String(r[OF_P.ID] || '');
    if (id) map[id] = _ofProgressToObj(r);
  });
  return map;
}

// ============================================================
// ④ 最低注文数・ロット
// ============================================================
function _ofLotRules() {
  return _ofValues(_ofSheet(OF_SHEET.LOT), OF_LOT_HEADERS.length)
    .filter(function(r) { return String(r[0]).trim(); })
    .map(function(r) {
      return {
        keyword: String(r[0]).trim(),
        moq:     Number(r[1]) || 0,
        lot:     Number(r[2]) || 0,
        note:    String(r[3] || ''),
      };
    });
}

/** 品名＋仕様にキーワードを含む最長一致ルールを返す */
function _ofFindLotRule(rules, itemName, spec) {
  var text = normalizeText(String(itemName || '') + ' ' + String(spec || ''));
  var hit = null;
  rules.forEach(function(rule) {
    var kw = normalizeText(rule.keyword);
    if (kw && text.indexOf(kw) >= 0 && (!hit || kw.length > normalizeText(hit.keyword).length)) hit = rule;
  });
  return hit;
}

/** 数量をルールで検査。問題なければ null、あればメッセージ */
function _ofCheckQty(rule, qty) {
  if (!rule) return null;
  var q = Number(qty) || 0;
  var msgs = [];
  if (rule.moq && q < rule.moq) msgs.push('最低注文数 ' + rule.moq + ' 未満');
  if (rule.lot && q % rule.lot !== 0) {
    var up = Math.ceil(q / rule.lot) * rule.lot;
    msgs.push('ロット ' + rule.lot + ' 単位外（推奨 ' + up + '）');
  }
  return msgs.length ? msgs.join(' / ') : null;
}

function apiOfLotRuleList() {
  try { return { success: true, items: _ofLotRules() }; }
  catch (e) { return { success: false, error: e.message }; }
}

/** p.items を丸ごと保存（mode:'merge' ならキーワード単位で上書き・追加） */
function apiOfLotRuleSave(p) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var incoming = (p && p.items) || [];
    var byKey = {};
    var order = [];
    if (p && p.mode === 'merge') {
      _ofLotRules().forEach(function(r) { byKey[r.keyword] = r; order.push(r.keyword); });
    }
    incoming.forEach(function(r) {
      var kw = String(r.keyword || '').trim();
      if (!kw) return;
      if (!byKey[kw]) order.push(kw);
      byKey[kw] = { keyword: kw, moq: Number(r.moq) || 0, lot: Number(r.lot) || 0, note: String(r.note || '') };
    });
    var sheet = _ofSheet(OF_SHEET.LOT);
    var last  = sheet.getLastRow();
    if (last > 1) sheet.getRange(2, 1, last - 1, OF_LOT_HEADERS.length).clearContent();
    var now  = nowJST();
    var rows = order.map(function(k) { var r = byKey[k]; return [r.keyword, r.moq || '', r.lot || '', r.note, now]; });
    if (rows.length) sheet.getRange(2, 1, rows.length, OF_LOT_HEADERS.length).setValues(rows);
    return { success: true, count: rows.length };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// ① 受注ステータス 一覧／保存
// ============================================================
function apiOfList() {
  try {
    var base = _apiGetAll();
    if (!base.success) return base;
    var progress = _ofProgressMap();
    var rules    = _ofLotRules();

    // 注文明細のロット警告件数
    var lotWarn = {};
    if (rules.length) {
      var os = getSpreadsheet().getSheetByName(CONFIG.SHEET_ORDERS);
      if (os && os.getLastRow() > 1) {
        os.getRange(2, 1, os.getLastRow() - 1, ORDER_COLS.QTY).getValues().forEach(function(r) {
          var id = String(r[ORDER_COLS.MGMT_ID - 1] || '');
          if (!id) return;
          var rule = _ofFindLotRule(rules, r[ORDER_COLS.ITEM_NAME - 1], r[ORDER_COLS.SPEC - 1]);
          if (_ofCheckQty(rule, r[ORDER_COLS.QTY - 1])) lotWarn[id] = (lotWarn[id] || 0) + 1;
        });
      }
    }

    // 社内メモ件数
    var memoCount = {};
    _ofValues(_ofSheet(OF_SHEET.MEMO), OF_MEMO_HEADERS.length).forEach(function(r) {
      var id = String(r[1] || '');
      if (id) memoCount[id] = (memoCount[id] || 0) + 1;
    });

    var items = base.items.map(function(it) {
      var pg = progress[it.id] || _ofProgressToObj([it.id]);
      if (it.status === CONFIG.STATUS.DELIVERED && !progress[it.id]) pg.progress = '納品完了';
      return {
        id: it.id, orderNo: it.orderNo, quoteNo: it.quoteNo, subject: it.subject, client: it.client,
        status: it.status, orderType: it.orderType, orderDate: it.orderDate, deliveryDate: it.deliveryDate,
        orderAmount: it.orderAmount, modelCode: it.modelCode, assignee: it.assignee, orderPdfUrl: it.orderPdfUrl,
        progress: pg.progress, answerDate: pg.answerDate, shipPlanDate: pg.shipPlanDate, shipDate: pg.shipDate,
        carrier: pg.carrier, trackingNo: pg.trackingNo, notifyTo: pg.notifyTo,
        notifiedAt: pg.notifiedAt, notifiedType: pg.notifiedType, updatedBy: pg.updatedBy, updatedAt: pg.updatedAt,
        lotWarnings: lotWarn[it.id] || 0, memoCount: memoCount[it.id] || 0,
      };
    });
    return { success: true, statuses: OF_STATUSES, items: items };
  } catch (e) {
    Logger.log('[apiOfList] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

/** 受注進捗を1件 upsert（呼び出し元でロック取得済みの前提で使う内部版） */
function _ofUpsertProgress(sheet, p, user) {
  var data = _ofValues(sheet, OF_PROGRESS_HEADERS.length);
  var idx  = -1;
  for (var i = 0; i < data.length; i++) { if (String(data[i][OF_P.ID]) === String(p.mgmtId)) { idx = i; break; } }
  var row = idx >= 0 ? data[idx].slice() : [p.mgmtId, OF_STATUSES[0], '', '', '', '', '', '', '', '', '', ''];
  var before = _ofProgressToObj(row);

  if (p.progress !== undefined) {
    if (OF_STATUSES.indexOf(p.progress) < 0) throw new Error('無効な進捗ステータス: ' + p.progress);
    row[OF_P.STATUS] = p.progress;
  }
  if (p.answerDate   !== undefined) row[OF_P.ANSWER_DATE] = _ofDate(p.answerDate);
  if (p.shipPlanDate !== undefined) row[OF_P.SHIP_PLAN]   = _ofDate(p.shipPlanDate);
  if (p.shipDate     !== undefined) row[OF_P.SHIP_DATE]   = _ofDate(p.shipDate);
  if (p.carrier      !== undefined) row[OF_P.CARRIER]     = String(p.carrier);
  if (p.trackingNo   !== undefined) row[OF_P.TRACKING]    = String(p.trackingNo);
  if (p.notifyTo     !== undefined) row[OF_P.NOTIFY_TO]   = String(p.notifyTo);
  if (p.notifiedAt   !== undefined) row[OF_P.NOTIFIED_AT] = String(p.notifiedAt);
  if (p.notifiedType !== undefined) row[OF_P.NOTIFIED_TYPE] = String(p.notifiedType);
  // 出荷日が入って未出荷ステータスなら自動で「出荷済」へ
  if (p.progress === undefined && row[OF_P.SHIP_DATE] && OF_STATUSES.indexOf(row[OF_P.STATUS]) < OF_STATUSES.indexOf('出荷済')) {
    row[OF_P.STATUS] = '出荷済';
  }
  row[OF_P.UPDATED_BY] = user;
  row[OF_P.UPDATED_AT] = nowJST();

  if (idx >= 0) sheet.getRange(idx + 2, 1, 1, OF_PROGRESS_HEADERS.length).setValues([row]);
  else          sheet.appendRow(row);

  return { before: before, after: _ofProgressToObj(row) };
}

/** 受注進捗の変更を管理シートへ同期（納期・ステータス） */
function _ofSyncMgmt(mgmtId, before, after) {
  var upd = { mgmtId: mgmtId };
  var changed = false;
  if (after.answerDate && after.answerDate !== before.answerDate) { upd.deliveryDate = after.answerDate; changed = true; }
  var ms = OF_TO_MGMT_STATUS[after.progress];
  if (ms && after.progress !== before.progress) { upd.status = ms; changed = true; }
  if (changed) _apiUpdateMgmt(upd);
}

function apiOfSave(p) {
  if (!p || !p.mgmtId) return { success: false, error: '管理IDが必要です' };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var res = _ofUpsertProgress(_ofSheet(OF_SHEET.PROGRESS), p, _ofUser());
    _ofSyncMgmt(p.mgmtId, res.before, res.after);
    return { success: true, item: res.after };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 詳細（進捗＋明細ロットチェック＋メモ）
// ============================================================
function apiOfDetail(p) {
  try {
    if (!p || !p.mgmtId) return { success: false, error: '管理IDが必要です' };
    var detail = _apiGetDetail({ mgmtId: p.mgmtId });
    if (!detail.success) return detail;
    var rules = _ofLotRules();
    var lines = (detail.orderLines || []).map(function(l) {
      var rule = _ofFindLotRule(rules, l.itemName, l.spec);
      return {
        lineNo: l.lineNo, itemName: l.itemName, spec: l.spec, qty: l.qty, unit: l.unit,
        unitPrice: l.unitPrice, amount: l.amount, firstDelivery: l.firstDelivery,
        rule: rule, warning: _ofCheckQty(rule, l.qty),
      };
    });
    var progress = _ofProgressMap()[p.mgmtId] || _ofProgressToObj([p.mgmtId]);
    var mgmtRow  = getAllMgmtData().filter(function(r) { return String(r[MGMT_COLS.ID - 1]) === String(p.mgmtId); })[0];
    var mgmt     = mgmtRow ? _rowToObject(mgmtRow) : { id: p.mgmtId };
    if (!progress.notifyTo && mgmt.client) progress.notifyTo = _ofClientEmail(mgmt.client);
    return { success: true, mgmt: mgmt, progress: progress, lines: lines, memos: _ofMemos(p.mgmtId), statuses: OF_STATUSES };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// ③ 社内メモ
// ============================================================
function _ofMemos(mgmtId) {
  return _ofValues(_ofSheet(OF_SHEET.MEMO), OF_MEMO_HEADERS.length)
    .filter(function(r) { return String(r[1]) === String(mgmtId); })
    .map(function(r) { return { id: String(r[0]), mgmtId: String(r[1]), author: String(r[2]), body: String(r[3]), postedAt: String(r[4]) }; })
    .sort(function(a, b) { return a.postedAt < b.postedAt ? 1 : -1; });
}

function apiOfMemoList(p) {
  try { return { success: true, items: _ofMemos(p.mgmtId), me: _ofUser() }; }
  catch (e) { return { success: false, error: e.message }; }
}

function apiOfMemoAdd(p) {
  try {
    var body = String((p && p.body) || '').trim();
    if (!p.mgmtId || !body) return { success: false, error: '管理IDと本文が必要です' };
    var id = 'MM-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 900 + 100);
    _ofSheet(OF_SHEET.MEMO).appendRow([id, p.mgmtId, _ofUser(), body, nowJST()]);
    return { success: true, items: _ofMemos(p.mgmtId) };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/** 自分の投稿のみ削除可（管理者は全件可） */
function apiOfMemoDelete(p) {
  try {
    var sheet = _ofSheet(OF_SHEET.MEMO);
    var data  = _ofValues(sheet, OF_MEMO_HEADERS.length);
    var me    = _ofUser();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) !== String(p.memoId)) continue;
      if (String(data[i][2]) !== me && !_ofIsAdmin(me)) return { success: false, error: '他の人のメモは削除できません' };
      var mgmtId = String(data[i][1]);
      sheet.deleteRow(i + 2);
      return { success: true, items: _ofMemos(mgmtId) };
    }
    return { success: false, error: 'メモが見つかりません' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function _ofIsAdmin(email) {
  var s = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAILS') || '';
  if (!s.trim()) return true;
  return s.split(',').map(function(x) { return x.trim(); }).indexOf(email) >= 0;
}

// ============================================================
// ② 納期回答・出荷通知（Gmail 下書き作成）
// ============================================================
function _ofClientEmail(client) {
  try {
    var map = JSON.parse(PropertiesService.getScriptProperties().getProperty(OF_CLIENT_EMAIL_PROP) || '{}');
    return map[client] || '';
  } catch (e) { return ''; }
}

function _ofRememberClientEmail(client, email) {
  if (!client || !email) return;
  var props = PropertiesService.getScriptProperties();
  var map = {};
  try { map = JSON.parse(props.getProperty(OF_CLIENT_EMAIL_PROP) || '{}'); } catch (e) {}
  map[client] = email;
  props.setProperty(OF_CLIENT_EMAIL_PROP, JSON.stringify(map));
}

/** 通知文面を組み立てる（下書き作成前のプレビューにも使う） */
function _ofBuildNotice(type, mgmt, pg, lines) {
  var no = mgmt.orderNo || mgmt.quoteNo || mgmt.id;
  var head = (mgmt.client ? mgmt.client + '\nご担当者様\n\n' : '') + 'いつもお世話になっております。\n\n';
  var itemsText = (lines || []).slice(0, 30).map(function(l) {
    return '  ・' + (l.itemName || '') + (l.spec ? '（' + l.spec + '）' : '') + '　' + (l.qty || '') + (l.unit || '');
  }).join('\n');
  var subject, body;
  if (type === 'ship') {
    subject = '【出荷のご案内】ご注文番号 ' + no;
    body = head + '下記ご注文の商品を出荷いたしましたのでご案内申し上げます。\n\n'
      + '■ ご注文番号：' + no + '\n'
      + (mgmt.subject ? '■ 件名：' + mgmt.subject + '\n' : '')
      + '■ 出荷日：' + (pg.shipDate || '（未設定）') + '\n'
      + (pg.carrier    ? '■ 運送会社：' + pg.carrier + '\n' : '')
      + (pg.trackingNo ? '■ 送り状番号：' + pg.trackingNo + '\n' : '')
      + (pg.answerDate ? '■ お届け予定日：' + pg.answerDate + '\n' : '')
      + (itemsText ? '\n【出荷明細】\n' + itemsText + '\n' : '')
      + '\n到着まで今しばらくお待ちくださいませ。\n何卒よろしくお願いいたします。';
  } else {
    subject = '【納期のご回答】ご注文番号 ' + no;
    body = head + '下記ご注文につきまして、納期をご回答申し上げます。\n\n'
      + '■ ご注文番号：' + no + '\n'
      + (mgmt.subject ? '■ 件名：' + mgmt.subject + '\n' : '')
      + '■ 納品予定日：' + (pg.answerDate || '（未設定）') + '\n'
      + (pg.shipPlanDate ? '■ 出荷予定日：' + pg.shipPlanDate + '\n' : '')
      + (itemsText ? '\n【対象明細】\n' + itemsText + '\n' : '')
      + '\n変更が生じた場合は改めてご連絡いたします。\n何卒よろしくお願いいたします。';
  }
  return { subject: subject, body: body };
}

function apiOfNoticePreview(p) {
  try {
    var d = apiOfDetail({ mgmtId: p.mgmtId });
    if (!d.success) return d;
    var pg = Object.assign({}, d.progress, p.override || {});
    return Object.assign({ success: true, to: pg.notifyTo || '' }, _ofBuildNotice(p.type, d.mgmt, pg, d.lines));
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/** 送信はせず Gmail の下書きを作る。担当者が内容を確認してから送信する運用。 */
function apiOfNoticeDraft(p) {
  try {
    if (!p || !p.mgmtId) return { success: false, error: '管理IDが必要です' };
    var to = String(p.to || '').trim();
    if (!to) return { success: false, error: '宛先メールアドレスを入力してください' };
    if (!p.subject || !p.body) return { success: false, error: '件名と本文が必要です' };
    var draft = GmailApp.createDraft(to, p.subject, p.body);
    var type  = p.type === 'ship' ? '出荷通知' : '納期回答';
    apiOfSave({ mgmtId: p.mgmtId, notifyTo: to, notifiedAt: nowJST(), notifiedType: type + '（下書き）' });
    if (p.client) _ofRememberClientEmail(p.client, to);
    _ofSheet(OF_SHEET.MEMO).appendRow([
      'MM-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 900 + 100),
      p.mgmtId, _ofUser(), '📧 ' + type + 'メールの下書きを作成しました（宛先: ' + to + '）', nowJST()
    ]);
    return { success: true, draftId: draft.getId(), gmailUrl: 'https://mail.google.com/mail/u/0/#drafts' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// ⑤ CSV 一括更新（案件＋受注進捗）
// ============================================================
// CSV見出し → { target: 'mgmt' | 'progress', key }
var OF_CSV_COLUMNS = [
  { header: '管理ID',       target: 'id',       key: 'mgmtId' },
  { header: '注文番号',     target: 'mgmt',     key: 'orderNo' },
  { header: '見積番号',     target: 'mgmt',     key: 'quoteNo' },
  { header: '件名',         target: 'mgmt',     key: 'subject' },
  { header: '顧客名',       target: 'mgmt',     key: 'client' },
  { header: 'ステータス',   target: 'mgmt',     key: 'status' },
  { header: '注文種別',     target: 'mgmt',     key: 'orderType' },
  { header: '発注日',       target: 'mgmt',     key: 'orderDate' },
  { header: '納期',         target: 'mgmt',     key: 'deliveryDate' },
  { header: '注文金額',     target: 'mgmt',     key: 'orderAmount' },
  { header: '機種コード',   target: 'mgmt',     key: 'modelCode' },
  { header: '担当者',       target: 'mgmt',     key: 'assignee' },
  { header: '進捗ステータス', target: 'progress', key: 'progress' },
  { header: '納期回答日',   target: 'progress', key: 'answerDate' },
  { header: '出荷予定日',   target: 'progress', key: 'shipPlanDate' },
  { header: '出荷日',       target: 'progress', key: 'shipDate' },
  { header: '運送会社',     target: 'progress', key: 'carrier' },
  { header: '送り状番号',   target: 'progress', key: 'trackingNo' },
  { header: '通知先メール', target: 'progress', key: 'notifyTo' },
];
var OF_DATE_KEYS = ['orderDate', 'deliveryDate', 'answerDate', 'shipPlanDate', 'shipDate'];

function apiOfCsvColumns() {
  return { success: true, columns: OF_CSV_COLUMNS.map(function(c) { return c.header; }) };
}

function _ofNormCell(key, v) {
  var s = String(v == null ? '' : v).trim();
  if (OF_DATE_KEYS.indexOf(key) >= 0) return _ofDate(s);
  if (key === 'orderAmount') return s === '' ? '' : String(Number(s.replace(/[,¥円\s]/g, '')) || 0);
  return s;
}

/**
 * p.rows: [{見出し: 値, ...}]。空欄セルは「変更しない」扱い。
 * 戻り値 changes: [{mgmtId, label, diffs:[{header, before, after}]}] / errors: [{rowNo, message}]
 */
function _ofCsvDiff(rows) {
  var list = apiOfList();
  if (!list.success) throw new Error(list.error);
  var byId = {};
  list.items.forEach(function(it) { byId[it.id] = it; });
  var changes = [], errors = [];
  rows.forEach(function(row, i) {
    var rowNo = i + 2;
    var id = String(row['管理ID'] || '').trim();
    if (!id) { errors.push({ rowNo: rowNo, message: '管理IDが空です' }); return; }
    var cur = byId[id];
    if (!cur) { errors.push({ rowNo: rowNo, message: '管理ID ' + id + ' が見つかりません（受注済み案件のみ対象）' }); return; }
    var diffs = [], mgmt = {}, progress = {};
    OF_CSV_COLUMNS.forEach(function(c) {
      if (c.target === 'id' || !(c.header in row)) return;
      var after = _ofNormCell(c.key, row[c.header]);
      if (after === '') return;
      var before = _ofNormCell(c.key, cur[c.key]);
      if (after === before) return;
      if (c.key === 'progress' && OF_STATUSES.indexOf(after) < 0) {
        errors.push({ rowNo: rowNo, message: '進捗ステータス「' + after + '」は無効です（' + OF_STATUSES.join('/') + '）' });
        return;
      }
      diffs.push({ header: c.header, before: before, after: after });
      (c.target === 'mgmt' ? mgmt : progress)[c.key] = after;
    });
    if (diffs.length) changes.push({ mgmtId: id, label: (cur.orderNo || cur.quoteNo || id) + ' ' + (cur.client || ''), diffs: diffs, mgmt: mgmt, progress: progress });
  });
  return { changes: changes, errors: errors };
}

function apiOfCsvPreview(p) {
  try {
    var d = _ofCsvDiff((p && p.rows) || []);
    return {
      success: true, errors: d.errors,
      changes: d.changes.map(function(c) { return { mgmtId: c.mgmtId, label: c.label, diffs: c.diffs }; }),
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function apiOfCsvApply(p) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var d = _ofCsvDiff((p && p.rows) || []);
    var sheet = _ofSheet(OF_SHEET.PROGRESS);
    var user  = _ofUser();
    var applied = 0, failed = [];
    d.changes.forEach(function(c) {
      try {
        if (Object.keys(c.mgmt).length) {
          var r = _apiUpdateMgmt(Object.assign({ mgmtId: c.mgmtId }, c.mgmt));
          if (!r.success) throw new Error(r.error);
        }
        if (Object.keys(c.progress).length) {
          var res = _ofUpsertProgress(sheet, Object.assign({ mgmtId: c.mgmtId }, c.progress), user);
          _ofSyncMgmt(c.mgmtId, res.before, res.after);
        }
        applied++;
      } catch (e2) {
        failed.push({ mgmtId: c.mgmtId, message: e2.message });
      }
    });
    return { success: true, applied: applied, failed: failed, errors: d.errors };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}
