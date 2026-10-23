// ============================================================
// 46_board_price_table.gs
// 基板価格表（基板ごとの 実装費・材料費・管理費・値引き・PCB単価 と 利益／逆ざや判定）
//
//   縦 = 基板名 ／ 横 = 実装費(売値・原価) 材料費(売値・原価) 管理費 利益 値引き 合計金額 PCB単価 … 実利益
//
//   ・利益     = 見積書に記載する利益（手入力）
//   ・合計金額 = 実装費売値 + 材料費売値 + 管理費 + 利益 − 値引き        （自動計算。値引きはマイナスで入れても差し引く）
//   ・実利益   = 合計金額 − 原価合計                                      （自動計算・シートの最後の列）
//       原価合計 = 実装費原価 + 材料費原価（設定ONの時だけ PCB単価も原価に含める）
//   ・逆ざやライン = 他の原価を据え置いたとき、その原価が いくらまで上がると利益が0になるか
//
//   ★ 金額を変更する保存は「変更理由（履歴）」が無いと必ず失敗する（サーバー側で強制）。
//      スプレッドシートを直接書き換えた場合も、編集トリガー(bpOnEdit)が履歴に「直接編集」として残す。
//   ★ 実装費・PCB の仕入見積書は Drive にアップロードして基板に添付（版を残す）。
//
//   シート: 基板価格表 / 基板価格_履歴 / 基板価格_仕入見積書
// ============================================================

var BP_SHEET      = '基板価格表';
var BP_HIST_SHEET = '基板価格_履歴';
var BP_DOC_SHEET  = '基板価格_仕入見積書';
var BP_FOLDER_NAME = '基板価格表_仕入見積書';
var BP_PROP_INCLUDE_PCB = 'BP_INCLUDE_PCB';
var BP_PROP_FOLDER_ID   = 'BP_DOC_FOLDER_ID';

// 金額項目（保存・履歴の対象）。kind: 'sale' = 価格改定 / 'cost' = 原価更新
var BP_FIELDS = [
  { key: 'mountSale', label: '実装費(売値)', kind: 'sale', col: 3 },
  { key: 'mountCost', label: '実装費(原価)', kind: 'cost', col: 4 },
  { key: 'matSale',   label: '材料費(売値)', kind: 'sale', col: 5 },
  { key: 'matCost',   label: '材料費(原価)', kind: 'cost', col: 6 },
  { key: 'mgmtFee',   label: '管理費',       kind: 'sale', col: 7 },
  { key: 'quoteProfit', label: '利益(見積記載)', kind: 'sale', col: 8 },   // 見積書に記載する利益（手入力）
  { key: 'discount',  label: '値引き',       kind: 'sale', col: 9 },
  { key: 'pcbPrice',  label: 'PCB単価',      kind: 'cost', col: 11 },
];
var BP_HEADERS = ['基板ID', '基板名', '実装費_売値', '実装費_原価', '材料費_売値', '材料費_原価', '管理費', '利益', '値引き',
                  '合計金額(自動)', 'PCB単価', '実利益率%(自動)', '逆ざや(自動)', '備考', '登録日時', '更新日時', '更新者', '実利益(自動)'];
var BP_COL = { ID: 1, NAME: 2, QUOTE_PROFIT: 8, TOTAL: 10, RATE: 12, INVERTED: 13, REMARKS: 14, CREATED: 15, UPDATED: 16, UPDATER: 17, PROFIT: 18 };
var BP_NCOLS = 18;   // PROFIT(18列目) = 実利益（自動）。8列目の「利益」は見積記載用の手入力
var BP_PROP_LAYOUT = 'BP_LAYOUT_V2';   // 8列目を「利益(自動)」→「利益(手入力)」に変えた版

var BP_HIST_HEADERS = ['履歴ID', '変更ID', '基板ID', '基板名', '種別', '項目', '変更前', '変更後', '差額', '変更率%', '改定日(適用日)',
                       '依頼元', '状態', '理由・メモ', '変更後の利益', '変更後の逆ざや', '入力者', '登録日時'];
var BP_HIST_NCOLS = 18;
var BP_HIST_KINDS = ['価格改定', '原価更新'];
var BP_HIST_STATUSES = ['適用済み', '依頼中（未反映）', '見送り'];

var BP_DOC_HEADERS = ['書類ID', '基板ID', '基板名', '区分', 'ファイル名', 'URL', 'メモ', '登録者', '登録日時'];
var BP_DOC_KINDS = ['実装費', 'PCB'];

// ------------------------------------------------------------
// 純粋関数（計算）
// ------------------------------------------------------------
function _bpNum(x) {
  if (x === '' || x === null || x === undefined) return null;
  // 全角数字・全角の小数点・カンマ・¥ も可。金額は小数第2位まで（3位以下は四捨五入）
  var s = String(x).normalize('NFKC').replace(/[,，\s¥￥円]/g, '').replace(/。/g, '.');
  if (s === '' || !/^-?(\d+\.?\d*|\.\d+)$/.test(s)) return null;
  var n = Number(s);
  return isNaN(n) ? null : Math.round(n * 100) / 100;
}
function _bpRound(n, d) { var p = Math.pow(10, d || 0); return Math.round(n * p) / p; }

/**
 * 合計金額・利益・利益率・逆ざやライン を計算する。
 * v: { mountSale, mountCost, matSale, matCost, mgmtFee, discount, pcbPrice }（空欄可）
 */
function _bpCalc(v, includePcb) {
  v = v || {};
  var n = function (k) { return _bpNum(v[k]); };
  var z = function (k) { var x = n(k); return x === null ? 0 : x; };

  var hasSale = ['mountSale', 'matSale', 'mgmtFee', 'quoteProfit'].some(function (k) { return n(k) !== null; });
  // 合計金額 = 実装費売値 + 材料費売値 + 管理費 + 利益(見積記載) − 値引き（値引きはマイナスで入っていても差し引く）
  var total = hasSale ? _bpRound(z('mountSale') + z('matSale') + z('mgmtFee') + z('quoteProfit') - Math.abs(z('discount')), 2) : null;

  var costKeys = ['mountCost', 'matCost'].concat(includePcb ? ['pcbPrice'] : []);
  var hasCost = costKeys.some(function (k) { return n(k) !== null; });
  var cost = hasCost ? _bpRound(costKeys.reduce(function (s, k) { return s + z(k); }, 0), 2) : null;

  var profit = (total !== null && hasCost) ? _bpRound(total - cost, 2) : null;
  var rate = (profit !== null && total > 0) ? _bpRound(profit / total * 100, 1) : null;
  var headroomRate = (profit !== null && cost > 0) ? _bpRound(profit / cost * 100, 1) : null;
  var limits = {};
  if (profit !== null) costKeys.forEach(function (k) { limits[k] = _bpRound(z(k) + profit, 2); });
  return {
    total: total, cost: cost, profit: profit, rate: rate,
    inverted: profit !== null && profit < 0,
    headroomRate: headroomRate,   // 原価合計がこの%上がると利益0
    limits: limits,               // 原価項目ごとの逆ざやライン（その原価だけが上がった場合）
    costKeys: costKeys,
  };
}

/** 変更前後の金額項目を比べて差分を返す（空欄→空欄は差分なし） */
function _bpDiff(before, after) {
  var out = [];
  BP_FIELDS.forEach(function (f) {
    var b = _bpNum(before[f.key]), a = _bpNum(after[f.key]);
    if (b === a) return;
    out.push({
      key: f.key, label: f.label, kind: f.kind, before: b, after: a,
      diff: (a !== null && b !== null) ? _bpRound(a - b, 2) : null,
      pct: (a !== null && b !== null && b !== 0) ? _bpRound((a - b) / b * 100, 1) : null,
    });
  });
  return out;
}

// ------------------------------------------------------------
// シート初期化
// ------------------------------------------------------------
function _bpSetHeader(sheet, headers, color) {
  var hr = sheet.getRange(1, 1, 1, headers.length);
  hr.setValues([headers]);
  hr.setBackground(color);
  hr.setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function _bpInitSheets() {
  var ss = getSpreadsheet();
  var created = false;

  var main = ss.getSheetByName(BP_SHEET);
  if (!main) {
    main = ss.insertSheet(BP_SHEET);
    _bpSetHeader(main, BP_HEADERS, '#E3F2FD');
    main.setFrozenColumns(2);
    main.setColumnWidth(BP_COL.NAME, 200);
    main.getRange(2, BP_COL.CREATED, main.getMaxRows() - 1, 2).setNumberFormat('@');   // 登録日時・更新日時を文字列のまま保つ
    // 直接編集は警告を出す（編集しても bpOnEdit が履歴に残す）
    try {
      main.protect().setDescription('基板価格表：金額は画面（価格表→基板価格表）から、変更理由を付けて編集してください').setWarningOnly(true);
    } catch (e) { Logger.log('[BP protect] ' + e.message); }
    created = true;
  }
  var hist = ss.getSheetByName(BP_HIST_SHEET);
  if (!hist) {
    hist = ss.insertSheet(BP_HIST_SHEET);
    _bpSetHeader(hist, BP_HIST_HEADERS, '#FFF3E0');
    hist.getRange(2, 11, hist.getMaxRows() - 1, 1).setNumberFormat('@');   // 日付が勝手にDate化されないように
    hist.getRange(2, 18, hist.getMaxRows() - 1, 1).setNumberFormat('@');
    hist.setColumnWidth(14, 280);
  }
  var doc = ss.getSheetByName(BP_DOC_SHEET);
  if (!doc) {
    doc = ss.insertSheet(BP_DOC_SHEET);
    _bpSetHeader(doc, BP_DOC_HEADERS, '#E8F5E9');
    doc.getRange(2, 9, doc.getMaxRows() - 1, 1).setNumberFormat('@');
  }
  if (!created) _bpMigrateLayout(main);
  if (created) { try { setupBoardPriceTrigger(); } catch (e2) { Logger.log('[BP trigger] ' + e2.message); } }
  return { main: main, hist: hist, doc: doc };
}

/** スプレッドシートを直接編集したときに履歴へ残す編集トリガーを登録（重複登録しない） */
function setupBoardPriceTrigger() {
  var exists = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'bpOnEdit'; });
  if (exists) return { success: true, already: true };
  ScriptApp.newTrigger('bpOnEdit').forSpreadsheet(getSpreadsheet()).onEdit().create();
  return { success: true, already: false };
}

/**
 * 旧版（8列目＝利益(自動)）のシートを新版へ移行（1回だけ）
 *   8列目：自動計算の値を消して「利益（見積記載・手入力）」にする
 *   18列目：「実利益(自動)」を追加し、全行の 合計金額・実利益・実利益率・逆ざや を再計算
 */
function _bpMigrateLayout(main) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(BP_PROP_LAYOUT) === '1') return;
  var h8 = String(main.getRange(1, BP_COL.QUOTE_PROFIT).getValue() || '');
  if (main.getMaxColumns() < BP_NCOLS) main.insertColumnsAfter(main.getMaxColumns(), BP_NCOLS - main.getMaxColumns());
  var last = main.getLastRow();
  if (h8.indexOf('自動') >= 0 && last > 1) main.getRange(2, BP_COL.QUOTE_PROFIT, last - 1, 1).clearContent();
  _bpSetHeader(main, BP_HEADERS, '#E3F2FD');
  main.getRange(1, BP_COL.QUOTE_PROFIT).setBackground('#FFF9C4').setNote('見積書に記載する利益（手入力）。合計金額に含まれます');
  main.getRange(1, BP_COL.PROFIT).setNote('実利益 = 合計金額 − 原価合計（自動計算）');
  _bpRecalcAll(main, _bpIncludePcb());
  props.setProperty(BP_PROP_LAYOUT, '1');
  Logger.log('[BP] 基板価格表を新レイアウト（利益=手入力／実利益=自動）へ移行しました');
}

/** 全行の自動計算列（合計金額・実利益・実利益率・逆ざや）を再計算 */
function _bpRecalcAll(main, includePcb) {
  var last = main.getLastRow();
  if (last <= 1) return;
  var rg = main.getRange(2, 1, last - 1, BP_NCOLS);
  var data = rg.getValues();
  data.forEach(function (r) {
    if (r[0] === '') return;
    var calc = _bpCalc(_bpRowToValues(r), includePcb);
    r[BP_COL.PROFIT - 1] = calc.profit === null ? '' : calc.profit;
    r[BP_COL.TOTAL - 1] = calc.total === null ? '' : calc.total;
    r[BP_COL.RATE - 1] = calc.rate === null ? '' : calc.rate;
    r[BP_COL.INVERTED - 1] = calc.inverted ? '⚠️逆ざや' : '';
  });
  rg.setValues(data);
}

// ------------------------------------------------------------
// 小物
// ------------------------------------------------------------
function _bpUser() {
  try { return Session.getActiveUser().getEmail() || '(不明)'; } catch (e) { return '(不明)'; }
}
function _bpId(prefix) {
  return prefix + '-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 9000 + 1000);
}
function _bpIncludePcb() {
  return PropertiesService.getScriptProperties().getProperty(BP_PROP_INCLUDE_PCB) === '1';
}
function _bpNormName(s) { return String(s == null ? '' : s).normalize('NFKC').toUpperCase().replace(/[\s　]/g, ''); }
function _bpValidDate(s) { return /^\d{4}\/\d{1,2}\/\d{1,2}$/.test(String(s || '')); }
function _bpNormDate(s) {
  var m = String(s || '').replace(/-/g, '/').match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  return m ? m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2) : '';
}
function _bpFieldByKey(key) {
  for (var i = 0; i < BP_FIELDS.length; i++) if (BP_FIELDS[i].key === key) return BP_FIELDS[i];
  return null;
}
function _bpFieldByCol(col) {
  for (var i = 0; i < BP_FIELDS.length; i++) if (BP_FIELDS[i].col === col) return BP_FIELDS[i];
  return null;
}

function _bpRowToValues(r) {
  var v = {};
  BP_FIELDS.forEach(function (f) { v[f.key] = _bpNum(r[f.col - 1]); });
  return v;
}

/** 基板価格表の1行ぶんの配列（入力値＋自動計算値）を作る */
function _bpBuildRow(id, name, values, remarks, createdAt, updatedAt, updater, includePcb) {
  var calc = _bpCalc(values, includePcb);
  var row = new Array(BP_NCOLS);
  for (var i = 0; i < BP_NCOLS; i++) row[i] = '';
  row[BP_COL.ID - 1] = id;
  row[BP_COL.NAME - 1] = name;
  BP_FIELDS.forEach(function (f) { var x = _bpNum(values[f.key]); row[f.col - 1] = x === null ? '' : x; });
  row[BP_COL.PROFIT - 1]   = calc.profit === null ? '' : calc.profit;
  row[BP_COL.TOTAL - 1]    = calc.total === null ? '' : calc.total;
  row[BP_COL.RATE - 1]     = calc.rate === null ? '' : calc.rate;
  row[BP_COL.INVERTED - 1] = calc.inverted ? '⚠️逆ざや' : '';
  row[BP_COL.REMARKS - 1]  = remarks || '';
  row[BP_COL.CREATED - 1]  = createdAt;
  row[BP_COL.UPDATED - 1]  = updatedAt;
  row[BP_COL.UPDATER - 1]  = updater;
  return row;
}

function _bpHistRow(o) {
  return [
    o.id || _bpId('BH'), o.editId || '', o.boardId, o.boardName, o.kind, o.label,
    o.before === null || o.before === undefined ? '' : o.before,
    o.after === null || o.after === undefined ? '' : o.after,
    o.diff === null || o.diff === undefined ? '' : o.diff,
    o.pct === null || o.pct === undefined ? '' : o.pct,
    o.date || '', o.requester || '', o.status || '適用済み', o.reason || '',
    o.profitAfter === null || o.profitAfter === undefined ? '' : o.profitAfter,
    o.invertedAfter ? '⚠️逆ざや' : '', o.user || _bpUser(), nowJST(),
  ];
}
function _bpAppendHist(sheet, rows) {
  if (!rows.length) return;
  var start = sheet.getLastRow() + 1;
  if (start + rows.length - 1 > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), rows.length + 50);
  // 日付・日時の列は文字列固定（'2026/10/22' がDate化されないように）。金額の列は数値のまま書く
  sheet.getRange(start, 11, rows.length, 1).setNumberFormat('@');
  sheet.getRange(start, 18, rows.length, 1).setNumberFormat('@');
  sheet.getRange(start, 1, rows.length, BP_HIST_NCOLS).setValues(rows);
}

function _bpFindRow(sheet, id) {
  var last = sheet.getLastRow();
  if (last <= 1) return -1;
  var ids = sheet.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); });
  var i = ids.indexOf(String(id));
  return i < 0 ? -1 : i + 2;
}

/** 基板の価格だけを軽く取得（機種の基板構成の合計計算用。履歴・添付は読まない） */
function _bpBoardsLite() {
  var sh = _bpInitSheets();
  var includePcb = _bpIncludePcb();
  var last = sh.main.getLastRow();
  var data = last > 1 ? sh.main.getRange(2, 1, last - 1, BP_NCOLS).getValues() : [];
  return data.filter(function (r) { return r[0] !== ''; }).map(function (r) {
    var values = _bpRowToValues(r);
    return { id: String(r[0]), name: String(r[BP_COL.NAME - 1] || ''), values: values, calc: _bpCalc(values, includePcb) };
  });
}

// ------------------------------------------------------------
// API: 一覧
// ------------------------------------------------------------
function apiBoardPriceList() {
  try {
    var sh = _bpInitSheets();
    var includePcb = _bpIncludePcb();
    var last = sh.main.getLastRow();
    var data = last > 1 ? sh.main.getRange(2, 1, last - 1, BP_NCOLS).getValues() : [];

    // 添付：基板×区分ごとの最新と件数
    var docs = {};
    var dl = sh.doc.getLastRow();
    if (dl > 1) {
      sh.doc.getRange(2, 1, dl - 1, BP_DOC_HEADERS.length).getValues().forEach(function (d) {
        if (!d[0]) return;
        var key = d[1] + '|' + d[3];
        var cur = docs[key] || { count: 0 };
        cur.count++;
        cur.url = d[5]; cur.name = d[4]; cur.date = String(d[8] || '');   // 行順＝登録順なので最後が最新
        docs[key] = cur;
      });
    }
    // 履歴件数・最終改定日
    var hist = {};
    var hl = sh.hist.getLastRow();
    if (hl > 1) {
      sh.hist.getRange(2, 1, hl - 1, BP_HIST_NCOLS).getValues().forEach(function (h) {
        if (!h[0]) return;
        var cur = hist[h[2]] || { count: 0, lastDate: '', pending: 0 };
        cur.count++;
        var d = _bpNormDate(String(h[10] || ''));
        if (d && h[12] === '適用済み' && d > cur.lastDate) cur.lastDate = d;
        if (h[12] === '依頼中（未反映）') cur.pending++;
        hist[h[2]] = cur;
      });
    }

    var rows = data.filter(function (r) { return r[0] !== ''; }).map(function (r) {
      var values = _bpRowToValues(r);
      var calc = _bpCalc(values, includePcb);
      var id = String(r[0]);
      var hs = hist[id] || { count: 0, lastDate: '', pending: 0 };
      return {
        id: id, name: String(r[BP_COL.NAME - 1] || ''), values: values, calc: calc,
        remarks: String(r[BP_COL.REMARKS - 1] || ''),
        createdAt: String(r[BP_COL.CREATED - 1] || ''), updatedAt: String(r[BP_COL.UPDATED - 1] || ''), updater: String(r[BP_COL.UPDATER - 1] || ''),
        docMount: docs[id + '|実装費'] || null, docPcb: docs[id + '|PCB'] || null,
        histCount: hs.count, lastRevision: hs.lastDate, pendingRequests: hs.pending,
      };
    });
    rows.sort(function (a, b) { return a.name.localeCompare(b.name, 'ja'); });
    return { success: true, rows: rows, includePcb: includePcb, fields: BP_FIELDS, generatedAt: nowJST() };
  } catch (e) {
    Logger.log('[apiBoardPriceList] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

// ------------------------------------------------------------
// API: 保存（新規登録／金額変更）。金額が変わる時は理由(履歴)が必須。
//   payload: { id?, name, values:{...}, remarks, reason, effectiveDate('yyyy/MM/dd'), requester }
// ------------------------------------------------------------
function apiBoardPriceSave(payload) {
  var lock = LockService.getScriptLock();
  try {
    payload = payload || {};
    lock.waitLock(20000);
    var sh = _bpInitSheets();
    var includePcb = _bpIncludePcb();
    var name = String(payload.name || '').trim();
    if (!name) return { success: false, error: '基板名を入力してください' };

    // 入力値の検証（数値・0以上）
    var values = {};
    for (var i = 0; i < BP_FIELDS.length; i++) {
      var f = BP_FIELDS[i], raw = (payload.values || {})[f.key];
      var x = _bpNum(raw);
      if (raw !== '' && raw !== null && raw !== undefined && x === null) return { success: false, error: f.label + ' は数値で入力してください' };
      if (x !== null && x < 0 && f.key !== 'discount') return { success: false, error: f.label + ' は0以上で入力してください' };
      values[f.key] = x;
    }

    var reason = String(payload.reason || '').trim();
    var requester = String(payload.requester || '').trim();
    var date = _bpNormDate(payload.effectiveDate);
    var now = nowJST(), user = _bpUser();

    // 同名チェック
    var last = sh.main.getLastRow();
    var data = last > 1 ? sh.main.getRange(2, 1, last - 1, BP_NCOLS).getValues() : [];
    var dup = data.filter(function (r) { return r[0] !== '' && String(r[0]) !== String(payload.id || '') && _bpNormName(r[BP_COL.NAME - 1]) === _bpNormName(name); });
    if (dup.length) return { success: false, error: '同じ基板名がすでにあります: ' + name };

    var editId = _bpId('BE');
    var histRows = [];
    var boardId;

    if (!payload.id) {
      // ── 新規登録（入力済みの金額は「新規登録」として履歴の起点を残す）
      boardId = _bpId('BP');
      var rowNew = _bpBuildRow(boardId, name, values, String(payload.remarks || ''), now, now, user, includePcb);
      sh.main.appendRow(rowNew);
      var calcNew = _bpCalc(values, includePcb);
      BP_FIELDS.forEach(function (f) {
        if (values[f.key] === null) return;
        histRows.push(_bpHistRow({
          editId: editId, boardId: boardId, boardName: name, kind: f.kind === 'cost' ? '原価更新' : '価格改定', label: f.label,
          before: null, after: values[f.key], date: date || _toDateStr(new Date()), requester: requester,
          reason: reason || '新規登録', profitAfter: calcNew.profit, invertedAfter: calcNew.inverted,
        }));
      });
      _bpAppendHist(sh.hist, histRows);
      return { success: true, id: boardId, created: true, historyAdded: histRows.length };
    }

    // ── 既存の更新
    boardId = String(payload.id);
    var rowNum = _bpFindRow(sh.main, boardId);
    if (rowNum < 0) return { success: false, error: '基板が見つかりません（削除された可能性があります）' };
    var cur = sh.main.getRange(rowNum, 1, 1, BP_NCOLS).getValues()[0];
    var oldValues = _bpRowToValues(cur);
    var diffs = _bpDiff(oldValues, values);

    if (diffs.length) {
      // ★ 金額を変えるなら 理由と改定日が必須（履歴なしの金額変更は受け付けない）
      if (!reason) return { success: false, error: '金額を変更する場合は「変更理由」を入力してください（履歴として必ず残します）', needReason: true };
      if (!_bpValidDate(payload.effectiveDate)) return { success: false, error: '改定日（適用日）を入力してください', needReason: true };
    }

    var rowUpd = _bpBuildRow(boardId, name, values, String(payload.remarks || ''), cur[BP_COL.CREATED - 1], now, user, includePcb);
    sh.main.getRange(rowNum, 1, 1, BP_NCOLS).setValues([rowUpd]);

    if (diffs.length) {
      var calc = _bpCalc(values, includePcb);
      diffs.forEach(function (d) {
        histRows.push(_bpHistRow({
          editId: editId, boardId: boardId, boardName: name, kind: d.kind === 'cost' ? '原価更新' : '価格改定', label: d.label,
          before: d.before, after: d.after, diff: d.diff, pct: d.pct, date: date, requester: requester,
          reason: reason, profitAfter: calc.profit, invertedAfter: calc.inverted,
        }));
      });
      _bpAppendHist(sh.hist, histRows);
    }
    return { success: true, id: boardId, created: false, historyAdded: histRows.length };
  } catch (e) {
    Logger.log('[apiBoardPriceSave] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// ------------------------------------------------------------
// API: 履歴一覧 / 手入力で履歴だけ追加 / 状態・メモ更新
// ------------------------------------------------------------
function apiBoardPriceHistory(payload) {
  try {
    var sh = _bpInitSheets();
    var boardId = String((payload || {}).boardId || '');
    var last = sh.hist.getLastRow();
    var rows = last > 1 ? sh.hist.getRange(2, 1, last - 1, BP_HIST_NCOLS).getValues() : [];
    var out = rows.filter(function (h) { return h[0] && (!boardId || String(h[2]) === boardId); }).map(function (h) {
      return {
        id: String(h[0]), editId: String(h[1]), boardId: String(h[2]), boardName: String(h[3]), kind: String(h[4]), label: String(h[5]),
        before: h[6] === '' ? null : Number(h[6]), after: h[7] === '' ? null : Number(h[7]),
        diff: h[8] === '' ? null : Number(h[8]), pct: h[9] === '' ? null : Number(h[9]),
        date: _bpNormDate(String(h[10] || '')) || String(h[10] || ''), requester: String(h[11] || ''), status: String(h[12] || ''),
        reason: String(h[13] || ''), profitAfter: h[14] === '' ? null : Number(h[14]), invertedAfter: !!h[15],
        user: String(h[16] || ''), at: String(h[17] || ''),
        direct: /直接編集/.test(String(h[13] || '')),
      };
    });
    // 新しい順（改定日 → 登録日時）
    out.sort(function (a, b) { return String(b.date).localeCompare(String(a.date)) || String(b.at).localeCompare(String(a.at)); });
    return { success: true, rows: out };
  } catch (e) { return { success: false, error: e.message }; }
}

/**
 * 過去の改定・原価値上げ依頼を手入力で履歴だけ追加する（現在の金額は変えない）。
 *   payload: { boardId, kind('価格改定'|'原価更新'), fieldKey, before, after, date, requester, status, reason }
 */
function apiBoardPriceHistoryAdd(payload) {
  var lock = LockService.getScriptLock();
  try {
    payload = payload || {};
    lock.waitLock(20000);
    var sh = _bpInitSheets();
    var includePcb = _bpIncludePcb();
    var rowNum = _bpFindRow(sh.main, payload.boardId);
    if (rowNum < 0) return { success: false, error: '基板が見つかりません' };
    var f = _bpFieldByKey(payload.fieldKey);
    if (!f) return { success: false, error: '項目を選んでください' };
    var kind = String(payload.kind || '');
    if (BP_HIST_KINDS.indexOf(kind) < 0) return { success: false, error: '種別を選んでください' };
    if ((kind === '原価更新') !== (f.kind === 'cost')) return { success: false, error: '種別と項目が合っていません（売値系＝価格改定／原価系＝原価更新）' };
    var after = _bpNum(payload.after), before = _bpNum(payload.before);
    if (after === null) return { success: false, error: '変更後の金額を入力してください' };
    if (after < 0 || (before !== null && before < 0)) return { success: false, error: '金額は0以上で入力してください' };
    var date = _bpNormDate(payload.date);
    if (!date) return { success: false, error: '日付（改定日／依頼日）を入力してください' };
    var reason = String(payload.reason || '').trim();
    if (!reason) return { success: false, error: '理由・メモを入力してください' };
    var status = BP_HIST_STATUSES.indexOf(payload.status) >= 0 ? payload.status : '適用済み';

    var cur = sh.main.getRange(rowNum, 1, 1, BP_NCOLS).getValues()[0];
    // 「その値になったら利益はいくらか」を現在の他の金額で試算して残す（依頼中の判断材料）
    var trial = _bpRowToValues(cur); trial[f.key] = after;
    var sim = _bpCalc(trial, includePcb);
    var d = (before !== null) ? _bpRound(after - before, 2) : null;
    var p = (before !== null && before !== 0) ? _bpRound((after - before) / before * 100, 1) : null;
    _bpAppendHist(sh.hist, [_bpHistRow({
      editId: _bpId('BE'), boardId: String(payload.boardId), boardName: String(cur[BP_COL.NAME - 1]), kind: kind, label: f.label,
      before: before, after: after, diff: d, pct: p, date: date, requester: String(payload.requester || '').trim(),
      status: status, reason: reason + '（手入力）', profitAfter: sim.profit, invertedAfter: sim.inverted,
    })]);
    return { success: true, simProfit: sim.profit, simInverted: sim.inverted };
  } catch (e) {
    Logger.log('[apiBoardPriceHistoryAdd] ' + e.message);
    return { success: false, error: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** 履歴の「状態」変更とメモ追記（履歴の削除・金額の書換えはできない） */
function apiBoardPriceHistoryUpdate(payload) {
  var lock = LockService.getScriptLock();
  try {
    payload = payload || {};
    lock.waitLock(20000);
    var sh = _bpInitSheets();
    var last = sh.hist.getLastRow();
    if (last <= 1) return { success: false, error: '履歴がありません' };
    var ids = sh.hist.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); });
    var i = ids.indexOf(String(payload.id || ''));
    if (i < 0) return { success: false, error: '履歴が見つかりません' };
    var rowNum = i + 2;
    var oldStatus = String(sh.hist.getRange(rowNum, 13).getValue() || '');
    var memo = String(sh.hist.getRange(rowNum, 14).getValue() || '');
    var add = [];
    if (payload.status && BP_HIST_STATUSES.indexOf(payload.status) >= 0 && payload.status !== oldStatus) {
      sh.hist.getRange(rowNum, 13).setValue(payload.status);
      add.push('状態: ' + oldStatus + '→' + payload.status);
    }
    if (String(payload.memo || '').trim()) add.push(String(payload.memo).trim());
    if (!add.length) return { success: false, error: '変更内容がありません' };
    sh.hist.getRange(rowNum, 14).setValue(memo + (memo ? ' / ' : '') + '【' + nowJST().slice(0, 10) + ' ' + _bpUser() + '】' + add.join('、'));
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
  finally { try { lock.releaseLock(); } catch (e2) {} }
}

// ------------------------------------------------------------
// API: 仕入見積書の添付（実装費／PCB）
// ------------------------------------------------------------
function _bpDocFolder() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(BP_PROP_FOLDER_ID);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* 消えていたら作り直す */ } }
  var parent = DriveApp.getFolderById(CONFIG.WEB_UPLOAD_FOLDER_ID);
  var it = parent.getFoldersByName(BP_FOLDER_NAME);
  var folder = it.hasNext() ? it.next() : parent.createFolder(BP_FOLDER_NAME);
  props.setProperty(BP_PROP_FOLDER_ID, folder.getId());
  return folder;
}

function apiBoardPriceDocUpload(payload) {
  try {
    payload = payload || {};
    if (!payload.base64Data || !payload.fileName) return { success: false, error: 'ファイルデータが不足しています' };
    if (BP_DOC_KINDS.indexOf(payload.kind) < 0) return { success: false, error: '区分（実装費／PCB）を選んでください' };
    var sh = _bpInitSheets();
    var rowNum = _bpFindRow(sh.main, payload.boardId);
    if (rowNum < 0) return { success: false, error: '基板が見つかりません' };
    var boardName = String(sh.main.getRange(rowNum, BP_COL.NAME).getValue() || '');

    var safe = function (s) { return String(s).replace(/[\/\\:*?"<>|]/g, '_'); };
    var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
    var blob = Utilities.newBlob(Utilities.base64Decode(payload.base64Data), payload.mimeType || 'application/octet-stream',
                                 safe(boardName) + '_' + payload.kind + '仕入見積_' + stamp + '_' + safe(payload.fileName));
    var file = _bpDocFolder().createFile(blob);
    var id = _bpId('BD');
    sh.doc.appendRow([id, String(payload.boardId), boardName, payload.kind, String(payload.fileName), file.getUrl(),
                      String(payload.memo || ''), _bpUser(), nowJST()]);
    return { success: true, id: id, url: file.getUrl() };
  } catch (e) {
    Logger.log('[apiBoardPriceDocUpload] ' + e.message);
    return { success: false, error: e.message };
  }
}

function apiBoardPriceDocList(payload) {
  try {
    var sh = _bpInitSheets();
    var boardId = String((payload || {}).boardId || '');
    var last = sh.doc.getLastRow();
    var rows = last > 1 ? sh.doc.getRange(2, 1, last - 1, BP_DOC_HEADERS.length).getValues() : [];
    var out = rows.filter(function (d) { return d[0] && String(d[1]) === boardId; }).map(function (d) {
      return { id: String(d[0]), kind: String(d[3]), fileName: String(d[4]), url: String(d[5]), memo: String(d[6] || ''), user: String(d[7] || ''), at: String(d[8] || '') };
    });
    out.reverse(); // 新しい順
    return { success: true, rows: out };
  } catch (e) { return { success: false, error: e.message }; }
}

/** 添付リンクを外す（Drive上のファイルは消さない） */
function apiBoardPriceDocDelete(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var sh = _bpInitSheets();
    var last = sh.doc.getLastRow();
    if (last <= 1) return { success: false, error: '添付がありません' };
    var ids = sh.doc.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); });
    var i = ids.indexOf(String((payload || {}).id || ''));
    if (i < 0) return { success: false, error: '添付が見つかりません' };
    sh.doc.deleteRow(i + 2);
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
  finally { try { lock.releaseLock(); } catch (e2) {} }
}

// ------------------------------------------------------------
// API: 設定（PCB単価を原価に含めるか）。金額は変わらないので履歴は不要、自動計算列だけ再計算。
// ------------------------------------------------------------
function apiBoardPriceSetting(payload) {
  var lock = LockService.getScriptLock();
  try {
    payload = payload || {};
    lock.waitLock(20000);
    if (payload.includePcb !== undefined) {
      var inc = !!payload.includePcb;
      PropertiesService.getScriptProperties().setProperty(BP_PROP_INCLUDE_PCB, inc ? '1' : '0');
      var sh = _bpInitSheets();
      var last = sh.main.getLastRow();
      if (last > 1) {
        _bpRecalcAll(sh.main, inc);
      }
    }
    return { success: true, includePcb: _bpIncludePcb() };
  } catch (e) { return { success: false, error: e.message }; }
  finally { try { lock.releaseLock(); } catch (e2) {} }
}

// ------------------------------------------------------------
// 編集トリガー：スプレッドシートの直接編集も 必ず履歴に残す（インストーラブル onEdit）
// ------------------------------------------------------------
function bpOnEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (sheet.getName() !== BP_SHEET) return;
    var r1 = e.range.getRow(), c1 = e.range.getColumn(), nr = e.range.getNumRows(), nc = e.range.getNumColumns();
    if (r1 + nr - 1 < 2) return;
    var includePcb = _bpIncludePcb();
    var hist = _bpInitSheets().hist;
    var single = (nr === 1 && nc === 1);
    var histRows = [];
    var editId = _bpId('BE');
    var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
    var firstRow = Math.max(r1, 2), lastRow = r1 + nr - 1;

    for (var r = firstRow; r <= lastRow; r++) {
      var cur = sheet.getRange(r, 1, 1, BP_NCOLS).getValues()[0];
      if (cur[0] === '') continue;
      var calc = _bpCalc(_bpRowToValues(cur), includePcb);
      for (var c = c1; c < c1 + nc; c++) {
        var f = _bpFieldByCol(c);
        if (!f) continue;
        var before = single ? _bpNum(e.oldValue) : null;
        var after = _bpNum(cur[c - 1]);
        if (single && before === after) continue;
        histRows.push(_bpHistRow({
          editId: editId, boardId: String(cur[0]), boardName: String(cur[BP_COL.NAME - 1]), kind: f.kind === 'cost' ? '原価更新' : '価格改定', label: f.label,
          before: before, after: after, diff: (before !== null && after !== null) ? _bpRound(after - before, 2) : null,
          date: today, reason: '【スプレッドシート直接編集】理由は未入力です（画面の履歴で追記してください）',
          profitAfter: calc.profit, invertedAfter: calc.inverted, user: (e.user && e.user.getEmail && e.user.getEmail()) || _bpUser(),
        }));
      }
      // 自動計算列を戻す（実利益・合計・実利益率・逆ざやは常に計算値。8列目の利益は手入力）＋更新情報
      var calc2 = _bpCalc(_bpRowToValues(cur), includePcb);
      sheet.getRange(r, BP_COL.PROFIT).setValue(calc2.profit === null ? '' : calc2.profit);
      sheet.getRange(r, BP_COL.TOTAL).setValue(calc2.total === null ? '' : calc2.total);
      sheet.getRange(r, BP_COL.RATE).setValue(calc2.rate === null ? '' : calc2.rate);
      sheet.getRange(r, BP_COL.INVERTED).setValue(calc2.inverted ? '⚠️逆ざや' : '');
      sheet.getRange(r, BP_COL.UPDATED).setValue(nowJST());
      sheet.getRange(r, BP_COL.UPDATER).setValue(_bpUser());
    }
    _bpAppendHist(hist, histRows);
  } catch (err) {
    Logger.log('[bpOnEdit] ' + err.message);
  }
}
