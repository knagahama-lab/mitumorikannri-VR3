// ============================================================
// 34_customer_part_master.gs
// 客先部品コードマスタ（お客様の部品コード ⇔ 品名・単価・弊社見積）
//
//   ・注文書のOCR取込時に、明細ごとの「客先部品コード」を自動でマスタ登録
//       新規コード → 🆕 新規品番として登録（確認待ち）
//       既知コード → 🔁 リピート。前回の品名・単価・数量・弊社見積を引き当て、
//                    弊社見積へ自動で紐づけ（AI推定より優先）。単価変更は警告。
//   ・結果は案件の「社内メモ」に自動記録
//   ・過去の注文書シートからマスタを一括作成（apiCpmRebuild）
//
// 部品コードは「品名が変わればコードも変わる」前提（藤商事：見本機/量産で別コード）なので
// キーは「取引先 × 客先部品コード」。
// 注文書シートの 21列目（ORDER_COLS.CUST_PART_CODE）に明細ごとのコードを保存する。
// ============================================================

var CPM_SHEET   = '客先品番マスタ';
var CPM_HEADERS = ['取引先','客先部品コード','客先品名','図番','機種コード','区分','弊社見積No','見積管理ID','見積単価',
                   '前回単価','前回数量','前回発注日','前回発注書No','前回管理ID','発注回数','初回発注日','確認済','備考','更新日時','弊社基板名'];
var CPM = { CLIENT:0, CODE:1, NAME:2, DRAWING:3, MODEL:4, VARIANT:5, QUOTE_NO:6, QUOTE_ID:7, QUOTE_PRICE:8,
            LAST_PRICE:9, LAST_QTY:10, LAST_DATE:11, LAST_ORDER_NO:12, LAST_MGMT_ID:13, COUNT:14, FIRST_DATE:15,
            CONFIRMED:16, NOTE:17, UPDATED_AT:18, BOARD:19 };
var CPM_ORDER_CODE_COL = 21; // 注文書シート：20列目は明細ステータスで使用済み

// ── シート ──
function _cpmSheet() {
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName(CPM_SHEET);
  if (sh) return sh;
  sh = ss.insertSheet(CPM_SHEET);
  sh.getRange(1, 1, 1, CPM_HEADERS.length).setValues([CPM_HEADERS])
    .setBackground('#FCE7F3').setFontWeight('bold').setFontSize(10);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, sh.getMaxRows(), CPM_HEADERS.length).setNumberFormat('@'); // 部品コードの先頭0・日付の自動変換を防ぐ
  return sh;
}
function _cpmRows() {
  var sh = _cpmSheet(); var last = sh.getLastRow();
  return last > 1 ? sh.getRange(2, 1, last - 1, CPM_HEADERS.length).getValues() : [];
}
function _cpmEnsureOrderHeader(os) {
  if (String(os.getRange(1, CPM_ORDER_CODE_COL).getValue()) !== '客先部品コード') {
    os.getRange(1, CPM_ORDER_CODE_COL).setValue('客先部品コード').setFontWeight('bold').setBackground('#FEF7E0');
  }
}

// ── 正規化・抽出 ──
function _cpmHalf(s) {
  return String(s == null ? '' : s)
    .replace(/[Ａ-Ｚａ-ｚ０-９－]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); })
    .replace(/[　]/g, ' ');
}
function _cpmCode(s) { return _cpmHalf(s).replace(/\s+/g, '').toUpperCase(); }

/** 取引先キー：取引先マスタの分類名（藤商事/コナミ…）、未分類なら社名の正規化 */
function _cpmClientKey(client) {
  var name = String(client || '').trim();
  if (!name || (typeof _cpmIsSelf === 'function' && _cpmIsSelf(name))) return '(取引先不明)'; // 自社（受注側）は取引先にしない
  try {
    var c = classifyClientName(name);
    if (c && !c.isFallback) return c.name;
  } catch (e) {}
  return name.replace(/株式会社|（株）|\(株\)|有限会社|\s|　/g, '') || '(取引先不明)';
}

/**
 * 明細から 部品コード・品名・図番・区分 を取り出す。
 * OCR の partCode を優先し、無ければ品名先頭の「数字5桁以上を含む英数字トークン」をコードとみなす。
 */
function _cpmSplit(item) {
  var name = _cpmHalf(item.itemName || '').replace(/\s+/g, ' ').trim();
  var code = _cpmCode(item.partCode || '');
  var m = name.match(/^([0-9A-Z][0-9A-Z\-]{4,24})\s+(.+)$/i);
  if (m && (m[1].match(/[0-9]/g) || []).length >= 5) {
    if (!code) code = _cpmCode(m[1]);
    if (_cpmCode(m[1]) === code) name = m[2];
  } else if (code && _cpmCode(name).indexOf(code) === 0) {
    name = name.substring(code.length).trim();
  }
  var drawing = String(item.drawingNo || '').trim();
  if (!drawing) {
    var d = name.match(/[（(]([A-Z]{2,}[0-9][0-9A-Z\-\/]*)[)）]/i);
    if (d) drawing = d[1];
  }
  var variant = /見本機/.test(name) ? '見本機' : /量産/.test(name) ? '量産' : /試作/.test(name) ? '試作' : '';
  return { code: code, name: name, drawing: drawing, variant: variant };
}

function _cpmToObj(r, rowNo) {
  return {
    rowNo: rowNo, client: String(r[CPM.CLIENT]), code: String(r[CPM.CODE]), name: String(r[CPM.NAME]),
    drawing: String(r[CPM.DRAWING]), modelCode: String(r[CPM.MODEL]), variant: String(r[CPM.VARIANT]),
    quoteNo: String(r[CPM.QUOTE_NO]), quoteMgmtId: String(r[CPM.QUOTE_ID]), quotePrice: Number(r[CPM.QUOTE_PRICE]) || 0,
    lastPrice: Number(r[CPM.LAST_PRICE]) || 0, lastQty: Number(r[CPM.LAST_QTY]) || 0, lastDate: _ofDate(r[CPM.LAST_DATE]),
    lastOrderNo: String(r[CPM.LAST_ORDER_NO]), lastMgmtId: String(r[CPM.LAST_MGMT_ID]), count: Number(r[CPM.COUNT]) || 0,
    firstDate: _ofDate(r[CPM.FIRST_DATE]), confirmed: String(r[CPM.CONFIRMED]) === 'TRUE', note: String(r[CPM.NOTE]),
    updatedAt: String(r[CPM.UPDATED_AT]), boardName: String(r[CPM.BOARD] || ''),
  };
}
function _cpmIndex(rows) {
  var idx = {};
  rows.forEach(function(r, i) { idx[r[CPM.CLIENT] + '|' + _cpmCode(r[CPM.CODE])] = i; });
  return idx;
}

/** 見積No. → { mgmtId, 明細 } */
function _cpmQuoteInfo(quoteNo, itemName, drawing) {
  quoteNo = String(quoteNo || '').trim();
  if (!quoteNo || quoteNo === '(複数)') return null;
  var mg = getAllMgmtData().map(_rowToObject).filter(function(o) { return o.quoteNo === quoteNo && !(o.orderNo && !o.quotePdfUrl); })[0];
  if (!mg) return { quoteNo: quoteNo, mgmtId: '', price: 0 };
  var price = 0;
  var qs = getSpreadsheet().getSheetByName(CONFIG.SHEET_QUOTES);
  if (qs && qs.getLastRow() > 1) {
    var lines = qs.getRange(2, 1, qs.getLastRow() - 1, QUOTE_COLS.UNIT_PRICE).getValues()
      .filter(function(r) { return String(r[QUOTE_COLS.MGMT_ID - 1]) === mg.id; });
    var key = _cpmCode(drawing || '');
    var hit = lines.filter(function(r) { return key && _cpmCode(r[QUOTE_COLS.ITEM_NAME - 1] + r[QUOTE_COLS.SPEC - 1]).indexOf(key) >= 0; })[0]
           || (lines.length === 1 ? lines[0] : null);
    if (hit) price = Number(hit[QUOTE_COLS.UNIT_PRICE - 1]) || 0;
  }
  return { quoteNo: quoteNo, mgmtId: mg.id, price: price };
}

// ============================================================
// ★ 注文書取込時フック（02_ocr_and_processing.gs の _saveOrderData から呼ぶ）
// ============================================================
/**
 * @param {string} mgmtId   注文の管理ID
 * @param {Object} ocr      OCR結果
 * @param {{startRow:number,count:number}} written  今回書き込んだ注文書シートの範囲
 * @return {{links:Array, allLinked:boolean, results:Array}}
 */
function cpmOnOrderSaved(mgmtId, ocr, written) {
  var out = { links: [], allLinked: false, results: [] };
  if (!written || !written.count) return out;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = getSpreadsheet();
    var os = ss.getSheetByName(CONFIG.SHEET_ORDERS);
    _cpmEnsureOrderHeader(os);
    var mgmt = getAllMgmtData().map(_rowToObject).filter(function(o) { return o.id === String(mgmtId); })[0] || {};
    var clientKey = _cpmClientKey(mgmt.client || ocr.clientName);
    var orderNo   = String(ocr.documentNo || mgmt.orderNo || '');
    var orderDate = _ofDate(ocr.documentDate || mgmt.orderDate);
    var modelCode = String(ocr.modelCode || mgmt.modelCode || '');

    var sh   = _cpmSheet();
    var rows = _cpmRows();
    var idx  = _cpmIndex(rows);
    var codes = [];
    var now = nowJST();
    var withCode = 0;

    (ocr.lineItems || []).slice(0, written.count).forEach(function(item, i) {
      var sp = _cpmSplit(item);
      codes.push([sp.code]);
      if (!sp.code) return;
      withCode++;
      var lineNo = i + 1;
      var qty = Number(item.qty) || 0, price = Number(item.unitPrice) || 0;
      var key = clientKey + '|' + sp.code;
      var res = { lineNo: lineNo, code: sp.code, name: sp.name, qty: qty, price: price };

      if (idx[key] !== undefined) {
        var r = rows[idx[key]];
        var m = _cpmToObj(r, idx[key] + 2);
        res.type = 'repeat';
        res.prev = { date: m.lastDate, orderNo: m.lastOrderNo, qty: m.lastQty, price: m.lastPrice, name: m.name };
        res.quoteNo = m.quoteNo; res.quotePrice = m.quotePrice;
        if (m.lastPrice && price && m.lastPrice !== price) res.priceChanged = true;
        if (m.quotePrice && price && m.quotePrice !== price) res.quoteDiff = true;
        if (m.name && m.name !== sp.name) res.nameChanged = true;
        if (m.lastOrderNo !== orderNo) { // 同じ注文書の再取込（差し替え）は回数に数えない
          r[CPM.COUNT] = (Number(r[CPM.COUNT]) || 0) + 1;
          r[CPM.LAST_PRICE] = price; r[CPM.LAST_QTY] = qty; r[CPM.LAST_DATE] = orderDate;
          r[CPM.LAST_ORDER_NO] = orderNo; r[CPM.LAST_MGMT_ID] = mgmtId;
        }
        r[CPM.NAME] = sp.name;
        if (sp.drawing) r[CPM.DRAWING] = sp.drawing;
        if (modelCode && !r[CPM.MODEL]) r[CPM.MODEL] = modelCode;
        r[CPM.UPDATED_AT] = now;
        sh.getRange(idx[key] + 2, 1, 1, CPM_HEADERS.length).setValues([r]);
        if (m.quoteMgmtId) out.links.push({ orderLineNo: lineNo, quoteMgmtId: m.quoteMgmtId, quoteNo: m.quoteNo, score: 100 });
      } else {
        var q = _cpmQuoteInfo(ocr.linkedQuoteNo, sp.name, sp.drawing) || {};
        var nr = new Array(CPM_HEADERS.length).fill('');
        nr[CPM.CLIENT] = clientKey; nr[CPM.CODE] = sp.code; nr[CPM.NAME] = sp.name; nr[CPM.DRAWING] = sp.drawing;
        nr[CPM.MODEL] = modelCode; nr[CPM.VARIANT] = sp.variant;
        nr[CPM.QUOTE_NO] = q.quoteNo || ''; nr[CPM.QUOTE_ID] = q.mgmtId || ''; nr[CPM.QUOTE_PRICE] = q.price || '';
        nr[CPM.LAST_PRICE] = price; nr[CPM.LAST_QTY] = qty; nr[CPM.LAST_DATE] = orderDate; nr[CPM.LAST_ORDER_NO] = orderNo;
        nr[CPM.LAST_MGMT_ID] = mgmtId; nr[CPM.COUNT] = 1; nr[CPM.FIRST_DATE] = orderDate; nr[CPM.CONFIRMED] = 'FALSE';
        nr[CPM.UPDATED_AT] = now;
        sh.appendRow(nr);
        rows.push(nr); idx[key] = rows.length - 1;
        res.type = 'new'; res.quoteNo = q.quoteNo || '';
      }
      out.results.push(res);
    });

    if (codes.length) os.getRange(written.startRow, CPM_ORDER_CODE_COL, codes.length, 1).setNumberFormat('@').setValues(codes);
    out.allLinked = withCode > 0 && withCode === written.count && out.links.length === written.count;
    _cpmWriteMemo(mgmtId, out.results);
  } finally {
    lock.releaseLock();
  }
  return out;
}

/** 部品コード由来の紐づけを適用（AI推定の後に呼び、コード一致を優先させる） */
function cpmApplyLinks(mgmtId, links) {
  if (!links || !links.length) return;
  _applyOrderLinks_LineLevel(mgmtId, links);
}

function _cpmWriteMemo(mgmtId, results) {
  if (!results.length) return;
  var yen = function(n) { return '¥' + Number(n || 0).toLocaleString(); };
  var lines = results.map(function(r) {
    if (r.type === 'new') return '🆕 新規品番 ' + r.code + ' ' + r.name + '（' + r.qty + '個 ' + yen(r.price) + '）' + (r.quoteNo ? ' 見積 ' + r.quoteNo : ' ※弊社見積の紐づけ待ち');
    var s = '🔁 リピート ' + r.code + ' ' + r.name + '：前回 ' + (r.prev.date || '') + ' No.' + (r.prev.orderNo || '') + ' ' + r.prev.qty + '個 ' + yen(r.prev.price)
          + ' → 今回 ' + r.qty + '個 ' + yen(r.price) + (r.quoteNo ? '（見積 ' + r.quoteNo + '）' : '');
    if (r.priceChanged) s += '\n　⚠ 単価が前回と異なります（' + yen(r.prev.price) + ' → ' + yen(r.price) + '）';
    if (r.quoteDiff)    s += '\n　⚠ 見積単価 ' + yen(r.quotePrice) + ' と異なります';
    if (r.nameChanged)  s += '\n　ℹ 品名表記が前回と異なります（前回: ' + r.prev.name + '）';
    return s;
  });
  try {
    _ofSheet(OF_SHEET.MEMO).appendRow([
      'MM-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 900 + 100),
      mgmtId, 'システム（客先品番）', lines.join('\n'), nowJST()
    ]);
  } catch (e) { Logger.log('[cpm memo] ' + e.message); }
}

// ============================================================
// 注文書シートの明細（部品コードつき）を読む
// ============================================================
function _cpmOrderLines() {
  var os = getSpreadsheet().getSheetByName(CONFIG.SHEET_ORDERS);
  if (!os || os.getLastRow() <= 1) return [];
  var width = Math.max(CPM_ORDER_CODE_COL, os.getLastColumn());
  var mg = {};
  getAllMgmtData().forEach(function(r) { var o = _rowToObject(r); mg[o.id] = o; });
  return os.getRange(2, 1, os.getLastRow() - 1, width).getValues().map(function(r, i) {
    var item = { itemName: r[ORDER_COLS.ITEM_NAME - 1], partCode: r[CPM_ORDER_CODE_COL - 1] };
    var sp = _cpmSplit(item);
    var m = mg[String(r[ORDER_COLS.MGMT_ID - 1])] || {};
    return {
      row: i + 2, storedCode: String(r[CPM_ORDER_CODE_COL - 1] || ''), code: sp.code, name: sp.name, drawing: sp.drawing, variant: sp.variant,
      mgmtId: String(r[ORDER_COLS.MGMT_ID - 1]), orderNo: String(r[ORDER_COLS.ORDER_NO - 1] || m.orderNo || ''),
      orderDate: _ofDate(r[ORDER_COLS.ORDER_DATE - 1] || m.orderDate), modelCode: String(r[ORDER_COLS.MODEL_CODE - 1] || m.modelCode || ''),
      lineNo: Number(r[ORDER_COLS.LINE_NO - 1]) || 0, qty: Number(r[ORDER_COLS.QTY - 1]) || 0, price: Number(r[ORDER_COLS.UNIT_PRICE - 1]) || 0,
      linkedQuote: String(r[ORDER_COLS.LINKED_QUOTE - 1] || m.quoteNo || ''), orderPdfUrl: String(r[ORDER_COLS.PDF_URL - 1] || m.orderPdfUrl || ''),
      client: m.client || '', clientKey: _cpmClientKey(m.client || ''),
    };
  });
}

// ============================================================
// 過去の注文書からマスタを一括作成・更新（手入力済みの見積No./備考/確認済/基板名は保持）
// ============================================================
function apiCpmRebuild() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = getSpreadsheet();
    var os = ss.getSheetByName(CONFIG.SHEET_ORDERS);
    if (!os || os.getLastRow() <= 1) return { success: true, created: 0, updated: 0, codesFilled: 0 };
    _cpmEnsureOrderHeader(os);
    var lines = _cpmOrderLines();

    // 注文書シート21列目にコードを書き戻し（空のものだけ）
    var filled = 0;
    var colVals = lines.map(function(l) { if (!l.storedCode && l.code) { filled++; return [l.code]; } return [l.storedCode]; });
    if (filled) os.getRange(2, CPM_ORDER_CODE_COL, colVals.length, 1).setNumberFormat('@').setValues(colVals);

    // キーごとに集計
    var groups = {};
    lines.forEach(function(l) {
      if (!l.code) return;
      var k = l.clientKey + '|' + l.code;
      (groups[k] = groups[k] || []).push(l);
    });
    var sh = _cpmSheet(); var rows = _cpmRows(); var idx = _cpmIndex(rows);
    var created = 0, updated = 0, now = nowJST();
    Object.keys(groups).forEach(function(k) {
      var g = groups[k].sort(function(a, b) { return String(a.orderDate).localeCompare(String(b.orderDate)) || a.row - b.row; });
      var last = g[g.length - 1], first = g[0];
      var orders = {}; g.forEach(function(l) { orders[l.orderNo || l.mgmtId] = true; });
      var linked = g.slice().reverse().map(function(l) { return l.linkedQuote; }).filter(function(q) { return q && q !== '(複数)'; })[0] || '';
      var r = idx[k] !== undefined ? rows[idx[k]] : new Array(CPM_HEADERS.length).fill('');
      var isNew = idx[k] === undefined;
      r[CPM.CLIENT] = last.clientKey; r[CPM.CODE] = last.code; r[CPM.NAME] = last.name;
      r[CPM.DRAWING] = last.drawing || r[CPM.DRAWING]; r[CPM.MODEL] = last.modelCode || r[CPM.MODEL]; r[CPM.VARIANT] = last.variant || r[CPM.VARIANT];
      if (!r[CPM.QUOTE_NO] && linked) {
        var q = _cpmQuoteInfo(linked, last.name, last.drawing) || {};
        r[CPM.QUOTE_NO] = q.quoteNo || ''; r[CPM.QUOTE_ID] = q.mgmtId || ''; r[CPM.QUOTE_PRICE] = q.price || '';
      }
      r[CPM.LAST_PRICE] = last.price; r[CPM.LAST_QTY] = last.qty; r[CPM.LAST_DATE] = last.orderDate;
      r[CPM.LAST_ORDER_NO] = last.orderNo; r[CPM.LAST_MGMT_ID] = last.mgmtId;
      r[CPM.COUNT] = Object.keys(orders).length; r[CPM.FIRST_DATE] = first.orderDate;
      if (isNew) r[CPM.CONFIRMED] = 'FALSE';
      r[CPM.UPDATED_AT] = now;
      if (isNew) { rows.push(r); idx[k] = rows.length - 1; created++; } else updated++;
    });
    if (rows.length) sh.getRange(2, 1, rows.length, CPM_HEADERS.length).setValues(rows);
    return { success: true, created: created, updated: updated, codesFilled: filled, noCode: lines.filter(function(l) { return !l.code; }).length };
  } catch (e) {
    Logger.log('[apiCpmRebuild] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 一覧・詳細・保存・注文別の引き当て
// ============================================================
function apiCpmList(p) {
  try {
    var kw = normalizeText((p && p.keyword) || '');
    var items = _cpmRows().map(function(r, i) { return _cpmToObj(r, i + 2); }).filter(function(m) {
      if (p && p.client && m.client !== p.client) return false;
      if (p && p.onlyUnconfirmed && m.confirmed) return false;
      return !kw || normalizeText([m.code, m.name, m.drawing, m.modelCode, m.quoteNo, m.boardName].join(' ')).indexOf(kw) >= 0;
    });
    items.sort(function(a, b) { return String(b.lastDate).localeCompare(String(a.lastDate)); });
    var clients = {};
    _cpmRows().forEach(function(r) { clients[r[CPM.CLIENT]] = (clients[r[CPM.CLIENT]] || 0) + 1; });
    return { success: true, items: items, clients: clients };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/** 部品コードで引き当て：マスタ＋発注履歴＋弊社見積 */
function apiCpmLookup(p) {
  try {
    var code = _cpmCode(p.code);
    var clientKey = p.client ? _cpmClientKey(p.client) : '';
    var master = _cpmRows().map(function(r, i) { return _cpmToObj(r, i + 2); })
      .filter(function(m) { return _cpmCode(m.code) === code && (!clientKey || m.client === clientKey); })[0] || null;
    var history = _cpmOrderLines().filter(function(l) { return l.code === code && (!master || l.clientKey === master.client); })
      .sort(function(a, b) { return String(b.orderDate).localeCompare(String(a.orderDate)); })
      .map(function(l) { return { orderDate: l.orderDate, orderNo: l.orderNo, mgmtId: l.mgmtId, qty: l.qty, price: l.price, name: l.name, linkedQuote: l.linkedQuote, orderPdfUrl: l.orderPdfUrl }; });
    var quote = null;
    if (master && master.quoteNo) {
      var mg = getAllMgmtData().map(_rowToObject).filter(function(o) { return o.id === master.quoteMgmtId || (!master.quoteMgmtId && o.quoteNo === master.quoteNo && o.quotePdfUrl); })[0];
      if (mg) quote = { quoteNo: mg.quoteNo, mgmtId: mg.id, quoteDate: mg.quoteDate, subject: mg.subject, pdfUrl: mg.quotePdfUrl, amount: mg.quoteAmount };
    }
    return { success: true, master: master, history: history, quote: quote };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/** 手入力項目の保存（見積No.・見積単価・弊社基板名・区分・備考・確認済） */
function apiCpmSave(p) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var sh = _cpmSheet(); var rows = _cpmRows(); var idx = _cpmIndex(rows);
    var k = p.client + '|' + _cpmCode(p.code);
    if (idx[k] === undefined) return { success: false, error: 'マスタに存在しません: ' + p.code };
    var r = rows[idx[k]];
    if (p.quoteNo !== undefined && String(p.quoteNo).trim() !== String(r[CPM.QUOTE_NO])) {
      var q = _cpmQuoteInfo(p.quoteNo, r[CPM.NAME], r[CPM.DRAWING]) || { quoteNo: '', mgmtId: '', price: 0 };
      r[CPM.QUOTE_NO] = q.quoteNo; r[CPM.QUOTE_ID] = q.mgmtId;
      if (p.quotePrice === undefined || p.quotePrice === '') r[CPM.QUOTE_PRICE] = q.price || '';
    }
    if (p.quotePrice !== undefined && p.quotePrice !== '') r[CPM.QUOTE_PRICE] = Number(String(p.quotePrice).replace(/[,¥]/g, '')) || '';
    if (p.boardName !== undefined) r[CPM.BOARD] = p.boardName;
    if (p.variant   !== undefined) r[CPM.VARIANT] = p.variant;
    if (p.note      !== undefined) r[CPM.NOTE] = p.note;
    if (p.confirmed !== undefined) r[CPM.CONFIRMED] = p.confirmed ? 'TRUE' : 'FALSE';
    r[CPM.UPDATED_AT] = nowJST();
    sh.getRange(idx[k] + 2, 1, 1, CPM_HEADERS.length).setValues([r]);
    return { success: true, item: _cpmToObj(r, idx[k] + 2) };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

/** 注文（管理ID）の各明細について、マスタから前回情報・弊社見積を引き当てる（詳細画面用） */
function apiCpmForOrder(p) {
  try {
    var lines = _cpmOrderLines().filter(function(l) { return l.mgmtId === String(p.mgmtId); });
    var seen = {};
    lines = lines.filter(function(l) { var k = l.lineNo + '|' + l.code; if (seen[k]) return false; seen[k] = true; return true; });
    var master = {};
    _cpmRows().forEach(function(r, i) { var m = _cpmToObj(r, i + 2); master[m.client + '|' + _cpmCode(m.code)] = m; });
    var all = _cpmOrderLines();
    var items = lines.map(function(l) {
      var m = l.code ? master[l.clientKey + '|' + l.code] : null;
      var prev = all.filter(function(x) { return x.code && x.code === l.code && x.clientKey === l.clientKey && x.mgmtId !== l.mgmtId && String(x.orderDate) <= String(l.orderDate); })
        .sort(function(a, b) { return String(b.orderDate).localeCompare(String(a.orderDate)); })[0] || null;
      return {
        lineNo: l.lineNo, code: l.code, name: l.name, qty: l.qty, price: l.price, clientKey: l.clientKey,
        master: m, prev: prev ? { orderDate: prev.orderDate, orderNo: prev.orderNo, mgmtId: prev.mgmtId, qty: prev.qty, price: prev.price } : null,
      };
    });
    return { success: true, items: items };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// Excel / CSV インポート
//   ・キー「取引先 × 客先部品コード」が既に登録済みの行は登録せずスキップ（既存データは変更しない）
//   ・ファイル内で同じキーが重複する場合は最初の行のみ採用
//   ・p.dryRun=true で登録せずに判定結果だけ返す（プレビュー）
// ============================================================
var CPM_IMPORT_ALIASES = {
  client:     ['取引先', '得意先', '顧客名', '客先', '取引先名'],
  code:       ['客先部品コード', '部品コード', '客先品番', '品番', '部品番号', 'パーツコード'],
  name:       ['客先品名', '品名', '部品名', '名称'],
  drawing:    ['図番', '型式', '型番'],
  modelCode:  ['機種コード', '機種'],
  variant:    ['区分'],
  quoteNo:    ['弊社見積No', '弊社見積No.', '見積No', '見積No.', '見積番号'],
  quotePrice: ['見積単価'],
  boardName:  ['弊社基板名', '弊社品名', '基板名'],
  lastPrice:  ['前回単価', '単価'],
  lastQty:    ['前回数量', '数量'],
  lastDate:   ['前回発注日', '発注日'],
  lastOrderNo:['前回発注書No', '前回発注書No.', '発注書No', '注文番号'],
  note:       ['備考'],
};

function _cpmPick(row, key) {
  var names = CPM_IMPORT_ALIASES[key];
  for (var i = 0; i < names.length; i++) {
    if (row[names[i]] !== undefined && String(row[names[i]]).trim() !== '') return String(row[names[i]]).trim();
  }
  return '';
}
function _cpmNum(v) { var n = Number(String(v || '').replace(/[,¥円\s]/g, '')); return isNaN(n) ? 0 : n; }

/** 見積No.→{mgmtId, price} を一括解決する関数を返す（行ごとにシートを読まないためのキャッシュ） */
function _cpmQuoteResolver() {
  var byNo = {};
  getAllMgmtData().map(_rowToObject).forEach(function(o) {
    if (o.quoteNo && !(o.orderNo && !o.quotePdfUrl) && !byNo[o.quoteNo]) byNo[o.quoteNo] = o.id;
  });
  var linesById = null;
  return function(quoteNo, drawing) {
    quoteNo = String(quoteNo || '').trim();
    if (!quoteNo) return { quoteNo: '', mgmtId: '', price: 0 };
    var id = byNo[quoteNo] || '';
    if (!id) return { quoteNo: quoteNo, mgmtId: '', price: 0 };
    if (!linesById) {
      linesById = {};
      var qs = getSpreadsheet().getSheetByName(CONFIG.SHEET_QUOTES);
      if (qs && qs.getLastRow() > 1) {
        qs.getRange(2, 1, qs.getLastRow() - 1, QUOTE_COLS.UNIT_PRICE).getValues().forEach(function(r) {
          var k = String(r[QUOTE_COLS.MGMT_ID - 1]); (linesById[k] = linesById[k] || []).push(r);
        });
      }
    }
    var lines = linesById[id] || [];
    var key = _cpmCode(drawing || '');
    var hit = lines.filter(function(r) { return key && _cpmCode(r[QUOTE_COLS.ITEM_NAME - 1] + r[QUOTE_COLS.SPEC - 1]).indexOf(key) >= 0; })[0]
           || (lines.length === 1 ? lines[0] : null);
    return { quoteNo: quoteNo, mgmtId: id, price: hit ? Number(hit[QUOTE_COLS.UNIT_PRICE - 1]) || 0 : 0 };
  };
}

/**
 * @param {{rows:Object[], defaultClient:string, confirmed:boolean, dryRun:boolean}} p
 *   rows は見出し→値 のオブジェクト配列（Excel/CSVの1行目が見出し）
 */
function apiCpmImport(p) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var rows = (p && p.rows) || [];
    if (!rows.length) return { success: false, error: '取り込む行がありません' };
    if (rows.length > 5000) return { success: false, error: '一度に取り込めるのは5,000行までです（' + rows.length + '行）' };
    var sh = _cpmSheet();
    var existing = _cpmIndex(_cpmRows());
    var resolve = _cpmQuoteResolver();
    var seen = {}, adds = [], newItems = [], skipped = [], errors = [];
    var now = nowJST();

    rows.forEach(function(row, i) {
      var rowNo = (Number(row.__rowNo) || i + 2);
      var code = _cpmCode(_cpmPick(row, 'code'));
      var clientRaw = _cpmPick(row, 'client') || (p.defaultClient || '');
      if (!code) { errors.push({ rowNo: rowNo, message: '部品コードが空です' }); return; }
      if (!clientRaw) { errors.push({ rowNo: rowNo, code: code, message: '取引先が空です（列が無い場合は取込画面で取引先を選択）' }); return; }
      var clientKey = _cpmClientKey(clientRaw);
      var key = clientKey + '|' + code;
      var name = _cpmHalf(_cpmPick(row, 'name')).replace(/\s+/g, ' ').trim();
      if (existing[key] !== undefined) { skipped.push({ rowNo: rowNo, client: clientKey, code: code, name: name, reason: '登録済み' }); return; }
      if (seen[key]) { skipped.push({ rowNo: rowNo, client: clientKey, code: code, name: name, reason: 'ファイル内で重複（' + seen[key] + '行目を採用）' }); return; }
      seen[key] = rowNo;

      var sp = _cpmSplit({ itemName: name, partCode: code, drawingNo: _cpmPick(row, 'drawing') });
      var quoteNo = _cpmPick(row, 'quoteNo');
      var q = quoteNo ? resolve(quoteNo, sp.drawing) : { quoteNo: '', mgmtId: '', price: 0 };
      var qPrice = _cpmNum(_cpmPick(row, 'quotePrice')) || q.price || '';
      var lastDate = _ofDate(_cpmPick(row, 'lastDate'));
      var r = new Array(CPM_HEADERS.length).fill('');
      r[CPM.CLIENT] = clientKey; r[CPM.CODE] = code; r[CPM.NAME] = sp.name; r[CPM.DRAWING] = sp.drawing;
      r[CPM.MODEL] = _cpmPick(row, 'modelCode'); r[CPM.VARIANT] = _cpmPick(row, 'variant') || sp.variant;
      r[CPM.QUOTE_NO] = q.quoteNo; r[CPM.QUOTE_ID] = q.mgmtId; r[CPM.QUOTE_PRICE] = qPrice;
      r[CPM.LAST_PRICE] = _cpmNum(_cpmPick(row, 'lastPrice')) || ''; r[CPM.LAST_QTY] = _cpmNum(_cpmPick(row, 'lastQty')) || '';
      r[CPM.LAST_DATE] = lastDate; r[CPM.LAST_ORDER_NO] = _cpmPick(row, 'lastOrderNo');
      r[CPM.COUNT] = r[CPM.LAST_ORDER_NO] || lastDate ? 1 : 0; r[CPM.FIRST_DATE] = lastDate;
      r[CPM.CONFIRMED] = p.confirmed ? 'TRUE' : 'FALSE'; r[CPM.NOTE] = _cpmPick(row, 'note');
      r[CPM.UPDATED_AT] = now; r[CPM.BOARD] = _cpmPick(row, 'boardName');
      adds.push(r);
      newItems.push({ rowNo: rowNo, client: clientKey, code: code, name: sp.name, quoteNo: q.quoteNo, quoteFound: !quoteNo || !!q.mgmtId, quotePrice: qPrice });
    });

    if (!p.dryRun && adds.length) {
      var start = sh.getLastRow() + 1;
      sh.getRange(start, 1, adds.length, CPM_HEADERS.length).setNumberFormat('@').setValues(adds);
    }
    return {
      success: true, dryRun: !!p.dryRun,
      added: adds.length, skippedCount: skipped.length, errorCount: errors.length,
      newItems: newItems.slice(0, 500), skipped: skipped.slice(0, 500), errors: errors.slice(0, 200),
    };
  } catch (e) {
    Logger.log('[apiCpmImport] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 弊社見積の紐づけ登録（見積番号入力／見積書検索から）
//   p.quoteNo : 見積番号（空＋p.unlink=true で解除）
//   p.items   : [{client, code}]  複数の部品コードへまとめて登録可
// ============================================================
function apiCpmLinkQuote(p) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var quoteNo = String((p && p.quoteNo) || '').trim();
    var items = (p && p.items) || [];
    if (!items.length) return { success: false, error: '対象の部品コードがありません' };
    if (!quoteNo && !p.unlink) return { success: false, error: '見積番号を入力してください' };

    var resolve = _cpmQuoteResolver();
    if (quoteNo) {
      var probe = resolve(quoteNo, '');
      var inLedger = !probe.mgmtId && getAllLedgerData().some(function(r) {
        return String(r[LEDGER_COLS.QUOTE_NO - 1] || '').trim() === quoteNo;
      });
      if (!probe.mgmtId && !inLedger) return { success: false, error: '見積番号「' + quoteNo + '」が見積書一覧・見積台帳に見つかりません' };
    }

    var sh = _cpmSheet(); var rows = _cpmRows(); var idx = _cpmIndex(rows);
    var now = nowJST(), updated = [];
    items.forEach(function(it) {
      var k = it.client + '|' + _cpmCode(it.code);
      if (idx[k] === undefined) return;
      var r = rows[idx[k]];
      if (quoteNo) {
        var q = resolve(quoteNo, r[CPM.DRAWING]);
        r[CPM.QUOTE_NO] = quoteNo; r[CPM.QUOTE_ID] = q.mgmtId; r[CPM.QUOTE_PRICE] = q.price || '';
      } else {
        r[CPM.QUOTE_NO] = ''; r[CPM.QUOTE_ID] = ''; r[CPM.QUOTE_PRICE] = '';
      }
      r[CPM.UPDATED_AT] = now;
      sh.getRange(idx[k] + 2, 1, 1, CPM_HEADERS.length).setValues([r]);
      updated.push(_cpmToObj(r, idx[k] + 2));
    });
    if (!updated.length) return { success: false, error: 'マスタに該当する部品コードがありません' };
    return { success: true, items: updated };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}
