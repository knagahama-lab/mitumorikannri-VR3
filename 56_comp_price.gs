// ============================================================
// 56_comp_price.gs
// 構成表（54）× 客先品番マスタ（34）× 注文書 の単価管理
//
//   ① 構成表の各品番を「部品コード」で客先品番マスタ・注文書に紐づける（品名が変わっても追える）
//   ② 各品番に 現行単価（最新の注文の単価）／新単価（価格改定の予定）／次回の想定単価 を出す
//   ③ 注文書が来たら、部品コードでその注文を見つけ、単価が想定どおりかを自動でチェックする
//        想定単価 = 新単価（その注文の納入日・発注日が適用日以降なら）→ 1つ前の注文の単価 → 見積単価
//   ④ 価格UPの前に、提出した見積書の単価を「新単価（適用日つき）」として品番マスタに登録しておく
//
//   新単価は 客先品番マスタ の 21〜23列目（新単価／新単価適用日／新単価見積No）に保存する。
//   34 の処理は 20列目までしか書かないので、ここの値を上書きしない。
// ============================================================

var CP_COL_PRICE = 21, CP_COL_DATE = 22, CP_COL_QUOTE = 23;   // 客先品番マスタの列（1始まり）
var CP_HEADERS = ['新単価', '新単価適用日', '新単価見積No'];

function _cpStr(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd');
  return String(v).trim();
}
function _cpNum(v) {
  if (v === '' || v === null || v === undefined) return null;
  var n = Number(String(v).normalize('NFKC').replace(/[,¥円\s]/g, ''));
  return isNaN(n) ? null : Math.round(n * 100) / 100;
}
function _cpDate(v) { var s = _cpStr(v).replace(/-/g, '/'), m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})/); return m ? m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2) : ''; }
function _cpToday() { return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd'); }
/** 品名の比較用（全角半角・空白・カッコの違いをならす） */
function _cpNameKey(s) { return String(s || '').normalize('NFKC').toUpperCase().replace(/[\s　]+/g, '').replace(/[（]/g, '(').replace(/[）]/g, ')'); }
function _cpOurs(supplierCode, supplier) { return String(supplierCode).trim() === '9' || /サン電子/.test(String(supplier || '')); }
function _cpClientKey() { return _cpmClientKey('株式会社藤商事'); }

/** 客先品番マスタ：部品コード → { row, client, name, quoteNo, quotePrice, lastPrice, newPrice, newDate, newQuoteNo } （藤商事を優先） */
function _cpMaster() {
  var sh = _cpmSheet(), last = sh.getLastRow(), out = {}, fuji = _cpClientKey();
  if (last < 2) return out;
  var v = sh.getRange(2, 1, last - 1, CP_COL_QUOTE).getValues();
  v.forEach(function (r, i) {
    var code = _cpmCode(r[CPM.CODE]); if (!code) return;
    if (out[code] && out[code].client === fuji && String(r[CPM.CLIENT]) !== fuji) return;
    out[code] = { row: i + 2, client: String(r[CPM.CLIENT]), code: code, name: _cpStr(r[CPM.NAME]), quoteNo: _cpStr(r[CPM.QUOTE_NO]),
                  quotePrice: _cpNum(r[CPM.QUOTE_PRICE]), lastPrice: _cpNum(r[CPM.LAST_PRICE]), lastDate: _cpDate(r[CPM.LAST_DATE]),
                  newPrice: _cpNum(r[CP_COL_PRICE - 1]), newDate: _cpDate(r[CP_COL_DATE - 1]), newQuoteNo: _cpStr(r[CP_COL_QUOTE - 1]) };
  });
  return out;
}

/** 注文書の明細：部品コード別・品名別（新しい順） */
function _cpOrders() {
  var lines = _cpmOrderLines(), byCode = {}, byName = {};
  if (!lines.length) return { byCode: byCode, byName: byName };
  var os = getSpreadsheet().getSheetByName(CONFIG.SHEET_ORDERS);
  var del = os.getRange(2, ORDER_COLS.FIRST_DELIVERY, os.getLastRow() - 1, 1).getValues();
  var slip = os.getRange(2, ORDER_COLS.ORDER_SLIP_NO, os.getLastRow() - 1, 1).getValues();
  lines.forEach(function (l) {
    if (!l.orderNo && !l.price) return;
    var o = { orderNo: l.orderNo, slipNo: _cpStr(slip[l.row - 2][0]), date: l.orderDate, delivery: _cpDate(del[l.row - 2][0]), qty: l.qty, price: l.price,
              model: l.modelCode, pdfUrl: l.orderPdfUrl, name: l.name };
    var c = _cpmCode(l.code); if (c) (byCode[c] = byCode[c] || []).push(o);
    var n = _cpNameKey(l.name); if (n) (byName[n] = byName[n] || []).push(o);
  });
  var sort = function (m) {
    Object.keys(m).forEach(function (k) {
      var seen = {};
      m[k] = m[k].sort(function (a, b) { return String(b.delivery || b.date).localeCompare(String(a.delivery || a.date)) || String(b.orderNo).localeCompare(String(a.orderNo)); })
        .filter(function (o) { var x = o.orderNo + '|' + o.delivery + '|' + o.price; if (seen[x]) return false; seen[x] = 1; return true; });
    });
  };
  sort(byCode); sort(byName);
  return { byCode: byCode, byName: byName };
}

/** 1品番の単価の見立て */
function _cpJudge(m, orders, today) {
  var latest = orders[0] || null, prev = orders[1] || null;
  var np = m && m.newPrice, nd = m && m.newDate;
  var cur = latest ? latest.price : (m && m.lastPrice) || null;
  var curSrc = latest ? '最新の注文' : (m && m.lastPrice) ? '前回の注文' : '';
  var next = null, nextSrc = '';
  if (np !== null && np !== undefined && (!nd || nd <= today)) { next = np; nextSrc = '新単価（' + (nd ? nd + '〜' : '適用中') + '）'; }
  else if (cur) { next = cur; nextSrc = curSrc; }
  else if (m && m.quotePrice) { next = m.quotePrice; nextSrc = '見積単価' + (m.quoteNo ? '（' + m.quoteNo + '）' : ''); }
  var check = '', expect = null, expSrc = '';
  if (latest) {
    var base = latest.delivery || latest.date;
    if (np !== null && np !== undefined && (!nd || base >= nd)) { expect = np; expSrc = '新単価'; }
    else if (prev) { expect = prev.price; expSrc = '前回の注文（' + prev.orderNo + '）'; }
    else if (m && m.quotePrice) { expect = m.quotePrice; expSrc = '見積単価'; }
    if (expect === null) check = '比較なし';
    else check = Math.abs(latest.price - expect) < 0.005 ? '一致' : '不一致';
  }
  return { cur: cur, curSrc: curSrc, next: next, nextSrc: nextSrc, check: check, expect: expect, expSrc: expSrc,
           diff: (check === '不一致') ? Math.round((latest.price - expect) * 100) / 100 : null,
           future: (np !== null && np !== undefined && nd && nd > today) ? { price: np, date: nd } : null };
}

/** API：構成表の各品番の単価・注文チェック */
function apiCpList(p) {
  try {
    var rows = _csReadLines(), master = _cpMaster(), od = _cpOrders(), today = _cpToday();
    var lines = rows.map(function (r) {
      var code = _cpmCode(r[6]), name = _cpStr(r[7]), m = master[code] || null;
      var byCode = code ? (od.byCode[code] || []) : [], orders = byCode.length ? byCode : (od.byName[_cpNameKey(name)] || []);
      var j = _cpJudge(m, orders, today), o = orders[0] || null;
      return {
        model: _cpStr(r[0]), ver: _cpStr(r[1]), seq: Number(r[2]) || 0, level: Number(r[3]) || 1, supplierCode: _cpStr(r[4]), supplier: _cpStr(r[5]),
        code: code, name: name, dest: _cpStr(r[8]), unitQty: _cpNum(r[9]), note: _cpStr(r[14]),
        ours: _cpOurs(r[4], r[5]), inMaster: !!m, masterName: m ? m.name : '',
        quoteNo: m ? m.quoteNo : '', quotePrice: m ? m.quotePrice : null,
        newPrice: m ? m.newPrice : null, newDate: m ? m.newDate : '', newQuoteNo: m ? m.newQuoteNo : '',
        cur: j.cur, curSrc: j.curSrc, next: j.next, nextSrc: j.nextSrc, future: j.future,
        order: o ? { orderNo: o.orderNo, slipNo: o.slipNo, date: o.date, delivery: o.delivery, qty: o.qty, price: o.price, pdfUrl: o.pdfUrl, by: byCode.length ? 'コード' : '品名' } : null,
        orderCount: orders.length, check: j.check, expect: j.expect, expSrc: j.expSrc, diff: j.diff,
      };
    });
    var models = {};
    lines.forEach(function (l) {
      var g = models[l.model] || (models[l.model] = { model: l.model, ver: l.ver, ours: 0, noMaster: 0, ordered: 0, mismatch: 0, future: 0, noPrice: 0 });
      if (!l.ours) return;
      g.ours++;
      if (!l.inMaster) g.noMaster++;
      if (l.order) g.ordered++;
      if (l.check === '不一致') g.mismatch++;
      if (l.future) g.future++;
      if (l.next === null) g.noPrice++;
    });
    return { success: true, today: today, lines: lines, models: Object.keys(models).sort().map(function (k) { return models[k]; }) };
  } catch (e) {
    Logger.log('[apiCpList] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

/** 客先品番マスタの新単価の列見出し */
function _cpEnsureHeader(sh) {
  var h = sh.getRange(1, CP_COL_PRICE, 1, 3).getValues()[0];
  if (String(h[0]) !== CP_HEADERS[0]) {
    sh.getRange(1, CP_COL_PRICE, 1, 3).setValues([CP_HEADERS]).setFontWeight('bold').setBackground('#FEF08A');
    sh.getRange(2, CP_COL_DATE, Math.max(1, sh.getMaxRows() - 1), 2).setNumberFormat('@');
  }
}

/**
 * 新単価の登録・品番マスタへの登録
 *   p.items: [{ code, name, model, price, date, quoteNo }]   price 空 = 品番マスタへの登録だけ
 *   p.clear: true … 新単価を消す
 */
function apiCpSetPrice(p) {
  try {
    p = p || {};
    var items = (p.items || []).filter(function (x) { return _cpmCode(x.code); });
    if (!items.length) return { success: false, error: '部品コードがありません' };
    var sh = _cpmSheet(); _cpEnsureHeader(sh);
    var master = _cpMaster(), fuji = _cpClientKey(), now = nowJST(), add = [], upd = 0, reg = 0;
    items.forEach(function (x) {
      var code = _cpmCode(x.code), price = p.clear ? null : _cpNum(x.price), date = p.clear ? '' : _cpDate(x.date), q = p.clear ? '' : _cpStr(x.quoteNo);
      if (!p.clear && x.price !== '' && x.price !== null && x.price !== undefined && price === null) throw new Error('単価は数値で入力してください（' + code + '）');
      if (!p.clear && x.date && !date) throw new Error('適用日が正しくありません（' + code + '）');
      var m = master[code];
      if (m) {
        if (p.clear || price !== null) {
          sh.getRange(m.row, CP_COL_PRICE, 1, 3).setNumberFormat('@').setValues([[price === null ? '' : String(price), date, q]]);
          sh.getRange(m.row, CPM.UPDATED_AT + 1).setValue(now);
          upd++;
        }
        return;
      }
      var r = new Array(CP_COL_QUOTE).fill('');
      r[CPM.CLIENT] = fuji; r[CPM.CODE] = code; r[CPM.NAME] = _cpStr(x.name); r[CPM.MODEL] = _cpStr(x.model);
      r[CPM.COUNT] = 0; r[CPM.CONFIRMED] = 'TRUE'; r[CPM.NOTE] = '構成表から登録'; r[CPM.UPDATED_AT] = now;
      r[CP_COL_PRICE - 1] = price === null ? '' : String(price); r[CP_COL_DATE - 1] = date; r[CP_COL_QUOTE - 1] = q;
      master[code] = { row: -1 };
      add.push(r); reg++; if (price !== null) upd++;
    });
    if (add.length) {
      var lock = LockService.getScriptLock(), got = false;
      try { got = lock.tryLock(1500); } catch (e) {}
      if (got) {
        try { var at = sh.getLastRow() + 1; sh.getRange(at, 1, add.length, CP_COL_QUOTE).setNumberFormat('@').setValues(add); SpreadsheetApp.flush(); }
        finally { lock.releaseLock(); }
      } else {
        add.forEach(function (r) { sh.appendRow([r[0]]); sh.getRange(sh.getLastRow(), 1, 1, CP_COL_QUOTE).setNumberFormat('@').setValues([r]); });
      }
    }
    return { success: true, registered: reg, updated: upd };
  } catch (e) {
    Logger.log('[apiCpSetPrice] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

/** 見積書の明細と、構成表（機種）の品番の対応候補  p: { quoteNo, model } */
function apiCpQuoteLines(p) {
  try {
    p = p || {};
    var no = _cpNameKey(p.quoteNo);
    if (!no) return { success: false, error: '見積番号を入れてください' };
    var qs = getSpreadsheet().getSheetByName(CONFIG.SHEET_QUOTES), out = [], date = '', client = '';
    if (qs && qs.getLastRow() > 1) {
      qs.getRange(2, 1, qs.getLastRow() - 1, Math.max(QUOTE_COLS.PDF_URL, qs.getLastColumn())).getValues().forEach(function (r) {
        if (_cpNameKey(r[QUOTE_COLS.QUOTE_NO - 1]) !== no) return;
        date = date || _cpDate(r[QUOTE_COLS.ISSUE_DATE - 1]); client = client || _cpStr(r[QUOTE_COLS.DEST_COMPANY - 1]);
        var name = _cpStr(r[QUOTE_COLS.ITEM_NAME - 1]), spec = _cpStr(r[QUOTE_COLS.SPEC - 1]);
        if (!name && !spec) return;
        out.push({ name: name, spec: spec, qty: _cpNum(r[QUOTE_COLS.QTY - 1]), price: _cpNum(r[QUOTE_COLS.UNIT_PRICE - 1]), pdfUrl: _cpStr(r[QUOTE_COLS.PDF_URL - 1]) });
      });
    }
    // 見積に紐づけ済みの部品コード（見積書一覧の「客先部品コード」）
    var linked = [];
    _cpmRows().forEach(function (r) { if (_cpNameKey(r[CPM.QUOTE_NO]) === no) linked.push({ code: _cpmCode(r[CPM.CODE]), name: _cpStr(r[CPM.NAME]) }); });
    // 構成表の品番（その機種の自社分）
    var cand = _csReadLines().filter(function (r) { return (!p.model || _cpStr(r[0]) === p.model) && _cpOurs(r[4], r[5]) && _cpmCode(r[6]); })
      .map(function (r) { return { code: _cpmCode(r[6]), name: _cpStr(r[7]), model: _cpStr(r[0]) }; });
    var seen = {}; cand = cand.filter(function (c) { var k = c.model + '|' + c.code; if (seen[k]) return false; seen[k] = 1; return true; });
    out.forEach(function (l) {
      var best = null, bs = 0;
      cand.forEach(function (c) { var s = _cpScore(l.name + ' ' + l.spec, c.name); if (s > bs) { bs = s; best = c; } });
      if (!best || bs < 60) {
        linked.forEach(function (c) { var s = _cpScore(l.name + ' ' + l.spec, c.name); if (s > bs) { bs = s; best = c; } });
      }
      l.match = bs >= 60 ? best.code : ''; l.score = bs;
    });
    return { success: true, quoteNo: _cpStr(p.quoteNo), date: date, client: client, lines: out, candidates: cand, linked: linked };
  } catch (e) {
    Logger.log('[apiCpQuoteLines] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}
/** 見積の品名と構成表の品名の近さ（100=同じ） */
function _cpScore(a, b) {
  var x = _cpNameKey(a), y = _cpNameKey(b);
  if (!x || !y) return 0;
  if (x === y || x.indexOf(y) >= 0 || y.indexOf(x) >= 0) return x === y ? 100 : 85;
  var bx = _aqBoards(a), by = _aqBoards(b);
  if (!bx.length || bx.slice().sort().join() !== by.slice().sort().join()) return 0;
  var pcb = /PCB/.test(x) === /PCB/.test(y), done = /\(完\)/.test(x) === /\(完\)/.test(y);
  return pcb && done ? 70 : 30;
}
