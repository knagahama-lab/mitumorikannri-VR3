// ============================================================
// 36_price_table.gs
// 価格一覧表（品名ごとの 現在単価／旧単価／弊社見積／客先部品コード）
//
//   見積書シートの全明細を 分類×品名×仕様 でまとめ、見積日の新しい順に並べて
//     現在単価 … 最新の見積の単価
//     旧単価   … 現在単価と異なる直近の単価（同額が続く場合は「変更なし」）
//   客先部品コードは 客先部品コードマスタ（34）の「弊社見積No」から引き当て、
//   紐づけが無いものは 型番・分類の一致で推定（guessed=true）する。
// ============================================================

function _ptNorm(s) { return String(s == null ? '' : s).normalize('NFKC').toUpperCase().replace(/[\s　]/g, ''); }
function _ptTokens(s) {
  var t = String(s || '').normalize('NFKC').toUpperCase();
  return (t.match(/[A-Z0-9][A-Z0-9+\-]*[0-9][A-Z0-9+\-]*/g) || [])
    .map(function(x) { return x.replace(/^[-+]+|[-+]+$/g, ''); })
    .filter(function(x) { return x.length >= 4 && !/^\d{1,4}$/.test(x); });
}
function _ptDateKey(v) { var s = _toDateStr(v); return s ? String(s).replace(/-/g, '/') : ''; }

function apiPriceTable() {
  try {
    var ss = getSpreadsheet();
    var qs = ss.getSheetByName(CONFIG.SHEET_QUOTES);
    if (!qs || qs.getLastRow() <= 1) return { success: true, rows: [] };
    var data = qs.getRange(2, 1, qs.getLastRow() - 1, QUOTE_COLS.FOLDER_URL).getValues();

    // 取消・失注の見積は現在単価の対象外（履歴には残す）
    var dead = {};
    getAllMgmtData().forEach(function(r) {
      var st = String(r[MGMT_COLS.STATUS - 1] || '');
      if (st === 'キャンセル' || st === '失注') dead[String(r[MGMT_COLS.ID - 1])] = true;
    });

    var groups = {};
    data.forEach(function(r) {
      var name = String(r[QUOTE_COLS.ITEM_NAME - 1] || '').trim();
      if (!name) return;
      var qty = Number(r[QUOTE_COLS.QTY - 1]) || 0;
      var amount = Number(r[QUOTE_COLS.AMOUNT - 1]) || 0;
      var price = Number(r[QUOTE_COLS.UNIT_PRICE - 1]) || 0;
      if (!price && amount) price = qty > 1 ? Math.round(amount / qty) : amount;
      if (!price) return;
      var spec = String(r[QUOTE_COLS.SPEC - 1] || '').trim();
      var cat = _classifyPriceItem(name.normalize('NFKC')); // 全角ＰＣＢ等も同じ分類に
      var key = cat + '|' + _ptNorm(name) + '|' + _ptNorm(spec);
      var mgmtId = String(r[QUOTE_COLS.MGMT_ID - 1] || '');
      (groups[key] = groups[key] || { category: cat, name: name, spec: spec, hist: [] }).hist.push({
        name: name, spec: spec, price: price, quoteNo: String(r[QUOTE_COLS.QUOTE_NO - 1] || ''), mgmtId: mgmtId,
        date: _ptDateKey(r[QUOTE_COLS.ISSUE_DATE - 1]), client: String(r[QUOTE_COLS.DEST_COMPANY - 1] || ''),
        pdfUrl: String(r[QUOTE_COLS.PDF_URL - 1] || ''), dead: !!dead[mgmtId],
      });
    });

    // 客先部品コード：弊社見積No → マスタ行
    var byQuote = {}, masters = [];
    var cpmSheet = ss.getSheetByName(typeof CPM_SHEET !== 'undefined' ? CPM_SHEET : '客先品番マスタ');
    if (cpmSheet && cpmSheet.getLastRow() > 1) {
      cpmSheet.getRange(2, 1, cpmSheet.getLastRow() - 1, 20).getValues().forEach(function(r) {
        var code = String(r[1] || '').trim();
        if (!code) return;
        var m = { client: String(r[0]), code: code, name: String(r[2]), drawing: String(r[3]), quoteNo: String(r[6] || '').trim(),
                  cat: _classifyPriceItem(String(r[2]).normalize('NFKC')), tokens: _ptTokens(r[2] + ' ' + r[3]), lastPrice: Number(r[9]) || 0 };
        masters.push(m);
        if (m.quoteNo) (byQuote[m.quoteNo] = byQuote[m.quoteNo] || []).push(m);
      });
    }
    var shares = function(a, b) { return a.some(function(t) { return b.indexOf(t) >= 0; }); };

    var rows = Object.keys(groups).map(function(k) {
      var g = groups[k];
      var h = g.hist.sort(function(a, b) { return b.date.localeCompare(a.date) || b.quoteNo.localeCompare(a.quoteNo); });
      var live = h.filter(function(x) { return !x.dead; });
      var cur = live[0] || h[0];
      var old = null;
      for (var i = h.indexOf(cur) + 1; i < h.length; i++) { if (h[i].price !== cur.price) { old = h[i]; break; } }
      var quoteNos = {};
      h.forEach(function(x) { if (x.quoteNo) quoteNos[x.quoteNo] = true; });
      var tokens = _ptTokens(g.name + ' ' + g.spec);

      // 見積の紐づけから（同じ見積に複数品番があれば 分類＋型番で絞る）
      var codes = [], seen = {};
      Object.keys(quoteNos).forEach(function(q) {
        var cand = (byQuote[q] || []).filter(function(m) { return m.cat === g.category; });
        if (cand.length > 1 && tokens.length) cand = cand.filter(function(m) { return shares(m.tokens, tokens); });
        cand.forEach(function(m) { if (!seen[m.client + m.code]) { seen[m.client + m.code] = 1; codes.push({ code: m.code, client: m.client, name: m.name, guessed: false }); } });
      });
      // 紐づけが無ければ 型番(6文字以上)＋分類の一致で推定
      if (!codes.length && tokens.length) {
        var strong = tokens.filter(function(t) { return t.length >= 6; });
        masters.forEach(function(m) {
          if (codes.length >= 3 || m.cat !== g.category || !strong.length || !shares(m.tokens, strong) || seen[m.client + m.code]) return;
          seen[m.client + m.code] = 1;
          codes.push({ code: m.code, client: m.client, name: m.name, guessed: true });
        });
      }
      return {
        category: g.category, name: cur.name, spec: cur.spec, count: Object.keys(quoteNos).length,
        cur: cur ? { price: cur.price, quoteNo: cur.quoteNo, mgmtId: cur.mgmtId, date: cur.date, client: cur.client, pdfUrl: cur.pdfUrl } : null,
        old: old ? { price: old.price, quoteNo: old.quoteNo, mgmtId: old.mgmtId, date: old.date, pdfUrl: old.pdfUrl } : null,
        codes: codes,
      };
    });
    var order = { '基板': 0, 'PCB': 1, 'その他': 2 };
    rows.sort(function(a, b) { return (order[a.category] - order[b.category]) || a.name.localeCompare(b.name, 'ja'); });
    return { success: true, rows: rows, generatedAt: nowJST() };
  } catch (e) {
    Logger.log('[apiPriceTable] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}
