// ============================================================
// 37_quote_partcodes.gs
// 見積書 ⇔ 客先部品コード（見積書一覧の「客先部品コード」列）
//
//   apiCpqList()         … 見積No.ごとの登録済み部品コードと、そのコードの注文明細（発注チェック用）
//   apiCpqRegister(p)    … 見積に部品コードを登録（マスタに無いコードは新規作成して登録）
//   apiCpqUnregister(p)  … 見積から部品コードの登録を解除
// 部品コード⇔見積の関係は 客先品番マスタ の「弊社見積No」列で保持する（1コード＝1見積）。
// ============================================================

function apiCpqList() {
  try {
    var byQuote = {}, all = [];
    var lines = {};
    _cpmOrderLines().forEach(function(l) {
      if (!l.code || !l.orderNo) return;
      (lines[l.clientKey + '|' + l.code] = lines[l.clientKey + '|' + l.code] || []).push(
        { orderNo: l.orderNo, orderDate: l.orderDate, mgmtId: l.mgmtId, qty: l.qty, price: l.price, orderPdfUrl: l.orderPdfUrl });
    });
    _cpmRows().forEach(function(r) {
      var m = { client: String(r[CPM.CLIENT]), code: String(r[CPM.CODE]), name: String(r[CPM.NAME]), drawing: String(r[CPM.DRAWING]),
                quoteNo: String(r[CPM.QUOTE_NO] || '').trim(), variant: String(r[CPM.VARIANT] || '') };
      all.push(m);
      if (!m.quoteNo) return;
      var ol = (lines[m.client + '|' + _cpmCode(m.code)] || []).slice()
        .sort(function(a, b) { return String(b.orderDate).localeCompare(String(a.orderDate)); });
      var seen = {};
      ol = ol.filter(function(o) { if (seen[o.orderNo]) return false; seen[o.orderNo] = 1; return true; }).slice(0, 30);
      (byQuote[m.quoteNo] = byQuote[m.quoteNo] || []).push(Object.assign({ orders: ol }, m));
    });
    return { success: true, byQuote: byQuote, all: all };
  } catch (e) {
    Logger.log('[apiCpqList] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

/** p: { quoteNo, client, code, name } */
function apiCpqRegister(p) {
  var quoteNo = String((p && p.quoteNo) || '').trim();
  var code = _cpmCode((p && p.code) || '');
  if (!quoteNo) return { success: false, error: '見積番号がありません' };
  if (!code) return { success: false, error: '客先部品コードを入力してください' };
  var clientKey = _cpmClientKey(p.client || '');
  if (clientKey === '(取引先不明)') return { success: false, error: '取引先を判定できません。取引先を選択してください' };

  var lock = LockService.getScriptLock();
  var created = false;
  try {
    lock.waitLock(20000);
    var idx = _cpmIndex(_cpmRows());
    if (idx[clientKey + '|' + code] === undefined) {
      var name = _cpmHalf(String(p.name || '')).replace(/\s+/g, ' ').trim();
      var sp = _cpmSplit({ itemName: name, partCode: code });
      var r = new Array(CPM_HEADERS.length).fill('');
      r[CPM.CLIENT] = clientKey; r[CPM.CODE] = code; r[CPM.NAME] = sp.name; r[CPM.DRAWING] = sp.drawing; r[CPM.VARIANT] = sp.variant;
      r[CPM.COUNT] = 0; r[CPM.CONFIRMED] = 'TRUE'; r[CPM.NOTE] = '見積書一覧から登録'; r[CPM.UPDATED_AT] = nowJST();
      var sh = _cpmSheet();
      sh.getRange(sh.getLastRow() + 1, 1, 1, CPM_HEADERS.length).setNumberFormat('@').setValues([r]);
      created = true;
    }
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
  var res = apiCpmLinkQuote({ quoteNo: quoteNo, items: [{ client: clientKey, code: code }] });
  if (!res.success) return res;
  return { success: true, created: created, item: res.items[0] };
}

/** p: { client, code } */
function apiCpqUnregister(p) {
  return apiCpmLinkQuote({ quoteNo: '', unlink: true, items: [{ client: p.client, code: p.code }] });
}
