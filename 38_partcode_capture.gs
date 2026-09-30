// ============================================================
// 38_partcode_capture.gs
// 客先部品コードの取りこぼし対策（登録を増やす）
//
//   apiCpmCaptureStatus()     … 注文明細のうち部品コード取得済み／未取得の件数
//   apiCpmReocrPending(p)     … 部品コード未取得の注文PDFを再OCRしてコードを補完（分割実行）
//   apiCpmManualAdd(p)        … 手動で1件登録（登録済みは登録しない）
//   apiCpmSetLineCode(p)      … 注文明細の1行に部品コードを手入力 → マスタへ反映
//   _cpmFindLooseIdx()        … 取引先が自社・不明で登録された同じコードを探す（34から利用）
//
// 注文書シート21列目（客先部品コード）の「-」は「再読取しても部品コードが無かった」印。
// ============================================================

/** 取引先が自社・不明で登録されている同じコードの行（1件だけの場合）の位置。無ければ -1 */
function _cpmFindLooseIdx(rows, code, clientKey) {
  code = _cpmCode(code);
  var hits = [];
  rows.forEach(function(r, i) {
    if (_cpmCode(r[CPM.CODE]) !== code || r[CPM.CLIENT] === clientKey) return;
    var bad = (typeof _cpmIsBadClient === 'function') ? _cpmIsBadClient(r[CPM.CLIENT]) : r[CPM.CLIENT] === '(取引先不明)';
    if (bad) hits.push(i);
  });
  return hits.length === 1 ? hits[0] : -1;
}

function _cpmPendingOrders(lines) {
  var pend = {};
  lines.forEach(function(l) {
    if (l.code || l.reocrDone || !l.orderPdfUrl || !l.mgmtId) return;
    (pend[l.mgmtId] = pend[l.mgmtId] || { mgmtId: l.mgmtId, orderNo: l.orderNo, orderDate: l.orderDate, pdfUrl: l.orderPdfUrl, rows: [] }).rows.push(l);
  });
  return Object.keys(pend).map(function(k) { return pend[k]; })
    .sort(function(a, b) { return String(b.orderDate).localeCompare(String(a.orderDate)); }); // 新しい注文から
}

function apiCpmCaptureStatus() {
  try {
    var lines = _cpmOrderLines();
    var withCode = lines.filter(function(l) { return l.code; }).length;
    var confirmedNone = lines.filter(function(l) { return !l.code && l.reocrDone; }).length;
    var pending = _cpmPendingOrders(lines);
    return { success: true, lines: lines.length, withCode: withCode, noCode: lines.length - withCode,
             confirmedNone: confirmedNone, pendingOrders: pending.length,
             pendingLines: pending.reduce(function(s, o) { return s + o.rows.length; }, 0) };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * 部品コード未取得の注文PDFを再OCR（新しい注文から p.max 件、既定10件。実行時間4分半で打ち切り）
 * 読み取れたコードを注文書シートへ書き込み、最後にマスタを再集計する。
 */
function apiCpmReocrPending(p) {
  try {
    var max = Math.min(Number(p && p.max) || 10, 40);
    var started = Date.now();
    var os = getSpreadsheet().getSheetByName(CONFIG.SHEET_ORDERS);
    var pending = _cpmPendingOrders(_cpmOrderLines());
    var processed = 0, found = 0, failed = [], details = [];
    for (var i = 0; i < pending.length && processed < max; i++) {
      if (Date.now() - started > 270000) break;
      var o = pending[i];
      processed++;
      var fid = (String(o.pdfUrl).match(/[-\w]{25,}/) || [])[0];
      var ocr = null;
      try { ocr = fid ? extractPdfData(DriveApp.getFileById(fid), 'order') : null; } catch (e) { ocr = null; }
      if (!ocr || !ocr.lineItems) { failed.push(o.orderNo || o.mgmtId); continue; }
      var items = ocr.lineItems;
      var rows = o.rows.slice().sort(function(a, b) { return a.lineNo - b.lineNo; });
      var got = 0;
      rows.forEach(function(l, k) {
        var it = (items.length === rows.length) ? items[k] : null;
        if (!it) it = items.filter(function(x) { return Number(x.qty) === l.qty && Number(x.unitPrice) === l.price; })[0] || null;
        if (!it) it = items.filter(function(x) { return l.name && _cpmCode(x.itemName).indexOf(_cpmCode(l.name).substring(0, 8)) >= 0; })[0] || null;
        var code = it ? _cpmSplit(it).code : '';
        os.getRange(l.row, CPM_ORDER_CODE_COL).setNumberFormat('@').setValue(code || '-');
        if (code) { got++; found++; }
      });
      details.push({ orderNo: o.orderNo, lines: rows.length, found: got });
    }
    var rb = found ? apiCpmRebuild() : { created: 0, updated: 0 };
    return { success: true, processed: processed, found: found, failed: failed, details: details,
             created: rb.created || 0, updated: rb.updated || 0, remaining: Math.max(0, pending.length - processed) };
  } catch (e) {
    Logger.log('[apiCpmReocrPending] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

/** 手動で1件登録（取引先×部品コードが登録済みなら登録しない） */
function apiCpmManualAdd(p) {
  p = p || {};
  var res = apiCpmImport({
    rows: [{ '取引先': p.client || '', '客先部品コード': p.code || '', '客先品名': p.name || '', '図番': p.drawing || '',
             '機種コード': p.modelCode || '', '区分': p.variant || '', '弊社見積No': p.quoteNo || '', '見積単価': p.quotePrice || '',
             '弊社基板名': p.boardName || '', '備考': p.note || '' }],
    defaultClient: p.client || '', confirmed: true, dryRun: false,
  });
  if (!res.success) return res;
  if (res.added) return { success: true, item: res.newItems[0] };
  if (res.skippedCount) return { success: false, error: '部品コード ' + p.code + ' は ' + (res.skipped[0] || {}).client + ' で登録済みです' };
  return { success: false, error: (res.errors[0] || {}).message || '登録できませんでした' };
}

/** 注文明細の1行に部品コードを設定（p.none=true で「部品コード無し」を記録） → マスタへ反映 */
function apiCpmSetLineCode(p) {
  try {
    var code = p.none ? '-' : _cpmCode(p.code || '');
    if (!code) return { success: false, error: '部品コードを入力してください' };
    var os = getSpreadsheet().getSheetByName(CONFIG.SHEET_ORDERS);
    var n = 0;
    _cpmOrderLines().forEach(function(l) {
      if (l.mgmtId === String(p.mgmtId) && l.lineNo === Number(p.lineNo)) {
        os.getRange(l.row, CPM_ORDER_CODE_COL).setNumberFormat('@').setValue(code); n++;
      }
    });
    if (!n) return { success: false, error: '注文明細が見つかりません' };
    var rb = apiCpmRebuild();
    return { success: true, updatedRows: n, created: rb.created || 0 };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
