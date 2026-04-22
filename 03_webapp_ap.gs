// ============================================================
// ★ 安全版：案件データ取得（エラー回避処理済み）
// ============================================================
function _apiGetAll() {
  try {
    var rows = getAllMgmtData();

    // IS_LATEST フィルター
    rows = rows.filter(function(r) {
      var colIdx = (MGMT_COLS.IS_LATEST || 0) - 1;
      if (colIdx < 0 || colIdx >= r.length) return true;
      var v = String(r[colIdx] || '');
      return v === '' || v.toUpperCase() === 'TRUE';
    });

    // 非表示ステータスを除外
    rows = rows.filter(function(r) {
      var hidden = CONFIG.STATUS_HIDDEN || [];
      return hidden.indexOf(String(r[MGMT_COLS.STATUS - 1] || '')) < 0;
    });

    // 注文番号がある行のみ
    var orderRows = rows.filter(function(r) {
      return String(r[MGMT_COLS.ORDER_NO - 1]).trim() !== '';
    });

    var items = _deduplicateMgmtRows(orderRows);

    // ★追加：検索用に注文書の明細データを安全に取得する
    var ss = getSpreadsheet();
    var os = ss.getSheetByName(CONFIG.SHEET_ORDERS);
    if (os && os.getLastRow() > 1) {
      // getRangeではなく getDataRange() を使うことで列数エラーを完全回避
      var orderData = os.getDataRange().getValues();
      orderData.shift(); // 1行目のヘッダーを除外

      items.forEach(function(item) {
        var lines = orderData.filter(function(r) { return String(r[0]) === item.id; });
        if (lines.length > 0) {
          item.detailText = JSON.stringify(lines);
        }
      });
    }

    // 発注日の新しい順
    items.sort(function(a, b) {
      var da = String(a.orderDate || a.quoteDate || '');
      var db = String(b.orderDate || b.quoteDate || '');
      return db.localeCompare(da);
    });

    return { success: true, total: items.length, items: items };
  } catch(e) {
    Logger.log('[_apiGetAll ERROR] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}


// ============================================================
// ★ 安全版：見積書一覧 API（エラー回避処理済み）
// ============================================================
function _apiQuoteListGetAll() {
  try {
    var ss         = getSpreadsheet();
    var quoteSheet = ss.getSheetByName(CONFIG.SHEET_QUOTES);
    var mgmtData  = getAllMgmtData();
    
    // 見積番号がある行だけを抽出
    var allRows   = mgmtData.filter(function(r) { return String(r[MGMT_COLS.QUOTE_NO - 1]).trim() !== ''; });
    
    // 重複を排除
    var seenQNo  = {};
    var quoteRows = allRows.filter(function(r) {
      var qNo = String(r[MGMT_COLS.QUOTE_NO - 1]).trim();
      if (seenQNo[qNo]) return false;
      seenQNo[qNo] = true;
      return true;
    });

    var quoteLineMap = {};
    if (quoteSheet && quoteSheet.getLastRow() > 1) {
      // getDataRange() で列数エラーを回避
      var qData = quoteSheet.getDataRange().getValues();
      qData.shift(); // 1行目のヘッダーを除外

      qData.forEach(function(r) {
        var mgmtId = String(r[QUOTE_COLS.MGMT_ID - 1] || '');
        if (!mgmtId) return;
        if (!quoteLineMap[mgmtId]) {
          quoteLineMap[mgmtId] = {
            issueDate:   _toDateStr(r[QUOTE_COLS.ISSUE_DATE  - 1]),
            destCompany: String(r[QUOTE_COLS.DEST_COMPANY - 1] || ''),
            destPerson:  String(r[QUOTE_COLS.DEST_PERSON  - 1] || ''),
            lines: [] // 明細行を保存
          };
        }
        quoteLineMap[mgmtId].lines.push(r);
      });
    }

    var items = quoteRows.map(function(r) {
      var mgmtId  = String(r[MGMT_COLS.ID - 1] || '');
      var lineInfo = quoteLineMap[mgmtId] || { lines: [] };
      return {
        id:          mgmtId,
        quoteNo:     String(r[MGMT_COLS.QUOTE_NO - 1]      || ''),
        issueDate:   lineInfo.issueDate   || _toDateStr(r[MGMT_COLS.QUOTE_DATE - 1]),
        destCompany: lineInfo.destCompany || String(r[MGMT_COLS.CLIENT - 1] || ''),
        destPerson:  lineInfo.destPerson  || '',
        quoteAmount: _toNum(r[MGMT_COLS.QUOTE_AMOUNT - 1]),
        status:      String(r[MGMT_COLS.STATUS - 1]        || ''),
        quotePdfUrl: String(r[MGMT_COLS.QUOTE_PDF_URL - 1] || ''),
        orderNo:     String(r[MGMT_COLS.ORDER_NO - 1]      || ''),
        linked:      _isLinkedVal(r[MGMT_COLS.LINKED - 1]),
        orderType:   String(r[MGMT_COLS.ORDER_TYPE - 1]    || ''),
        modelCode:   String(r[MGMT_COLS.MODEL_CODE - 1]    || ''),
        subject:     String(r[MGMT_COLS.SUBJECT - 1]       || ''),
        // ★検索用に明細データを隠し持たせる
        detailText:  JSON.stringify(lineInfo.lines)
      };
    });

    items.sort(function(a, b) {
      var da = String(a.issueDate || a.quoteDate || '');
      var db = String(b.issueDate || b.quoteDate || '');
      return db.localeCompare(da);
    });

    return { success: true, total: items.length, items: items };
  } catch(e) { return { success: false, error: e.message }; }
}