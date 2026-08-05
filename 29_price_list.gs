// ============================================================
// 29_price_list.gs
// 価格表（基板価格・PCB価格・その他価格）
//
// 見積書シート（QUOTE_COLS）の明細行から、品名に含まれるキーワードで
// 基板／PCB／その他に自動分類し、品名+仕様ごとに最新の見積単価を
// 集計してシートに保持する。原価・粗利は手入力で、原価が見積単価を
// 上回った場合（逆ざや）はフラグを立てて画面上で警告表示する。
//
// 自動更新: 見積書PDFが新規登録・差し替えされた際に、その案件
// (mgmtId) 分の明細のみを差分更新する（_upsertPriceListForMgmt）。
// 手動更新: 画面上の「🔄 価格表を再集計」ボタンで全件再集計できる
// （apiPriceListRebuild）。
// ============================================================

var PRICE_LIST_SHEET = '価格表';
var PRICE_LIST_CATEGORIES = ['基板', 'PCB', 'その他'];

function _initPriceListSheet() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(PRICE_LIST_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PRICE_LIST_SHEET);
    var h  = ['RowID', 'カテゴリ', '品名', '仕様', '見積単価', '見積番号', 'mgmtId', 'PDF URL', '原価', '粗利額', '粗利率(%)', '逆ざや', '見積日', '更新日時'];
    var hr = sheet.getRange(1, 1, 1, h.length);
    hr.setValues([h]);
    hr.setBackground('#FFF3E0');
    hr.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(3, 220);
  }
  return sheet;
}

// 品名から 基板／PCB／その他 を自動判定
function _classifyPriceItem(itemName) {
  var s = String(itemName || '');
  if (/PCB/i.test(s)) return 'PCB';
  if (s.indexOf('基板') >= 0) return '基板';
  return 'その他';
}

// 品名+仕様ごとの最新見積単価データを 価格表 シートへ upsert する（手入力の原価は保持）
function _upsertPriceListRows(items) {
  if (!items || !items.length) return { appended: 0, updated: 0 };
  var sheet = _initPriceListSheet();
  var last  = sheet.getLastRow();
  var existingRows = last > 1 ? sheet.getRange(2, 1, last - 1, 14).getValues() : [];
  var keyToIdx = {};
  existingRows.forEach(function(r, i) {
    keyToIdx[String(r[1]) + '|' + String(r[2]) + '|' + String(r[3])] = i;
  });
  var now = nowJST();
  var appended = 0, updated = 0;
  items.forEach(function(it) {
    var key = it.category + '|' + it.itemName + '|' + it.spec;
    if (keyToIdx.hasOwnProperty(key)) {
      var idx = keyToIdx[key];
      var row = idx + 2;
      var cost = existingRows[idx][8]; // 手入力の原価は維持する
      var hasCost = cost !== '' && cost !== null && cost !== undefined;
      var margin = hasCost ? (it.unitPrice - Number(cost)) : '';
      var marginRate = (hasCost && it.unitPrice) ? Math.round((margin / it.unitPrice) * 1000) / 10 : '';
      var inverted = hasCost && Number(cost) > it.unitPrice;
      sheet.getRange(row, 1, 1, 14).setValues([[
        existingRows[idx][0], it.category, it.itemName, it.spec, it.unitPrice, it.quoteNo, it.mgmtId, it.pdfUrl,
        cost, margin, marginRate, inverted ? '⚠️逆ざや' : '', it.issueDate, now,
      ]]);
      updated++;
    } else {
      var id = 'PL-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + Math.floor(Math.random() * 1000);
      sheet.appendRow([id, it.category, it.itemName, it.spec, it.unitPrice, it.quoteNo, it.mgmtId, it.pdfUrl, '', '', '', '', it.issueDate, now]);
      appended++;
    }
  });
  return { appended: appended, updated: updated };
}

// 見積書シートの明細行を items[] へ変換（unitPrice=0 や品名なしの行は除外）
function _quoteLinesToPriceItems(rows) {
  var map = {}; // 同一 品名+仕様 は最新の見積日のものだけ残す
  rows.forEach(function(r) {
    var itemName = String(r[QUOTE_COLS.ITEM_NAME - 1] || '').trim();
    var unitPrice = Number(r[QUOTE_COLS.UNIT_PRICE - 1] || 0);
    if (!itemName || !unitPrice) return;
    var spec = String(r[QUOTE_COLS.SPEC - 1] || '').trim();
    var category = _classifyPriceItem(itemName);
    var issueDateRaw = r[QUOTE_COLS.ISSUE_DATE - 1];
    var d = issueDateRaw instanceof Date ? issueDateRaw.getTime() : 0;
    var key = category + '|' + itemName + '|' + spec;
    var existing = map[key];
    if (existing && existing._d > d) return;
    map[key] = {
      category: category, itemName: itemName, spec: spec, unitPrice: unitPrice,
      quoteNo: String(r[QUOTE_COLS.QUOTE_NO - 1] || ''),
      mgmtId: String(r[QUOTE_COLS.MGMT_ID - 1] || ''),
      pdfUrl: String(r[QUOTE_COLS.PDF_URL - 1] || ''),
      issueDate: _toDateStr(issueDateRaw),
      _d: d,
    };
  });
  return Object.keys(map).map(function(k) { return map[k]; });
}

// 特定の案件(mgmtId)の明細だけ差分更新（見積書の新規登録・PDF差し替え時に呼ぶ）
function _upsertPriceListForMgmt(mgmtId) {
  try {
    if (!mgmtId) return;
    var ss = getSpreadsheet();
    var quoteSheet = ss.getSheetByName(CONFIG.SHEET_QUOTES);
    if (!quoteSheet || quoteSheet.getLastRow() <= 1) return;
    var last = quoteSheet.getLastRow();
    var data = quoteSheet.getRange(2, 1, last - 1, 15).getValues()
      .filter(function(r) { return String(r[QUOTE_COLS.MGMT_ID - 1]) === String(mgmtId); });
    var items = _quoteLinesToPriceItems(data);
    if (items.length) _upsertPriceListRows(items);
  } catch (e) { Logger.log('[_upsertPriceListForMgmt] ' + e.message); }
}

// 全件再集計（見積書シート全体をスキャン）
function apiPriceListRebuild() {
  try {
    var ss = getSpreadsheet();
    var quoteSheet = ss.getSheetByName(CONFIG.SHEET_QUOTES);
    if (!quoteSheet || quoteSheet.getLastRow() <= 1) return { success: true, appended: 0, updated: 0 };
    var last = quoteSheet.getLastRow();
    var data = quoteSheet.getRange(2, 1, last - 1, 15).getValues();
    var items = _quoteLinesToPriceItems(data);
    var res = _upsertPriceListRows(items);
    return { success: true, appended: res.appended, updated: res.updated, total: items.length };
  } catch (e) { return { success: false, error: e.message }; }
}

// 一覧取得
function apiPriceListGet(payload) {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(PRICE_LIST_SHEET);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, rows: [] };
    var last = sheet.getLastRow();
    var data = sheet.getRange(2, 1, last - 1, 14).getValues();
    var category = (payload || {}).category;
    var rows = data
      .filter(function(r) { return !category || String(r[1]) === category; })
      .map(function(r) {
        var cost = (r[8] === '' || r[8] === null) ? '' : Number(r[8]);
        return {
          id:         String(r[0] || ''),
          category:   String(r[1] || ''),
          itemName:   String(r[2] || ''),
          spec:       String(r[3] || ''),
          unitPrice:  Number(r[4] || 0),
          quoteNo:    String(r[5] || ''),
          mgmtId:     String(r[6] || ''),
          pdfUrl:     String(r[7] || ''),
          cost:       cost,
          margin:     (r[9]  === '' || r[9]  === null) ? '' : Number(r[9]),
          marginRate: (r[10] === '' || r[10] === null) ? '' : Number(r[10]),
          inverted:   !!r[11],
          quoteDate:  _toDateStr(r[12]),
          updatedAt:  _toDateStr(r[13]),
        };
      });
    return { success: true, rows: rows };
  } catch (e) { return { success: false, error: e.message }; }
}

// 原価（手入力）の保存。粗利額・粗利率・逆ざやフラグはサーバ側で再計算する。
function apiPriceListSaveCost(payload) {
  try {
    payload = payload || {};
    var id = String(payload.id || '').trim();
    if (!id) return { success: false, error: 'IDが必要です' };
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(PRICE_LIST_SHEET);
    if (!sheet) return { success: false, error: 'シートがありません' };
    var last = sheet.getLastRow();
    if (last <= 1) return { success: false, error: 'データなし' };
    var ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat().map(function(v) { return String(v); });
    var idx = ids.indexOf(id);
    if (idx < 0) return { success: false, error: 'IDが見つかりません' };
    var row = idx + 2;

    var hasCost = payload.cost !== '' && payload.cost !== undefined && payload.cost !== null;
    var cost = hasCost ? Number(payload.cost) : '';
    var unitPrice = Number(sheet.getRange(row, 5).getValue() || 0);
    var margin = hasCost ? (unitPrice - cost) : '';
    var marginRate = (hasCost && unitPrice) ? Math.round((margin / unitPrice) * 1000) / 10 : '';
    var inverted = hasCost && cost > unitPrice;

    sheet.getRange(row, 9, 1, 4).setValues([[cost, margin, marginRate, inverted ? '⚠️逆ざや' : '']]);
    sheet.getRange(row, 14).setValue(nowJST());
    return { success: true, margin: margin, marginRate: marginRate, inverted: inverted };
  } catch (e) { return { success: false, error: e.message }; }
}
