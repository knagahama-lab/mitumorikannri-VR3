// ============================================================
// 見積書・注文書管理システム
// ファイル 11: DBスキーマ正規化 & ナレッジDB構築
// ============================================================
//
// 【依存関係】（既存ファイルの関数を使用）
//   01 config and setup.gs : CONFIG, MGMT_COLS, QUOTE_COLS, ORDER_COLS
//                            getSpreadsheet(), nowJST(), _toDateStr(), getAllMgmtData()
//   03_webapp_ap.gs        : _isLinkedVal()
//
// 【追加するシート】
//   ・見積明細DB   — 見積書の明細行を1行1レコードで管理
//   ・注文明細DB   — 注文書の明細行を1行1レコードで管理
//   ・単価マスタ   — 品名別・顧客別の単価履歴（ナレッジ化の核心）
//   ・案件サマリ   — 全文検索高速化用の統合ビュー
//
// 【セットアップ手順】
//   1. GASプロジェクトにこのファイルを追加（既存ファイルの変更不要）
//   2. initialSetupNormalized() を実行 → 4シートが自動作成 + 初回同期
//   3. 以降は毎時トリガーで自動同期される
// ============================================================

// ===== 正規化DB シート名 =====
var SHEET_QUOTE_DETAIL  = '見積明細DB';
var SHEET_ORDER_DETAIL  = '注文明細DB';
var SHEET_UNIT_PRICE    = '単価マスタ';
var SHEET_CASE_SUMMARY  = '案件サマリ';

// ===== 見積明細DB 列定義 (18列) =====
var QUOTE_DETAIL_COLS = {
  ROW_ID       : 1,
  MGMT_ID      : 2,
  QUOTE_NO     : 3,
  ISSUE_DATE   : 4,
  DEST_COMPANY : 5,
  DEST_PERSON  : 6,
  LINE_NO      : 7,
  ITEM_NAME    : 8,
  SPEC         : 9,
  QTY          : 10,
  UNIT         : 11,
  UNIT_PRICE   : 12,
  AMOUNT       : 13,
  REMARKS      : 14,
  PDF_URL      : 15,
  STATUS       : 16,
  ORDER_NO     : 17,
  SYNCED_AT    : 18,
};

// ===== 注文明細DB 列定義 (22列) =====
var ORDER_DETAIL_COLS = {
  ROW_ID          : 1,
  MGMT_ID         : 2,
  ORDER_NO        : 3,
  LINKED_QUOTE    : 4,
  ORDER_TYPE      : 5,
  ORDER_DATE      : 6,
  MODEL_CODE      : 7,
  ORDER_SLIP_NO   : 8,
  LINE_NO         : 9,
  ITEM_NAME       : 10,
  SPEC            : 11,
  FIRST_DELIVERY  : 12,
  DELIVERY_DEST   : 13,
  QTY             : 14,
  UNIT            : 15,
  UNIT_PRICE      : 16,
  AMOUNT          : 17,
  REMARKS         : 18,
  PDF_URL         : 19,
  COMPARE_STATUS  : 20,
  COMPARE_DETAIL  : 21,
  SYNCED_AT       : 22,
};

// ===== 単価マスタ 列定義 (13列) =====
var UNIT_PRICE_COLS = {
  MASTER_ID    : 1,
  ITEM_NAME    : 2,
  ITEM_NAME_RAW: 3,
  SPEC         : 4,
  CLIENT       : 5,
  UNIT_PRICE   : 6,
  UNIT         : 7,
  QUOTE_NO     : 8,
  QUOTE_DATE   : 9,
  ORDER_NO     : 10,
  ORDER_DATE   : 11,
  IS_ORDERED   : 12,
  UPDATED_AT   : 13,
};

// ===== 案件サマリ 列定義 (19列) =====
var CASE_SUMMARY_COLS = {
  MGMT_ID      : 1,
  QUOTE_NO     : 2,
  ORDER_NO     : 3,
  SUBJECT      : 4,
  CLIENT       : 5,
  STATUS       : 6,
  QUOTE_DATE   : 7,
  ORDER_DATE   : 8,
  QUOTE_AMOUNT : 9,
  ORDER_AMOUNT : 10,
  QUOTE_PDF    : 11,
  ORDER_PDF    : 12,
  MODEL_CODE   : 13,
  ORDER_TYPE   : 14,
  DELIVERY_DATE: 15,
  MEMO         : 16,
  ITEM_NAMES   : 17,
  UNIT_PRICES  : 18,
  UPDATED_AT   : 19,
};

// ============================================================
// 初回セットアップ（一度だけ実行）
// ============================================================
function initialSetupNormalized() {
  var ss = getSpreadsheet();
  _createNormalizedSheets(ss);
  syncAllDetailDB();
  rebuildUnitPriceMaster();
  _buildCaseSummary();
  _registerNormalizedTriggers();

  Logger.log('正規化DB初期セットアップ完了');
  try {
    SpreadsheetApp.getUi().alert(
      '✅ 正規化DB構築完了\n\n' +
      '追加されたシート:\n' +
      '  ・見積明細DB\n  ・注文明細DB\n  ・単価マスタ\n  ・案件サマリ\n\n' +
      'トリガー:\n' +
      '  ・syncAllDetailDB → 毎時\n' +
      '  ・rebuildUnitPriceMaster → 毎日3時'
    );
  } catch(e) {}
}

function _createNormalizedSheets(ss) {
  _ensureNormSheet(ss, SHEET_QUOTE_DETAIL, [
    '行ID','管理ID','見積番号','発行日','送り先会社名','送り先担当者',
    '行No','品名','仕様','数量','単位','単価','金額','備考','PDF URL',
    'ステータス','受注番号','同期日時'
  ], '#E3F2FD');

  _ensureNormSheet(ss, SHEET_ORDER_DETAIL, [
    '行ID','管理ID','注文番号','紐づけ見積番号','注文種別','発注日',
    '機種コード','発注伝票番号','行No','品名','仕様',
    '初回納品日','納品先','数量','単位','単価','金額','備考','PDF URL',
    '単価比較結果','差異内容','同期日時'
  ], '#FFF3E0');

  _ensureNormSheet(ss, SHEET_UNIT_PRICE, [
    'マスタID','品名（正規化）','品名（原文）','仕様','顧客名',
    '単価','単位','参照見積番号','見積日','受注番号','受注日',
    '受注実績','更新日時'
  ], '#F3E5F5');

  _ensureNormSheet(ss, SHEET_CASE_SUMMARY, [
    '管理ID','見積番号','注文番号','件名','顧客名','ステータス',
    '見積日','発注日','見積金額','注文金額',
    '見積書PDF','注文書PDF','機種コード','注文種別','納期','メモ',
    '主要品名リスト','単価情報JSON','更新日時'
  ], '#E8F5E9');
}

function _ensureNormSheet(ss, name, headers, color) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    var r = sheet.getRange(1, 1, 1, headers.length);
    r.setValues([headers]).setBackground(color).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ============================================================
// メイン同期処理（毎時トリガー）
// ============================================================
function syncAllDetailDB() {
  Logger.log('[NORM] 明細DB同期開始');
  _syncQuoteDetails();
  _syncOrderDetails();
  _buildCaseSummary();
  Logger.log('[NORM] 明細DB同期完了');
}

// ===== 見積明細DBへの同期 =====
function _syncQuoteDetails() {
  var ss        = getSpreadsheet();
  var srcSheet  = ss.getSheetByName(CONFIG.SHEET_QUOTES);   // '見積書シート'
  var dstSheet  = ss.getSheetByName(SHEET_QUOTE_DETAIL);
  var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);

  if (!srcSheet || srcSheet.getLastRow() <= 1) return;

  var existingIds = _getNormExistingIds(dstSheet);
  var mgmtMap     = _buildMgmtMap(mgmtSheet);

  var srcData = srcSheet.getRange(2, 1, srcSheet.getLastRow() - 1, 15).getValues();
  var newRows = [];

  srcData.forEach(function(row) {
    var mgmtId  = String(row[QUOTE_COLS.MGMT_ID  - 1] || '').trim();
    var quoteNo = String(row[QUOTE_COLS.QUOTE_NO  - 1] || '').trim();
    var lineNo  = row[QUOTE_COLS.LINE_NO - 1];
    if (!mgmtId || !quoteNo) return;

    var rowId = 'QD-' + quoteNo.replace(/[^a-zA-Z0-9\-]/g, '') + '-' + String(lineNo || '0');
    if (existingIds[rowId]) return;

    var mgmt = mgmtMap[mgmtId] || {};
    newRows.push([
      rowId,
      mgmtId,
      quoteNo,
      _toDateStr(row[QUOTE_COLS.ISSUE_DATE    - 1]),
      row[QUOTE_COLS.DEST_COMPANY - 1] || '',
      row[QUOTE_COLS.DEST_PERSON  - 1] || '',
      lineNo,
      row[QUOTE_COLS.ITEM_NAME    - 1] || '',
      row[QUOTE_COLS.SPEC         - 1] || '',
      row[QUOTE_COLS.QTY          - 1] || '',
      row[QUOTE_COLS.UNIT         - 1] || '',
      row[QUOTE_COLS.UNIT_PRICE   - 1] || '',
      row[QUOTE_COLS.AMOUNT       - 1] || '',
      row[QUOTE_COLS.REMARKS      - 1] || '',
      row[QUOTE_COLS.PDF_URL      - 1] || '',
      mgmt.status  || '',
      mgmt.orderNo || '',
      nowJST(),
    ]);
  });

  if (newRows.length > 0) {
    dstSheet.getRange(dstSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
            .setValues(newRows);
    Logger.log('[NORM] 見積明細 ' + newRows.length + '行 追加');
  }

  // ステータス・受注番号の差分更新
  _updateQuoteDetailStatus(dstSheet, mgmtMap);
}

function _updateQuoteDetailStatus(dstSheet, mgmtMap) {
  if (dstSheet.getLastRow() <= 1) return;
  var ncols = Object.keys(QUOTE_DETAIL_COLS).length;
  var data  = dstSheet.getRange(2, 1, dstSheet.getLastRow() - 1, ncols).getValues();
  var updates = [];
  data.forEach(function(row, i) {
    var mgmtId = String(row[QUOTE_DETAIL_COLS.MGMT_ID - 1] || '');
    var mgmt   = mgmtMap[mgmtId];
    if (!mgmt) return;
    var curStatus  = String(row[QUOTE_DETAIL_COLS.STATUS   - 1] || '');
    var curOrderNo = String(row[QUOTE_DETAIL_COLS.ORDER_NO - 1] || '');
    if (curStatus !== mgmt.status || curOrderNo !== mgmt.orderNo) {
      updates.push({ rowIdx: i + 2, status: mgmt.status, orderNo: mgmt.orderNo });
    }
  });
  updates.forEach(function(u) {
    dstSheet.getRange(u.rowIdx, QUOTE_DETAIL_COLS.STATUS  ).setValue(u.status);
    dstSheet.getRange(u.rowIdx, QUOTE_DETAIL_COLS.ORDER_NO).setValue(u.orderNo);
    dstSheet.getRange(u.rowIdx, QUOTE_DETAIL_COLS.SYNCED_AT).setValue(nowJST());
  });
  if (updates.length) Logger.log('[NORM] 見積明細ステータス更新: ' + updates.length + '行');
}

// ===== 注文明細DBへの同期 =====
function _syncOrderDetails() {
  var ss       = getSpreadsheet();
  var srcSheet = ss.getSheetByName(CONFIG.SHEET_ORDERS);   // '注文書シート'
  var dstSheet = ss.getSheetByName(SHEET_ORDER_DETAIL);

  if (!srcSheet || srcSheet.getLastRow() <= 1) return;

  var existingIds = _getNormExistingIds(dstSheet);
  var srcData = srcSheet.getRange(2, 1, srcSheet.getLastRow() - 1, 19).getValues();
  var newRows = [];

  srcData.forEach(function(row) {
    var mgmtId  = String(row[ORDER_COLS.MGMT_ID  - 1] || '').trim();
    var orderNo = String(row[ORDER_COLS.ORDER_NO  - 1] || '').trim();
    var lineNo  = row[ORDER_COLS.LINE_NO - 1];
    if (!mgmtId || !orderNo) return;

    var rowId = 'OD-' + orderNo.replace(/[^a-zA-Z0-9\-]/g, '') + '-' + String(lineNo || '0');
    if (existingIds[rowId]) return;

    newRows.push([
      rowId,
      mgmtId,
      orderNo,
      row[ORDER_COLS.LINKED_QUOTE   - 1] || '',
      row[ORDER_COLS.ORDER_TYPE     - 1] || '',
      _toDateStr(row[ORDER_COLS.ORDER_DATE  - 1]),
      row[ORDER_COLS.MODEL_CODE     - 1] || '',
      row[ORDER_COLS.ORDER_SLIP_NO  - 1] || '',
      lineNo,
      row[ORDER_COLS.ITEM_NAME      - 1] || '',
      row[ORDER_COLS.SPEC           - 1] || '',
      _toDateStr(row[ORDER_COLS.FIRST_DELIVERY - 1]),
      row[ORDER_COLS.DELIVERY_DEST  - 1] || '',
      row[ORDER_COLS.QTY            - 1] || '',
      row[ORDER_COLS.UNIT           - 1] || '',
      row[ORDER_COLS.UNIT_PRICE     - 1] || '',
      row[ORDER_COLS.AMOUNT         - 1] || '',
      row[ORDER_COLS.REMARKS        - 1] || '',
      row[ORDER_COLS.PDF_URL        - 1] || '',
      '未比較',
      '',
      nowJST(),
    ]);
  });

  if (newRows.length > 0) {
    dstSheet.getRange(dstSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
            .setValues(newRows);
    Logger.log('[NORM] 注文明細 ' + newRows.length + '行 追加');
  }
}

// ============================================================
// 案件サマリ構築
// ============================================================
function _buildCaseSummary() {
  var ss            = getSpreadsheet();
  var summarySheet  = ss.getSheetByName(SHEET_CASE_SUMMARY);
  var quoteDetailSh = ss.getSheetByName(SHEET_QUOTE_DETAIL);
  var mgmtSheet     = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);

  var mgmtData = getAllMgmtData();
  if (!mgmtData.length) return;

  var quoteDetailIndex = _buildQuoteDetailIndex(quoteDetailSh);

  var rows = mgmtData.map(function(mgmt) {
    var mgmtId  = String(mgmt[MGMT_COLS.ID - 1] || '');
    var details = quoteDetailIndex[mgmtId] || [];
    var itemNames = details.map(function(d) { return d.itemName; }).filter(Boolean).join(',');
    var priceJson = JSON.stringify(details.map(function(d) {
      return { item: d.itemName, spec: d.spec, qty: d.qty, price: d.unitPrice, amount: d.amount };
    }));
    if (priceJson.length > 50000) priceJson = priceJson.substring(0, 50000);

    return [
      mgmtId,
      String(mgmt[MGMT_COLS.QUOTE_NO      - 1] || ''),
      String(mgmt[MGMT_COLS.ORDER_NO      - 1] || ''),
      String(mgmt[MGMT_COLS.SUBJECT       - 1] || ''),
      String(mgmt[MGMT_COLS.CLIENT        - 1] || ''),
      String(mgmt[MGMT_COLS.STATUS        - 1] || ''),
      _toDateStr(mgmt[MGMT_COLS.QUOTE_DATE    - 1]),
      _toDateStr(mgmt[MGMT_COLS.ORDER_DATE    - 1]),
      mgmt[MGMT_COLS.QUOTE_AMOUNT - 1] || '',
      mgmt[MGMT_COLS.ORDER_AMOUNT - 1] || '',
      String(mgmt[MGMT_COLS.QUOTE_PDF_URL - 1] || ''),
      String(mgmt[MGMT_COLS.ORDER_PDF_URL - 1] || ''),
      String(mgmt[MGMT_COLS.MODEL_CODE    - 1] || ''),
      String(mgmt[MGMT_COLS.ORDER_TYPE    - 1] || ''),
      _toDateStr(mgmt[MGMT_COLS.DELIVERY_DATE - 1]),
      String(mgmt[MGMT_COLS.MEMO          - 1] || ''),
      itemNames,
      priceJson,
      nowJST(),
    ];
  });

  if (!rows.length) return;
  var lastRow = summarySheet.getLastRow();
  if (lastRow > 1) summarySheet.deleteRows(2, lastRow - 1);
  summarySheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  Logger.log('[NORM] 案件サマリ再構築: ' + rows.length + '件');
}

function _buildQuoteDetailIndex(sheet) {
  var index = {};
  if (!sheet || sheet.getLastRow() <= 1) return index;
  var ncols = Object.keys(QUOTE_DETAIL_COLS).length;
  var data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, ncols).getValues();
  data.forEach(function(row) {
    var mgmtId = String(row[QUOTE_DETAIL_COLS.MGMT_ID - 1] || '');
    if (!mgmtId) return;
    if (!index[mgmtId]) index[mgmtId] = [];
    index[mgmtId].push({
      lineNo   : row[QUOTE_DETAIL_COLS.LINE_NO    - 1],
      itemName : String(row[QUOTE_DETAIL_COLS.ITEM_NAME  - 1] || ''),
      spec     : String(row[QUOTE_DETAIL_COLS.SPEC       - 1] || ''),
      qty      : row[QUOTE_DETAIL_COLS.QTY        - 1],
      unitPrice: row[QUOTE_DETAIL_COLS.UNIT_PRICE - 1],
      amount   : row[QUOTE_DETAIL_COLS.AMOUNT     - 1],
    });
  });
  return index;
}

// ============================================================
// 単価マスタ再構築（毎日3時 or 手動実行）
// ============================================================
function rebuildUnitPriceMaster() {
  Logger.log('[UNIT_PRICE] 単価マスタ再構築開始');
  var ss             = getSpreadsheet();
  var masterSheet    = ss.getSheetByName(SHEET_UNIT_PRICE);
  var quoteDetailSh  = ss.getSheetByName(SHEET_QUOTE_DETAIL);
  var mgmtSheet      = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);

  if (!quoteDetailSh || quoteDetailSh.getLastRow() <= 1) {
    Logger.log('[UNIT_PRICE] 見積明細DBが空のためスキップ。先にsyncAllDetailDB()を実行してください');
    return;
  }

  var mgmtMap    = _buildMgmtMap(mgmtSheet);
  var ncols      = Object.keys(QUOTE_DETAIL_COLS).length;
  var detailData = quoteDetailSh.getRange(
    2, 1, quoteDetailSh.getLastRow() - 1, ncols
  ).getValues();

  // 品名(正規化) + 仕様 + 顧客 でグループ化 → 最新単価を採用
  var priceMap = {};
  detailData.forEach(function(row) {
    var itemName  = String(row[QUOTE_DETAIL_COLS.ITEM_NAME  - 1] || '').trim();
    var spec      = String(row[QUOTE_DETAIL_COLS.SPEC       - 1] || '').trim();
    var unitPrice = row[QUOTE_DETAIL_COLS.UNIT_PRICE - 1];
    var mgmtId    = String(row[QUOTE_DETAIL_COLS.MGMT_ID   - 1] || '');
    var quoteNo   = String(row[QUOTE_DETAIL_COLS.QUOTE_NO  - 1] || '');
    var issueDate = String(row[QUOTE_DETAIL_COLS.ISSUE_DATE - 1] || '');
    var unit      = String(row[QUOTE_DETAIL_COLS.UNIT      - 1] || '');
    var orderNo   = String(row[QUOTE_DETAIL_COLS.ORDER_NO  - 1] || '');

    if (!itemName) return;
    var price = Number(unitPrice);
    if (isNaN(price) || price <= 0) return;

    var mgmt   = mgmtMap[mgmtId] || {};
    var client = mgmt.client || '';
    var key    = _normItemName(itemName) + '|' + spec + '|' + client;

    if (!priceMap[key] || issueDate > priceMap[key].issueDate) {
      priceMap[key] = {
        itemNameNorm: _normItemName(itemName),
        itemNameRaw : itemName,
        spec        : spec,
        client      : client,
        unitPrice   : price,
        unit        : unit,
        quoteNo     : quoteNo,
        issueDate   : issueDate,
        orderNo     : orderNo,
        orderDate   : mgmt.orderDate || '',
        isOrdered   : !!orderNo,
      };
    }
  });

  var keys      = Object.keys(priceMap);
  var masterRows = keys.map(function(key, i) {
    var p = priceMap[key];
    return [
      'UP-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd') + '-' + String(i + 1).padStart(4, '0'),
      p.itemNameNorm, p.itemNameRaw, p.spec, p.client,
      p.unitPrice, p.unit, p.quoteNo, p.issueDate,
      p.orderNo, p.orderDate, p.isOrdered ? 'TRUE' : 'FALSE', nowJST(),
    ];
  });

  var lastRow = masterSheet.getLastRow();
  if (lastRow > 1) masterSheet.deleteRows(2, lastRow - 1);
  if (masterRows.length > 0) {
    masterSheet.getRange(2, 1, masterRows.length, masterRows[0].length).setValues(masterRows);
  }
  Logger.log('[UNIT_PRICE] 単価マスタ再構築完了: ' + masterRows.length + '品目');
}

// ============================================================
// 外部公開API（12_price_compare_v2.gs / 07 chatbot api.gs から呼び出し）
// ============================================================

/**
 * 品名・仕様から過去単価を検索（スコア降順で最大10件）
 */
function searchUnitPrice(itemName, spec, client) {
  var ss          = getSpreadsheet();
  var masterSheet = ss.getSheetByName(SHEET_UNIT_PRICE);
  if (!masterSheet || masterSheet.getLastRow() <= 1) return [];

  var ncols = Object.keys(UNIT_PRICE_COLS).length;
  var data  = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, ncols).getValues();
  var query = _normItemName(itemName);
  var results = [];

  data.forEach(function(row) {
    var masterItem   = String(row[UNIT_PRICE_COLS.ITEM_NAME   - 1] || '');
    var masterSpec   = String(row[UNIT_PRICE_COLS.SPEC        - 1] || '');
    var masterClient = String(row[UNIT_PRICE_COLS.CLIENT      - 1] || '');

    var score = 0;
    if (masterItem === query)                score += 100;
    else if (masterItem.indexOf(query) >= 0) score +=  60;
    else if (query.indexOf(masterItem) >= 0) score +=  40;
    if (score === 0) return;

    if (spec   && masterSpec   && masterSpec.indexOf(spec)     >= 0) score += 20;
    if (client && masterClient && masterClient.indexOf(client) >= 0) score += 30;

    results.push({
      score     : score,
      itemName  : String(row[UNIT_PRICE_COLS.ITEM_NAME_RAW - 1] || ''),
      spec      : masterSpec,
      client    : masterClient,
      unitPrice : Number(row[UNIT_PRICE_COLS.UNIT_PRICE - 1] || 0),
      unit      : String(row[UNIT_PRICE_COLS.UNIT       - 1] || ''),
      quoteNo   : String(row[UNIT_PRICE_COLS.QUOTE_NO   - 1] || ''),
      quoteDate : String(row[UNIT_PRICE_COLS.QUOTE_DATE  - 1] || ''),
      isOrdered : row[UNIT_PRICE_COLS.IS_ORDERED - 1] === 'TRUE',
    });
  });

  results.sort(function(a, b) { return b.score - a.score; });
  return results.slice(0, 10);
}

/**
 * 案件サマリからキーワード検索（チャットボット用）
 */
function searchCaseSummary(keywords) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CASE_SUMMARY);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var ncols = Object.keys(CASE_SUMMARY_COLS).length;
  var data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, ncols).getValues();
  var results = [];

  data.forEach(function(row) {
    var searchText = [
      row[CASE_SUMMARY_COLS.MGMT_ID    - 1],
      row[CASE_SUMMARY_COLS.QUOTE_NO   - 1],
      row[CASE_SUMMARY_COLS.ORDER_NO   - 1],
      row[CASE_SUMMARY_COLS.SUBJECT    - 1],
      row[CASE_SUMMARY_COLS.CLIENT     - 1],
      row[CASE_SUMMARY_COLS.MODEL_CODE - 1],
      row[CASE_SUMMARY_COLS.ITEM_NAMES - 1],
    ].join(' ').toLowerCase();

    var score = 0;
    keywords.forEach(function(kw) {
      if (searchText.indexOf(kw.toLowerCase()) >= 0) score++;
    });
    if (score === 0) return;

    var unitPriceData = [];
    try { unitPriceData = JSON.parse(row[CASE_SUMMARY_COLS.UNIT_PRICES - 1] || '[]'); } catch(e) {}

    results.push({
      score       : score,
      mgmtId      : String(row[CASE_SUMMARY_COLS.MGMT_ID      - 1] || ''),
      quoteNo     : String(row[CASE_SUMMARY_COLS.QUOTE_NO     - 1] || ''),
      orderNo     : String(row[CASE_SUMMARY_COLS.ORDER_NO     - 1] || ''),
      subject     : String(row[CASE_SUMMARY_COLS.SUBJECT      - 1] || ''),
      client      : String(row[CASE_SUMMARY_COLS.CLIENT       - 1] || ''),
      status      : String(row[CASE_SUMMARY_COLS.STATUS       - 1] || ''),
      quoteDate   : String(row[CASE_SUMMARY_COLS.QUOTE_DATE   - 1] || ''),
      orderDate   : String(row[CASE_SUMMARY_COLS.ORDER_DATE   - 1] || ''),
      quoteAmount : row[CASE_SUMMARY_COLS.QUOTE_AMOUNT - 1] || '',
      orderAmount : row[CASE_SUMMARY_COLS.ORDER_AMOUNT - 1] || '',
      quotePdf    : String(row[CASE_SUMMARY_COLS.QUOTE_PDF    - 1] || ''),
      orderPdf    : String(row[CASE_SUMMARY_COLS.ORDER_PDF    - 1] || ''),
      modelCode   : String(row[CASE_SUMMARY_COLS.MODEL_CODE   - 1] || ''),
      itemNames   : String(row[CASE_SUMMARY_COLS.ITEM_NAMES   - 1] || ''),
      unitPrices  : unitPriceData,
    });
  });

  results.sort(function(a, b) { return b.score - a.score; });
  return results.slice(0, 30);
}

/**
 * 注文明細DBの単価比較結果を更新（12_price_compare_v2.gs から呼び出し）
 */
function updateOrderDetailCompareResult(orderNo, lineNo, compareStatus, compareDetail) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ORDER_DETAIL);
  if (!sheet || sheet.getLastRow() <= 1) return false;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1,
                             ORDER_DETAIL_COLS.LINE_NO).getValues();
  for (var i = 0; i < data.length; i++) {
    var rowOrderNo = String(data[i][ORDER_DETAIL_COLS.ORDER_NO - 1] || '');
    var rowLineNo  = data[i][ORDER_DETAIL_COLS.LINE_NO - 1];
    if (rowOrderNo === String(orderNo) && String(rowLineNo) === String(lineNo)) {
      var rowIdx = i + 2;
      sheet.getRange(rowIdx, ORDER_DETAIL_COLS.COMPARE_STATUS).setValue(compareStatus);
      sheet.getRange(rowIdx, ORDER_DETAIL_COLS.COMPARE_DETAIL).setValue(compareDetail);
      sheet.getRange(rowIdx, ORDER_DETAIL_COLS.SYNCED_AT).setValue(nowJST());
      return true;
    }
  }
  return false;
}

// ============================================================
// 内部ユーティリティ
// ============================================================
function _buildMgmtMap(mgmtSheet) {
  var map = {};
  if (!mgmtSheet || mgmtSheet.getLastRow() <= 1) return map;
  // getAllMgmtData() は最低27列読み込むので直接利用
  var data = getAllMgmtData();
  data.forEach(function(row) {
    var id = String(row[MGMT_COLS.ID - 1] || '').trim();
    if (!id) return;
    map[id] = {
      status   : String(row[MGMT_COLS.STATUS    - 1] || ''),
      client   : String(row[MGMT_COLS.CLIENT    - 1] || ''),
      orderNo  : String(row[MGMT_COLS.ORDER_NO  - 1] || ''),
      orderDate: _toDateStr(row[MGMT_COLS.ORDER_DATE  - 1]),
      quoteDate: _toDateStr(row[MGMT_COLS.QUOTE_DATE  - 1]),
    };
  });
  return map;
}

function _getNormExistingIds(sheet) {
  var map = {};
  if (!sheet || sheet.getLastRow() <= 1) return map;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
       .forEach(function(r) { if (r[0]) map[String(r[0])] = true; });
  return map;
}

/**
 * 品名の正規化（全角→半角、スペース除去、小文字化）
 */
function _normItemName(name) {
  if (!name) return '';
  return String(name)
    .trim()
    .replace(/\s+/g, '')
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(c) {
      return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
    })
    .toLowerCase();
}

function _registerNormalizedTriggers() {
  var existing = ScriptApp.getProjectTriggers().map(function(t) {
    return t.getHandlerFunction();
  });
  if (existing.indexOf('syncAllDetailDB') < 0) {
    ScriptApp.newTrigger('syncAllDetailDB').timeBased().everyHours(1).create();
    Logger.log('[NORM] syncAllDetailDB トリガー登録');
  }
  if (existing.indexOf('rebuildUnitPriceMaster') < 0) {
    ScriptApp.newTrigger('rebuildUnitPriceMaster').timeBased()
             .atHour(3).everyDays(1).create();
    Logger.log('[NORM] rebuildUnitPriceMaster トリガー登録（毎日3時）');
  }
}

// ============================================================
// デバッグ用
// ============================================================
function debugNormalize() {
  Logger.log('=== 正規化DB状態確認 ===');
  var ss = getSpreadsheet();
  [SHEET_QUOTE_DETAIL, SHEET_ORDER_DETAIL, SHEET_UNIT_PRICE, SHEET_CASE_SUMMARY]
    .forEach(function(name) {
      var s = ss.getSheetByName(name);
      Logger.log(name + ': ' + (s ? (s.getLastRow() - 1) + '行' : '未作成'));
    });
}
