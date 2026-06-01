// ============================================================
// 25_quoteset.gs
// 見積セット管理 API
// 4種の見積書（基板回路設計費・基板実装費・量産基板・注残処理）を
// セットとして管理し、各書類の進捗をトラッキングする
// ============================================================

var QS_TYPES = ['基板回路設計費', '基板実装費', '量産基板', '注残処理'];
var QS_STATUSES = ['未作成', '作成中', '承認待ち', '発行済', '却下'];

// シートヘッダー（26列）
var QS_HEADERS = [
  'セットID', '見積依頼日', '案件名', '顧客名', '担当者', '楽楽販売番号',
  '基板回路設計費_見積番号', '基板回路設計費_金額', '基板回路設計費_状態',
  '基板実装費_見積番号',     '基板実装費_金額',     '基板実装費_状態',
  '量産基板_見積番号',       '量産基板_金額',       '量産基板_状態',
  '注残処理_見積番号',       '注残処理_金額',       '注残処理_状態',
  '基板カテゴリ', '進捗', '備考', '追加見積JSON', '機種コード', '機種名', '登録日時', '更新日時'
];

// ============================================================
// シート初期化
// ============================================================

function initQuoteSetSheet() {
  var ss    = getSpreadsheet();
  var name  = CONFIG.SHEET_QUOTE_SET;
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    var hr = sheet.getRange(1, 1, 1, QS_HEADERS.length);
    hr.setValues([QS_HEADERS]);
    hr.setBackground('#E8F0FE');
    hr.setFontWeight('bold');
    hr.setFontSize(10);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, QS_HEADERS.length);
    Logger.log('見積セット管理シート作成完了');
  }
  return sheet;
}

// ============================================================
// ID生成
// ============================================================

function _generateQsId() {
  return 'QS-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd') + '-' +
    (Math.floor(Math.random() * 9000) + 1000);
}

// ============================================================
// 進捗計算
// ============================================================

function _calcQsProgress(statuses) {
  // statuses: 固定4種＋追加見積の状態配列
  if (statuses.length === 0) return '未着手';
  if (statuses.some(function(s) { return s === '却下'; })) return '要確認';
  if (statuses.every(function(s) { return s === '発行済'; })) return '完了';
  if (statuses.every(function(s) { return s === '未作成'; })) return '未着手';
  return '進行中';
}

// ============================================================
// 行 → オブジェクト変換
// ============================================================

function _qsRowToObj(row) {
  var types = QS_TYPES;
  var items = {};
  types.forEach(function(t, i) {
    var base = 6 + i * 3; // col index (0-based)
    items[t] = {
      quoteNo: String(row[base]     || ''),
      amount:  row[base + 1] !== '' && row[base + 1] !== null ? Number(row[base + 1]) : null,
      status:  String(row[base + 2] || '未作成'),
    };
  });

  // 合計金額
  var total = 0;
  types.forEach(function(t) {
    if (items[t].amount !== null) total += items[t].amount;
  });

  // 追加見積 JSON パース
  var extraItems = [];
  try {
    var rawExtra = String(row[21] || '');
    if (rawExtra) extraItems = JSON.parse(rawExtra);
    if (!Array.isArray(extraItems)) extraItems = [];
  } catch(e2) { extraItems = []; }

  // 合計金額に追加見積を加算
  extraItems.forEach(function(ei) {
    if (ei.amount !== null && ei.amount !== undefined && ei.amount !== '') total += Number(ei.amount) || 0;
  });

  // 進捗計算（固定4種＋追加見積の状態を合算）
  var statuses = types.map(function(t) { return items[t].status; });
  extraItems.forEach(function(ei) { statuses.push(String(ei.status || '未作成')); });
  var progress = _calcQsProgress(statuses);

  return {
    id:             String(row[0]  || ''),
    requestDate:    _toDateStr(row[1]),
    subject:        String(row[2]  || ''),
    client:         String(row[3]  || ''),
    assignee:       String(row[4]  || ''),
    rakurakuNo:     String(row[5]  || ''),
    items:          items,
    extraItems:     extraItems,
    totalAmount:    total,
    boardCategory:  String(row[18] || ''),
    progress:       String(row[19] || progress),
    memo:           String(row[20] || ''),
    modelCode:      String(row[22] || ''),
    modelName:      String(row[23] || ''),
    createdAt:      _toDateStr(row[24]),
    updatedAt:      _toDateStr(row[25]),
  };
}

// ============================================================
// API: 一覧取得
// ============================================================

function apiQuoteSetList() {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_QUOTE_SET);
    if (!sheet) {
      // シートがなければ初期化して空を返す
      initQuoteSetSheet();
      return { success: true, items: [] };
    }
    var last = sheet.getLastRow();
    if (last <= 1) return { success: true, items: [] };

    var rows  = sheet.getRange(2, 1, last - 1, QS_HEADERS.length).getValues();
    var items = rows
      .filter(function(r) { return String(r[0]).trim() !== ''; })
      .map(function(r)    { return _qsRowToObj(r); });

    return { success: true, items: items };
  } catch (e) {
    Logger.log('[apiQuoteSetList ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// API: 保存（新規 / 更新）
// ============================================================

function apiQuoteSetSave(payload) {
  try {
    payload = payload || {};
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_QUOTE_SET) || initQuoteSetSheet();
    var now   = nowJST();

    var isNew = !payload.id || String(payload.id).trim() === '';
    var id    = isNew ? _generateQsId() : String(payload.id).trim();

    // items オブジェクトを展開
    var items  = payload.items || {};
    var types  = QS_TYPES;

    // 進捗計算（固定4種＋追加見積）
    var statuses = types.map(function(t) {
      return String((items[t] && items[t].status) || '未作成');
    });
    var extraItems = [];
    try {
      var ei = payload.extraItems;
      if (Array.isArray(ei)) extraItems = ei;
    } catch(e2) {}
    extraItems.forEach(function(ei) { statuses.push(String(ei.status || '未作成')); });
    var progress = _calcQsProgress(statuses);

    var rowData = [
      id,
      payload.requestDate || '',
      payload.subject     || '',
      payload.client      || '',
      payload.assignee    || '',
      payload.rakurakuNo  || '',
    ];

    types.forEach(function(t) {
      var item = items[t] || {};
      rowData.push(item.quoteNo  || '');
      rowData.push(item.amount   !== undefined && item.amount !== null ? Number(item.amount) : '');
      rowData.push(item.status   || '未作成');
    });

    rowData.push(payload.boardCategory || '');       // 基板カテゴリ [18]
    rowData.push(progress);                          // 進捗 [19]
    rowData.push(payload.memo || '');                // 備考 [20]
    rowData.push(JSON.stringify(extraItems));         // 追加見積JSON [21]
    rowData.push(payload.modelCode || '');           // 機種コード [22]
    rowData.push(payload.modelName || '');           // 機種名 [23]

    if (isNew) {
      rowData.push(now); // 登録日時 [24]
      rowData.push(now); // 更新日時 [25]
      sheet.appendRow(rowData);
    } else {
      // 既存行を探して更新
      var last = sheet.getLastRow();
      if (last <= 1) return { success: false, error: 'データが見つかりません: ' + id };
      var ids  = sheet.getRange(2, 1, last - 1, 1).getValues().flat().map(function(v){ return String(v); });
      var idx  = ids.indexOf(id);
      if (idx < 0) return { success: false, error: 'IDが見つかりません: ' + id };
      var rowNum = idx + 2;

      // 登録日時は既存値を保持（列25 = 1-based）
      var existCreated = sheet.getRange(rowNum, 25).getValue();
      rowData.push(existCreated || now); // 登録日時
      rowData.push(now);                 // 更新日時
      sheet.getRange(rowNum, 1, 1, QS_HEADERS.length).setValues([rowData]);
    }

    return { success: true, id: id, progress: progress };
  } catch (e) {
    Logger.log('[apiQuoteSetSave ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// API: 削除
// ============================================================

function apiQuoteSetDelete(payload) {
  try {
    payload = payload || {};
    var id = String(payload.id || '').trim();
    if (!id) return { success: false, error: 'IDが必要です' };

    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_QUOTE_SET);
    if (!sheet) return { success: false, error: 'シートが存在しません' };

    var last = sheet.getLastRow();
    if (last <= 1) return { success: false, error: 'データなし' };

    var ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat().map(function(v){ return String(v); });
    var idx = ids.indexOf(id);
    if (idx < 0) return { success: false, error: 'IDが見つかりません: ' + id };

    sheet.deleteRow(idx + 2);
    return { success: true, id: id };
  } catch (e) {
    Logger.log('[apiQuoteSetDelete ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}
