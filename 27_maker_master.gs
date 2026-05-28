// ============================================================
// 27_maker_master.gs
// 遊技機メーカー管理
// 競合情報・決算資料の格納、販売訪問管理
// ============================================================

var MAKER_SHEET   = 'メーカーマスタ';
var MAKER_HEADERS = [
  'メーカーID', 'メーカー名', 'ジャンル', '住所', '電話', 'メール', 'URL',
  '担当者', '最終訪問日', '次回訪問予定', '競合情報', '決算資料URL', '備考',
  '登録日時', '更新日時'
];
// ジャンル候補
var MAKER_GENRES = ['遊技機メーカー', '部品メーカー', '販売代理店', 'その他'];

// ============================================================
// シート初期化
// ============================================================

function initMakerMasterSheet() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(MAKER_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(MAKER_SHEET);
    var hr = sheet.getRange(1, 1, 1, MAKER_HEADERS.length);
    hr.setValues([MAKER_HEADERS]);
    hr.setBackground('#FCE4EC');
    hr.setFontWeight('bold');
    hr.setFontSize(10);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(11, 300); // 競合情報
    sheet.setColumnWidth(12, 240); // 決算資料URL
    sheet.autoResizeColumns(1, 10);
    Logger.log('メーカーマスタシート作成完了');
  }
  return sheet;
}

function _getMakerSheet() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(MAKER_SHEET);
  if (!sheet) { initMakerMasterSheet(); sheet = ss.getSheetByName(MAKER_SHEET); }
  return sheet;
}

function _generateMakerId() {
  return 'MK-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd') + '-' +
    (Math.floor(Math.random() * 9000) + 1000);
}

function _makerRowToObj(r) {
  return {
    id:            String(r[0]  || ''),
    name:          String(r[1]  || ''),
    genre:         String(r[2]  || ''),
    address:       String(r[3]  || ''),
    tel:           String(r[4]  || ''),
    email:         String(r[5]  || ''),
    url:           String(r[6]  || ''),
    assignee:      String(r[7]  || ''),
    lastVisit:     _toDateStr(r[8]),
    nextVisit:     _toDateStr(r[9]),
    competitorInfo:String(r[10] || ''),
    settlementUrl: String(r[11] || ''),
    memo:          String(r[12] || ''),
    createdAt:     _toDateStr(r[13]),
    updatedAt:     _toDateStr(r[14]),
  };
}

// ============================================================
// API: 一覧取得
// ============================================================

function apiMakerList() {
  try {
    var sheet = _getMakerSheet();
    var last  = sheet.getLastRow();
    if (last <= 1) return { success: true, items: [] };
    var rows  = sheet.getRange(2, 1, last - 1, MAKER_HEADERS.length).getValues();
    var items = rows
      .filter(function(r) { return String(r[0]).trim() !== ''; })
      .map(function(r)    { return _makerRowToObj(r); });
    // 次回訪問日 昇順ソート（未設定は後ろ）
    items.sort(function(a, b) {
      var na = a.nextVisit || 'zzzzz';
      var nb = b.nextVisit || 'zzzzz';
      return na.localeCompare(nb);
    });
    return { success: true, items: items };
  } catch(e) { return { success: false, error: e.message }; }
}

// ============================================================
// API: 保存（新規/更新）
// ============================================================

function apiMakerSave(payload) {
  try {
    payload = payload || {};
    var name = String(payload.name || '').trim();
    if (!name) return { success: false, error: 'メーカー名は必須です' };
    var sheet = _getMakerSheet();
    var now   = nowJST();
    var isNew = !payload.id || String(payload.id).trim() === '';
    var id    = isNew ? _generateMakerId() : String(payload.id).trim();

    var row = [
      id,
      name,
      payload.genre         || '',
      payload.address       || '',
      payload.tel           || '',
      payload.email         || '',
      payload.url           || '',
      payload.assignee      || '',
      payload.lastVisit     || '',
      payload.nextVisit     || '',
      payload.competitorInfo|| '',
      payload.settlementUrl || '',
      payload.memo          || '',
    ];

    if (isNew) {
      row.push(now); // 登録日時
      row.push(now); // 更新日時
      sheet.appendRow(row);
    } else {
      var last = sheet.getLastRow();
      if (last <= 1) return { success: false, error: 'データなし' };
      var ids  = sheet.getRange(2, 1, last - 1, 1).getValues().flat().map(function(v){ return String(v); });
      var idx  = ids.indexOf(id);
      if (idx < 0) return { success: false, error: 'IDが見つかりません: ' + id };
      var rowNum  = idx + 2;
      var created = sheet.getRange(rowNum, 14).getValue();
      row.push(created || now); // 登録日時
      row.push(now);            // 更新日時
      sheet.getRange(rowNum, 1, 1, MAKER_HEADERS.length).setValues([row]);
    }
    return { success: true, id: id };
  } catch(e) { return { success: false, error: e.message }; }
}

// ============================================================
// API: 削除
// ============================================================

function apiMakerDelete(payload) {
  try {
    var id    = String((payload||{}).id||'').trim();
    if (!id) return { success: false, error: 'IDが必要です' };
    var sheet = _getMakerSheet();
    var last  = sheet.getLastRow();
    if (last <= 1) return { success: false, error: 'データなし' };
    var ids   = sheet.getRange(2, 1, last - 1, 1).getValues().flat().map(function(v){ return String(v); });
    var idx   = ids.indexOf(id);
    if (idx < 0) return { success: false, error: 'IDが見つかりません: ' + id };
    sheet.deleteRow(idx + 2);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}
