// ============================================================
// 26_model_extensions.gs
// 機種マスタ 拡張データ管理
//
// ① 機種付属ファイル（仕様書PDF・指示書・納入指示書・構成表）
// ② 適合状況・12Dコード一覧（見本機 / 量産）
// ③ 売上数管理
// ============================================================

// ── シート名 ──
var MEXT_SHEET_FILES      = '機種ファイル管理';
var MEXT_SHEET_COMPLIANCE = '適合状況管理';
var MEXT_SHEET_SALES      = '機種売上管理';

// ── ファイル種別 ──
var MEXT_FILE_TYPES = ['仕様書PDF', '指示書', '納入指示書', '構成表', 'その他'];

// ============================================================
// ① 機種付属ファイル管理
//    ヘッダー: RowID | 機種コード | ファイル種別 | ファイル名 | Drive URL | 備考 | 登録日時
// ============================================================

function _initModelFilesSheet() {
  var ss      = getSpreadsheet();
  var sheet   = ss.getSheetByName(MEXT_SHEET_FILES);
  if (!sheet) {
    sheet = ss.insertSheet(MEXT_SHEET_FILES);
    var h = ['RowID', '機種コード', 'ファイル種別', 'ファイル名', 'Drive URL', '備考', '登録日時'];
    var hr = sheet.getRange(1, 1, 1, h.length);
    hr.setValues([h]);
    hr.setBackground('#E8F5E9');
    hr.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(5, 240);
  }
  return sheet;
}

function apiModelFilesList(payload) {
  try {
    var modelCode = String((payload || {}).modelCode || '').trim();
    if (!modelCode) return { success: false, error: '機種コードが必要です' };
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(MEXT_SHEET_FILES);
    if (!sheet) return { success: true, files: [] };
    var last  = sheet.getLastRow();
    if (last <= 1) return { success: true, files: [] };
    var rows  = sheet.getRange(2, 1, last - 1, 7).getValues();
    var files = rows
      .filter(function(r) { return String(r[1]).trim() === modelCode; })
      .map(function(r) {
        return {
          id:       String(r[0] || ''),
          fileType: String(r[2] || ''),
          fileName: String(r[3] || ''),
          url:      String(r[4] || ''),
          memo:     String(r[5] || ''),
          createdAt:_toDateStr(r[6]),
        };
      });
    return { success: true, files: files };
  } catch(e) { return { success: false, error: e.message }; }
}

function apiModelFileSave(payload) {
  try {
    payload = payload || {};
    var modelCode = String(payload.modelCode || '').trim();
    if (!modelCode) return { success: false, error: '機種コードが必要です' };
    var sheet = _initModelFilesSheet();
    var now   = nowJST();
    var isNew = !payload.id || String(payload.id).trim() === '';
    var id    = isNew ? ('MF-' + Utilities.formatDate(new Date(),'Asia/Tokyo','yyyyMMddHHmmss') + Math.floor(Math.random()*1000)) : String(payload.id).trim();
    var row   = [id, modelCode, payload.fileType||'その他', payload.fileName||'', payload.url||'', payload.memo||'', now];
    if (isNew) {
      sheet.appendRow(row);
    } else {
      var last = sheet.getLastRow();
      if (last > 1) {
        var ids = sheet.getRange(2,1,last-1,1).getValues().flat().map(function(v){return String(v);});
        var idx = ids.indexOf(id);
        if (idx >= 0) { sheet.getRange(idx+2,1,1,7).setValues([row]); }
        else           { sheet.appendRow(row); }
      } else { sheet.appendRow(row); }
    }
    return { success: true, id: id };
  } catch(e) { return { success: false, error: e.message }; }
}

function apiModelFileDelete(payload) {
  try {
    var id    = String((payload||{}).id||'').trim();
    if (!id) return { success: false, error: 'IDが必要です' };
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(MEXT_SHEET_FILES);
    if (!sheet) return { success: false, error: 'シートがありません' };
    var last  = sheet.getLastRow();
    if (last <= 1) return { success: false, error: 'データなし' };
    var ids   = sheet.getRange(2,1,last-1,1).getValues().flat().map(function(v){return String(v);});
    var idx   = ids.indexOf(id);
    if (idx < 0) return { success: false, error: 'IDが見つかりません' };
    sheet.deleteRow(idx + 2);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ============================================================
// ② 適合状況・12Dコード一覧
//    ヘッダー: RowID | 機種コード | コード番号 | 基板 | 基板種類
//             | 見本機_台数 | 見本機_初回実装 | 見本機_初回組立 | 見本機_出荷日 | 見本機_備考 | 見本機_ロム種類
//             | 量産_台数 | 量産_初回実装 | 量産_初回組立 | 量産_出荷日 | 量産_備考 | 量産_ロム種類
// ============================================================

var COMP_HEADERS = [
  'RowID', '機種コード', 'コード番号', '基板', '基板種類',
  '見本機_台数', '見本機_初回実装', '見本機_初回組立', '見本機_出荷日', '見本機_備考', '見本機_ロム種類',
  '量産_台数', '量産_初回実装', '量産_初回組立', '量産_出荷日', '量産_備考', '量産_ロム種類'
];

function _initComplianceSheet() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(MEXT_SHEET_COMPLIANCE);
  if (!sheet) {
    sheet = ss.insertSheet(MEXT_SHEET_COMPLIANCE);
    var hr = sheet.getRange(1, 1, 1, COMP_HEADERS.length);
    hr.setValues([COMP_HEADERS]);
    hr.setBackground('#FFF3E0');
    hr.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, COMP_HEADERS.length);
  }
  return sheet;
}

function _compRowToObj(r) {
  return {
    id:         String(r[0]  || ''),
    modelCode:  String(r[1]  || ''),
    codeNo:     String(r[2]  || ''),
    board:      String(r[3]  || ''),
    boardType:  String(r[4]  || ''),
    sample: {
      count:       r[5]  !== '' ? String(r[5])  : '',
      firstMount:  _toDateStr(r[6]),
      firstAssem:  _toDateStr(r[7]),
      shipDate:    _toDateStr(r[8]),
      memo:        String(r[9]  || ''),
      romType:     String(r[10] || ''),
    },
    mass: {
      count:       r[11] !== '' ? String(r[11]) : '',
      firstMount:  _toDateStr(r[12]),
      firstAssem:  _toDateStr(r[13]),
      shipDate:    _toDateStr(r[14]),
      memo:        String(r[15] || ''),
      romType:     String(r[16] || ''),
    },
  };
}

function apiComplianceList(payload) {
  try {
    var modelCode = String((payload||{}).modelCode||'').trim();
    if (!modelCode) return { success: false, error: '機種コードが必要です' };
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(MEXT_SHEET_COMPLIANCE);
    if (!sheet) return { success: true, rows: [] };
    var last  = sheet.getLastRow();
    if (last <= 1) return { success: true, rows: [] };
    var data  = sheet.getRange(2, 1, last - 1, COMP_HEADERS.length).getValues();
    var rows  = data
      .filter(function(r) { return String(r[1]).trim() === modelCode; })
      .map(function(r)    { return _compRowToObj(r); });
    return { success: true, rows: rows };
  } catch(e) { return { success: false, error: e.message }; }
}

function apiComplianceSave(payload) {
  try {
    payload = payload || {};
    var modelCode = String(payload.modelCode||'').trim();
    if (!modelCode) return { success: false, error: '機種コードが必要です' };
    var sheet = _initComplianceSheet();
    var isNew = !payload.id || String(payload.id).trim() === '';
    var id    = isNew ? ('CP-' + Utilities.formatDate(new Date(),'Asia/Tokyo','yyyyMMddHHmmss') + Math.floor(Math.random()*1000)) : String(payload.id).trim();
    var s = payload.sample || {};
    var m = payload.mass   || {};
    var row = [
      id, modelCode,
      payload.codeNo    || '',
      payload.board     || '',
      payload.boardType || '',
      s.count      || '', s.firstMount || '', s.firstAssem || '', s.shipDate || '', s.memo || '', s.romType || '',
      m.count      || '', m.firstMount || '', m.firstAssem || '', m.shipDate || '', m.memo || '', m.romType || '',
    ];
    if (isNew) {
      sheet.appendRow(row);
    } else {
      var last = sheet.getLastRow();
      if (last > 1) {
        var ids = sheet.getRange(2,1,last-1,1).getValues().flat().map(function(v){return String(v);});
        var idx = ids.indexOf(id);
        if (idx >= 0) { sheet.getRange(idx+2,1,1,COMP_HEADERS.length).setValues([row]); }
        else           { sheet.appendRow(row); }
      } else { sheet.appendRow(row); }
    }
    return { success: true, id: id };
  } catch(e) { return { success: false, error: e.message }; }
}

function apiComplianceDelete(payload) {
  try {
    var id    = String((payload||{}).id||'').trim();
    if (!id) return { success: false, error: 'IDが必要です' };
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(MEXT_SHEET_COMPLIANCE);
    if (!sheet) return { success: false, error: 'シートなし' };
    var last  = sheet.getLastRow();
    if (last <= 1) return { success: false, error: 'データなし' };
    var ids   = sheet.getRange(2,1,last-1,1).getValues().flat().map(function(v){return String(v);});
    var idx   = ids.indexOf(id);
    if (idx < 0) return { success: false, error: 'IDが見つかりません' };
    sheet.deleteRow(idx + 2);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ============================================================
// ③ 売上数管理
//    ヘッダー: RowID | 機種コード | 年月 | 台数 | 金額 | 再販区分 | 備考 | 登録日時
// ============================================================

var SALES_HEADERS = ['RowID', '機種コード', '年月', '台数', '金額', '再販区分', '備考', '登録日時'];

function _initModelSalesSheet() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(MEXT_SHEET_SALES);
  if (!sheet) {
    sheet = ss.insertSheet(MEXT_SHEET_SALES);
    var hr = sheet.getRange(1, 1, 1, SALES_HEADERS.length);
    hr.setValues([SALES_HEADERS]);
    hr.setBackground('#E3F2FD');
    hr.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, SALES_HEADERS.length);
  }
  return sheet;
}

function apiModelSalesList(payload) {
  try {
    var modelCode = String((payload||{}).modelCode||'').trim();
    if (!modelCode) return { success: false, error: '機種コードが必要です' };
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(MEXT_SHEET_SALES);
    if (!sheet) return { success: true, sales: [] };
    var last  = sheet.getLastRow();
    if (last <= 1) return { success: true, sales: [] };
    var data  = sheet.getRange(2, 1, last - 1, SALES_HEADERS.length).getValues();
    var sales = data
      .filter(function(r) { return String(r[1]).trim() === modelCode; })
      .map(function(r) {
        return {
          id:        String(r[0]||''),
          yearMonth: String(r[2]||''),
          count:     r[3] !== '' ? Number(r[3]) : null,
          amount:    r[4] !== '' ? Number(r[4]) : null,
          resale:    String(r[5]||''),
          memo:      String(r[6]||''),
          createdAt: _toDateStr(r[7]),
        };
      });
    // 年月降順
    sales.sort(function(a,b){ return (b.yearMonth||'').localeCompare(a.yearMonth||''); });
    return { success: true, sales: sales };
  } catch(e) { return { success: false, error: e.message }; }
}

function apiModelSalesSave(payload) {
  try {
    payload = payload || {};
    var modelCode = String(payload.modelCode||'').trim();
    if (!modelCode) return { success: false, error: '機種コードが必要です' };
    var sheet = _initModelSalesSheet();
    var now   = nowJST();
    var isNew = !payload.id || String(payload.id).trim() === '';
    var id    = isNew ? ('SL-' + Utilities.formatDate(new Date(),'Asia/Tokyo','yyyyMMddHHmmss') + Math.floor(Math.random()*1000)) : String(payload.id).trim();
    var row   = [
      id, modelCode,
      payload.yearMonth || '',
      payload.count  !== undefined && payload.count  !== null ? Number(payload.count)  : '',
      payload.amount !== undefined && payload.amount !== null ? Number(payload.amount) : '',
      payload.resale || '',
      payload.memo   || '',
      now,
    ];
    if (isNew) {
      sheet.appendRow(row);
    } else {
      var last = sheet.getLastRow();
      if (last > 1) {
        var ids = sheet.getRange(2,1,last-1,1).getValues().flat().map(function(v){return String(v);});
        var idx = ids.indexOf(id);
        if (idx >= 0) { sheet.getRange(idx+2,1,1,SALES_HEADERS.length).setValues([row]); }
        else           { sheet.appendRow(row); }
      } else { sheet.appendRow(row); }
    }
    return { success: true, id: id };
  } catch(e) { return { success: false, error: e.message }; }
}

function apiModelSalesDelete(payload) {
  try {
    var id    = String((payload||{}).id||'').trim();
    if (!id) return { success: false, error: 'IDが必要です' };
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(MEXT_SHEET_SALES);
    if (!sheet) return { success: false, error: 'シートなし' };
    var last  = sheet.getLastRow();
    if (last <= 1) return { success: false, error: 'データなし' };
    var ids   = sheet.getRange(2,1,last-1,1).getValues().flat().map(function(v){return String(v);});
    var idx   = ids.indexOf(id);
    if (idx < 0) return { success: false, error: 'IDが見つかりません' };
    sheet.deleteRow(idx + 2);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}
