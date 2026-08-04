// ============================================================
// 28_quote_documents.gs
// 見積関連書類管理
//
// 見積書PDFに加え、以下の関連書類を案件（mgmtId）単位で複数保存する:
//   ・社内回覧用資料（Excel）
//   ・お客様説明資料（Excel / PDF）
//   ・参考見積書_実装費原価（他社見積書）
//   ・参考見積書_組立費原価（他社見積書）
//   ・参考見積書_PCBメーカー
// 26_model_extensions.gs の「機種ファイル管理」と同じ「1レコードに複数ファイル」
// パターンを、機種コードではなく mgmtId（案件）をキーにして踏襲している。
// ============================================================

// ── シート名 ──
var QDOC_SHEET = '見積関連書類';

// ── 書類種別 ──
var QDOC_TYPES = [
  '社内回覧用資料',
  'お客様説明資料',
  '参考見積書_実装費原価',
  '参考見積書_組立費原価',
  '参考見積書_PCBメーカー',
];

// ============================================================
// 見積関連書類シート
//    ヘッダー: RowID | mgmtId | 書類種別 | ファイル名 | Drive URL | 備考 | 登録日時
// ============================================================

function _initQuoteDocsSheet() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(QDOC_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(QDOC_SHEET);
    var h  = ['RowID', 'mgmtId', '書類種別', 'ファイル名', 'Drive URL', '備考', '登録日時'];
    var hr = sheet.getRange(1, 1, 1, h.length);
    hr.setValues([h]);
    hr.setBackground('#E3F2FD');
    hr.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(5, 240);
  }
  return sheet;
}

// 一覧取得（mgmtId 指定）
function apiQuoteDocsList(payload) {
  try {
    var mgmtId = String((payload || {}).mgmtId || '').trim();
    if (!mgmtId) return { success: false, error: 'mgmtIdが必要です' };
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(QDOC_SHEET);
    if (!sheet) return { success: true, docTypes: QDOC_TYPES, files: [] };
    var last = sheet.getLastRow();
    if (last <= 1) return { success: true, docTypes: QDOC_TYPES, files: [] };
    var rows  = sheet.getRange(2, 1, last - 1, 7).getValues();
    var files = rows
      .filter(function(r) { return String(r[1]).trim() === mgmtId; })
      .map(function(r) {
        return {
          id:        String(r[0] || ''),
          mgmtId:    String(r[1] || ''),
          docType:   String(r[2] || ''),
          fileName:  String(r[3] || ''),
          url:       String(r[4] || ''),
          memo:      String(r[5] || ''),
          createdAt: _toDateStr(r[6]),
        };
      });
    return { success: true, docTypes: QDOC_TYPES, files: files };
  } catch (e) { return { success: false, error: e.message }; }
}

// メタデータの新規/更新保存（URLは既に確定している場合）
function apiQuoteDocSave(payload) {
  try {
    payload = payload || {};
    var mgmtId = String(payload.mgmtId || '').trim();
    if (!mgmtId) return { success: false, error: 'mgmtIdが必要です' };
    if (!payload.url) return { success: false, error: 'URLが必要です' };
    var sheet = _initQuoteDocsSheet();
    var now   = nowJST();
    var isNew = !payload.id || String(payload.id).trim() === '';
    var id    = isNew ? ('QD-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + Math.floor(Math.random() * 1000)) : String(payload.id).trim();
    var row   = [id, mgmtId, payload.docType || QDOC_TYPES[0], payload.fileName || '', payload.url || '', payload.memo || '', now];
    if (isNew) {
      sheet.appendRow(row);
    } else {
      var last = sheet.getLastRow();
      if (last > 1) {
        var ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat().map(function(v) { return String(v); });
        var idx = ids.indexOf(id);
        if (idx >= 0) { sheet.getRange(idx + 2, 1, 1, 7).setValues([row]); }
        else           { sheet.appendRow(row); }
      } else { sheet.appendRow(row); }
    }
    return { success: true, id: id };
  } catch (e) { return { success: false, error: e.message }; }
}

// ファイル本体アップロード（Base64） + シート登録を一括で行う
function apiQuoteDocUpload(payload) {
  try {
    payload = payload || {};
    var mgmtId = String(payload.mgmtId || '').trim();
    if (!mgmtId) return { success: false, error: 'mgmtIdが必要です' };
    if (!payload.base64Data || !payload.fileName) return { success: false, error: 'ファイルデータ不足' };
    if (QDOC_TYPES.indexOf(payload.docType) < 0) return { success: false, error: '不正な書類種別です' };

    var folder   = DriveApp.getFolderById(CONFIG.WEB_UPLOAD_FOLDER_ID);
    var mimeType = payload.mimeType || 'application/octet-stream';
    var safeName = String(payload.fileName).replace(/[/\\:*?"<>|]/g, '_');
    var blob     = Utilities.newBlob(Utilities.base64Decode(payload.base64Data), mimeType, mgmtId + '_' + payload.docType + '_' + safeName);
    var file     = folder.createFile(blob);
    var url      = file.getUrl();

    var saveRes = apiQuoteDocSave({
      mgmtId:   mgmtId,
      docType:  payload.docType,
      fileName: payload.fileName,
      url:      url,
      memo:     payload.memo || '',
    });
    if (!saveRes.success) return saveRes;
    return { success: true, id: saveRes.id, url: url, fileName: payload.fileName };
  } catch (e) { return { success: false, error: e.message }; }
}

// 削除
function apiQuoteDocDelete(payload) {
  try {
    var id = String((payload || {}).id || '').trim();
    if (!id) return { success: false, error: 'IDが必要です' };
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(QDOC_SHEET);
    if (!sheet) return { success: false, error: 'シートがありません' };
    var last = sheet.getLastRow();
    if (last <= 1) return { success: false, error: 'データなし' };
    var ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat().map(function(v) { return String(v); });
    var idx = ids.indexOf(id);
    if (idx < 0) return { success: false, error: 'IDが見つかりません' };
    sheet.deleteRow(idx + 2);
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}
