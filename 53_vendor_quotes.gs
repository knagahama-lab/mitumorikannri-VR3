// ============================================================
// 53_vendor_quotes.gs
// 依頼先見積書（実装会社・組立会社・部品メーカーなど、こちらが依頼して受け取る見積書）
//   ・PDFアップロードと件名は手入力（OCRなし）
//   ・保存先：見積管理システムのフォルダ（見積書の保存先フォルダの親）内の「依頼先見積書\<区分>」
//   ・シート：依頼先見積書
// ============================================================

var VQ_SHEET = '依頼先見積書';
var VQ_HEADERS = ['ID', '区分', '依頼先', '件名', '見積No', '見積日', '金額', '関連（機種・基板・弊社見積No）', 'PDF名', 'PDF URL', 'メモ', '登録者', '登録日時', '更新日時'];
var VQ_KINDS = ['実装会社', '組立会社', '部品メーカー'];
var VQ_FOLDER_NAME = '依頼先見積書';
var VQ_PROP_FOLDER = 'VQ_FOLDER_ID';

function _vqSheet() {
  var ss = getSpreadsheet(), sh = ss.getSheetByName(VQ_SHEET);
  if (!sh) {
    sh = ss.insertSheet(VQ_SHEET);
    sh.getRange(1, 1, 1, VQ_HEADERS.length).setValues([VQ_HEADERS]).setFontWeight('bold').setBackground('#E8F5E9');
    sh.setFrozenRows(1);
    sh.getRange(2, 6, sh.getMaxRows() - 1, 1).setNumberFormat('@');
    sh.getRange(2, 13, sh.getMaxRows() - 1, 2).setNumberFormat('@');
  }
  return sh;
}

/** 「依頼先見積書」フォルダ（見積管理システムのフォルダ＝見積書保存先の親フォルダの中） */
function _vqRootFolder() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(VQ_PROP_FOLDER);
  if (id) { try { var f0 = DriveApp.getFolderById(id); if (!f0.isTrashed()) return f0; } catch (e) {} }
  var quoteFolder = DriveApp.getFolderById(CONFIG.QUOTE_FOLDER_ID);
  var parents = quoteFolder.getParents();
  var base = parents.hasNext() ? parents.next() : quoteFolder;
  var it = base.getFoldersByName(VQ_FOLDER_NAME);
  var root = it.hasNext() ? it.next() : base.createFolder(VQ_FOLDER_NAME);
  props.setProperty(VQ_PROP_FOLDER, root.getId());
  return root;
}
function _vqKindFolder(kind) {
  var root = _vqRootFolder(), name = String(kind || 'その他').replace(/[\\\/:*?"<>|]/g, '_');
  var it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}
function _vqKinds() {
  try { var k = JSON.parse(PropertiesService.getScriptProperties().getProperty('VQ_KINDS') || 'null'); if (Array.isArray(k) && k.length) return k; } catch (e) {}
  return VQ_KINDS.slice();
}
function _vqNum(v) {
  if (v === '' || v == null) return '';
  var n = Number(String(v).normalize('NFKC').replace(/[,，¥￥円\s]/g, ''));
  if (isNaN(n)) throw new Error('金額は数値で入力してください');
  return Math.round(n * 100) / 100;
}
function _vqDate(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd');
  var m = String(v).normalize('NFKC').match(/(\d{4})[\/\-年.](\d{1,2})[\/\-月.](\d{1,2})/);
  return m ? m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2) : String(v);
}

function apiVqList() {
  try {
    var sh = _vqSheet(), last = sh.getLastRow();
    var rows = last > 1 ? sh.getRange(2, 1, last - 1, VQ_HEADERS.length).getValues() : [];
    var items = rows.filter(function (r) { return r[0]; }).map(function (r) {
      return { id: String(r[0]), kind: String(r[1] || ''), vendor: String(r[2] || ''), subject: String(r[3] || ''), quoteNo: String(r[4] || ''),
               date: _vqDate(r[5]), amount: r[6] === '' ? null : Number(r[6]), related: String(r[7] || ''), fileName: String(r[8] || ''),
               url: String(r[9] || ''), memo: String(r[10] || ''), user: String(r[11] || ''), createdAt: String(r[12] || ''), updatedAt: String(r[13] || '') };
    });
    items.sort(function (a, b) { return String(b.date || b.createdAt).localeCompare(String(a.date || a.createdAt)); });
    var folderUrl = ''; try { folderUrl = _vqRootFolder().getUrl(); } catch (e) {}
    return { success: true, items: items, kinds: _vqKinds(), folderUrl: folderUrl };
  } catch (e) { return { success: false, error: e.message }; }
}

/**
 * 登録・更新 p: { id?, kind, vendor, subject, quoteNo, date, amount, related, memo, file?: { name, mimeType, base64 } }
 *   file があれば Drive の「依頼先見積書\<区分>」へ保存して差し替え（以前のPDFは残す）
 */
function apiVqSave(p) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    p = p || {};
    var subject = String(p.subject || '').trim();
    if (!subject) return { success: false, error: '見積書の件名を入力してください' };
    if (!String(p.kind || '').trim()) return { success: false, error: '区分（実装会社・組立会社・部品メーカー）を選んでください' };
    var sh = _vqSheet(), now = nowJST();
    var user = ''; try { user = Session.getActiveUser().getEmail(); } catch (e) {}
    var rowNum = -1, cur = null;
    if (p.id) {
      var ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
      var i = ids.indexOf(String(p.id));
      if (i < 0) return { success: false, error: '対象の見積書が見つかりません' };
      rowNum = i + 2; cur = sh.getRange(rowNum, 1, 1, VQ_HEADERS.length).getValues()[0];
    }
    var fileName = cur ? cur[8] : '', url = cur ? cur[9] : '';
    if (p.file && p.file.base64) {
      var blob = Utilities.newBlob(Utilities.base64Decode(p.file.base64), p.file.mimeType || 'application/pdf', p.file.name || (subject + '.pdf'));
      var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
      var name = stamp + '_' + String(p.vendor || '').replace(/[\\\/:*?"<>|]/g, '_') + '_' + (p.file.name || subject + '.pdf');
      var f = _vqKindFolder(p.kind).createFile(blob.setName(name));
      fileName = p.file.name || name; url = f.getUrl();
    }
    var row = [cur ? cur[0] : 'VQ-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 900 + 100),
      String(p.kind).trim(), String(p.vendor || '').trim(), subject, String(p.quoteNo || '').trim(), _vqDate(p.date), _vqNum(p.amount),
      String(p.related || '').trim(), fileName, url, String(p.memo || ''), cur ? cur[11] : user, cur ? cur[12] : now, now];
    if (rowNum > 0) sh.getRange(rowNum, 1, 1, VQ_HEADERS.length).setValues([row]);
    else sh.getRange(sh.getLastRow() + 1, 1, 1, VQ_HEADERS.length).setValues([row]);
    return { success: true, id: row[0], url: url };
  } catch (e) { return { success: false, error: e.message }; }
  finally { try { lock.releaseLock(); } catch (e2) {} }
}

/** 削除（シートの行だけ。PDFはDriveに残す） */
function apiVqDelete(p) {
  try {
    var sh = _vqSheet();
    var ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
    var i = ids.indexOf(String((p || {}).id || ''));
    if (i < 0) return { success: false, error: '対象が見つかりません' };
    sh.deleteRow(i + 2);
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

/** 区分の追加・変更 p: { kinds: [...] } */
function apiVqKindsSave(p) {
  try {
    var k = ((p && p.kinds) || []).map(function (s) { return String(s).trim(); }).filter(Boolean);
    if (!k.length) return { success: false, error: '区分を1つ以上入力してください' };
    PropertiesService.getScriptProperties().setProperty('VQ_KINDS', JSON.stringify(k));
    return apiVqList();
  } catch (e) { return { success: false, error: e.message }; }
}
