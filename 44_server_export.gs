// ============================================================
// 44_server_export.gs
// 見積書PDF → Googleドライブ「見積管理_サーバー保存」→（PC側）会社サーバー（M:）の所定フォルダへ保存
//
//   GAS（週1回・金曜15時）：前回以降に登録された見積書PDFを、マイドライブ直下の「見積管理_サーバー保存」へ
//       「見積番号_取引先_品名_作成日.pdf」の名前でコピーし、振り分け用の情報（件名・基板名・機種コード等）を
//       _manifest_*.json に書く。
//   PC側（tools/server_sync/drive_to_server_export.ps1・金曜17時）：Googleドライブ パソコン版の同フォルダから
//       件名・取引先・基板名で振り分け先を決めて会社サーバーへコピーし、済んだものは「送信済み」へ移す。
//
//   ・2026/10/20 以降に追加された見積書だけが対象（過去の分は保存しない）
//   ・一度書き出した見積は「サーバー保存履歴」シートに記録し、二度書き出さない
//   ・トリガーは processDriveImports（5分ごと）から自動で作成される（手動設定は不要）
// ============================================================

var SERVER_EXPORT_ROOT_NAME = '見積管理_サーバー保存';
var SERVER_EXPORT_HIST_SHEET = 'サーバー保存履歴';
var SERVER_EXPORT_BASELINE_PROP = 'SERVER_EXPORT_START_OVERRIDE';   // 変更したい時だけプロパティに ISO 日時を設定
var SERVER_EXPORT_START = '2026-10-20T00:00:00+09:00';          // この日以降に追加された見積書から保存する
var SERVER_EXPORT_TRIGGER_CHECK_PROP = 'SERVER_EXPORT_TRIGGER_CHECKED';

/** processDriveImports から呼ぶ：週1回のトリガーが無ければ作る（1日1回だけ確認） */
function ensureServerExportTrigger() {
  try {
    var props = PropertiesService.getScriptProperties();
    var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
    if (props.getProperty(SERVER_EXPORT_TRIGGER_CHECK_PROP) === today) return;
    props.setProperty(SERVER_EXPORT_TRIGGER_CHECK_PROP, today);
    var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'exportQuotesForServer'; });
    if (!has) {
      ScriptApp.newTrigger('exportQuotesForServer').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(15).create();
      Logger.log('[SERVER EXPORT] 週1回（金曜15時）のトリガーを作成しました');
    }
  } catch (e) { Logger.log('[ensureServerExportTrigger] ' + e.message); }
}

/** マイドライブ直下の「見積管理_サーバー保存」（無ければ作成） */
function _serverExportFolder() {
  var it = DriveApp.getRootFolder().getFoldersByName(SERVER_EXPORT_ROOT_NAME);
  while (it.hasNext()) { var f = it.next(); if (!f.isTrashed()) return f; }
  return DriveApp.getRootFolder().createFolder(SERVER_EXPORT_ROOT_NAME);
}

function _serverExportHistSheet() {
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName(SERVER_EXPORT_HIST_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SERVER_EXPORT_HIST_SHEET);
    var h = ['管理ID', '見積番号', 'ファイル名', '書き出し日時'];
    sh.getRange(1, 1, 1, h.length).setValues([h]).setBackground('#E3F2FD').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function _seClean(s) {
  return String(s || '').replace(/[\\\/:*?"<>|\r\n\t]/g, '_').replace(/\s+/g, ' ').trim();
}
/** 取引先名を短く（株式会社・様などを除く） */
function _seClient(s) {
  return _seClean(s).replace(/株式会社|有限会社|（株）|\(株\)|御中|様/g, '').replace(/[\s　]+/g, '') || '取引先不明';
}
function _seDate(v) {
  var d = (v instanceof Date) ? v : (v ? new Date(v) : null);
  if (!d || isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyyMMdd');
}

/**
 * 週1回（金曜15時）：新しい見積書PDFを「見積管理_サーバー保存」へ書き出す
 * 手動で今すぐ実行してもよい。
 */
function exportQuotesForServer() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, error: '他の処理が実行中です' };
  try {
    var props = PropertiesService.getScriptProperties();
    var base = new Date(props.getProperty(SERVER_EXPORT_BASELINE_PROP) || SERVER_EXPORT_START);

    var hist = _serverExportHistSheet();
    var done = {};
    if (hist.getLastRow() > 1) {
      hist.getRange(2, 1, hist.getLastRow() - 1, 1).getValues().forEach(function (r) { done[String(r[0])] = true; });
    }

    var folder = null, entries = [], histRows = [], usedNames = {};
    getAllMgmtData().forEach(function (r) {
      var id = String(r[MGMT_COLS.ID - 1] || '').trim();
      var url = String(r[MGMT_COLS.QUOTE_PDF_URL - 1] || '').trim();
      var quoteNo = String(r[MGMT_COLS.QUOTE_NO - 1] || '').trim();
      if (!id || !url || !quoteNo || done[id]) return;
      var created = r[MGMT_COLS.CREATED_AT - 1];
      var cd = (created instanceof Date) ? created : (created ? new Date(created) : null);
      if (!cd || isNaN(cd.getTime()) || cd < base) return; // 運用開始前の見積は対象外

      var fid = _extractFileIdFromUrl(url);
      if (!fid) return;
      var file;
      try { file = DriveApp.getFileById(fid); } catch (e) { Logger.log('[SERVER EXPORT] PDFが開けません ' + quoteNo + ': ' + e.message); return; }

      var subject = _seClean(r[MGMT_COLS.SUBJECT - 1]);
      var board = _seClean(r[MGMT_COLS.BOARD_NAME - 1]);
      var client = _seClean(r[MGMT_COLS.CLIENT - 1]);
      var date = _seDate(r[MGMT_COLS.QUOTE_DATE - 1]) || _seDate(cd);
      var name = [_seClean(quoteNo), _seClient(client), subject || board || '品名なし', date].join('_');
      if (name.length > 150) name = name.substring(0, 150);
      var fileName = name + '.pdf', n = 2;
      while (usedNames[fileName]) { fileName = name + '_' + (n++) + '.pdf'; }
      usedNames[fileName] = true;

      if (!folder) folder = _serverExportFolder();
      folder.createFile(file.getBlob().setName(fileName));
      entries.push({
        file: fileName, mgmtId: id, quoteNo: quoteNo, client: client, subject: subject, board: board,
        model: _seClean(r[MGMT_COLS.MODEL_CODE - 1]), orderType: _seClean(r[MGMT_COLS.ORDER_TYPE - 1]), date: date,
      });
      histRows.push([id, quoteNo, fileName, new Date()]);
    });

    if (entries.length) {
      var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
      folder.createFile('_manifest_' + stamp + '.json', JSON.stringify({ createdAt: new Date().toISOString(), entries: entries }, null, 1), 'application/json');
      hist.getRange(hist.getLastRow() + 1, 1, histRows.length, 4).setValues(histRows);
    }
    Logger.log('[SERVER EXPORT] 見積書 ' + entries.length + '件を「' + SERVER_EXPORT_ROOT_NAME + '」へ書き出しました');
    return { success: true, count: entries.length };
  } catch (e) {
    Logger.log('[exportQuotesForServer] ' + e.message);
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}
