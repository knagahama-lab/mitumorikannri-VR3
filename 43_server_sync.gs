// ============================================================
// 43_server_sync.gs
// 会社サーバー（M:）→ Googleドライブ「見積管理_自動取込」フォルダ → 自動OCR取込
//
//   PC側（tools/server_sync/server_to_drive_sync.ps1）が、会社サーバーに保存された注文書PDFを
//   Googleドライブ パソコン版の「マイドライブ\見積管理_自動取込\注文書(試作)／注文書(量産)／見積書」へコピーする。
//   GAS側は 04 drive watcher の processDriveImports（5分ごと）から本ファイルの監視を呼び、
//   フォルダを「名前」で探して取り込む（フォルダIDの設定は不要）。
//   ファイル名の先頭の [取引先名] は取引先判定のヒントとして使う。
// ============================================================

var SERVER_SYNC_ROOT_NAME = '見積管理_自動取込';
var SERVER_SYNC_SUBFOLDERS = { trial: '注文書(試作)', mass: '注文書(量産)', quote: '見積書' };
var SERVER_SYNC_PROP = 'SERVER_SYNC_FOLDER_IDS';

/** 「見積管理_自動取込」と配下3フォルダのID（見つからなければ作成）。IDはプロパティにキャッシュ */
function _serverSyncFolders() {
  var props = PropertiesService.getScriptProperties();
  var cached = {};
  try { cached = JSON.parse(props.getProperty(SERVER_SYNC_PROP) || '{}'); } catch (e) {}
  var ok = function (id) { try { return id && !DriveApp.getFolderById(id).isTrashed(); } catch (e) { return false; } };
  if (cached.root && ok(cached.root) && ok(cached.trial) && ok(cached.mass) && ok(cached.quote)) return cached;

  var it = DriveApp.getFoldersByName(SERVER_SYNC_ROOT_NAME), root = null;
  while (it.hasNext()) { var f = it.next(); if (!f.isTrashed()) { root = f; break; } }
  if (!root) return null; // PC側で作られるまでは何もしない
  var ids = { root: root.getId() };
  Object.keys(SERVER_SYNC_SUBFOLDERS).forEach(function (k) {
    var sub = root.getFoldersByName(SERVER_SYNC_SUBFOLDERS[k]);
    ids[k] = (sub.hasNext() ? sub.next() : root.createFolder(SERVER_SYNC_SUBFOLDERS[k])).getId();
  });
  props.setProperty(SERVER_SYNC_PROP, JSON.stringify(ids));
  return ids;
}

/** processDriveImports から呼ぶ：サーバー連携フォルダの取込（処理済みは各フォルダ内の「処理済み」へ移動） */
function processServerSyncImports() {
  var ids = _serverSyncFolders();
  if (!ids) return 0;
  var n = 0;
  n += _watchFolder(ids.trial, CONFIG.ORDER_TRIAL_FOLDER_ID, 'order', CONFIG.ORDER_TYPE.TRIAL, '');
  n += _watchFolder(ids.mass,  CONFIG.ORDER_MASS_FOLDER_ID,  'order', CONFIG.ORDER_TYPE.MASS,  '');
  n += _watchFolder(ids.quote, CONFIG.QUOTE_FOLDER_ID,       'quote', '',                      '');
  if (n) Logger.log('[SERVER SYNC] サーバー連携フォルダから ' + n + '件を取り込みました');
  return n;
}

/** ファイル名先頭の [取引先名] を取り出す（例: "[藤商事]発注書206424.pdf" → "藤商事"） */
function _serverSyncClientFromName(name) {
  var m = String(name || '').match(/^\[([^\]]+)\]/);
  return m ? m[1] : '';
}
