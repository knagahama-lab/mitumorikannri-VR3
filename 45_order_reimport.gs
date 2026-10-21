// ============================================================
// 45_order_reimport.gs
// 注文書の再インポート（上書き・重複させない）と、登録済みの重複注文の整理
//
//   ■ 上書きのしくみ（02_ocr_and_processing.gs の _saveOrderData から呼ぶ）
//     同じ取引先・同じ注文番号の注文が登録済みなら、新規行を作らずその行を上書きする。
//       ・進捗ステータス・見積との紐づけ・社内メモはそのまま
//       ・注文明細は古い行を消して書き直し、明細ステータス（注文書シート20列目以降）は行番号ごとに引き継ぐ
//       ・以前に取り込んだPDFのコピー（保存先フォルダの「日時_ファイル名.pdf」）はゴミ箱へ
//       ・通知は出さない／客先部品コードのリピート回数は数えない
//
//   ■ 再インポート（管理コンソール「🔁 注文書 再インポート」）
//     インポート用フォルダ（試作・量産）と会社サーバー連携フォルダの「処理済み」にあるPDFを
//     インポート用フォルダへ戻し、処理済みの記録を消す → 5分ごとの自動取込で順に読み直す。
//     再インポート中（最大48時間・全部終わるまで）は通知を止める。
//
//   ■ 重複の整理：同じ取引先・同じ注文番号の注文が複数ある時、1件（見積と紐づいた行→新しい行）を残して削除。
//     実行前に「管理シート」「注文書シート」のバックアップ（シートのコピー）を作る。
//
//   ※ Googleドライブ・スプレッドシートだけを操作する。会社サーバー（M:）には一切触れない。
// ============================================================

var ORI_REIMPORT_PROP = 'ORDER_REIMPORT_UNTIL';

function _oriNorm(s) {
  return String(s || '').normalize('NFKC').toUpperCase().replace(/[\s　]+/g, '');
}
function _oriClientKey(s) {
  var name = String(s || '').trim();
  if (!name) return '';
  if (typeof _cpmIsSelf === 'function' && _cpmIsSelf(name)) return '';
  try { if (typeof _cpmClientKey === 'function') { var k = _cpmClientKey(name); return k === '(取引先不明)' ? '' : _oriNorm(k); } } catch (e) {}
  return _oriNorm(name.replace(/株式会社|（株）|\(株\)|有限会社/g, ''));
}

/** 同じ取引先・同じ注文番号の注文の行番号（無ければ -1）。取引先が片方不明の時は注文番号だけで判定 */
function _oriFindOrderRow(mgmtSheet, orderNo, client) {
  var no = _oriNorm(orderNo);
  if (no.length < 3) return -1;
  var last = mgmtSheet.getLastRow();
  if (last <= 1) return -1;
  var ck = _oriClientKey(client);
  var data = mgmtSheet.getRange(2, 1, last - 1, MGMT_COLS.ORDER_TYPE).getValues();
  var loose = -1; // 取引先が片方不明の行（取引先が一致する行を優先）
  for (var i = data.length - 1; i >= 0; i--) {
    if (_oriNorm(data[i][MGMT_COLS.ORDER_NO - 1]) !== no) continue;
    var rk = _oriClientKey(data[i][MGMT_COLS.CLIENT - 1]);
    if (ck && rk && ck === rk) return i + 2;
    if ((!ck || !rk) && loose < 0) loose = i + 2;
  }
  return loose;
}

/** 注文書シートから、その注文（管理ID＋注文番号）の明細を消す。20列目以降の値を行番号ごとに返す */
function _oriRemoveOrderLines(os, mgmtId, orderNo) {
  var out = { width: 0, byLine: {} };
  var last = os.getLastRow();
  if (last <= 1) return out;
  var lastCol = Math.max(os.getLastColumn(), 19);
  var data = os.getRange(2, 1, last - 1, lastCol).getValues();
  out.width = Math.max(0, lastCol - 19);
  var no = _oriNorm(orderNo), rows = [];
  data.forEach(function (r, i) {
    if (String(r[ORDER_COLS.MGMT_ID - 1]) !== String(mgmtId)) return;
    if (no && _oriNorm(r[ORDER_COLS.ORDER_NO - 1]) !== no) return;
    rows.push(i + 2);
    var ln = Number(r[ORDER_COLS.LINE_NO - 1]) || 0;
    if (out.width && ln && !out.byLine[ln]) {
      var ex = r.slice(19);
      if (typeof CPM_ORDER_CODE_COL !== 'undefined' && CPM_ORDER_CODE_COL - 20 < ex.length) ex[CPM_ORDER_CODE_COL - 20] = ''; // 部品コードは書き直す
      out.byLine[ln] = ex;
    }
  });
  _oriDeleteRows(os, rows);
  return out;
}

/** 行番号の配列を下からまとめて削除 */
function _oriDeleteRows(sheet, rows) {
  rows = rows.slice().sort(function (a, b) { return b - a; });
  var i = 0;
  while (i < rows.length) {
    var end = rows[i], start = end, j = i + 1;
    while (j < rows.length && rows[j] === start - 1) { start = rows[j]; j++; }
    sheet.deleteRows(start, end - start + 1);
    i = j;
  }
}

/** 以前取り込んだPDFのコピーをゴミ箱へ（システムが作った「日時_」付きのコピーで、他の行から使われていない物だけ） */
function _oriTrashOldCopy(mgmtSheet, oldUrl, exceptRow) {
  try {
    var id = _extractFileIdFromUrl(String(oldUrl));
    if (!id) return;
    var last = mgmtSheet.getLastRow();
    var used = mgmtSheet.getRange(2, MGMT_COLS.ORDER_PDF_URL, last - 1, 1).getValues().some(function (r, i) {
      return i + 2 !== exceptRow && String(r[0]).indexOf(id) >= 0;
    });
    if (used) return;
    var f = DriveApp.getFileById(id);
    if (!/^\d{14}_/.test(f.getName())) return; // 取り込み時に作ったコピー以外は触らない
    f.setTrashed(true);
    Logger.log('[REIMPORT] 古いPDFコピーをゴミ箱へ: ' + f.getName());
  } catch (e) { Logger.log('[REIMPORT trash] ' + e.message); }
}

/** 再インポート中か（通知を止める） */
function _oriIsReimporting() {
  var until = Number(PropertiesService.getScriptProperties().getProperty(ORI_REIMPORT_PROP) || 0);
  return until > Date.now();
}

/** 再インポートの対象フォルダ：[ {label, importId, processedFolder} ] */
function _oriTargets() {
  var list = [];
  var add = function (label, importId, processedId) {
    try {
      var imp = DriveApp.getFolderById(importId);
      var proc = null;
      if (processedId) { try { proc = DriveApp.getFolderById(processedId); } catch (e) {} }
      if (!proc) { var it = imp.getFoldersByName('処理済み'); if (it.hasNext()) proc = it.next(); }
      list.push({ label: label, imp: imp, proc: proc });
    } catch (e) { Logger.log('[REIMPORT target] ' + label + ': ' + e.message); }
  };
  add('注文書（試作）', CONFIG.IMPORT_ORDER_TRIAL_FOLDER_ID, CONFIG.PROCESSED_ORDER_TRIAL_FOLDER_ID || '');
  add('注文書（量産）', CONFIG.IMPORT_ORDER_MASS_FOLDER_ID, CONFIG.PROCESSED_ORDER_MASS_FOLDER_ID || '');
  try {
    var ids = (typeof _serverSyncFolders === 'function') ? _serverSyncFolders() : null;
    if (ids) { add('サーバー連携 注文書(試作)', ids.trial, ''); add('サーバー連携 注文書(量産)', ids.mass, ''); }
  } catch (e) {}
  return list;
}

function _oriCountPdf(folder) {
  var n = 0; if (!folder) return 0;
  var it = folder.getFilesByType(MimeType.PDF);
  while (it.hasNext()) { it.next(); n++; }
  return n;
}

/** 状況：各フォルダの「取込待ち」「処理済み」の件数 */
function apiOrderReimportStatus() {
  try {
    var rows = _oriTargets().map(function (t) {
      return { label: t.label, waiting: _oriCountPdf(t.imp), processed: _oriCountPdf(t.proc) };
    });
    var until = Number(PropertiesService.getScriptProperties().getProperty(ORI_REIMPORT_PROP) || 0);
    return { success: true, rows: rows, reimporting: until > Date.now(), until: until ? new Date(until).toISOString() : '' };
  } catch (e) { return { success: false, error: e.message }; }
}

/** 再インポート開始：処理済みのPDFをインポート用フォルダへ戻し、処理済みの記録を消す */
function apiOrderReimportStart() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, error: '自動取込の実行中です。少し待ってからもう一度押してください' };
  try {
    var ids = _getProcessedFileIds();
    var moved = 0, detail = [];
    _oriTargets().forEach(function (t) {
      var n = 0;
      if (t.proc) {
        var it = t.proc.getFilesByType(MimeType.PDF);
        while (it.hasNext()) {
          var f = it.next();
          delete ids[f.getId()];
          f.moveTo(t.imp);
          n++;
        }
      }
      // インポート用フォルダに残っている物も読み直す
      var it2 = t.imp.getFilesByType(MimeType.PDF);
      while (it2.hasNext()) delete ids[it2.next().getId()];
      moved += n;
      detail.push(t.label + ' ' + n + '件');
    });
    PropertiesService.getScriptProperties().setProperty(PROCESSED_IDS_KEY, JSON.stringify(ids));
    PropertiesService.getScriptProperties().setProperty(ORI_REIMPORT_PROP, String(Date.now() + 48 * 3600 * 1000));
    Logger.log('[REIMPORT] 再インポート開始: ' + detail.join('／'));
    return { success: true, moved: moved, detail: detail };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

/** processDriveImports の最後に呼ぶ：取込待ちが無くなったら再インポート終了（通知を再開） */
function _oriCheckReimportDone() {
  try {
    if (!PropertiesService.getScriptProperties().getProperty(ORI_REIMPORT_PROP)) return;
    var waiting = _oriTargets().reduce(function (s, t) { return s + _oriCountPdf(t.imp); }, 0);
    if (!waiting || !_oriIsReimporting()) {
      PropertiesService.getScriptProperties().deleteProperty(ORI_REIMPORT_PROP);
      Logger.log('[REIMPORT] 再インポートが終わりました（通知を再開）');
    }
  } catch (e) { Logger.log('[REIMPORT done] ' + e.message); }
}

/**
 * 重複注文の整理：同じ取引先・同じ注文番号の注文が複数ある時、1件を残して削除
 *   p.dryRun = true … 削除せず、対象だけ返す
 *   残す行：見積と紐づいた行（見積番号あり）→ 最後に更新された行
 */
function apiOrderDedupe(p) {
  var dry = !!(p && p.dryRun);
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, error: '他の処理が実行中です' };
  try {
    var ss = getSpreadsheet();
    var ms = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var os = ss.getSheetByName(CONFIG.SHEET_ORDERS);
    var last = ms.getLastRow();
    if (last <= 1) return { success: true, groups: [], removed: 0 };
    var data = ms.getRange(2, 1, last - 1, Math.max(ms.getLastColumn(), MGMT_COLS.UPDATED_AT)).getValues();

    // 注文番号ごとにまとめる（取引先不明の行は、同じ注文番号の取引先ありの行にまとめる）
    var byNo = {};
    data.forEach(function (r, i) {
      var no = _oriNorm(r[MGMT_COLS.ORDER_NO - 1]);
      if (no.length < 3 || !String(r[MGMT_COLS.ORDER_PDF_URL - 1] || '').trim()) return;
      (byNo[no] = byNo[no] || []).push({ row: i + 2, r: r, ck: _oriClientKey(r[MGMT_COLS.CLIENT - 1]) });
    });
    var groups = [];
    Object.keys(byNo).forEach(function (no) {
      var items = byNo[no];
      if (items.length < 2) return;
      var byClient = {};
      items.forEach(function (x) { if (x.ck) (byClient[x.ck] = byClient[x.ck] || []).push(x); });
      var keys = Object.keys(byClient);
      items.forEach(function (x) { if (!x.ck) { if (keys.length === 1) byClient[keys[0]].push(x); else (byClient[''] = byClient[''] || []).push(x); } });
      Object.keys(byClient).forEach(function (k) {
        var g = byClient[k];
        if (g.length < 2) return;
        g.sort(function (a, b) {
          var qa = String(a.r[MGMT_COLS.QUOTE_NO - 1] || '') ? 1 : 0, qb = String(b.r[MGMT_COLS.QUOTE_NO - 1] || '') ? 1 : 0;
          if (qa !== qb) return qb - qa;
          var ua = String(a.r[MGMT_COLS.UPDATED_AT - 1] || ''), ub = String(b.r[MGMT_COLS.UPDATED_AT - 1] || '');
          return ua < ub ? 1 : ua > ub ? -1 : b.row - a.row;
        });
        groups.push({ orderNo: String(g[0].r[MGMT_COLS.ORDER_NO - 1]), client: String(g[0].r[MGMT_COLS.CLIENT - 1] || ''),
                      keep: g[0], drop: g.slice(1) });
      });
    });

    var preview = groups.map(function (g) {
      return { orderNo: g.orderNo, client: g.client, keepId: String(g.keep.r[MGMT_COLS.ID - 1]),
               dropIds: g.drop.map(function (x) { return String(x.r[MGMT_COLS.ID - 1]); }) };
    });
    var removeCount = groups.reduce(function (s, g) { return s + g.drop.length; }, 0);
    if (dry || !removeCount) return { success: true, dryRun: dry, groups: preview, removed: 0, toRemove: removeCount };

    // バックアップ（シートのコピー）
    var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmm');
    ms.copyTo(ss).setName('バックアップ_管理シート_' + stamp);
    os.copyTo(ss).setName('バックアップ_注文書シート_' + stamp);

    var dropRows = [], dropIds = {}, dropPdf = [];
    groups.forEach(function (g) {
      g.drop.forEach(function (x) {
        dropRows.push(x.row);
        dropIds[String(x.r[MGMT_COLS.ID - 1])] = true;
        var u = String(x.r[MGMT_COLS.ORDER_PDF_URL - 1] || ''), ku = String(g.keep.r[MGMT_COLS.ORDER_PDF_URL - 1] || '');
        if (u && u !== ku) dropPdf.push(u);
      });
    });
    // 注文明細を削除
    if (os.getLastRow() > 1) {
      var ol = os.getRange(2, 1, os.getLastRow() - 1, 1).getValues(), orows = [];
      ol.forEach(function (r, i) { if (dropIds[String(r[0])]) orows.push(i + 2); });
      _oriDeleteRows(os, orows);
    }
    _oriDeleteRows(ms, dropRows);
    // 使われなくなったPDFコピーをゴミ箱へ
    dropPdf.forEach(function (u) { _oriTrashOldCopy(ms, u, -1); });
    Logger.log('[DEDUPE] 重複注文 ' + dropRows.length + '件を削除（バックアップ: ' + stamp + '）');
    return { success: true, groups: preview, removed: dropRows.length, backup: stamp };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}
