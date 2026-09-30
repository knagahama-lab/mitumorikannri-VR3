// ============================================================
// 35_client_fix.gs
// 注文の取引先が「自社（受注側）」になるのを防ぐ／既存データを付け替える
//
//   原因：注文書OCRが宛先（殿・御中＝自社）を取引先として読むことがある
//   対策：
//     ① 取込時補正 cpmFixOrderClient … 自社・空欄なら 発行元(issuerName) → PDF保存フォルダ名 の順で補う
//     ② 付け替え候補 apiCpmClientFixCandidates … 部品コード・フォルダ名・件名から取引先を推定
//     ③ 付け替え実行 apiCpmClientFix … 管理シートを更新し、客先品番マスタを再集計（手入力項目は引き継ぎ）
//   自社名はスクリプトプロパティ SELF_COMPANY_NAMES（カンマ区切り）で変更可。既定は「サン電子」。
// ============================================================

function _cpmSelfNames() {
  var props = PropertiesService.getScriptProperties();
  var names = String(props.getProperty('SELF_COMPANY_NAMES') || 'サン電子').split(',');
  try { var co = JSON.parse(props.getProperty('QE_COMPANY_INFO') || '{}'); if (co.name) names.push(co.name); } catch (e) {}
  return names.map(function(n) { return _cpmCompanyNorm(n); }).filter(Boolean);
}
function _cpmCompanyNorm(s) {
  return String(s || '').normalize('NFKC').replace(/株式会社|\(株\)|（株）|有限会社|\s|　/g, '');
}
function _cpmIsSelf(name, selfNames) {
  var n = _cpmCompanyNorm(name);
  if (!n) return false;
  return (selfNames || _cpmSelfNames()).some(function(s) { return n.indexOf(s) >= 0 || s.indexOf(n) >= 0; });
}
function _cpmIsBadClient(name, selfNames) {
  var s = String(name || '').trim();
  return !s || s === '(取引先不明)' || _cpmIsSelf(s, selfNames);
}

/** 取引先マスタのキーワードに一致する分類を返す（未一致は null） */
function _cpmMatchClientMaster(text) {
  var t = String(text || '').normalize('NFKC');
  if (!t) return null;
  var list = [];
  try { list = getClientMasterList(); } catch (e) {}
  for (var i = 0; i < list.length; i++) {
    var c = list[i];
    if (c.isFallback) continue;
    var kws = (c.keywords && c.keywords.length) ? c.keywords : [c.name];
    for (var j = 0; j < kws.length; j++) {
      if (kws[j] && t.indexOf(String(kws[j]).normalize('NFKC')) >= 0) return c;
    }
  }
  return null;
}

/** PDFの保存フォルダ名（3階層まで）から取引先を推定 */
var _cpmFolderCache = {};
function _cpmClientFromPdf(pdfUrl) {
  var m = String(pdfUrl || '').match(/[-\w]{25,}/);
  if (!m) return null;
  try {
    var names = [];
    var it = DriveApp.getFileById(m[0]).getParents();
    var folder = it.hasNext() ? it.next() : null;
    for (var depth = 0; folder && depth < 3; depth++) {
      var id = folder.getId();
      if (!_cpmFolderCache[id]) {
        var p = folder.getParents();
        _cpmFolderCache[id] = { name: folder.getName(), parent: p.hasNext() ? p.next() : null };
      }
      names.push(_cpmFolderCache[id].name);
      folder = _cpmFolderCache[id].parent;
    }
    var hitName = names.filter(function(n) { return _cpmMatchClientMaster(n); })[0];
    return hitName ? { client: _cpmMatchClientMaster(hitName), via: 'フォルダ「' + hitName + '」' } : null;
  } catch (e) {
    return null;
  }
}

/** 管理シートで実際に使われている表記（例：株式会社藤商事）を分類ごとに最多のもの1つ */
function _cpmCanonicalClientNames() {
  var counts = {};
  getAllMgmtData().forEach(function(r) {
    var name = String(r[MGMT_COLS.CLIENT - 1] || '').trim();
    var c = name ? _cpmMatchClientMaster(name) : null;
    if (!c) return;
    counts[c.name] = counts[c.name] || {};
    counts[c.name][name] = (counts[c.name][name] || 0) + 1;
  });
  var out = {};
  Object.keys(counts).forEach(function(k) {
    out[k] = Object.keys(counts[k]).sort(function(a, b) { return counts[k][b] - counts[k][a]; })[0];
  });
  return out;
}

/**
 * ★ 注文書取込時の補正（02_ocr_and_processing.gs の _saveOrderData 先頭で呼ぶ）
 */
function cpmFixOrderClient(ocr, pdfUrl) {
  if (!ocr) return ocr;
  var self = _cpmSelfNames();
  if (!_cpmIsBadClient(ocr.clientName, self)) return ocr;
  var original = ocr.clientName || '';
  if (ocr.issuerName && !_cpmIsSelf(ocr.issuerName, self)) {
    ocr.clientName = ocr.issuerName;
  } else {
    var g = _cpmClientFromPdf(pdfUrl);
    ocr.clientName = g ? (_cpmCanonicalClientNames()[g.client.name] || g.client.name) : '';
  }
  Logger.log('[cpmFixOrderClient] 取引先を補正: "' + original + '" → "' + ocr.clientName + '"');
  return ocr;
}

/** 付け替え候補：取引先が自社・空欄の注文 */
function apiCpmClientFixCandidates() {
  try {
    var self = _cpmSelfNames();
    var canon = _cpmCanonicalClientNames();
    var codeClient = {}; // 部品コード → 自社以外で登録済みの取引先
    _cpmRows().forEach(function(r) {
      if (!_cpmIsBadClient(r[CPM.CLIENT], self)) codeClient[_cpmCode(r[CPM.CODE])] = r[CPM.CLIENT];
    });
    var linesByMgmt = {};
    _cpmOrderLines().forEach(function(l) { if (l.code) (linesByMgmt[l.mgmtId] = linesByMgmt[l.mgmtId] || []).push(l.code); });

    var orders = getAllMgmtData().map(_rowToObject).filter(function(o) { return o.orderNo && _cpmIsBadClient(o.client, self); });
    var started = Date.now(), driveSkipped = 0;
    var items = orders.map(function(o) {
      var sug = '', via = '';
      var codes = linesByMgmt[o.id] || [];
      var byCode = codes.map(function(c) { return codeClient[c]; }).filter(Boolean)[0];
      if (byCode) { sug = byCode; via = '同じ部品コードの登録済み取引先'; }
      if (!sug) {
        if (Date.now() - started < 240000) { // 実行時間の上限（6分）に配慮
          var g = _cpmClientFromPdf(o.orderPdfUrl);
          if (g) { sug = g.client.name; via = g.via; }
        } else driveSkipped++;
      }
      if (!sug) {
        var t = _cpmMatchClientMaster([o.subject, o.memo].join(' '));
        if (t) { sug = t.name; via = '件名'; }
      }
      return { mgmtId: o.id, orderNo: o.orderNo, orderDate: o.orderDate, subject: o.subject, client: o.client,
               modelCode: o.modelCode, codes: codes.length, suggest: sug ? (canon[sug] || sug) : '', via: via, orderPdfUrl: o.orderPdfUrl };
    });
    var clients = [];
    try { clients = getClientMasterList().filter(function(c) { return !c.isFallback; }).map(function(c) { return canon[c.name] || c.name; }); } catch (e) {}
    return { success: true, items: items, clients: clients, selfNames: self, driveSkipped: driveSkipped };
  } catch (e) {
    Logger.log('[apiCpmClientFixCandidates] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

/**
 * 付け替え実行：p.items = [{mgmtId, client}]
 *   1) 管理シートの顧客名を更新
 *   2) 取引先が自社・不明の客先品番マスタ行を退避・削除して再集計
 *      （弊社見積No.・見積単価・弊社基板名・備考・確認済は部品コード単位で引き継ぎ）
 */
function apiCpmClientFix(p) {
  try {
    var items = ((p && p.items) || []).filter(function(it) { return it.mgmtId && String(it.client || '').trim(); });
    if (!items.length) return { success: false, error: '付け替える注文がありません' };
    var self = _cpmSelfNames();
    var fixed = 0, stash = {};

    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var sheet = getSpreadsheet().getSheetByName(CONFIG.SHEET_MANAGEMENT);
      var last = sheet.getLastRow();
      var ids = sheet.getRange(2, MGMT_COLS.ID, last - 1, 1).getValues().map(function(r) { return String(r[0]); });
      var clientCol = sheet.getRange(2, MGMT_COLS.CLIENT, last - 1, 1).getValues();
      items.forEach(function(it) {
        var i = ids.indexOf(String(it.mgmtId));
        if (i < 0) return;
        clientCol[i][0] = String(it.client).trim(); fixed++;
      });
      sheet.getRange(2, MGMT_COLS.CLIENT, last - 1, 1).setValues(clientCol);

      var sh = _cpmSheet();
      var rows = _cpmRows();
      var keep = [];
      rows.forEach(function(r) {
        if (_cpmIsBadClient(r[CPM.CLIENT], self)) stash[_cpmCode(r[CPM.CODE])] = r;
        else keep.push(r);
      });
      if (rows.length) sh.getRange(2, 1, rows.length, CPM_HEADERS.length).clearContent();
      if (keep.length) sh.getRange(2, 1, keep.length, CPM_HEADERS.length).setValues(keep);
    } finally {
      lock.releaseLock();
    }

    var rb = apiCpmRebuild(); // 付け替え後の顧客名で注文書シートから再集計（内部でロック取得）
    if (!rb.success) return rb;

    var carried = 0;
    lock.waitLock(30000);
    try {
      var sh2 = _cpmSheet();
      _cpmRows().forEach(function(r, i) {
        var s = stash[_cpmCode(r[CPM.CODE])];
        if (!s) return;
        var changed = false;
        [CPM.QUOTE_NO, CPM.QUOTE_ID, CPM.QUOTE_PRICE, CPM.BOARD, CPM.NOTE].forEach(function(c) {
          if (!String(r[c]).trim() && String(s[c]).trim()) { r[c] = s[c]; changed = true; }
        });
        if (String(s[CPM.CONFIRMED]) === 'TRUE' && String(r[CPM.CONFIRMED]) !== 'TRUE') { r[CPM.CONFIRMED] = 'TRUE'; changed = true; }
        if (changed) { sh2.getRange(i + 2, 1, 1, CPM_HEADERS.length).setValues([r]); carried++; }
      });
    } finally {
      lock.releaseLock();
    }
    var remaining = _cpmRows().filter(function(r) { return _cpmIsBadClient(r[CPM.CLIENT], self); }).length;
    return { success: true, fixedOrders: fixed, created: rb.created, updated: rb.updated, carried: carried, remaining: remaining };
  } catch (e) {
    Logger.log('[apiCpmClientFix] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}
