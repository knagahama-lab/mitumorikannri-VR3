// ============================================================
// 31_supplier_docs.gs
// 仕入先・商社の見積書／値上げ書面 管理
//
// 28_quote_documents.gs の「見積関連書類」は 1書類 = 1案件(mgmtId) の
// 従属レコードだったが、仕入先見積書は
//   ・1枚の仕入先見積書が複数の弊社見積に使われる
//   ・1件の弊社見積が複数の仕入先見積書を参照する
// という N:N になるため、書類本体（仕入先書類シート）と
// 紐づけ（仕入先書類リンクシート）を分離している。
//
// これにより
//   ・弊社見積 → 参照した仕入先見積書   （順引き）
//   ・仕入先見積書 → 使われた弊社見積   （逆引き）
// の双方向検索が可能になる。
//
// 値上げ書面は旧単価／新単価／適用開始日を持たせ、
// apiPriceIncreaseTimeline() で時系列に並べて参照できる。
// ============================================================

// ── シート名 ──
var SDOC_SHEET      = '仕入先書類';
var SDOC_LINK_SHEET = '仕入先書類リンク';

// ── 書類種別 ──
//    値上げ系（PRICE_CHANGE）は時系列タイムラインの対象になる。
var SDOC_TYPES = [
  '仕入先見積書',
  '商社見積書',
  '値上げ通知書',
  '価格改定通知',
  '単価表',
  'その他',
];

// 値上げ時系列の対象とする書類種別
var SDOC_PRICE_CHANGE_TYPES = ['値上げ通知書', '価格改定通知'];

var SDOC_HEADERS = [
  '書類ID', '書類種別', '仕入先名', '商社名', '件名', '発行日', '適用開始日',
  '対象品名', '型番', '数量', '旧単価', '新単価', '増減額', '増減率(%)', '通貨',
  'ファイル名', 'Drive URL', '備考', '登録日時', '更新日時',
];

var SDOC_LINK_HEADERS = [
  'リンクID', '書類ID', '対象種別', '対象ID', '見積番号', '備考', '登録日時',
];

// ============================================================
// シート初期化
// ============================================================

function _initSupplierDocSheet() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(SDOC_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SDOC_SHEET);
    var hr = sheet.getRange(1, 1, 1, SDOC_HEADERS.length);
    hr.setValues([SDOC_HEADERS]);
    hr.setBackground('#E8F5E9');
    hr.setFontWeight('bold');
    hr.setFontSize(10);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(5,  220); // 件名
    sheet.setColumnWidth(8,  200); // 対象品名
    sheet.setColumnWidth(17, 240); // Drive URL
  }
  return sheet;
}

function _initSupplierDocLinkSheet() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(SDOC_LINK_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SDOC_LINK_SHEET);
    var hr = sheet.getRange(1, 1, 1, SDOC_LINK_HEADERS.length);
    hr.setValues([SDOC_LINK_HEADERS]);
    hr.setBackground('#FFF8E1');
    hr.setFontWeight('bold');
    hr.setFontSize(10);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// 初期セットアップ（管理コンソールから一度だけ実行すればよい）
function initSupplierDocSheets() {
  _initSupplierDocSheet();
  _initSupplierDocLinkSheet();
  return { success: true };
}

// ============================================================
// 内部ヘルパー
// ============================================================

function _sdocGenId(prefix) {
  return prefix + '-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') +
    '-' + (Math.floor(Math.random() * 9000) + 1000);
}

// 日付を yyyy-MM-dd の比較可能キーに正規化（空文字は '' のまま）
function _sdocDateKey(v) {
  var s = _toDateStr(v);
  if (!s) return '';
  return String(s).replace(/\//g, '-');
}

function _sdocNum(v) {
  var n = Number(String(v == null ? '' : v).replace(/[,¥\s]/g, ''));
  return isNaN(n) ? 0 : n;
}

function _sdocRowToObj(r) {
  var oldPrice = _sdocNum(r[10]);
  var newPrice = _sdocNum(r[11]);
  return {
    id:         String(r[0]  || ''),
    docType:    String(r[1]  || ''),
    supplier:   String(r[2]  || ''),
    trader:     String(r[3]  || ''),
    subject:    String(r[4]  || ''),
    issueDate:  _toDateStr(r[5]),
    effectDate: _toDateStr(r[6]),
    itemName:   String(r[7]  || ''),
    modelNo:    String(r[8]  || ''),
    qty:        _sdocNum(r[9]),
    oldPrice:   oldPrice,
    newPrice:   newPrice,
    diff:       _sdocNum(r[12]),
    diffRate:   _sdocNum(r[13]),
    currency:   String(r[14] || 'JPY'),
    fileName:   String(r[15] || ''),
    url:        String(r[16] || ''),
    memo:       String(r[17] || ''),
    createdAt:  _toDateStr(r[18]),
    updatedAt:  _toDateStr(r[19]),
  };
}

function _sdocLinkRowToObj(r) {
  return {
    id:         String(r[0] || ''),
    docId:      String(r[1] || ''),
    targetType: String(r[2] || 'mgmt'),
    targetId:   String(r[3] || ''),
    quoteNo:    String(r[4] || ''),
    memo:       String(r[5] || ''),
    createdAt:  _toDateStr(r[6]),
  };
}

// 全書類を読み込む（内部用）
function _sdocReadAll() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(SDOC_SHEET);
  if (!sheet) return [];
  var last = sheet.getLastRow();
  if (last <= 1) return [];
  return sheet.getRange(2, 1, last - 1, SDOC_HEADERS.length).getValues()
    .filter(function(r) { return String(r[0]).trim() !== ''; })
    .map(function(r) { return _sdocRowToObj(r); });
}

// 全リンクを読み込む（内部用）
function _sdocReadAllLinks() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(SDOC_LINK_SHEET);
  if (!sheet) return [];
  var last = sheet.getLastRow();
  if (last <= 1) return [];
  return sheet.getRange(2, 1, last - 1, SDOC_LINK_HEADERS.length).getValues()
    .filter(function(r) { return String(r[0]).trim() !== ''; })
    .map(function(r) { return _sdocLinkRowToObj(r); });
}

// 弊社見積（管理シート + 見積台帳）の索引を作る。
// 逆引き（仕入先書類 → 弊社見積）で見積No.や宛先を表示するために使う。
function _sdocBuildQuoteIndex() {
  var idx = { byId: {}, byQuoteNo: {} };
  var ss  = getSpreadsheet();

  try {
    var mgmt = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    if (mgmt && mgmt.getLastRow() > 1) {
      var mLast = mgmt.getLastRow();
      var mRows = mgmt.getRange(2, 1, mLast - 1, mgmt.getLastColumn()).getValues();
      mRows.forEach(function(r) {
        var id = String(r[MGMT_COLS.ID - 1] || '').trim();
        if (!id) return;
        var o = {
          targetType: 'mgmt',
          targetId:   id,
          quoteNo:    String(r[MGMT_COLS.QUOTE_NO - 1]     || '').trim(),
          subject:    String(r[MGMT_COLS.SUBJECT - 1]      || ''),
          client:     String(r[MGMT_COLS.CLIENT - 1]       || ''),
          issueDate:  _toDateStr(r[MGMT_COLS.QUOTE_DATE - 1]),
          amount:     _sdocNum(r[MGMT_COLS.QUOTE_AMOUNT - 1]),
          modelCode:  String(r[MGMT_COLS.MODEL_CODE - 1]   || ''),
          pdfUrl:     String(r[MGMT_COLS.QUOTE_PDF_URL - 1]|| ''),
        };
        idx.byId[id] = o;
        if (o.quoteNo && !idx.byQuoteNo[o.quoteNo]) idx.byQuoteNo[o.quoteNo] = o;
      });
    }
  } catch (e) { Logger.log('[_sdocBuildQuoteIndex mgmt] ' + e.message); }

  try {
    var ledger = ss.getSheetByName(CONFIG.SHEET_LEDGER);
    if (ledger && ledger.getLastRow() > 1) {
      var lLast = ledger.getLastRow();
      var lRows = ledger.getRange(2, 1, lLast - 1, ledger.getLastColumn()).getValues();
      lRows.forEach(function(r) {
        var id = String(r[LEDGER_COLS.LEDGER_ID - 1] || '').trim();
        if (!id) return;
        var o = {
          targetType: 'ledger',
          targetId:   id,
          quoteNo:    String(r[LEDGER_COLS.QUOTE_NO - 1] || '').trim(),
          subject:    String(r[LEDGER_COLS.SUBJECT - 1]  || ''),
          client:     String(r[LEDGER_COLS.DEST - 1]     || ''),
          issueDate:  _toDateStr(r[LEDGER_COLS.ISSUE_DATE - 1]),
          amount:     _sdocNum(r[LEDGER_COLS.AMOUNT - 1]),
          modelCode:  String(r[LEDGER_COLS.MACHINE_CODE - 1] || ''),
          pdfUrl:     String(r[LEDGER_COLS.SAVE_URL - 1] || ''),
        };
        idx.byId[id] = o;
        if (o.quoteNo && !idx.byQuoteNo[o.quoteNo]) idx.byQuoteNo[o.quoteNo] = o;
      });
    }
  } catch (e) { Logger.log('[_sdocBuildQuoteIndex ledger] ' + e.message); }

  return idx;
}

// ============================================================
// API: 仕入先書類 一覧（フィルタ付き）
//   payload: { keyword, docType, supplier, dateFrom, dateTo, priceChangeOnly }
//   各書類に紐づく弊社見積（links）も付けて返す。
// ============================================================
function apiSupplierDocsList(payload) {
  try {
    payload = payload || {};
    var kw       = String(payload.keyword  || '').trim().toLowerCase();
    var docType  = String(payload.docType  || '').trim();
    var supplier = String(payload.supplier || '').trim();
    var from     = _sdocDateKey(payload.dateFrom);
    var to       = _sdocDateKey(payload.dateTo);

    var docs  = _sdocReadAll();
    var links = _sdocReadAllLinks();
    var qIdx  = _sdocBuildQuoteIndex();

    // 書類ID → 紐づく弊社見積[]
    var linkMap = {};
    links.forEach(function(l) {
      if (!linkMap[l.docId]) linkMap[l.docId] = [];
      var q = qIdx.byId[l.targetId] || (l.quoteNo ? qIdx.byQuoteNo[l.quoteNo] : null);
      linkMap[l.docId].push({
        linkId:     l.id,
        targetType: l.targetType,
        targetId:   l.targetId,
        quoteNo:    l.quoteNo || (q ? q.quoteNo : ''),
        subject:    q ? q.subject : '',
        client:     q ? q.client  : '',
        issueDate:  q ? q.issueDate : '',
        amount:     q ? q.amount  : 0,
        pdfUrl:     q ? q.pdfUrl  : '',
        memo:       l.memo,
      });
    });

    var items = docs.filter(function(d) {
      if (docType  && d.docType  !== docType)  return false;
      if (supplier && d.supplier !== supplier && d.trader !== supplier) return false;
      if (payload.priceChangeOnly && SDOC_PRICE_CHANGE_TYPES.indexOf(d.docType) < 0) return false;
      // 期間は「適用開始日 → なければ発行日」で判定
      var dk = _sdocDateKey(d.effectDate) || _sdocDateKey(d.issueDate);
      if (from && (!dk || dk < from)) return false;
      if (to   && (!dk || dk > to))   return false;
      if (kw) {
        var hay = [d.supplier, d.trader, d.subject, d.itemName, d.modelNo, d.fileName, d.memo, d.docType]
          .join(' ').toLowerCase();
        // 紐づけ先の見積No.でも引けるようにする（双方向検索）
        // ここも小文字化してから連結しないと、見積No.の英字で引けなくなる
        (linkMap[d.id] || []).forEach(function(l) {
          hay += ' ' + [l.quoteNo, l.subject, l.client].join(' ').toLowerCase();
        });
        if (hay.indexOf(kw) < 0) return false;
      }
      return true;
    }).map(function(d) {
      d.links     = linkMap[d.id] || [];
      d.linkCount = d.links.length;
      return d;
    });

    // 発行日の新しい順
    items.sort(function(a, b) {
      var ka = _sdocDateKey(a.issueDate) || _sdocDateKey(a.effectDate) || '';
      var kb = _sdocDateKey(b.issueDate) || _sdocDateKey(b.effectDate) || '';
      if (ka === kb) return String(b.id).localeCompare(String(a.id));
      return kb.localeCompare(ka);
    });

    // 絞り込み用の仕入先候補
    var suppliers = {};
    docs.forEach(function(d) {
      if (d.supplier) suppliers[d.supplier] = true;
      if (d.trader)   suppliers[d.trader]   = true;
    });

    return {
      success:   true,
      items:     items,
      docTypes:  SDOC_TYPES,
      suppliers: Object.keys(suppliers).sort(),
      priceChangeTypes: SDOC_PRICE_CHANGE_TYPES,
    };
  } catch (e) { return { success: false, error: e.message }; }
}

// ============================================================
// API: 仕入先書類 保存（新規/更新）
// ============================================================
function apiSupplierDocSave(payload) {
  try {
    payload = payload || {};
    var docType = String(payload.docType || '').trim() || SDOC_TYPES[0];
    if (SDOC_TYPES.indexOf(docType) < 0) return { success: false, error: '不正な書類種別です' };
    var supplier = String(payload.supplier || '').trim();
    var trader   = String(payload.trader   || '').trim();
    if (!supplier && !trader) return { success: false, error: '仕入先名または商社名は必須です' };

    var sheet = _initSupplierDocSheet();
    var now   = nowJST();
    var isNew = !payload.id || String(payload.id).trim() === '';
    var id    = isNew ? _sdocGenId('SD') : String(payload.id).trim();

    var oldPrice = _sdocNum(payload.oldPrice);
    var newPrice = _sdocNum(payload.newPrice);
    var diff     = (oldPrice || newPrice) ? (newPrice - oldPrice) : 0;
    var diffRate = oldPrice > 0 ? Math.round(((newPrice - oldPrice) / oldPrice) * 1000) / 10 : 0;

    var row = [
      id,
      docType,
      supplier,
      trader,
      String(payload.subject    || ''),
      String(payload.issueDate  || ''),
      String(payload.effectDate || ''),
      String(payload.itemName   || ''),
      String(payload.modelNo    || ''),
      _sdocNum(payload.qty),
      oldPrice,
      newPrice,
      diff,
      diffRate,
      String(payload.currency || 'JPY'),
      String(payload.fileName || ''),
      String(payload.url      || ''),
      String(payload.memo     || ''),
    ];

    if (isNew) {
      row.push(now); // 登録日時
      row.push(now); // 更新日時
      sheet.appendRow(row);
    } else {
      var last = sheet.getLastRow();
      if (last <= 1) return { success: false, error: 'データなし' };
      var ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat().map(function(v) { return String(v); });
      var idx = ids.indexOf(id);
      if (idx < 0) return { success: false, error: 'IDが見つかりません: ' + id };
      var rowNum  = idx + 2;
      var created = sheet.getRange(rowNum, 19).getValue();
      row.push(created || now);
      row.push(now);
      sheet.getRange(rowNum, 1, 1, SDOC_HEADERS.length).setValues([row]);
    }
    return { success: true, id: id };
  } catch (e) { return { success: false, error: e.message }; }
}

// ============================================================
// API: 仕入先書類 ファイルアップロード + 保存
//   既存レコードへのファイル差し替えにも使える（payload.id 指定時）。
// ============================================================
function apiSupplierDocUpload(payload) {
  try {
    payload = payload || {};
    if (!payload.base64Data || !payload.fileName) return { success: false, error: 'ファイルデータ不足' };

    var folder   = DriveApp.getFolderById(CONFIG.WEB_UPLOAD_FOLDER_ID);
    var mimeType = payload.mimeType || 'application/octet-stream';
    var safeName = String(payload.fileName).replace(/[/\\:*?"<>|]/g, '_');
    var prefix   = (String(payload.supplier || payload.trader || 'SUPPLIER')).replace(/[/\\:*?"<>|]/g, '_');
    var blob     = Utilities.newBlob(
      Utilities.base64Decode(payload.base64Data), mimeType,
      prefix + '_' + (payload.docType || '仕入先書類') + '_' + safeName
    );
    var file = folder.createFile(blob);

    var saveRes = apiSupplierDocSave(Object.assign({}, payload, {
      fileName: payload.fileName,
      url:      file.getUrl(),
    }));
    if (!saveRes.success) return saveRes;
    return { success: true, id: saveRes.id, url: file.getUrl(), fileName: payload.fileName };
  } catch (e) { return { success: false, error: e.message }; }
}

// ============================================================
// API: 仕入先書類 削除（紐づけも同時に削除）
// ============================================================
function apiSupplierDocDelete(payload) {
  try {
    var id = String((payload || {}).id || '').trim();
    if (!id) return { success: false, error: 'IDが必要です' };
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(SDOC_SHEET);
    if (!sheet) return { success: false, error: 'シートがありません' };
    var last = sheet.getLastRow();
    if (last <= 1) return { success: false, error: 'データなし' };
    var ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat().map(function(v) { return String(v); });
    var idx = ids.indexOf(id);
    if (idx < 0) return { success: false, error: 'IDが見つかりません' };
    sheet.deleteRow(idx + 2);

    // 紐づけレコードも掃除（下から消して行ズレを防ぐ）
    var lSheet = ss.getSheetByName(SDOC_LINK_SHEET);
    if (lSheet && lSheet.getLastRow() > 1) {
      var lLast   = lSheet.getLastRow();
      var docIds  = lSheet.getRange(2, 2, lLast - 1, 1).getValues().flat().map(function(v) { return String(v); });
      for (var i = docIds.length - 1; i >= 0; i--) {
        if (docIds[i] === id) lSheet.deleteRow(i + 2);
      }
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// ============================================================
// API: 紐づけ追加（仕入先書類 ⇔ 弊社見積）
//   payload: { docId, targetId, targetType('mgmt'|'ledger'), quoteNo, memo }
//   targetId が未指定でも quoteNo から解決を試みる。
// ============================================================
function apiSupplierDocLinkSave(payload) {
  try {
    payload = payload || {};
    var docId = String(payload.docId || '').trim();
    if (!docId) return { success: false, error: '書類IDが必要です' };

    var targetId   = String(payload.targetId || '').trim();
    var quoteNo    = String(payload.quoteNo  || '').trim();
    var targetType = String(payload.targetType || '').trim();

    if (!targetId && !quoteNo) return { success: false, error: '紐づける見積を指定してください' };

    var qIdx = _sdocBuildQuoteIndex();
    if (!targetId && quoteNo) {
      var found = qIdx.byQuoteNo[quoteNo];
      if (!found) return { success: false, error: '見積番号が見つかりません: ' + quoteNo };
      targetId   = found.targetId;
      targetType = found.targetType;
    }
    if (!targetType) {
      targetType = (qIdx.byId[targetId] && qIdx.byId[targetId].targetType) || 'mgmt';
    }
    if (!quoteNo && qIdx.byId[targetId]) quoteNo = qIdx.byId[targetId].quoteNo;

    var sheet = _initSupplierDocLinkSheet();

    // 同じ組み合わせの重複を防ぐ
    var dup = _sdocReadAllLinks().some(function(l) {
      return l.docId === docId && l.targetId === targetId;
    });
    if (dup) return { success: false, error: 'すでに紐づけ済みです' };

    var id = _sdocGenId('SDL');
    sheet.appendRow([id, docId, targetType, targetId, quoteNo, String(payload.memo || ''), nowJST()]);
    return { success: true, id: id };
  } catch (e) { return { success: false, error: e.message }; }
}

// ============================================================
// API: 紐づけ削除
// ============================================================
function apiSupplierDocLinkDelete(payload) {
  try {
    var id = String((payload || {}).id || '').trim();
    if (!id) return { success: false, error: 'リンクIDが必要です' };
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(SDOC_LINK_SHEET);
    if (!sheet) return { success: false, error: 'シートがありません' };
    var last = sheet.getLastRow();
    if (last <= 1) return { success: false, error: 'データなし' };
    var ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat().map(function(v) { return String(v); });
    var idx = ids.indexOf(id);
    if (idx < 0) return { success: false, error: 'リンクIDが見つかりません' };
    sheet.deleteRow(idx + 2);
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// ============================================================
// API: 順引き — 弊社見積 → 参照している仕入先書類
//   payload: { targetId, quoteNo }
// ============================================================
function apiSupplierDocsForQuote(payload) {
  try {
    payload = payload || {};
    var targetId = String(payload.targetId || '').trim();
    var quoteNo  = String(payload.quoteNo  || '').trim();
    if (!targetId && !quoteNo) return { success: false, error: '見積の指定が必要です' };

    var links = _sdocReadAllLinks().filter(function(l) {
      if (targetId && l.targetId === targetId) return true;
      if (quoteNo  && l.quoteNo  === quoteNo)  return true;
      return false;
    });

    var docMap = {};
    _sdocReadAll().forEach(function(d) { docMap[d.id] = d; });

    var items = links.map(function(l) {
      var d = docMap[l.docId];
      if (!d) return null;
      d = Object.assign({}, d);
      d.linkId   = l.id;
      d.linkMemo = l.memo;
      return d;
    }).filter(function(d) { return !!d; });

    items.sort(function(a, b) {
      var ka = _sdocDateKey(a.issueDate) || '';
      var kb = _sdocDateKey(b.issueDate) || '';
      return kb.localeCompare(ka);
    });

    return { success: true, items: items, docTypes: SDOC_TYPES };
  } catch (e) { return { success: false, error: e.message }; }
}

// ============================================================
// API: 逆引き — 仕入先書類 → その書類を使った弊社見積
//   payload: { docId }
// ============================================================
function apiSupplierQuotesForDoc(payload) {
  try {
    var docId = String((payload || {}).docId || '').trim();
    if (!docId) return { success: false, error: '書類IDが必要です' };

    var qIdx  = _sdocBuildQuoteIndex();
    var items = _sdocReadAllLinks()
      .filter(function(l) { return l.docId === docId; })
      .map(function(l) {
        var q = qIdx.byId[l.targetId] || (l.quoteNo ? qIdx.byQuoteNo[l.quoteNo] : null);
        return {
          linkId:     l.id,
          targetType: l.targetType,
          targetId:   l.targetId,
          quoteNo:    l.quoteNo || (q ? q.quoteNo : ''),
          subject:    q ? q.subject   : '',
          client:     q ? q.client    : '',
          issueDate:  q ? q.issueDate : '',
          amount:     q ? q.amount    : 0,
          modelCode:  q ? q.modelCode : '',
          pdfUrl:     q ? q.pdfUrl    : '',
          memo:       l.memo,
          missing:    !q,
        };
      });

    items.sort(function(a, b) {
      var ka = _sdocDateKey(a.issueDate) || '';
      var kb = _sdocDateKey(b.issueDate) || '';
      return kb.localeCompare(ka);
    });

    return { success: true, items: items };
  } catch (e) { return { success: false, error: e.message }; }
}

// ============================================================
// API: 見積書一覧の「🏭 仕入」列用 — 対象ID → 仕入先書類[] の一括マップ
// ============================================================
function apiSupplierDocsMapAll() {
  try {
    var docMap = {};
    _sdocReadAll().forEach(function(d) { docMap[d.id] = d; });

    var map = {};
    _sdocReadAllLinks().forEach(function(l) {
      var d = docMap[l.docId];
      if (!d) return;
      // mgmtId / ledgerId / 見積番号 のいずれからでも引けるようにする
      [l.targetId, l.quoteNo].forEach(function(key) {
        key = String(key || '').trim();
        if (!key) return;
        if (!map[key]) map[key] = [];
        // 同一書類の重複登録を避ける
        if (map[key].some(function(x) { return x.id === d.id; })) return;
        map[key].push({
          id:         d.id,
          docType:    d.docType,
          supplier:   d.supplier,
          trader:     d.trader,
          subject:    d.subject,
          issueDate:  d.issueDate,
          effectDate: d.effectDate,
          fileName:   d.fileName,
          url:        d.url,
          diffRate:   d.diffRate,
        });
      });
    });
    return { success: true, map: map, docTypes: SDOC_TYPES };
  } catch (e) { return { success: false, error: e.message }; }
}

// ============================================================
// API: 値上げ時系列
//   payload: { supplier, itemName, dateFrom, dateTo, includeQuotes }
//   適用開始日（無ければ発行日）の昇順で並べ、
//   仕入先＋品名ごとの単価推移も併せて返す。
// ============================================================
function apiPriceIncreaseTimeline(payload) {
  try {
    payload = payload || {};
    var supplier = String(payload.supplier || '').trim();
    var itemKw   = String(payload.itemName || '').trim().toLowerCase();
    var from     = _sdocDateKey(payload.dateFrom);
    var to       = _sdocDateKey(payload.dateTo);

    var docs  = _sdocReadAll().filter(function(d) {
      return SDOC_PRICE_CHANGE_TYPES.indexOf(d.docType) >= 0;
    });

    var links = _sdocReadAllLinks();
    var qIdx  = _sdocBuildQuoteIndex();
    var linkMap = {};
    links.forEach(function(l) {
      if (!linkMap[l.docId]) linkMap[l.docId] = [];
      var q = qIdx.byId[l.targetId] || (l.quoteNo ? qIdx.byQuoteNo[l.quoteNo] : null);
      linkMap[l.docId].push({
        targetId: l.targetId,
        quoteNo:  l.quoteNo || (q ? q.quoteNo : ''),
        client:   q ? q.client : '',
      });
    });

    var events = docs.filter(function(d) {
      if (supplier && d.supplier !== supplier && d.trader !== supplier) return false;
      if (itemKw) {
        var hay = (d.itemName + ' ' + d.modelNo + ' ' + d.subject).toLowerCase();
        if (hay.indexOf(itemKw) < 0) return false;
      }
      var dk = _sdocDateKey(d.effectDate) || _sdocDateKey(d.issueDate);
      if (from && (!dk || dk < from)) return false;
      if (to   && (!dk || dk > to))   return false;
      return true;
    }).map(function(d) {
      var dk = _sdocDateKey(d.effectDate) || _sdocDateKey(d.issueDate);
      return {
        id:         d.id,
        dateKey:    dk,
        date:       d.effectDate || d.issueDate,
        issueDate:  d.issueDate,
        effectDate: d.effectDate,
        docType:    d.docType,
        supplier:   d.supplier || d.trader,
        trader:     d.trader,
        itemName:   d.itemName,
        modelNo:    d.modelNo,
        oldPrice:   d.oldPrice,
        newPrice:   d.newPrice,
        diff:       d.diff,
        diffRate:   d.diffRate,
        currency:   d.currency,
        subject:    d.subject,
        fileName:   d.fileName,
        url:        d.url,
        memo:       d.memo,
        quotes:     linkMap[d.id] || [],
      };
    });

    // 適用開始日の昇順（古い→新しい）＝時系列
    events.sort(function(a, b) {
      if (a.dateKey === b.dateKey) return String(a.id).localeCompare(String(b.id));
      if (!a.dateKey) return 1;
      if (!b.dateKey) return -1;
      return a.dateKey.localeCompare(b.dateKey);
    });

    // 仕入先＋品名ごとの単価推移（グラフ・比較用）
    var seriesMap = {};
    events.forEach(function(ev) {
      var key = (ev.supplier || '（仕入先未設定）') + ' / ' + (ev.itemName || ev.modelNo || '（品名未設定）');
      if (!seriesMap[key]) seriesMap[key] = { key: key, supplier: ev.supplier, itemName: ev.itemName || ev.modelNo, points: [] };
      seriesMap[key].points.push({
        date:     ev.date,
        dateKey:  ev.dateKey,
        oldPrice: ev.oldPrice,
        newPrice: ev.newPrice,
        diffRate: ev.diffRate,
        docId:    ev.id,
      });
    });

    var series = Object.keys(seriesMap).map(function(k) {
      var s      = seriesMap[k];
      var first  = s.points[0];
      var lastPt = s.points[s.points.length - 1];
      // 累計上昇率: 最初の旧単価 → 最後の新単価
      var base   = first.oldPrice || 0;
      s.firstPrice  = base;
      s.latestPrice = lastPt.newPrice || 0;
      s.totalRate   = base > 0 ? Math.round(((s.latestPrice - base) / base) * 1000) / 10 : 0;
      s.count       = s.points.length;
      return s;
    }).sort(function(a, b) { return b.totalRate - a.totalRate; });

    // 絞り込み用の候補
    var suppliers = {};
    _sdocReadAll().forEach(function(d) {
      if (SDOC_PRICE_CHANGE_TYPES.indexOf(d.docType) < 0) return;
      if (d.supplier) suppliers[d.supplier] = true;
      if (d.trader)   suppliers[d.trader]   = true;
    });

    return {
      success:   true,
      events:    events,
      series:    series,
      suppliers: Object.keys(suppliers).sort(),
    };
  } catch (e) { return { success: false, error: e.message }; }
}
