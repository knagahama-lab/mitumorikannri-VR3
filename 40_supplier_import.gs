// ============================================================
// 40_supplier_import.gs
// 仕入先見積書（他社見積書PDF）の一括インポート＋OCR自動登録
//
//   apiSdocImportPdf(p)        … ブラウザから受け取ったPDF1件を保存→OCR→「仕入先書類」へ登録（連続インポート用）
//   apiSdocImportFolder(p)     … Driveフォルダ内の未登録PDFを分割実行でOCR登録（残り件数を返す）
//   apiSdocWatchGet/Set(p)     … 監視フォルダの設定（1時間ごとに自動取込：sdocAutoImportTrigger）
//
//   主な仕入先：実装会社（岐阜電子工業）、組立工場（イードリーム）。
//   仕入先名は OCR の発行元 → 既知の仕入先名（SDOC_KNOWN_SUPPLIERS）→ フォルダ名 の順で判定。
//   明細は全行を「明細(JSON)」列（21列目）に保存し、品名・型番は検索用に連結、備考に明細一覧を記載。
//   同じDriveファイル／同じファイル名の書類は重複として登録しない。
// ============================================================

var SDOC_LINES_COL = 21; // 仕入先書類シート：明細(JSON)
var SDOC_WATCH_PROP = 'SDOC_WATCH_FOLDER';

function _sdocKnownSuppliers() {
  var raw = PropertiesService.getScriptProperties().getProperty('SDOC_KNOWN_SUPPLIERS');
  return (raw ? raw.split(',') : ['岐阜電子工業', 'イードリーム']).map(function(s) { return s.trim(); }).filter(Boolean);
}
function _sdocMatchKnown(text) {
  var norm = function(x) { return (typeof _cpmCompanyNorm === 'function') ? _cpmCompanyNorm(x) : String(x || '').replace(/株式会社|\s/g, ''); };
  var t = norm(text);
  var hit = _sdocKnownSuppliers().filter(function(k) {
    var n = norm(k);
    return n && t.indexOf(n) >= 0;
  })[0];
  return hit || '';
}

function _sdocOcrPrompt() {
  var self = (PropertiesService.getScriptProperties().getProperty('SELF_COMPANY_NAMES') || 'サン電子').split(',')[0];
  return 'あなたはOCR専門家です。添付PDFは取引先（仕入先・外注先・商社）から「' + self + '」宛に届いた書類です。以下のJSON形式のみで返してください。説明文不要。\n' +
    '{\n' +
    ' "docKind": "見積書 / 値上げ通知書 / 価格改定通知 / 単価表 / その他 のいずれか",\n' +
    ' "issuerName": "発行元の会社名（社名ロゴ・住所・印がある側。宛先の「' + self + '」ではない）",\n' +
    ' "traderName": "商社名（発行元とは別に商社が介在する場合のみ。なければ空文字）",\n' +
    ' "documentNo": "見積番号・書類番号",\n' +
    ' "issueDate": "発行日(YYYY/MM/DD)",\n' +
    ' "effectDate": "価格の適用開始日(YYYY/MM/DD)。なければ空文字",\n' +
    ' "subject": "件名（機種・基板名などが分かれば含める）",\n' +
    ' "currency": "JPY など",\n' +
    ' "lineItems": [\n' +
    '   {"itemName":"品名","modelNo":"型番・品番・図番","qty":数量,"unit":"単位","unitPrice":単価,"oldUnitPrice":旧単価(値上げ書面のみ。なければ0),"amount":金額,"remarks":"備考"}\n' +
    ' ],\n' +
    ' "subtotal": 小計(数値), "tax": 消費税(数値), "total": 合計(数値)\n' +
    '}\n' +
    'ルール: 有効なJSONのみ。金額・数量は数値。不明は空文字か0。小計・合計・消費税の行は lineItems に含めない。';
}

/** PDF Blob を OCR（Gemini）して JSON を返す。失敗時は null */
function _sdocOcrBlob(blob) {
  var body = {
    contents: [{ parts: [{ text: _sdocOcrPrompt() }, { inline_data: { mime_type: 'application/pdf', data: Utilities.base64Encode(blob.getBytes()) } }] }],
    generationConfig: { temperature: 0.1, responseMimeType: 'application/json' },
  };
  var models = [CONFIG.GEMINI_PRIMARY_MODEL, CONFIG.GEMINI_FALLBACK_MODEL];
  for (var i = 0; i < models.length; i++) {
    var res = _callGeminiApi(models[i], body);
    try {
      var text = res && res.candidates && res.candidates[0] && res.candidates[0].content.parts[0].text;
      if (!text) continue;
      return JSON.parse(String(text).replace(/```json|```/g, '').trim());
    } catch (e) { Logger.log('[_sdocOcrBlob] ' + e.message); }
  }
  return null;
}

/** 登録済みの Drive ファイルID・ファイル名（重複判定用） */
function _sdocExistingKeys() {
  var ids = {}, names = {};
  _sdocReadAll().forEach(function(d) {
    var m = String(d.url || '').match(/[-\w]{25,}/);
    if (m) ids[m[0]] = d.id;
    if (d.fileName) names[String(d.fileName).trim()] = d.id;
  });
  return { ids: ids, names: names };
}

/**
 * OCR結果 → 仕入先書類 1件として登録
 * @param {Object} ocr   _sdocOcrBlob の結果（null 可：その場合はファイル名等だけで登録）
 * @param {Object} meta  { fileName, url, supplierHint, docTypeHint, folderName }
 */
function _sdocRegisterFromOcr(ocr, meta) {
  ocr = ocr || {};
  var lines = (ocr.lineItems || []).filter(function(l) { return l && (l.itemName || l.modelNo || Number(l.unitPrice)); });
  // 仕入先：手動指定 → OCR発行元（既知名に寄せる） → フォルダ名・ファイル名から既知名
  var supplier = String(meta.supplierHint || '').trim()
    || _sdocMatchKnown(ocr.issuerName) || String(ocr.issuerName || '').trim()
    || _sdocMatchKnown(meta.folderName + ' ' + meta.fileName) || '（仕入先不明）';
  var trader = String(ocr.traderName || '').trim();
  var kind = String(ocr.docKind || '');
  var docType = meta.docTypeHint && meta.docTypeHint !== 'auto' ? meta.docTypeHint
    : /値上げ/.test(kind) ? '値上げ通知書' : /価格改定/.test(kind) ? '価格改定通知' : /単価表/.test(kind) ? '単価表'
    : /見積/.test(kind) || !kind ? (trader ? '商社見積書' : '仕入先見積書') : 'その他';
  var yen = function(n) { return Number(n || 0).toLocaleString(); };
  var first = lines[0] || {};
  var names = lines.map(function(l) { return String(l.itemName || '').trim(); }).filter(String);
  var models = lines.map(function(l) { return String(l.modelNo || '').trim(); }).filter(String);
  var memo = [];
  if (ocr.documentNo) memo.push('書類番号: ' + ocr.documentNo);
  if (lines.length) {
    memo.push('【明細 ' + lines.length + '行】');
    lines.slice(0, 40).forEach(function(l, i) {
      memo.push((i + 1) + '. ' + [l.itemName, l.modelNo].filter(Boolean).join(' / ') + '　' + (l.qty || '') + (l.unit || '') +
        '　@' + yen(l.unitPrice) + (Number(l.oldUnitPrice) ? '（旧 @' + yen(l.oldUnitPrice) + '）' : '') + (l.amount ? '　= ' + yen(l.amount) : ''));
    });
  }
  if (ocr.total || ocr.subtotal) memo.push('小計 ' + yen(ocr.subtotal) + ' / 合計 ' + yen(ocr.total));
  memo.push(ocr && ocr.lineItems ? '（OCR自動登録）' : '（OCR失敗：ファイルのみ登録。内容を確認してください）');

  var single = lines.length === 1;
  var res = apiSupplierDocSave({
    docType: docType, supplier: supplier, trader: trader,
    subject: String(ocr.subject || '').trim() || String(meta.fileName || '').replace(/\.pdf$/i, ''),
    issueDate: String(ocr.issueDate || '').replace(/-/g, '/'), effectDate: String(ocr.effectDate || '').replace(/-/g, '/'),
    itemName: (single ? names[0] : names.slice(0, 8).join('・') + (names.length > 8 ? ' ほか' + (names.length - 8) + '件' : '')) || '',
    modelNo: models.slice(0, 8).join('・'),
    qty: single ? first.qty : '', oldPrice: single ? (first.oldUnitPrice || '') : '', newPrice: single ? first.unitPrice : (ocr.subtotal || ocr.total || ''),
    currency: ocr.currency || 'JPY', fileName: meta.fileName, url: meta.url, memo: memo.join('\n'),
  });
  if (!res.success) return res;
  try { // 明細(JSON) を21列目へ
    var sheet = _initSupplierDocSheet();
    if (String(sheet.getRange(1, SDOC_LINES_COL).getValue()) !== '明細(JSON)') sheet.getRange(1, SDOC_LINES_COL).setValue('明細(JSON)').setFontWeight('bold');
    var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(function(r) { return String(r[0]); });
    var idx = ids.lastIndexOf(res.id);
    if (idx >= 0) sheet.getRange(idx + 2, SDOC_LINES_COL).setValue(JSON.stringify({ documentNo: ocr.documentNo || '', lines: lines, subtotal: ocr.subtotal || 0, tax: ocr.tax || 0, total: ocr.total || 0 }).substring(0, 49000));
  } catch (e) { Logger.log('[sdoc lines] ' + e.message); }
  return { success: true, id: res.id, supplier: supplier, docType: docType, subject: String(ocr.subject || ''), issueDate: ocr.issueDate || '',
           documentNo: ocr.documentNo || '', lines: lines.length, total: ocr.total || ocr.subtotal || 0, ocrOk: !!(ocr && ocr.lineItems) };
}

// ============================================================
// ① 連続インポート（ブラウザから1件ずつ送る）
// ============================================================
function apiSdocImportPdf(p) {
  try {
    p = p || {};
    if (!p.base64Data || !p.fileName) return { success: false, error: 'ファイルデータ不足' };
    var fileName = String(p.fileName).trim();
    var dupId = p.force ? '' : _sdocExistingKeys().names[fileName];
    if (dupId) return { success: true, skipped: true, reason: '同じファイル名の書類が登録済み', id: dupId };
    var folderId = PropertiesService.getScriptProperties().getProperty('SDOC_UPLOAD_FOLDER_ID') || CONFIG.WEB_UPLOAD_FOLDER_ID;
    var safe = fileName.replace(/[/\\:*?"<>|]/g, '_');
    var blob = Utilities.newBlob(Utilities.base64Decode(p.base64Data), p.mimeType || 'application/pdf', '仕入先書類_' + safe);
    var file = DriveApp.getFolderById(folderId).createFile(blob);
    var ocr = p.skipOcr ? null : _sdocOcrBlob(file.getBlob());
    return _sdocRegisterFromOcr(ocr, { fileName: fileName, url: file.getUrl(), supplierHint: p.supplier, docTypeHint: p.docType, folderName: '' });
  } catch (e) {
    Logger.log('[apiSdocImportPdf] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

// ============================================================
// ② Driveフォルダインポート（未登録PDFだけ、分割実行）
// ============================================================
function _sdocFolderPdfs(folder, recursive, out, pathName) {
  var it = folder.getFilesByType(MimeType.PDF);
  while (it.hasNext()) { var f = it.next(); out.push({ file: f, folderName: pathName }); }
  if (recursive) {
    var sub = folder.getFolders();
    while (sub.hasNext()) { var s = sub.next(); _sdocFolderPdfs(s, true, out, pathName + ' ' + s.getName()); }
  }
  return out;
}

function apiSdocImportFolder(p) {
  try {
    p = p || {};
    var fid = (String(p.folderUrl || '').match(/[-\w]{25,}/) || [])[0];
    if (!fid) return { success: false, error: 'DriveフォルダのURLを入力してください' };
    var folder = DriveApp.getFolderById(fid);
    var started = Date.now();
    var max = Math.min(Number(p.max) || 8, 20);
    var keys = _sdocExistingKeys();
    var all = _sdocFolderPdfs(folder, !!p.recursive, [], folder.getName());
    var pending = all.filter(function(x) { return !keys.ids[x.file.getId()] && !keys.names[x.file.getName()]; });
    if (p.dryRun) return { success: true, folderName: folder.getName(), total: all.length, pending: pending.length, registered: all.length - pending.length };
    var results = [];
    for (var i = 0; i < pending.length && results.length < max; i++) {
      if (Date.now() - started > 240000) break; // 実行時間（6分）の上限に配慮
      var x = pending[i];
      var r;
      try {
        var ocr = _sdocOcrBlob(x.file.getBlob());
        r = _sdocRegisterFromOcr(ocr, { fileName: x.file.getName(), url: x.file.getUrl(), supplierHint: p.supplier, docTypeHint: p.docType, folderName: x.folderName });
      } catch (e) { r = { success: false, error: e.message }; }
      r.fileName = x.file.getName();
      results.push(r);
    }
    return { success: true, folderName: folder.getName(), total: all.length, processed: results.length,
             remaining: Math.max(0, pending.length - results.length), results: results };
  } catch (e) {
    Logger.log('[apiSdocImportFolder] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message.indexOf('No item with the given ID') >= 0 ? 'フォルダが見つからないか、アクセス権がありません' : e.message };
  }
}

// ============================================================
// 監視フォルダ（1時間ごとに自動でOCR登録）
// ============================================================
function apiSdocWatchGet() {
  try {
    var w = JSON.parse(PropertiesService.getScriptProperties().getProperty(SDOC_WATCH_PROP) || '{}');
    var on = ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === 'sdocAutoImportTrigger'; });
    return { success: true, folderUrl: w.folderUrl || '', recursive: !!w.recursive, supplier: w.supplier || '', enabled: on, lastRun: w.lastRun || '', lastResult: w.lastResult || '' };
  } catch (e) { return { success: false, error: e.message }; }
}

function apiSdocWatchSet(p) {
  try {
    p = p || {};
    var props = PropertiesService.getScriptProperties();
    var w = JSON.parse(props.getProperty(SDOC_WATCH_PROP) || '{}');
    if (p.enabled) {
      var fid = (String(p.folderUrl || '').match(/[-\w]{25,}/) || [])[0];
      if (!fid) return { success: false, error: 'DriveフォルダのURLを入力してください' };
      DriveApp.getFolderById(fid).getName(); // アクセス確認
      w.folderUrl = p.folderUrl; w.recursive = !!p.recursive; w.supplier = p.supplier || '';
    }
    props.setProperty(SDOC_WATCH_PROP, JSON.stringify(w));
    ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() === 'sdocAutoImportTrigger') ScriptApp.deleteTrigger(t); });
    if (p.enabled) ScriptApp.newTrigger('sdocAutoImportTrigger').timeBased().everyHours(1).create();
    return apiSdocWatchGet();
  } catch (e) { return { success: false, error: e.message }; }
}

/** 時間主導トリガー：監視フォルダの新しいPDFを取り込む */
function sdocAutoImportTrigger() {
  var props = PropertiesService.getScriptProperties();
  var w = JSON.parse(props.getProperty(SDOC_WATCH_PROP) || '{}');
  if (!w.folderUrl) return;
  var r = apiSdocImportFolder({ folderUrl: w.folderUrl, recursive: w.recursive, supplier: w.supplier, max: 15 });
  w.lastRun = nowJST();
  w.lastResult = r.success ? ('登録 ' + (r.results || []).filter(function(x) { return x.success && !x.skipped; }).length + '件／残り ' + r.remaining + '件') : ('エラー: ' + r.error);
  props.setProperty(SDOC_WATCH_PROP, JSON.stringify(w));
  // 通知ログ（🔔）にも記録
  try {
    var n = (r.results || []).filter(function(x) { return x.success && !x.skipped; });
    if (n.length && typeof _ntLog === 'function') {
      _ntLog('仕入先見積', { body: '📥 仕入先書類を自動登録しました（' + n.length + '件）\n' + n.map(function(x) { return '・' + x.supplier + '　' + (x.subject || x.fileName); }).join('\n') });
    }
  } catch (e) {}
}
