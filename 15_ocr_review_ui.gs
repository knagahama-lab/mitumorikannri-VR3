// ============================================================
// 見積書・注文書管理システム
// ファイル 15: OCR確認UI バックエンドAPI
// ============================================================
//
// 【フロー】
//   1. PDFアップロード → OCR実行 → 結果をセッションに保持（未DB登録）
//   2. フロントで確認・修正
//   3. 「承認して登録」で初めてDBに書き込む
//
// 【03_webapp_ap.gs handleApiRequest に追記するcase】
//   case 'ocrPreview':    res = apiOcrPreview(payload);    break;
//   case 'ocrApprove':    res = apiOcrApprove(payload);    break;
//   case 'ocrGetPending': res = apiOcrGetPending(payload); break;
//   case 'ocrDiscard':    res = apiOcrDiscard(payload);    break;
// ============================================================

var OCR_PENDING_KEY_PREFIX = 'OCR_PENDING_';
var OCR_PENDING_TTL_MS     = 30 * 60 * 1000; // 30分で自動破棄

// ============================================================
// API: OCRプレビュー（DBには書かない）
// ============================================================
function apiOcrPreview(p) {
  try {
    if (!p.base64Data || !p.fileName || !p.docType) {
      return { success: false, error: '必須パラメータ不足' };
    }

    // DriveにPDFを一時保存
    var blob     = Utilities.newBlob(Utilities.base64Decode(p.base64Data), 'application/pdf', p.fileName);
    var folderId = (p.docType === 'order')
      ? _getOrderFolderId(p.orderType || '')
      : CONFIG.WEB_UPLOAD_FOLDER_ID;
    var folder   = DriveApp.getFolderById(folderId);
    var saved    = 'OCR_PREVIEW_' + nowJST().replace(/[\\/: ]/g,'') + '_' + p.fileName;
    var file     = folder.createFile(blob.setName(saved));
    var pdfUrl   = file.getUrl();

    // OCR実行（13_ocr_enhanced.gsの強化版が使われる）
    var ocr = extractPdfData(file, p.docType);
    if (!ocr) {
      file.setTrashed(true); // 失敗したら一時ファイルを削除
      _logOcrResult(p.fileName, 'ocr_failed', null, 'OCRプレビュー失敗');
      return { success: false, error: 'OCR解析に失敗しました。PDFの内容を確認してください。' };
    }

    // 品質スコア計算
    var quality = _calcOcrQuality(ocr);

    // セッションIDを生成してプロパティに保存（30分有効）
    var sessionId = 'sess_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') +
                    '_' + Math.random().toString(36).substring(2, 7);

    var pending = {
      sessionId  : sessionId,
      docType    : p.docType,
      orderType  : p.orderType || ocr.orderType || '',
      pdfUrl     : pdfUrl,
      fileId     : file.getId(),
      folderId   : folderId,
      fileName   : saved,
      ocrResult  : ocr,
      quality    : quality,
      createdAt  : new Date().getTime(),
    };

    PropertiesService.getScriptProperties()
      .setProperty(OCR_PENDING_KEY_PREFIX + sessionId, JSON.stringify(pending));

    _logOcrResult(p.fileName, 'preview', null,
      '品質スコア:' + quality + ' 明細:' + (ocr.lineItems ? ocr.lineItems.length : 0) + '行');

    return {
      success   : true,
      sessionId : sessionId,
      ocrResult : ocr,
      quality   : quality,
      pdfUrl    : pdfUrl,
      warnings  : _buildOcrWarnings(ocr, quality),
    };
  } catch(e) {
    Logger.log('[OCR PREVIEW ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// API: 確認後に承認してDBに登録
// ============================================================
function apiOcrApprove(p) {
  try {
    if (!p.sessionId) return { success: false, error: 'sessionIdが必要です' };

    // セッションから保留データを取得
    var key = OCR_PENDING_KEY_PREFIX + p.sessionId;
    var raw = PropertiesService.getScriptProperties().getProperty(key);
    if (!raw) return { success: false, error: 'セッションが見つかりません（タイムアウトした可能性があります）' };

    var pending = JSON.parse(raw);

    // TTLチェック（30分）
    if (new Date().getTime() - pending.createdAt > OCR_PENDING_TTL_MS) {
      PropertiesService.getScriptProperties().deleteProperty(key);
      return { success: false, error: 'セッションの有効期限が切れました。再度PDFをアップロードしてください。' };
    }

    // フロントで修正されたOCR結果でオーバーライド
    var ocr = p.correctedOcr || pending.ocrResult;

    // 修正されたlineItemsの数値を再正規化
    if (ocr.lineItems) {
      ocr.lineItems = ocr.lineItems.map(function(item) {
        item.unitPrice = _ocr_normAmt(item.unitPrice);
        item.amount    = _ocr_normAmt(item.amount);
        item.qty       = _ocr_normQty(item.qty);
        if (!item.amount && item.unitPrice > 0 && item.qty > 0) {
          item.amount = item.unitPrice * item.qty;
        }
        return item;
      });
    }

    var mockMsgId = 'OCR_APPROVED_' + Date.now();
    var finalMgmtId;

    if (pending.docType === 'quote') {
      finalMgmtId = _processQuotePdfFromFile(pending.pdfUrl, getFolderUrl(pending.folderId), ocr, mockMsgId);
    } else {
      var finalType = pending.orderType || ocr.orderType || '';
      finalMgmtId = _saveOrderData(ocr, finalType, pending.pdfUrl, getFolderUrl(pending.folderId), mockMsgId, pending.fileName);
    }

    // セッション削除
    PropertiesService.getScriptProperties().deleteProperty(key);

    _logOcrResult(pending.fileName, 'approved', finalMgmtId,
      '承認登録完了 明細:' + (ocr.lineItems ? ocr.lineItems.length : 0) + '行');

    return {
      success     : true,
      mgmtId      : finalMgmtId,
      documentNo  : ocr.documentNo,
      clientName  : ocr.destCompany || ocr.clientName,
      totalAmount : ocr.totalAmount,
      lineCount   : ocr.lineItems ? ocr.lineItems.length : 0,
    };
  } catch(e) {
    Logger.log('[OCR APPROVE ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// API: 破棄（一時ファイルを削除してセッションを消す）
// ============================================================
function apiOcrDiscard(p) {
  try {
    if (!p.sessionId) return { success: false, error: 'sessionIdが必要です' };
    var key = OCR_PENDING_KEY_PREFIX + p.sessionId;
    var raw = PropertiesService.getScriptProperties().getProperty(key);
    if (raw) {
      try {
        var pending = JSON.parse(raw);
        DriveApp.getFileById(pending.fileId).setTrashed(true);
      } catch(e) {}
      PropertiesService.getScriptProperties().deleteProperty(key);
    }
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// OCR警告メッセージ生成（フロントに表示する）
// ============================================================
function _buildOcrWarnings(ocr, quality) {
  var warnings = [];

  if (quality < 50) {
    warnings.push({ level: 'error', msg: 'OCR精度が低い可能性があります。全項目を確認してください。' });
  } else if (quality < 75) {
    warnings.push({ level: 'warn', msg: '一部の項目が正確に読み取れていない可能性があります。' });
  }

  if (!ocr.documentNo) warnings.push({ level: 'warn', msg: '見積番号・注文番号が読み取れませんでした。手動入力してください。' });
  if (!ocr.totalAmount || ocr.totalAmount === 0) warnings.push({ level: 'warn', msg: '合計金額が読み取れませんでした。' });
  if (!ocr.lineItems || ocr.lineItems.length === 0) warnings.push({ level: 'error', msg: '明細行が読み取れませんでした。PDFに明細が含まれているか確認してください。' });

  var zeroPrice = ocr.lineItems ? ocr.lineItems.filter(function(i){ return !i.unitPrice; }).length : 0;
  if (zeroPrice > 0) warnings.push({ level: 'warn', msg: zeroPrice + '行の単価が読み取れませんでした。確認してください。' });

  return warnings;
}

// ============================================================
// 03_webapp_ap.gs へ追記するcase（コメントとして記載）
// ============================================================
//
// handleApiRequest の switch文 default の直前に追記:
//
//   case 'ocrPreview':    res = apiOcrPreview(payload);    break;
//   case 'ocrApprove':    res = apiOcrApprove(payload);    break;
//   case 'ocrDiscard':    res = apiOcrDiscard(payload);    break;
