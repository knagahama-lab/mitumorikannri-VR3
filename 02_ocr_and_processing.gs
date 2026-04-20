// ============================================================
// 見積書・注文書管理システム
// ファイル 2/4: Gemini OCR・メール監視・データ転記
// ============================================================

// ============================================================
// メール監視（設定シート連動）
// ============================================================

function processNewEmails() {
  try {
    var configs = getEmailConfigs();
    if (configs.length === 0) {
      Logger.log('[EMAIL] 有効な設定なし。「メール監視設定」シートを確認してください。');
      return;
    }

    // 送信済みトレイ（見積書）
    var sentThreads = GmailApp.search('in:sent has:attachment filename:pdf', 0, 30);
    _processThreads(sentThreads, configs, false);

    // 受信トレイ（注文書）
    var inboxThreads = GmailApp.search('in:inbox has:attachment filename:pdf', 0, 30);
    _processThreads(inboxThreads, configs, true);

  } catch(e) {
    Logger.log('[processNewEmails ERROR] ' + e.message);
  }
}

function _processThreads(threads, configs, isInbox) {
  for (var i = 0; i < threads.length; i++) {
    var msgs = threads[i].getMessages();
    for (var j = 0; j < msgs.length; j++) {
      var msg   = msgs[j];
      var msgId = msg.getId();
      if (isMessageAlreadyProcessed(msgId)) continue;

      var fromAddr = msg.getFrom();
      var toAddr   = msg.getTo();
      var subject  = msg.getSubject();
      var atts     = msg.getAttachments();

      for (var k = 0; k < atts.length; k++) {
        var att  = atts[k];
        var name = att.getName();
        if (!name.toLowerCase().endsWith('.pdf')) continue;

        // 設定シートでマッチング
        var cfg = matchEmailConfig(name, subject, fromAddr, toAddr);
        if (!cfg) continue;

        try {
          if (cfg.docType === 'quote') {
            _processQuotePdf(att, msg, msgId);
          } else {
            _processOrderPdf(att, msg, msgId, cfg.orderType || '');
          }
        } catch(e) {
          Logger.log('[PROCESS ERROR] ' + name + ': ' + e.message);
        }
      }
    }
  }
}


// ============================================================
// 見積書処理
// ============================================================

function _processQuotePdf(attachment, gmailMsg, msgId) {
  var folderId  = CONFIG.QUOTE_FOLDER_ID;
  var folder    = DriveApp.getFolderById(folderId);
  var fileName  = nowJST().replace(/[\/: ]/g,'') + '_' + attachment.getName();
  var file      = folder.createFile(attachment.copyBlob().setName(fileName));
  var pdfUrl    = file.getUrl();
  var folderUrl = getFolderUrl(folderId);

  var ocr = extractPdfData(file, 'quote');
  if (!ocr) { Logger.log('[OCR SKIP] ' + fileName); return; }

  var mgmtId    = generateMgmtId();
  var ss        = getSpreadsheet();
  var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var newRow    = mgmtSheet.getLastRow() + 1;

  var row = new Array(27).fill('');
  row[MGMT_COLS.ID - 1]               = mgmtId;
  row[MGMT_COLS.QUOTE_NO - 1]         = ocr.documentNo    || '';
  row[MGMT_COLS.SUBJECT - 1]          = ocr.subject       || gmailMsg.getSubject();
  row[MGMT_COLS.CLIENT - 1]           = ocr.destCompany   || ocr.clientName || '';
  row[MGMT_COLS.STATUS - 1]           = CONFIG.STATUS.SENT;
  row[MGMT_COLS.QUOTE_DATE - 1]       = ocr.issueDate     || ocr.documentDate || '';
  row[MGMT_COLS.QUOTE_AMOUNT - 1]     = ocr.subtotal      || 0;
  row[MGMT_COLS.TAX - 1]              = ocr.tax           || 0;
  row[MGMT_COLS.TOTAL - 1]            = ocr.totalAmount   || 0;
  row[MGMT_COLS.QUOTE_PDF_URL - 1]    = pdfUrl;
  row[MGMT_COLS.DRIVE_FOLDER_URL - 1] = folderUrl;
  row[MGMT_COLS.LINKED - 1]           = 'FALSE';
  row[MGMT_COLS.CREATED_AT - 1]       = nowJST();
  row[MGMT_COLS.UPDATED_AT - 1]       = nowJST();
  row[MGMT_COLS.GMAIL_MSG_ID - 1]     = msgId;
  mgmtSheet.getRange(newRow, 1, 1, 27).setValues([row]);

  _writeQuoteLines(ss, mgmtSheet, newRow, mgmtId, ocr, pdfUrl, folderUrl);

  // 見積台帳の自動更新
  try {
    _apiLedgerUpdateUrl({
      quoteNo:   ocr.documentNo  || '',
      dest:      ocr.destCompany || ocr.clientName || '',
      subject:   ocr.subject     || (gmailMsg ? gmailMsg.getSubject() : ''),
      issueDate: ocr.issueDate   || ocr.documentDate || nowJST().substring(0, 10),
      saveUrl:   pdfUrl,
    });
  } catch(ledgerErr) {
    Logger.log('[LEDGER UPDATE SKIP] ' + ledgerErr.message);
  }

  // AI紐付け試行
  try {
    aiLinkQuoteToOrder(mgmtId);
  } catch(e) {
    Logger.log('[AI LINK ERROR] ' + e.message);
  }

  // チャット通知
  _sendChatNotification(mgmtId, 'quote');

  Logger.log('[QUOTE OK] ' + mgmtId);
}

function _writeQuoteLines(ss, mgmtSheet, mgmtRow, mgmtId, ocr, pdfUrl, folderUrl) {
  if (!ocr.lineItems || ocr.lineItems.length === 0) return;
  var qs = ss.getSheetByName(CONFIG.SHEET_QUOTES);
  var lines = ocr.lineItems.map(function(item, idx) {
    return [
      mgmtId,
      ocr.documentNo   || '',
      ocr.issueDate    || ocr.documentDate || '',
      ocr.destCompany  || '',
      ocr.destPerson   || '',
      idx + 1,
      item.itemName    || '',
      item.spec        || '',
      item.qty         || 0,
      item.unit        || '',
      item.unitPrice   || 0,
      item.amount      || 0,
      item.remarks     || '',
      pdfUrl,
      folderUrl,
    ];
  });
  var startRow = qs.getLastRow() + 1;
  qs.getRange(startRow, 1, lines.length, 15).setValues(lines);
  mgmtSheet.getRange(mgmtRow, MGMT_COLS.QUOTE_SHEET_ROW).setValue(startRow);
}


// ============================================================
// 注文書処理
// ============================================================

function _processOrderPdf(attachment, gmailMsg, msgId, orderType) {
  var folderId  = _getOrderFolderId(orderType);
  var folder    = DriveApp.getFolderById(folderId);
  var fileName  = nowJST().replace(/[\/: ]/g,'') + '_' + attachment.getName();
  var file      = folder.createFile(attachment.copyBlob().setName(fileName));
  var pdfUrl    = file.getUrl();
  var folderUrl = getFolderUrl(folderId);

  var ocr = extractPdfData(file, 'order');
  if (!ocr) { Logger.log('[OCR SKIP] ' + fileName); return; }

  var finalOrderType = orderType || ocr.orderType || '';
  _saveOrderData(ocr, finalOrderType, pdfUrl, folderUrl, msgId, gmailMsg.getSubject());
}

function _saveOrderData(ocr, orderType, pdfUrl, folderUrl, msgId, fallbackSubject) {
  var ss            = getSpreadsheet();
  var mgmtSheet     = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var action        = ocr.actionType || 'new';
  var linkedQuoteNo = ocr.linkedQuoteNo || '';
  var finalMgmtId, updateRow;

  var mgmtRow = -1;
  if (action === 'revision' || action === 'cancellation') {
    mgmtRow = _findExistingMgmtRowForOrder(ss, ocr.documentNo, ocr.subject || fallbackSubject);
  } else if (linkedQuoteNo) {
    mgmtRow = findMgmtRowByQuoteNo(linkedQuoteNo);
  }

  if (mgmtRow > 0) {
    finalMgmtId = mgmtSheet.getRange(mgmtRow, MGMT_COLS.ID).getValue();
    updateRow   = mgmtRow;

    if (action === 'cancellation') {
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.STATUS).setValue(CONFIG.STATUS.CANCELLED);
      var currentMemo = mgmtSheet.getRange(mgmtRow, MGMT_COLS.MEMO).getValue();
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.MEMO).setValue(currentMemo + "\n[CANCEL] " + (ocr.reason || 'キャンセル通知受領'));
    } else {
      var status = (action === 'revision') ? CONFIG.STATUS.REVISED : CONFIG.STATUS.RECEIVED;
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_NO).setValue(ocr.documentNo    || '');
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_DATE).setValue(ocr.documentDate || '');
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_AMOUNT).setValue(ocr.subtotal   || 0);

      var oldPdf = mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_PDF_URL).getValue();
      if (oldPdf && oldPdf !== pdfUrl) {
        var currentMemo = mgmtSheet.getRange(mgmtRow, MGMT_COLS.MEMO).getValue();
        mgmtSheet.getRange(mgmtRow, MGMT_COLS.MEMO).setValue(currentMemo + "\n[OLD PDF] " + oldPdf);
      }

      mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_PDF_URL).setValue(pdfUrl);
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.DRIVE_FOLDER_URL).setValue(folderUrl);
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.STATUS).setValue(status);
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.LINKED).setValue('TRUE');
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_TYPE).setValue(orderType);
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.MODEL_CODE).setValue(ocr.modelCode    || '');
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_SLIP_NO).setValue(ocr.orderSlipNo || '');
    }
    mgmtSheet.getRange(mgmtRow, MGMT_COLS.UPDATED_AT).setValue(nowJST());

  } else {
    finalMgmtId = generateMgmtId();
    var newRow  = mgmtSheet.getLastRow() + 1;
    updateRow   = newRow;
    var row = new Array(27).fill('');
    row[MGMT_COLS.ID - 1]               = finalMgmtId;
    row[MGMT_COLS.ORDER_NO - 1]         = ocr.documentNo    || '';
    row[MGMT_COLS.SUBJECT - 1]          = ocr.subject       || fallbackSubject;
    row[MGMT_COLS.CLIENT - 1]           = ocr.clientName    || '';
    row[MGMT_COLS.STATUS - 1]           = CONFIG.STATUS.RECEIVED;
    row[MGMT_COLS.ORDER_DATE - 1]       = ocr.documentDate  || '';
    row[MGMT_COLS.ORDER_AMOUNT - 1]     = ocr.subtotal      || 0;
    row[MGMT_COLS.ORDER_PDF_URL - 1]    = pdfUrl;
    row[MGMT_COLS.DRIVE_FOLDER_URL - 1] = folderUrl;
    row[MGMT_COLS.LINKED - 1]           = 'FALSE';
    row[MGMT_COLS.ORDER_TYPE - 1]       = orderType;
    row[MGMT_COLS.MODEL_CODE - 1]       = ocr.modelCode     || '';
    row[MGMT_COLS.ORDER_SLIP_NO - 1]    = ocr.orderSlipNo   || '';
    row[MGMT_COLS.CREATED_AT - 1]       = nowJST();
    row[MGMT_COLS.UPDATED_AT - 1]       = nowJST();
    row[MGMT_COLS.GMAIL_MSG_ID - 1]     = msgId || '';
    mgmtSheet.getRange(newRow, 1, 1, 27).setValues([row]);
  }

  if (action !== 'cancellation') {
    _writeOrderLines(ss, mgmtSheet, updateRow, finalMgmtId, ocr, orderType, pdfUrl, folderUrl);

    try {
      aiLinkOrderToQuote(finalMgmtId);
    } catch(e) {
      Logger.log('[AI LINK ERROR] ' + e.message);
    }
  }

  _sendChatNotification(finalMgmtId, 'order', action);

  return finalMgmtId;
}

function _findExistingMgmtRowForOrder(ss, orderNo, subject) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var data  = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (orderNo && data[i][MGMT_COLS.ORDER_NO - 1] === orderNo) return i + 1;
    if (subject && data[i][MGMT_COLS.SUBJECT - 1] === subject) return i + 1;
  }
  return -1;
}

function _writeOrderLines(ss, mgmtSheet, mgmtRow, mgmtId, ocr, orderType, pdfUrl, folderUrl) {
  if (!ocr.lineItems || ocr.lineItems.length === 0) return;
  var os = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  var lines = ocr.lineItems.map(function(item, idx) {
    return [
      mgmtId, ocr.documentNo||'', ocr.linkedQuoteNo||'', orderType,
      ocr.documentDate||'', ocr.modelCode||'', ocr.orderSlipNo||'',
      idx+1,
      item.itemName||'', item.spec||'',
      item.firstDelivery||'', item.deliveryDest||'',
      item.qty||0, item.unit||'', item.unitPrice||0, item.amount||0, item.remarks||'',
      pdfUrl, folderUrl,
    ];
  });
  var startRow = os.getLastRow() + 1;
  os.getRange(startRow, 1, lines.length, 19).setValues(lines);
  mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_SHEET_ROW).setValue(startRow);
}


// ============================================================
// Gemini OCR（Drive OCR + Gemini テキスト解析）
// ・Drive OCRでPDF→テキスト変換（無料・制限なし）
// ・GeminiにはテキストのみをJSON化（429激減）
// ・Drive OCR失敗時はBase64直接送信にフォールバック
// ============================================================

function extractPdfData(driveFile, docType) {
  // Step1: Drive APIでPDF→テキスト抽出（無料）
  var ocrText = _extractTextByDriveOcr(driveFile);

  if (!ocrText || ocrText.trim().length < 10) {
    Logger.log('[OCR] Drive OCR失敗またはテキスト不足 → Base64フォールバック');
    return _extractPdfDataByBase64(driveFile, docType);
  }

  Logger.log('[OCR TEXT] ' + ocrText.substring(0, 300));

  // Step2: テキストをGeminiでJSON化
  return _parseOcrTextWithGemini(ocrText, docType);
}

/**
 * Drive APIのOCR機能でPDFからテキストを抽出
 * 完全無料・レート制限なし・日本語対応
 */
function _extractTextByDriveOcr(driveFile) {
  try {
    var resource = {
      title: 'OCR_TEMP_' + Date.now(),
      mimeType: MimeType.GOOGLE_DOCS
    };
    var docFile = Drive.Files.insert(resource, driveFile.getBlob(), {
      ocr: true,
      ocrLanguage: 'ja',
      convert: true
    });

    var docId = docFile.id;
    Utilities.sleep(1000); // 変換完了待機

    var doc  = DocumentApp.openById(docId);
    var text = doc.getBody().getText();

    // 一時ファイル削除
    try { DriveApp.getFileById(docId).setTrashed(true); } catch(e) {}

    return text;
  } catch(e) {
    Logger.log('[DRIVE OCR ERROR] ' + e.message);
    return null;
  }
}

/**
 * OCRテキストをGeminiで構造化JSON化
 * テキストのみ送信なので429が大幅減
 */
function _parseOcrTextWithGemini(ocrText, docType) {
  var prompt = _buildOcrPrompt(docType) +
    '\n\n以下のOCRテキストを解析してJSONを返してください：\n\n' +
    ocrText.substring(0, 8000);

  var body = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.1 }
  };

  var result = _callGeminiApi(CONFIG.GEMINI_PRIMARY_MODEL, body);
  if (!result) result = _callGeminiApi(CONFIG.GEMINI_FALLBACK_MODEL, body);
  if (!result) { Logger.log('[GEMINI] 全モデル失敗'); return null; }

  try {
    var text = '';
    if (result.candidates && result.candidates[0] &&
        result.candidates[0].content && result.candidates[0].content.parts) {
      text = result.candidates[0].content.parts[0].text || '';
    }
    text = text.replace(/```json|```/g, '').trim();
    Logger.log('[OCR RAW] ' + text.substring(0, 500));
    return JSON.parse(text);
  } catch(e) {
    Logger.log('[OCR PARSE ERROR] ' + e.message);
    return null;
  }
}

/**
 * フォールバック：Base64でPDFを直接送信（旧方式）
 * Drive OCRが失敗した場合 or 有料APIに切り替えた場合に使用
 */
function _extractPdfDataByBase64(driveFile, docType) {
  var base64 = Utilities.base64Encode(driveFile.getBlob().getBytes());
  var body = {
    contents: [{ parts: [
      { text: _buildOcrPrompt(docType) },
      { inline_data: { mime_type: 'application/pdf', data: base64 } }
    ]}],
    generationConfig: { temperature: 0.1 }
  };

  var result = _callGeminiApi(CONFIG.GEMINI_PRIMARY_MODEL, body);
  if (!result) result = _callGeminiApi(CONFIG.GEMINI_FALLBACK_MODEL, body);
  if (!result) { Logger.log('[GEMINI] Base64フォールバックも失敗'); return null; }

  try {
    var text = '';
    if (result.candidates && result.candidates[0] &&
        result.candidates[0].content && result.candidates[0].content.parts) {
      text = result.candidates[0].content.parts[0].text || '';
    }
    text = text.replace(/```json|```/g, '').trim();
    Logger.log('[OCR RAW fallback] ' + text.substring(0, 500));
    return JSON.parse(text);
  } catch(e) {
    Logger.log('[OCR PARSE ERROR fallback] ' + e.message);
    return null;
  }
}

/**
 * Gemini API呼び出し（リトライ・nullガード付き）
 * 無料枠：429時は30秒待機×3回
 * 有料枠：429はほぼ発生しないが念のため残す
 */
function _callGeminiApi(model, body) {
  var key = CONFIG.GEMINI_API_KEY ||
    PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('GEMINI_API_KEY未設定');

  var url      = CONFIG.GEMINI_API_ENDPOINT + model + ':generateContent?key=' + key;
  var maxRetry = 3;
  var waitMs   = 30000;

  for (var attempt = 1; attempt <= maxRetry; attempt++) {
    var res;
    try {
      res = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(body),
        muteHttpExceptions: true,
      });
    } catch(fetchErr) {
      Logger.log('[GEMINI FETCH ERROR] ' + model + ' attempt' + attempt + ': ' + fetchErr.message);
      if (attempt < maxRetry) Utilities.sleep(waitMs);
      continue;
    }

    if (!res) {
      Logger.log('[GEMINI] ' + model + ' attempt' + attempt + ': レスポンスがnull');
      if (attempt < maxRetry) Utilities.sleep(waitMs);
      continue;
    }

    var code = res.getResponseCode();
    var text = res.getContentText();

    if (code === 429) {
      Logger.log('[GEMINI 429] ' + model + ' attempt' + attempt + ' → ' + waitMs/1000 + 's待機');
      if (attempt < maxRetry) Utilities.sleep(waitMs);
      continue;
    }

    if (code !== 200) {
      Logger.log('[GEMINI ERROR] ' + model + ': HTTP ' + code + ': ' + text.substring(0, 200));
      return null;
    }

    try {
      return JSON.parse(text);
    } catch(parseErr) {
      Logger.log('[GEMINI PARSE ERROR] ' + model + ': ' + parseErr.message);
      return null;
    }
  }

  Logger.log('[GEMINI] ' + model + ': 全リトライ失敗');
  return null;
}

function _buildOcrPrompt(docType) {
  if (docType === 'quote') {
    return 'あなたはOCR専門家です。添付PDF（見積書）を解析し、以下のJSON形式のみで返してください。説明文不要。\n' +
      '{\n' +
      '  "documentNo": "見積番号",\n' +
      '  "issueDate": "発行日(YYYY/MM/DD)",\n' +
      '  "documentDate": "見積日(YYYY/MM/DD)",\n' +
      '  "destCompany": "送り先・宛先の会社名",\n' +
      '  "destPerson": "送り先担当者名（なければ空文字）",\n' +
      '  "clientName": "顧客企業名",\n' +
      '  "subject": "件名",\n' +
      '  "subtotal": 小計(数値),\n' +
      '  "tax": 消費税(数値),\n' +
      '  "totalAmount": 合計(数値),\n' +
      '  "lineItems": [\n' +
      '    {"itemName":"品名","spec":"仕様","qty":数量,"unit":"単位","unitPrice":単価,"amount":金額,"remarks":"備考"}\n' +
      '  ]\n' +
      '}\n' +
      'ルール: 有効なJSONのみ。金額は数値。不明は空文字か0。合計行はlineItemsに含めない。';
  } else {
    return 'あなたはOCR専門家であり、業務フローの判定官です。添付PDF（発注書・注文書）を解析し、以下のJSON形式のみで返してください。\n' +
      '{\n' +
      '  "actionType": "new" または "revision" または "cancellation" (書類が新規ならnew、差し替え/更新ならrevision、取消通知ならcancellation),\n' +
      '  "reason": "差し替えやキャンセルの理由（あれば。なければ空文字）",\n' +
      '  "documentNo": "発注書番号",\n' +
      '  "documentDate": "発注日(YYYY/MM/DD)",\n' +
      '  "clientName": "発注先企業名",\n' +
      '  "subject": "件名",\n' +
      '  "modelCode": "機種コード",\n' +
      '  "orderSlipNo": "発注伝票番号",\n' +
      '  "linkedQuoteNo": "紐づく見積番号（なければ空文字）",\n' +
      '  "orderType": "試作 または 量産（不明なら空文字）",\n' +
      '  "subtotal": 小計(数値),\n' +
      '  "tax": 消費税(数値),\n' +
      '  "totalAmount": 合計(数値),\n' +
      '  "lineItems": [\n' +
      '    {"itemName":"品名","spec":"仕様","firstDelivery":"初回納入日(YYYY/MM/DD)","deliveryDest":"納入先","qty":数量,"unit":"単位","unitPrice":単価,"amount":金額,"remarks":"備考（取消線行はキャンセルと記載）"}\n' +
      '  ]\n' +
      '}\n' +
      '※重要: 書類内に「差し替え」「訂正」「版数更新」等の文言があればrevision、「中止」「取消」「キャンセル」等があればcancellationと判定してください。\n' +
      'ルール: 有効なJSONのみ。金額は数値。不明は空文字か0。合計行はlineItemsに含めない。';
  }
}


// ============================================================
// 手動アップロード（Webダッシュボードから）
// ============================================================

function processUploadedPdf(base64Data, fileName, docType, orderType) {
  try {
    var blob      = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', fileName);
    var folderId  = (docType === 'order') ? _getOrderFolderId(orderType||'') : CONFIG.WEB_UPLOAD_FOLDER_ID;
    var folder    = DriveApp.getFolderById(folderId);
    var saved     = 'MANUAL_' + nowJST().replace(/[\/: ]/g,'') + '_' + fileName;
    var file      = folder.createFile(blob.setName(saved));
    var pdfUrl    = file.getUrl();
    var folderUrl = getFolderUrl(folderId);

    Logger.log('[UPLOAD] 保存: ' + saved + ' → ' + folderUrl);

    var ocr = extractPdfData(file, docType);
    if (!ocr) return { success: false, error: 'OCR解析失敗。PDFを確認してください。' };

    var mockMsgId = 'MANUAL_' + Date.now();
    var finalMgmtId;
    if (docType === 'quote') {
      finalMgmtId = _processQuotePdfFromFile(pdfUrl, folderUrl, ocr, mockMsgId);
    } else {
      var finalType = orderType || ocr.orderType || '';
      finalMgmtId = _saveOrderData(ocr, finalType, pdfUrl, folderUrl, mockMsgId, fileName);
    }

    return {
      mgmtId:      finalMgmtId || ocr.mgmtId,
      documentNo:  ocr.documentNo,
      clientName:  ocr.destCompany || ocr.clientName,
      totalAmount: ocr.totalAmount,
      lineCount:   ocr.lineItems ? ocr.lineItems.length : 0,
      savedFolder: folderUrl,
      orderType:   orderType || ocr.orderType || '',
      modelCode:   ocr.modelCode   || '',
      orderSlipNo: ocr.orderSlipNo || '',
      issueDate:   ocr.issueDate   || '',
      destCompany: ocr.destCompany || '',
      destPerson:  ocr.destPerson  || '',
    };
  } catch(e) {
    Logger.log('[UPLOAD ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

function _processQuotePdfFromFile(pdfUrl, folderUrl, ocr, msgId) {
  var mgmtId    = generateMgmtId();
  var ss        = getSpreadsheet();
  var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var newRow    = mgmtSheet.getLastRow() + 1;

  var row = new Array(27).fill('');
  row[MGMT_COLS.ID - 1]               = mgmtId;
  row[MGMT_COLS.QUOTE_NO - 1]         = ocr.documentNo   || '';
  row[MGMT_COLS.SUBJECT - 1]          = ocr.subject      || '';
  row[MGMT_COLS.CLIENT - 1]           = ocr.destCompany  || ocr.clientName || '';
  row[MGMT_COLS.STATUS - 1]           = CONFIG.STATUS.SENT;
  row[MGMT_COLS.QUOTE_DATE - 1]       = ocr.issueDate    || ocr.documentDate || '';
  row[MGMT_COLS.QUOTE_AMOUNT - 1]     = ocr.subtotal     || 0;
  row[MGMT_COLS.TAX - 1]              = ocr.tax          || 0;
  row[MGMT_COLS.TOTAL - 1]            = ocr.totalAmount  || 0;
  row[MGMT_COLS.QUOTE_PDF_URL - 1]    = pdfUrl;
  row[MGMT_COLS.DRIVE_FOLDER_URL - 1] = folderUrl;
  row[MGMT_COLS.LINKED - 1]           = 'FALSE';
  row[MGMT_COLS.CREATED_AT - 1]       = nowJST();
  row[MGMT_COLS.UPDATED_AT - 1]       = nowJST();
  row[MGMT_COLS.GMAIL_MSG_ID - 1]     = msgId;
  mgmtSheet.getRange(newRow, 1, 1, 27).setValues([row]);

  _writeQuoteLines(ss, mgmtSheet, newRow, mgmtId, ocr, pdfUrl, folderUrl);

  try {
    aiLinkQuoteToOrder(mgmtId);
  } catch(e) {
    Logger.log('[AI LINK ERROR] ' + e.message);
  }

  _sendChatNotification(mgmtId, 'quote');

  return mgmtId;
}


// ============================================================
// 通知・設定・手動紐づけ（API用拡張）
// ============================================================

function _getChatWebhookUrl() {
  return CONFIG.GOOGLE_CHAT_WEBHOOK_URL ||
    PropertiesService.getScriptProperties().getProperty('GOOGLE_CHAT_WEBHOOK_URL') || '';
}
function testOcrDebug() {
  // テスト対象のファイルID（ログに出ていたID）
  var fileId = '1H6lqjWUK9rkUHnjnEW1AZWWrnLIZuy-N';
  var file = DriveApp.getFileById(fileId);
  
  Logger.log('ファイル名: ' + file.getName());
  
  // Step1: Drive OCRテスト
  var ocrText = _extractTextByDriveOcr(file);
  Logger.log('Drive OCR結果文字数: ' + (ocrText ? ocrText.length : 'null'));
  Logger.log('Drive OCR先頭200文字: ' + (ocrText ? ocrText.substring(0, 200) : 'null'));
  
  // Step2: Gemini疎通テスト
  var testBody = {
    contents: [{ parts: [{ text: 'テストです。「OK」とだけ返してください。' }] }],
    generationConfig: { temperature: 0.1 }
  };
  var result = _callGeminiApi(CONFIG.GEMINI_PRIMARY_MODEL, testBody);
  Logger.log('Gemini疎通: ' + JSON.stringify(result));
}