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

        // ★設定シートでマッチング
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

  // ===== 見積台帳の自動更新 =====
  // OCRで取得した宛先・件名で台帳の対象行を特定し、保存先URLとステータスを更新
  try {
    _apiLedgerUpdateUrl({
      quoteNo:   ocr.documentNo  || '',
      dest:      ocr.destCompany || ocr.clientName || '',
      subject:   ocr.subject     || (gmailMsg ? gmailMsg.getSubject() : ''),
      issueDate: ocr.issueDate   || ocr.documentDate || nowJST().substring(0, 10),
      saveUrl:   pdfUrl,
    });
    Logger.log('[LEDGER UPDATE] 台帳自動更新試行完了');
  } catch(ledgerErr) {
    Logger.log('[LEDGER UPDATE SKIP] ' + ledgerErr.message);
  }

  // ★チャット通知の送信
  _sendChatNotification({
    id: mgmtId, 
    subject: row[MGMT_COLS.SUBJECT - 1], 
    client: row[MGMT_COLS.CLIENT - 1], 
    quoteAmount: row[MGMT_COLS.QUOTE_AMOUNT - 1]
  }, 'quote');

  Logger.log('[QUOTE OK] ' + mgmtId);
}

function _writeQuoteLines(ss, mgmtSheet, mgmtRow, mgmtId, ocr, pdfUrl, folderUrl) {
  if (!ocr.lineItems || ocr.lineItems.length === 0) return;
  var qs = ss.getSheetByName(CONFIG.SHEET_QUOTES);
  var lines = ocr.lineItems.map(function(item, idx) {
    return [
      mgmtId,
      ocr.documentNo   || '',
      ocr.issueDate    || ocr.documentDate || '',  // ★発行日
      ocr.destCompany  || '',                      // ★送り先会社名
      ocr.destPerson   || '',                      // ★送り先担当者名
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
  var linkedQuoteNo = ocr.linkedQuoteNo || '';
  var mgmtRow       = linkedQuoteNo ? findMgmtRowByQuoteNo(linkedQuoteNo) : -1;
  var ss            = getSpreadsheet();
  var mgmtSheet     = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var finalMgmtId, updateRow;

  if (mgmtRow > 0) {
    finalMgmtId = mgmtSheet.getRange(mgmtRow, MGMT_COLS.ID).getValue();
    updateRow   = mgmtRow;
    mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_NO).setValue(ocr.documentNo    || '');
    mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_DATE).setValue(ocr.documentDate || '');
    mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_AMOUNT).setValue(ocr.subtotal   || 0);
    mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_PDF_URL).setValue(pdfUrl);
    mgmtSheet.getRange(mgmtRow, MGMT_COLS.DRIVE_FOLDER_URL).setValue(folderUrl);
    mgmtSheet.getRange(mgmtRow, MGMT_COLS.STATUS).setValue(CONFIG.STATUS.RECEIVED);
    mgmtSheet.getRange(mgmtRow, MGMT_COLS.LINKED).setValue('TRUE');
    mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_TYPE).setValue(orderType);
    mgmtSheet.getRange(mgmtRow, MGMT_COLS.MODEL_CODE).setValue(ocr.modelCode    || '');
    mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_SLIP_NO).setValue(ocr.orderSlipNo || '');
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
    row[MGMT_COLS.STATUS - 1]           = CONFIG.STATUS.PLANNED;  // ★紐づけ完了後にORDERED/RECEIVEDに変更
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

  _writeOrderLines(ss, mgmtSheet, updateRow, finalMgmtId, ocr, orderType, pdfUrl, folderUrl);
  return finalMgmtId;

  // ★チャット通知の送信（AI紐づけを含む）
  _sendChatNotification({
    id: finalMgmtId, 
    subject: ocr.subject || fallbackSubject, 
    client: ocr.clientName, 
    orderAmount: ocr.subtotal
  }, 'order');
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
// Gemini OCR
// ============================================================

function extractPdfData(driveFile, docType) {
  var base64 = Utilities.base64Encode(driveFile.getBlob().getBytes());
  var body = {
    contents: [{ parts: [
      { text: _buildOcrPrompt(docType) },
      { inline_data: { mime_type: 'application/pdf', data: base64 } }
    ]}],
    generationConfig: { temperature: 0.1, responseMimeType: 'application/json' },
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
    text = text.replace(/```json|```/g,'').trim();
    Logger.log('[OCR RAW] ' + text.substring(0,500));
    return JSON.parse(text);
  } catch(e) {
    Logger.log('[OCR PARSE ERROR] ' + e.message);
    return null;
  }
}

function _callGeminiApi(model, body) {
  var key = CONFIG.GEMINI_API_KEY ||
    PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('GEMINI_API_KEY未設定');
  var url = CONFIG.GEMINI_API_ENDPOINT + model + ':generateContent?key=' + key;
  try {
    var res = fetchWithRetry(url, {
      method:'post', contentType:'application/json',
      payload: JSON.stringify(body), muteHttpExceptions: true,
    });
    return JSON.parse(res.getContentText());
  } catch(e) {
    Logger.log('[GEMINI ERROR] ' + model + ': ' + e.message);
    return null;
  }
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
    return 'あなたはOCR専門家です。添付PDF（発注書）を解析し、以下のJSON形式のみで返してください。説明文不要。\n' +
      '{\n' +
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
      'ルール: 有効なJSONのみ。金額は数値。不明は空文字か0。合計行はlineItemsに含めない。取消線行もlineItemsに含めremarksにキャンセルと記載。';
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
  return mgmtId;

  // ★チャット通知の送信
  _sendChatNotification({
    id: mgmtId, 
    subject: row[MGMT_COLS.SUBJECT - 1], 
    client: row[MGMT_COLS.CLIENT - 1], 
    quoteAmount: row[MGMT_COLS.QUOTE_AMOUNT - 1]
  }, 'quote');
}

// ============================================================
// 通知・設定・手動紐づけ（API用拡張）
// ============================================================

/**
 * Google Chatへの通知を送信する（見積情報付き）
 */
function _sendChatNotification(mgmtObj, docType) {
  var webhookUrl = _getChatWebhookUrl();
  if (!webhookUrl) return;

  var title = docType === 'quote' ? "📄 見積書を登録しました" : "📦 注文書を受領しました";
  
  var text = "【" + title + "】\n";
  text += "案件名: " + (mgmtObj.subject || "なし") + "\n";
  text += "顧客名: " + (mgmtObj.client || "不明") + "\n";
  text += "金額: ¥" + (Number(mgmtObj.orderAmount || mgmtObj.quoteAmount || 0).toLocaleString()) + "\n";

  // 注文書の場合、紐づいた見積情報を探す（アグレッシブに）
  if (docType === 'order') {
    var matchRes = matchOrderToQuote(mgmtObj.id, true); // とりあえず一番近いものに強制紐づけ
    if (matchRes.autoLinked && matchRes.match) {
      var q = matchRes.match;
      text += "\n🔗 *紐づけ済み見積情報 (AI自動)*\n";
      text += "・見積No: " + (q.quoteNo || "不明") + "\n";
      text += "・件名: " + (q.subject || "不明") + "\n";
      text += "・金額: ¥" + (Number(q.amount || 0).toLocaleString()) + "\n";
      text += "・見積PDF: " + (q.quotePdfUrl || "URLなし") + "\n";
    } else {
      text += "\n⚠️ 該当する見積書が見つかりませんでした\n";
    }
  }

  text += "\n🌐 システムで確認:\n" + ScriptApp.getService().getUrl();

  var payload = { "text": text };
  try {
    UrlFetchApp.fetch(webhookUrl, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch(e) {
    Logger.log('[CHAT NOTIFY ERROR] ' + e.message);
  }
}

function _getChatWebhookUrl() {
  return CONFIG.GOOGLE_CHAT_WEBHOOK_URL || PropertiesService.getScriptProperties().getProperty('GOOGLE_CHAT_WEBHOOK_URL') || '';
}