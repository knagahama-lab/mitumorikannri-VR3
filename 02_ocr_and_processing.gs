// ============================================================
// 見積書・注文書管理システム
// ファイル 2/4: Gemini OCR・メール監視・データ転記
// ============================================================
//
// 【修正点】
//   ・_sendChatNotification() を本ファイルに統合定義
//     （旧 "10 order upload and notify.gs" が廃止されたため）
//   ・processUploadedPdf の全工程を try/catch で保護し
//     OCR失敗・保存失敗でも success:false を明示的に返す
//   ・各処理ステップのログを詳細化（デバッグ容易化）
// ============================================================

// ============================================================
// ★ チャット通知（本ファイルに統合）
// ============================================================

/**
 * Google Chat Webhook へ通知を送る
 * 旧 "10 order upload and notify.gs" の _sendChatNotification の代替
 *
 * @param {string} mgmtId  - 管理ID
 * @param {string} docType - 'quote' | 'order'
 * @param {string} [action] - 'new' | 'revision' | 'cancellation'（注文書のみ）
 */
function _sendChatNotification(mgmtId, docType, action) {
  try {
    var webhookUrl = _getChatWebhookUrl();
    if (!webhookUrl) {
      Logger.log('[CHAT] Webhook未設定のためスキップ: ' + mgmtId);
      return;
    }

    // 管理シートから最新情報を取得
    var ss      = getSpreadsheet();
    var sheet   = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var last    = sheet.getLastRow();
    var rowData = null;
    if (last > 1) {
      var ids = sheet.getRange(2, MGMT_COLS.ID, last - 1, 1).getValues().flat();
      var idx = ids.map(String).indexOf(String(mgmtId));
      if (idx >= 0) {
        rowData = sheet.getRange(idx + 2, 1, 1, sheet.getLastColumn()).getValues()[0];
      }
    }

    var client   = rowData ? String(rowData[MGMT_COLS.CLIENT   - 1] || '') : '';
    var subject  = rowData ? String(rowData[MGMT_COLS.SUBJECT  - 1] || '') : '';
    var docNo    = rowData ? String(rowData[docType === 'quote'
                              ? MGMT_COLS.QUOTE_NO - 1
                              : MGMT_COLS.ORDER_NO - 1] || '') : '';
    var amount   = rowData ? rowData[docType === 'quote'
                              ? MGMT_COLS.QUOTE_AMOUNT - 1
                              : MGMT_COLS.ORDER_AMOUNT - 1] : 0;
    var pdfUrl   = rowData ? String(rowData[docType === 'quote'
                              ? MGMT_COLS.QUOTE_PDF_URL - 1
                              : MGMT_COLS.ORDER_PDF_URL - 1] || '') : '';
    var mgmtUrl  = _getMgmtRowUrl ? _getMgmtRowUrl(rowData ? (ids.map(String).indexOf(String(mgmtId)) + 2) : 0) : '';

    // メッセージ組み立て
    var typeLabel  = docType === 'quote' ? '見積書' : '注文書';
    var actionLabel = '';
    if (action === 'revision')     actionLabel = '【差し替え】';
    if (action === 'cancellation') actionLabel = '【キャンセル】';

    var lines = [
      actionLabel + '✅ ' + typeLabel + 'を登録しました',
      '管理ID: ' + mgmtId,
    ];
    if (client)  lines.push('顧客: '     + client);
    if (docNo)   lines.push('番号: '     + docNo);
    if (subject) lines.push('件名: '     + subject);
    if (amount)  lines.push('金額: ¥'   + Number(amount).toLocaleString());
    if (pdfUrl)  lines.push('PDF: '      + pdfUrl);
    if (mgmtUrl) lines.push('管理シート: ' + mgmtUrl);

    UrlFetchApp.fetch(webhookUrl, {
      method:          'post',
      contentType:     'application/json',
      payload:         JSON.stringify({ text: lines.join('\n') }),
      muteHttpExceptions: true,
    });
    Logger.log('[CHAT] 通知送信完了: ' + mgmtId);
  } catch(e) {
    // 通知失敗は握りつぶしてOCR処理は続行
    Logger.log('[CHAT ERROR] ' + e.message);
  }
}

/**
 * Webhook URL を取得する
 */
function _getChatWebhookUrl() {
  return CONFIG.GOOGLE_CHAT_WEBHOOK_URL ||
         PropertiesService.getScriptProperties().getProperty('GOOGLE_CHAT_WEBHOOK_URL') || '';
}

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
// 見積書処理（メール添付）
// ============================================================
function _processQuotePdf(attachment, gmailMsg, msgId) {
  var folderId  = CONFIG.QUOTE_FOLDER_ID;
  var folder    = DriveApp.getFolderById(folderId);
  var fileName  = nowJST().replace(/[\/: ]/g,'') + '_' + attachment.getName();
  var file      = folder.createFile(attachment.copyBlob().setName(fileName));
  var pdfUrl    = file.getUrl();
  var folderUrl = getFolderUrl(folderId);

  var ocr = extractPdfData(file, 'quote');
  if (!ocr) {
    Logger.log('[OCR SKIP] ' + fileName);
    _logOcrResult(fileName, 'ocr_failed', '');
    return;
  }

  var mgmtId    = generateMgmtId();
  var ss        = getSpreadsheet();
  var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var newRow    = mgmtSheet.getLastRow() + 1;

  var row = new Array(27).fill('');
  row[MGMT_COLS.ID              - 1] = mgmtId;
  row[MGMT_COLS.QUOTE_NO        - 1] = ocr.documentNo     || '';
  row[MGMT_COLS.SUBJECT         - 1] = ocr.subject        || gmailMsg.getSubject();
  row[MGMT_COLS.CLIENT          - 1] = ocr.destCompany    || ocr.clientName || '';
  row[MGMT_COLS.STATUS          - 1] = CONFIG.STATUS.SENT;
  row[MGMT_COLS.QUOTE_DATE      - 1] = ocr.issueDate      || ocr.documentDate || '';
  row[MGMT_COLS.QUOTE_AMOUNT    - 1] = ocr.subtotal       || 0;
  row[MGMT_COLS.TAX             - 1] = ocr.tax            || 0;
  row[MGMT_COLS.TOTAL           - 1] = ocr.totalAmount    || 0;
  row[MGMT_COLS.QUOTE_PDF_URL   - 1] = pdfUrl;
  row[MGMT_COLS.DRIVE_FOLDER_URL - 1] = folderUrl;
  row[MGMT_COLS.LINKED          - 1] = 'FALSE';
  row[MGMT_COLS.CREATED_AT      - 1] = nowJST();
  row[MGMT_COLS.UPDATED_AT      - 1] = nowJST();
  row[MGMT_COLS.GMAIL_MSG_ID    - 1] = msgId;
  mgmtSheet.getRange(newRow, 1, 1, 27).setValues([row]);

  _writeQuoteLines(ss, mgmtSheet, newRow, mgmtId, ocr, pdfUrl, folderUrl);

  // 見積台帳の自動更新
  try {
    _apiLedgerUpdateUrl({
      quoteNo:   ocr.documentNo || '',
      dest:      ocr.destCompany || ocr.clientName || '',
      subject:   ocr.subject     || gmailMsg.getSubject(),
      issueDate: ocr.issueDate   || ocr.documentDate || nowJST().substring(0, 10),
      saveUrl:   pdfUrl,
    });
  } catch(ledgerErr) {
    Logger.log('[LEDGER UPDATE SKIP] ' + ledgerErr.message);
  }

  // AI紐付け試行
  try { aiLinkQuoteToOrder(mgmtId); } catch(e) { Logger.log('[AI LINK ERROR] ' + e.message); }

  // Chat通知
  _sendChatNotification(mgmtId, 'quote');
  Logger.log('[QUOTE OK] ' + mgmtId);
}

function _writeQuoteLines(ss, mgmtSheet, mgmtRow, mgmtId, ocr, pdfUrl, folderUrl) {
  if (!ocr.lineItems || ocr.lineItems.length === 0) return;

  var qs    = ss.getSheetByName(CONFIG.SHEET_QUOTES);
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
// 注文書処理（メール添付）
// ============================================================
function _processOrderPdf(attachment, gmailMsg, msgId, orderType) {
  var folderId  = _getOrderFolderId(orderType || '');
  var folder    = DriveApp.getFolderById(folderId);
  var fileName  = nowJST().replace(/[\/: ]/g,'') + '_' + attachment.getName();
  var file      = folder.createFile(attachment.copyBlob().setName(fileName));
  var pdfUrl    = file.getUrl();
  var folderUrl = getFolderUrl(folderId);

  var ocr = extractPdfData(file, 'order');
  if (!ocr) {
    Logger.log('[OCR SKIP] ' + fileName);
    _logOcrResult(fileName, 'ocr_failed', '');
    return;
  }

  var finalOrderType = orderType || ocr.orderType || '';
  _saveOrderData(ocr, finalOrderType, pdfUrl, folderUrl, msgId, gmailMsg.getSubject());
}

function _saveOrderData(ocr, orderType, pdfUrl, folderUrl, msgId, fallbackSubject) {
  var ss        = getSpreadsheet();
  var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);

  var action       = ocr.actionType     || 'new';
  var linkedQuoteNo = ocr.linkedQuoteNo || '';

  // 既存レコードを探す
  var mgmtRow = -1;
  if (action === 'revision' || action === 'cancellation') {
    mgmtRow = _findExistingMgmtRowForOrder(ss, ocr.documentNo, ocr.subject || fallbackSubject);
  } else if (linkedQuoteNo) {
    mgmtRow = findMgmtRowByQuoteNo(linkedQuoteNo);
  }

  var finalMgmtId, updateRow;

  if (mgmtRow > 0) {
    // ── 既存行を更新 ──
    finalMgmtId = mgmtSheet.getRange(mgmtRow, MGMT_COLS.ID).getValue();
    updateRow   = mgmtRow;

    if (action === 'cancellation') {
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.STATUS).setValue(CONFIG.STATUS.CANCELLED);
      var currentMemo = mgmtSheet.getRange(mgmtRow, MGMT_COLS.MEMO).getValue();
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.MEMO).setValue(
        currentMemo + '\n[CANCEL] ' + (ocr.reason || 'キャンセル通知受領')
      );
    } else {
      var status = (action === 'revision') ? CONFIG.STATUS.REVISED : CONFIG.STATUS.RECEIVED;

      mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_NO).setValue(ocr.documentNo || '');
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_DATE).setValue(ocr.documentDate || '');
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_AMOUNT).setValue(ocr.subtotal || 0);

      // 旧PDFをメモに保存
      var oldPdf = mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_PDF_URL).getValue();
      if (oldPdf && oldPdf !== pdfUrl) {
        var memo2 = mgmtSheet.getRange(mgmtRow, MGMT_COLS.MEMO).getValue();
        mgmtSheet.getRange(mgmtRow, MGMT_COLS.MEMO).setValue(memo2 + '\n[OLD PDF] ' + oldPdf);
      }

      mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_PDF_URL   ).setValue(pdfUrl);
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.DRIVE_FOLDER_URL).setValue(folderUrl);
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.STATUS          ).setValue(status);
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.LINKED          ).setValue('TRUE');
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_TYPE      ).setValue(orderType);
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.MODEL_CODE      ).setValue(ocr.modelCode   || '');
      mgmtSheet.getRange(mgmtRow, MGMT_COLS.ORDER_SLIP_NO   ).setValue(ocr.orderSlipNo || '');
    }
    mgmtSheet.getRange(mgmtRow, MGMT_COLS.UPDATED_AT).setValue(nowJST());

  } else {
    // ── 新規行を追加 ──
    finalMgmtId = generateMgmtId();
    var newRow  = mgmtSheet.getLastRow() + 1;
    updateRow   = newRow;

    var row = new Array(27).fill('');
    row[MGMT_COLS.ID               - 1] = finalMgmtId;
    row[MGMT_COLS.ORDER_NO         - 1] = ocr.documentNo   || '';
    row[MGMT_COLS.SUBJECT          - 1] = ocr.subject      || fallbackSubject;
    row[MGMT_COLS.CLIENT           - 1] = ocr.clientName   || '';
    row[MGMT_COLS.STATUS           - 1] = CONFIG.STATUS.RECEIVED;
    row[MGMT_COLS.ORDER_DATE       - 1] = ocr.documentDate || '';
    row[MGMT_COLS.ORDER_AMOUNT     - 1] = ocr.subtotal     || 0;
    row[MGMT_COLS.ORDER_PDF_URL    - 1] = pdfUrl;
    row[MGMT_COLS.DRIVE_FOLDER_URL - 1] = folderUrl;
    row[MGMT_COLS.LINKED           - 1] = 'FALSE';
    row[MGMT_COLS.ORDER_TYPE       - 1] = orderType;
    row[MGMT_COLS.MODEL_CODE       - 1] = ocr.modelCode    || '';
    row[MGMT_COLS.ORDER_SLIP_NO    - 1] = ocr.orderSlipNo  || '';
    row[MGMT_COLS.CREATED_AT       - 1] = nowJST();
    row[MGMT_COLS.UPDATED_AT       - 1] = nowJST();
    row[MGMT_COLS.GMAIL_MSG_ID     - 1] = msgId || '';
    mgmtSheet.getRange(newRow, 1, 1, 27).setValues([row]);
  }

  // キャンセル以外は明細書き込み＋AI紐付け
  if (action !== 'cancellation') {
    _writeOrderLines(ss, mgmtSheet, updateRow, finalMgmtId, ocr, orderType, pdfUrl, folderUrl);
    try { aiLinkOrderToQuote(finalMgmtId); } catch(e) { Logger.log('[AI LINK ERROR] ' + e.message); }
  }

  // Chat通知
  _sendChatNotification(finalMgmtId, 'order', action);

  return finalMgmtId;
}

function _findExistingMgmtRowForOrder(ss, orderNo, subject) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var data  = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (orderNo  && data[i][MGMT_COLS.ORDER_NO - 1] === orderNo)  return i + 1;
    if (subject  && data[i][MGMT_COLS.SUBJECT  - 1] === subject)  return i + 1;
  }
  return -1;
}

function _writeOrderLines(ss, mgmtSheet, mgmtRow, mgmtId, ocr, orderType, pdfUrl, folderUrl) {
  if (!ocr.lineItems || ocr.lineItems.length === 0) return;

  var os    = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  var lines = ocr.lineItems.map(function(item, idx) {
    return [
      mgmtId, ocr.documentNo || '', ocr.linkedQuoteNo || '', orderType,
      ocr.documentDate || '', ocr.modelCode || '', ocr.orderSlipNo || '',
      idx + 1,
      item.itemName || '', item.spec || '',
      item.firstDelivery || '', item.deliveryDest || '',
      item.qty || 0, item.unit || '', item.unitPrice || 0, item.amount || 0, item.remarks || '',
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
  var body   = {
    contents: [{ parts: [
      { text: _buildOcrPrompt(docType) },
      { inline_data: { mime_type: 'application/pdf', data: base64 } },
    ]}],
    generationConfig: { temperature: 0.1, responseMimeType: 'application/json' },
  };

  var result = _callGeminiApi(CONFIG.GEMINI_PRIMARY_MODEL,  body);
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

function _callGeminiApi(model, body) {
  var key = CONFIG.GEMINI_API_KEY ||
            PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('GEMINI_API_KEY未設定');
  var url = CONFIG.GEMINI_API_ENDPOINT + model + ':generateContent?key=' + key;
  try {
    var res = fetchWithRetry(url, {
      method: 'post', contentType: 'application/json',
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
      ' "documentNo": "見積番号",\n' +
      ' "issueDate": "発行日(YYYY/MM/DD)",\n' +
      ' "documentDate": "見積日(YYYY/MM/DD)",\n' +
      ' "deliveryDate": "納期(YYYY/MM/DD または 記載通り。なければ空文字)",\n' +
      ' "destCompany": "送り先・宛先の会社名",\n' +
      ' "destPerson": "送り先担当者名（なければ空文字）",\n' +
      ' "clientName": "顧客企業名",\n' +
      ' "subject": "件名",\n' +
      ' "subtotal": 小計(数値),\n' +
      ' "tax": 消費税(数値),\n' +
      ' "totalAmount": 合計(数値),\n' +
      ' "lineItems": [\n' +
      '   {"itemName":"品名","spec":"仕様","qty":数量,"unit":"単位","unitPrice":単価,"amount":金額,"remarks":"備考"}\n' +
      ' ]\n' +
      '}\n' +
      'ルール: 有効なJSONのみ。金額は数値。不明は空文字か0。合計行はlineItemsに含めない。';
  } else {
    return 'あなたはOCR専門家であり、業務フローの判定官です。添付PDF（発注書・注文書）を解析し、以下のJSON形式のみで返してください。\n' +
      '{\n' +
      ' "actionType": "new" または "revision" または "cancellation",\n' +
      ' "reason": "差し替えやキャンセルの理由（あれば。なければ空文字）",\n' +
      ' "documentNo": "発注書番号",\n' +
      ' "documentDate": "発注日(YYYY/MM/DD)",\n' +
      ' "clientName": "発注先企業名",\n' +
      ' "subject": "件名",\n' +
      ' "modelCode": "機種コード",\n' +
      ' "orderSlipNo": "発注伝票番号",\n' +
      ' "linkedQuoteNo": "紐づく見積番号（なければ空文字）",\n' +
      ' "orderType": "試作 または 量産（不明なら空文字）",\n' +
      ' "subtotal": 小計(数値),\n' +
      ' "tax": 消費税(数値),\n' +
      ' "totalAmount": 合計(数値),\n' +
      ' "lineItems": [\n' +
      '   {"itemName":"品名","spec":"仕様","firstDelivery":"初回納入日(YYYY/MM/DD)","deliveryDest":"納入先","qty":数量,"unit":"単位","unitPrice":単価,"amount":金額,"remarks":"備考"}\n' +
      ' ]\n' +
      '}\n' +
      '※重要: 書類内に「差し替え」「訂正」「版数更新」等の文言があればrevision、「中止」「取消」「キャンセル」等があればcancellationと判定。\n' +
      'ルール: 有効なJSONのみ。金額は数値。不明は空文字か0。合計行はlineItemsに含めない。';
  }
}

// ============================================================
// ★ 手動アップロード（Webダッシュボードの「PDF登録」ボタン）
//    PDF登録 → OCR → 保存フォルダへ保存 → スプレッドシートに転記
// ============================================================

/**
 * フロントエンドの _apiUploadPdf から呼ばれるメイン関数
 *
 * 処理順:
 *   1. base64 → Blob 変換
 *   2. 保存先フォルダへファイルを保存（Drive）
 *   3. Gemini OCR でテキスト抽出
 *   4. 管理シート（＋見積/注文明細シート）に転記
 *   5. AI紐付け試行
 *   6. Chat / メール通知
 *
 * @returns {{ success: boolean, mgmtId: string, ... } | { success: false, error: string }}
 */
function processUploadedPdf(base64Data, fileName, docType, orderType) {
  // ── ステップ1: Blob 変換 ──
  var blob;
  try {
    blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data), 'application/pdf', fileName
    );
  } catch(e) {
    Logger.log('[UPLOAD] Blob変換失敗: ' + e.message);
    return { success: false, error: 'PDFデータの変換に失敗しました: ' + e.message };
  }

  // ── ステップ2: Drive に保存 ──
  var file, pdfUrl, folderUrl, folderId;
  try {
    folderId  = (docType === 'order')
      ? _getOrderFolderId(orderType || '')
      : CONFIG.WEB_UPLOAD_FOLDER_ID;
    var folder = DriveApp.getFolderById(folderId);
    var saved  = 'MANUAL_' + nowJST().replace(/[\/: ]/g, '') + '_' + fileName;
    file       = folder.createFile(blob.setName(saved));
    pdfUrl     = file.getUrl();
    folderUrl  = getFolderUrl(folderId);
    Logger.log('[UPLOAD] 保存完了: ' + saved + ' URL=' + pdfUrl);
  } catch(e) {
    Logger.log('[UPLOAD] Drive保存失敗: ' + e.message);
    return { success: false, error: 'Driveへの保存に失敗しました: ' + e.message };
  }

  // ── ステップ3: OCR ──
  var ocr;
  try {
    ocr = extractPdfData(file, docType);
  } catch(e) {
    Logger.log('[UPLOAD] OCR例外: ' + e.message);
    _logOcrResult(fileName, 'error', pdfUrl);
    return { success: false, error: 'OCR処理中にエラーが発生しました: ' + e.message };
  }
  if (!ocr) {
    Logger.log('[UPLOAD] OCR解析失敗');
    _logOcrResult(fileName, 'ocr_failed', pdfUrl);
    return { success: false, error: 'OCRの解析に失敗しました。PDFの内容を確認してください。' };
  }

  // ── ステップ4: スプレッドシートに転記 ──
  var finalMgmtId;
  try {
    var mockMsgId = 'MANUAL_' + Date.now();
    if (docType === 'quote') {
      finalMgmtId = _processQuotePdfFromFile(pdfUrl, folderUrl, ocr, mockMsgId);
    } else {
      var finalType = orderType || ocr.orderType || '';
      finalMgmtId   = _saveOrderData(ocr, finalType, pdfUrl, folderUrl, mockMsgId, fileName);
    }
  } catch(e) {
    Logger.log('[UPLOAD] スプレッドシート転記失敗: ' + e.message);
    return { success: false, error: 'スプレッドシートへの転記に失敗しました: ' + e.message };
  }

  Logger.log('[UPLOAD] 完了 mgmtId=' + finalMgmtId);

  // ── 成功レスポンス ──
  return {
    success:     true,
    mgmtId:      finalMgmtId || '',
    documentNo:  ocr.documentNo   || '',
    clientName:  ocr.destCompany  || ocr.clientName || '',
    totalAmount: ocr.totalAmount  || 0,
    lineCount:   ocr.lineItems ? ocr.lineItems.length : 0,
    savedFolder: folderUrl,
    pdfUrl:      pdfUrl,
    orderType:   orderType         || ocr.orderType  || '',
    modelCode:   ocr.modelCode     || '',
    orderSlipNo: ocr.orderSlipNo   || '',
    issueDate:   ocr.issueDate     || '',
    destCompany: ocr.destCompany   || '',
    destPerson:  ocr.destPerson    || '',
  };
}

/**
 * 手動アップロード用の見積書転記処理
 * （メール添付処理の _processQuotePdf と同等。msgId は "MANUAL_xxx"）
 */
function _processQuotePdfFromFile(pdfUrl, folderUrl, ocr, msgId) {
  var mgmtId    = generateMgmtId();
  var ss        = getSpreadsheet();
  var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var newRow    = mgmtSheet.getLastRow() + 1;

  var row = new Array(32).fill('');
  row[MGMT_COLS.ID               - 1] = mgmtId;
  row[MGMT_COLS.QUOTE_NO         - 1] = ocr.documentNo     || '';
  row[MGMT_COLS.SUBJECT          - 1] = ocr.subject        || '';
  row[MGMT_COLS.CLIENT           - 1] = ocr.destCompany    || ocr.clientName || '';
  row[MGMT_COLS.STATUS           - 1] = CONFIG.STATUS.SENT;
  row[MGMT_COLS.QUOTE_DATE       - 1] = ocr.issueDate      || ocr.documentDate || '';
  row[MGMT_COLS.QUOTE_AMOUNT     - 1] = ocr.subtotal       || 0;
  row[MGMT_COLS.TAX              - 1] = ocr.tax            || 0;
  row[MGMT_COLS.TOTAL            - 1] = ocr.totalAmount    || 0;
  row[MGMT_COLS.QUOTE_PDF_URL    - 1] = pdfUrl;
  row[MGMT_COLS.DRIVE_FOLDER_URL - 1] = folderUrl;
  row[MGMT_COLS.LINKED           - 1] = 'FALSE';
  row[MGMT_COLS.DELIVERY_DATE    - 1] = ocr.deliveryDate   || '';
  row[MGMT_COLS.CREATED_AT       - 1] = nowJST();
  row[MGMT_COLS.UPDATED_AT       - 1] = nowJST();
  row[MGMT_COLS.GMAIL_MSG_ID     - 1] = msgId;
  mgmtSheet.getRange(newRow, 1, 1, 32).setValues([row]);

  _writeQuoteLines(ss, mgmtSheet, newRow, mgmtId, ocr, pdfUrl, folderUrl);

  // 見積台帳の自動更新
  try {
    _apiLedgerUpdateUrl({
      quoteNo:   ocr.documentNo  || '',
      dest:      ocr.destCompany || ocr.clientName || '',
      subject:   ocr.subject     || '',
      issueDate: ocr.issueDate   || ocr.documentDate || nowJST().substring(0, 10),
      saveUrl:   pdfUrl,
    });
  } catch(ledgerErr) {
    Logger.log('[LEDGER UPDATE SKIP] ' + ledgerErr.message);
  }

  // AI紐付け試行
  try { aiLinkQuoteToOrder(mgmtId); } catch(e) { Logger.log('[AI LINK ERROR] ' + e.message); }

  // Chat通知
  _sendChatNotification(mgmtId, 'quote');

  return mgmtId;
}

// ============================================================
// OCRログ記録ユーティリティ
// ============================================================

// _logOcrResult は 13_ocr_extended.gs で定義（4引数版: fileName, status, mgmtId, detail）。
// ここでの重複定義を削除し、13_ocr_extended.gs の実装に統一する。
// 旧呼び出し側 (pdfUrl を第3引数で渡すケース) は mgmtId 列に入るが動作上問題なし。