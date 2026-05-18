// ============================================================
// 見積書・注文書管理システム
// ファイル 13: OCR 拡張パッケージ（統合版）
// ============================================================
//
// 【統合元】
//   13_ocr_enhanced.gs   — OCR強化パッチ（プロンプト・正規化・ログ）
//   17_ocr_hybrid.gs     — OCRハイブリッド戦略（テキスト抽出+Vision+AIリレー）
//   14_ocr_retry_batch.gs — 夜間バッチOCRリトライ
//
// 【優先順位】
//   extractPdfData()     → 17版（ハイブリッド方式）が最終実装
//   processUploadedPdf() → 17版がハイブリッドOCR対応
//   _buildOcrPrompt()    → 13版（高精度プロンプト）
//   _ocr_normalize() 等  → 13版のユーティリティ群
// ============================================================

var OCR_LOG_SHEET_NAME = 'OCR処理ログ';

// ============================================================
// OCRハイブリッド戦略（17版 — 最終 extractPdfData）
// ============================================================
function extractPdfData(driveFile, docType) {
  Logger.log('[OCR HYBRID] 開始: ' + docType + ' / ' + driveFile.getName());

  var text = _extractTextFromPdf(driveFile);

  if (text && text.trim().length >= 30) {
    Logger.log('[OCR HYBRID] テキスト抽出成功: ' + text.length + '文字 → Gemini構造化');
    var result = _geminiStructureText(text, docType);
    if (result && (result.documentNo || result.totalAmount)) {
      _logOcrResult(driveFile.getName(), 'success', null,
        (docType==='quote'?'見積書':'注文書') + 'テキスト+Gemini構造化: ' + (result.documentNo||'?'));
      return result;
    }
    Logger.log('[OCR HYBRID] テキスト構造化が不十分 → Visionへフォールバック');
  }

  Logger.log('[OCR HYBRID] Gemini Vision(画像解析)試行');
  var apiKey = getGeminiApiKey();
  if (!apiKey) {
    _logOcrResult(driveFile.getName(), 'ocr_failed', null, 'APIキー未設定。手動入力が必要です。');
    return null;
  }

  var visionResult = _geminiVisionOcr(driveFile, docType, apiKey);
  if (visionResult) {
    _logOcrResult(driveFile.getName(), 'success', null,
      (docType==='quote'?'見積書':'注文書') + 'Vision OCR: ' + (visionResult.documentNo||'?'));
    return visionResult;
  }

  _logOcrResult(driveFile.getName(), 'ocr_failed', null, '全OCR手法失敗。手動入力してください。');
  return null;
}

// ============================================================
// processUploadedPdf（17版 — ハイブリッドOCR対応）
// ============================================================
// ============================================================
// processUploadedPdf（指定フォルダへの保存専用に修正）
// ============================================================
function processUploadedPdf(base64Data, fileName, docType, orderType) {
  try {
    var blob      = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', fileName);
    
    // ★ 保存先をユーザー指定のフォルダIDに固定
    var folderId  = '1Y66PDSi35ScuIyS0Jgm0l3p2l7MEM2Jk'; 
    var folder    = DriveApp.getFolderById(folderId);
    
    // ファイル名のプレフィックスを付けて保存
    var saved     = 'MANUAL_' + nowJST().replace(/[\\/: ]/g,'') + '_' + fileName;
    var file      = folder.createFile(blob.setName(saved));
    var pdfUrl    = file.getUrl();
    var folderUrl = folder.getUrl(); 

    Logger.log('[UPLOAD ONLY] 指定フォルダへの保存に成功しました: ' + saved);

    // 即時OCRはスキップするため、UI（画面側）がクラッシュしない用の空テンプレートを定義
    var ocr = {
      documentNo: '', issueDate: '', documentDate: '',
      destCompany: '', destPerson: '', clientName: '',
      subject: '', subtotal: 0, tax: 0, totalAmount: 0,
      lineItems: [],
      actionType: docType === 'order' ? 'new' : undefined,
    };

    // 画面側に返すレスポンス（保存成功のステータスのみを返却）
    return {
      success     : true,
      mgmtId      : null,
      documentNo  : '',
      clientName  : '',
      totalAmount : 0,
      lineCount   : 0,
      savedFolder : folderUrl,
      orderType   : orderType || '',
      modelCode   : '',
      orderSlipNo : '',
      issueDate   : '',
      destCompany : '',
      destPerson  : '',
      ocrResult   : ocr,
      quality     : 0,
      warnings    : [{ level: 'info', msg: 'ファイルは指定フォルダに保存されました。自動OCRの実行をお待ちください。' }],
      pdfUrl      : pdfUrl,
    };
    
  } catch(e) {
    Logger.log('[UPLOAD ONLY ERROR] ' + e.message);
    _logOcrResult(fileName, 'error', null, e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// 見積書: PDFテキスト + ルールベース構造化
// ============================================================
function _extractQuoteFromPdfText(driveFile) {
  var text = _extractTextFromPdf(driveFile);
  if (!text || text.trim().length < 20) {
    Logger.log('[OCR HYBRID] 見積書テキスト抽出失敗 → 手動入力UIへ');
    _logOcrResult(driveFile.getName(), 'ocr_failed', null, '見積書PDFからテキストを抽出できませんでした。');
    return null;
  }

  Logger.log('[OCR HYBRID] 見積書テキスト抽出成功: ' + text.length + '文字');
  var result = _parseQuoteText(text);
  if (!result.documentNo && !result.totalAmount) {
    Logger.log('[OCR HYBRID] 見積書構造化不十分 → Geminiテキスト補完');
    result = _geminiStructureText(text, 'quote') || result;
  }

  _logOcrResult(driveFile.getName(), 'success', null,
    '見積書: ' + (result.documentNo||'番号不明') +
    ' 金額:' + (result.totalAmount||0) +
    ' 明細:' + (result.lineItems ? result.lineItems.length : 0) + '行');
  return result;
}

// ============================================================
// 注文書: テキスト抽出 → Gemini構造化 → Vision フォールバック
// ============================================================
function _extractOrderHybrid(driveFile) {
  var text = _extractTextFromPdf(driveFile);
  if (text && text.trim().length >= 30) {
    Logger.log('[OCR HYBRID] 注文書テキスト抽出成功 → Gemini構造化');
    var result = _geminiStructureText(text, 'order');
    if (result) {
      _logOcrResult(driveFile.getName(), 'success', null, '注文書テキスト+Gemini構造化: ' + (result.documentNo||'?'));
      return result;
    }
  }

  Logger.log('[OCR HYBRID] 注文書テキスト抽出失敗 → Gemini Vision試行');
  var apiKey = getGeminiApiKey();
  if (!apiKey) {
    _logOcrResult(driveFile.getName(), 'ocr_failed', null, 'スキャンPDF + APIキー未設定。手動入力が必要です。');
    return null;
  }

  var visionResult = _geminiVisionOcr(driveFile, 'order', apiKey);
  if (visionResult) {
    _logOcrResult(driveFile.getName(), 'success', null, '注文書Vision OCR: ' + (visionResult.documentNo||'?'));
    return visionResult;
  }

  _logOcrResult(driveFile.getName(), 'ocr_failed', null, '全OCR手法失敗。手動入力してください。');
  return null;
}

// ============================================================
// PDFからテキストを直接抽出（GASネイティブ）
// ============================================================
function _extractTextFromPdf(driveFile) {
  try {
    var blob = driveFile.getBlob();
    var metadata = {
      name    : '_ocr_tmp_' + Date.now(),
      mimeType: 'application/vnd.google-apps.document',
    };
    var insertedFile = Drive.Files.create(metadata, blob, { ocrLanguage: 'ja', fields: 'id' });
    Utilities.sleep(2000);

    var docFile = DriveApp.getFileById(insertedFile.id);
    var doc     = DocumentApp.openById(insertedFile.id);
    var text    = doc.getBody().getText();
    docFile.setTrashed(true);

    if (text && text.trim().length > 10) {
      Logger.log('[TEXT EXTRACT] Drive OCR成功: ' + text.length + '文字');
      return text;
    }
    Logger.log('[TEXT EXTRACT] Drive OCRテキストなし（スキャンPDFの可能性）');
    return '';
  } catch(e) {
    Logger.log('[TEXT EXTRACT ERROR] ' + e.message);
    return '';
  }
}

// ============================================================
// Geminiにテキストを渡して構造化（AIリレー方式）
// ============================================================
function _geminiStructureText(text, docType) {
  var apiKey = getGeminiApiKey();
  if (!apiKey) return null;

  var modelsToTry = [
    'gemini-3.1-flash-lite-preview',
    'gemini-2.0-flash-lite-001',
    'gemini-3-flash-preview',
    'gemini-2.5-flash',
  ];

  var prompt = _buildTextStructurePrompt(docType, text);
  var body   = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.0, responseMimeType: 'application/json', maxOutputTokens: 2048 },
  };
  var options = {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(body), muteHttpExceptions: true,
  };

  for (var i = 0; i < modelsToTry.length; i++) {
    var model = modelsToTry[i];
    Logger.log('[GEMINI TEXT] ' + model + ' で試行中...');
    var url = CONFIG.GEMINI_API_ENDPOINT + model + ':generateContent?key=' + apiKey;
    try {
      var res = UrlFetchApp.fetch(url, options);
      if (res.getResponseCode() === 200) {
        var json = JSON.parse(res.getContentText());
        var raw  = json.candidates && json.candidates[0] &&
                   json.candidates[0].content && json.candidates[0].content.parts
                   ? json.candidates[0].content.parts[0].text || '' : '';
        raw = raw.replace(/```json|```/gi, '').trim();
        var parsed = _ocr_safeParseJson(raw);
        return parsed ? _ocr_normalize(parsed, docType) : null;
      } else {
        Logger.log('[GEMINI TEXT] ' + model + ' 失敗: HTTP ' + res.getResponseCode());
      }
    } catch(e) {
      Logger.log('[GEMINI TEXT ERROR] ' + model + ' 例外: ' + e.message);
    }
  }
  return null;
}

// ============================================================
// Gemini Vision OCR（スキャンPDF用・AIリレー方式）
// ============================================================
function _geminiVisionOcr(driveFile, docType, apiKey) {
  try {
    var base64 = Utilities.base64Encode(driveFile.getBlob().getBytes());
    var body   = {
      contents: [{ parts: [
        { text: _buildOcrPrompt(docType) },
        { inline_data: { mime_type: 'application/pdf', data: base64 } },
      ]}],
      generationConfig: { temperature: 0.05, responseMimeType: 'application/json', maxOutputTokens: 2048 },
    };
    var options = {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(body), muteHttpExceptions: true,
    };

    var modelsToTry = [
      'gemini-3.1-flash-lite-preview',
      'gemini-2.0-flash-lite-001',
      'gemini-3-flash-preview',
      'gemini-2.5-flash',
    ];

    for (var i = 0; i < modelsToTry.length; i++) {
      var model = modelsToTry[i];
      Logger.log('[VISION] ' + model + ' で試行中...');
      var url = CONFIG.GEMINI_API_ENDPOINT + model + ':generateContent?key=' + apiKey;
      var res = UrlFetchApp.fetch(url, options);
      if (res.getResponseCode() === 200) {
        var json = JSON.parse(res.getContentText());
        var raw  = json.candidates && json.candidates[0] &&
                   json.candidates[0].content && json.candidates[0].content.parts
                   ? json.candidates[0].content.parts[0].text || '' : '';
        raw = raw.replace(/```json|```/gi, '').trim();
        var parsed = _ocr_safeParseJson(raw);
        if (parsed) return _ocr_normalize(parsed, docType);
      } else {
        Logger.log('[VISION] ' + model + ' 失敗: HTTP ' + res.getResponseCode());
      }
      Utilities.sleep(2000);
    }
    return null;
  } catch(e) {
    Logger.log('[GEMINI VISION ERROR] ' + e.message);
    return null;
  }
}

// ============================================================
// テキスト構造化プロンプト（短いテキスト入力用）
// ============================================================
function _buildTextStructurePrompt(docType, text) {
  var truncated = text.length > 3000 ? text.substring(0, 3000) + '\n...(以下省略)' : text;

  if (docType === 'quote') {
    return [
      '以下は見積書から抽出したテキストです。このテキストを解析して、指定のJSON形式のみで返してください。',
      '説明文不要。JSONのみ出力。',
      '',
      '## 出力JSON形式',
      '{"documentNo":"見積番号","issueDate":"発行日YYYY/MM/DD","documentDate":"見積日","destCompany":"宛先会社名","destPerson":"担当者","clientName":"顧客名","subject":"件名","subtotal":小計数値,"tax":消費税数値,"totalAmount":合計数値,"lineItems":[{"itemName":"品名","spec":"仕様","qty":数量,"unit":"単位","unitPrice":単価,"amount":金額,"remarks":"備考"}]}',
      '',
      '## ルール',
      '- 金額は必ず数値型。カンマ・円・¥を除去',
      '- 合計行・消費税行はlineItemsに含めない',
      '- 不明な項目は空文字か0',
      '',
      '## 見積書テキスト',
      truncated,
    ].join('\n');
  } else {
    return [
      '以下は発注書・注文書から抽出したテキストです。このテキストを解析して、指定のJSON形式のみで返してください。',
      '説明文不要。JSONのみ出力。',
      '',
      '## 出力JSON形式',
      '{"actionType":"new/revision/cancellation","reason":"","documentNo":"注文番号","documentDate":"発注日YYYY/MM/DD","clientName":"発注者","subject":"件名","modelCode":"機種コード","orderSlipNo":"伝票番号","linkedQuoteNo":"見積番号","orderType":"試作/量産/空","subtotal":数値,"tax":数値,"totalAmount":数値,"lineItems":[{"itemName":"品名","spec":"仕様","firstDelivery":"納入日","deliveryDest":"納入先","qty":数量,"unit":"単位","unitPrice":単価,"amount":金額,"remarks":"備考"}]}',
      '',
      '## ルール',
      '- 金額は必ず数値型',
      '- 合計行はlineItemsに含めない',
      '- 差し替え/訂正→revision、取消→cancellation、それ以外→new',
      '',
      '## 注文書テキスト',
      truncated,
    ].join('\n');
  }
}

// ============================================================
// 見積書テキストのルールベース解析
// ============================================================
function _parseQuoteText(text) {
  var result = {
    documentNo: '', issueDate: '', documentDate: '',
    destCompany: '', destPerson: '', clientName: '',
    subject: '', subtotal: 0, tax: 0, totalAmount: 0, lineItems: [],
  };

  var lines = text.split(/\n/).map(function(l){ return l.trim(); }).filter(Boolean);

  lines.forEach(function(line) {
    var noMatch = line.match(/見積[　\s]*[NnＮｎ][oOｏＯ][\.．]?\s*([A-Za-z0-9\-－一ー０-９Ａ-Ｚａ-ｚ]+)/);
    if (noMatch && !result.documentNo) result.documentNo = noMatch[1];

    var noMatch2 = line.match(/[NnＮｎ][oOｏＯ][\.．\s]([A-Za-z0-9\-－ー０-９]+)/);
    if (noMatch2 && !result.documentNo) result.documentNo = noMatch2[1];

    var dateMatch = line.match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
    if (dateMatch && !result.issueDate) {
      result.issueDate = dateMatch[1] + '/' + dateMatch[2].padStart(2,'0') + '/' + dateMatch[3].padStart(2,'0');
      result.documentDate = result.issueDate;
    }

    var totalMatch = line.match(/(?:税込[　\s]*合計|合計[　\s]*金額|御請求[　\s]*金額)[　\s]*[¥￥]?([\d,，０-９]+)/);
    if (totalMatch && !result.totalAmount) result.totalAmount = _ocr_normAmt(totalMatch[1]);

    var subMatch = line.match(/小[　\s]*計[　\s]*[¥￥]?([\d,，０-９]+)/);
    if (subMatch && !result.subtotal) result.subtotal = _ocr_normAmt(subMatch[1]);

    var taxMatch = line.match(/消費税[（(]?\d+%?[)）]?[　\s]*[¥￥]?([\d,，０-９]+)/);
    if (taxMatch && !result.tax) result.tax = _ocr_normAmt(taxMatch[1]);

    var subjectMatch = line.match(/件[　\s]*名[：:][　\s]*(.+)/);
    if (subjectMatch && !result.subject) result.subject = subjectMatch[1].trim();

    var destMatch = line.match(/(.{2,20}(?:株式会社|有限会社|合同会社|㈱|㈲).*?)(?:\s+御中|　御中)?$/);
    if (destMatch && !result.destCompany && line.indexOf('御中') >= 0) {
      result.destCompany = destMatch[1].trim();
    }
  });

  result.lineItems = _parseLineItemsFromText(lines);
  if (!result.totalAmount && result.subtotal > 0) result.totalAmount = result.subtotal + result.tax;
  return result;
}

function _parseLineItemsFromText(lines) {
  var items    = [];
  var skipWords = ['合計','小計','消費税','税込','税抜','値引','割引'];

  lines.forEach(function(line) {
    var nums = line.match(/[\d,，０-９]{3,}/g);
    if (!nums || nums.length < 2) return;
    if (skipWords.some(function(w){ return line.indexOf(w) >= 0; })) return;

    var numVals = nums.map(_ocr_normAmt).filter(function(n){ return n > 0; });
    if (numVals.length < 2) return;

    var amount    = numVals[numVals.length - 1];
    var unitPrice = numVals.length >= 2 ? numVals[numVals.length - 2] : 0;
    var qty       = numVals.length >= 3 ? numVals[0] : 1;

    var firstNumIdx = line.search(/[\d０-９]/);
    var itemName    = firstNumIdx > 0 ? line.substring(0, firstNumIdx).trim() : '';
    if (!itemName || itemName.length < 1) return;
    if (amount < 100) return;

    items.push({ itemName: itemName, spec: '', qty: qty, unit: '式', unitPrice: unitPrice, amount: amount, remarks: '' });
  });

  return items.slice(0, 30);
}

// ============================================================
// OCRプロンプト（13版 高精度版）
// ============================================================
function _buildOcrPrompt(docType) {
  var commonRules = [
    '## 絶対ルール',
    '- 出力はJSON形式のみ。説明文・マークダウン・```を一切含めないこと',
    '- 金額・単価・数量は必ず数値型（文字列不可）。¥/円/,/スペースを除去してから数値化',
    '- 全角数字は半角に変換（１２３→123）',
    '- 日付はYYYY/MM/DD形式。和暦は西暦に変換（令和6年→2024、R6→2024）',
    '- 読み取れない項目は空文字""か0（数値フィールド）。nullは使わない',
    '- lineItemsの合計行・税行（「合計」「小計」「消費税」のみの行）は含めない',
    '- 同じ品名が複数行あれば全行を個別に含める（まとめない）',
  ].join('\n');

  if (docType === 'quote') {
    return [
      'あなたは高精度OCR専門AIです。添付の見積書PDFを解析し、以下のJSON形式のみで返してください。',
      '',
      '{',
      '  "documentNo": "見積番号（No.2025-123、製見2024-0001 等）",',
      '  "issueDate": "発行日 YYYY/MM/DD",',
      '  "documentDate": "見積日 YYYY/MM/DD（issueDateと同じでよい）",',
      '  "destCompany": "宛先会社名（〇〇株式会社 御中 の形式が多い）",',
      '  "destPerson": "宛先担当者名（なければ空文字）",',
      '  "clientName": "顧客名（destCompanyと同じ場合が多い）",',
      '  "subject": "件名",',
      '  "subtotal": 小計(数値),',
      '  "tax": 消費税(数値),',
      '  "totalAmount": 税込合計(数値),',
      '  "lineItems": [',
      '    {"itemName":"品名","spec":"仕様・型番","qty":数量,"unit":"単位","unitPrice":単価,"amount":金額,"remarks":"備考"}',
      '  ]',
      '}',
      '',
      commonRules,
      '',
      '## 読み取りヒント',
      '- 単価が空欄で金額のみある場合: unitPrice = amount / qty で計算して設定',
      '- 型番・規格が品名と同じ列に混在する場合は spec に分離する',
      '- 「一式」「式」は unit に設定し qty=1 にする',
    ].join('\n');
  } else {
    return [
      'あなたは高精度OCR専門AIです。添付の発注書・注文書PDFを解析し、以下のJSON形式のみで返してください。',
      '',
      '{',
      '  "actionType": "new / revision / cancellation",',
      '  "reason": "差し替えやキャンセルの理由（新規なら空文字）",',
      '  "documentNo": "発注書番号・注文番号",',
      '  "documentDate": "発注日 YYYY/MM/DD",',
      '  "clientName": "発注者（注文を出した会社名）",',
      '  "subject": "件名",',
      '  "modelCode": "機種コード・型番（なければ空文字）",',
      '  "orderSlipNo": "発注伝票番号（なければ空文字）",',
      '  "linkedQuoteNo": "対応する見積番号（記載があれば。なければ空文字）",',
      '  "orderType": "試作 または 量産（不明なら空文字）",',
      '  "subtotal": 小計(数値),',
      '  "tax": 消費税(数値),',
      '  "totalAmount": 税込合計(数値),',
      '  "lineItems": [',
      '    {"itemName":"品名","spec":"仕様","firstDelivery":"初回納入日YYYY/MM/DD","deliveryDest":"納入先","qty":数量,"unit":"単位","unitPrice":単価,"amount":金額,"remarks":"備考"}',
      '  ]',
      '}',
      '',
      commonRules,
      '',
      '## actionType判定基準',
      '- 「差し替え」「訂正」「版数更新」→ revision',
      '- 「中止」「取消」「キャンセル」→ cancellation',
      '- それ以外 → new',
      '- 取消線が引かれた行は remarks に「キャンセル」と記載しamount=0',
    ].join('\n');
  }
}

// ============================================================
// 正規化ロジック（13版）
// ============================================================

function _ocr_normalize(data, docType) {
  if (!data || typeof data !== 'object') return null;

  data.issueDate    = _ocr_normDate(data.issueDate    || data.documentDate || '');
  data.documentDate = data.issueDate;
  data.subtotal     = _ocr_normAmt(data.subtotal);
  data.tax          = _ocr_normAmt(data.tax);
  data.totalAmount  = _ocr_normAmt(data.totalAmount);

  if (!data.totalAmount && data.subtotal > 0) data.totalAmount = data.subtotal + (data.tax || 0);
  if (!data.subtotal && data.totalAmount > 0) data.subtotal = data.tax ? data.totalAmount - data.tax : data.totalAmount;

  if (Array.isArray(data.lineItems)) {
    data.lineItems = data.lineItems
      .filter(function(item) {
        if (!item || !item.itemName) return false;
        var n = String(item.itemName).trim();
        return n && n !== '合計' && n !== '小計' && n !== '消費税' && n !== '税込合計';
      })
      .map(function(item) {
        item.qty       = _ocr_normQty(item.qty);
        item.unitPrice = _ocr_normAmt(item.unitPrice);
        item.amount    = _ocr_normAmt(item.amount);
        item.unit      = String(item.unit || '式').trim();
        if (!item.unitPrice && item.amount > 0 && item.qty > 0) item.unitPrice = Math.round(item.amount / item.qty);
        if (!item.amount && item.unitPrice > 0 && item.qty > 0) item.amount = item.unitPrice * item.qty;
        item.itemName = String(item.itemName || '').trim();
        item.spec     = String(item.spec     || '').trim();
        item.remarks  = String(item.remarks  || '').trim();
        if (docType === 'order') {
          item.firstDelivery = _ocr_normDate(item.firstDelivery || '');
          item.deliveryDest  = String(item.deliveryDest || '').trim();
        }
        return item;
      });
  } else {
    data.lineItems = [];
  }

  if (docType === 'order') data.actionType = data.actionType || 'new';
  return data;
}

function _ocr_normAmt(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : Math.round(val);
  var s = String(val)
    .replace(/[０-９]/g, function(c){ return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); })
    .replace(/[¥￥円,\s　]/g, '');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : Math.round(n);
}

function _ocr_normQty(val) {
  if (val === null || val === undefined || val === '') return 1;
  if (typeof val === 'number') return isNaN(val) || val <= 0 ? 1 : val;
  var s = String(val).replace(/[０-９]/g, function(c){ return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); }).replace(/[,\s　]/g,'');
  var n = parseFloat(s);
  return isNaN(n) || n <= 0 ? 1 : n;
}

function _ocr_normDate(val) {
  if (!val) return '';
  var s = String(val).trim().replace(/[．。・]/g,'/').replace(/-/g,'/');
  var m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (m) return m[1] + '/' + m[2].padStart(2,'0') + '/' + m[3].padStart(2,'0');
  var r = s.match(/R(\d+)[\/.年](\d+)[\/.月](\d+)/);
  if (r) return (2018 + parseInt(r[1])) + '/' + r[2].padStart(2,'0') + '/' + r[3].padStart(2,'0');
  return s;
}

function _ocr_extractText(raw) {
  try {
    if (raw.candidates && raw.candidates[0] && raw.candidates[0].content && raw.candidates[0].content.parts) {
      return (raw.candidates[0].content.parts[0].text || '').replace(/```json|```/gi,'').trim();
    }
    if (raw.promptFeedback && raw.promptFeedback.blockReason) Logger.log('[OCR] ブロック: ' + raw.promptFeedback.blockReason);
    return null;
  } catch(e) { return null; }
}

function _ocr_safeParseJson(text) {
  try { return JSON.parse(text); } catch(e) {
    try {
      var fixed  = text.replace(/,\s*$/, '').replace(/,\s*\]/,']').replace(/,\s*\}/,'}');
      var opens  = (fixed.match(/\{/g)||[]).length;
      var closes = (fixed.match(/\}/g)||[]).length;
      for (var i = 0; i < opens - closes; i++) fixed += '}';
      return JSON.parse(fixed);
    } catch(e2) { return null; }
  }
}

function _ocr_extractPartial(text, docType) {
  try {
    var r  = { lineItems: [] };
    var dm = text.match(/"documentNo"\s*:\s*"([^"]+)"/);       if (dm) r.documentNo  = dm[1];
    var im = text.match(/"(?:issueDate|documentDate)"\s*:\s*"([^"]+)"/); if (im) r.issueDate = im[1];
    var cm = text.match(/"(?:destCompany|clientName)"\s*:\s*"([^"]+)"/); if (cm) r.destCompany = cm[1];
    var tm = text.match(/"totalAmount"\s*:\s*([\d.]+)/);        if (tm) r.totalAmount = Number(tm[1]);
    var sm = text.match(/"subtotal"\s*:\s*([\d.]+)/);           if (sm) r.subtotal    = Number(sm[1]);
    Logger.log('[OCR] 部分抽出: documentNo=' + r.documentNo);
    return Object.keys(r).length > 1 ? r : null;
  } catch(e) { return null; }
}

// ============================================================
// OCR品質スコア（0–100）
// ============================================================
function _calcOcrQuality(ocr) {
  if (!ocr) return 0;
  var s = 0;
  if (ocr.documentNo)  s += 20;
  if (ocr.issueDate)   s += 15;
  if (ocr.destCompany || ocr.clientName) s += 15;
  if (ocr.totalAmount > 0) s += 20;
  if (ocr.lineItems && ocr.lineItems.length > 0) {
    s += 20;
    var wp = ocr.lineItems.filter(function(i){ return i.unitPrice > 0; }).length;
    s += Math.round((wp / ocr.lineItems.length) * 10);
  }
  return Math.min(100, s);
}

// _buildOcrWarnings は 15_ocr_review_ui.gs に定義済み（そちらが使用される）

// ============================================================
// OCR処理ログ記録
// ============================================================
function _logOcrResult(fileName, status, mgmtId, detail) {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(OCR_LOG_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(OCR_LOG_SHEET_NAME);
      sheet.getRange(1,1,1,5).setValues([['日時','ファイル名','ステータス','管理ID','詳細']])
           .setBackground('#263238').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setFrozenRows(1);
      [1,2,3,4,5].forEach(function(i,idx){ sheet.setColumnWidth(i,[160,280,100,160,400][idx]); });
    }
    var row = sheet.getLastRow() + 1;
    sheet.getRange(row,1,1,5).setValues([[
      nowJST(), String(fileName||'').substring(0,100),
      status, mgmtId||'', String(detail||'').substring(0,200),
    ]]);
    var bg = { success:'#E8F5E9', ocr_failed:'#FFF3E0', error:'#FFEBEE' }[status];
    if (bg) sheet.getRange(row,1,1,5).setBackground(bg);
  } catch(e) {
    Logger.log('[_logOcrResult] ' + e.message);
  }
}

// ============================================================
// 夜間バッチOCRリトライ（14_ocr_retry_batch.gs）
// ============================================================

var OCR_RETRY_CONFIG = {
  MODEL          : 'gemini-2.5-flash',
  SLEEP_MS       : 15000,
  MAX_EXEC_SECONDS: 270,
  MAX_RETRY      : 3,
  STATUS_COL_NAME: 'OCRステータス',
  RETRY_COL_NAME : 'リトライ回数',
  BATCH_API_KEY_NAME: 'GEMINI_API_KEY_BATCH',
};

function runOcrRetryBatch() {
  var startTime = new Date().getTime();
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var sheet     = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  if (!sheet) { Logger.log('[OCR RETRY] 管理シートが見つかりません。'); return; }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1) return;

  var range   = sheet.getRange(1, 1, lastRow, lastCol);
  var values  = range.getValues();
  var headers = values[0];

  var statusColIdx = headers.indexOf(OCR_RETRY_CONFIG.STATUS_COL_NAME);
  var retryColIdx  = headers.indexOf(OCR_RETRY_CONFIG.RETRY_COL_NAME);

  if (statusColIdx === -1) {
    statusColIdx = lastCol;
    sheet.getRange(1, statusColIdx + 1).setValue(OCR_RETRY_CONFIG.STATUS_COL_NAME);
    headers.push(OCR_RETRY_CONFIG.STATUS_COL_NAME);
    lastCol++;
  }
  if (retryColIdx === -1) {
    retryColIdx = lastCol;
    sheet.getRange(1, retryColIdx + 1).setValue(OCR_RETRY_CONFIG.RETRY_COL_NAME);
    headers.push(OCR_RETRY_CONFIG.RETRY_COL_NAME);
    lastCol++;
  }

  var targetColIndexes = [
    MGMT_COLS.QUOTE_NO - 1, MGMT_COLS.ORDER_NO - 1,
    MGMT_COLS.SUBJECT - 1,  MGMT_COLS.CLIENT - 1,
    MGMT_COLS.QUOTE_DATE - 1, MGMT_COLS.ORDER_DATE - 1,
    MGMT_COLS.QUOTE_AMOUNT - 1, MGMT_COLS.ORDER_AMOUNT - 1,
  ];

  for (var i = 1; i < values.length; i++) {
    var elapsedSeconds = (new Date().getTime() - startTime) / 1000;
    if (elapsedSeconds > OCR_RETRY_CONFIG.MAX_EXEC_SECONDS) {
      Logger.log('[OCR RETRY] タイムアウト回避のため中断。完了行: ' + i);
      break;
    }

    var rowId    = values[i][MGMT_COLS.ID - 1] || 'Unknown';
    var pdfUrl   = values[i][MGMT_COLS.QUOTE_PDF_URL - 1] || values[i][MGMT_COLS.ORDER_PDF_URL - 1];
    var status   = values[i][statusColIdx] || '';
    var retryCnt = Number(values[i][retryColIdx]) || 0;

    if (status === '補完待ち' && retryCnt < OCR_RETRY_CONFIG.MAX_RETRY) {

      if (!pdfUrl) { sheet.getRange(i+1, statusColIdx+1).setValue('手動確認'); continue; }

      var missingHeaders = [];
      var missingIndexes = [];
      targetColIndexes.forEach(function(colIdx) {
        var val = String(values[i][colIdx] || '').trim();
        if (val === '' || val === '—' || val === '-') {
          missingHeaders.push(headers[colIdx]);
          missingIndexes.push(colIdx);
        }
      });

      if (missingHeaders.length === 0) { sheet.getRange(i+1, statusColIdx+1).setValue('完了'); continue; }

      Logger.log('[OCR RETRY] 対象ID: ' + rowId + ' 不足項目: ' + missingHeaders.join(', '));

      try {
        var fileId = _extractFileIdFromUrl(pdfUrl);
        if (!fileId) throw new Error('Invalid PDF URL');

        var file     = DriveApp.getFileById(fileId);
        var mimeType = file.getMimeType() || 'application/pdf';
        var base64   = Utilities.base64Encode(file.getBlob().getBytes());

        var prompt = 'この画像は書類（見積書または注文書）です。\n' +
                     '前回のデータベース登録時に以下の項目が読み取りできませんでした。\n' +
                     '画像から [' + missingHeaders.join(', ') + '] のみを推測・抽出し、厳格なJSON形式で返してください。\n' +
                     'JSONキーには日本語のカラム名をそのまま利用してください。値が取得できない場合は空文字としてください。\n' +
                     'ルール: 有効なJSONのみ。マークダウンや説明は一切記載しないこと。金額などは数字のみ。';

        var body = {
          contents: [{ parts: [
            { text: prompt },
            { inline_data: { mime_type: mimeType, data: base64 } }
          ]}],
          generationConfig: { temperature: 0.1, responseMimeType: 'application/json' }
        };

        var apiRes        = _callGeminiApiOcrRetry(OCR_RETRY_CONFIG.MODEL, body);
        var extractedData = {};
        if (apiRes) extractedData = _parseGeminiJsonRetryResponse(apiRes);

        var isAllFilled = true;
        missingIndexes.forEach(function(colIdx) {
          var hName  = headers[colIdx];
          var extVal = extractedData[hName];
          if (extVal && String(extVal).trim() !== '') {
            sheet.getRange(i+1, colIdx+1).setValue(extVal);
          } else {
            isAllFilled = false;
          }
        });

        var nextRetry = retryCnt + 1;
        sheet.getRange(i+1, retryColIdx+1).setValue(nextRetry);
        if (isAllFilled) {
          sheet.getRange(i+1, statusColIdx+1).setValue('完了');
        } else if (nextRetry >= OCR_RETRY_CONFIG.MAX_RETRY) {
          sheet.getRange(i+1, statusColIdx+1).setValue('手動確認');
        }

      } catch(e) {
        Logger.log('[OCR RETRY ERROR] ' + rowId + ': ' + e.message);
      }

      Utilities.sleep(OCR_RETRY_CONFIG.SLEEP_MS);
    }
  }
}

function resetEmptyFieldsForOcrRetry() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  if (!sheet) return;

  var values       = sheet.getDataRange().getValues();
  var headers      = values[0];
  var statusColIdx = headers.indexOf(OCR_RETRY_CONFIG.STATUS_COL_NAME);
  var retryColIdx  = headers.indexOf(OCR_RETRY_CONFIG.RETRY_COL_NAME);

  if (statusColIdx === -1) { Logger.log('OCRステータス列がありません。runOcrRetryBatchを一度実行してください。'); return; }

  var resetCount = 0;
  for (var i = 1; i < values.length; i++) {
    var subject = String(values[i][MGMT_COLS.SUBJECT - 1] || '').trim();
    var qDate   = String(values[i][MGMT_COLS.QUOTE_DATE - 1] || '').trim();
    if (subject === '' || subject === '—' || subject === '-' || qDate === '') {
      sheet.getRange(i+1, statusColIdx+1).setValue('補完待ち');
      sheet.getRange(i+1, retryColIdx+1).setValue(0);
      resetCount++;
    }
  }
  Logger.log(resetCount + ' 件を「補完待ち」にセットしました。');
}

function _extractFileIdFromUrl(url) {
  var match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

function _callGeminiApiOcrRetry(model, body) {
  var batchKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY_BATCH');
  var key = '';
  if (batchKey) {
    Logger.log('[DEBUG] 夜間用キー(BATCH)を使用: ' + batchKey.substring(0, 5) + '...');
    key = batchKey;
  } else {
    Logger.log('[DEBUG] 夜間用キーなし → 通常キーを使用');
    key = typeof CONFIG !== 'undefined' && CONFIG.GEMINI_API_KEY
          ? CONFIG.GEMINI_API_KEY
          : PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  }
  if (!key) throw new Error('APIキーがどこにも設定されていません');

  var url = CONFIG.GEMINI_API_ENDPOINT + model + ':generateContent?key=' + key;
  var res = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(body), muteHttpExceptions: true,
  });
  return JSON.parse(res.getContentText());
}

function _parseGeminiJsonRetryResponse(result) {
  try {
    var text = '';
    if (result.candidates && result.candidates[0] &&
        result.candidates[0].content && result.candidates[0].content.parts) {
      text = result.candidates[0].content.parts[0].text || '';
    }
    text = text.replace(/```json|```/gi, '').trim();
    if (!text) return {};
    return JSON.parse(text);
  } catch(e) {
    Logger.log('[JSON PARSE ERROR] ' + e.message);
    return {};
  }
}
