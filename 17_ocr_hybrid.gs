// ============================================================
// 見積書・注文書管理システム
// ファイル 17: OCR ハイブリッド戦略【コストゼロ版】
// ============================================================
//
// 【戦略】
//   見積書（自社PDF）: PDFテキスト直接抽出 → 構造化
//                      Gemini API 不使用・コストゼロ・精度100%
//
//   注文書（取引先PDF）: ① PDFテキスト直接抽出を試みる
//                        ② テキストが取れた → Geminiで構造化（少トークン）
//                        ③ スキャンPDFで取れない → Gemini Vision（最終手段）
//                        ④ Gemini失敗 → 手動入力UIへ誘導
//
// 【GASファイルの上書き優先順位】
//   このファイル(17)は 13_ocr_enhanced.gs より後に定義されるため
//   extractPdfData / _buildOcrPrompt を上書きします
//
// 【メール転記フロー（見積書）】
//   メール添付PDF → Drive保存 → PDFテキスト抽出 → 構造化 → DB登録
//   Gemini API 呼び出し: ゼロ回
//
// 【注文書フロー】
//   取引先PDF → Drive保存 → テキスト抽出試行
//     成功 → Geminiテキスト構造化（トークン少・低コスト）
//     失敗（スキャン） → Gemini Vision（必要時のみ）
//     Vision失敗 → OCR確認UI（手動入力）
// ============================================================

// ============================================================
// extractPdfData の最終上書き版
// ============================================================
function extractPdfData(driveFile, docType) {
  Logger.log('[OCR HYBRID] 開始: ' + docType + ' / ' + driveFile.getName());

  if (docType === 'quote') {
    // ===== 見積書: PDFテキスト直接抽出（Gemini不使用）=====
    return _extractQuoteFromPdfText(driveFile);
  } else {
    // ===== 注文書: テキスト抽出 → 失敗時Gemini =====
    return _extractOrderHybrid(driveFile);
  }
}

// ============================================================
// 見積書: PDFテキスト直接抽出 + ルールベース構造化
// ============================================================
function _extractQuoteFromPdfText(driveFile) {
  var text = _extractTextFromPdf(driveFile);

  if (!text || text.trim().length < 20) {
    Logger.log('[OCR HYBRID] 見積書テキスト抽出失敗 → 手動入力UIへ');
    _logOcrResult(driveFile.getName(), 'ocr_failed', null,
      '見積書PDFからテキストを抽出できませんでした。スキャンPDFの可能性があります。');
    return null;
  }

  Logger.log('[OCR HYBRID] 見積書テキスト抽出成功: ' + text.length + '文字');
  _logOcrResult(driveFile.getName(), 'text_extract', null,
    '見積書テキスト直接抽出: ' + text.length + '文字');

  // ルールベースで構造化
  var result = _parseQuoteText(text);

  // テキストが取れたが構造化が不十分な場合、Geminiで補完
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
  // ① PDFテキスト直接抽出を試みる
  var text = _extractTextFromPdf(driveFile);

  if (text && text.trim().length >= 30) {
    Logger.log('[OCR HYBRID] 注文書テキスト抽出成功: ' + text.length + '文字 → Gemini構造化');

    // ② テキストをGeminiに渡して構造化（Vision不使用 = 低コスト）
    var result = _geminiStructureText(text, 'order');
    if (result) {
      _logOcrResult(driveFile.getName(), 'success', null,
        '注文書テキスト+Gemini構造化: ' + (result.documentNo||'?'));
      return result;
    }
  }

  // ③ スキャンPDF等でテキスト取れない → Gemini Vision（最終手段）
  Logger.log('[OCR HYBRID] 注文書テキスト抽出失敗 → Gemini Vision試行');
  var apiKey = getGeminiApiKey();
  if (!apiKey) {
    Logger.log('[OCR HYBRID] APIキー未設定 → 手動入力UIへ誘導');
    _logOcrResult(driveFile.getName(), 'ocr_failed', null,
      'スキャンPDF + APIキー未設定。手動入力が必要です。');
    return null;
  }

  var visionResult = _geminiVisionOcr(driveFile, 'order', apiKey);
  if (visionResult) {
    _logOcrResult(driveFile.getName(), 'success', null,
      '注文書Vision OCR: ' + (visionResult.documentNo||'?'));
    return visionResult;
  }

  // ④ 全て失敗 → nullを返してOCR確認UIで手動入力
  _logOcrResult(driveFile.getName(), 'ocr_failed', null,
    '全OCR手法失敗。OCR確認UIで手動入力してください。');
  return null;
}

// ============================================================
// PDFからテキストを直接抽出（GASネイティブ）
// ============================================================
function _extractTextFromPdf(driveFile) {
  try {
    // 方法1: Google Docs変換 → テキスト取得
    var blob   = driveFile.getBlob();
    var folder = DriveApp.getFolderById(
      driveFile.getParents().hasNext()
        ? driveFile.getParents().next().getId()
        : DriveApp.getRootFolder().getId()
    );

    // Drive API v3でPDF→Docsに変換してテキスト抽出
    var metadata = {
      name    : '_ocr_tmp_' + Date.now(),
      mimeType: 'application/vnd.google-apps.document',
    };
    var insertedFile = Drive.Files.create(metadata, blob, {
      ocrLanguage: 'ja',
      fields      : 'id',
    });
    Utilities.sleep(2000); // 変換完了待機

    var docFile = DriveApp.getFileById(insertedFile.id);
    var doc     = DocumentApp.openById(insertedFile.id);
    var text    = doc.getBody().getText();

    // 一時ファイルを削除
    docFile.setTrashed(true);

    if (text && text.trim().length > 10) {
      Logger.log('[TEXT EXTRACT] Drive OCR成功: ' + text.length + '文字');
      return text;
    }

    Logger.log('[TEXT EXTRACT] Drive OCRテキストなし（スキャンPDFの可能性）');
    return '';

  } catch(e) {
    Logger.log('[TEXT EXTRACT ERROR] ' + e.message);
    // Drive APIが使えない場合は空を返す
    return '';
  }
}

// ============================================================
// Geminiにテキストを渡して構造化JSON生成（Vision不使用・低コスト）
// ============================================================
function _geminiStructureText(text, docType) {
  var apiKey = getGeminiApiKey();
  if (!apiKey) return null;

  var prompt = docType === 'quote'
    ? _buildTextStructurePrompt('quote', text)
    : _buildTextStructurePrompt('order', text);

  var body = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature      : 0.0,
      responseMimeType : 'application/json',
      maxOutputTokens  : 2048,
    },
  };

  try {
    var url = CONFIG.GEMINI_API_ENDPOINT +
      'gemini-2.0-flash-lite' +  // ★最安モデルを使用
      ':generateContent?key=' + apiKey;

    var res  = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(body), muteHttpExceptions: true,
    });

    if (res.getResponseCode() !== 200) {
      Logger.log('[GEMINI TEXT] HTTP ' + res.getResponseCode());
      return null;
    }

    var json = JSON.parse(res.getContentText());
    var raw  = json.candidates && json.candidates[0] &&
               json.candidates[0].content &&
               json.candidates[0].content.parts
               ? json.candidates[0].content.parts[0].text || '' : '';

    raw = raw.replace(/```json|```/gi, '').trim();
    var parsed = _ocr_safeParseJson(raw);
    return parsed ? _ocr_normalize(parsed, docType) : null;

  } catch(e) {
    Logger.log('[GEMINI TEXT ERROR] ' + e.message);
    return null;
  }
}

// ============================================================
// Gemini Vision OCR（スキャンPDF用・最終手段）
// ============================================================
function _geminiVisionOcr(driveFile, docType, apiKey) {
  try {
    var base64 = Utilities.base64Encode(driveFile.getBlob().getBytes());
    var body   = {
      contents: [{ parts: [
        { text: _buildOcrPrompt(docType) },
        { inline_data: { mime_type: 'application/pdf', data: base64 } },
      ]}],
      generationConfig: {
        temperature      : 0.05,
        responseMimeType : 'application/json',
        maxOutputTokens  : 2048,
      },
    };

    // gemini-2.0-flash-lite（最安・無料枠対応）を優先
    var models = ['gemini-2.0-flash-lite', CONFIG.GEMINI_PRIMARY_MODEL];
    for (var i = 0; i < models.length; i++) {
      var url = CONFIG.GEMINI_API_ENDPOINT + models[i] + ':generateContent?key=' + apiKey;
      var res = UrlFetchApp.fetch(url, {
        method: 'post', contentType: 'application/json',
        payload: JSON.stringify(body), muteHttpExceptions: true,
      });
      if (res.getResponseCode() === 200) {
        var json = JSON.parse(res.getContentText());
        var raw  = json.candidates && json.candidates[0] &&
                   json.candidates[0].content &&
                   json.candidates[0].content.parts
                   ? json.candidates[0].content.parts[0].text || '' : '';
        raw = raw.replace(/```json|```/gi, '').trim();
        var parsed = _ocr_safeParseJson(raw);
        if (parsed) return _ocr_normalize(parsed, docType);
      }
      Logger.log('[VISION] ' + models[i] + ' HTTP ' + res.getResponseCode());
      Utilities.sleep(2000);
    }
    return null;
  } catch(e) {
    Logger.log('[GEMINI VISION ERROR] ' + e.message);
    return null;
  }
}

// ============================================================
// テキスト → 構造化JSON プロンプト（テキスト入力用・短い）
// ============================================================
function _buildTextStructurePrompt(docType, text) {
  // テキストが長すぎる場合は先頭3000文字に制限
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
// 見積書テキストのルールベース解析（Gemini不使用）
// ============================================================
function _parseQuoteText(text) {
  var result = {
    documentNo  : '',
    issueDate   : '',
    documentDate: '',
    destCompany : '',
    destPerson  : '',
    clientName  : '',
    subject     : '',
    subtotal    : 0,
    tax         : 0,
    totalAmount : 0,
    lineItems   : [],
  };

  var lines = text.split(/\n/).map(function(l){ return l.trim(); }).filter(Boolean);

  lines.forEach(function(line) {
    // 見積番号
    var noMatch = line.match(/見積[　\s]*[NnＮｎ][oOｏＯ][\.．]?\s*([A-Za-z0-9\-－一\u30FC０-９Ａ-Ｚａ-ｚ]+)/);
    if (noMatch && !result.documentNo) result.documentNo = noMatch[1];

    // 番号パターン（No. XXXX）
    var noMatch2 = line.match(/[NnＮｎ][oOｏＯ][\.．\s]([A-Za-z0-9\-－\u30FC０-９]+)/);
    if (noMatch2 && !result.documentNo) result.documentNo = noMatch2[1];

    // 発行日・見積日
    var dateMatch = line.match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
    if (dateMatch && !result.issueDate) {
      result.issueDate = dateMatch[1] + '/' +
        dateMatch[2].padStart(2,'0') + '/' + dateMatch[3].padStart(2,'0');
      result.documentDate = result.issueDate;
    }

    // 合計金額（税込）
    var totalMatch = line.match(/(?:税込[　\s]*合計|合計[　\s]*金額|御請求[　\s]*金額)[　\s]*[¥￥]?([\d,，０-９]+)/);
    if (totalMatch && !result.totalAmount) {
      result.totalAmount = _ocr_normAmt(totalMatch[1]);
    }

    // 小計
    var subMatch = line.match(/小[　\s]*計[　\s]*[¥￥]?([\d,，０-９]+)/);
    if (subMatch && !result.subtotal) {
      result.subtotal = _ocr_normAmt(subMatch[1]);
    }

    // 消費税
    var taxMatch = line.match(/消費税[（(]?\d+%?[)）]?[　\s]*[¥￥]?([\d,，０-９]+)/);
    if (taxMatch && !result.tax) {
      result.tax = _ocr_normAmt(taxMatch[1]);
    }

    // 件名
    var subjectMatch = line.match(/件[　\s]*名[：:][　\s]*(.+)/);
    if (subjectMatch && !result.subject) result.subject = subjectMatch[1].trim();

    // 宛先（〇〇株式会社 御中）
    var destMatch = line.match(/(.{2,20}(?:株式会社|有限会社|合同会社|㈱|㈲).*?)(?:\s+御中|　御中)?$/);
    if (destMatch && !result.destCompany && line.indexOf('御中') >= 0) {
      result.destCompany = destMatch[1].trim();
    }
  });

  // 明細行の解析（数値が3つ以上並ぶ行を明細として解釈）
  var lineItems = _parseLineItemsFromText(lines);
  result.lineItems = lineItems;

  // totalAmountがない場合はsubtotal+taxから計算
  if (!result.totalAmount && result.subtotal > 0) {
    result.totalAmount = result.subtotal + result.tax;
  }

  return result;
}

/**
 * テキスト行から明細行を解析
 * 品名・数量・単価・金額が並ぶ行を検出
 */
function _parseLineItemsFromText(lines) {
  var items = [];
  // 除外キーワード（合計行等）
  var skipWords = ['合計','小計','消費税','税込','税抜','値引','割引'];

  lines.forEach(function(line) {
    // 明細の特徴: 文字列 + 数値が複数含まれる行
    var nums = line.match(/[\d,，０-９]{3,}/g);
    if (!nums || nums.length < 2) return;

    // スキップ対象
    if (skipWords.some(function(w){ return line.indexOf(w) >= 0; })) return;

    // 数値を正規化
    var numVals = nums.map(_ocr_normAmt).filter(function(n){ return n > 0; });
    if (numVals.length < 2) return;

    // 最後の数値=金額、その1つ前=単価、最初の数値=数量と推定
    var amount    = numVals[numVals.length - 1];
    var unitPrice = numVals.length >= 2 ? numVals[numVals.length - 2] : 0;
    var qty       = numVals.length >= 3 ? numVals[0] : 1;

    // 品名: 行の先頭から最初の数値の前まで
    var firstNumIdx = line.search(/[\d０-９]/);
    var itemName    = firstNumIdx > 0 ? line.substring(0, firstNumIdx).trim() : '';

    if (!itemName || itemName.length < 1) return;
    if (amount < 100) return; // 金額が小さすぎる行は除外

    items.push({
      itemName : itemName,
      spec     : '',
      qty      : qty,
      unit     : '式',
      unitPrice: unitPrice,
      amount   : amount,
      remarks  : '',
    });
  });

  return items.slice(0, 30); // 最大30行
}

// ============================================================
// processUploadedPdf の上書き（ハイブリッドOCR対応）
// ============================================================
function processUploadedPdf(base64Data, fileName, docType, orderType) {
  try {
    var blob      = Utilities.newBlob(
      Utilities.base64Decode(base64Data), 'application/pdf', fileName
    );
    var folderId  = (docType === 'order')
      ? _getOrderFolderId(orderType || '')
      : CONFIG.WEB_UPLOAD_FOLDER_ID;
    var folder    = DriveApp.getFolderById(folderId);
    var saved     = 'MANUAL_' + nowJST().replace(/[\\/: ]/g,'') + '_' + fileName;
    var file      = folder.createFile(blob.setName(saved));
    var pdfUrl    = file.getUrl();
    var folderUrl = getFolderUrl(folderId);

    Logger.log('[UPLOAD HYBRID] 保存: ' + saved);

    var ocr = extractPdfData(file, docType); // ←ハイブリッド版が呼ばれる

    var quality = _calcOcrQuality(ocr);
    var warnings = _buildOcrWarnings(ocr, quality);

    // OCRが完全失敗した場合でも確認UIに渡せる空テンプレートを返す
    if (!ocr) {
      ocr = {
        documentNo: '', issueDate: '', documentDate: '',
        destCompany: '', destPerson: '', clientName: '',
        subject: '', subtotal: 0, tax: 0, totalAmount: 0,
        lineItems: [],
        actionType: docType === 'order' ? 'new' : undefined,
      };
      quality = 0;
      warnings = [{
        level: 'error',
        msg: 'OCRでテキストを取得できませんでした。手動で入力してください。'
      }];
    }

    return {
      success     : true,
      mgmtId      : null, // OCR確認UI経由で登録するためここではnull
      documentNo  : ocr.documentNo,
      clientName  : ocr.destCompany || ocr.clientName,
      totalAmount : ocr.totalAmount,
      lineCount   : ocr.lineItems ? ocr.lineItems.length : 0,
      savedFolder : folderUrl,
      orderType   : orderType || (ocr && ocr.orderType) || '',
      modelCode   : (ocr && ocr.modelCode)   || '',
      orderSlipNo : (ocr && ocr.orderSlipNo) || '',
      issueDate   : (ocr && ocr.issueDate)   || '',
      destCompany : (ocr && ocr.destCompany) || '',
      destPerson  : (ocr && ocr.destPerson)  || '',
      ocrResult   : ocr,
      quality     : quality,
      warnings    : warnings,
      pdfUrl      : pdfUrl,
    };
  } catch(e) {
    Logger.log('[UPLOAD HYBRID ERROR] ' + e.message);
    _logOcrResult(fileName, 'error', null, e.message);
    return { success: false, error: e.message };
  }
}
