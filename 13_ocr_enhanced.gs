// ============================================================
// 見積書・注文書管理システム
// ファイル 13: OCR強化パッチ【Phase2 品質向上】
// ============================================================
//
// 【このファイルの役割】
//   02_ocr_and_processing.gs の extractPdfData / _buildOcrPrompt を
//   高精度版に置き換える。GASは同名関数が複数あると後に定義した方が優先。
//   このファイルをプロジェクトに追加するだけでOCR精度が向上する。
//
// 【強化内容】
//   1. 温度パラメータを0.05に下げて決定的な出力を実現
//   2. プロンプトを大幅強化（表記ゆれ・フォーマット差・和暦・全角数字を吸収）
//   3. JSONパース失敗時の部分抽出リカバリ
//   4. 抽出後の数値正規化（¥/円/カンマ/全角→数値）
//   5. OCR品質スコア（0-100）をWebアプリに返す
//   6. 全処理結果を「OCR処理ログ」シートに記録
// ============================================================

var OCR_LOG_SHEET_NAME = 'OCR処理ログ';

// ============================================================
// extractPdfData の上書き（02の同名関数より後に定義するため優先される）
// ============================================================
function extractPdfData(driveFile, docType) {
  var base64 = Utilities.base64Encode(driveFile.getBlob().getBytes());
  var prompt = _buildOcrPrompt(docType);

  var body = {
    contents: [{ parts: [
      { text: prompt },
      { inline_data: { mime_type: 'application/pdf', data: base64 } }
    ]}],
    generationConfig: {
      temperature      : 0.05,
      responseMimeType : 'application/json',
      maxOutputTokens  : 4096,
    },
  };

  // ① プライマリモデル
  var raw = _callGeminiApi(CONFIG.GEMINI_PRIMARY_MODEL, body);
  // ② フォールバック
  if (!raw) {
    Logger.log('[OCR] プライマリ失敗 → フォールバック試行');
    raw = _callGeminiApi(CONFIG.GEMINI_FALLBACK_MODEL, body);
  }
  if (!raw) { Logger.log('[OCR] 全モデル失敗'); return null; }

  var text = _ocr_extractText(raw);
  if (!text) return null;
  Logger.log('[OCR RAW] ' + text.substring(0, 600));

  // ③ パース → 部分抽出リカバリ
  var parsed = _ocr_safeParseJson(text);
  if (!parsed) {
    Logger.log('[OCR] JSONパース失敗 → 部分抽出試行');
    parsed = _ocr_extractPartial(text, docType);
  }
  if (!parsed) return null;

  // ④ 数値正規化・バリデーション
  return _ocr_normalize(parsed, docType);
}

// ============================================================
// _buildOcrPrompt の上書き
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
// 正規化ロジック
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
    .replace(/[０-９]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); })
    .replace(/[¥￥円,\s　]/g, '');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : Math.round(n);
}

function _ocr_normQty(val) {
  if (val === null || val === undefined || val === '') return 1;
  if (typeof val === 'number') return isNaN(val) || val <= 0 ? 1 : val;
  var s = String(val).replace(/[０-９]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); }).replace(/[,\s　]/g,'');
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
      var fixed = text.replace(/,\s*$/, '').replace(/,\s*\]/,']').replace(/,\s*\}/,'}');
      var opens = (fixed.match(/\{/g)||[]).length, closes = (fixed.match(/\}/g)||[]).length;
      for (var i = 0; i < opens - closes; i++) fixed += '}';
      return JSON.parse(fixed);
    } catch(e2) { return null; }
  }
}

function _ocr_extractPartial(text, docType) {
  try {
    var r = { lineItems: [] };
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
// OCR品質スコア（processUploadedPdfの返り値に追加）
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
    var bg = {success:'#E8F5E9', ocr_failed:'#FFF3E0', error:'#FFEBEE'}[status];
    if (bg) sheet.getRange(row,1,1,5).setBackground(bg);
  } catch(e) {
    Logger.log('[_logOcrResult] ' + e.message);
  }
}
