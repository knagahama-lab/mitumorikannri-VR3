// ============================================================
// 見積・注文 管理システム
// ファイル 7/7: Gemini AIチャットボット API
// ============================================================
//
// 機能:
//   - スプレッドシートデータの全文検索
//   - DrivePDFキャッシュからファイル名検索
//   - 関連PDFの中身をGemini Vision APIで読み取り
//   - 全情報をGemini AIに渡して自然言語で回答
//
// コスト目安: 1回の検索 約5〜15円（PDF読み込み数による）
// ============================================================

var CHATBOT_CONFIG = {
  MAX_PDF_READ: 5,    // 1回の検索で読み込む最大PDF数
  MAX_SS_ROWS: 300,  // スプレッドシートから渡す最大行数
  MAX_TOKENS: 4000, // Gemini への最大トークン数
};

// ============================================================
// メインエントリーポイント
// ============================================================

function apiChatbotQuery(p) {
  try {
    var question = String(p.question || '').trim();
    var history = p.history || [];  // 会話履歴
    if (!question) return { success: false, error: '質問が空です' };

    var apiKey = CONFIG.GEMINI_API_KEY;
    if (!apiKey) return { success: false, error: 'GEMINI_API_KEY が設定されていません' };

    Logger.log('[CHATBOT] 質問: ' + question);

    // ===== Step1: スプレッドシートデータを収集 =====
    var ssContext = _buildSpreadsheetContext(question);

    // ===== Step2: Driveキャッシュからファイル候補を取得 =====
    var driveFiles = _searchDriveFilesForChat(question);

    // ===== Step3: 関連PDFの中身を読み取り =====
    var pdfContents = _readPdfContents(driveFiles, apiKey, question);

    // ===== Step4: Geminiに全情報を渡して回答生成 =====
    var answer = _generateChatAnswer(question, history, ssContext, driveFiles, pdfContents, apiKey);

    return {
      success: true,
      answer: answer,
      ssHits: ssContext.hitCount,
      driveHits: driveFiles.length,
      pdfRead: pdfContents.length,
    };
  } catch (e) {
    Logger.log('[CHATBOT ERROR] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

// ============================================================
// Step1: スプレッドシートから関連データを収集
// ============================================================

function _buildSpreadsheetContext(question) {
  var keywords = _extractKeywords(question);
  var results = [];

  // 管理シート検索
  var mgmtData = getAllMgmtData();
  mgmtData.forEach(function (row) {
    var text = [
      row[MGMT_COLS.QUOTE_NO - 1],
      row[MGMT_COLS.ORDER_NO - 1],
      row[MGMT_COLS.SUBJECT - 1],
      row[MGMT_COLS.CLIENT - 1],
      row[MGMT_COLS.MODEL_CODE - 1],
      row[MGMT_COLS.ORDER_SLIP_NO - 1],
    ].join(' ').toLowerCase();

    if (keywords.some(function (k) { return text.indexOf(k) >= 0; })) {
      results.push({
        type: '管理',
        quoteNo: String(row[MGMT_COLS.QUOTE_NO - 1] || ''),
        orderNo: String(row[MGMT_COLS.ORDER_NO - 1] || ''),
        subject: String(row[MGMT_COLS.SUBJECT - 1] || ''),
        client: String(row[MGMT_COLS.CLIENT - 1] || ''),
        status: String(row[MGMT_COLS.STATUS - 1] || ''),
        amount: String(row[MGMT_COLS.ORDER_AMOUNT - 1] || row[MGMT_COLS.QUOTE_AMOUNT - 1] || ''),
        model: String(row[MGMT_COLS.MODEL_CODE - 1] || ''),
        date: String(row[MGMT_COLS.ORDER_DATE - 1] || row[MGMT_COLS.QUOTE_DATE - 1] || ''),
        pdfUrl: String(row[MGMT_COLS.ORDER_PDF_URL - 1] || row[MGMT_COLS.QUOTE_PDF_URL - 1] || ''),
      });
    }
  });

  // 見積書シート検索（品名・仕様）
  var ss = getSpreadsheet();
  var qs = ss.getSheetByName(CONFIG.SHEET_QUOTES);
  if (qs && qs.getLastRow() > 1) {
    var qData = qs.getRange(2, 1, Math.min(qs.getLastRow() - 1, 500), 15).getValues();
    var seenMgmtIds = {};
    qData.forEach(function (row) {
      var itemName = String(row[6] || '');
      var spec = String(row[7] || '');
      var mgmtId = String(row[0] || '');
      var text = (itemName + ' ' + spec).toLowerCase();
      if (seenMgmtIds[mgmtId]) return;
      if (keywords.some(function (k) { return text.indexOf(k) >= 0; })) {
        seenMgmtIds[mgmtId] = true;
        results.push({
          type: '見積明細',
          quoteNo: String(row[1] || ''),
          subject: itemName + (spec ? '（' + spec + '）' : ''),
          client: String(row[3] || ''),
          amount: String(row[11] || ''),
          date: _toDateStr(row[2]),
          pdfUrl: String(row[13] || ''),
        });
      }
    });
  }

  // 注文書シート検索（品名・仕様）
  var os = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  if (os && os.getLastRow() > 1) {
    var oData = os.getRange(2, 1, Math.min(os.getLastRow() - 1, 500), 19).getValues();
    var seenOIds = {};
    oData.forEach(function (row) {
      var itemName = String(row[8] || '');
      var spec = String(row[9] || '');
      var mgmtId = String(row[0] || '');
      var text = (itemName + ' ' + spec).toLowerCase();
      if (seenOIds[mgmtId]) return;
      if (keywords.some(function (k) { return text.indexOf(k) >= 0; })) {
        seenOIds[mgmtId] = true;
        results.push({
          type: '注文明細',
          orderNo: String(row[1] || ''),
          subject: itemName + (spec ? '（' + spec + '）' : ''),
          model: String(row[5] || ''),
          amount: String(row[15] || ''),
          date: _toDateStr(row[4]),
          pdfUrl: String(row[17] || ''),
        });
      }
    });
  }

  // 見積台帳検索
  var ledger = getAllLedgerData();
  ledger.forEach(function (row) {
    var text = [row[1], row[3], row[4], row[5]].join(' ').toLowerCase();
    if (keywords.some(function (k) { return text.indexOf(k) >= 0; })) {
      var obj = _ledgerRowToObject(row);
      results.push({
        type: '見積台帳',
        quoteNo: obj.quoteNo,
        subject: obj.subject,
        client: obj.dest,
        status: obj.status,
        date: obj.issueDate,
        pdfUrl: obj.saveUrl,
      });
    }
  });

  Logger.log('[CHATBOT] SS hits: ' + results.length);
  return {
    hits: results.slice(0, CHATBOT_CONFIG.MAX_SS_ROWS),
    hitCount: results.length,
  };
}

// ============================================================
// Step2: Driveキャッシュからファイル候補を取得
// ============================================================

function _searchDriveFilesForChat(question) {
  try {
    var props = PropertiesService.getScriptProperties();
    var cached = props.getProperty(DRIVE_CACHE_KEY);
    if (!cached) return [];

    var files = JSON.parse(cached);
    var keywords = _extractKeywords(question);

    var matches = files.filter(function (f) {
      var name = String(f.name || '').toLowerCase();
      return keywords.some(function (k) { return name.indexOf(k) >= 0; });
    });

    // 更新日降順で上位を返す
    matches.sort(function (a, b) {
      return String(b.updatedAt || '').localeCompare(String(a.updatedAt || ''));
    });

    Logger.log('[CHATBOT] Drive hits: ' + matches.length);
    return matches.slice(0, 20); // 上位20件をPDF読み込み候補に
  } catch (e) {
    Logger.log('[CHATBOT] Drive search error: ' + e.message);
    return [];
  }
}

// ============================================================
// Step3: 関連PDFの中身をGemini Vision APIで読み取り
// ============================================================

function _readPdfContents(driveFiles, apiKey, question) {
  var results = [];
  var readCount = 0;
  var maxRead = CHATBOT_CONFIG.MAX_PDF_READ;

  for (var i = 0; i < driveFiles.length && readCount < maxRead; i++) {
    var f = driveFiles[i];
    try {
      var file = DriveApp.getFileById(f.id);
      var blob = file.getBlob();
      var b64 = Utilities.base64Encode(blob.getBytes());

      // Gemini Vision でPDF内容を抽出
      var endpoint = CONFIG.GEMINI_API_ENDPOINT + CONFIG.GEMINI_PRIMARY_MODEL + ':generateContent?key=' + apiKey;
      var payload = {
        contents: [{
          parts: [
            { text: '以下のPDFから、見積番号・注文番号・品名・仕様・数量・単価・金額・会社名・日付・件名などの重要情報を箇条書きで抽出してください。200文字以内で簡潔に。' },
            { inline_data: { mime_type: 'application/pdf', data: b64 } }
          ]
        }],
        generationConfig: { maxOutputTokens: 300, temperature: 0 }
      };

      var res = UrlFetchApp.fetch(endpoint, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      });
      var json = JSON.parse(res.getContentText());
      var text = (json.candidates && json.candidates[0] &&
        json.candidates[0].content && json.candidates[0].content.parts)
        ? json.candidates[0].content.parts.map(function (p) { return p.text || ''; }).join('')
        : '';

      if (text) {
        results.push({ name: f.name, url: f.url, content: text });
        readCount++;
      }
    } catch (e) {
      Logger.log('[CHATBOT] PDF read error: ' + f.name + ' ' + e.message);
    }
  }

  Logger.log('[CHATBOT] PDF read: ' + results.length + '件');
  return results;
}

// ============================================================
// Step4: Geminiで回答生成
// ============================================================

function _generateChatAnswer(question, history, ssContext, driveFiles, pdfContents, apiKey) {
  // コンテキストを構築
  var contextParts = [];

  if (ssContext.hits.length > 0) {
    contextParts.push('【スプレッドシートの検索結果（' + ssContext.hits.length + '件）】');
    ssContext.hits.slice(0, 50).forEach(function (h) {
      var line = '[' + h.type + '] ';
      if (h.quoteNo) line += '見積No:' + h.quoteNo + ' ';
      if (h.orderNo) line += '注文No:' + h.orderNo + ' ';
      if (h.subject) line += '件名:' + h.subject + ' ';
      if (h.client) line += '顧客:' + h.client + ' ';
      if (h.model) line += '機種:' + h.model + ' ';
      if (h.amount) line += '金額:' + h.amount + ' ';
      if (h.status) line += 'ステータス:' + h.status + ' ';
      if (h.date) line += '日付:' + h.date + ' ';
      if (h.pdfUrl) line += 'PDF:' + h.pdfUrl;
      contextParts.push(line);
    });
  }

  if (driveFiles.length > 0) {
    contextParts.push('\n【DriveのPDFファイル候補（' + driveFiles.length + '件）】');
    driveFiles.slice(0, 10).forEach(function (f) {
      contextParts.push('ファイル名:' + f.name + ' 更新日:' + (f.updatedAt || '') + ' URL:' + (f.url || ''));
    });
  }

  if (pdfContents.length > 0) {
    contextParts.push('\n【PDFの中身（読み込み済み ' + pdfContents.length + '件）】');
    pdfContents.forEach(function (p) {
      contextParts.push('--- ' + p.name + ' ---');
      contextParts.push(p.content);
    });
  }

  // 会話履歴を構築
  var messages = [];

  // システムプロンプト
  var systemPrompt =
    'あなたは見積書・注文書管理システムのAIアシスタントです。' +
    'ユーザーの質問に対して、提供されたデータを元に正確に回答してください。' +
    'PDFのURLや見積番号・注文番号を具体的に示してください。' +
    '見つからない場合はその旨を伝えてください。回答は日本語で簡潔に。';

  // 過去の会話履歴
  history.forEach(function (h) {
    messages.push({ role: h.role, parts: [{ text: h.content }] });
  });

  // 今回の質問（コンテキスト付き）
  var userMessage = question;
  if (contextParts.length > 0) {
    userMessage = '以下のデータを参考に質問に答えてください。\n\n' +
      contextParts.join('\n') + '\n\n質問: ' + question;
  } else {
    userMessage = question + '\n\n（注: 関連するデータが見つかりませんでした）';
  }
  messages.push({ role: 'user', parts: [{ text: userMessage }] });

  var endpoint = CONFIG.GEMINI_API_ENDPOINT + CONFIG.GEMINI_PRIMARY_MODEL + ':generateContent?key=' + apiKey;
  var payload = {
    system_instruction: { parts: [{ text: systemPrompt }] },
    contents: messages,
    generationConfig: { maxOutputTokens: 1000, temperature: 0.3 },
  };

  var res = UrlFetchApp.fetch(endpoint, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  var json = JSON.parse(res.getContentText());

  if (json.error) {
    Logger.log('[CHATBOT] Gemini error: ' + JSON.stringify(json.error));
    return 'エラー: ' + json.error.message;
  }

  return (json.candidates && json.candidates[0] &&
    json.candidates[0].content && json.candidates[0].content.parts)
    ? json.candidates[0].content.parts.map(function (p) { return p.text || ''; }).join('')
    : '回答を生成できませんでした。';
}

// ============================================================
// ユーティリティ
// ============================================================

function _extractKeywords(text) {
  // 助詞・一般的な語を除いて2文字以上のキーワードを抽出
  var stopWords = ['について', 'から', 'まで', 'です', 'ます', 'ください', 'した', 'して',
    'ある', 'いる', 'する', 'ない', 'また', 'その', 'この', 'それ', 'これ'];
  var words = text.toLowerCase()
    .replace(/[、。！？\.\,!?]/g, ' ')
    .split(/\s+/)
    .filter(function (w) {
      return w.length >= 2 && !stopWords.some(function (sw) { return w === sw; });
    });
  return words.length > 0 ? words : [text.toLowerCase()];
}