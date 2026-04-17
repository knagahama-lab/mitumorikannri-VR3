// ============================================================
// 見積書・注文書管理システム
// ファイル 16: AIチャット強化版 with Gemini Function Calling
// ============================================================
//
// 【07 chatbot api.gs との違い】
//   旧: 検索して答えるだけ（読み取り専用）
//   新: Gemini Function Calling でDBを操作できる
//       「受注登録して」→ 実際にステータスを更新
//       「A社の見積と注文を紐づけて」→ 実際に紐づけ実行
//       「単価比較して」→ 比較結果を構造化して返す
//
// 【利用可能なアクション（Geminiが自律的に選択）】
//   - search_cases          : 案件・見積・注文の検索
//   - get_case_detail       : 案件の詳細（明細行含む）
//   - compare_prices        : 見積書と注文書の単価比較
//   - update_status         : ステータス更新
//   - link_quote_order      : 見積書と注文書の手動紐づけ
//   - search_unit_price     : 過去の単価を調べる
//   - get_ocr_log           : OCRの処理状況を確認
//   - get_unlinked_orders   : 未紐づけ注文書の一覧
// ============================================================

// ============================================================
// メインAPI（03_webapp_ap.gsから呼ばれる）
// ============================================================
function apiChatbotQuery(p) {
  try {
    var question = String(p.question || '').trim();
    var history  = p.history  || [];
    var mode     = p.mode     || 'auto'; // auto / search_only

    if (!question) return { success: false, error: '質問が空です' };

    var apiKey = CONFIG.GEMINI_API_KEY ||
      PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return { success: false, error: 'GEMINI_API_KEY が設定されていません' };

    Logger.log('[CHAT v2] 質問: ' + question);

    // Function Calling で実行（Geminiが必要な関数を自律選択）
    var result = _runChatWithFunctionCalling(question, history, apiKey, mode);
    return result;

  } catch(e) {
    Logger.log('[CHAT v2 ERROR] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

// ============================================================
// Gemini Function Calling 実行ループ
// ============================================================
function _runChatWithFunctionCalling(question, history, apiKey, mode) {
  var tools = _buildFunctionDefinitions();
  var messages = _buildInitialMessages(question, history);
  var actionLog = [];   // 実行したアクションの記録
  var maxCycles = 5;    // 無限ループ防止

  for (var cycle = 0; cycle < maxCycles; cycle++) {
    var response = _callGeminiWithTools(messages, tools, apiKey);
    if (!response) return { success: false, error: 'Gemini APIの呼び出しに失敗しました' };

    var candidates = response.candidates;
    if (!candidates || !candidates[0]) break;

    var content = candidates[0].content;
    if (!content || !content.parts) break;

    // Function Callがあるか確認
    var funcCalls = content.parts.filter(function(p) { return p.functionCall; });

    if (funcCalls.length === 0) {
      // 関数呼び出しなし → テキスト回答を返す
      var answer = content.parts
        .filter(function(p) { return p.text; })
        .map(function(p) { return p.text; })
        .join('');

      return {
        success   : true,
        answer    : answer,
        actionLog : actionLog,
        mode      : 'function_calling',
      };
    }

    // Function Call を実行
    messages.push({ role: 'model', parts: content.parts });

    var funcResults = [];
    funcCalls.forEach(function(part) {
      var fc   = part.functionCall;
      var name = fc.name;
      var args = fc.args || {};

      Logger.log('[CHAT v2] Function Call: ' + name + ' args=' + JSON.stringify(args));

      var result = _executeFunction(name, args);
      actionLog.push({ function: name, args: args, success: result.success });

      funcResults.push({
        functionResponse: {
          name    : name,
          response: result,
        }
      });
    });

    messages.push({ role: 'user', parts: funcResults });
  }

  return { success: false, error: '応答の生成に失敗しました' };
}

// ============================================================
// Gemini API 呼び出し（Tools付き）
// ============================================================
function _callGeminiWithTools(messages, tools, apiKey) {
  var url = CONFIG.GEMINI_API_ENDPOINT +
    CONFIG.GEMINI_PRIMARY_MODEL + ':generateContent?key=' + apiKey;

  var systemPrompt = [
    'あなたは見積書・注文書管理システムの優秀なAIアシスタントです。',
    'ユーザーの依頼を理解し、必要な関数を呼び出してデータを取得・操作し、',
    '具体的な情報（番号・金額・日付・URL）を含めて日本語で回答してください。',
    '',
    '【回答スタイル】',
    '- 検索結果は表形式または箇条書きで見やすく整理する',
    '- 金額は「¥1,234,000」のようにカンマ区切りで表示',
    '- URLは「[PDFを開く]」のようにリンクテキストで表示',
    '- アクション（ステータス更新・紐づけ等）を実行した場合は結果を明確に報告',
    '- 見つからない場合は「該当する案件が見つかりませんでした」と正直に伝える',
    '',
    '【注意事項】',
    '- ステータス更新や紐づけ等の破壊的操作は、ユーザーが明示的に依頼した場合のみ実行',
    '- 曖昧な場合は実行前に確認を取る',
  ].join('\n');

  var payload = {
    system_instruction: { parts: [{ text: systemPrompt }] },
    contents          : messages,
    tools             : [{ functionDeclarations: tools }],
    generationConfig  : {
      temperature    : 0.2,
      maxOutputTokens: 2048,
    },
  };

  try {
    var res = UrlFetchApp.fetch(url, {
      method         : 'post',
      contentType    : 'application/json',
      payload        : JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    var code = res.getResponseCode();
    if (code !== 200) {
      Logger.log('[CHAT v2] HTTP ' + code + ': ' + res.getContentText().substring(0,300));
      return null;
    }
    return JSON.parse(res.getContentText());
  } catch(e) {
    Logger.log('[CHAT v2] API error: ' + e.message);
    return null;
  }
}

// ============================================================
// Function定義（Geminiに渡すスキーマ）
// ============================================================
function _buildFunctionDefinitions() {
  return [
    {
      name       : 'search_cases',
      description: '案件・見積書・注文書をキーワードで検索する。顧客名、見積番号、注文番号、品名、機種コードなどで検索可能。',
      parameters : {
        type      : 'object',
        properties: {
          keyword : { type: 'string', description: '検索キーワード（顧客名・番号・品名など）' },
          status  : { type: 'string', description: 'ステータスフィルタ（送信済み/受領/受注済み/納品済み）' },
          doc_type: { type: 'string', description: '書類種別（quote=見積書 / order=注文書 / all=両方）' },
        },
        required: ['keyword'],
      },
    },
    {
      name       : 'get_case_detail',
      description: '特定の案件の詳細情報（明細行・単価・PDF URLなど）を取得する。',
      parameters : {
        type      : 'object',
        properties: {
          mgmt_id : { type: 'string', description: '管理ID（QM-...形式）' },
          quote_no: { type: 'string', description: '見積番号（mgmt_idがない場合）' },
          order_no: { type: 'string', description: '注文番号（mgmt_idがない場合）' },
        },
      },
    },
    {
      name       : 'compare_prices',
      description: '見積書と注文書の単価を行ごとに比較し、差異を検出する。紐づけ済みの案件に対して実行する。',
      parameters : {
        type      : 'object',
        properties: {
          mgmt_id: { type: 'string', description: '管理ID（QM-...形式）' },
        },
        required: ['mgmt_id'],
      },
    },
    {
      name       : 'update_status',
      description: '案件のステータスを更新する。受注登録・納品完了などに使用。必ずユーザーに確認してから実行する。',
      parameters : {
        type      : 'object',
        properties: {
          mgmt_id   : { type: 'string', description: '管理ID（QM-...形式）' },
          new_status: { type: 'string', description: '新しいステータス（送信済み/受領/受注済み/納品済み/キャンセル）' },
        },
        required: ['mgmt_id', 'new_status'],
      },
    },
    {
      name       : 'link_quote_order',
      description: '見積書と注文書を手動で紐づける。必ずユーザーに確認してから実行する。',
      parameters : {
        type      : 'object',
        properties: {
          order_mgmt_id: { type: 'string', description: '注文書の管理ID' },
          quote_mgmt_id: { type: 'string', description: '見積書の管理ID' },
        },
        required: ['order_mgmt_id', 'quote_mgmt_id'],
      },
    },
    {
      name       : 'search_unit_price',
      description: '過去の見積書から品名の単価履歴を検索する。「この品名の過去の単価は？」などに使用。',
      parameters : {
        type      : 'object',
        properties: {
          item_name: { type: 'string', description: '品名' },
          client   : { type: 'string', description: '顧客名（省略可）' },
        },
        required: ['item_name'],
      },
    },
    {
      name       : 'get_unlinked_orders',
      description: '見積書未紐づけの注文書一覧を取得する。「紐づけできていない注文書は？」などに使用。',
      parameters : {
        type      : 'object',
        properties: {
          limit: { type: 'number', description: '取得件数（デフォルト10）' },
        },
      },
    },
    {
      name       : 'get_ocr_log',
      description: 'OCR処理ログを確認する。「OCRが失敗しているファイルは？」などに使用。',
      parameters : {
        type      : 'object',
        properties: {
          status: { type: 'string', description: 'フィルタ（success/ocr_failed/error/all）' },
          limit : { type: 'number', description: '取得件数（デフォルト20）' },
        },
      },
    },
    {
      name       : 'get_summary_stats',
      description: '案件の集計情報を取得する。「今月の受注は？」「未紐づけは何件？」などに使用。',
      parameters : {
        type      : 'object',
        properties: {
          period: { type: 'string', description: '期間（today/this_month/last_month/all）' },
        },
      },
    },
  ];
}

// ============================================================
// Function実行（Geminiが選択した関数を実際に動かす）
// ============================================================
function _executeFunction(name, args) {
  try {
    switch(name) {

      case 'search_cases':
        return _fn_searchCases(args);

      case 'get_case_detail':
        return _fn_getCaseDetail(args);

      case 'compare_prices':
        return _fn_comparePrices(args);

      case 'update_status':
        return _fn_updateStatus(args);

      case 'link_quote_order':
        return _fn_linkQuoteOrder(args);

      case 'search_unit_price':
        return _fn_searchUnitPrice(args);

      case 'get_unlinked_orders':
        return _fn_getUnlinkedOrders(args);

      case 'get_ocr_log':
        return _fn_getOcrLog(args);

      case 'get_summary_stats':
        return _fn_getSummaryStats(args);

      default:
        return { success: false, error: '不明な関数: ' + name };
    }
  } catch(e) {
    Logger.log('[CHAT v2 FN] ' + name + ' error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// 各Function の実装
// ============================================================

function _fn_searchCases(args) {
  var keyword  = String(args.keyword  || '').trim();
  var status   = String(args.status   || '').trim();
  var docType  = String(args.doc_type || 'all').trim();

  // 案件サマリDBがあれば優先使用（11_db_normalize.gs）
  var ss = getSpreadsheet();
  var summarySheet = ss.getSheetByName('案件サマリ');
  if (summarySheet && summarySheet.getLastRow() > 1) {
    var kws = keyword.split(/\s+/).filter(Boolean);
    var results = searchCaseSummary(kws);
    if (status) results = results.filter(function(r){ return r.status === status; });
    if (docType === 'quote') results = results.filter(function(r){ return r.quoteNo; });
    if (docType === 'order') results = results.filter(function(r){ return r.orderNo; });
    return {
      success: true,
      count  : results.length,
      items  : results.slice(0, 20).map(function(r) {
        return {
          mgmtId     : r.mgmtId,
          quoteNo    : r.quoteNo,
          orderNo    : r.orderNo,
          client     : r.client,
          subject    : r.subject,
          status     : r.status,
          quoteAmount: r.quoteAmount,
          orderAmount: r.orderAmount,
          quoteDate  : r.quoteDate,
          orderDate  : r.orderDate,
          quotePdf   : r.quotePdf,
          orderPdf   : r.orderPdf,
          itemNames  : r.itemNames,
        };
      }),
    };
  }

  // フォールバック: 管理シート直接検索
  var mgmtData = getAllMgmtData();
  var kw = keyword.toLowerCase();
  var found = mgmtData.filter(function(row) {
    var text = [
      row[MGMT_COLS.QUOTE_NO - 1], row[MGMT_COLS.ORDER_NO - 1],
      row[MGMT_COLS.SUBJECT - 1],  row[MGMT_COLS.CLIENT - 1],
      row[MGMT_COLS.MODEL_CODE - 1],
    ].join(' ').toLowerCase();
    var statusMatch = !status || String(row[MGMT_COLS.STATUS - 1]) === status;
    return text.indexOf(kw) >= 0 && statusMatch;
  }).slice(0, 20).map(_rowToObject);

  return { success: true, count: found.length, items: found };
}

function _fn_getCaseDetail(args) {
  var ss        = getSpreadsheet();
  var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var mgmtData  = getAllMgmtData();
  var targetRow = null;

  if (args.mgmt_id) {
    targetRow = mgmtData.find(function(r){ return String(r[MGMT_COLS.ID-1]) === args.mgmt_id; });
  } else if (args.quote_no) {
    targetRow = mgmtData.find(function(r){ return String(r[MGMT_COLS.QUOTE_NO-1]) === args.quote_no; });
  } else if (args.order_no) {
    targetRow = mgmtData.find(function(r){ return String(r[MGMT_COLS.ORDER_NO-1]) === args.order_no; });
  }

  if (!targetRow) return { success: false, error: '案件が見つかりません' };

  var mgmtId = String(targetRow[MGMT_COLS.ID-1]);
  var obj    = _rowToObject(targetRow);

  // 見積明細
  var qs = ss.getSheetByName(CONFIG.SHEET_QUOTES);
  var quoteLines = [];
  if (qs && qs.getLastRow() > 1) {
    quoteLines = qs.getRange(2,1,qs.getLastRow()-1,15).getValues()
      .filter(function(r){ return String(r[0]) === mgmtId; })
      .map(function(r){ return {
        lineNo: r[5], itemName: r[6], spec: r[7],
        qty: r[8], unit: r[9], unitPrice: r[10], amount: r[11],
      }; });
  }

  // 注文明細
  var os = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  var orderLines = [];
  if (os && os.getLastRow() > 1) {
    orderLines = os.getRange(2,1,os.getLastRow()-1,19).getValues()
      .filter(function(r){ return String(r[0]) === mgmtId; })
      .map(function(r){ return {
        lineNo: r[7], itemName: r[8], spec: r[9],
        qty: r[12], unit: r[13], unitPrice: r[14], amount: r[15],
      }; });
  }

  return {
    success    : true,
    case       : obj,
    quoteLines : quoteLines,
    orderLines : orderLines,
  };
}

function _fn_comparePrices(args) {
  // 12_price_compare_v2.gs の関数を呼び出す
  if (typeof comparePriceByMgmtId !== 'function') {
    return { success: false, error: '単価比較エンジン（12_price_compare_v2.gs）が見つかりません' };
  }
  var result = comparePriceByMgmtId(args.mgmt_id);
  return {
    success     : true,
    status      : result.status,
    canAutoOrder: result.canAutoOrder,
    quoteNo     : result.quoteNo,
    orderNo     : result.orderNo,
    client      : result.client,
    lineResults : (result.lineResults || []).map(function(l) {
      return {
        lineNo        : l.lineNo,
        itemName      : l.itemName,
        matchStatus   : l.matchStatus,
        orderUnitPrice: l.orderUnitPrice,
        quoteUnitPrice: l.quoteUnitPrice,
        message       : l.message,
      };
    }),
    summary: result.summary,
  };
}

function _fn_updateStatus(args) {
  var valid = ['送信済み','受領','受注済み','納品済み','キャンセル','作成予定'];
  if (valid.indexOf(args.new_status) < 0) {
    return { success: false, error: '無効なステータス: ' + args.new_status };
  }
  var ss        = getSpreadsheet();
  var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var rowIdx    = _getMgmtRowIndex(mgmtSheet, args.mgmt_id);
  if (rowIdx < 0) return { success: false, error: '管理IDが見つかりません: ' + args.mgmt_id };

  var oldStatus = mgmtSheet.getRange(rowIdx, MGMT_COLS.STATUS).getValue();
  mgmtSheet.getRange(rowIdx, MGMT_COLS.STATUS    ).setValue(args.new_status);
  mgmtSheet.getRange(rowIdx, MGMT_COLS.UPDATED_AT).setValue(nowJST());

  Logger.log('[CHAT v2] ステータス更新: ' + args.mgmt_id + ' ' + oldStatus + ' → ' + args.new_status);
  return {
    success   : true,
    mgmtId    : args.mgmt_id,
    oldStatus : oldStatus,
    newStatus : args.new_status,
    updatedAt : nowJST(),
  };
}

function _fn_linkQuoteOrder(args) {
  // 00 missing utils.gs の confirmManualLink を呼び出す
  if (typeof confirmManualLink !== 'function') {
    return { success: false, error: 'confirmManualLink関数が見つかりません' };
  }
  var result = confirmManualLink(args.order_mgmt_id, args.quote_mgmt_id);
  return result;
}

function _fn_searchUnitPrice(args) {
  // 11_db_normalize.gs の searchUnitPrice を呼び出す
  if (typeof searchUnitPrice !== 'function') {
    return { success: false, error: '単価マスタ（11_db_normalize.gs）が見つかりません' };
  }
  var results = searchUnitPrice(args.item_name, '', args.client || '');
  return {
    success: true,
    count  : results.length,
    items  : results.slice(0, 10),
  };
}

function _fn_getUnlinkedOrders(args) {
  var limit   = Number(args.limit || 10);
  var mgmtData = getAllMgmtData();
  var unlinked = mgmtData.filter(function(row) {
    var orderNo = String(row[MGMT_COLS.ORDER_NO - 1] || '').trim();
    var quoteNo = String(row[MGMT_COLS.QUOTE_NO - 1] || '').trim();
    var linked  = _isLinkedVal(row[MGMT_COLS.LINKED  - 1]);
    return orderNo && !quoteNo && !linked;
  }).slice(0, limit).map(_rowToObject);

  return {
    success: true,
    count  : unlinked.length,
    items  : unlinked,
  };
}

function _fn_getOcrLog(args) {
  var status = String(args.status || 'all');
  var limit  = Number(args.limit  || 20);
  var ss     = getSpreadsheet();
  var sheet  = ss.getSheetByName('OCR処理ログ');
  if (!sheet || sheet.getLastRow() <= 1) {
    return { success: true, count: 0, items: [], message: 'OCR処理ログがまだありません' };
  }
  var data  = sheet.getRange(2, 1, sheet.getLastRow()-1, 5).getValues();
  var items = data.map(function(r) {
    return { date: r[0], fileName: r[1], status: r[2], mgmtId: r[3], detail: r[4] };
  });
  if (status !== 'all') {
    items = items.filter(function(r){ return r.status === status; });
  }
  // 新しい順
  items.reverse();
  return { success: true, count: items.length, items: items.slice(0, limit) };
}

function _fn_getSummaryStats(args) {
  var period   = String(args.period || 'this_month');
  var today    = new Date();
  var ym       = today.getFullYear() + '/' + String(today.getMonth()+1).padStart(2,'0');
  var lastYm   = new Date(today.getFullYear(), today.getMonth()-1, 1);
  var lastYmStr = lastYm.getFullYear() + '/' + String(lastYm.getMonth()+1).padStart(2,'0');

  var mgmtData  = getAllMgmtData();
  var allObjects = mgmtData.map(_rowToObject);

  var filtered = allObjects;
  if (period === 'this_month') {
    filtered = allObjects.filter(function(r){
      return String(r.orderDate||r.quoteDate||'').startsWith(ym);
    });
  } else if (period === 'last_month') {
    filtered = allObjects.filter(function(r){
      return String(r.orderDate||r.quoteDate||'').startsWith(lastYmStr);
    });
  } else if (period === 'today') {
    var todayStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');
    filtered = allObjects.filter(function(r){
      return String(r.orderDate||r.quoteDate||'').startsWith(todayStr);
    });
  }

  var unlinked  = allObjects.filter(function(r){ return r.orderNo && !r.quoteNo && !r.linked; });
  var totalAmt  = filtered.reduce(function(s,r){ return s+(Number(r.orderAmount)||0); }, 0);

  var byStatus = {};
  filtered.forEach(function(r){
    byStatus[r.status] = (byStatus[r.status]||0) + 1;
  });

  return {
    success        : true,
    period         : period,
    totalCases     : filtered.length,
    totalAmount    : totalAmt,
    byStatus       : byStatus,
    unlinkedCount  : unlinked.length,
  };
}

// ============================================================
// メッセージ構築
// ============================================================
function _buildInitialMessages(question, history) {
  var messages = [];

  // 会話履歴を追加
  history.forEach(function(h) {
    messages.push({
      role : h.role === 'assistant' ? 'model' : 'user',
      parts: [{ text: h.content }],
    });
  });

  // 今回の質問
  messages.push({ role: 'user', parts: [{ text: question }] });
  return messages;
}
