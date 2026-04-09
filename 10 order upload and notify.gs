/**
 * Gemini APIを用いた高精度マッチング機能
 */

const MATCHING_THRESHOLD_AUTO = 80;
const MATCHING_THRESHOLD_CANDIDATE = 50;

/**
 * 注文書と見積書をGemini APIで紐付ける（AI判定）
 * @param {string} orderMgmtId 注文書の管理ID
 * @returns {Object} 判定結果
 */
function aiLinkOrderToQuote(orderMgmtId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mgmtSheet = ss.getSheetByName('案件管理');
  const mgmtData = mgmtSheet.getDataRange().getValues();
  const headers = mgmtData[0];
  
  const orderIdx = mgmtData.findIndex(r => r[0] === orderMgmtId);
  if (orderIdx === -1) return { success: false, error: '注文データが見つかりません' };
  
  const orderRow = mgmtData[orderIdx];
  const orderData = {};
  headers.forEach((h, i) => orderData[h] = orderRow[i]);
  
  // マッチング候補の見積書を抽出
  const quoteGroups = _buildQuoteGroups(mgmtData, headers);
  
  // Gemini APIによるマッチング推論
  const aiResult = matchWithGeminiAPI(orderData, quoteGroups);
  
  if (!aiResult || !aiResult.matches) {
    return { success: false, error: 'AI解析に失敗しました' };
  }
  
  // スコア順にソート
  const matches = aiResult.matches.sort((a, b) => b.score - a.score);
  const bestMatch = matches[0];
  
  let status = 'no_match';
  if (bestMatch && bestMatch.score >= MATCHING_THRESHOLD_AUTO) {
    // 自動紐付け実行
    _applyOrderLink(orderMgmtId, bestMatch.quoteMgmtId);
    status = 'auto_linked';
  } else if (bestMatch && bestMatch.score >= MATCHING_THRESHOLD_CANDIDATE) {
    status = 'candidates_found';
  }
  
  return {
    success: true,
    status: status,
    bestMatch: bestMatch,
    candidates: matches.filter(m => m.score >= MATCHING_THRESHOLD_CANDIDATE),
    reason: bestMatch ? bestMatch.reason : '候補が見つかりませんでした'
  };
}

/**
 * Gemini APIを使用して注文書に最適な見積書を選択する
 */
function matchWithGeminiAPI(orderData, quoteGroups) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません');
  
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + apiKey;
  
  const prompt = `
あなたは優秀な営業事務アシスタントです。アップロードされた「注文書」のデータと、システム内の「見積書候補」を比較し、最も可能性の高い紐付け先を判定してください。

【注文書データ】
番号: ${orderData['注文書番号'] || '不明'}
機種: ${orderData['機種コード'] || '不明'}
顧客: ${orderData['顧客名'] || '不明'}
金額: ${orderData['注文金額'] || '不明'}
件名: ${orderData['件名'] || '不明'}

【見積書候補リスト】
${quoteGroups.map(q => `ID:${q.mgmtId} | 番号:${q.quoteNo} | 機種:${q.modelCode} | 顧客:${q.client} | 金額:${q.amount} | 件名:${q.subject}`).join('\n')}

【判定ルール】
1. 注文書番号と見積書番号が関連しているか（例：見積枝番、注文書側の参照番号）。
2. 機種コードが一致、または酷似しているか。
3. 顧客名が一致しているか（株式会社の有無、屋号のみ等の表記ゆれを考慮）。
4. 金額が一致、または消費税(10%)の有無による差、OCR誤字による微差（1文字違い等）を許容。
5. 件名に含まれるキーワードが一致しているか。

【返却形式】
JSONのみで回答してください。
{
  "matches": [
    {
      "quoteMgmtId": "管理ID",
      "quoteNo": "見積番号",
      "score": 0から100の数値(確信度),
      "reason": "選定理由（簡潔に）"
    }
  ]
}
`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: "application/json"
    }
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const resCode = response.getResponseCode();
  const resText = response.getContentText();
  
  if (resCode !== 200) {
    Logger.log('Gemini Error: ' + resText);
    return null;
  }
  
  const json = JSON.parse(resText);
  try {
    return JSON.parse(json.candidates[0].content.parts[0].text);
  } catch (e) {
    Logger.log('JSON Parse Error: ' + e.message);
    return null;
  }
}

/**
 * 見積書候補のリストを整形
 */
function _buildQuoteGroups(mgmtData, headers) {
  const quotes = [];
  const qNoIdx = headers.indexOf('見積書番号');
  const mIdx = headers.indexOf('機種コード');
  const cIdx = headers.indexOf('顧客名');
  const aIdx = headers.indexOf('注文金額');
  const sIdx = headers.indexOf('件名');
  const lIdx = headers.indexOf('紐付け済み'); // 紐付け済みフラグがあれば考慮
  
  // データの正規化
  for (let i = 1; i < mgmtData.length; i++) {
    const row = mgmtData[i];
    // 既に紐付け済みのものは除外したいが、再判定のために含める
    if (row[qNoIdx]) {
      quotes.push({
        mgmtId: row[0],
        quoteNo: row[qNoIdx],
        modelCode: row[mIdx],
        client: row[cIdx],
        amount: row[aIdx],
        subject: row[sIdx]
      });
    }
  }
  return quotes;
}

/**
 * 紐付けをスプレッドシートに反映
 */
function _applyOrderLink(orderMgmtId, quoteMgmtId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  const last = sheet.getLastRow();
  if (last <= 1) return;
  
  const data = sheet.getRange(1, 1, last, 27).getValues();
  let orderRowIdx = -1;
  let quoteRowIdx = -1;
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][MGMT_COLS.ID - 1]) === orderMgmtId) orderRowIdx = i + 1;
    if (String(data[i][MGMT_COLS.ID - 1]) === quoteMgmtId) quoteRowIdx = i + 1;
  }
  
  if (orderRowIdx !== -1 && quoteRowIdx !== -1) {
    const quoteNo = data[quoteRowIdx - 1][MGMT_COLS.QUOTE_NO - 1];
    const quotePdfUrl = data[quoteRowIdx - 1][MGMT_COLS.QUOTE_PDF_URL - 1];
    const orderNo = data[orderRowIdx - 1][MGMT_COLS.ORDER_NO - 1];
    const orderPdfUrl = data[orderRowIdx - 1][MGMT_COLS.ORDER_PDF_URL - 1];
    const orderAmt = data[orderRowIdx - 1][MGMT_COLS.ORDER_AMOUNT - 1];
    const orderDt = data[orderRowIdx - 1][MGMT_COLS.ORDER_DATE - 1];

    // 注文書側の更新
    sheet.getRange(orderRowIdx, MGMT_COLS.QUOTE_NO).setValue(quoteNo);
    sheet.getRange(orderRowIdx, MGMT_COLS.QUOTE_PDF_URL).setValue(quotePdfUrl);
    sheet.getRange(orderRowIdx, MGMT_COLS.LINKED).setValue('TRUE');
    sheet.getRange(orderRowIdx, MGMT_COLS.STATUS).setValue(CONFIG.STATUS.RECEIVED);
    sheet.getRange(orderRowIdx, MGMT_COLS.UPDATED_AT).setValue(nowJST());

    // 見積書側の更新
    sheet.getRange(quoteRowIdx, MGMT_COLS.ORDER_NO).setValue(orderNo);
    sheet.getRange(quoteRowIdx, MGMT_COLS.ORDER_PDF_URL).setValue(orderPdfUrl);
    sheet.getRange(quoteRowIdx, MGMT_COLS.ORDER_AMOUNT).setValue(orderAmt);
    sheet.getRange(quoteRowIdx, MGMT_COLS.ORDER_DATE).setValue(orderDt);
    sheet.getRange(quoteRowIdx, MGMT_COLS.LINKED).setValue('TRUE');
    sheet.getRange(quoteRowIdx, MGMT_COLS.STATUS).setValue(CONFIG.STATUS.RECEIVED);
    sheet.getRange(quoteRowIdx, MGMT_COLS.UPDATED_AT).setValue(nowJST());
    
    // 注文書管理シート（ファイル依存分）の同期
    _syncDocSheetLink(orderMgmtId, quoteNo);
  }
}

/**
 * 注文書詳細シートの「見積番号(紐づけ)」列を更新
 */
function _syncDocSheetLink(orderMgmtId, quoteNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const os = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  if (!os || os.getLastRow() <= 1) return;
  const ids = os.getRange(2, 1, os.getLastRow() - 1, 1).getValues().flat();
  ids.forEach((id, i) => {
    if (String(id) === orderMgmtId) {
      os.getRange(i + 2, 3).setValue(quoteNo); // INDEX 3 = 見積番号(紐づけ)
    }
  });
}

/**
 * テスト用関数（GASエディタから実行可能）
 */
function testAiLinkOrder() {
  // 存在する適当な注文管理IDを指定してください
  const testId = 'MO-20260401-001'; 
  const result = aiLinkOrderToQuote(testId);
  Logger.log(JSON.stringify(result, null, 2));
}
