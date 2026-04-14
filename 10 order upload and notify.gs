/**
 * Gemini APIを用いた高精度マッチング機能 (明細レベル対応)
 */

const MATCHING_THRESHOLD_AUTO = 80;
const MATCHING_THRESHOLD_CANDIDATE = 50;

/**
 * 注文書受領時：見積書を明細単位でAI紐付け
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

  // 明細データの取得
  const orderLines = _getOrderLines(orderMgmtId);
  if (orderLines.length === 0) return { success: false, error: '注文明細が見つかりません' };
  
  // マッチング候補の見積書を抽出
  const quoteGroups = _buildQuoteGroups(mgmtData, headers);
  const quoteLines = _getAllQuoteLines();
  
  // Gemini APIによる明細レベルのマッチング推論
  const aiResult = matchWithGeminiAPI_LineLevel(orderData, orderLines, quoteGroups, quoteLines);
  
  if (!aiResult || !aiResult.matches) {
    return { success: false, error: 'AI解析に失敗しました' };
  }
  
  // 明細単位の適用
  _applyOrderLinks_LineLevel(orderMgmtId, aiResult.matches);
  
  return {
    success: true,
    status: aiResult.isMixed ? 'mixed_linked' : 'auto_linked',
    matches: aiResult.matches,
    reason: '明細レベルでのマッチングを完了しました'
  };
}

/**
 * 見積書アップロード時：注文書をドキュメント単位で探索（AI判定）
 */
function aiLinkQuoteToOrder(quoteMgmtId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  const mgmtData = mgmtSheet.getDataRange().getValues();
  const headers = mgmtData[0];
  
  const quoteIdx = mgmtData.findIndex(r => r[0] === quoteMgmtId);
  if (quoteIdx === -1) return { success: false, error: '見積データが見つかりません' };
  
  const quoteRow = mgmtData[quoteIdx];
  const quoteData = {};
  headers.forEach((h, i) => quoteData[h] = quoteRow[i]);
  
  // マッチング候補の注文書を抽出
  const orderGroups = _buildOrderGroups(mgmtData, headers);
  
  // Gemini APIによるマッチング推論
  const aiResult = matchQuoteToOrderWithGemini(quoteData, orderGroups);
  
  if (!aiResult || !aiResult.matches) {
    return { success: false, error: 'AI解析に失敗しました' };
  }
  
  const matches = aiResult.matches.sort((a, b) => b.score - a.score);
  const bestMatch = matches[0];
  
  if (bestMatch && bestMatch.score >= MATCHING_THRESHOLD_AUTO) {
    // 1件のみ紐付け適用
    _applyOrderLink_DocLevel(bestMatch.orderMgmtId, quoteMgmtId);
    return { success: true, status: 'auto_linked', bestMatch: bestMatch };
  }
  
  return { success: true, status: 'no_match', candidates: matches };
}

// ============================================================
// Gemini API 推論ロジック
// ============================================================

function matchWithGeminiAPI_LineLevel(orderData, orderLines, quoteGroups, quoteLines) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません');
  
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + apiKey;
  
  const quoteDetailSummary = quoteGroups.map(q => {
    const details = quoteLines.filter(ql => ql['管理ID'] === q.mgmtId);
    return {
      mgmtId: q.mgmtId,
      quoteNo: q.quoteNo,
      issueDate: q.issueDate || q['発行日'] || '',
      client: q.client || q['顧客名'] || '',
      totalAmount: q.totalAmount || q['見積金額'] || '',
      items: details.map(d => {
        const up = d['単価'] || d['単価（円）'] || '';
        const amt = d['金額'] || '';
        return `${d['品名']} / ${d['仕様']} / 数量:${d['数量']} / 単価:${up} / 金額:${amt}`;
      }).join('; ')
    };
  });

  const prompt = `
あなたは優秀な営業事務アシスタントです。注文書の各明細に最適な見積書（管理ID）を特定してください。
1つの注文に異なる見積書の内容が混在している場合があります。

【注文書】番号:${orderData['注文書番号']} | 顧客:${orderData['顧客名']} | 発注日:${orderData['発注日']||''}
【注文明細】
${orderLines.map(l => `行:${l.lineNo} | 品名:${l['品名']} | 仕様:${l['仕様']} | 数量:${l['数量']} | 単価:${l['単価']||''} | 金額:${l['金額']||''}`).join('\n')}

【見積書候補】（発行日・単価が注文書と近いものを優先してください）
${quoteDetailSummary.map(q => `ID:${q.mgmtId} | 見積番号:${q.quoteNo} | 顧客:${q.client} | 発行日:${q.issueDate} | 合計:${q.totalAmount} | 明細:[${q.items}]`).join('\n')}

マッチング判断基準（優先順位順）：
1. 品名・仕様の一致度（最重要）
2. 単価・金額の一致度
3. 発行日が発注日より前であること
4. 顧客名の一致

【返却形式】 JSONのみ。
{
  "isMixed": boolean,
  "matches": [
    { "orderLineNo": 注文明細の行番号, "quoteMgmtId": "管理ID", "quoteNo": "見積番号", "score": 0-100, "reason": "理由（品名/単価/発行日の一致根拠を明記）" }
  ]
}
`;

  const response = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { responseMimeType: "application/json" } }),
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) return null;
  try {
    return JSON.parse(JSON.parse(response.getContentText()).candidates[0].content.parts[0].text);
  } catch (e) { return null; }
}

function matchQuoteToOrderWithGemini(quoteData, orderGroups) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return null;
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + apiKey;
  const prompt = `見積書に最適な注文書を選んでください。\n見積:${quoteData['見積書番号']}|${quoteData['顧客名']}\n注文候補:${orderGroups.map(o => `ID:${o.mgmtId}|番号:${o.orderNo}`).join('\n')}\nJSONのみ:{"matches":[{"orderMgmtId":"ID","score":0-100}]}`;
  const response = UrlFetchApp.fetch(url, { method:'post', contentType:'application/json', payload: JSON.stringify({ contents:[{parts:[{text:prompt}]}], generationConfig:{responseMimeType:"application/json"} }), muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) return null;
  try { return JSON.parse(JSON.parse(response.getContentText()).candidates[0].content.parts[0].text); } catch(e){ return null; }
}

// ============================================================
// データ反映ロジック
// ============================================================

function _applyOrderLinks_LineLevel(orderMgmtId, matches) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  const orderSheet = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  const mgmtData = mgmtSheet.getDataRange().getValues();
  const orderLineData = orderSheet.getDataRange().getValues();
  
  const linkedQuoteIds = new Set();
  
  matches.forEach(m => {
    if (m.quoteMgmtId && m.score >= 50) {
      linkedQuoteIds.add(m.quoteMgmtId);
      for (let i = 1; i < orderLineData.length; i++) {
        if (orderLineData[i][0] === orderMgmtId && (orderLineData[i][7] === m.orderLineNo || i === m.orderLineNo)) {
          orderSheet.getRange(i + 1, 3).setValue(m.quoteNo);
        }
      }
      _updateQuoteSideLink(m.quoteMgmtId, orderMgmtId);
    }
  });
  
  const mid = orderMgmtId;
  const rowIdx = mgmtData.findIndex(r => r[0] === mid);
  if (rowIdx !== -1) {
    const qIds = Array.from(linkedQuoteIds);
    if (qIds.length > 1) {
      mgmtSheet.getRange(rowIdx+1, MGMT_COLS.QUOTE_NO).setValue('(複数)');
      mgmtSheet.getRange(rowIdx+1, MGMT_COLS.LINKED).setValue('TRUE');
    } else if (qIds.length === 1) {
      const qRow = mgmtData.find(r => r[0] === qIds[0]);
      if (qRow) {
        mgmtSheet.getRange(rowIdx+1, MGMT_COLS.QUOTE_NO).setValue(qRow[MGMT_COLS.QUOTE_NO-1]);
        mgmtSheet.getRange(rowIdx+1, MGMT_COLS.QUOTE_PDF_URL).setValue(qRow[MGMT_COLS.QUOTE_PDF_URL-1]);
        mgmtSheet.getRange(rowIdx+1, MGMT_COLS.LINKED).setValue('TRUE');
      }
    }
    mgmtSheet.getRange(rowIdx+1, MGMT_COLS.STATUS).setValue(CONFIG.STATUS.RECEIVED);
    mgmtSheet.getRange(rowIdx+1, MGMT_COLS.UPDATED_AT).setValue(nowJST());
  }
}

function _applyOrderLink_DocLevel(orderMgmtId, quoteMgmtId) {
  _applyOrderLinks_LineLevel(orderMgmtId, [{ orderLineNo: null, quoteMgmtId: quoteMgmtId, score: 100 }]);
}

function _updateQuoteSideLink(quoteMgmtId, orderMgmtId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  const data = mgmtSheet.getDataRange().getValues();
  const qIdx = data.findIndex(r => r[0] === quoteMgmtId);
  const oIdx = data.findIndex(r => r[0] === orderMgmtId);
  if (qIdx !== -1 && oIdx !== -1) {
    mgmtSheet.getRange(qIdx+1, MGMT_COLS.ORDER_NO).setValue(data[oIdx][MGMT_COLS.ORDER_NO-1]);
    mgmtSheet.getRange(qIdx+1, MGMT_COLS.ORDER_PDF_URL).setValue(data[oIdx][MGMT_COLS.ORDER_PDF_URL-1]);
    mgmtSheet.getRange(qIdx+1, MGMT_COLS.LINKED).setValue('TRUE');
    mgmtSheet.getRange(qIdx+1, MGMT_COLS.STATUS).setValue(CONFIG.STATUS.RECEIVED);
  }
}

// ============================================================
// 共通ユーティリティ
// ============================================================

function _getOrderLines(mgmtId) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ORDERS).getDataRange().getValues();
  const head = data[0];
  return data.slice(1).filter(r => r[0] === mgmtId).map((r, i) => {
    const o = { lineNo: i+1 };
    head.forEach((h, j) => o[h] = r[j]);
    return o;
  });
}

function _getAllQuoteLines() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_QUOTES).getDataRange().getValues();
  const head = data[0];
  return data.slice(1).map(r => { const o = {}; head.forEach((h, j) => o[h]=r[j]); return o; });
}

function _buildQuoteGroups(mgmtData, headers) {
  const qNoIdx = headers.indexOf('見積書番号');
  return mgmtData.slice(1).filter(r => r[qNoIdx]).map(r => ({ mgmtId: r[0], quoteNo: r[qNoIdx], modelCode: r[headers.indexOf('機種コード')], client: r[headers.indexOf('顧客名')], amount: r[headers.indexOf('枚数')], subject: r[headers.indexOf('件名')] }));
}

function _buildOrderGroups(mgmtData, headers) {
  return mgmtData.slice(1).filter(r => String(r[0]).startsWith('MO')).map(r => ({ mgmtId: r[0], orderNo: r[headers.indexOf('注文書番号')] }));
}

function _sendChatNotification(mgmtId, docType, actionType) {
  const webhookUrl = _getChatWebhookUrl();
  if (!webhookUrl) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const row = data.find(r => String(r[0]) === String(mgmtId));
  if (!row) return;
  
  const getVal = (h) => row[headers.indexOf(h)] || 'なし';
  const action = actionType || 'new';

  let title = docType === 'quote' ? "📄 見積書を登録" : "📦 注文書を受領";
  let icon = "🔹";
  
  if (action === 'revision') {
    title = "🔄 注文書の差し替え受領";
    icon = "⚠️";
  } else if (action === 'cancellation') {
    title = "❌ 注文書のキャンセル";
    icon = "🚫";
  }

  let text = `【${title}】\n`;
  text += `━━━━━━━━━━━━━━\n`;
  text += `${icon} 案件: ${getVal('件名')}\n`;
  text += `${icon} 顧客: ${getVal('顧客名')}\n`;
  
  if (action !== 'cancellation') {
    text += `${icon} 金額: ¥${Number(getVal(docType==='quote'?'見積金額':'注文金額')).toLocaleString()}\n`;
    text += `${icon} URL: ${getVal(docType==='quote'?'見積書PDF':'注文書PDF')}\n`;
  } else {
    text += `🚨 この注文はキャンセルとして処理されました。\n`;
    text += `理由: ${getVal('備考')}\n`;
  }
  
  if (getVal('紐付け済み') === 'TRUE' && action !== 'cancellation') {
    text += `\n✅ AI紐付け完了: ${docType==='quote'?getVal('注文書番号'):getVal('見積書番号')}\n`;
  }
  
  text += `\n🌐 システムで確認: ${ScriptApp.getService().getUrl()}`;
  
  UrlFetchApp.fetch(webhookUrl, { 
    method: "post", 
    contentType: "application/json", 
    payload: JSON.stringify({ "text": text }), 
    muteHttpExceptions: true 
  });
}

function _getChatWebhookUrl() {
  return PropertiesService.getScriptProperties().getProperty('GOOGLE_CHAT_WEBHOOK_URL');
}
