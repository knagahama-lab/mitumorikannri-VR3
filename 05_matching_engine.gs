// ============================================================
// 05_matching_engine.gs
// 注文書 ↔ 見積書 紐づけエンジン（統合版）
//
// 統合元:
//   - 05 matching engine.gs   （バッチ司令塔）
//   - 10 order link engine.gs  （スコアリング・API・通知）
//   - 10 order upload and notify.gs （Gemini AI マッチング）
// ============================================================

// ============================================================
// 設定（重複定数を統一）
// ============================================================
var MATCH_CONFIG = {
  AUTO_LINK_THRESHOLD:   82,   // 自動紐づけスコア閾値
  CANDIDATE_THRESHOLD:   50,   // 候補として表示するスコア最低値
  MAX_CANDIDATES:         5,   // 最大候補数
  PRICE_TOLERANCE_PCT:   10,   // 金額一致と見なす許容誤差(%)
  SHEET_CANDIDATES: '紐づけ候補',
  MATCHING_API_KEY_NAME: 'GEMINI_API_KEY_MATCHING'
};

// ============================================================
// 公開 API
// ============================================================

/**
 * 注文書IDを受け取り、最適な見積書を1件紐付けるか候補を返す
 */
function matchOrderToQuote(orderMgmtId) {
  var result = aiLinkOrderToQuote(orderMgmtId);
  if (result.success) {
    if (result.status === 'auto_linked') {
      return {
        success: true,
        status: 'auto_linked',
        quoteMgmtId: result.bestMatch ? result.bestMatch.quoteMgmtId : null,
        quoteNo:     result.bestMatch ? result.bestMatch.quoteNo     : null,
        score:       result.bestMatch ? result.bestMatch.score       : null,
        reason:      result.bestMatch ? result.bestMatch.reason      : null
      };
    } else if (result.status === 'candidates_found') {
      return {
        success: true,
        status: 'candidates_found',
        candidates: (result.candidates || []).map(function(c) {
          return { quoteMgmtId: c.quoteMgmtId, quoteNo: c.quoteNo, score: c.score, reason: c.reason };
        })
      };
    }
  }
  return { success: true, status: 'no_match', candidates: [] };
}

/**
 * 紐づけられていないすべての注文書に対して一括マッチングを実行（司令塔）
 */
function runBatchMatching() {
  try {
    var ss = getSpreadsheet();
    var mgmtData = getAllMgmtData();
    var quotes   = _getAllQuotesForMatching(ss);

    var autoCount      = 0;
    var candidateCount = 0;
    var candidatesStore = [];

    mgmtData.forEach(function(r) {
      var mgmtId  = String(r[MGMT_COLS.ID - 1]      || '').trim();
      var orderNo = String(r[MGMT_COLS.ORDER_NO - 1] || '').trim();
      var quoteNo = String(r[MGMT_COLS.QUOTE_NO - 1] || '').trim();
      var isLinked = r[MGMT_COLS.LINKED - 1] === true || r[MGMT_COLS.LINKED - 1] === 'TRUE';

      if (!orderNo || quoteNo || isLinked) return;

      var orderInfo = _getOrderInfo(ss, mgmtId);
      if (!orderInfo) return;

      var scored = _scoreQuotesForOrder(orderInfo, quotes);
      if (!scored.length) return;

      var best = scored[0];
      if (best.score >= MATCH_CONFIG.AUTO_LINK_THRESHOLD) {
        try {
          _applyOrderQuoteLink(ss, mgmtId, best.quoteMgmtId, best.quoteNo);
          autoCount++;
          Logger.log('[BATCH AUTO LINK] ' + mgmtId + ' → ' + best.quoteNo + ' score:' + best.score);
        } catch (e) {
          Logger.log('[BATCH LINK ERROR] ' + e.message);
        }
      } else if (best.score >= MATCH_CONFIG.CANDIDATE_THRESHOLD) {
        candidatesStore.push({
          orderMgmtId: mgmtId,
          orderNo:     orderNo,
          orderClient: orderInfo.client,
          orderDate:   orderInfo.orderDate,
          orderAmount: orderInfo.orderAmount,
          orderPdfUrl: orderInfo.pdfUrl,
          candidates:  scored.filter(function(c) { return c.score >= MATCH_CONFIG.CANDIDATE_THRESHOLD; })
        });
        candidateCount++;
      }
    });

    PropertiesService.getScriptProperties().setProperty('AI_MATCHING_CANDIDATES', JSON.stringify(candidatesStore));
    Logger.log('[BATCH MATCH] 自動確定:' + autoCount + ' 候補抽出:' + candidateCount);
    return { autoCount: autoCount, candidateCount: candidateCount };

  } catch (e) {
    Logger.log('runBatchMatching エラー: ' + e.message);
    throw e;
  }
}

/** 定期実行トリガー用 */
function autoMatchNewOrders() {
  try {
    Logger.log('定期実行: AI自動マッチングを開始します...');
    var result = runBatchMatching();
    Logger.log('定期実行完了: 自動確定 ' + result.autoCount + '件, 候補抽出 ' + result.candidateCount + '件');
  } catch (e) {
    Logger.log('定期実行エラー: ' + e.message);
  }
}

// ============================================================
// API エンドポイント（handleApiRequest から呼ぶ）
// ============================================================

/** 注文書に対する見積書候補を返す */
function _apiGetOrderLinkCandidates(p) {
  try {
    var mgmtId = String(p.mgmtId || '').trim();
    if (!mgmtId) return { success: false, error: '管理IDが必要です' };

    var ss        = getSpreadsheet();
    var orderInfo = _getOrderInfo(ss, mgmtId);
    if (!orderInfo) return { success: true, candidates: [], linkedQuoteNo: null };

    var linkedInfo = orderInfo.quoteNo ? _getLinkedQuoteInfo(ss, orderInfo.quoteNo, mgmtId) : null;

    var quotes     = _getAllQuotesForMatching(ss);
    var scored     = _scoreQuotesForOrder(orderInfo, quotes);
    var candidates = scored
      .filter(function(c) { return c.score >= MATCH_CONFIG.CANDIDATE_THRESHOLD; })
      .slice(0, MATCH_CONFIG.MAX_CANDIDATES);

    // 高スコアなら自動確定
    if (candidates.length > 0 && candidates[0].score >= MATCH_CONFIG.AUTO_LINK_THRESHOLD && !orderInfo.quoteNo) {
      try {
        _applyOrderQuoteLink(ss, mgmtId, candidates[0].quoteMgmtId, candidates[0].quoteNo);
      } catch (e) {
        Logger.log('[AUTO LINK ERROR] ' + e.message);
      }
    }

    return {
      success:       true,
      candidates:    candidates,
      linkedQuoteNo: linkedInfo ? linkedInfo.quoteNo    : (orderInfo.quoteNo || null),
      linkedMgmtId:  linkedInfo ? linkedInfo.mgmtId     : null,
      linkedAmount:  linkedInfo ? linkedInfo.quoteAmount : null,
      linkedPdfUrl:  linkedInfo ? linkedInfo.quotePdfUrl : null
    };
  } catch (e) {
    Logger.log('[GET CANDIDATES ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

/** 手動で紐づけを確定する */
function _apiConfirmOrderLink(p) {
  try {
    var orderMgmtId = String(p.orderMgmtId || '').trim();
    var quoteMgmtId = String(p.quoteMgmtId || '').trim();
    var quoteNo     = String(p.quoteNo     || '').trim();
    if (!orderMgmtId || !quoteMgmtId) return { success: false, error: '管理IDが必要です' };

    _applyOrderQuoteLink(getSpreadsheet(), orderMgmtId, quoteMgmtId, quoteNo);
    return { success: true };
  } catch (e) {
    Logger.log('[CONFIRM LINK ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// Gemini AI マッチング
// ============================================================

/** 注文書受領時：見積書を明細単位でAI紐付け */
function aiLinkOrderToQuote(orderMgmtId) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  if (!mgmtSheet) return { success: false, error: '管理シートが見つかりません: ' + CONFIG.SHEET_MANAGEMENT };

  var mgmtData = mgmtSheet.getDataRange().getValues();
  var headers  = mgmtData[0];
  var orderIdx = mgmtData.findIndex(function(r) { return r[0] === orderMgmtId; });
  if (orderIdx === -1) return { success: false, error: '注文データが見つかりません' };

  var orderRow  = mgmtData[orderIdx];
  var orderData = {};
  headers.forEach(function(h, i) { orderData[h] = orderRow[i]; });

  var orderLines = _getOrderLines(orderMgmtId);
  if (orderLines.length === 0) return { success: false, error: '注文明細が見つかりません' };

  var quoteGroups = _buildQuoteGroups(mgmtData, headers);
  var quoteLines  = _getAllQuoteLines();

  var aiResult = _matchWithGeminiLineLevel(orderData, orderLines, quoteGroups, quoteLines);
  if (!aiResult || !aiResult.matches) return { success: false, error: 'AI解析に失敗しました' };

  _applyOrderLinks_LineLevel(orderMgmtId, aiResult.matches);

  return {
    success: true,
    status:  aiResult.isMixed ? 'mixed_linked' : 'auto_linked',
    matches: aiResult.matches,
    reason:  '明細レベルでのマッチングを完了しました'
  };
}

/** 見積書アップロード時：注文書をドキュメント単位でAI探索 */
function aiLinkQuoteToOrder(quoteMgmtId) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var mgmtData  = mgmtSheet.getDataRange().getValues();
  var headers   = mgmtData[0];

  var quoteIdx = mgmtData.findIndex(function(r) { return r[0] === quoteMgmtId; });
  if (quoteIdx === -1) return { success: false, error: '見積データが見つかりません' };

  var quoteRow  = mgmtData[quoteIdx];
  var quoteData = {};
  headers.forEach(function(h, i) { quoteData[h] = quoteRow[i]; });

  var orderGroups = _buildOrderGroups(mgmtData, headers);
  var aiResult    = _matchQuoteToOrderWithGemini(quoteData, orderGroups);
  if (!aiResult || !aiResult.matches) return { success: false, error: 'AI解析に失敗しました' };

  var matches   = aiResult.matches.sort(function(a, b) { return b.score - a.score; });
  var bestMatch = matches[0];

  if (bestMatch && bestMatch.score >= MATCH_CONFIG.AUTO_LINK_THRESHOLD) {
    _applyOrderLink_DocLevel(bestMatch.orderMgmtId, quoteMgmtId);
    return { success: true, status: 'auto_linked', bestMatch: bestMatch };
  }
  return { success: true, status: 'no_match', candidates: matches };
}

// ============================================================
// Gemini API 推論（内部）
// ============================================================

function _matchWithGeminiLineLevel(orderData, orderLines, quoteGroups, quoteLines) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません');

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
            CONFIG.GEMINI_PRIMARY_MODEL + ':generateContent?key=' + apiKey;

  var quoteDetailSummary = quoteGroups.map(function(q) {
    var details = quoteLines.filter(function(ql) { return ql['管理ID'] === q.mgmtId; });
    return {
      mgmtId:      q.mgmtId,
      quoteNo:     q.quoteNo,
      issueDate:   q.issueDate || q['発行日'] || '',
      client:      q.client    || q['顧客名'] || '',
      totalAmount: q.totalAmount || q['見積金額'] || '',
      items: details.map(function(d) {
        return (d['品名'] || '') + ' / ' + (d['仕様'] || '') +
               ' / 数量:' + (d['数量'] || '') +
               ' / 単価:' + (d['単価'] || d['単価（円）'] || '') +
               ' / 金額:' + (d['金額'] || '');
      }).join('; ')
    };
  });

  var prompt =
    'あなたは優秀な営業事務アシスタントです。注文書の各明細に最適な見積書（管理ID）を特定してください。\n' +
    '1つの注文に異なる見積書の内容が混在している場合があります。\n\n' +
    '【注文書】番号:' + (orderData['注文書番号'] || '') +
    ' | 顧客:' + (orderData['顧客名'] || '') +
    ' | 発注日:' + (orderData['発注日'] || '') + '\n' +
    '【注文明細】\n' +
    orderLines.map(function(l) {
      return '行:' + l.lineNo + ' | 品名:' + (l['品名'] || '') + ' | 仕様:' + (l['仕様'] || '') +
             ' | 数量:' + (l['数量'] || '') + ' | 単価:' + (l['単価'] || '') + ' | 金額:' + (l['金額'] || '');
    }).join('\n') + '\n\n' +
    '【見積書候補】\n' +
    quoteDetailSummary.map(function(q) {
      return 'ID:' + q.mgmtId + ' | 見積番号:' + q.quoteNo +
             ' | 顧客:' + q.client + ' | 発行日:' + q.issueDate +
             ' | 合計:' + q.totalAmount + ' | 明細:[' + q.items + ']';
    }).join('\n') + '\n\n' +
    'マッチング判断基準（優先順位順）：\n' +
    '1. 品名・仕様の一致度（最重要）\n' +
    '2. 単価・金額の一致度\n' +
    '3. 発行日が発注日より前であること\n' +
    '4. 顧客名の一致\n\n' +
    '【返却形式】 JSONのみ。\n' +
    '{"isMixed":boolean,"matches":[{"orderLineNo":行番号,"quoteMgmtId":"管理ID","quoteNo":"見積番号","score":0-100,"reason":"理由"}]}';

  var response = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { responseMimeType: 'application/json' }
    }),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) return null;
  try {
    return JSON.parse(JSON.parse(response.getContentText()).candidates[0].content.parts[0].text);
  } catch (e) { return null; }
}

function _matchQuoteToOrderWithGemini(quoteData, orderGroups) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return null;

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
            CONFIG.GEMINI_PRIMARY_MODEL + ':generateContent?key=' + apiKey;

  var prompt = '見積書に最適な注文書を選んでください。\n' +
    '見積:' + (quoteData['見積番号'] || '') + '|' + (quoteData['顧客名'] || '') + '\n' +
    '注文候補:' + orderGroups.map(function(o) { return 'ID:' + o.mgmtId + '|番号:' + o.orderNo; }).join('\n') + '\n' +
    'JSONのみ:{"matches":[{"orderMgmtId":"管理ID","score":0-100}]}';

  var response = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { responseMimeType: 'application/json' }
    }),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) return null;
  try {
    return JSON.parse(JSON.parse(response.getContentText()).candidates[0].content.parts[0].text);
  } catch (e) { return null; }
}

// ============================================================
// スコアリングロジック
// ============================================================

function _scoreQuotesForOrder(order, quotes) {
  return quotes.map(function(q) {
    var score   = 0;
    var reasons = [];

    // ① 顧客名一致
    if (order.client && q.client) {
      var clientSim = _textSimilarity(
        _normalizeCompanyName(order.client),
        _normalizeCompanyName(q.client)
      );
      if (clientSim >= 0.9)      { score += 30; reasons.push('顧客名一致'); }
      else if (clientSim >= 0.6) { score += 15; reasons.push('顧客名部分一致'); }
    }

    // ② 機種コード一致
    if (order.modelCode && q.modelCode) {
      if (order.modelCode === q.modelCode) {
        score += 25; reasons.push('機種コード完全一致');
      } else if (order.modelCode.indexOf(q.modelCode) >= 0 || q.modelCode.indexOf(order.modelCode) >= 0) {
        score += 12; reasons.push('機種コード部分一致');
      }
    }

    // ③ 明細品名の一致
    var itemMatchScore = _scoreItemsMatch(order.items, q.items);
    if (itemMatchScore > 0) {
      score += itemMatchScore;
      reasons.push('明細品名一致×' + Math.round(itemMatchScore / 10));
    }

    // ④ 金額の一致
    if (order.orderAmount > 0 && q.quoteAmount > 0) {
      var pctDiff = Math.abs(order.orderAmount - q.quoteAmount) / order.orderAmount * 100;
      if      (pctDiff <= 1)                              { score += 20; reasons.push('金額完全一致'); }
      else if (pctDiff <= MATCH_CONFIG.PRICE_TOLERANCE_PCT) { score += 12; reasons.push('金額ほぼ一致'); }
      else if (pctDiff <= 20)                              { score += 4; }
    }

    // ⑤ 件名・キーワード一致
    if (order.subject && q.subject) {
      if (_textSimilarity(order.subject, q.subject) >= 0.5) { score += 8; reasons.push('件名類似'); }
    }

    // ⑥ 日付の近さ（注文日が見積日の30日以内）
    if (order.orderDate && q.issueDate) {
      try {
        var orderD  = new Date(order.orderDate.replace(/\//g, '-'));
        var quoteD  = new Date(q.issueDate.replace(/\//g, '-'));
        var dayDiff = (orderD - quoteD) / (1000 * 60 * 60 * 24);
        if (dayDiff >= 0 && dayDiff <= 30) { score += 6; reasons.push('発注日と見積日が近い'); }
        else if (dayDiff >= -7 && dayDiff < 0) { score += 3; }
      } catch (e2) {}
    }

    return {
      quoteMgmtId: q.quoteMgmtId,
      quoteNo:     q.quoteNo,
      client:      q.client,
      destCompany: q.destCompany,
      issueDate:   q.issueDate,
      quoteAmount: q.quoteAmount,
      quotePdfUrl: q.quotePdfUrl,
      subject:     q.subject,
      score:       Math.min(score, 100),
      reason:      reasons.slice(0, 3).join('・')
    };
  })
  .filter(function(c) { return c.score > 0; })
  .sort(function(a, b) { return b.score - a.score; });
}

function _scoreItemsMatch(orderItems, quoteItems) {
  if (!orderItems || !orderItems.length || !quoteItems || !quoteItems.length) return 0;
  var totalScore = 0;
  var matchCount = 0;

  orderItems.forEach(function(oi) {
    var oName   = _normItemName(oi.itemName);
    var bestSim = 0;
    var bestQItem = null;

    quoteItems.forEach(function(qi) {
      var sim = _textSimilarity(oName, _normItemName(qi.itemName));
      if (sim > bestSim) { bestSim = sim; bestQItem = qi; }
    });

    if (bestSim >= 0.85) {
      matchCount++;
      totalScore += 10;
      if (bestQItem && oi.unitPrice > 0 && bestQItem.unitPrice > 0) {
        var priceDiff = Math.abs(oi.unitPrice - bestQItem.unitPrice) / oi.unitPrice * 100;
        if (priceDiff <= 5) totalScore += 5;
      }
    } else if (bestSim >= 0.6) {
      matchCount++;
      totalScore += 5;
    }
  });

  if (orderItems.length > 0 && matchCount === orderItems.length) totalScore += 10;
  return Math.min(totalScore, 40);
}

// ============================================================
// データ反映ロジック
// ============================================================

/** 紐づけを管理シートに書き込む */
function _applyOrderQuoteLink(ss, orderMgmtId, quoteMgmtId, quoteNo) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var last  = sheet.getLastRow();
  if (last <= 1) throw new Error('管理シートにデータがありません');

  var data    = sheet.getRange(2, 1, last - 1, Object.keys(MGMT_COLS).length + 2).getValues();
  var orderNo = '';

  // 注文書行を更新
  data.forEach(function(row, i) {
    if (String(row[MGMT_COLS.ID - 1] || '').trim() !== orderMgmtId) return;
    var rowNum = i + 2;
    sheet.getRange(rowNum, MGMT_COLS.QUOTE_NO).setValue(quoteNo);
    sheet.getRange(rowNum, MGMT_COLS.LINKED).setValue('TRUE');
    sheet.getRange(rowNum, MGMT_COLS.STATUS).setValue('受領');
    sheet.getRange(rowNum, MGMT_COLS.UPDATED_AT).setValue(nowJST());
    orderNo = String(row[MGMT_COLS.ORDER_NO - 1] || '').trim();
  });

  // 見積書行を更新
  if (orderNo) {
    data.forEach(function(row, i) {
      if (String(row[MGMT_COLS.ID - 1] || '').trim() !== quoteMgmtId) return;
      var rowNum = i + 2;
      sheet.getRange(rowNum, MGMT_COLS.ORDER_NO).setValue(orderNo);
      sheet.getRange(rowNum, MGMT_COLS.LINKED).setValue('TRUE');
      sheet.getRange(rowNum, MGMT_COLS.UPDATED_AT).setValue(nowJST());
    });
  }

  // 「紐づけ候補」シートを処理済みにする
  try {
    var candSheet = ss.getSheetByName(MATCH_CONFIG.SHEET_CANDIDATES);
    if (candSheet && candSheet.getLastRow() > 1) {
      var candData = candSheet.getRange(2, 1, candSheet.getLastRow() - 1, 16).getValues();
      candData.forEach(function(r, i) {
        if (String(r[0]) === orderMgmtId) {
          candSheet.getRange(i + 2, 16).setValue('手動確定済み');
          candSheet.getRange(i + 2, 17).setValue(nowJST());
        }
      });
    }
  } catch (e2) {}
}

/** 明細レベルで紐づけを適用 */
function _applyOrderLinks_LineLevel(orderMgmtId, matches) {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var mgmtSheet   = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var orderSheet  = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  var mgmtData    = mgmtSheet.getDataRange().getValues();
  var orderLineData = orderSheet.getDataRange().getValues();
  var linkedQuoteIds = {};

  matches.forEach(function(m) {
    if (!m.quoteMgmtId || m.score < 50) return;
    linkedQuoteIds[m.quoteMgmtId] = true;
    for (var i = 1; i < orderLineData.length; i++) {
      if (orderLineData[i][0] === orderMgmtId &&
          (orderLineData[i][7] === m.orderLineNo || i === m.orderLineNo)) {
        orderSheet.getRange(i + 1, 3).setValue(m.quoteNo);
      }
    }
    _updateQuoteSideLink(m.quoteMgmtId, orderMgmtId);
  });

  var qIds   = Object.keys(linkedQuoteIds);
  var rowIdx = mgmtData.findIndex(function(r) { return r[0] === orderMgmtId; });
  if (rowIdx !== -1) {
    if (qIds.length > 1) {
      mgmtSheet.getRange(rowIdx + 1, MGMT_COLS.QUOTE_NO).setValue('(複数)');
      mgmtSheet.getRange(rowIdx + 1, MGMT_COLS.LINKED).setValue('TRUE');
    } else if (qIds.length === 1) {
      var qRow = mgmtData.find(function(r) { return r[0] === qIds[0]; });
      if (qRow) {
        mgmtSheet.getRange(rowIdx + 1, MGMT_COLS.QUOTE_NO).setValue(qRow[MGMT_COLS.QUOTE_NO - 1]);
        mgmtSheet.getRange(rowIdx + 1, MGMT_COLS.QUOTE_PDF_URL).setValue(qRow[MGMT_COLS.QUOTE_PDF_URL - 1]);
        mgmtSheet.getRange(rowIdx + 1, MGMT_COLS.LINKED).setValue('TRUE');
      }
    }
    mgmtSheet.getRange(rowIdx + 1, MGMT_COLS.STATUS).setValue(CONFIG.STATUS.RECEIVED);
    mgmtSheet.getRange(rowIdx + 1, MGMT_COLS.UPDATED_AT).setValue(nowJST());
  }
}

function _applyOrderLink_DocLevel(orderMgmtId, quoteMgmtId) {
  _applyOrderLinks_LineLevel(orderMgmtId, [{ orderLineNo: null, quoteMgmtId: quoteMgmtId, score: 100 }]);
}

function _updateQuoteSideLink(quoteMgmtId, orderMgmtId) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var data      = mgmtSheet.getDataRange().getValues();
  var qIdx      = data.findIndex(function(r) { return r[0] === quoteMgmtId; });
  var oIdx      = data.findIndex(function(r) { return r[0] === orderMgmtId; });
  if (qIdx !== -1 && oIdx !== -1) {
    mgmtSheet.getRange(qIdx + 1, MGMT_COLS.ORDER_NO).setValue(data[oIdx][MGMT_COLS.ORDER_NO - 1]);
    mgmtSheet.getRange(qIdx + 1, MGMT_COLS.ORDER_PDF_URL).setValue(data[oIdx][MGMT_COLS.ORDER_PDF_URL - 1]);
    mgmtSheet.getRange(qIdx + 1, MGMT_COLS.LINKED).setValue('TRUE');
    mgmtSheet.getRange(qIdx + 1, MGMT_COLS.STATUS).setValue(CONFIG.STATUS.RECEIVED);
  }
}

// ============================================================
// データ取得ユーティリティ
// ============================================================

function _getOrderInfo(ss, mgmtId) {
  var orderSheet = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  var mgmtData   = getAllMgmtData();
  var mgmtRow    = mgmtData.find(function(r) {
    return String(r[MGMT_COLS.ID - 1]).trim() === mgmtId;
  });
  if (!mgmtRow) return null;

  var orderItems  = [];
  var orderAmount = 0;
  if (orderSheet && orderSheet.getLastRow() > 1) {
    var od = orderSheet.getRange(2, 1, orderSheet.getLastRow() - 1, 20).getValues();
    od.forEach(function(r) {
      if (String(r[0]).trim() !== mgmtId) return;
      var itemName  = String(r[8]  || '').trim();
      var spec      = String(r[9]  || '').trim();
      var qty       = Number(r[12]) || 0;
      var unitPrice = Number(r[14]) || 0;
      var amount    = Number(r[15]) || 0;
      if (itemName) {
        orderItems.push({ itemName: itemName, spec: spec, qty: qty, unitPrice: unitPrice, amount: amount });
        orderAmount += amount;
      }
    });
  }

  return {
    mgmtId:      mgmtId,
    orderNo:     String(mgmtRow[MGMT_COLS.ORDER_NO - 1]      || '').trim(),
    quoteNo:     String(mgmtRow[MGMT_COLS.QUOTE_NO - 1]      || '').trim(),
    client:      String(mgmtRow[MGMT_COLS.CLIENT - 1]        || '').trim(),
    subject:     String(mgmtRow[MGMT_COLS.SUBJECT - 1]       || '').trim(),
    modelCode:   String(mgmtRow[MGMT_COLS.MODEL_CODE - 1]    || '').trim(),
    orderDate:   String(mgmtRow[MGMT_COLS.ORDER_DATE - 1]    || '').trim(),
    orderAmount: orderAmount || (Number(mgmtRow[MGMT_COLS.ORDER_AMOUNT - 1]) || 0),
    orderType:   String(mgmtRow[MGMT_COLS.ORDER_TYPE - 1]    || '').trim(),
    pdfUrl:      String(mgmtRow[MGMT_COLS.ORDER_PDF_URL - 1] || '').trim(),
    items:       orderItems
  };
}

function _getLinkedQuoteInfo(ss, quoteNo, orderMgmtId) {
  if (!quoteNo) return null;
  var mgmtData = getAllMgmtData();
  var qRow = mgmtData.find(function(r) {
    return String(r[MGMT_COLS.QUOTE_NO - 1]).trim() === quoteNo &&
           String(r[MGMT_COLS.ORDER_NO  - 1]).trim() !== '';
  }) || mgmtData.find(function(r) {
    return String(r[MGMT_COLS.QUOTE_NO - 1]).trim() === quoteNo;
  });
  if (!qRow) return null;
  return {
    mgmtId:      String(qRow[MGMT_COLS.ID - 1]            || '').trim(),
    quoteNo:     quoteNo,
    quoteAmount: Number(qRow[MGMT_COLS.QUOTE_AMOUNT - 1]) || 0,
    quotePdfUrl: String(qRow[MGMT_COLS.QUOTE_PDF_URL - 1] || '').trim()
  };
}

function _getAllQuotesForMatching(ss) {
  var quoteSheet    = ss.getSheetByName(CONFIG.SHEET_QUOTES);
  var mgmtData      = getAllMgmtData();
  var quoteItemsMap = {};

  if (quoteSheet && quoteSheet.getLastRow() > 1) {
    var qd = quoteSheet.getRange(2, 1, quoteSheet.getLastRow() - 1, 15).getValues();
    qd.forEach(function(r) {
      var mid = String(r[QUOTE_COLS.MGMT_ID - 1] || '').trim();
      if (!mid) return;
      if (!quoteItemsMap[mid]) quoteItemsMap[mid] = [];
      quoteItemsMap[mid].push({
        itemName:  String(r[QUOTE_COLS.ITEM_NAME  - 1] || '').trim(),
        spec:      String(r[QUOTE_COLS.SPEC       - 1] || '').trim(),
        qty:       Number(r[QUOTE_COLS.QTY        - 1]) || 0,
        unitPrice: Number(r[QUOTE_COLS.UNIT_PRICE - 1]) || 0,
        amount:    Number(r[QUOTE_COLS.AMOUNT     - 1]) || 0
      });
    });
  }

  var seen   = {};
  var quotes = [];
  mgmtData.forEach(function(r) {
    var qNo = String(r[MGMT_COLS.QUOTE_NO - 1] || '').trim();
    if (!qNo || seen[qNo]) return;
    seen[qNo]   = true;
    var mid     = String(r[MGMT_COLS.ID - 1] || '').trim();
    quotes.push({
      quoteMgmtId: mid,
      quoteNo:     qNo,
      client:      String(r[MGMT_COLS.CLIENT      - 1] || '').trim(),
      subject:     String(r[MGMT_COLS.SUBJECT     - 1] || '').trim(),
      modelCode:   String(r[MGMT_COLS.MODEL_CODE  - 1] || '').trim(),
      issueDate:   _toDateStr(r[MGMT_COLS.QUOTE_DATE - 1]),
      quoteAmount: Number(r[MGMT_COLS.QUOTE_AMOUNT - 1]) || 0,
      quotePdfUrl: String(r[MGMT_COLS.QUOTE_PDF_URL - 1] || '').trim(),
      destCompany: String(r[MGMT_COLS.CLIENT      - 1] || '').trim(),
      items:       quoteItemsMap[mid] || []
    });
  });
  return quotes;
}

function _getOrderLines(mgmtId) {
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ORDERS).getDataRange().getValues();
  var head = data[0];
  return data.slice(1).filter(function(r) { return r[0] === mgmtId; }).map(function(r, i) {
    var o = { lineNo: i + 1 };
    head.forEach(function(h, j) { o[h] = r[j]; });
    return o;
  });
}

function _getAllQuoteLines() {
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_QUOTES).getDataRange().getValues();
  var head = data[0];
  return data.slice(1).map(function(r) {
    var o = {};
    head.forEach(function(h, j) { o[h] = r[j]; });
    return o;
  });
}

function _buildQuoteGroups(mgmtData, headers) {
  var qNoIdx     = headers.indexOf('見積番号');
  if (qNoIdx < 0) qNoIdx = MGMT_COLS.QUOTE_NO - 1;
  var clientIdx  = headers.indexOf('顧客名');
  var subjectIdx = headers.indexOf('件名');
  var modelIdx   = headers.indexOf('機種コード');

  return mgmtData.slice(1)
    .filter(function(r) { return r[0] && String(r[qNoIdx] || '').trim() !== ''; })
    .map(function(r) {
      return {
        mgmtId:      r[0],
        quoteNo:     r[qNoIdx],
        modelCode:   modelIdx   >= 0 ? r[modelIdx]   : '',
        client:      clientIdx  >= 0 ? r[clientIdx]  : '',
        subject:     subjectIdx >= 0 ? r[subjectIdx] : '',
        totalAmount: r[MGMT_COLS.TOTAL - 1] || r[MGMT_COLS.QUOTE_AMOUNT - 1] || '',
        issueDate:   r[MGMT_COLS.QUOTE_DATE - 1] || ''
      };
    });
}

function _buildOrderGroups(mgmtData, headers) {
  var orderNoIdx = headers.indexOf('注文番号');
  if (orderNoIdx < 0) orderNoIdx = MGMT_COLS.ORDER_NO - 1;
  return mgmtData.slice(1)
    .filter(function(r) { return r[0] && String(r[0]).trim() !== '' && String(r[orderNoIdx] || '').trim() !== ''; })
    .map(function(r) { return { mgmtId: r[0], orderNo: r[orderNoIdx] }; });
}

// ============================================================
// 通知（統合: Telegram + Google Chat）
// ============================================================

/** 注文書到着時に見積書候補を含めて通知する */
function notifyOrderArrivalWithCandidates(mgmtId, orderNo, client, orderAmount, orderPdfUrl) {
  try {
    var ss        = getSpreadsheet();
    var orderInfo = _getOrderInfo(ss, mgmtId);
    if (!orderInfo) { Logger.log('[NOTIFY] orderInfo not found for ' + mgmtId); return; }

    var quotes = _getAllQuotesForMatching(ss);
    var top3   = _scoreQuotesForOrder(orderInfo, quotes)
      .filter(function(c) { return c.score >= MATCH_CONFIG.CANDIDATE_THRESHOLD; })
      .slice(0, 3);

    var msg = '📦 *新規注文書を受信しました*\n\n';
    msg += '📋 発注番号: `' + orderNo + '`\n';
    msg += '🏢 発注元: ' + client + '\n';
    if (orderAmount) msg += '💴 金額: ¥' + Number(orderAmount).toLocaleString() + '\n';
    if (orderPdfUrl) msg += '📄 発注書PDF: ' + orderPdfUrl + '\n';
    msg += '\n';

    if (top3.length > 0) {
      msg += '🔍 *対応する見積書候補（上位' + top3.length + '件）:*\n';
      top3.forEach(function(c, i) {
        var emoji = c.score >= 80 ? '🟢' : (c.score >= 60 ? '🟡' : '🔴');
        msg += '\n' + (i + 1) + '. ' + emoji + ' *見積No: ' + (c.quoteNo || '—') + '*\n';
        msg += '   件名: ' + (c.subject || c.destCompany || '—') + '\n';
        msg += '   発行日: ' + (c.issueDate || '—') + ' ／ 金額: ¥' + (c.quoteAmount ? Number(c.quoteAmount).toLocaleString() : '—') + '\n';
        msg += '   一致度: ' + c.score + '点 (' + (c.reason || '') + ')\n';
        if (c.quotePdfUrl) msg += '   📄 見積書PDF: ' + c.quotePdfUrl + '\n';
      });
    } else {
      msg += '⚠️ 対応する見積書候補が見つかりませんでした。\n手動で紐づけ画面から確認してください。';
    }
    msg += '\n\n🔗 管理システムで確認・紐づけ:\n' + _getWebAppUrl();

    _sendTelegramIfConfigured(msg);
    _sendGoogleChatNotification(msg, null, null, null);
    _saveMatchingCandidatesToSheet(ss, mgmtId, orderNo, client, orderAmount, top3);

  } catch (e) {
    Logger.log('[NOTIFY ERROR] ' + e.message + '\n' + e.stack);
  }
}

/** 注文書アップロード完了後のフック */
function onOrderUploaded(mgmtId, orderNo, client, orderAmount, pdfUrl) {
  try {
    Utilities.sleep(500);
    notifyOrderArrivalWithCandidates(mgmtId, orderNo, client, orderAmount, pdfUrl);
  } catch (e) {
    Logger.log('[ON ORDER UPLOADED ERROR] ' + e.message);
  }
}

/**
 * Google Chat 通知（詳細版）
 * docType: 'quote' | 'order'
 * actionType: 'new' | 'revision' | 'cancellation'
 */
function _sendGoogleChatNotification(message, mgmtId, docType, actionType) {
  try {
    var webhookUrl = PropertiesService.getScriptProperties().getProperty('GOOGLE_CHAT_WEBHOOK_URL') ||
                    PropertiesService.getScriptProperties().getProperty('GCHAT_WEBHOOK_URL');
    if (!webhookUrl) { Logger.log('[GCHAT] WebhookURL未設定のため通知スキップ'); return; }

    // mgmtId 指定がある場合は管理シートから詳細テキストを構築
    var text = message;
    if (mgmtId && docType) {
      var ss    = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
      var data  = sheet.getDataRange().getValues();
      var row   = data.find(function(r) { return String(r[0]) === String(mgmtId); });
      if (row) {
        var getCol = function(colNum) { return String(row[colNum - 1] || 'なし'); };
        var action = actionType || 'new';
        var title  = docType === 'quote' ? '📄 見積書を登録' : '📦 注文書を受領';
        var icon   = '🔹';
        if (action === 'revision')     { title = '🔄 注文書の差し替え受領'; icon = '⚠️'; }
        else if (action === 'cancellation') { title = '❌ 注文書のキャンセル'; icon = '🚫'; }

        text  = '【' + title + '】\n━━━━━━━━━━━━━━\n';
        text += icon + ' 案件: ' + getCol(MGMT_COLS.SUBJECT) + '\n';
        text += icon + ' 顧客: ' + getCol(MGMT_COLS.CLIENT)  + '\n';
        if (action !== 'cancellation') {
          var amount    = docType === 'quote' ? getCol(MGMT_COLS.QUOTE_AMOUNT) : getCol(MGMT_COLS.ORDER_AMOUNT);
          var amountNum = parseFloat(String(amount).replace(/[^\d.]/g, ''));
          text += icon + ' 金額: ¥' + (!isNaN(amountNum) && amountNum > 0 ? amountNum.toLocaleString() : '—') + '\n';
          var pdfUrl = docType === 'quote' ? getCol(MGMT_COLS.QUOTE_PDF_URL) : getCol(MGMT_COLS.ORDER_PDF_URL);
          if (pdfUrl && pdfUrl !== 'なし') text += icon + ' PDF: ' + pdfUrl + '\n';
        } else {
          text += '🚨 この注文はキャンセルとして処理されました。\n';
          var memo = getCol(MGMT_COLS.MEMO);
          if (memo && memo !== 'なし') text += '理由: ' + memo + '\n';
        }
        try { text += '\n🌐 システムで確認: ' + ScriptApp.getService().getUrl(); } catch (e2) {}
      }
    }

    UrlFetchApp.fetch(webhookUrl, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ text: text }),
      muteHttpExceptions: true
    });
    Logger.log('[GCHAT] 送信完了');
  } catch (e) {
    Logger.log('[GCHAT ERROR] ' + e.message);
  }
}

function _sendTelegramIfConfigured(message) {
  try {
    var token  = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
    var chatId = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID');
    if (!token || !chatId) return;
    UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'POST', contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, text: message, parse_mode: 'Markdown', disable_web_page_preview: true }),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('[TELEGRAM ERROR] ' + e.message);
  }
}

function _saveMatchingCandidatesToSheet(ss, orderMgmtId, orderNo, client, orderAmount, candidates) {
  try {
    var sheet = ss.getSheetByName(MATCH_CONFIG.SHEET_CANDIDATES);
    if (!sheet) {
      sheet = ss.insertSheet(MATCH_CONFIG.SHEET_CANDIDATES);
      sheet.appendRow([
        '注文管理ID','注文番号','顧客名','発注日','注文金額',
        '候補1_管理ID','候補1_見積No','候補1_顧客','候補1_スコア','候補1_理由',
        '候補2_管理ID','候補2_スコア',
        '候補3_管理ID','候補3_スコア',
        '備考','ステータス','更新日時'
      ]);
    }
    var lastRow        = sheet.getLastRow();
    var existingRowNum = -1;
    if (lastRow > 1) {
      var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(function(r) { return String(r[0]); });
      var idx = ids.indexOf(orderMgmtId);
      if (idx >= 0) existingRowNum = idx + 2;
    }
    var c1  = candidates[0] || {};
    var c2  = candidates[1] || {};
    var c3  = candidates[2] || {};
    var row = [
      orderMgmtId, orderNo, client, '', orderAmount || '',
      c1.quoteMgmtId||'', c1.quoteNo||'', c1.destCompany||'', c1.score||'', c1.reason||'',
      c2.quoteMgmtId||'', c2.score||'',
      c3.quoteMgmtId||'', c3.score||'',
      '', '未対応', nowJST()
    ];
    if (existingRowNum > 0) {
      sheet.getRange(existingRowNum, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
  } catch (e) {
    Logger.log('[SAVE CANDIDATES ERROR] ' + e.message);
  }
}

// ============================================================
// テキスト正規化・類似度計算
// ============================================================

function _normalizeCompanyName(s) {
  return String(s)
    .replace(/株式会社|有限会社|合同会社|合資会社/g, '')
    .replace(/（株）|\(株\)|（有）|\(有\)/g, '')
    .replace(/[\s　]+/g, '')
    .toLowerCase();
}

function _normItemName(s) {
  return String(s)
    .replace(/[ａ-ｚＡ-Ｚ０-９]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); })
    .replace(/[ァ-ン]/g, function(c)            { return String.fromCharCode(c.charCodeAt(0) - 0x60); })
    .replace(/[\s　]+/g, '')
    .toLowerCase();
}

function _textSimilarity(a, b) {
  if (!a || !b) return 0;
  if (a === b)  return 1;
  if (a.indexOf(b) >= 0 || b.indexOf(a) >= 0) return 0.9;
  var bigramsA = _bigrams(a);
  var bigramsB = _bigrams(b);
  if (!bigramsA.length || !bigramsB.length) return 0;
  var setA = {};
  bigramsA.forEach(function(bg) { setA[bg] = (setA[bg] || 0) + 1; });
  var intersect = 0;
  bigramsB.forEach(function(bg) { if (setA[bg] && setA[bg] > 0) { intersect++; setA[bg]--; } });
  return (2 * intersect) / (bigramsA.length + bigramsB.length);
}

function _bigrams(s) {
  var result = [];
  for (var i = 0; i < s.length - 1; i++) result.push(s.slice(i, i + 2));
  return result;
}

// ============================================================
// その他ユーティリティ
// ============================================================

function _getWebAppUrl() {
  try { return ScriptApp.getService().getUrl(); } catch (e) { return '（URLを取得できませんでした）'; }
}

function _toDateStr(v) {
  if (!v) return '';
  try {
    var d = (v instanceof Date) ? v : new Date(v);
    if (isNaN(d.getTime())) return String(v);
    return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
  } catch (e) { return String(v); }
}

/** 注文書一覧の行に「紐づけ」ボタンHTMLを生成 */
function generateLinkButtonHtml(item) {
  var isLinked = item.linked || item.quoteNo;
  if (isLinked) {
    return '<button class="btn btn-sm" style="background:#e6f4ea;color:#137333;border:1px solid #a5d6a7;border-radius:6px;padding:3px 10px;font-size:11px;cursor:pointer;" ' +
           'onclick="event.stopPropagation();openOrderLinkModal(\'' + item.id + '\', null)">' +
           '✅ ' + (item.quoteNo || '紐づけ済') + '</button>';
  }
  return '<button class="btn btn-sm" style="background:#fff3e0;color:#e65100;border:1px solid #ffe0b2;border-radius:6px;padding:3px 10px;font-size:11px;cursor:pointer;" ' +
         'onclick="event.stopPropagation();openOrderLinkModal(\'' + item.id + '\', null)">' +
         '🔗 見積書を紐づける</button>';
}
