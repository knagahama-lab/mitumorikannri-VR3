// ============================================================
// 見積書・注文書管理システム
// ファイル 10: AI紐づけエンジン（Gemini 明細レベル対応）★修正版
// ============================================================

var MATCHING_THRESHOLD_AUTO      = 80;
var MATCHING_THRESHOLD_CANDIDATE = 50;

// ============================================================
// 注文書受領時：見積書を明細単位でAI紐付け
// ============================================================

function aiLinkOrderToQuote(orderMgmtId) {
  try {
    var ss        = getSpreadsheet();                              // ★ CONFIG.SPREADSHEET_ID を使う正しい関数
    var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);   // ★ '管理シート' (CONFIG参照)
    if (!mgmtSheet) return { success: false, error: '管理シートが見つかりません: ' + CONFIG.SHEET_MANAGEMENT };

    var mgmtData = mgmtSheet.getDataRange().getValues();
    if (mgmtData.length <= 1) return { success: false, error: 'データなし' };
    var headers = mgmtData[0];

    var orderIdx = mgmtData.findIndex(function(r) { return String(r[0]) === String(orderMgmtId); });
    if (orderIdx === -1) return { success: false, error: '注文データが見つかりません: ' + orderMgmtId };

    var orderRow  = mgmtData[orderIdx];
    var orderData = {};
    headers.forEach(function(h, i) { orderData[h] = orderRow[i]; });

    // 明細データの取得
    var orderLines = _getOrderLines(orderMgmtId);
    if (orderLines.length === 0) {
      // 明細なしでも候補検索は続ける（ドキュメントレベルで紐づけ試行）
      Logger.log('[AI LINK] 明細なし。ドキュメントレベルで試みます: ' + orderMgmtId);
    }

    // マッチング候補の見積書を抽出（未紐づけ見積書のみ）
    var quoteGroups = _buildQuoteGroups_Safe(ss, mgmtData, headers);
    var quoteLines  = _getAllQuoteLines_Safe(ss);

    if (quoteGroups.length === 0) {
      Logger.log('[AI LINK] 未紐づけ見積書なし');
      return { success: true, status: 'no_quotes', candidates: [] };
    }

    // Gemini API による明細レベルのマッチング推論
    var aiResult = _matchWithGeminiSafe(orderData, orderLines, quoteGroups, quoteLines);

    if (!aiResult || !aiResult.matches || aiResult.matches.length === 0) {
      // AI失敗時：スコアベースのシンプルマッチング
      var simpleResult = _simpleScoreMatch(orderData, quoteGroups);
      return _applyMatchResult(orderMgmtId, simpleResult, mgmtSheet, mgmtData);
    }

    return _applyMatchResult(orderMgmtId, aiResult.matches, mgmtSheet, mgmtData);

  } catch(e) {
    Logger.log('[aiLinkOrderToQuote ERROR] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

// ============================================================
// 見積書アップロード時：注文書をドキュメント単位で探索
// ============================================================

function aiLinkQuoteToOrder(quoteMgmtId) {
  try {
    var ss        = getSpreadsheet();
    var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    if (!mgmtSheet) return { success: false, error: '管理シートが見つかりません' };

    var mgmtData = mgmtSheet.getDataRange().getValues();
    if (mgmtData.length <= 1) return { success: false, error: 'データなし' };
    var headers = mgmtData[0];

    var quoteIdx = mgmtData.findIndex(function(r) { return String(r[0]) === String(quoteMgmtId); });
    if (quoteIdx === -1) return { success: false, error: '見積データが見つかりません: ' + quoteMgmtId };

    var quoteRow  = mgmtData[quoteIdx];
    var quoteData = {};
    headers.forEach(function(h, i) { quoteData[h] = quoteRow[i]; });

    // 未紐づけ注文書を抽出
    var orderGroups = _buildOrderGroups_Safe(mgmtData, headers);
    if (orderGroups.length === 0) {
      return { success: true, status: 'no_orders', candidates: [] };
    }

    // Gemini API によるマッチング推論
    var aiResult = _matchQuoteToOrderWithGemini(quoteData, orderGroups);

    if (!aiResult || !aiResult.matches || aiResult.matches.length === 0) {
      return { success: true, status: 'no_match', candidates: [] };
    }

    var matches    = aiResult.matches.sort(function(a, b) { return b.score - a.score; });
    var bestMatch  = matches[0];

    if (bestMatch && bestMatch.score >= MATCHING_THRESHOLD_AUTO) {
      _applyOrderLink_DocLevel(bestMatch.orderMgmtId, quoteMgmtId);
      return { success: true, status: 'auto_linked', bestMatch: bestMatch };
    }

    if (bestMatch && bestMatch.score >= MATCHING_THRESHOLD_CANDIDATE) {
      return { success: true, status: 'candidates_found', candidates: matches };
    }

    return { success: true, status: 'no_match', candidates: matches };

  } catch(e) {
    Logger.log('[aiLinkQuoteToOrder ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// Gemini API 推論ロジック（エラー耐性強化版）
// ============================================================

function _matchWithGeminiSafe(orderData, orderLines, quoteGroups, quoteLines) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) { Logger.log('[GEMINI] APIキー未設定'); return null; }

    var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;

    var quoteDetailSummary = quoteGroups.slice(0, 30).map(function(q) {
      var details = quoteLines.filter(function(ql) { return String(ql.mgmtId) === String(q.mgmtId); });
      return {
        mgmtId:      q.mgmtId,
        quoteNo:     q.quoteNo,
        issueDate:   q.issueDate || '',
        client:      q.client   || '',
        totalAmount: q.amount   || '',
        items:       details.slice(0, 5).map(function(d) {
          return (d.itemName || '') + ' / ' + (d.spec || '') + ' / 単価:' + (d.unitPrice || '');
        }).join('; ') || q.subject || '',
      };
    });

    var prompt = 'あなたは見積・注文管理システムの紐づけエンジンです。注文書に最適な見積書を特定してください。\n\n' +
      '【注文書】番号:' + (orderData['注文番号'] || orderData[headers_safe('ORDER_NO')] || '') +
      ' | 顧客:' + (orderData['顧客名'] || '') +
      ' | 発注日:' + (orderData['発注日'] || '') + '\n' +
      '【注文明細】\n' +
      (orderLines.length > 0
        ? orderLines.slice(0,10).map(function(l) {
            return '行:' + l.lineNo + ' | 品名:' + (l.itemName || '') + ' | 仕様:' + (l.spec || '') + ' | 単価:' + (l.unitPrice || '');
          }).join('\n')
        : '（明細なし）') + '\n\n' +
      '【見積書候補（最大30件）】\n' +
      quoteDetailSummary.map(function(q) {
        return 'ID:' + q.mgmtId + ' | 見積番号:' + q.quoteNo + ' | 顧客:' + q.client + ' | 発行日:' + q.issueDate + ' | 合計:' + q.totalAmount + ' | 明細:[' + q.items + ']';
      }).join('\n') + '\n\n' +
      '【返却形式】JSONのみ。スコア0-100。80以上で自動紐づけ対象。\n' +
      '{"isMixed":false,"matches":[{"orderLineNo":null,"quoteMgmtId":"ID","quoteNo":"見積番号","score":0-100,"reason":"理由"}]}';

    var response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { responseMimeType: 'application/json', temperature: 0.1 },
      }),
      muteHttpExceptions: true,
    });

    if (response.getResponseCode() !== 200) {
      Logger.log('[GEMINI MATCH] HTTP ' + response.getResponseCode());
      return null;
    }

    var body = JSON.parse(response.getContentText());
    var text = body.candidates && body.candidates[0] && body.candidates[0].content &&
               body.candidates[0].content.parts ? body.candidates[0].content.parts[0].text : '';
    text = text.replace(/```json|```/g, '').trim();
    return JSON.parse(text);

  } catch(e) {
    Logger.log('[GEMINI MATCH ERROR] ' + e.message);
    return null;
  }
}

// ダミーヘッダー参照ヘルパー（内部で安全に列名を取得）
function headers_safe(colKey) {
  var map = {
    'ORDER_NO': '注文番号',
    'QUOTE_NO': '見積番号',
  };
  return map[colKey] || '';
}

function _matchQuoteToOrderWithGemini(quoteData, orderGroups) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return null;
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;
    var prompt = '見積書に最適な注文書を選んでください（JSONのみ）\n見積:' +
      (quoteData['見積番号'] || '') + '|' + (quoteData['顧客名'] || '') + '\n注文候補:\n' +
      orderGroups.map(function(o) { return 'ID:' + o.mgmtId + '|番号:' + o.orderNo + '|顧客:' + (o.client||''); }).join('\n') +
      '\n{"matches":[{"orderMgmtId":"ID","score":0-100,"reason":"理由"}]}';
    var response = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ contents:[{ parts:[{ text:prompt }] }], generationConfig:{ responseMimeType:'application/json' } }),
      muteHttpExceptions: true,
    });
    if (response.getResponseCode() !== 200) return null;
    var body = JSON.parse(response.getContentText());
    var text = body.candidates[0].content.parts[0].text.replace(/```json|```/g,'').trim();
    return JSON.parse(text);
  } catch(e) {
    Logger.log('[matchQuoteToOrder ERROR] ' + e.message);
    return null;
  }
}

// ============================================================
// スコアベースのシンプルマッチング（AIフォールバック）
// ============================================================

function _simpleScoreMatch(orderData, quoteGroups) {
  var orderClient  = normalizeText(String(orderData['顧客名'] || orderData[Object.keys(orderData).find(function(k){return k.indexOf('顧客')>=0;})|| ''] || ''));
  var orderNo      = normalizeText(String(orderData['注文番号'] || ''));
  var orderDate    = String(orderData['発注日'] || '');

  var scored = quoteGroups.map(function(q) {
    var score  = 0;
    var qClient = normalizeText(q.client || '');
    var qNo     = normalizeText(q.quoteNo || '');

    // 顧客名一致
    if (orderClient && qClient && (orderClient.indexOf(qClient) >= 0 || qClient.indexOf(orderClient) >= 0)) score += 40;
    // 発行日が発注日より前
    if (q.issueDate && orderDate && q.issueDate < orderDate) score += 10;
    // 発行日が近い（3ヶ月以内）
    if (q.issueDate && orderDate) {
      var diff = Math.abs(new Date(orderDate.replace(/\//g,'-')).getTime() - new Date(q.issueDate.replace(/\//g,'-')).getTime());
      if (diff < 90 * 86400 * 1000) score += 15;
    }

    return { orderLineNo: null, quoteMgmtId: q.mgmtId, quoteNo: q.quoteNo, score: score, reason: 'スコアベース（顧客名・発行日）' };
  });

  scored.sort(function(a, b) { return b.score - a.score; });
  return scored.slice(0, 5);
}

// ============================================================
// マッチング結果の適用
// ============================================================

function _applyMatchResult(orderMgmtId, matches, mgmtSheet, mgmtData) {
  if (!matches || matches.length === 0) {
    return { success: true, status: 'no_match', candidates: [] };
  }

  var best = matches[0];

  // ★ 80点以上は自動紐づけ
  if (best.score >= MATCHING_THRESHOLD_AUTO) {
    try {
      _applyOrderLinks_LineLevel(orderMgmtId, [best]);
      return {
        success: true,
        status: 'auto_linked',
        bestMatch: best,
        score: best.score,
        quoteNo: best.quoteNo,
      };
    } catch(e) {
      Logger.log('[APPLY MATCH ERROR] ' + e.message);
    }
  }

  // 50点以上は候補として保存
  var candidates = matches.filter(function(m) { return m.score >= MATCHING_THRESHOLD_CANDIDATE; });

  if (candidates.length > 0) {
    // スクリプトプロパティに候補を保存
    try {
      var raw   = PropertiesService.getScriptProperties().getProperty('AI_MATCHING_CANDIDATES') || '[]';
      var stored = JSON.parse(raw);
      // 既存エントリを置き換え
      stored = stored.filter(function(s) { return String(s.orderMgmtId) !== String(orderMgmtId); });
      var oRow = mgmtData.find(function(r) { return String(r[0]) === String(orderMgmtId); });
      stored.push({
        orderMgmtId: orderMgmtId,
        orderNo:     oRow ? String(oRow[MGMT_COLS.ORDER_NO - 1] || '') : '',
        orderClient: oRow ? String(oRow[MGMT_COLS.CLIENT - 1]   || '') : '',
        orderDate:   oRow ? _toDateStr(oRow[MGMT_COLS.ORDER_DATE - 1])  : '',
        orderAmount: oRow ? _toNum(oRow[MGMT_COLS.ORDER_AMOUNT - 1])    : '',
        orderPdfUrl: oRow ? String(oRow[MGMT_COLS.ORDER_PDF_URL - 1] || '') : '',
        candidates:  candidates,
      });
      PropertiesService.getScriptProperties().setProperty('AI_MATCHING_CANDIDATES', JSON.stringify(stored));
    } catch(propErr) {
      Logger.log('[PROP STORE ERROR] ' + propErr.message);
    }

    return { success: true, status: 'candidates_found', candidates: candidates };
  }

  // 候補すら0件の場合もプロパティに記録（フォールバック表示のため）
  try {
    var raw2   = PropertiesService.getScriptProperties().getProperty('AI_MATCHING_CANDIDATES') || '[]';
    var stored2 = JSON.parse(raw2);
    stored2 = stored2.filter(function(s) { return String(s.orderMgmtId) !== String(orderMgmtId); });
    var oRow2 = mgmtData.find(function(r) { return String(r[0]) === String(orderMgmtId); });
    stored2.push({
      orderMgmtId: orderMgmtId,
      orderNo:     oRow2 ? String(oRow2[MGMT_COLS.ORDER_NO - 1] || '') : '',
      orderClient: oRow2 ? String(oRow2[MGMT_COLS.CLIENT - 1]   || '') : '',
      orderDate:   oRow2 ? _toDateStr(oRow2[MGMT_COLS.ORDER_DATE - 1])  : '',
      orderAmount: oRow2 ? _toNum(oRow2[MGMT_COLS.ORDER_AMOUNT - 1])    : '',
      orderPdfUrl: oRow2 ? String(oRow2[MGMT_COLS.ORDER_PDF_URL - 1] || '') : '',
      candidates:  [], // ← getMatchingCandidates() が補填する
    });
    PropertiesService.getScriptProperties().setProperty('AI_MATCHING_CANDIDATES', JSON.stringify(stored2));
  } catch(e2) {}

  return { success: true, status: 'no_match', candidates: [] };
}

// ============================================================
// データ反映ロジック
// ============================================================

function _applyOrderLinks_LineLevel(orderMgmtId, matches) {
  try {
    var ss        = getSpreadsheet();
    var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var orderSheet = ss.getSheetByName(CONFIG.SHEET_ORDERS);
    var mgmtData  = mgmtSheet.getDataRange().getValues();
    var orderLineData = (orderSheet && orderSheet.getLastRow() > 1)
      ? orderSheet.getRange(2, 1, orderSheet.getLastRow()-1, 19).getValues()
      : [];

    var linkedQuoteIds = {};

    matches.forEach(function(m) {
      if (!m.quoteMgmtId || m.score < MATCHING_THRESHOLD_CANDIDATE) return;
      linkedQuoteIds[m.quoteMgmtId] = true;

      // 注文書明細の紐づけ列（col 3 = LINKED_QUOTE）を更新
      if (orderLineData.length > 0 && m.orderLineNo != null) {
        for (var i = 0; i < orderLineData.length; i++) {
          if (String(orderLineData[i][0]) === String(orderMgmtId) &&
              (orderLineData[i][7] === m.orderLineNo || (i + 1) === m.orderLineNo)) {
            orderSheet.getRange(i + 2, 3).setValue(m.quoteNo || '');
          }
        }
      }
      _updateQuoteSideLink(m.quoteMgmtId, orderMgmtId);
    });

    var qIds = Object.keys(linkedQuoteIds);
    var rowIdx = mgmtData.findIndex(function(r) { return String(r[0]) === String(orderMgmtId); });

    if (rowIdx !== -1) {
      var sheetRow = rowIdx + 1;
      if (qIds.length > 1) {
        mgmtSheet.getRange(sheetRow, MGMT_COLS.QUOTE_NO).setValue('(複数)');
        mgmtSheet.getRange(sheetRow, MGMT_COLS.LINKED).setValue('TRUE');
      } else if (qIds.length === 1) {
        var qRow = mgmtData.find(function(r) { return String(r[0]) === qIds[0]; });
        if (qRow) {
          mgmtSheet.getRange(sheetRow, MGMT_COLS.QUOTE_NO).setValue(qRow[MGMT_COLS.QUOTE_NO - 1] || '');
          mgmtSheet.getRange(sheetRow, MGMT_COLS.QUOTE_PDF_URL).setValue(qRow[MGMT_COLS.QUOTE_PDF_URL - 1] || '');
          mgmtSheet.getRange(sheetRow, MGMT_COLS.LINKED).setValue('TRUE');
        }
      }
      if (qIds.length > 0) {
        mgmtSheet.getRange(sheetRow, MGMT_COLS.STATUS).setValue(CONFIG.STATUS.RECEIVED);
        mgmtSheet.getRange(sheetRow, MGMT_COLS.UPDATED_AT).setValue(nowJST());
      }
    }
  } catch(e) {
    Logger.log('[_applyOrderLinks_LineLevel ERROR] ' + e.message);
  }
}

function _applyOrderLink_DocLevel(orderMgmtId, quoteMgmtId) {
  var ss       = getSpreadsheet();
  var mgmtData = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT).getDataRange().getValues();
  var qRow     = mgmtData.find(function(r) { return String(r[0]) === String(quoteMgmtId); });
  var quoteNo  = qRow ? String(qRow[MGMT_COLS.QUOTE_NO - 1] || '') : '';
  _applyOrderLinks_LineLevel(orderMgmtId, [{
    orderLineNo: null, quoteMgmtId: quoteMgmtId, quoteNo: quoteNo, score: 100,
  }]);
}

function _updateQuoteSideLink(quoteMgmtId, orderMgmtId) {
  try {
    var ss        = getSpreadsheet();
    var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var data      = mgmtSheet.getDataRange().getValues();
    var qIdx      = data.findIndex(function(r) { return String(r[0]) === String(quoteMgmtId); });
    var oIdx      = data.findIndex(function(r) { return String(r[0]) === String(orderMgmtId); });
    if (qIdx === -1 || oIdx === -1) return;
    mgmtSheet.getRange(qIdx+1, MGMT_COLS.ORDER_NO).setValue(data[oIdx][MGMT_COLS.ORDER_NO-1] || '');
    mgmtSheet.getRange(qIdx+1, MGMT_COLS.ORDER_PDF_URL).setValue(data[oIdx][MGMT_COLS.ORDER_PDF_URL-1] || '');
    mgmtSheet.getRange(qIdx+1, MGMT_COLS.LINKED).setValue('TRUE');
    mgmtSheet.getRange(qIdx+1, MGMT_COLS.STATUS).setValue(CONFIG.STATUS.RECEIVED);
    mgmtSheet.getRange(qIdx+1, MGMT_COLS.UPDATED_AT).setValue(nowJST());
  } catch(e) {
    Logger.log('[_updateQuoteSideLink ERROR] ' + e.message);
  }
}

// ============================================================
// データ取得ヘルパー（安全版）
// ============================================================

function _getOrderLines(mgmtId) {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_ORDERS);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 19).getValues();
    return data
      .filter(function(r) { return String(r[0]) === String(mgmtId); })
      .map(function(r, i) {
        return {
          lineNo:    r[7] || (i + 1),
          itemName:  String(r[8]  || ''),
          spec:      String(r[9]  || ''),
          qty:       r[12],
          unit:      String(r[13] || ''),
          unitPrice: r[14],
          amount:    r[15],
          remarks:   String(r[16] || ''),
        };
      });
  } catch(e) {
    Logger.log('[_getOrderLines ERROR] ' + e.message);
    return [];
  }
}

function _getAllQuoteLines_Safe(ss) {
  try {
    var sheet = ss.getSheetByName(CONFIG.SHEET_QUOTES);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    return sheet.getRange(2, 1, sheet.getLastRow()-1, 15).getValues()
      .filter(function(r) { return r[0] && r[6]; })
      .map(function(r) {
        return {
          mgmtId:    String(r[0] || ''),
          quoteNo:   String(r[1] || ''),
          itemName:  String(r[6] || ''),
          spec:      String(r[7] || ''),
          qty:       r[8],
          unitPrice: r[10],
          amount:    r[11],
          pdfUrl:    String(r[13] || ''),
        };
      });
  } catch(e) {
    Logger.log('[_getAllQuoteLines_Safe ERROR] ' + e.message);
    return [];
  }
}

function _buildQuoteGroups_Safe(ss, mgmtData, headers) {
  try {
    // 未紐づけの見積書のみ
    return mgmtData.slice(1).filter(function(r) {
      return String(r[MGMT_COLS.QUOTE_NO - 1]).trim() !== '' &&
             !_isLinkedVal(r[MGMT_COLS.LINKED - 1]);
    }).map(function(r) {
      return {
        mgmtId:    String(r[MGMT_COLS.ID - 1]          || ''),
        quoteNo:   String(r[MGMT_COLS.QUOTE_NO - 1]     || ''),
        client:    String(r[MGMT_COLS.CLIENT - 1]        || ''),
        issueDate: _toDateStr(r[MGMT_COLS.QUOTE_DATE - 1]),
        amount:    _toNum(r[MGMT_COLS.QUOTE_AMOUNT - 1]),
        subject:   String(r[MGMT_COLS.SUBJECT - 1]       || ''),
        pdfUrl:    String(r[MGMT_COLS.QUOTE_PDF_URL - 1] || ''),
      };
    });
  } catch(e) {
    Logger.log('[_buildQuoteGroups_Safe ERROR] ' + e.message);
    return [];
  }
}

function _buildOrderGroups_Safe(mgmtData, headers) {
  try {
    return mgmtData.slice(1).filter(function(r) {
      return String(r[MGMT_COLS.ORDER_NO - 1]).trim() !== '' &&
             !_isLinkedVal(r[MGMT_COLS.LINKED - 1]);
    }).map(function(r) {
      return {
        mgmtId:  String(r[MGMT_COLS.ID - 1]       || ''),
        orderNo: String(r[MGMT_COLS.ORDER_NO - 1]  || ''),
        client:  String(r[MGMT_COLS.CLIENT - 1]    || ''),
      };
    });
  } catch(e) {
    Logger.log('[_buildOrderGroups_Safe ERROR] ' + e.message);
    return [];
  }
}

// ============================================================
// Chat通知
// ============================================================

function _sendChatNotification(mgmtId, docType) {
  try {
    var webhookUrl = _getChatWebhookUrl();
    if (!webhookUrl) return;
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    if (!sheet || sheet.getLastRow() <= 1) return;
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 27).getValues();
    var row  = data.find(function(r) { return String(r[0]) === String(mgmtId); });
    if (!row) return;
    var subject  = String(row[MGMT_COLS.SUBJECT - 1]    || 'なし');
    var client   = String(row[MGMT_COLS.CLIENT - 1]     || 'なし');
    var amountCol = docType === 'quote' ? MGMT_COLS.QUOTE_AMOUNT : MGMT_COLS.ORDER_AMOUNT;
    var pdfCol    = docType === 'quote' ? MGMT_COLS.QUOTE_PDF_URL: MGMT_COLS.ORDER_PDF_URL;
    var amount   = Number(row[amountCol - 1] || 0).toLocaleString();
    var pdfUrl   = String(row[pdfCol - 1]    || '');
    var linked   = _isLinkedVal(row[MGMT_COLS.LINKED - 1]);
    var title    = docType === 'quote' ? '📄 見積書を登録' : '📦 注文書を受領';
    var text = '【' + title + '】\n案件: ' + subject + '\n顧客: ' + client + '\n金額: ¥' + amount;
    if (pdfUrl) text += '\nPDF: ' + pdfUrl;
    if (linked) text += '\n✅ AI紐付け完了';
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ text: text }),
      muteHttpExceptions: true,
    });
  } catch(e) {
    Logger.log('[_sendChatNotification ERROR] ' + e.message);
  }
}

function _sendOrderRegistrationToChat(mgmtId, uploadInfo, linkResult) {
  try {
    var webhookUrl = _getChatWebhookUrl();
    if (!webhookUrl) return;
    var lr    = linkResult || {};
    var lines = [
      '【📦 注文書を登録（Drive自動取込）】',
      '発注書番号: ' + (uploadInfo.documentNo || '—'),
      '種別: '      + (uploadInfo.orderType   || '—'),
    ];
    if (lr.status === 'auto_linked')       lines.push('✅ AI自動紐づけ完了');
    else if (lr.status === 'candidates_found') lines.push('⚠️ 紐づけ候補あり（要確認）');
    else                                   lines.push('❌ 紐づく見積書が見つかりませんでした');
    lines.push('▶ ' + ScriptApp.getService().getUrl());
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ text: lines.join('\n') }),
      muteHttpExceptions: true,
    });
  } catch(e) {
    Logger.log('[_sendOrderRegistrationToChat ERROR] ' + e.message);
  }
}

function _findLatestOrderMgmtId() {
  try {
    var sheet = getSpreadsheet().getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var last  = sheet.getLastRow();
    if (last <= 1) return null;
    var val = sheet.getRange(last, MGMT_COLS.ID).getValue();
    return val ? String(val) : null;
  } catch(e) { return null; }
}