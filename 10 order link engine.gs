// ============================================================
// 10_order_link_engine.gs
// 注文書 ↔ 見積書 紐づけエンジン強化版
// ─ ① 候補取得 API  (_apiGetOrderLinkCandidates)
// ─ ② 紐づけ確定 API (_apiConfirmOrderLink)
// ─ ③ Drive通知連携 (notifyOrderArrivalWithCandidates)
// ─ ④ 注文書アップロード時の自動マッチング
// ============================================================

// ============================================================
// 設定
// ============================================================
var ORDER_LINK_CONFIG = {
  AUTO_LINK_THRESHOLD: 82,    // 自動紐づけスコア閾値
  CANDIDATE_THRESHOLD: 45,    // 候補として表示するスコア最低値
  MAX_CANDIDATES: 5,          // 最大候補数
  PRICE_TOLERANCE_PCT: 10,    // 金額一致と見なす許容誤差(%)
};

// ============================================================
// API: 注文書に対する見積書候補を返す
// handleApiRequest('getOrderLinkCandidates', { mgmtId }) から呼ぶ
// ============================================================
function _apiGetOrderLinkCandidates(p) {
  try {
    var mgmtId = String(p.mgmtId || '').trim();
    if (!mgmtId) return { success: false, error: '管理IDが必要です' };

    var ss = getSpreadsheet();

    // ── 注文書の情報を取得 ──
    var orderInfo = _getOrderInfo(ss, mgmtId);
    if (!orderInfo) return { success: true, candidates: [], linkedQuoteNo: null };

    // すでに紐づけ済みか確認
    var linkedInfo = null;
    if (orderInfo.quoteNo) {
      linkedInfo = _getLinkedQuoteInfo(ss, orderInfo.quoteNo, orderInfo.mgmtId);
    }

    // ── 全見積書を取得してスコアリング ──
    var quotes = _getAllQuotesForMatching(ss);
    var scored = _scoreQuotesForOrder(orderInfo, quotes);

    // スコア閾値でフィルタリング
    var candidates = scored
      .filter(function(c) { return c.score >= ORDER_LINK_CONFIG.CANDIDATE_THRESHOLD; })
      .slice(0, ORDER_LINK_CONFIG.MAX_CANDIDATES);

    // 高スコアが自動紐づけ閾値以上なら自動確定
    if (candidates.length > 0 && candidates[0].score >= ORDER_LINK_CONFIG.AUTO_LINK_THRESHOLD && !orderInfo.quoteNo) {
      try {
        _applyOrderQuoteLink(ss, mgmtId, candidates[0].quoteMgmtId, candidates[0].quoteNo);
        Logger.log('[AUTO LINK] ' + mgmtId + ' → ' + candidates[0].quoteNo + ' (score:' + candidates[0].score + ')');
      } catch(e) {
        Logger.log('[AUTO LINK ERROR] ' + e.message);
      }
    }

    return {
      success: true,
      candidates: candidates,
      linkedQuoteNo: linkedInfo ? linkedInfo.quoteNo : (orderInfo.quoteNo || null),
      linkedMgmtId:  linkedInfo ? linkedInfo.mgmtId : null,
      linkedAmount:  linkedInfo ? linkedInfo.quoteAmount : null,
      linkedPdfUrl:  linkedInfo ? linkedInfo.quotePdfUrl : null,
    };
  } catch(e) {
    Logger.log('[GET CANDIDATES ERROR] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

// ============================================================
// API: 手動で紐づけを確定する
// handleApiRequest('confirmOrderLink', { orderMgmtId, quoteMgmtId, quoteNo }) から呼ぶ
// ============================================================
function _apiConfirmOrderLink(p) {
  try {
    var orderMgmtId = String(p.orderMgmtId || '').trim();
    var quoteMgmtId = String(p.quoteMgmtId || '').trim();
    var quoteNo     = String(p.quoteNo     || '').trim();

    if (!orderMgmtId || !quoteMgmtId) return { success: false, error: '管理IDが必要です' };

    var ss = getSpreadsheet();
    _applyOrderQuoteLink(ss, orderMgmtId, quoteMgmtId, quoteNo);

    return { success: true };
  } catch(e) {
    Logger.log('[CONFIRM LINK ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// 内部: 紐づけを管理シートに書き込む
// ============================================================
function _applyOrderQuoteLink(ss, orderMgmtId, quoteMgmtId, quoteNo) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var last = sheet.getLastRow();
  if (last <= 1) throw new Error('管理シートにデータがありません');

  var data = sheet.getRange(2, 1, last - 1, Object.keys(MGMT_COLS).length + 2).getValues();

  // 注文書行を更新
  data.forEach(function(row, i) {
    var id = String(row[MGMT_COLS.ID - 1] || '').trim();
    if (id !== orderMgmtId) return;

    var rowNum = i + 2;
    sheet.getRange(rowNum, MGMT_COLS.QUOTE_NO).setValue(quoteNo);
    sheet.getRange(rowNum, MGMT_COLS.LINKED).setValue('TRUE');
    sheet.getRange(rowNum, MGMT_COLS.STATUS).setValue('受領');
    sheet.getRange(rowNum, MGMT_COLS.UPDATED_AT).setValue(nowJST());
  });

  // 見積書行を更新（注文書番号を紐づける）
  var orderNo = '';
  data.forEach(function(row) {
    if (String(row[MGMT_COLS.ID - 1]).trim() === orderMgmtId) {
      orderNo = String(row[MGMT_COLS.ORDER_NO - 1] || '').trim();
    }
  });

  if (orderNo) {
    data.forEach(function(row, i) {
      var id = String(row[MGMT_COLS.ID - 1] || '').trim();
      if (id !== quoteMgmtId) return;

      var rowNum = i + 2;
      if (orderNo) sheet.getRange(rowNum, MGMT_COLS.ORDER_NO).setValue(orderNo);
      sheet.getRange(rowNum, MGMT_COLS.LINKED).setValue('TRUE');
      sheet.getRange(rowNum, MGMT_COLS.UPDATED_AT).setValue(nowJST());
    });
  }

  // 「紐づけ候補」シートがあれば処理済みにする
  try {
    var candSheet = ss.getSheetByName('紐づけ候補');
    if (candSheet && candSheet.getLastRow() > 1) {
      var candData = candSheet.getRange(2, 1, candSheet.getLastRow() - 1, 16).getValues();
      candData.forEach(function(r, i) {
        if (String(r[0]) === orderMgmtId) {
          candSheet.getRange(i + 2, 16).setValue('手動確定済み');
          candSheet.getRange(i + 2, 17).setValue(nowJST());
        }
      });
    }
  } catch(e2) {}
}

// ============================================================
// 内部: 注文書情報を取得
// ============================================================
function _getOrderInfo(ss, mgmtId) {
  var mgmtSheet  = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var orderSheet = ss.getSheetByName(CONFIG.SHEET_ORDERS);

  var mgmtData = getAllMgmtData();
  var mgmtRow = mgmtData.find(function(r) {
    return String(r[MGMT_COLS.ID - 1]).trim() === mgmtId;
  });
  if (!mgmtRow) return null;

  // 注文書明細を取得（品名・仕様リスト）
  var orderItems = [];
  var orderAmount = 0;
  if (orderSheet && orderSheet.getLastRow() > 1) {
    var od = orderSheet.getRange(2, 1, orderSheet.getLastRow() - 1, 20).getValues();
    od.forEach(function(r) {
      if (String(r[0]).trim() !== mgmtId) return;
      var itemName = String(r[8] || '').trim();
      var spec     = String(r[9] || '').trim();
      var qty      = Number(r[12]) || 0;
      var unitPrice= Number(r[14]) || 0;
      var amount   = Number(r[15]) || 0;
      if (itemName) {
        orderItems.push({ itemName: itemName, spec: spec, qty: qty, unitPrice: unitPrice, amount: amount });
        orderAmount += amount;
      }
    });
  }

  return {
    mgmtId:      mgmtId,
    orderNo:     String(mgmtRow[MGMT_COLS.ORDER_NO - 1]    || '').trim(),
    quoteNo:     String(mgmtRow[MGMT_COLS.QUOTE_NO - 1]    || '').trim(),
    client:      String(mgmtRow[MGMT_COLS.CLIENT - 1]      || '').trim(),
    subject:     String(mgmtRow[MGMT_COLS.SUBJECT - 1]     || '').trim(),
    modelCode:   String(mgmtRow[MGMT_COLS.MODEL_CODE - 1]  || '').trim(),
    orderDate:   String(mgmtRow[MGMT_COLS.ORDER_DATE - 1]  || '').trim(),
    orderAmount: orderAmount || (Number(mgmtRow[MGMT_COLS.ORDER_AMOUNT - 1]) || 0),
    orderType:   String(mgmtRow[MGMT_COLS.ORDER_TYPE - 1]  || '').trim(),
    pdfUrl:      String(mgmtRow[MGMT_COLS.ORDER_PDF_URL - 1] || '').trim(),
    items:       orderItems,
  };
}

// ============================================================
// 内部: 紐づけ済み見積書情報を取得
// ============================================================
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
    mgmtId:       String(qRow[MGMT_COLS.ID - 1]            || '').trim(),
    quoteNo:      quoteNo,
    quoteAmount:  Number(qRow[MGMT_COLS.QUOTE_AMOUNT - 1]) || 0,
    quotePdfUrl:  String(qRow[MGMT_COLS.QUOTE_PDF_URL - 1] || '').trim(),
  };
}

// ============================================================
// 内部: マッチング用に全見積書を取得
// ============================================================
function _getAllQuotesForMatching(ss) {
  var mgmtData   = getAllMgmtData();
  var quoteSheet = ss.getSheetByName(CONFIG.SHEET_QUOTES);

  // 見積書の明細をmgmtId → items にマップ
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
        amount:    Number(r[QUOTE_COLS.AMOUNT     - 1]) || 0,
      });
    });
  }

  // 重複排除して見積書一覧を構築
  var seen = {};
  var quotes = [];
  mgmtData.forEach(function(r) {
    var qNo = String(r[MGMT_COLS.QUOTE_NO - 1] || '').trim();
    if (!qNo) return;
    if (seen[qNo]) return;
    seen[qNo] = true;

    var mid = String(r[MGMT_COLS.ID - 1] || '').trim();
    quotes.push({
      quoteMgmtId:  mid,
      quoteNo:      qNo,
      client:       String(r[MGMT_COLS.CLIENT      - 1] || '').trim(),
      subject:      String(r[MGMT_COLS.SUBJECT     - 1] || '').trim(),
      modelCode:    String(r[MGMT_COLS.MODEL_CODE  - 1] || '').trim(),
      issueDate:    _toDateStr(r[MGMT_COLS.QUOTE_DATE - 1]),
      quoteAmount:  Number(r[MGMT_COLS.QUOTE_AMOUNT - 1]) || 0,
      quotePdfUrl:  String(r[MGMT_COLS.QUOTE_PDF_URL - 1] || '').trim(),
      destCompany:  String(r[MGMT_COLS.CLIENT - 1] || '').trim(),
      items:        quoteItemsMap[mid] || [],
    });
  });
  return quotes;
}

// ============================================================
// スコアリングロジック（注文書 vs 見積書）
// ============================================================
function _scoreQuotesForOrder(order, quotes) {
  return quotes.map(function(q) {
    var score = 0;
    var reasons = [];

    // ① 顧客名一致
    if (order.client && q.client) {
      var clientSim = _textSimilarity(
        _normalizeCompanyName(order.client),
        _normalizeCompanyName(q.client)
      );
      if (clientSim >= 0.9) { score += 30; reasons.push('顧客名一致'); }
      else if (clientSim >= 0.6) { score += 15; reasons.push('顧客名部分一致'); }
    }

    // ② 機種コード一致
    if (order.modelCode && q.modelCode) {
      if (order.modelCode === q.modelCode) { score += 25; reasons.push('機種コード完全一致'); }
      else if (order.modelCode.indexOf(q.modelCode) >= 0 || q.modelCode.indexOf(order.modelCode) >= 0) {
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
      if (pctDiff <= 1) { score += 20; reasons.push('金額完全一致'); }
      else if (pctDiff <= ORDER_LINK_CONFIG.PRICE_TOLERANCE_PCT) { score += 12; reasons.push('金額ほぼ一致'); }
      else if (pctDiff <= 20) { score += 4; }
    }

    // ⑤ 件名・キーワード一致
    if (order.subject && q.subject) {
      if (_textSimilarity(order.subject, q.subject) >= 0.5) { score += 8; reasons.push('件名類似'); }
    }

    // ⑥ 日付の近さ（注文日が見積日の30日以内であれば加点）
    if (order.orderDate && q.issueDate) {
      try {
        var orderD = new Date(order.orderDate.replace(/\//g,'-'));
        var quoteD = new Date(q.issueDate.replace(/\//g,'-'));
        var dayDiff = (orderD - quoteD) / (1000 * 60 * 60 * 24);
        if (dayDiff >= 0 && dayDiff <= 30) { score += 6; reasons.push('発注日と見積日が近い'); }
        else if (dayDiff >= -7 && dayDiff < 0) { score += 3; }
      } catch(e2) {}
    }

    return {
      quoteMgmtId:  q.quoteMgmtId,
      quoteNo:      q.quoteNo,
      client:       q.client,
      destCompany:  q.destCompany,
      issueDate:    q.issueDate,
      quoteAmount:  q.quoteAmount,
      quotePdfUrl:  q.quotePdfUrl,
      subject:      q.subject,
      score:        Math.min(score, 100),
      reason:       reasons.slice(0, 3).join('・'),
    };
  })
  .filter(function(c) { return c.score > 0; })
  .sort(function(a, b) { return b.score - a.score; });
}

// ============================================================
// 明細品名スコアリング
// ============================================================
function _scoreItemsMatch(orderItems, quoteItems) {
  if (!orderItems.length || !quoteItems.length) return 0;
  var totalScore = 0;
  var matchCount = 0;

  orderItems.forEach(function(oi) {
    var oName = _normItemName(oi.itemName);
    var bestSim = 0;
    var bestQItem = null;

    quoteItems.forEach(function(qi) {
      var qName = _normItemName(qi.itemName);
      var sim = _textSimilarity(oName, qName);
      if (sim > bestSim) {
        bestSim = sim;
        bestQItem = qi;
      }
    });

    if (bestSim >= 0.85) {
      matchCount++;
      totalScore += 10;
      // 単価も一致するか確認
      if (bestQItem && oi.unitPrice > 0 && bestQItem.unitPrice > 0) {
        var priceDiff = Math.abs(oi.unitPrice - bestQItem.unitPrice) / oi.unitPrice * 100;
        if (priceDiff <= 5) totalScore += 5;
      }
    } else if (bestSim >= 0.6) {
      matchCount++;
      totalScore += 5;
    }
  });

  // 全明細が一致した場合にボーナス
  if (orderItems.length > 0 && matchCount === orderItems.length) {
    totalScore += 10;
  }
  return Math.min(totalScore, 40);
}

// ============================================================
// テキスト正規化・類似度計算
// ============================================================
function _normalizeCompanyName(s) {
  return String(s)
    .replace(/株式会社|有限会社|合同会社|合資会社/g, '')
    .replace(/（株）|\(株\)|\（有\）|\(有\)/g, '')
    .replace(/\s+/g, '')
    .replace(/　/g, '')
    .toLowerCase();
}

function _normItemName(s) {
  return String(s)
    .replace(/[ａ-ｚＡ-Ｚ０-９]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); })
    .replace(/[ァ-ン]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0x60); })
    .replace(/\s+/g, '')
    .toLowerCase();
}

function _textSimilarity(a, b) {
  if (!a || !b) return 0;
  if (a === b) return 1;
  if (a.indexOf(b) >= 0 || b.indexOf(a) >= 0) return 0.9;

  // bigram similarity
  var bigramsA = _bigrams(a);
  var bigramsB = _bigrams(b);
  if (!bigramsA.length || !bigramsB.length) return 0;

  var setA = {};
  bigramsA.forEach(function(bg) { setA[bg] = (setA[bg]||0) + 1; });

  var intersect = 0;
  bigramsB.forEach(function(bg) {
    if (setA[bg] && setA[bg] > 0) { intersect++; setA[bg]--; }
  });

  return (2 * intersect) / (bigramsA.length + bigramsB.length);
}

function _bigrams(s) {
  var result = [];
  for (var i = 0; i < s.length - 1; i++) result.push(s.slice(i, i+2));
  return result;
}

// ============================================================
// ④ Drive通知: 注文書到着時に見積書候補を含めて通知する
//    既存の notifyNewOrder / sendTelegramNotification から呼ぶ
// ============================================================
function notifyOrderArrivalWithCandidates(mgmtId, orderNo, client, orderAmount, orderPdfUrl) {
  try {
    var ss = getSpreadsheet();
    var orderInfo = _getOrderInfo(ss, mgmtId);
    if (!orderInfo) {
      Logger.log('[NOTIFY] orderInfo not found for ' + mgmtId);
      return;
    }

    var quotes = _getAllQuotesForMatching(ss);
    var scored = _scoreQuotesForOrder(orderInfo, quotes);
    var top3   = scored
      .filter(function(c) { return c.score >= ORDER_LINK_CONFIG.CANDIDATE_THRESHOLD; })
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
        var scoreEmoji = c.score >= 80 ? '🟢' : (c.score >= 60 ? '🟡' : '🔴');
        msg += '\n' + (i+1) + '. ' + scoreEmoji + ' *見積No: ' + (c.quoteNo||'—') + '*\n';
        msg += '   件名: ' + (c.subject || c.destCompany || '—') + '\n';
        msg += '   発行日: ' + (c.issueDate||'—') + ' ／ 金額: ¥' + (c.quoteAmount ? Number(c.quoteAmount).toLocaleString() : '—') + '\n';
        msg += '   一致度: ' + c.score + '点 (' + (c.reason||'') + ')\n';
        if (c.quotePdfUrl) msg += '   📄 見積書PDF: ' + c.quotePdfUrl + '\n';
      });
    } else {
      msg += '⚠️ 対応する見積書候補が見つかりませんでした。\n';
      msg += '手動で紐づけ画面から確認してください。';
    }

    msg += '\n\n🔗 管理システムで確認・紐づけ:\n' + _getWebAppUrl();

    // Telegram送信
    _sendTelegramIfConfigured(msg);

    // Google Chat Webhook送信
    _sendGoogleChatIfConfigured(msg);

    // スプレッドシートに候補を記録
    _saveMatchingCandidatesToSheet(ss, mgmtId, orderNo, client, orderAmount, top3);

  } catch(e) {
    Logger.log('[NOTIFY ERROR] ' + e.message + '\n' + e.stack);
  }
}

// ============================================================
// Telegram送信（設定済みの場合のみ）
// ============================================================
function _sendTelegramIfConfigured(message) {
  try {
    var token  = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
    var chatId = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID');
    if (!token || !chatId) return;

    var url = 'https://api.telegram.org/bot' + token + '/sendMessage';
    UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({
        chat_id: chatId,
        text: message,
        parse_mode: 'Markdown',
        disable_web_page_preview: true,
      }),
      muteHttpExceptions: true,
    });
  } catch(e) {
    Logger.log('[TELEGRAM ERROR] ' + e.message);
  }
}

// ============================================================
// Google Chat Webhook送信（設定済みの場合のみ）
// ============================================================
function _sendGoogleChatIfConfigured(message) {
  try {
    var webhookUrl = PropertiesService.getScriptProperties().getProperty('GCHAT_WEBHOOK_URL');
    if (!webhookUrl) return;

    UrlFetchApp.fetch(webhookUrl, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({ text: message }),
      muteHttpExceptions: true,
    });
  } catch(e) {
    Logger.log('[GCHAT ERROR] ' + e.message);
  }
}

// ============================================================
// 候補をスプレッドシートに記録
// ============================================================
function _saveMatchingCandidatesToSheet(ss, orderMgmtId, orderNo, client, orderAmount, candidates) {
  try {
    var sheet = ss.getSheetByName('紐づけ候補');
    if (!sheet) {
      // シートがなければ作成
      sheet = ss.insertSheet('紐づけ候補');
      sheet.appendRow([
        '注文管理ID','注文番号','顧客名','発注日','注文金額',
        '候補1_管理ID','候補1_見積No','候補1_顧客','候補1_スコア','候補1_理由',
        '候補2_管理ID','候補2_スコア',
        '候補3_管理ID','候補3_スコア',
        '備考', 'ステータス', '更新日時'
      ]);
    }

    // 既存行を確認（同じ管理IDが既にある場合は更新）
    var lastRow = sheet.getLastRow();
    var existingRowNum = -1;
    if (lastRow > 1) {
      var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
      var idx = ids.map(String).indexOf(orderMgmtId);
      if (idx >= 0) existingRowNum = idx + 2;
    }

    var now = nowJST();
    var c1 = candidates[0] || {};
    var c2 = candidates[1] || {};
    var c3 = candidates[2] || {};

    var rowData = [
      orderMgmtId, orderNo, client, '', orderAmount || '',
      c1.quoteMgmtId||'', c1.quoteNo||'', c1.destCompany||'', c1.score||'', c1.reason||'',
      c2.quoteMgmtId||'', c2.score||'',
      c3.quoteMgmtId||'', c3.score||'',
      '', '未対応', now
    ];

    if (existingRowNum > 0) {
      sheet.getRange(existingRowNum, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
  } catch(e) {
    Logger.log('[SAVE CANDIDATES ERROR] ' + e.message);
  }
}

// ============================================================
// WebアプリURL取得
// ============================================================
function _getWebAppUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch(e) {
    return '（URLを取得できませんでした）';
  }
}

// ============================================================
// 注文書アップロード完了後に自動実行するフック
// 既存の processUploadedPdf / _apiUploadOrderWithLink から呼ぶ
// ============================================================
function onOrderUploaded(mgmtId, orderNo, client, orderAmount, pdfUrl) {
  try {
    Utilities.sleep(500); // スプレッドシートの書き込み完了待ち
    notifyOrderArrivalWithCandidates(mgmtId, orderNo, client, orderAmount, pdfUrl);
  } catch(e) {
    Logger.log('[ON ORDER UPLOADED ERROR] ' + e.message);
  }
}

// ============================================================
// ⑤ 定期バッチ: 未紐づけ注文書を一括マッチング（cronから呼ぶ）
// ============================================================
function batchLinkUnmatchedOrders() {
  var ss = getSpreadsheet();
  var mgmtData = getAllMgmtData();
  var quotes   = _getAllQuotesForMatching(ss);

  var autoCount = 0;
  var skipCount = 0;

  mgmtData.forEach(function(r) {
    var mgmtId  = String(r[MGMT_COLS.ID - 1]      || '').trim();
    var orderNo = String(r[MGMT_COLS.ORDER_NO - 1] || '').trim();
    var quoteNo = String(r[MGMT_COLS.QUOTE_NO - 1] || '').trim();
    var linked  = r[MGMT_COLS.LINKED - 1];
    var isLinked = linked === true || linked === 'TRUE';

    if (!orderNo || quoteNo || isLinked) { skipCount++; return; }

    var orderInfo = _getOrderInfo(ss, mgmtId);
    if (!orderInfo) return;

    var scored = _scoreQuotesForOrder(orderInfo, quotes);
    if (scored.length > 0 && scored[0].score >= ORDER_LINK_CONFIG.AUTO_LINK_THRESHOLD) {
      try {
        _applyOrderQuoteLink(ss, mgmtId, scored[0].quoteMgmtId, scored[0].quoteNo);
        autoCount++;
        Logger.log('[BATCH AUTO LINK] ' + mgmtId + ' → ' + scored[0].quoteNo);
      } catch(e) {
        Logger.log('[BATCH LINK ERROR] ' + e.message);
      }
    }
  });

  Logger.log('[BATCH LINK] 自動紐づけ: ' + autoCount + '件, スキップ: ' + skipCount + '件');
  return { autoLinked: autoCount, skipped: skipCount };
}

// ============================================================
// ⑥ 注文書一覧テーブルへの「紐づけ」ボタン追加用
//    Dashboard.js.html の renderOrderTable から呼ばれる想定
// ============================================================
/**
 * 注文書一覧の行に「紐づけ」列を追加するためのHTML生成
 * @param {Object} item - 注文書オブジェクト
 * @returns {string} HTML
 */
function generateLinkButtonHtml(item) {
  var isLinked = item.linked || item.quoteNo;
  if (isLinked) {
    return '<button class="btn btn-sm" style="background:#e6f4ea;color:#137333;border:1px solid #a5d6a7;border-radius:6px;padding:3px 10px;font-size:11px;cursor:pointer;" ' +
      'onclick="event.stopPropagation();openOrderLinkModal(\'' + item.id + '\', null)">' +
      '✅ ' + (item.quoteNo || '紐づけ済') + '</button>';
  } else {
    return '<button class="btn btn-sm" style="background:#fff3e0;color:#e65100;border:1px solid #ffe0b2;border-radius:6px;padding:3px 10px;font-size:11px;cursor:pointer;" ' +
      'onclick="event.stopPropagation();openOrderLinkModal(\'' + item.id + '\', null)">' +
      '🔗 見積書を紐づける</button>';
  }
}