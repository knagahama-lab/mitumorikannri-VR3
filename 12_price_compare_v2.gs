// ============================================================
// 見積書・注文書管理システム
// ファイル 12: 単価自動比較エンジン（受注判定コア）
// ============================================================
//
// 【依存関係】（既存ファイルの関数を使用）
//   01 config and setup.gs : CONFIG, MGMT_COLS, QUOTE_COLS, ORDER_COLS
//                            getSpreadsheet(), nowJST(), _toDateStr(), getAllMgmtData()
//   03_webapp_ap.gs        : _isLinkedVal()
//   11_db_normalize.gs     : SHEET_QUOTE_DETAIL, SHEET_ORDER_DETAIL
//                            QUOTE_DETAIL_COLS, ORDER_DETAIL_COLS
//                            searchUnitPrice(), updateOrderDetailCompareResult()
//                            _buildCaseSummary()
//
// 【06 board api.gs との関係】
//   既存: apiComparePriceToBOM   → 見積書 vs BOM部品コスト比較（利益率チェック）
//   新規: apiOrderPriceCompare   → 見積書 vs 注文書の単価照合（受注判定）← このファイル
//   ※ 名前が異なるため競合なし
//
// 【03_webapp_ap.gs への追記（handleApiRequest switch文の default 直前に追加）】
//   case 'orderPriceCompare':      res = apiOrderPriceCompare(payload); break;
//   case 'batchOrderPriceCompare': res = apiBatchOrderPriceCompare(payload); break;
//   case 'searchUnitPrice':
//     res = { success: true, results: searchUnitPrice(payload.itemName, payload.spec, payload.client) };
//     break;
//   case 'searchCaseSummary':
//     var kws = String(payload.keywords||'').split(/[\s,　]+/).filter(Boolean);
//     res = { success: true, results: searchCaseSummary(kws) };
//     break;
//   case 'syncDetailDB':    syncAllDetailDB(); res = { success: true }; break;
//   case 'rebuildUnitPrice': rebuildUnitPriceMaster(); res = { success: true }; break;
// ============================================================

var PRICE_COMPARE_CONFIG = {
  PRICE_TOLERANCE_RATE : 0.01,  // 単価許容誤差 1%（端数吸収）
  PRICE_TOLERANCE_ABS  : 1,     // 絶対誤差 1円以内も許容
  ITEM_MATCH_THRESHOLD : 60,    // 品名マッチングスコア閾値（%）
  SHEET_COMPARE_LOG    : '単価比較ログ',
};

var COMPARE_STATUS_VALS = {
  ALL_MATCH    : '✅ 全行一致',
  PARTIAL_MATCH: '⚠️ 一部差異',
  UNMATCHED    : '❌ 未対応品目あり',
  NO_QUOTE     : '📋 見積未紐づけ',
  ERROR        : '🔴 エラー',
};

// ============================================================
// WebアプリAPI（03_webapp_ap.gs の handleApiRequest から呼び出し）
// ============================================================
function apiOrderPriceCompare(p) {
  try {
    var mgmtId = String(p.mgmtId || '').trim();
    if (!mgmtId) return { success: false, error: '管理IDが必要です' };
    var result = comparePriceByMgmtId(mgmtId);
    return { success: true, result: result };
  } catch(e) {
    Logger.log('[COMPARE] ERROR: ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

function apiBatchOrderPriceCompare(p) {
  try {
    var result = batchComparePrices();
    return { success: true, result: result };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// 1件処理：管理IDを指定して比較実行
// ============================================================
function comparePriceByMgmtId(mgmtId) {
  Logger.log('[COMPARE] 開始: ' + mgmtId);

  var ss        = getSpreadsheet();
  var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var mgmtRow   = _getMgmtRowById(mgmtSheet, mgmtId);

  if (!mgmtRow) {
    return _makeErrorResult(mgmtId, '管理IDが見つかりません: ' + mgmtId);
  }

  var orderNo     = String(mgmtRow[MGMT_COLS.ORDER_NO    - 1] || '').trim();
  var linkedQuote = String(mgmtRow[MGMT_COLS.QUOTE_NO    - 1] || '').trim();
  var client      = String(mgmtRow[MGMT_COLS.CLIENT      - 1] || '').trim();
  var subject     = String(mgmtRow[MGMT_COLS.SUBJECT     - 1] || '').trim();
  var orderType   = String(mgmtRow[MGMT_COLS.ORDER_TYPE  - 1] || '').trim();

  if (!orderNo) {
    return _makeErrorResult(mgmtId, '注文番号がありません。注文書が登録されていない可能性があります。');
  }
  if (!linkedQuote) {
    return {
      mgmtId: mgmtId, orderNo: orderNo, client: client, subject: subject,
      status: COMPARE_STATUS_VALS.NO_QUOTE,
      message: '見積書が紐づいていません。先に紐づけマッチングを実行してください。',
      lineResults: [], canAutoOrder: false,
    };
  }

  var quoteLines = _getQuoteLines(ss, linkedQuote);
  var orderLines = _getOrderLines(ss, orderNo);

  if (quoteLines.length === 0) {
    return _makeErrorResult(mgmtId,
      '見積書（' + linkedQuote + '）の明細データが見つかりません。OCR処理済みか確認してください。');
  }
  if (orderLines.length === 0) {
    return _makeErrorResult(mgmtId,
      '注文書（' + orderNo + '）の明細データが見つかりません。OCR処理済みか確認してください。');
  }

  var lineResults = _compareLines(orderLines, quoteLines, client);

  var hasUnmatched = lineResults.some(function(r) { return r.matchStatus === 'unmatched'; });
  var hasDiff      = lineResults.some(function(r) { return r.matchStatus === 'price_diff'; });
  var allOk        = lineResults.every(function(r) { return r.matchStatus === 'ok'; });

  var overallStatus = allOk              ? COMPARE_STATUS_VALS.ALL_MATCH
                    : hasUnmatched       ? COMPARE_STATUS_VALS.UNMATCHED
                    : hasDiff            ? COMPARE_STATUS_VALS.PARTIAL_MATCH
                    :                      COMPARE_STATUS_VALS.PARTIAL_MATCH;

  var canAutoOrder = allOk;

  var result = {
    mgmtId      : mgmtId,
    orderNo     : orderNo,
    quoteNo     : linkedQuote,
    client      : client,
    subject     : subject,
    orderType   : orderType,
    status      : overallStatus,
    canAutoOrder: canAutoOrder,
    lineResults : lineResults,
    summary     : _buildSummaryText(lineResults, overallStatus),
    comparedAt  : nowJST(),
  };

  _writeCompareResultToSheets(ss, mgmtId, orderNo, lineResults, overallStatus);

  if (canAutoOrder) {
    _autoRegisterOrder(mgmtSheet, mgmtId, result);
  } else {
    _sendDiffAlert(result);
  }

  _appendCompareLog(ss, result);

  Logger.log('[COMPARE] 完了: ' + mgmtId + ' → ' + overallStatus);
  return result;
}

// ============================================================
// バッチ処理（未比較の全注文書を一括実行）
// ============================================================
function batchComparePrices() {
  var data = getAllMgmtData(); // 01 config and setup.gs の関数
  // MGMT_COLS.LINKED = 18（列定義で確認済み）
  var LINKED_COL_IDX = MGMT_COLS.LINKED - 1; // 0ベース = 17

  var targets = data.filter(function(row) {
    var mgmtId  = String(row[MGMT_COLS.ID      - 1] || '');
    var orderNo = String(row[MGMT_COLS.ORDER_NO - 1] || '');
    var status  = String(row[MGMT_COLS.STATUS   - 1] || '');
    return orderNo && mgmtId &&
           status !== CONFIG.STATUS.ORDERED &&
           status !== CONFIG.STATUS.DELIVERED;
  });

  Logger.log('[BATCH_COMPARE] 対象: ' + targets.length + '件');

  var results = [];
  targets.forEach(function(row) {
    var mgmtId = String(row[MGMT_COLS.ID - 1] || '');
    try {
      var r = comparePriceByMgmtId(mgmtId);
      results.push({ mgmtId: mgmtId, status: r.status, canAutoOrder: r.canAutoOrder });
      Utilities.sleep(300);
    } catch(e) {
      results.push({ mgmtId: mgmtId, status: 'ERROR', error: e.message });
    }
  });

  var autoOrderCount = results.filter(function(r) { return r.canAutoOrder; }).length;
  var diffCount      = results.filter(function(r) { return !r.canAutoOrder && r.status !== 'ERROR'; }).length;

  Logger.log('[BATCH_COMPARE] 完了: 自動受注=' + autoOrderCount + '件, 差異あり=' + diffCount + '件');
  return { processed: results.length, autoOrdered: autoOrderCount, diffFound: diffCount, results: results };
}

// ============================================================
// 明細行マッチング・比較ロジック
// ============================================================
function _compareLines(orderLines, quoteLines, client) {
  var results = [];

  orderLines.forEach(function(orderLine) {
    var best = null;
    var bestScore = 0;

    quoteLines.forEach(function(quoteLine) {
      var score = _calcItemMatchScore(orderLine, quoteLine);
      if (score > bestScore) { bestScore = score; best = quoteLine; }
    });

    if (bestScore < PRICE_COMPARE_CONFIG.ITEM_MATCH_THRESHOLD) {
      var lineResult = _makeLineResult(orderLine, null, 'unmatched', bestScore,
        '⚠️ 対応する見積明細が見つかりません（類似度' + bestScore + '%）');

      // 単価マスタから補足候補を検索（11_db_normalize.gs の関数）
      var masterHits = searchUnitPrice(orderLine.itemName, orderLine.spec, client);
      if (masterHits.length > 0) {
        lineResult.masterSuggestion = masterHits[0];
        lineResult.message += '。単価マスタに類似品あり: ' +
          masterHits[0].itemName + ' ¥' + masterHits[0].unitPrice.toLocaleString();
      }
      results.push(lineResult);
      return;
    }

    var orderPrice = _toNumber(orderLine.unitPrice);
    var quotePrice = _toNumber(best.unitPrice);

    if (orderPrice === null || quotePrice === null) {
      results.push(_makeLineResult(orderLine, best, 'no_price', bestScore,
        '単価データなし（OCR未読み取りの可能性）'));
      return;
    }

    var diff     = orderPrice - quotePrice;
    var diffRate = quotePrice > 0 ? Math.abs(diff) / quotePrice : 0;
    var withinTolerance =
      Math.abs(diff) <= PRICE_COMPARE_CONFIG.PRICE_TOLERANCE_ABS ||
      diffRate <= PRICE_COMPARE_CONFIG.PRICE_TOLERANCE_RATE;

    if (withinTolerance) {
      results.push(_makeLineResult(orderLine, best, 'ok', bestScore,
        '✅ 単価一致（見積:¥' + quotePrice.toLocaleString() +
        ' / 注文:¥' + orderPrice.toLocaleString() + '）'));
    } else {
      var diffStr = (diff > 0 ? '+' : '-') + '¥' + Math.abs(diff).toLocaleString();
      results.push(_makeLineResult(orderLine, best, 'price_diff', bestScore,
        '❌ 単価差異（見積:¥' + quotePrice.toLocaleString() +
        ' → 注文:¥' + orderPrice.toLocaleString() +
        '、差額:' + diffStr + '、差率:' + (diffRate * 100).toFixed(1) + '%）',
        { quotePrice: quotePrice, orderPrice: orderPrice, diff: diff, diffRate: diffRate }
      ));
    }
  });

  return results;
}

function _makeLineResult(orderLine, quoteLine, matchStatus, score, message, diffDetail) {
  return {
    lineNo        : orderLine.lineNo,
    itemName      : orderLine.itemName,
    spec          : orderLine.spec,
    qty           : orderLine.qty,
    orderUnitPrice: orderLine.unitPrice,
    quoteItemName : quoteLine ? quoteLine.itemName  : '',
    quoteUnitPrice: quoteLine ? quoteLine.unitPrice : '',
    matchScore    : score,
    matchStatus   : matchStatus,  // ok / price_diff / unmatched / no_price
    message       : message,
    diffDetail    : diffDetail || null,
  };
}

/**
 * 品名マッチングスコアを計算（0-100）
 */
function _calcItemMatchScore(orderLine, quoteLine) {
  var oName = _normItemName(String(orderLine.itemName || ''));
  var qName = _normItemName(String(quoteLine.itemName || ''));
  var oSpec = _normItemName(String(orderLine.spec || ''));
  var qSpec = _normItemName(String(quoteLine.spec || ''));

  if (!oName || !qName) return 0;

  var nameScore = 0;
  if (oName === qName)                             nameScore = 100;
  else if (oName.indexOf(qName) >= 0 || qName.indexOf(oName) >= 0) nameScore = 80;
  else nameScore = _lcsRatio(oName, qName) * 100;

  var specBonus = 0;
  if (oSpec && qSpec) {
    if (oSpec === qSpec)                                        specBonus = 20;
    else if (oSpec.indexOf(qSpec) >= 0 || qSpec.indexOf(oSpec) >= 0) specBonus = 10;
  }

  var qtyBonus = 0;
  if (orderLine.qty && quoteLine.qty &&
      String(orderLine.qty) === String(quoteLine.qty)) qtyBonus = 5;

  return Math.min(100, Math.round(nameScore * 0.8 + specBonus + qtyBonus));
}

/** 最長共通部分列の長さ比（LCS ratio） */
function _lcsRatio(a, b) {
  if (!a || !b) return 0;
  var m = a.length, n = b.length;
  if (m > 30 || n > 30) {
    // 長文は簡易共通文字数スコア
    var common = 0, bArr = b.split('');
    a.split('').forEach(function(c) {
      var i = bArr.indexOf(c);
      if (i >= 0) { common++; bArr.splice(i, 1); }
    });
    return common / Math.max(m, n);
  }
  var dp = [];
  for (var i = 0; i <= m; i++) { dp[i] = []; for (var j = 0; j <= n; j++) dp[i][j] = 0; }
  for (var i = 1; i <= m; i++) {
    for (var j = 1; j <= n; j++) {
      dp[i][j] = a[i-1] === b[j-1] ? dp[i-1][j-1] + 1 : Math.max(dp[i-1][j], dp[i][j-1]);
    }
  }
  return dp[m][n] / Math.max(m, n);
}

/**
 * 品名の正規化（11_db_normalize.gs の _normItemName を使用）
 * ※ここでは再定義せず、11_db_normalize.gsの関数をそのまま呼び出します
 */
// function _normItemName は 11_db_normalize.gs で定義済み

// ============================================================
// 自動受注登録
// ============================================================
function _autoRegisterOrder(mgmtSheet, mgmtId, compareResult) {
  Logger.log('[AUTO_ORDER] 受注自動登録: ' + mgmtId);
  var rowIdx = _getMgmtRowIndex(mgmtSheet, mgmtId);
  if (rowIdx > 0) {
    mgmtSheet.getRange(rowIdx, MGMT_COLS.STATUS    ).setValue(CONFIG.STATUS.ORDERED);
    mgmtSheet.getRange(rowIdx, MGMT_COLS.UPDATED_AT).setValue(nowJST());
  }
  _sendOrderConfirmNotification(compareResult);
  Logger.log('[AUTO_ORDER] 完了: ' + mgmtId + ' → ' + CONFIG.STATUS.ORDERED);
}

// ============================================================
// スプレッドシートへの書き込み
// ============================================================
function _writeCompareResultToSheets(ss, mgmtId, orderNo, lineResults, overallStatus) {
  lineResults.forEach(function(line) {
    var detailJson = JSON.stringify({
      matchScore   : line.matchScore,
      matchStatus  : line.matchStatus,
      quoteItemName: line.quoteItemName,
      quoteUnitPrice: line.quoteUnitPrice,
      diffDetail   : line.diffDetail,
    });
    // 11_db_normalize.gsの関数
    updateOrderDetailCompareResult(orderNo, line.lineNo, line.matchStatus, detailJson);
  });
  _buildCaseSummary();
}

// ============================================================
// 通知
// ============================================================
function _sendOrderConfirmNotification(result) {
  var subject = '✅ 【自動受注登録完了】' + result.subject + '（' + result.client + '）';
  var body = [
    '以下の案件が自動的に受注登録されました。',
    '',
    '管理ID: ' + result.mgmtId,
    '顧客名: ' + result.client,
    '件名:   ' + result.subject,
    '見積No: ' + result.quoteNo,
    '注文No: ' + result.orderNo,
    '注文種別:' + result.orderType,
    '',
    '■ 単価比較結果',
    result.summary,
    '',
    '次のステップ: 製品発注の準備を開始してください。',
    '確認日時: ' + result.comparedAt,
  ].join('\n');

  try {
    GmailApp.sendEmail(Session.getActiveUser().getEmail(), subject, body);
  } catch(e) {
    Logger.log('[NOTIFY] メール送信失敗: ' + e.message);
  }

  // Google Chat通知（CONFIG.GOOGLE_CHAT_WEBHOOK_URL が設定されている場合）
  var webhookUrl = CONFIG.GOOGLE_CHAT_WEBHOOK_URL ||
    PropertiesService.getScriptProperties().getProperty('GOOGLE_CHAT_WEBHOOK_URL') || '';
  if (webhookUrl) {
    try {
      UrlFetchApp.fetch(webhookUrl, {
        method: 'post', contentType: 'application/json',
        payload: JSON.stringify({
          text: '✅ *受注自動登録完了*\n顧客: ' + result.client +
                '\n件名: ' + result.subject +
                '\n見積No: ' + result.quoteNo + ' / 注文No: ' + result.orderNo
        }),
        muteHttpExceptions: true,
      });
    } catch(e) {
      Logger.log('[NOTIFY] Chat送信失敗: ' + e.message);
    }
  }
}

function _sendDiffAlert(result) {
  var diffLines = result.lineResults.filter(function(r) {
    return r.matchStatus === 'price_diff' || r.matchStatus === 'unmatched';
  });
  if (!diffLines.length) return;

  var subject = '⚠️ 【単価差異あり・要確認】' + result.subject + '（' + result.client + '）';
  var body = [
    '単価差異または未対応品目が見つかりました。確認をお願いします。',
    '',
    '管理ID: ' + result.mgmtId,
    '顧客名: ' + result.client,
    '件名:   ' + result.subject,
    '見積No: ' + result.quoteNo,
    '注文No: ' + result.orderNo,
    '',
    '■ 差異・未対応の明細行 (' + diffLines.length + '行)',
    diffLines.map(function(r) {
      return '  [行' + r.lineNo + '] ' + r.itemName + '\n    ' + r.message;
    }).join('\n'),
    '',
    '確認日時: ' + result.comparedAt,
  ].join('\n');

  try {
    GmailApp.sendEmail(Session.getActiveUser().getEmail(), subject, body);
  } catch(e) {
    Logger.log('[NOTIFY] 差異アラートメール送信失敗: ' + e.message);
  }
}

// ============================================================
// 比較ログシート
// ============================================================
function _appendCompareLog(ss, result) {
  var logSheet = ss.getSheetByName(PRICE_COMPARE_CONFIG.SHEET_COMPARE_LOG);
  if (!logSheet) {
    logSheet = ss.insertSheet(PRICE_COMPARE_CONFIG.SHEET_COMPARE_LOG);
    var headers = ['日時','管理ID','顧客名','件名','見積番号','注文番号',
                   '判定結果','自動受注','一致行数','差異行数','未マッチ行数'];
    logSheet.getRange(1, 1, 1, headers.length).setValues([headers])
            .setBackground('#ECEFF1').setFontWeight('bold');
    logSheet.setFrozenRows(1);
  }

  var okCount      = result.lineResults.filter(function(r) { return r.matchStatus === 'ok'; }).length;
  var diffCount    = result.lineResults.filter(function(r) { return r.matchStatus === 'price_diff'; }).length;
  var unmatchCount = result.lineResults.filter(function(r) { return r.matchStatus === 'unmatched'; }).length;

  logSheet.appendRow([
    result.comparedAt, result.mgmtId, result.client, result.subject,
    result.quoteNo, result.orderNo, result.status,
    result.canAutoOrder ? '✅ 自動受注' : '手動確認要',
    okCount, diffCount, unmatchCount,
  ]);
}

// ============================================================
// データ取得ユーティリティ
// ============================================================
function _getQuoteLines(ss, quoteNo) {
  // まず見積明細DBを参照（11_db_normalize.gs が存在する場合）
  var dstSheet = ss.getSheetByName(SHEET_QUOTE_DETAIL);
  if (dstSheet && dstSheet.getLastRow() > 1) {
    var ncols = Object.keys(QUOTE_DETAIL_COLS).length;
    var data  = dstSheet.getRange(2, 1, dstSheet.getLastRow() - 1, ncols).getValues();
    var lines = data.filter(function(row) {
      return String(row[QUOTE_DETAIL_COLS.QUOTE_NO - 1] || '') === quoteNo;
    }).map(function(row) {
      return {
        lineNo   : row[QUOTE_DETAIL_COLS.LINE_NO    - 1],
        itemName : String(row[QUOTE_DETAIL_COLS.ITEM_NAME  - 1] || ''),
        spec     : String(row[QUOTE_DETAIL_COLS.SPEC       - 1] || ''),
        qty      : row[QUOTE_DETAIL_COLS.QTY        - 1],
        unit     : String(row[QUOTE_DETAIL_COLS.UNIT - 1] || ''),
        unitPrice: row[QUOTE_DETAIL_COLS.UNIT_PRICE - 1],
        amount   : row[QUOTE_DETAIL_COLS.AMOUNT     - 1],
      };
    });
    if (lines.length > 0) return lines;
  }

  // フォールバック: 見積書シート（CONFIG.SHEET_QUOTES = '見積書シート'）を直接参照
  var srcSheet = ss.getSheetByName(CONFIG.SHEET_QUOTES);
  if (!srcSheet || srcSheet.getLastRow() <= 1) return [];
  return srcSheet.getRange(2, 1, srcSheet.getLastRow() - 1, 15).getValues()
    .filter(function(r) { return String(r[QUOTE_COLS.QUOTE_NO - 1] || '') === quoteNo; })
    .map(function(r) {
      return {
        lineNo   : r[QUOTE_COLS.LINE_NO    - 1],
        itemName : String(r[QUOTE_COLS.ITEM_NAME  - 1] || ''),
        spec     : String(r[QUOTE_COLS.SPEC       - 1] || ''),
        qty      : r[QUOTE_COLS.QTY        - 1],
        unit     : String(r[QUOTE_COLS.UNIT - 1] || ''),
        unitPrice: r[QUOTE_COLS.UNIT_PRICE - 1],
        amount   : r[QUOTE_COLS.AMOUNT     - 1],
      };
    });
}

function _getOrderLines(ss, orderNo) {
  // まず注文明細DBを参照
  var dstSheet = ss.getSheetByName(SHEET_ORDER_DETAIL);
  if (dstSheet && dstSheet.getLastRow() > 1) {
    var ncols = Object.keys(ORDER_DETAIL_COLS).length;
    var data  = dstSheet.getRange(2, 1, dstSheet.getLastRow() - 1, ncols).getValues();
    var lines = data.filter(function(row) {
      return String(row[ORDER_DETAIL_COLS.ORDER_NO - 1] || '') === orderNo;
    }).map(function(row) {
      return {
        lineNo   : row[ORDER_DETAIL_COLS.LINE_NO    - 1],
        itemName : String(row[ORDER_DETAIL_COLS.ITEM_NAME  - 1] || ''),
        spec     : String(row[ORDER_DETAIL_COLS.SPEC       - 1] || ''),
        qty      : row[ORDER_DETAIL_COLS.QTY        - 1],
        unit     : String(row[ORDER_DETAIL_COLS.UNIT - 1] || ''),
        unitPrice: row[ORDER_DETAIL_COLS.UNIT_PRICE - 1],
        amount   : row[ORDER_DETAIL_COLS.AMOUNT     - 1],
      };
    });
    if (lines.length > 0) return lines;
  }

  // フォールバック: 注文書シート（CONFIG.SHEET_ORDERS = '注文書シート'）を直接参照
  var srcSheet = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  if (!srcSheet || srcSheet.getLastRow() <= 1) return [];
  return srcSheet.getRange(2, 1, srcSheet.getLastRow() - 1, 19).getValues()
    .filter(function(r) { return String(r[ORDER_COLS.ORDER_NO - 1] || '') === orderNo; })
    .map(function(r) {
      return {
        lineNo   : r[ORDER_COLS.LINE_NO    - 1],
        itemName : String(r[ORDER_COLS.ITEM_NAME  - 1] || ''),
        spec     : String(r[ORDER_COLS.SPEC       - 1] || ''),
        qty      : r[ORDER_COLS.QTY        - 1],
        unit     : String(r[ORDER_COLS.UNIT - 1] || ''),
        unitPrice: r[ORDER_COLS.UNIT_PRICE - 1],
        amount   : r[ORDER_COLS.AMOUNT     - 1],
      };
    });
}

/**
 * 管理シートから指定IDの行データを取得
 * ※ 06 board api.gs で呼ばれているが未定義だったため、ここで定義
 */
function _getMgmtRowById(sheet, mgmtId) {
  if (!sheet || sheet.getLastRow() <= 1) return null;
  var actualCols = Math.max(sheet.getLastColumn(), 27);
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, actualCols).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][MGMT_COLS.ID - 1] || '') === mgmtId) return data[i];
  }
  return null;
}

function _getMgmtRowIndex(sheet, mgmtId) {
  if (!sheet || sheet.getLastRow() <= 1) return -1;
  var ids = sheet.getRange(2, MGMT_COLS.ID, sheet.getLastRow() - 1, 1).getValues().flat();
  var idx = ids.map(String).indexOf(String(mgmtId));
  return idx >= 0 ? idx + 2 : -1;
}

function _buildSummaryText(lineResults, overallStatus) {
  var lines = ['総合判定: ' + overallStatus, ''];
  lineResults.forEach(function(r) {
    lines.push('[行' + r.lineNo + '] ' + r.itemName +
               (r.spec ? '（' + r.spec + '）' : '') +
               '\n  → ' + r.message);
  });
  return lines.join('\n');
}

function _makeErrorResult(mgmtId, message) {
  return { mgmtId: mgmtId, status: COMPARE_STATUS_VALS.ERROR,
           message: message, lineResults: [], canAutoOrder: false };
}

function _toNumber(v) {
  if (v === '' || v === null || v === undefined) return null;
  var n = Number(String(v).replace(/[,¥￥\s]/g, ''));
  return isNaN(n) ? null : n;
}

// ============================================================
// デバッグ用
// ============================================================
function testCompareLatestOrder() {
  var data = getAllMgmtData();
  var latest = null;
  for (var i = data.length - 1; i >= 0; i--) {
    if (String(data[i][MGMT_COLS.ORDER_NO - 1] || '').trim()) {
      latest = data[i]; break;
    }
  }
  if (!latest) { Logger.log('注文書のある案件がありません'); return; }
  var mgmtId = String(latest[MGMT_COLS.ID - 1] || '');
  Logger.log('テスト対象管理ID: ' + mgmtId);
  var result = comparePriceByMgmtId(mgmtId);
  Logger.log('結果: ' + result.status);
  Logger.log('自動受注可: ' + result.canAutoOrder);
  result.lineResults.forEach(function(r) {
    Logger.log('[行' + r.lineNo + '] ' + r.matchStatus + ' ' + r.message);
  });
}
