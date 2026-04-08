// ============================================================
// 見積書・注文書管理システム
// ファイル 5/5: 紐づけエンジン（スコアリング方式）
// ============================================================

var MATCH_CONFIG = {
  AUTO_LINK_THRESHOLD:         80,
  AUTO_LINK_THRESHOLD_NOLINES: 45,
  CANDIDATE_THRESHOLD:         35,
  DATE_RANGE_DAYS:             90,
  AMOUNT_TOLERANCE:            0.15,
  QTY_TOLERANCE:               0.10,
  SHEET_CANDIDATES:            '紐づけ候補',
};

function _isLinked(val) {
  return val === true || val === 'TRUE' || val === 'true';
}

// ============================================================
// メイン処理
// ============================================================

function matchOrderToQuote(orderMgmtId) {
  var ss         = getSpreadsheet();
  var mgmtSheet  = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var orderSheet = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  var quoteSheet = ss.getSheetByName(CONFIG.SHEET_QUOTES);

  var orderMgmt = _getMgmtRowById(mgmtSheet, orderMgmtId);
  if (!orderMgmt) {
    Logger.log('[MATCH] 管理ID未発見: ' + orderMgmtId);
    return { success: false, error: '管理ID未発見: ' + orderMgmtId };
  }

  if (_isLinked(orderMgmt[MGMT_COLS.LINKED - 1])) {
    return { success: true, status: 'already_linked' };
  }

  var orderLines = _getLinesById(orderSheet, orderMgmtId, 19);
  Logger.log('[MATCH] ' + orderMgmtId + ' 明細:' + orderLines.length + '行');

  var allMgmt    = getAllMgmtData();
  var candidates = [];

  for (var i = 0; i < allMgmt.length; i++) {
    var row         = allMgmt[i];
    var quoteMgmtId = String(row[MGMT_COLS.ID - 1] || '');
    if (quoteMgmtId === orderMgmtId) continue;
    if (!String(row[MGMT_COLS.QUOTE_NO - 1])) continue;
    if (_isLinked(row[MGMT_COLS.LINKED - 1])) continue;

    var quoteLines    = _getLinesById(quoteSheet, quoteMgmtId, 15);
    var score         = _calcMatchScore(orderMgmt, orderLines, row, quoteLines);
    var candThreshold = orderLines.length > 0
      ? MATCH_CONFIG.CANDIDATE_THRESHOLD
      : Math.floor(MATCH_CONFIG.CANDIDATE_THRESHOLD * 0.6);

    if (score.total >= candThreshold) {
      candidates.push({
        quoteMgmtId: quoteMgmtId,
        quoteNo:     String(row[MGMT_COLS.QUOTE_NO - 1] || ''),
        client:      String(row[MGMT_COLS.CLIENT - 1] || ''),
        quoteDate:   _toDateStr(row[MGMT_COLS.QUOTE_DATE - 1]),
        score:       score.total,
        breakdown:   score.breakdown,
      });
    }
  }

  candidates.sort(function(a, b) { return b.score - a.score; });
  Logger.log('[MATCH] 候補:' + candidates.length + '件 最高:' + (candidates[0] ? candidates[0].score : 0) + '点');

  if (candidates.length === 0) {
    _saveCandidates(orderMgmtId, orderMgmt, []);
    return { success: true, status: 'no_candidates', candidates: [] };
  }

  var best          = candidates[0];
  var autoThreshold = orderLines.length > 0
    ? MATCH_CONFIG.AUTO_LINK_THRESHOLD
    : MATCH_CONFIG.AUTO_LINK_THRESHOLD_NOLINES;

  if (best.score >= autoThreshold) {
    _executeLink(mgmtSheet, orderMgmtId, best.quoteMgmtId, best.score, '自動');
    return { success: true, status: 'auto_linked', quoteMgmtId: best.quoteMgmtId,
             quoteNo: best.quoteNo, score: best.score, breakdown: best.breakdown };
  }

  _saveCandidates(orderMgmtId, orderMgmt, candidates.slice(0, 5));
  return { success: true, status: 'candidates_found', candidates: candidates.slice(0, 5) };
}

function batchMatchAllUnlinked() {
  Logger.log('[BATCH MATCH] 開始');
  var all            = getAllMgmtData();
  var autoCount      = 0;
  var candidateCount = 0;
  var noMatchCount   = 0;
  var processed      = 0;

  for (var i = 0; i < all.length; i++) {
    var row     = all[i];
    var orderNo = String(row[MGMT_COLS.ORDER_NO - 1] || '');
    if (!orderNo) continue;
    if (_isLinked(row[MGMT_COLS.LINKED - 1])) continue;
    var mgmtId = String(row[MGMT_COLS.ID - 1] || '');
    if (!mgmtId) continue;

    processed++;
    var result = matchOrderToQuote(mgmtId);
    if (result.status === 'auto_linked')           autoCount++;
    else if (result.status === 'candidates_found') candidateCount++;
    else                                           noMatchCount++;

    if (processed % 10 === 0) Utilities.sleep(1000);
  }

  Logger.log('[BATCH MATCH] 完了 自動:' + autoCount + ' 候補:' + candidateCount + ' 対象外:' + noMatchCount + ' 合計:' + processed);
  return { autoCount: autoCount, candidateCount: candidateCount, noMatchCount: noMatchCount };
}

// ============================================================
// スコアリング
// ============================================================

function _calcMatchScore(orderMgmt, orderLines, quoteMgmt, quoteLines) {
  var bd = { keyword: 0, qty: 0, amount: 0, date: 0, client: 0, model: 0 };

  // S1: 品名・仕様キーワード（最大40点）
  if (orderLines.length > 0 && quoteLines.length > 0) {
    var oKws     = _extractKeywordsFromLines(orderLines, [ORDER_COLS.ITEM_NAME - 1, ORDER_COLS.SPEC - 1]);
    var qKws     = _extractKeywordsFromLines(quoteLines, [6, 7]);
    bd.keyword   = Math.round(_keywordMatchScore(oKws, qKws) * 40);
  } else if (orderLines.length === 0 && quoteLines.length > 0) {
    var qKws2    = _extractKeywordsFromLines(quoteLines, [6, 7]);
    var oClKws   = _extractKeywordsFromLines([[orderMgmt[MGMT_COLS.CLIENT - 1] || '']], [0]);
    bd.keyword   = Math.round(_keywordMatchScore(oClKws, qKws2) * 15);
  }

  // S2: 数量（最大20点、明細がある場合のみ）
  if (orderLines.length > 0 && quoteLines.length > 0) {
    var oQtys  = _extractNums(orderLines, ORDER_COLS.QTY - 1);
    var qQtys  = _extractNums(quoteLines, 8);
    bd.qty     = Math.round(_numSetMatchScore(oQtys, qQtys, MATCH_CONFIG.QTY_TOLERANCE) * 20);
  }

  // S3: 金額（最大15点）
  var oAmt = Number(orderMgmt[MGMT_COLS.ORDER_AMOUNT - 1]) || 0;
  var qAmt = Number(quoteMgmt[MGMT_COLS.QUOTE_AMOUNT - 1]) || 0;
  if (oAmt > 0 && qAmt > 0) {
    var diff = Math.abs(oAmt - qAmt) / Math.max(oAmt, qAmt);
    if (diff <= MATCH_CONFIG.AMOUNT_TOLERANCE) {
      bd.amount = Math.round((1 - diff / MATCH_CONFIG.AMOUNT_TOLERANCE) * 15);
    }
  }

  // S4: 日付近接（最大10点）
  var oDate = _parseDate(orderMgmt[MGMT_COLS.ORDER_DATE - 1]);
  var qDate = _parseDate(quoteMgmt[MGMT_COLS.QUOTE_DATE - 1]);
  if (oDate && qDate) {
    var days = (oDate - qDate) / 86400000;
    if (days >= 0 && days <= MATCH_CONFIG.DATE_RANGE_DAYS) {
      bd.date = Math.round((1 - days / MATCH_CONFIG.DATE_RANGE_DAYS) * 10);
    } else if (days >= -7 && days < 0) {
      bd.date = 3;
    }
  }

  // S5: 顧客名（最大10点）
  var oClient = _normalizeCompany(String(orderMgmt[MGMT_COLS.CLIENT - 1] || ''));
  var qClient = _normalizeCompany(String(quoteMgmt[MGMT_COLS.CLIENT - 1] || ''));
  if (oClient && qClient) {
    if (oClient === qClient) {
      bd.client = 10;
    } else if (oClient.indexOf(qClient) >= 0 || qClient.indexOf(oClient) >= 0) {
      bd.client = 6;
    } else if (_longestCommonSubstr(oClient, qClient) >= 2) {
      bd.client = 3;
    }
  }

  // S6: 機種コード（ボーナス最大15点）
  var oModel = String(orderMgmt[MGMT_COLS.MODEL_CODE - 1] || '').trim();
  var qModel = String(quoteMgmt[MGMT_COLS.MODEL_CODE - 1] || '').trim();
  if (oModel && qModel) {
    if (oModel === qModel) {
      bd.model = 15;
    } else if (oModel.indexOf(qModel) >= 0 || qModel.indexOf(oModel) >= 0) {
      bd.model = 8;
    }
  }

  return { total: bd.keyword + bd.qty + bd.amount + bd.date + bd.client + bd.model, breakdown: bd };
}

// ============================================================
// キーワード・数値抽出ヘルパー
// ============================================================

function _extractKeywordsFromLines(lines, colIndices) {
  var seen     = {};
  var keywords = [];
  lines.forEach(function(line) {
    colIndices.forEach(function(idx) {
      var val = String(line[idx] || '').trim();
      if (!val) return;
      val.split(/[\s\-_\(\)（）・\/＋]+/)
        .map(function(t) { return t.trim().toLowerCase(); })
        .filter(function(t) { return t.length >= 2; })
        .forEach(function(t) {
          if (!seen[t]) { seen[t] = true; keywords.push(t); }
        });
    });
  });
  return keywords;
}

function _keywordMatchScore(set1, set2) {
  if (!set1.length || !set2.length) return 0;
  var matched = 0;
  set1.forEach(function(k) {
    if (set2.some(function(q) { return q.indexOf(k) >= 0 || k.indexOf(q) >= 0; })) matched++;
  });
  return matched / Math.max(set1.length, set2.length);
}

function _extractNums(lines, colIdx) {
  return lines.map(function(l) { return Number(l[colIdx]) || 0; })
              .filter(function(n) { return n > 0; });
}

function _numSetMatchScore(set1, set2, tol) {
  if (!set1.length || !set2.length) return 0;
  var matched = 0;
  set1.forEach(function(n1) {
    if (set2.some(function(n2) { return Math.abs(n1 - n2) / Math.max(n1, n2) <= tol; })) matched++;
  });
  return matched / Math.max(set1.length, set2.length);
}

function _normalizeCompany(name) {
  return name.replace(/株式会社|有限会社|合同会社|（株）|\(株\)|（有）|\(有\)/g, '')
             .replace(/\s+/g, '').toLowerCase().trim();
}

function _longestCommonSubstr(s1, s2) {
  var max = 0;
  for (var i = 0; i < s1.length; i++) {
    for (var j = 0; j < s2.length; j++) {
      var len = 0;
      while (i+len < s1.length && j+len < s2.length && s1[i+len] === s2[j+len]) len++;
      if (len > max) max = len;
    }
  }
  return max;
}

function _parseDate(val) {
  if (!val || val === '') return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  var s = String(val).trim();
  if (!s) return null;
  try {
    var d = new Date(s.replace(/\//g, '-'));
    return isNaN(d.getTime()) ? null : d;
  } catch(e) { return null; }
}

// ============================================================
// シート操作
// ============================================================

function _getMgmtRowById(sheet, mgmtId) {
  var last = sheet.getLastRow();
  if (last <= 1) return null;
  var data = sheet.getRange(2, 1, last - 1, 27).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][MGMT_COLS.ID - 1]) === String(mgmtId)) return data[i];
  }
  return null;
}

function _getLinesById(sheet, mgmtId, colCount) {
  var last = sheet.getLastRow();
  if (last <= 1) return [];
  return sheet.getRange(2, 1, last - 1, colCount).getValues()
    .filter(function(r) { return String(r[0]) === String(mgmtId); });
}

function _executeLink(mgmtSheet, orderMgmtId, quoteMgmtId, score, method) {
  var last = mgmtSheet.getLastRow();
  if (last <= 1) return;
  var ids = mgmtSheet.getRange(2, MGMT_COLS.ID, last - 1, 1).getValues().flat()
    .map(function(v) { return String(v); });

  var orderIdx = ids.indexOf(String(orderMgmtId));
  var quoteIdx = ids.indexOf(String(quoteMgmtId));
  if (orderIdx < 0 || quoteIdx < 0) {
    Logger.log('[LINK ERROR] 行未発見 order=' + orderIdx + ' quote=' + quoteIdx);
    return;
  }

  var orderRow = orderIdx + 2;
  var quoteRow = quoteIdx + 2;

  var quoteNo  = mgmtSheet.getRange(quoteRow, MGMT_COLS.QUOTE_NO).getValue();
  var quoteUrl = mgmtSheet.getRange(quoteRow, MGMT_COLS.QUOTE_PDF_URL).getValue();
  var orderNo  = mgmtSheet.getRange(orderRow, MGMT_COLS.ORDER_NO).getValue();
  var orderUrl = mgmtSheet.getRange(orderRow, MGMT_COLS.ORDER_PDF_URL).getValue();
  var orderAmt = mgmtSheet.getRange(orderRow, MGMT_COLS.ORDER_AMOUNT).getValue();
  var orderDt  = mgmtSheet.getRange(orderRow, MGMT_COLS.ORDER_DATE).getValue();

  mgmtSheet.getRange(orderRow, MGMT_COLS.QUOTE_NO).setValue(quoteNo);
  mgmtSheet.getRange(orderRow, MGMT_COLS.QUOTE_PDF_URL).setValue(quoteUrl);
  mgmtSheet.getRange(orderRow, MGMT_COLS.LINKED).setValue('TRUE');
  mgmtSheet.getRange(orderRow, MGMT_COLS.STATUS).setValue(CONFIG.STATUS.RECEIVED);
  mgmtSheet.getRange(orderRow, MGMT_COLS.UPDATED_AT).setValue(nowJST());

  mgmtSheet.getRange(quoteRow, MGMT_COLS.ORDER_NO).setValue(orderNo);
  mgmtSheet.getRange(quoteRow, MGMT_COLS.ORDER_PDF_URL).setValue(orderUrl);
  mgmtSheet.getRange(quoteRow, MGMT_COLS.ORDER_AMOUNT).setValue(orderAmt);
  mgmtSheet.getRange(quoteRow, MGMT_COLS.ORDER_DATE).setValue(orderDt);
  mgmtSheet.getRange(quoteRow, MGMT_COLS.LINKED).setValue('TRUE');
  mgmtSheet.getRange(quoteRow, MGMT_COLS.STATUS).setValue(CONFIG.STATUS.RECEIVED);
  mgmtSheet.getRange(quoteRow, MGMT_COLS.UPDATED_AT).setValue(nowJST());

  Logger.log('[LINK] ' + method + ': ' + orderMgmtId + ' ↔ ' + quoteMgmtId + ' (' + score + '点)');
}

function _saveCandidates(orderMgmtId, orderMgmt, candidates) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(MATCH_CONFIG.SHEET_CANDIDATES);
  if (!sheet) {
    sheet = ss.insertSheet(MATCH_CONFIG.SHEET_CANDIDATES);
    var h = ['注文管理ID','注文番号','顧客名','注文日','注文金額',
             '候補1_管理ID','候補1_見積番号','候補1_顧客名','候補1_スコア','候補1_内訳',
             '候補2_管理ID','候補2_スコア','候補3_管理ID','候補3_スコア',
             '確定見積管理ID','処理結果','処理日時'];
    var hr = sheet.getRange(1, 1, 1, h.length);
    hr.setValues([h]);
    hr.setBackground('#E8D5F5');
    hr.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  var last = sheet.getLastRow();
  if (last > 1) {
    var existing = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
    var idx = existing.indexOf(String(orderMgmtId));
    if (idx >= 0) sheet.deleteRow(idx + 2);
  }

  var c1 = candidates[0] || {};
  var c2 = candidates[1] || {};
  var c3 = candidates[2] || {};
  var bd = c1.breakdown || {};
  var bd1 = c1.score
    ? '品名:' + (bd.keyword||0) + ' 数量:' + (bd.qty||0) + ' 金額:' + (bd.amount||0) +
      ' 日付:' + (bd.date||0) + ' 顧客:' + (bd.client||0) + ' 機種:' + (bd.model||0)
    : '';

  sheet.appendRow([
    orderMgmtId,
    String(orderMgmt[MGMT_COLS.ORDER_NO - 1]    || ''),
    String(orderMgmt[MGMT_COLS.CLIENT - 1]       || ''),
    _toDateStr(orderMgmt[MGMT_COLS.ORDER_DATE - 1]),
    Number(orderMgmt[MGMT_COLS.ORDER_AMOUNT - 1]) || '',
    c1.quoteMgmtId||'', c1.quoteNo||'', c1.client||'', c1.score||'', bd1,
    c2.quoteMgmtId||'', c2.score||'',
    c3.quoteMgmtId||'', c3.score||'',
    '', candidates.length > 0 ? '候補あり' : '候補なし', nowJST(),
  ]);
}

// ============================================================
// 手動確定
// ============================================================

function confirmManualLink(orderMgmtId, quoteMgmtId) {
  var ss        = getSpreadsheet();
  var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  _executeLink(mgmtSheet, orderMgmtId, quoteMgmtId, 0, '手動');

  var cand = ss.getSheetByName(MATCH_CONFIG.SHEET_CANDIDATES);
  if (cand && cand.getLastRow() > 1) {
    var ids = cand.getRange(2, 1, cand.getLastRow() - 1, 1).getValues().flat();
    var idx = ids.indexOf(String(orderMgmtId));
    if (idx >= 0) {
      cand.getRange(idx + 2, 15).setValue(quoteMgmtId);
      cand.getRange(idx + 2, 16).setValue('手動確定済み');
      cand.getRange(idx + 2, 17).setValue(nowJST());
    }
  }
  return { success: true };
}

// ============================================================
// 定期実行
// ============================================================

function autoMatchNewOrders() {
  batchMatchAllUnlinked();
}

// ============================================================
// デバッグ用
// ============================================================

function debugMatching() {
  var all      = getAllMgmtData();
  var orders   = all.filter(function(r) { return String(r[MGMT_COLS.ORDER_NO - 1]) !== ''; });
  var unlinked = orders.filter(function(r) { return !_isLinked(r[MGMT_COLS.LINKED - 1]); });
  var quotes   = all.filter(function(r) { return String(r[MGMT_COLS.QUOTE_NO - 1]) !== ''; });

  Logger.log('全件数: ' + all.length);
  Logger.log('注文書件数: ' + orders.length);
  Logger.log('未紐づけ注文書: ' + unlinked.length);
  Logger.log('見積書件数: ' + quotes.length);

  if (unlinked.length === 0) { Logger.log('未紐づけ案件なし'); return; }

  var s = unlinked[0];
  Logger.log('サンプル管理ID: '     + s[MGMT_COLS.ID - 1]);
  Logger.log('サンプル注文番号: '   + s[MGMT_COLS.ORDER_NO - 1]);
  Logger.log('サンプル顧客名: '     + s[MGMT_COLS.CLIENT - 1]);
  Logger.log('サンプルLINKED: '    + s[MGMT_COLS.LINKED - 1]);
  Logger.log('サンプルORDER_DATE: ' + s[MGMT_COLS.ORDER_DATE - 1]);

  var ss2        = getSpreadsheet();
  var os         = ss2.getSheetByName(CONFIG.SHEET_ORDERS);
  var qs         = ss2.getSheetByName(CONFIG.SHEET_QUOTES);
  var testId     = String(s[MGMT_COLS.ID - 1]);
  var orderLines = _getLinesById(os, testId, 19);

  Logger.log('注文書シート行数: '   + os.getLastRow());
  Logger.log('見積書シート行数: '   + qs.getLastRow());
  Logger.log('注文書明細行数: '     + orderLines.length);
  Logger.log('明細あり: ' + (orderLines.length > 0 ? 'YES' : 'NO - 管理シート情報のみでスコアリング'));

  Logger.log('=== スコアテスト ===');
  quotes.forEach(function(q) {
    var qId    = String(q[MGMT_COLS.ID - 1]);
    var qLines = _getLinesById(qs, qId, 15);
    var score  = _calcMatchScore(s, orderLines, q, qLines);
    var th     = orderLines.length > 0 ? MATCH_CONFIG.AUTO_LINK_THRESHOLD : MATCH_CONFIG.AUTO_LINK_THRESHOLD_NOLINES;
    var mark   = score.total >= th ? '★自動' : (score.total >= 20 ? '△候補' : '');
    Logger.log('vs ' + q[MGMT_COLS.QUOTE_NO-1] + ' (' + qId + '): ' + score.total + '点 ' + mark + ' ' + JSON.stringify(score.breakdown));
  });

  var cand = ss2.getSheetByName('紐づけ候補');
  Logger.log('紐づけ候補シート: ' + (cand ? 'あり 行数:' + cand.getLastRow() : 'なし'));
}

function debugIdMismatch() {
  var ss         = getSpreadsheet();
  var orderSheet = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  var oLast      = orderSheet.getLastRow();
  var orderSheetIds = oLast > 1
    ? orderSheet.getRange(2, 1, oLast - 1, 1).getValues().flat().map(String)
    : [];
  var all           = getAllMgmtData();
  var mgmtOrderIds  = all
    .filter(function(r) { return String(r[MGMT_COLS.ORDER_NO - 1]) !== ''; })
    .map(function(r) { return String(r[MGMT_COLS.ID - 1]); });
  var matched = mgmtOrderIds.filter(function(id) { return orderSheetIds.indexOf(id) >= 0; });
  Logger.log('管理シート注文書: ' + mgmtOrderIds.length + '件');
  Logger.log('注文書シートID数: ' + orderSheetIds.length + '件');
  Logger.log('一致数: ' + matched.length + '件');
  Logger.log('注文書シート先頭5: ' + JSON.stringify(orderSheetIds.slice(0,5)));
  Logger.log('管理シート先頭5: '   + JSON.stringify(mgmtOrderIds.slice(0,5)));
}