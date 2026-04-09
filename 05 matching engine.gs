// ============================================================
// 見積書・注文書管理システム
// ファイル 5/5: 紐づけエンジン（Gemini AI & バッチ処理）
// ============================================================

var MATCH_CONFIG = {
  AUTO_LINK_THRESHOLD: 80,
  CANDIDATE_THRESHOLD: 50,
  SHEET_CANDIDATES: '紐づけ候補',
};

/**
 * 注文書IDを受け取り、最適な見積書を1件紐付けるか候補を返す
 */
function matchOrderToQuote(orderMgmtId) {
  // 実体は 10 order upload and notify.gs の aiLinkOrderToQuote を利用
  const result = aiLinkOrderToQuote(orderMgmtId);
  
  if (result.success) {
    if (result.status === 'auto_linked') {
      return { 
        success: true, 
        status: 'auto_linked', 
        quoteMgmtId: result.bestMatch.quoteMgmtId,
        quoteNo: result.bestMatch.quoteNo,
        score: result.bestMatch.score,
        reason: result.bestMatch.reason
      };
    } else if (result.status === 'candidates_found') {
      return { 
        success: true, 
        status: 'candidates_found', 
        candidates: result.candidates.map(c => ({
          quoteMgmtId: c.quoteMgmtId,
          quoteNo: c.quoteNo,
          score: c.score,
          reason: c.reason
        }))
      };
    }
  }
  
  return { success: true, status: 'no_match', candidates: [] };
}

/**
 * 紐づけられていないすべての注文書に対して一括マッチングを実行
 */
function runBatchMatching() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mgmtSheet = ss.getSheetByName('案件管理');
  const data = mgmtSheet.getDataRange().getValues();
  const headers = data[0];
  
  const idIdx = headers.indexOf('管理ID');
  const qNoIdx = headers.indexOf('見積番号');
  const linkedIdx = headers.indexOf('紐付け済み');
  
  let autoCount = 0;
  let candidateCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const mgmtId = String(row[idIdx]);
    const linked = row[linkedIdx] === 'TRUE' || row[linkedIdx] === true;
    
    if (mgmtId.startsWith('MO') && !row[qNoIdx] && !linked) {
      const result = aiLinkOrderToQuote(mgmtId);
      if (result.success) {
        if (result.status === 'auto_linked') autoCount++;
        else if (result.status === 'candidates_found') candidateCount++;
      }
    }
  }
  
  return { 
    success: true, 
    message: `一括マッチング完了: 自動紐付け ${autoCount}件, 候補発見 ${candidateCount}件` 
  };
}

/**
 * ダッシュボードUI用の紐付け候補リストを取得
 */
function getMatchingCandidates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mgmtSheet = ss.getSheetByName('案件管理');
  const data = mgmtSheet.getDataRange().getValues();
  const headers = data[0];
  
  const idIdx = headers.indexOf('管理ID');
  const qNoIdx = headers.indexOf('見積番号');
  const clientIdx = headers.indexOf('顧客名');
  const subjIdx = headers.indexOf('件名');
  const amountIdx = headers.indexOf('注文金額');
  const linkedIdx = headers.indexOf('紐付け済み');

  const items = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const mgmtId = String(row[idIdx]);
    const linked = row[linkedIdx] === 'TRUE' || row[linkedIdx] === true;
    
    if (mgmtId.startsWith('MO') && !row[qNoIdx] && !linked) {
      // リアルタイムで判定結果を取得
      const result = aiLinkOrderToQuote(mgmtId);
      if (result.success && result.candidates && result.candidates.length > 0) {
        items.push({
          orderMgmtId: mgmtId,
          orderNo: row[subjIdx] || '名称不明',
          client: row[clientIdx],
          orderAmount: row[amountIdx],
          candidates: result.candidates
        });
      }
    }
  }
  
  return { success: true, items: items };
}

/**
 * 手動で紐付けを確定する
 */
function confirmManualLink(orderMgmtId, quoteMgmtId) {
  try {
    // 10 order upload and notify.gs 内の _applyOrderLink を利用
    _applyOrderLink(orderMgmtId, quoteMgmtId);
    return { success: true, message: '紐付けを確定しました' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * 定期実行用の関数
 */
function autoMatchNewOrders() {
  runBatchMatching();
}