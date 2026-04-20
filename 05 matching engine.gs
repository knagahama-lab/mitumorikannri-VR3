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
// ============================================================
// 🤖 AIマッチング司令塔 & Gemini API 連携
// ============================================================
function runBatchMatching() {
  try {
    var allOrders = getAllMgmtData().map(_rowToObject);
    var unlinkedOrders = allOrders.filter(function(o) { return o.orderNo && !o.quoteNo; });

    var ss = getSpreadsheet();
    
    // ★ エラー修正：見積書のシートをあらゆる名前のパターンで探し出して確実に取得する
    var qsName = (typeof CONFIG !== 'undefined' && CONFIG.SHEET_QUOTES) ? CONFIG.SHEET_QUOTES : null;
    var qs = qsName ? ss.getSheetByName(qsName) : null;
    
    // CONFIGの指定でダメなら一般的なシート名でフォールバック検索
    if (!qs) qs = ss.getSheetByName('見積管理');
    if (!qs) qs = ss.getSheetByName('見積一覧');
    if (!qs) qs = ss.getSheetByName('見積書');
    if (!qs) qs = ss.getSheetByName('見積');
    if (!qs) qs = ss.getSheetByName('見積データ');

    // それでも見つからない場合はエラーメッセージを返して安全に停止
    if (!qs) {
      throw new Error("見積書のシートが見つかりません。スプレッドシートのシート名を確認してください。");
    }

    // シートが空っぽ（ヘッダーしかない）場合はゼロ件で終了
    if (qs.getLastRow() <= 1) {
      return { autoCount: 0, candidateCount: 0 };
    }

    var qData = qs.getDataRange().getValues();
    var qHeaders = qData[0];
    
    var unlinkedQuotes = [];
    for (var i = 1; i < qData.length; i++) {
      var row = qData[i];
      var obj = {};
      qHeaders.forEach(function(h, idx) { obj[h] = row[idx]; });
      
      // 発注書番号が空欄＝未紐づけの見積書として抽出
      // 環境により列位置が違うため、オブジェクトとインデックスの両方で安全に判定
      var isLinked = row[15] === true || row[15] === 'TRUE' || obj['紐づけ済'] || obj['発注書番号'];
      
      if (!isLinked) {
         unlinkedQuotes.push({
           quoteId: obj['id'] || obj['管理ID'] || row[0], 
           quoteNo: obj['見積書番号'] || obj['見積番号'] || row[1],
           client: obj['宛先会社名'] || obj['顧客名'] || obj['提出先'] || row[3],
           amount: obj['見積金額'] || obj['合計金額'] || row[11],
           modelCode: obj['機種コード'] || row[5],
           date: obj['発行日'] || row[2]
         });
      }
    }

    var candidatesStore = [];
    var autoCount = 0;
    var candidateCount = 0;

    for (var j = 0; j < unlinkedOrders.length; j++) {
      var order = unlinkedOrders[j];
      var orderForAI = {
        orderNo: order.orderNo,
        client: order.client,
        amount: order.orderAmount,
        modelCode: order.modelCode,
        text: order.memo
      };

      // AI関数を呼び出し（別ファイルにある matchWithGeminiAPI を実行）
      var aiResults = matchWithGeminiAPI(orderForAI, unlinkedQuotes);
      if (!aiResults || aiResults.length === 0) continue;

      aiResults.sort(function(a, b) { return b.score - a.score; });
      var bestMatch = aiResults[0];

      if (bestMatch.score >= 80) {
        if (typeof _apiConfirmOrderLink === 'function') {
          _apiConfirmOrderLink({ orderMgmtId: order.id, quoteMgmtId: bestMatch.quoteId });
        }
        autoCount++;
      } else if (bestMatch.score >= 50) {
        candidatesStore.push({
          orderMgmtId: order.id,
          orderNo: order.orderNo,
          orderClient: order.client,
          orderDate: order.orderDate,
          orderAmount: order.orderAmount,
          orderPdfUrl: order.orderPdfUrl,
          candidates: aiResults.filter(function(c) { return c.score >= 50; })
        });
        candidateCount++;
      }
    }

    PropertiesService.getScriptProperties().setProperty('AI_MATCHING_CANDIDATES', JSON.stringify(candidatesStore));
    return { autoCount: autoCount, candidateCount: candidateCount };

  } catch (e) {
    Logger.log("runBatchMatching エラー: " + e.message);
    throw e;
  }
}
// ============================================================
// 🤖 自動マッチング実行用トリガー関数 (定期実行用)
// ============================================================
function autoMatchNewOrders() {
  try {
    Logger.log("定期実行: AI自動マッチングを開始します...");
    
    // 最新のマッチング司令塔（runBatchMatching）を呼び出す
    var result = runBatchMatching(); 
    
    Logger.log("定期実行完了: 自動確定 " + result.autoCount + "件, 候補抽出 " + result.candidateCount + "件");
  } catch (e) {
    Logger.log("定期実行エラー: " + e.message);
  }
}