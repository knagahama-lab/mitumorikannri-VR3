// ============================================================
// 見積書・注文書管理システム
// ファイル 0: 不足ユーティリティ関数群 ★白画面修正
// ============================================================
// このファイルは他のファイルから参照されているが未定義だった
// 関数をすべて補完します。GASプロジェクトに追加してください。
// ============================================================

// ============================================================
// 文字列正規化（検索用）
// ============================================================

function normalizeText(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .toLowerCase()
    // 全角英数字→半角
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    })
    // 全角スペース→半角
    .replace(/[\u3000]/g, ' ')
    .trim();
}

// ============================================================
// 設定オブジェクト取得
// ============================================================

function _loadSettingsObj() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('SYS_SETTINGS');
    if (raw) return JSON.parse(raw);
  } catch(e) {}
  return {
    webhookUrl:  '',
    notifyOrder: true,
    notifyQuote: true,
    notifyDl:    true,
    alertDays:   3,
  };
}

// ============================================================
// AI紐づけ候補の取得（プロパティ＋シート両対応）
// ============================================================

function getMatchingCandidates() {
  try {
    // ① スクリプトプロパティから取得（05_matching_engine.gs の runBatchMatching が保存する場所）
    var raw = PropertiesService.getScriptProperties().getProperty('AI_MATCHING_CANDIDATES');
    var stored = raw ? JSON.parse(raw) : [];

    // ② 管理シートから「未紐づけ注文書」を取得
    var allMgmt = getAllMgmtData();
    var unlinkedOrders = allMgmt.filter(function(r) {
      return String(r[MGMT_COLS.ORDER_NO - 1]).trim() !== '' &&
             !_isLinkedVal(r[MGMT_COLS.LINKED - 1]);
    });

    // ③ 見積書データ
    var ss = getSpreadsheet();
    var qs = ss.getSheetByName(CONFIG.SHEET_QUOTES);
    var quoteLineMap = {};
    if (qs && qs.getLastRow() > 1) {
      qs.getRange(2, 1, qs.getLastRow() - 1, 15).getValues()
        .filter(function(r) { return r[0] && r[6]; })
        .forEach(function(r) {
          var mid = String(r[0]);
          if (!quoteLineMap[mid]) quoteLineMap[mid] = [];
          quoteLineMap[mid].push({
            itemName:  String(r[6] || ''),
            spec:      String(r[7] || ''),
            unitPrice: r[10],
            pdfUrl:    String(r[13] || ''),
          });
        });
    }

    // ④ 未紐づけ見積書の一覧（フォールバック用）
    var quotesMgmt = getAllMgmtData().filter(function(r) {
      return String(r[MGMT_COLS.QUOTE_NO - 1]).trim() !== '' &&
             !_isLinkedVal(r[MGMT_COLS.LINKED - 1]);
    });
    var unlinkedQuoteObjs = quotesMgmt.map(function(r) {
      var mid = String(r[MGMT_COLS.ID - 1]);
      var lines = quoteLineMap[mid] || [];
      return {
        quoteId:    mid,
        quoteNo:    String(r[MGMT_COLS.QUOTE_NO - 1]      || ''),
        client:     String(r[MGMT_COLS.CLIENT - 1]         || ''),
        issueDate:  _toDateStr(r[MGMT_COLS.QUOTE_DATE - 1]),
        amount:     _toNum(r[MGMT_COLS.QUOTE_AMOUNT - 1]),
        subject:    String(r[MGMT_COLS.SUBJECT - 1]        || ''),
        quoteUrl:   String(r[MGMT_COLS.QUOTE_PDF_URL - 1]  || ''),
        items:      lines,
      };
    });

    // ⑤ stored に含まれていない未紐づけ注文書を追加（0件候補でも表示するため）
    var storedOrderIds = {};
    stored.forEach(function(s) { storedOrderIds[s.orderMgmtId] = true; });

    unlinkedOrders.forEach(function(r) {
      var mid = String(r[MGMT_COLS.ID - 1]);
      if (storedOrderIds[mid]) return;

      // 顧客名で絞り込んだフォールバック候補
      var orderClient = String(r[MGMT_COLS.CLIENT - 1] || '');
      var fallback = unlinkedQuoteObjs.filter(function(q) {
        return !orderClient || q.client.indexOf(orderClient) >= 0 || orderClient.indexOf(q.client) >= 0;
      });
      if (fallback.length === 0) fallback = unlinkedQuoteObjs.slice(0, 3);

      var candidates = fallback.slice(0, 3).map(function(q) {
        return {
          quoteId:   q.quoteId,
          quoteNo:   q.quoteNo,
          client:    q.client,
          issueDate: q.issueDate,
          amount:    q.amount,
          subject:   q.subject,
          quoteUrl:  q.quoteUrl,
          items:     q.items,
          score:     0,
          reason:    '類似候補が見つからなかったため、未紐づけ見積書を参考表示しています。',
        };
      });

      stored.push({
        orderMgmtId:  mid,
        orderNo:      String(r[MGMT_COLS.ORDER_NO - 1]   || ''),
        orderClient:  orderClient,
        orderDate:    _toDateStr(r[MGMT_COLS.ORDER_DATE - 1]),
        orderAmount:  _toNum(r[MGMT_COLS.ORDER_AMOUNT - 1]),
        orderPdfUrl:  String(r[MGMT_COLS.ORDER_PDF_URL - 1] || ''),
        candidates:   candidates,
      });
    });

    // ⑥ candidates が0件のエントリにも未紐づけ見積書を補填
    stored = stored.map(function(item) {
      if (item.candidates && item.candidates.length > 0) return item;

      var orderClient = item.orderClient || '';
      var fallback = unlinkedQuoteObjs.filter(function(q) {
        return !orderClient || q.client.indexOf(orderClient) >= 0 || orderClient.indexOf(q.client) >= 0;
      });
      if (fallback.length === 0) fallback = unlinkedQuoteObjs.slice(0, 3);

      item.candidates = fallback.slice(0, 3).map(function(q) {
        return {
          quoteId:   q.quoteId,
          quoteNo:   q.quoteNo,
          client:    q.client,
          issueDate: q.issueDate,
          amount:    q.amount,
          subject:   q.subject,
          quoteUrl:  q.quoteUrl,
          items:     q.items,
          score:     0,
          reason:    '類似候補が見つからなかったため、未紐づけ見積書を参考表示しています。',
        };
      });
      return item;
    });

    return { success: true, items: stored };
  } catch(e) {
    Logger.log('[getMatchingCandidates ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// 手動紐づけ確定
// ============================================================

function confirmManualLink(orderMgmtId, quoteMgmtId) {
  try {
    if (!orderMgmtId || !quoteMgmtId) {
      return { success: false, error: '管理IDが不足しています' };
    }
    var ss        = getSpreadsheet();
    var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var last      = mgmtSheet.getLastRow();
    if (last <= 1) return { success: false, error: 'データなし' };

    var allData  = mgmtSheet.getRange(2, 1, last - 1, 27).getValues();
    var ids      = allData.map(function(r) { return String(r[MGMT_COLS.ID - 1]); });

    // 注文書行を特定
    var oIdx = ids.indexOf(String(orderMgmtId));
    if (oIdx < 0) return { success: false, error: '注文書管理IDが見つかりません: ' + orderMgmtId };
    var oRow = oIdx + 2;

    // 見積書行を特定
    var qIdx = ids.indexOf(String(quoteMgmtId));
    if (qIdx < 0) return { success: false, error: '見積書管理IDが見つかりません: ' + quoteMgmtId };
    var qRow = qIdx + 2;

    var quoteNo    = String(allData[qIdx][MGMT_COLS.QUOTE_NO - 1]     || '');
    var quotePdfUrl= String(allData[qIdx][MGMT_COLS.QUOTE_PDF_URL - 1]|| '');
    var orderNo    = String(allData[oIdx][MGMT_COLS.ORDER_NO - 1]     || '');
    var orderPdfUrl= String(allData[oIdx][MGMT_COLS.ORDER_PDF_URL - 1]|| '');

    // 注文書側を更新
    mgmtSheet.getRange(oRow, MGMT_COLS.QUOTE_NO).setValue(quoteNo);
    mgmtSheet.getRange(oRow, MGMT_COLS.QUOTE_PDF_URL).setValue(quotePdfUrl);
    mgmtSheet.getRange(oRow, MGMT_COLS.LINKED).setValue('TRUE');
    mgmtSheet.getRange(oRow, MGMT_COLS.STATUS).setValue(CONFIG.STATUS.RECEIVED);
    mgmtSheet.getRange(oRow, MGMT_COLS.UPDATED_AT).setValue(nowJST());

    // 見積書側を更新
    mgmtSheet.getRange(qRow, MGMT_COLS.ORDER_NO).setValue(orderNo);
    mgmtSheet.getRange(qRow, MGMT_COLS.ORDER_PDF_URL).setValue(orderPdfUrl);
    mgmtSheet.getRange(qRow, MGMT_COLS.LINKED).setValue('TRUE');
    mgmtSheet.getRange(qRow, MGMT_COLS.STATUS).setValue(CONFIG.STATUS.RECEIVED);
    mgmtSheet.getRange(qRow, MGMT_COLS.UPDATED_AT).setValue(nowJST());

    // 候補リストからこのエントリを削除
    try {
      var raw = PropertiesService.getScriptProperties().getProperty('AI_MATCHING_CANDIDATES');
      if (raw) {
        var stored = JSON.parse(raw);
        stored = stored.filter(function(item) {
          return String(item.orderMgmtId) !== String(orderMgmtId);
        });
        PropertiesService.getScriptProperties()
          .setProperty('AI_MATCHING_CANDIDATES', JSON.stringify(stored));
      }
    } catch(propErr) {
      Logger.log('[CONFIRM LINK PROP ERROR] ' + propErr.message);
    }

    return { success: true, quoteNo: quoteNo, orderNo: orderNo };
  } catch(e) {
    Logger.log('[confirmManualLink ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// Drive監視用：最新の注文書mgmtIdを取得
// ============================================================

function _findLatestOrderMgmtId() {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var last  = sheet.getLastRow();
    if (last <= 1) return null;
    // 最後の行（最新登録）を返す
    var row = sheet.getRange(last, MGMT_COLS.ID).getValue();
    return row ? String(row) : null;
  } catch(e) {
    Logger.log('[_findLatestOrderMgmtId ERROR] ' + e.message);
    return null;
  }
}

// ============================================================
// Drive監視用：注文書登録をChatに通知
// ============================================================

function _sendOrderRegistrationToChat(mgmtId, uploadInfo, linkResult) {
  try {
    var webhookUrl = _getChatWebhookUrl();
    if (!webhookUrl) return;

    var lr  = linkResult || {};
    var lines = [
      '【📦 注文書を登録（Drive自動取込）】',
      '発注書番号: ' + (uploadInfo.documentNo || '—'),
      '種別: '      + (uploadInfo.orderType   || '—'),
    ];
    if (lr.status === 'auto_linked') {
      lines.push('✅ AI自動紐づけ完了');
    } else if (lr.status === 'candidates_found') {
      lines.push('⚠️ 紐づけ候補あり（要確認）');
    } else {
      lines.push('❌ 紐づく見積書が見つかりませんでした');
    }
    var appUrl = ScriptApp.getService().getUrl();
    lines.push('▶ ' + appUrl);

    UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: lines.join('\n') }),
      muteHttpExceptions: true,
    });
  } catch(e) {
    Logger.log('[_sendOrderRegistrationToChat ERROR] ' + e.message);
  }
}