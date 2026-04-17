// ============================================================
// 見積書・注文書管理システム
// ファイル 3/4: Webアプリ doGet / API ルーター
// ============================================================

// ============================================================
// ★ 共通ユーティリティ: エラーハンドリング・監査ログ・メール通知
// ============================================================

/**
 * 統一エラーレスポンス生成（内部エラー詳細を外部に漏らさない）
 */
function _errorResponse(e, context) {
  Logger.log('[ERROR][' + (context || 'unknown') + '] ' + e.message + '\n' + e.stack);
  return { success: false, error: (context || 'エラー') + 'が発生しました。時間をおいて再試行してください。' };
}

/**
 * 監査ログをスプレッドシートの「監査ログ」シートに書き込む
 * @param {string} action  - 操作種別 (例: 'quote_import', 'api_key_update')
 * @param {string} detail  - 詳細情報
 * @param {string} result  - 'success' or 'error'
 */
function _writeAuditLog(action, detail, result) {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName('監査ログ');
    if (!sheet) {
      sheet = ss.insertSheet('監査ログ');
      sheet.appendRow(['日時', '操作者', '操作種別', '詳細', '結果']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#e8f0fe');
    }
    var user = Session.getActiveUser().getEmail() || 'system';
    sheet.appendRow([
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
      user, action, String(detail).substring(0, 300), result || 'success'
    ]);
  } catch(e) {
    Logger.log('[AUDIT LOG ERROR] ' + e.message);
  }
}

/**
 * Gmail メール通知送信（NOTIFY_EMAILS に設定されたアドレスへ）
 * @param {string} subject - 件名
 * @param {string} body    - 本文（プレーンテキスト）
 * @param {string} htmlBody - HTML形式の本文（任意）
 */
function _sendNotifyEmail(subject, body, htmlBody) {
  try {
    var props    = PropertiesService.getScriptProperties();
    var toStr    = props.getProperty('NOTIFY_EMAILS') || '';
    var settings = _loadSettingsObj();
    // 設定でメール通知が有効かつ送信先が設定されている場合のみ送信
    if (!toStr || toStr.trim() === '') return;
    var toEmails = toStr.split(',').map(function(s){ return s.trim(); }).filter(Boolean);
    if (toEmails.length === 0) return;
    var options = { name: '見積・注文管理システム', noReply: false };
    if (htmlBody) options.htmlBody = htmlBody;
    toEmails.forEach(function(email) {
      try {
        GmailApp.sendEmail(email, subject, body, options);
      } catch(e2) {
        Logger.log('[GMAIL SEND ERROR] to:' + email + ' / ' + e2.message);
      }
    });
    Logger.log('[NOTIFY EMAIL SENT] to:' + toEmails.join(',') + ' / ' + subject);
  } catch(e) {
    Logger.log('[NOTIFY EMAIL ERROR] ' + e.message);
  }
}

/**
 * HTMLメール本文テンプレートを生成
 */
function _buildEmailHtml(title, rows, appUrl) {
  var rowsHtml = rows.map(function(r) {
    return '<tr><td style="padding:6px 12px;color:#555;font-size:13px;border-bottom:1px solid #eee;">' + r[0] +
           '</td><td style="padding:6px 12px;font-size:13px;font-weight:600;border-bottom:1px solid #eee;">' + (r[1] || '—') + '</td></tr>';
  }).join('');
  return '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f5f7fa;font-family:-apple-system,sans-serif;">' +
    '<div style="max-width:560px;margin:30px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.08);">' +
    '<div style="background:linear-gradient(135deg,#1e40af,#3b82f6);padding:24px 28px;">' +
    '<h2 style="color:#fff;margin:0;font-size:16px;font-weight:700;">📋 ' + title + '</h2>' +
    '<p style="color:rgba(255,255,255,0.8);margin:6px 0 0;font-size:12px;">見積・注文管理システム / 自動通知</p>' +
    '</div>' +
    '<table style="width:100%;border-collapse:collapse;margin:0;">' + rowsHtml + '</table>' +
    '<div style="padding:20px 28px;border-top:1px solid #e5e7eb;">' +
    '<a href="' + (appUrl || '#') + '" style="display:inline-block;background:#1e40af;color:#fff;padding:10px 22px;border-radius:8px;text-decoration:none;font-size:13px;font-weight:600;">🔗 システムで確認する</a>' +
    '</div></div></body></html>';
}

function doGet(e) {
  var userEmail = Session.getActiveUser().getEmail() || '';
  var adminEmailsStr = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAILS') || '';
  var isAdmin = false;
  if (!adminEmailsStr || adminEmailsStr.trim() === '') {
    isAdmin = true; // 未設定時は全員を管理者扱いとする
  } else {
    var adminEmails = adminEmailsStr.split(',').map(function(s){return s.trim();});
    isAdmin = (adminEmails.indexOf(userEmail) >= 0);
  }

  // URLの末尾に「?page=bom」がついている場合はBOM管理画面を開く
  if (e && e.parameter && e.parameter.page === 'bom') {
    var bomTmpl = HtmlService.createTemplateFromFile('BomDashboard');
    bomTmpl.isAdmin = isAdmin;
    return bomTmpl.evaluate()
      .setTitle('BOM・部品管理システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (e && e.parameter && e.parameter.action === 'viewBom') {
    var fileId = e.parameter.fileId;
    if (!fileId) return HtmlService.createHtmlOutput('Error: No file ID provided');
    try {
      var file = DriveApp.getFileById(fileId);
      var content = file.getBlob().getDataAsString('utf-8');
      return HtmlService.createHtmlOutput(content);
    } catch(err) {
      return HtmlService.createHtmlOutput('Error: ' + err.message);
    }
  }

  // 通常のアクセス（パラメータなし）の場合は、これまでの見積・注文管理システムを開く
  var dashboardTmpl = HtmlService.createTemplateFromFile('Dashboard');
  dashboardTmpl.isAdmin = isAdmin;
  dashboardTmpl.userEmail = userEmail;

  return dashboardTmpl.evaluate()
    .setTitle('見積・注文 管理システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// HTMLテンプレートから別ファイルをインクルードするためのヘルパー関数
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// ★ 安全な通信ルーター（フリーズ完全回避版）
// ============================================================
function handleApiRequest(action, payload) {
  try {
    payload = payload || {};
    var res;
    
    switch (action) {
      case 'getAll':              res = _apiGetAll(); break;
      case 'search':              res = _apiSearch(payload); break;
      case 'uploadPdf':           res = _apiUploadPdf(payload); break;
      case 'updateStatus':        res = _apiUpdateStatus(payload); break;
      case 'getDetail':           res = _apiGetDetail(payload); break;
      case 'getTodos':            res = _apiGetTodos(); break;
      case 'saveTodo':            res = _apiSaveTodo(payload); break;
      case 'deleteTodo':          res = _apiDeleteTodo(payload); break;
      case 'getCalendar':         res = _apiGetCalendar(payload); break;
      case 'getCandidates':       res = _apiGetCandidates(); break;
      case 'confirmLink':         res = _apiConfirmLink(payload); break;
      case 'runMatching':         res = _apiRunMatching(payload); break;
      case 'quoteListGetAll':     res = _apiQuoteListGetAll(); break;
      case 'getQuoteDetail':      res = _apiGetQuoteDetail(payload); break;
      case 'ledgerGetAll':        res = _apiLedgerGetAll(); break;
      case 'ledgerSave':          res = _apiLedgerSave(payload); break;
      case 'ledgerDelete':        res = _apiLedgerDelete(payload); break;
      case 'ledgerUpdateUrl':     res = _apiLedgerUpdateUrl(payload); break;
      case 'ledgerUploadFile':    res = _apiLedgerUploadFile(payload); break;
      case 'ledgerGetMachines':   res = _apiLedgerGetMachines(); break;
      case 'ledgerCreateMachine': res = _apiLedgerCreateMachine(payload); break;
      case 'appendBoardRow':      res = _apiAppendBoardRow(payload); break;
      case 'updateMgmt':          res = _apiUpdateMgmt(payload); break;
      case 'deleteMgmt':          res = _apiDeleteMgmt(payload); break;
      case 'chatbotQuery':        res = apiChatbotQuery(payload); break;
      case 'modelInfoGet':        res = _apiModelInfoGet(payload); break;
      case 'modelInfoSave':       res = _apiModelInfoSave(payload); break;
      case 'modelInfoUpload':     res = _apiModelInfoUpload(payload); break;
      case 'driveSearch':         res = _apiDriveSearch(payload); break;
      case 'driveRefreshCache':   res = { success: true, count: refreshDrivePdfCache() }; break;
      case 'driveDeleteFile':     res = _apiDriveDeleteFile(payload); break;
      case 'qsGetAll':            res = _apiQsGetAll(payload); break;
      case 'qsSave':              res = _apiQsSave(payload); break;
      case 'qsDelete':            res = _apiQsDelete(payload); break;
      case 'qsUploadFile':        res = _apiQsUploadFile(payload); break;
      case 'qsGetMachines':       res = _apiQsGetMachines(); break;
      case 'partsGetAll':         res = _apiPartsGetAll(payload); break;
      case 'partsSave':           res = _apiPartsSave(payload); break;
      case 'partsDelete':         res = _apiPartsDelete(payload); break;
      case 'partsImportCSV':      res = _apiPartsImportCSV(payload); break;
      case 'partsExportCSV':      res = _apiPartsExportCSV(payload); break;
      case 'pcbGetAll':           res = _apiPcbGetAll(payload); break;
      case 'pcbSave':             res = _apiPcbSave(payload); break;
      case 'pcbDelete':           res = _apiPcbDelete(payload); break;
      case 'pcbImportCSV':        res = _apiPcbImportCSV(payload); break;
      case 'pcbExportCSV':        res = _apiPcbExportCSV(payload); break;
      case 'ensurePartsSheets':   res = ensurePartsSheets() || { success: true }; break;
      case 'boardGetAll':         res = apiBoardGetAll(); break;
      case 'boardGetParts':       res = apiBoardGetParts(); break;
      case 'boardGetMachines':    res = apiBoardGetMachines(); break;
      case 'boardAddNew':         res = _apiBoardAddNew(payload); break;
      case 'boardSavePart':       res = apiBoardSavePart(payload); break;
      case 'boardDeletePart':     res = apiBoardDeletePart(payload.id); break;
      case 'boardSaveMachine':    res = apiBoardSaveMachine(payload); break;
      case 'boardDeleteMachine':  res = apiBoardDeleteMachine(payload.id); break;
      case 'boardSaveBoard':      res = apiBoardSaveBoard(payload); break;
      case 'boardDeleteBoard':    res = apiBoardDeleteBoard(payload.id); break;
      case 'boardSaveBOM':        res = apiBoardSaveBOM(payload.boardId, payload.lines); break;
      case 'boardImportBOMCSV':   res = apiBoardImportBOMCSV(payload.csvText); break;
      case 'machineAddNew':       res = _apiMachineAddNew(payload); break;
      case 'boardGetDetail':      res = apiBoardGetDetail(payload.boardId, payload.boardName); break;
      case 'boardGetAnalysis':    res = apiGetBoardAnalysis(); break;
      case 'boardGetOrders':      res = apiGetOrdersWithBoardInfo(); break;
      case 'boardComparePrice':   res = apiComparePriceToBOM(payload.mgmtId); break;
      case 'searchDetail':        res = _apiSearchDetail(payload); break;
      case 'updateLineStatus':    res = _apiUpdateLineStatus(payload); break;
      case 'createRevision':      res = _apiCreateRevision(payload); break;
      case 'checkDeadlines':      res = _apiCheckDeadlines(); break;
      case 'saveSettings':        res = _apiSaveSettings(payload); break;
      case 'loadSettings':        res = _apiLoadSettings(); break;
      case 'testGemini':          res = _apiTestGeminiConnection(); break;
      case 'testWebhook':         res = _apiTestWebhook(payload); break;
      case 'sendAnnouncement':    res = _apiSendAnnouncement(); break;
      case 'uploadOrderWithLink': res = _apiUploadOrderWithLink(payload); break;
      case 'confirmOrderLink':    res = _apiConfirmOrderLink(payload); break;
      case 'getMatchingCandidates': res = _apiGetMatchingCandidates(payload); break;
      case 'runBatchMatching':    res = _apiRunBatchMatching(payload); break;
      // ===== BOM管理API（スプレッドシート共有） =====
      case 'bomGetAll':           return apiBomGetAll();
      case 'bomSavePart':         return apiBomSavePart(payload);
      case 'bomDeletePart':       return apiBomDeletePart(payload.id);
      case 'bomSaveProduct':      return apiBomSaveProduct(payload);
      case 'bomDeleteProduct':    return apiBomDeleteProduct(payload.id);
      case 'bomSaveBoard':        return apiBomSaveBoard(payload);
      case 'bomDeleteBoard':      return apiBomDeleteBoard(payload.id);
      case 'bomSaveBomRow':       return apiBomSaveBomRow(payload);
      case 'bomDeleteBomRow':     return apiBomDeleteBomRow(payload.id);
      case 'bomImportFinalList':  return apiBomImportFinalList(payload.rows);
      case 'bomImportK10Parts':   return apiBomImportK10Parts(payload.rows);
      case 'reregisterTriggers':
        try { _registerTriggers(); return { success: true }; }
        catch(e) { return { success: false, error: e.message }; }

      // ===== 単価比較・正規化DB API（11_db_normalize.gs / 12_price_compare_v2.gs）=====
      case 'orderPriceCompare':       res = apiOrderPriceCompare(payload); break;
      case 'batchOrderPriceCompare':  res = apiBatchOrderPriceCompare(payload); break;
      case 'searchUnitPrice':
        res = { success: true, results: searchUnitPrice(payload.itemName, payload.spec, payload.client) };
        break;
      case 'searchCaseSummary':
        var _kws = String(payload.keywords || '').split(/[\s,　]+/).filter(Boolean);
        res = { success: true, results: searchCaseSummary(_kws) };
        break;
      case 'syncDetailDB':
        syncAllDetailDB();
        res = { success: true, message: '明細DB同期完了' };
        break;
      case 'rebuildUnitPrice':
        rebuildUnitPriceMaster();
        res = { success: true, message: '単価マスタ再構築完了' };
        break;

      // ===== OCR確認UI API（15_ocr_review_ui.gs）=====
      case 'ocrPreview':  res = apiOcrPreview(payload);  break;
      case 'ocrApprove':  res = apiOcrApprove(payload);  break;
      case 'ocrDiscard':  res = apiOcrDiscard(payload);  break;

      // ===== ★ 新機能 API =====
      case 'getAnalytics':     res = _apiGetAnalytics(payload);    break;
      case 'exportCsv':        res = _apiExportCsv(payload);       break;
      case 'getExportUrl':     res = _apiGetExportUrl(payload);    break;
      case 'aiMonthlySummary': res = _apiAiMonthlySummary(payload); break;
      case 'invalidateCache':  _invalidateMgmtCache(); res = { success: true }; break;
      case 'getAuditLog':      res = _apiGetAuditLog(payload);     break;

      default: return { success: false, error: '不明なアクション: ' + action };
    }
    
    // ★ 魔法の一行：ここでデータを強制的に「通信用」に変換し、フリーズを100%防ぎます
    return JSON.parse(JSON.stringify(res));

  } catch(e) {
    Logger.log('[API ERROR] ' + action + ': ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

// ============================================================
// ★ CacheService キャッシュ（速度改善）
// ============================================================
var MGMT_CACHE_KEY   = 'MGMT_DATA_CACHE';
var MGMT_CACHE_TTL   = 30; // 秒

function _getCachedMgmtData() {
  try {
    var cache  = CacheService.getScriptCache();
    var cached = cache.get(MGMT_CACHE_KEY);
    if (cached) return JSON.parse(cached);
  } catch(e) {}
  return null;
}

function _setCachedMgmtData(data) {
  try {
    var cache = CacheService.getScriptCache();
    var str   = JSON.stringify(data);
    // 100KB制限内に収める
    if (str.length < 90000) {
      cache.put(MGMT_CACHE_KEY, str, MGMT_CACHE_TTL);
    }
  } catch(e) {}
}

function _invalidateMgmtCache() {
  try { CacheService.getScriptCache().remove(MGMT_CACHE_KEY); } catch(e) {}
}

// データ書き込み後に必ずキャッシュを無効化する
function getAllMgmtDataWithCache() {
  var cached = _getCachedMgmtData();
  if (cached) return cached;
  var data = getAllMgmtData();
  _setCachedMgmtData(data);
  return data;
}

// ============================================================
// ★ 分析データAPI（グラフ用）
// ============================================================
function _apiGetAnalytics(p) {
  try {
    var rows    = getAllMgmtData().map(_rowToObject);
    var today   = new Date();
    var nowYM   = today.getFullYear() + '/' + String(today.getMonth()+1).padStart(2,'0');

    // ---- 月別受注金額（過去6ヶ月）----
    var months = [];
    for (var m = 5; m >= 0; m--) {
      var d   = new Date(today.getFullYear(), today.getMonth() - m, 1);
      var ym  = d.getFullYear() + '/' + String(d.getMonth()+1).padStart(2,'0');
      var lbl = String(d.getMonth()+1) + '月';
      var tot = rows.filter(function(r){ return String(r.orderDate||r.quoteDate||'').startsWith(ym); })
                    .reduce(function(s,r){ return s + (Number(r.orderAmount)||0); }, 0);
      months.push({ ym: ym, label: lbl, amount: tot });
    }

    // ---- ステータス別件数 ----
    var byStatus = {};
    rows.forEach(function(r) {
      var st = r.status || '不明';
      byStatus[st] = (byStatus[st]||0) + 1;
    });

    // ---- 顧客別受注金額 TOP5 ----
    var byClient = {};
    rows.forEach(function(r) {
      var cl = r.client || '不明';
      byClient[cl] = (byClient[cl]||0) + (Number(r.orderAmount)||0);
    });
    var topClients = Object.keys(byClient)
      .map(function(k){ return { client: k, amount: byClient[k] }; })
      .sort(function(a,b){ return b.amount - a.amount; })
      .slice(0, 5);

    // ---- 試作 vs 量産 ----
    var trialCount = rows.filter(function(r){ return r.orderType === '試作'; }).length;
    var massCount  = rows.filter(function(r){ return r.orderType === '量産'; }).length;

    // ---- 今月サマリー ----
    var thisMonth = rows.filter(function(r){
      return String(r.orderDate||r.quoteDate||'').startsWith(nowYM);
    });
    var totalThisMonth = thisMonth.reduce(function(s,r){ return s+(Number(r.orderAmount)||0); }, 0);
    var unlinked = rows.filter(function(r){ return r.orderNo && !r.quoteNo && !r.linked; }).length;

    return {
      success:      true,
      monthlyAmounts: months,
      byStatus:     byStatus,
      topClients:   topClients,
      orderTypeRatio: { trial: trialCount, mass: massCount },
      thisMonth:    { count: thisMonth.length, amount: totalThisMonth },
      unlinkedCount: unlinked,
      totalCases:   rows.length,
    };
  } catch(e) {
    Logger.log('[_apiGetAnalytics] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// ★ CSVエクスポートAPI
// ============================================================
function _apiExportCsv(p) {
  try {
    var sheetType = String(p.sheetType || 'management');
    var ss    = getSpreadsheet();
    var sheet, headers, rowMapper;

    if (sheetType === 'management') {
      sheet  = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
      headers = ['管理ID','見積番号','注文番号','件名','顧客名','ステータス',
                 '見積日','発注日','見積金額','注文金額','消費税','合計金額',
                 '注文種別','機種コード','担当者','納期','メモ','更新日時'];
      rowMapper = function(r) {
        return [r[0],r[1],r[2],r[3],r[4],r[5],
                _toDateStr(r[6]),_toDateStr(r[7]),r[8],r[9],r[10],r[11],
                r[18],r[19],r[21],_toDateStr(r[22]),r[23],_toDateStr(r[25])];
      };
    } else if (sheetType === 'quotes') {
      sheet   = ss.getSheetByName(CONFIG.SHEET_QUOTES);
      headers = ['管理ID','見積番号','発行日','宛先企業','担当者','行No','品名','仕様','数量','単位','単価','金額','備考'];
      rowMapper = function(r) { return r.slice(0, 13); };
    } else {
      sheet   = ss.getSheetByName(CONFIG.SHEET_ORDERS);
      headers = ['管理ID','注文番号','見積番号','注文種別','発注日','機種コード','品名','数量','単価','金額'];
      rowMapper = function(r) { return [r[0],r[1],r[2],r[3],_toDateStr(r[4]),r[5],r[8],r[12],r[14],r[15]]; };
    }

    if (!sheet || sheet.getLastRow() <= 1) return { success: true, csv: '', count: 0 };

    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).getValues();
    var csvRows = [headers];
    data.forEach(function(r) {
      if (!r[0]) return;
      csvRows.push(rowMapper(r).map(function(v) {
        var s = String(v == null ? '' : v);
        return s.indexOf(',') >= 0 || s.indexOf('"') >= 0 || s.indexOf('\n') >= 0
          ? '"' + s.replace(/"/g, '""') + '"' : s;
      }));
    });

    var csv = csvRows.map(function(r){ return r.join(','); }).join('\n');
    // BOM付きUTF-8（Excelで文字化けしない）
    var bom = '\uFEFF';
    var b64 = Utilities.base64Encode(Utilities.newBlob(bom + csv, 'text/csv').getBytes());

    return {
      success: true,
      csv:     b64,
      count:   csvRows.length - 1,
      filename: sheetType + '_' + Utilities.formatDate(new Date(),'Asia/Tokyo','yyyyMMdd') + '.csv',
    };
  } catch(e) {
    Logger.log('[_apiExportCsv] ' + e.message);
    return { success: false, error: e.message };
  }
}

// スプレッドシートをExcel形式でダウンロードするURL
function _apiGetExportUrl(p) {
  try {
    var ss = getSpreadsheet();
    var sheetName = String(p.sheetName || CONFIG.SHEET_MANAGEMENT);
    var sheet = ss.getSheetByName(sheetName);
    var gid   = sheet ? sheet.getSheetId() : 0;
    var url   = 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
                '/export?format=xlsx&gid=' + gid;
    return { success: true, url: url };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// ★ AI月次サマリー生成
// ============================================================
function _apiAiMonthlySummary(p) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') ||
                 CONFIG.GEMINI_API_KEY;
    if (!apiKey) return { success: false, error: 'GEMINI_API_KEY未設定' };

    // 分析データを取得
    var analytics = _apiGetAnalytics({});
    if (!analytics.success) return analytics;

    var today     = new Date();
    var monthLabel = today.getFullYear() + '年' + (today.getMonth()+1) + '月';

    var monthlyStr = analytics.monthlyAmounts.map(function(m) {
      return m.label + ': ¥' + Number(m.amount).toLocaleString();
    }).join('、');

    var statusStr = Object.keys(analytics.byStatus)
      .map(function(k){ return k + ':' + analytics.byStatus[k] + '件'; }).join('、');

    var clientStr = analytics.topClients
      .map(function(c){ return c.client + '(¥' + Number(c.amount).toLocaleString() + ')'; }).join('、');

    var prompt = [
      '以下は見積・注文書管理システムの業務データです。' + monthLabel + 'を対象に、',
      '経営者向けの日本語サマリーを400文字以内で作成してください。',
      '重要なポイント・注意点・改善提案を必ず含めてください。',
      '',
      '【今月データ】',
      '・件数: ' + analytics.thisMonth.count + '件',
      '・金額: ¥' + Number(analytics.thisMonth.amount).toLocaleString(),
      '・未紐づけ注文: ' + analytics.unlinkedCount + '件',
      '',
      '【直近6ヶ月の月別金額】',
      monthlyStr,
      '',
      '【ステータス別件数（全体）】',
      statusStr,
      '',
      '【受注上位顧客】',
      clientStr,
      '',
      '【受注種別】試作:' + analytics.orderTypeRatio.trial + '件 / 量産:' + analytics.orderTypeRatio.mass + '件',
    ].join('\n');

    var model    = CONFIG.GEMINI_PRIMARY_MODEL || 'gemini-1.5-flash';
    var url      = CONFIG.GEMINI_API_ENDPOINT + model + ':generateContent?key=' + apiKey;
    var response = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { maxOutputTokens: 512, temperature: 0.4 }
      }),
      muteHttpExceptions: true,
    });

    var json = JSON.parse(response.getContentText());
    if (json.error) return { success: false, error: json.error.message };

    var text = (json.candidates && json.candidates[0] &&
                json.candidates[0].content && json.candidates[0].content.parts)
               ? json.candidates[0].content.parts.map(function(p){ return p.text||''; }).join('')
               : 'サマリー生成に失敗しました。';

    return {
      success:   true,
      summary:   text,
      period:    monthLabel,
      analytics: analytics,
    };
  } catch(e) {
    Logger.log('[_apiAiMonthlySummary] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// ★ 安全版：案件データ取得
// ============================================================
function _apiGetAll() {

  try {
    var rows = getAllMgmtData();

    // IS_LATEST フィルター（列が存在しない場合も安全に処理）
    rows = rows.filter(function(r) {
      var colIdx = (MGMT_COLS.IS_LATEST || 0) - 1;
      if (colIdx < 0 || colIdx >= r.length) return true; // 列がなければ全件通す
      var v = String(r[colIdx] || '');
      return v === '' || v.toUpperCase() === 'TRUE';
    });

    // 非表示ステータスを除外
    rows = rows.filter(function(r) {
      var hidden = CONFIG.STATUS_HIDDEN || [];
      return hidden.indexOf(String(r[MGMT_COLS.STATUS - 1] || '')) < 0;
    });

    // 注文番号がある行のみ
    var orderRows = rows.filter(function(r) {
      return String(r[MGMT_COLS.ORDER_NO - 1]).trim() !== '';
    });

    var items = _deduplicateMgmtRows(orderRows);

    // 発注日の新しい順
    items.sort(function(a, b) {
      var da = String(a.orderDate || a.quoteDate || '');
      var db = String(b.orderDate || b.quoteDate || '');
      return db.localeCompare(da);
    });

    return { success: true, total: items.length, items: items };
  } catch(e) {
    Logger.log('[_apiGetAll ERROR] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

function _deduplicateMgmtRows(rows) {
  var seen   = {};
  var result = [];
  rows.forEach(function(r) {
    var obj     = _rowToObject(r);
    var orderNo = String(obj.orderNo  || '').trim();
    var quoteNo = String(obj.quoteNo  || '').trim();
    var key = orderNo || quoteNo || obj.id;
    if (!key) return;
    if (seen[key]) {
      var existing = seen[key];
      if (!existing.orderNo     && obj.orderNo)     existing.orderNo     = obj.orderNo;
      if (!existing.quoteNo     && obj.quoteNo)     existing.quoteNo     = obj.quoteNo;
      if (!existing.orderAmount && obj.orderAmount) existing.orderAmount = obj.orderAmount;
      if (!existing.quoteAmount && obj.quoteAmount) existing.quoteAmount = obj.quoteAmount;
      if (!existing.orderDate   && obj.orderDate)   existing.orderDate   = obj.orderDate;
      if (!existing.quoteDate   && obj.quoteDate)   existing.quoteDate   = obj.quoteDate;
      if (!existing.modelCode   && obj.modelCode)   existing.modelCode   = obj.modelCode;
      if (!existing.orderSlipNo && obj.orderSlipNo) existing.orderSlipNo = obj.orderSlipNo;
      if (!existing.deliveryDate && obj.deliveryDate) existing.deliveryDate = obj.deliveryDate;
    } else {
      seen[key] = obj;
      result.push(obj);
    }
  });
  return result;
}

function _apiSearch(p) {
  var kw    = normalizeText(p.keyword || '');
  var st    = p.status    || '';
  var ot    = p.orderType || '';
  var month = p.yearMonth || '';
  var showHidden = p.showHidden || false;
  var data = getAllMgmtData();
  if (!showHidden) {
    data = data.filter(function(r) {
      var hidden = CONFIG.STATUS_HIDDEN || [];
      return hidden.indexOf(String(r[MGMT_COLS.STATUS - 1] || '')) < 0;
    });
  }
  if (kw) data = data.filter(function(r) {
    return [MGMT_COLS.QUOTE_NO,MGMT_COLS.ORDER_NO,MGMT_COLS.ORDER_SLIP_NO,
            MGMT_COLS.MODEL_CODE,MGMT_COLS.SUBJECT,MGMT_COLS.CLIENT]
      .some(function(col) { return normalizeText(String(r[col-1] || '')).indexOf(kw) >= 0; });
  });
  if (st)    data = data.filter(function(r) { return String(r[MGMT_COLS.STATUS-1])      === st; });
  if (ot)    data = data.filter(function(r) { return String(r[MGMT_COLS.ORDER_TYPE-1]) === ot; });
  if (month) data = data.filter(function(r) {
    return String(r[MGMT_COLS.QUOTE_DATE-1]).indexOf(month) === 0 ||
           String(r[MGMT_COLS.ORDER_DATE-1]).indexOf(month) === 0;
  });
  data = data.filter(function(r) {
    return String(r[MGMT_COLS.ORDER_NO - 1]).trim() !== '';
  });
  var items = _deduplicateMgmtRows(data);
  items.sort(function(a, b) {
    var da = String(a.orderDate || '');
    var db = String(b.orderDate || '');
    return db.localeCompare(da);
  });
  return { success: true, total: items.length, items: items };
}

function _apiUploadPdf(p) {
  if (!p.base64Data || !p.fileName || !p.docType)
    return { success: false, error: '必須パラメータ不足' };
  return processUploadedPdf(p.base64Data, p.fileName, p.docType, p.orderType || '');
}

function _apiUpdateStatus(p) {
  if (!p.mgmtId || !p.newStatus) return { success: false, error: '管理IDとステータスが必要' };
  var valid = CONFIG.STATUS_LIST || ['作成予定','送信済み','受領','受注済み','保留','キャンセル','失注','納品済み'];
  if (valid.indexOf(p.newStatus) < 0) return { success: false, error: '無効なステータス' };
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var last  = sheet.getLastRow();
  if (last <= 1) return { success: false, error: 'データなし' };
  var ids = sheet.getRange(2, MGMT_COLS.ID, last-1, 1).getValues().flat();
  var idx = ids.findIndex(function(v) { return String(v) === String(p.mgmtId); });
  if (idx < 0) return { success: false, error: '未発見: ' + p.mgmtId };
  sheet.getRange(idx+2, MGMT_COLS.STATUS).setValue(p.newStatus);
  sheet.getRange(idx+2, MGMT_COLS.UPDATED_AT).setValue(nowJST());
  return { success: true };
}

function _apiAppendBoardRow(p) {
  try {
    if (!p.sheetName || !p.rowData) return { success: false, error: 'パラメータ不足' };
    var ss    = _getBoardSS();
    var sheet = ss.getSheetByName(p.sheetName);
    if (!sheet) return { success: false, error: 'シート「' + p.sheetName + '」が見つかりません' };
    sheet.appendRow(p.rowData);
    Logger.log('[APPEND ROW] ' + p.sheetName + ': ' + JSON.stringify(p.rowData));
    return { success: true };
  } catch(e) {
    Logger.log('[APPEND ROW ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

function _apiUpdateMgmt(p) {
  try {
    if (!p.mgmtId) return { success: false, error: '管理IDが必要です' };
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var last  = sheet.getLastRow();
    if (last <= 1) return { success: false, error: 'データなし' };
    var ids = sheet.getRange(2, MGMT_COLS.ID, last - 1, 1).getValues().flat();
    var idx = ids.map(String).indexOf(String(p.mgmtId));
    if (idx < 0) return { success: false, error: '管理ID未発見: ' + p.mgmtId };
    var row = idx + 2;
    var fields = {
      quoteNo:      MGMT_COLS.QUOTE_NO,
      orderNo:      MGMT_COLS.ORDER_NO,
      subject:      MGMT_COLS.SUBJECT,
      client:       MGMT_COLS.CLIENT,
      status:       MGMT_COLS.STATUS,
      quoteDate:    MGMT_COLS.QUOTE_DATE,
      orderDate:    MGMT_COLS.ORDER_DATE,
      quoteAmount:  MGMT_COLS.QUOTE_AMOUNT,
      orderAmount:  MGMT_COLS.ORDER_AMOUNT,
      orderType:    MGMT_COLS.ORDER_TYPE,
      modelCode:    MGMT_COLS.MODEL_CODE,
      orderSlipNo:  MGMT_COLS.ORDER_SLIP_NO,
      assignee:     MGMT_COLS.ASSIGNEE,
      deliveryDate:  MGMT_COLS.DELIVERY_DATE,
      orderDeadline: MGMT_COLS.ORDER_DEADLINE,
      revisionNo:    MGMT_COLS.REVISION_NO,
      memo:         MGMT_COLS.MEMO,
      linked:       MGMT_COLS.LINKED,
    };
    Object.keys(fields).forEach(function(key) {
      if (p[key] !== undefined) {
        sheet.getRange(row, fields[key]).setValue(p[key]);
      }
    });
    sheet.getRange(row, MGMT_COLS.UPDATED_AT).setValue(nowJST());
    if (p.linkedQuoteNo !== undefined && p.orderNo) {
      _updateLinkedSheetByMgmtId(ss, CONFIG.SHEET_ORDERS, p.mgmtId, 3, p.linkedQuoteNo);
    }
    if (p.linkedOrderNo !== undefined && p.quoteNo) {
      if (p.linkedOrderNo) sheet.getRange(row, MGMT_COLS.ORDER_NO).setValue(p.linkedOrderNo);
      sheet.getRange(row, MGMT_COLS.LINKED).setValue('TRUE');
      sheet.getRange(row, MGMT_COLS.STATUS).setValue(CONFIG.STATUS.RECEIVED);
    }
    return { success: true };
  } catch(e) {
    Logger.log('[UPDATE MGMT ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

function _updateLinkedSheetByMgmtId(ss, sheetName, mgmtId, colIdx, value) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return;
  var mgmtIds = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
  mgmtIds.forEach(function(id, i) {
    if (String(id) === String(mgmtId)) {
      sheet.getRange(i + 2, colIdx).setValue(value);
    }
  });
}

function _apiDeleteMgmt(p) {
  try {
    if (!p.mgmtId) return { success: false, error: '管理IDが必要です' };
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var last  = sheet.getLastRow();
    if (last <= 1) return { success: false, error: 'データなし' };
    var ids = sheet.getRange(2, MGMT_COLS.ID, last - 1, 1).getValues().flat();
    var idx = ids.map(String).indexOf(String(p.mgmtId));
    if (idx < 0) return { success: false, error: '管理ID未発見' };
    sheet.deleteRow(idx + 2);
    if (p.deleteRelated) {
      _deleteRelatedRows(ss, CONFIG.SHEET_QUOTES, p.mgmtId);
      _deleteRelatedRows(ss, CONFIG.SHEET_ORDERS, p.mgmtId);
    }
    return { success: true };
  } catch(e) {
    Logger.log('[DELETE MGMT ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

function _deleteRelatedRows(ss, sheetName, mgmtId) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]) === String(mgmtId)) {
      sheet.deleteRow(i + 2);
    }
  }
}

function _apiGetDetail(p) {
  if (!p.mgmtId) return { success: false, error: '管理IDが必要' };
  var ss = getSpreadsheet();
  var allMgmt    = getAllMgmtData();
  var targetRow  = allMgmt.find(function(r) { return String(r[MGMT_COLS.ID-1]) === String(p.mgmtId); });
  var relatedIds = [String(p.mgmtId)];
  if (targetRow) {
    var orderNo = String(targetRow[MGMT_COLS.ORDER_NO-1] || '').trim();
    var quoteNo = String(targetRow[MGMT_COLS.QUOTE_NO-1] || '').trim();
    allMgmt.forEach(function(r) {
      var id  = String(r[MGMT_COLS.ID-1] || '');
      var oNo = String(r[MGMT_COLS.ORDER_NO-1] || '').trim();
      var qNo = String(r[MGMT_COLS.QUOTE_NO-1] || '').trim();
      if (id === p.mgmtId) return;
      if ((orderNo && oNo === orderNo) || (quoteNo && qNo === quoteNo)) {
        relatedIds.push(id);
      }
    });
  }
  
  var qs    = ss.getSheetByName(CONFIG.SHEET_QUOTES);
  var quoteLines = [];
  if (qs && qs.getLastRow() > 1) {
    var qLast = qs.getLastRow();
    var seenQLines = {};
    qs.getRange(2,1,qLast-1,15).getValues()
      .filter(function(r) { return relatedIds.indexOf(String(r[0])) >= 0; })
      .forEach(function(r) {
        var key = [r[6], r[7], r[8], r[10]].join('|');
        if (seenQLines[key]) return;
        seenQLines[key] = true;
        quoteLines.push({
          quoteNo:r[1], issueDate:_toDateStr(r[2]), destCompany:r[3], destPerson:r[4],
          lineNo:r[5], itemName:r[6], spec:r[7], qty:r[8], unit:r[9],
          unitPrice:r[10], amount:r[11], remarks:r[12], pdfUrl:r[13], folderUrl:r[14]
        });
      });
  }
  
  var os    = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  var orderLines = [];
  if (os && os.getLastRow() > 1) {
    var oLast = os.getLastRow();
    var seenLines = {};
    os.getRange(2,1,oLast-1,19).getValues()
      .filter(function(r) { return relatedIds.indexOf(String(r[0])) >= 0; })
      .forEach(function(r) {
        var key = [r[8], r[9], r[12], r[14]].join('|');
        if (seenLines[key]) return;
        seenLines[key] = true;
        orderLines.push({
          orderType:r[3], orderDate:_toDateStr(r[4]), modelCode:r[5], orderSlipNo:r[6],
          lineNo:r[7], itemName:r[8], spec:r[9], firstDelivery:_toDateStr(r[10]),
          deliveryDest:r[11], qty:r[12], unit:r[13], unitPrice:r[14],
          amount:r[15], remarks:r[16], pdfUrl:r[17], folderUrl:r[18]
        });
      });
  }
  
  // 全見積書明細も返す（注文書明細の候補表示に使用）
  var allQuoteLinesForMatch = [];
  try {
    if (qs && qs.getLastRow() > 1) {
      allQuoteLinesForMatch = qs.getRange(2, 1, qs.getLastRow()-1, 15).getValues()
        .filter(function(r) { return r[6] && r[10]; }) // 品名・単価がある行のみ
        .map(function(r) {
          return {
            mgmtId:    String(r[0] || ''),
            quoteNo:   String(r[1] || ''),
            itemName:  String(r[6] || ''),
            spec:      String(r[7] || ''),
            unitPrice: r[10],
            pdfUrl:    String(r[13] || ''),
          };
        });
    }
  } catch(ex) { Logger.log('[allQuoteLines ERROR] ' + ex.message); }

  return { 
    success: true, 
    mgmtId: p.mgmtId, 
    quoteLines: quoteLines, 
    orderLines: orderLines,
    allQuoteLines: allQuoteLinesForMatch
  };
}

function _apiModelInfoGet(p) {
  try {
    var all = getAllModelInfoData().map(_modelInfoRowToObject);
    if (p.boardId) {
      var found = all.find(function(r) { return r.boardId === String(p.boardId).trim(); });
      return { success: true, item: found || null };
    }
    if (p.modelCode) {
      var items = all.filter(function(r) { return r.modelCode === String(p.modelCode).trim(); });
      return { success: true, items: items };
    }
    return { success: true, items: all };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _apiModelInfoSave(p) {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_MODEL_INFO);
    if (!sheet) return { success: false, error: '基板情報管理シートが見つかりません。' };
    var boardId = String(p.boardId || '').trim();
    if (!boardId) return { success: false, error: '基板IDは必須です' };
    var last = sheet.getLastRow();
    var existingRow = -1;
    if (last > 1) {
      var ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
      var idx = ids.map(String).indexOf(boardId);
      if (idx >= 0) existingRow = idx + 2;
    }
    var rowData = [
      boardId, p.modelCode || '', p.quoteUrl || '', p.orderUrl || '',
      p.purchaseUrl1 || '', p.purchaseUrl2 || '', p.purchaseUrl3 || '',
      p.localServerUrl || '', p.comment || '', nowJST(),
    ];
    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, 10).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    return { success: true, boardId: boardId };
  } catch(e) {
    Logger.log('[MODEL INFO SAVE] ' + e.message);
    return { success: false, error: e.message };
  }
}

function _apiModelInfoUpload(p) {
  try {
    if (!p.base64Data || !p.fileName) return { success: false, error: 'ファイルデータが必要です' };
    var folder   = DriveApp.getFolderById(CONFIG.WEB_UPLOAD_FOLDER_ID);
    var blob     = Utilities.newBlob(Utilities.base64Decode(p.base64Data), 'application/pdf', p.fileName);
    var file     = folder.createFile(blob);
    return { success: true, url: file.getUrl(), fileName: p.fileName };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// Drive検索関連
var DRIVE_SEARCH_FOLDER_ID = '1oAkPV-O4FZezbGmv7sTlNA4wFljMOyv6';
var DRIVE_CACHE_KEY     = 'DRIVE_PDF_CACHE';
var DRIVE_CACHE_TS_KEY  = 'DRIVE_PDF_CACHE_TS';
var DRIVE_CACHE_TTL_MS  = 60 * 60 * 1000;

function _apiDriveSearch(p) {
  try {
    var keyword  = String(p.keyword  || '').trim().toLowerCase();
    var dateFrom = String(p.dateFrom || '').trim();
    var dateTo   = String(p.dateTo   || '').trim();
    var forceRefresh = p.forceRefresh === true;
    var allFiles = _getDrivePdfCache(forceRefresh);
    var filtered = allFiles.filter(function(f) {
      if (keyword && f.name.toLowerCase().indexOf(keyword) < 0) return false;
      if (dateFrom && f.updatedAt < dateFrom) return false;
      if (dateTo   && f.updatedAt > dateTo + 'z') return false;
      return true;
    });
    filtered.sort(function(a, b) {
      return String(b.updatedAt).localeCompare(String(a.updatedAt));
    });
    return { success: true, total: filtered.length, items: filtered.slice(0, 200), cacheTotal: allFiles.length, cached: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _getDrivePdfCache(forceRefresh) {
  var props = PropertiesService.getScriptProperties();
  if (!forceRefresh) {
    var ts  = props.getProperty(DRIVE_CACHE_TS_KEY);
    var now = new Date().getTime();
    if (ts && (now - Number(ts)) < DRIVE_CACHE_TTL_MS) {
      try {
        var cached = props.getProperty(DRIVE_CACHE_KEY);
        if (cached) return JSON.parse(cached);
      } catch(e) { }
    }
  }
  var files = _buildDrivePdfIndex();
  var slim = files.map(function(f) {
    return { id:f.id, name:f.name, url:f.url, updatedAt:f.updatedAt, createdAt:f.createdAt, size:f.size };
  });
  try {
    props.setProperty(DRIVE_CACHE_KEY, JSON.stringify(slim));
    props.setProperty(DRIVE_CACHE_TS_KEY, String(new Date().getTime()));
  } catch(e) {
    var trimmed = slim.slice(0, 2000);
    props.setProperty(DRIVE_CACHE_KEY, JSON.stringify(trimmed));
    props.setProperty(DRIVE_CACHE_TS_KEY, String(new Date().getTime()));
  }
  return slim;
}

function _buildDrivePdfIndex() {
  var MITSUMORI_ROOT = '106Mb1Ucnk_zn-n2UpJkeH78Z2nbcmBCB';
  var allFolderIds   = _getAllSubFolderIds(MITSUMORI_ROOT);
  allFolderIds.unshift(MITSUMORI_ROOT);
  var results  = [];
  var seen     = {};
  var batchSize = 10;
  for (var i = 0; i < allFolderIds.length; i += batchSize) {
    var batch    = allFolderIds.slice(i, i + batchSize);
    var orParts  = batch.map(function(id) { return "'" + id + "' in parents"; });
    var q = '(' + orParts.join(' or ') + ') and mimeType = "application/pdf" and trashed = false';
    try {
      var pageToken = null;
      do {
        var params = { q: q, pageSize: 200, fields: 'nextPageToken, files(id, name, webViewLink, size, modifiedTime, createdTime)', orderBy: 'modifiedTime desc' };
        if (pageToken) params.pageToken = pageToken;
        var resp = Drive.Files.list(params);
        (resp.files || []).forEach(function(f) {
          if (seen[f.id]) return;
          seen[f.id] = true;
          results.push({
            id:        f.id,
            name:      f.name,
            url:       f.webViewLink,
            size:      f.size ? Math.round(Number(f.size) / 1024) : 0,
            updatedAt: f.modifiedTime ? f.modifiedTime.substring(0, 10) : '',
            createdAt: f.createdTime  ? f.createdTime.substring(0, 10)  : '',
          });
        });
        pageToken = resp.nextPageToken;
      } while (pageToken);
    } catch(fe) { }
  }
  return results;
}

function refreshDrivePdfCache() {
  var files = _buildDrivePdfIndex();
  var slim  = files.map(function(f) {
    return { id:f.id, name:f.name, url:f.url, updatedAt:f.updatedAt, createdAt:f.createdAt, size:f.size };
  });
  var props = PropertiesService.getScriptProperties();
  try { props.setProperty(DRIVE_CACHE_KEY, JSON.stringify(slim)); }
  catch(e) { props.setProperty(DRIVE_CACHE_KEY, JSON.stringify(slim.slice(0, 2000))); }
  props.setProperty(DRIVE_CACHE_TS_KEY, String(new Date().getTime()));
  return slim.length;
}

function _apiDriveSearchFallback(p) {
  return { success: false, error: "Not supported" };
}

function _getAllSubFolderIds(rootFolderId) {
  var allIds   = [];
  var queue    = [rootFolderId];
  var visited  = {};
  visited[rootFolderId] = true;
  var maxDepth = 4;
  var depth    = 0;
  while (queue.length > 0 && depth < maxDepth) {
    var currentBatch = queue.slice();
    queue = [];
    depth++;
    var orParts = currentBatch.map(function(id) { return "'" + id + "' in parents"; });
    var batchSize = 10;
    for (var i = 0; i < orParts.length; i += batchSize) {
      var batch = orParts.slice(i, i + batchSize);
      var q = '(' + batch.join(' or ') + ') and mimeType = "application/vnd.google-apps.folder" and trashed = false';
      try {
        var resp = Drive.Files.list({ q: q, fields: 'files(id, name)', pageSize: 200 });
        (resp.files || []).forEach(function(f) {
          if (!visited[f.id]) { visited[f.id] = true; allIds.push(f.id); queue.push(f.id); }
        });
      } catch(e) { }
    }
  }
  return allIds;
}

function _apiDriveDeleteFile(p) {
  try {
    if (!p.fileId) return { success: false, error: 'fileIdが必要です' };
    DriveApp.getFileById(p.fileId).setTrashed(true);
    var props  = PropertiesService.getScriptProperties();
    var cached = props.getProperty(DRIVE_CACHE_KEY);
    if (cached) {
      var files   = JSON.parse(cached);
      var updated = files.filter(function(f) { return f.id !== p.fileId; });
      props.setProperty(DRIVE_CACHE_KEY, JSON.stringify(updated));
    }
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function _apiGetQuoteDetail(p) {
  try {
    if (!p || !p.mgmtId) return { success: false, error: '管理IDが必要です' };
    var ss = getSpreadsheet();
    var mgmtData = getAllMgmtData();
    var targetRow = mgmtData.find(function(r) {
      return String(r[MGMT_COLS.ID - 1]) === String(p.mgmtId);
    });
    if (!targetRow) return { success: false, error: '管理IDが見つかりません: ' + p.mgmtId };
    
    var quoteNo = String(targetRow[MGMT_COLS.QUOTE_NO - 1] || '').trim();
    var relatedIds = [String(p.mgmtId)];
    if (quoteNo) {
      mgmtData.forEach(function(r) {
        var id = String(r[MGMT_COLS.ID - 1] || '');
        if (id === String(p.mgmtId)) return;
        if (String(r[MGMT_COLS.QUOTE_NO - 1] || '').trim() === quoteNo) {
          relatedIds.push(id);
        }
      });
    }
    
    var mgmt = _rowToObject(targetRow);
    var quoteSheet = ss.getSheetByName(CONFIG.SHEET_QUOTES);
    var quoteLines = [];
    if (quoteSheet && quoteSheet.getLastRow() > 1) {
      var qData = quoteSheet.getRange(2, 1, quoteSheet.getLastRow() - 1, 15).getValues();
      var seen = {};
      qData.filter(function(r) {
        return relatedIds.indexOf(String(r[0])) >= 0;
      }).forEach(function(r) {
        var key = [r[6], r[7], r[8], r[10]].join('|');
        if (seen[key]) return;
        seen[key] = true;
        quoteLines.push({
          mgmtId:     String(r[0]),
          quoteNo:    String(r[1] || ''),
          issueDate:  _toDateStr(r[2]),
          destCompany:String(r[3] || ''),
          destPerson: String(r[4] || ''),
          lineNo:     r[5],
          itemName:   String(r[6] || ''),
          spec:       String(r[7] || ''),
          qty:        r[8],
          unit:       String(r[9] || ''),
          unitPrice:  r[10],
          amount:     r[11],
          remarks:    String(r[12] || ''),
          pdfUrl:     String(r[13] || ''),
          folderUrl:  String(r[14] || ''),
        });
      });
    }
    return {
      success:    true,
      mgmt:       mgmt,
      quoteLines: quoteLines,
    };
  } catch(e) {
    Logger.log('[GET QUOTE DETAIL ERROR] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

function _apiGetCandidates() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName('紐づけ候補');
  if (!sheet || sheet.getLastRow() <= 1) return { success: true, items: [] };
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();
  var items = data
    .filter(function(r) { return r[0] && r[15] !== '手動確定済み' && r[11] !== 'auto_linked'; })
    .map(function(r) {
      return {
        orderMgmtId:  String(r[0]),  orderNo:    String(r[1]),
        orderClient:  String(r[2]),  orderDate:  _toDateStr(r[3]),
        orderAmount:  r[4],
        c1Id:    String(r[5]),  c1No:    String(r[6]),
        c1Client:String(r[7]), c1Score: r[8],   c1Detail: String(r[9]),
        c2Id:    String(r[10]), c2Score: r[11],
        c3Id:    String(r[12]), c3Score: r[13],
        status:  String(r[15]), updatedAt: _toDateStr(r[16]),
      };
    });
  return { success: true, items: items };
}

function _apiConfirmLink(p) {
  if (!p.orderMgmtId || !p.quoteMgmtId) return { success: false, error: '管理IDが必要です' };
  return confirmManualLink(p.orderMgmtId, p.quoteMgmtId);
}

function _apiRunMatching(p) {
  var mgmtId = p && p.mgmtId;
  if (mgmtId) return { success: true, result: matchOrderToQuote(mgmtId) };
  return { success: true, result: runBatchMatching() };
}

// ===== 見積書一覧 API =====
function _apiQuoteListGetAll() {
  try {
    var ss         = getSpreadsheet();
    var quoteSheet = ss.getSheetByName(CONFIG.SHEET_QUOTES);
    var mgmtData  = getAllMgmtData();
    
    // 見積番号がある行だけを抽出
    var allRows   = mgmtData.filter(function(r) { return String(r[MGMT_COLS.QUOTE_NO - 1]).trim() !== ''; });
    
    // 重複を排除
    var seenQNo  = {};
    var quoteRows = allRows.filter(function(r) {
      var qNo = String(r[MGMT_COLS.QUOTE_NO - 1]).trim();
      if (seenQNo[qNo]) return false;
      seenQNo[qNo] = true;
      return true;
    });

    var quoteLineMap = {};
    if (quoteSheet && quoteSheet.getLastRow() > 1) {
      var qData = quoteSheet.getRange(2, 1, quoteSheet.getLastRow() - 1, 15).getValues();
      qData.forEach(function(r) {
        var mgmtId = String(r[QUOTE_COLS.MGMT_ID - 1] || '');
        if (!mgmtId || quoteLineMap[mgmtId]) return;
        quoteLineMap[mgmtId] = {
          issueDate:   _toDateStr(r[QUOTE_COLS.ISSUE_DATE  - 1]),
          destCompany: String(r[QUOTE_COLS.DEST_COMPANY - 1] || ''),
          destPerson:  String(r[QUOTE_COLS.DEST_PERSON  - 1] || ''),
        };
      });
    }

    var items = quoteRows.map(function(r) {
      var mgmtId  = String(r[MGMT_COLS.ID - 1] || '');
      var lineInfo = quoteLineMap[mgmtId] || {};
      return {
        id:          mgmtId,
        quoteNo:     String(r[MGMT_COLS.QUOTE_NO - 1]      || ''),
        issueDate:   lineInfo.issueDate   || _toDateStr(r[MGMT_COLS.QUOTE_DATE - 1]),
        destCompany: lineInfo.destCompany || String(r[MGMT_COLS.CLIENT - 1] || ''),
        destPerson:  lineInfo.destPerson  || '',
        quoteAmount: _toNum(r[MGMT_COLS.QUOTE_AMOUNT - 1]),
        status:      String(r[MGMT_COLS.STATUS - 1]        || ''),
        quotePdfUrl: String(r[MGMT_COLS.QUOTE_PDF_URL - 1] || ''),
        orderNo:     String(r[MGMT_COLS.ORDER_NO - 1]      || ''),
        linked:      _isLinkedVal(r[MGMT_COLS.LINKED - 1]),
        orderType:   String(r[MGMT_COLS.ORDER_TYPE - 1]    || ''),
        modelCode:   String(r[MGMT_COLS.MODEL_CODE - 1]    || ''),
        // ★ ここが重要！件名（SUBJECT）をデータに追加
        subject:     String(r[MGMT_COLS.SUBJECT - 1]       || ''),
      };
    });

    items.sort(function(a, b) {
      var da = String(a.issueDate || a.quoteDate || '');
      var db = String(b.issueDate || b.quoteDate || '');
      return db.localeCompare(da);
    });

    return { success: true, total: items.length, items: items };
  } catch(e) { return { success: false, error: e.message }; }
}

function _isLinkedVal(val) { return val === true || val === 'TRUE' || val === 'true'; }

function _apiLedgerGetAll() { return { success: true, items: getAllLedgerData().map(_ledgerRowToObject) }; }

function _apiLedgerSave(p) {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_LEDGER);
    if (!sheet) return { success: false, error: '見積台帳シートが存在しません。' };
    var isNew    = !p.ledgerId;
    var ledgerId = isNew ? generateLedgerId() : p.ledgerId;
    if (isNew) {
      sheet.appendRow([
        ledgerId, p.quoteNo || '', p.issueDate || '', p.dest || '',
        p.category || '', p.subject || '', p.status || LEDGER_STATUS.PENDING,
        p.saveUrl || '', p.machineCode || '', p.boardName || '',
        p.modelNo || '', p.amount !== undefined && p.amount !== '' ? Number(p.amount) : '',
        p.submitTo || '', p.remarks || '',
      ]);
    } else {
      var last = sheet.getLastRow();
      if (last <= 1) return { success: false, error: '対象行が見つかりません' };
      var ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
      var idx = ids.indexOf(String(ledgerId));
      if (idx < 0) return { success: false, error: '台帳IDが見つかりません' };
      var row = idx + 2;
      var fields = {
        quoteNo: LEDGER_COLS.QUOTE_NO, issueDate: LEDGER_COLS.ISSUE_DATE,
        dest: LEDGER_COLS.DEST, category: LEDGER_COLS.CATEGORY,
        subject: LEDGER_COLS.SUBJECT, status: LEDGER_COLS.STATUS,
        saveUrl: LEDGER_COLS.SAVE_URL, machineCode: LEDGER_COLS.MACHINE_CODE,
        boardName: LEDGER_COLS.BOARD_NAME, modelNo: LEDGER_COLS.MODEL_NO,
        amount: LEDGER_COLS.AMOUNT, submitTo: LEDGER_COLS.SUBMIT_TO, remarks: LEDGER_COLS.REMARKS,
      };
      Object.keys(fields).forEach(function(key) {
        if (p[key] !== undefined) sheet.getRange(row, fields[key]).setValue(p[key]);
      });
    }
    return { success: true, ledgerId: ledgerId };
  } catch(e) { return { success: false, error: e.message }; }
}

function _apiLedgerDelete(p) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_LEDGER);
    if (!sheet || !p.ledgerId) return { success: false, error: 'パラメータ不足' };
    var last = sheet.getLastRow();
    var ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
    var idx = ids.indexOf(String(p.ledgerId));
    if (idx < 0) return { success: false, error: '対象行が見つかりません' };
    sheet.deleteRow(idx + 2);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function _apiLedgerUpdateUrl(p) {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_LEDGER);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, matched: false };
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
    var bestIdx = -1, bestScore = 0;
    var normSubject = _normStr(p.subject || '');
    var normDest    = _normStr(p.dest    || '');
    data.forEach(function(row, i) {
      if (String(row[LEDGER_COLS.STATUS - 1]) === LEDGER_STATUS.SENT) return;
      var rowSubject = _normStr(String(row[LEDGER_COLS.SUBJECT - 1]  || ''));
      var rowDest    = _normStr(String(row[LEDGER_COLS.DEST - 1]     || ''));
      var score = 0;
      if (normSubject && rowSubject && _strIncludes(normSubject, rowSubject)) score += 50;
      if (normDest    && rowDest    && _strIncludes(normDest,    rowDest))    score += 40;
      if (p.quoteNo && String(row[LEDGER_COLS.QUOTE_NO - 1]) === String(p.quoteNo)) score += 30;
      if (score > bestScore) { bestScore = score; bestIdx = i; }
    });
    if (bestIdx < 0 || bestScore < 40) return { success: true, matched: false };
    var targetRow = bestIdx + 2;
    sheet.getRange(targetRow, LEDGER_COLS.SAVE_URL).setValue(p.saveUrl || '');
    sheet.getRange(targetRow, LEDGER_COLS.STATUS).setValue(LEDGER_STATUS.SENT);
    if (p.issueDate) sheet.getRange(targetRow, LEDGER_COLS.ISSUE_DATE).setValue(p.issueDate);
    return { success: true, matched: true, row: targetRow, score: bestScore };
  } catch(e) { return { success: false, error: e.message }; }
}

function _normStr(s) { return String(s).replace(/\s+/g,'').replace(/　/g,'').toLowerCase().replace(/株式会社|有限会社|（株）|\(株\)/g,''); }
function _strIncludes(a, b) { return a.indexOf(b) >= 0 || b.indexOf(a) >= 0; }

function _apiLedgerUploadFile(p) {
  try {
    if (!p.base64Data || !p.fileName) return { success: false, error: 'ファイルデータ不足' };
    var folder   = DriveApp.getFolderById(CONFIG.WEB_UPLOAD_FOLDER_ID);
    var mimeType = p.mimeType || 'application/pdf';
    var prefix   = (p.machineCode ? p.machineCode + '_' : '') + (p.boardName ? p.boardName + '_' : '');
    prefix       = prefix.replace(/[/\\:*?"<>|]/g, '_').substring(0, 30);
    var blob     = Utilities.newBlob(Utilities.base64Decode(p.base64Data), mimeType, prefix + p.fileName);
    var file     = folder.createFile(blob);
    return { success: true, url: file.getUrl(), fileId: file.getId(), fileName: p.fileName };
  } catch(e) { return { success: false, error: e.message }; }
}

function _apiLedgerGetMachines() {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_LEDGER);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, machines: [] };
    var col = sheet.getRange(2, LEDGER_COLS.MACHINE_CODE, sheet.getLastRow() - 1, 1).getValues().flat();
    var seen = {}, machines = [];
    col.forEach(function(v) { var m = String(v||'').trim(); if (m && !seen[m]) { seen[m] = true; machines.push(m); } });
    return { success: true, machines: machines.sort() };
  } catch(e) { return { success: false, error: e.message }; }
}

function _apiLedgerCreateMachine(p) {
  try {
    var machineCode = String(p.machineCode || '').trim();
    if (!machineCode) return { success: false, error: '機種コードは必須です' };
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_LEDGER);
    if (!sheet) return { success: false, error: '見積台帳シートが存在しません' };
    if (sheet.getLastRow() > 1) {
      var codes = sheet.getRange(2, LEDGER_COLS.MACHINE_CODE, sheet.getLastRow() - 1, 2).getValues();
      for (var i = 0; i < codes.length; i++) {
        if (String(codes[i][0]).trim() === machineCode && String(codes[i][1]).trim() === '（機種フォルダ）') {
          return { success: false, error: '機種「' + machineCode + '」はすでに登録されています' };
        }
      }
    }
    var ledgerId = generateLedgerId();
    sheet.appendRow([ledgerId, '', '', '', '', '（機種フォルダ）', '__MACHINE_FOLDER__', '', machineCode, '', '', '', '', p.remarks || '']);
    return { success: true, ledgerId: ledgerId, machineCode: machineCode };
  } catch(e) { return { success: false, error: e.message }; }
}

function _apiGetTodos() {
  var items = getAllTodoData().map(function(r) { return { id:r[0], title:r[1], client:r[2], dueDate:r[3], priority:r[4], status:r[5], linkedMgmt:r[6], memo:r[7] }; });
  return { success: true, items: items };
}

function _apiSaveTodo(p) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_TODO);
  if (!sheet) return { success: false, error: 'Todoシートなし' };
  if (p.id) {
    var last = sheet.getLastRow();
    if (last > 1) {
      var ids = sheet.getRange(2,1,last-1,1).getValues().flat();
      var idx = ids.findIndex(function(v) { return String(v) === String(p.id); });
      if (idx >= 0) {
        sheet.getRange(idx+2, 1, 1, 8).setValues([[p.id, p.title||'', p.client||'', p.dueDate||'', p.priority||'中', p.status||'未着手', p.linkedMgmt||'', p.memo||'']]);
        return { success: true, id: p.id };
      }
    }
  }
  var newId = generateTodoId();
  sheet.appendRow([newId, p.title||'', p.client||'', p.dueDate||'', p.priority||'中', p.status||'未着手', p.linkedMgmt||'', p.memo||'']);
  return { success: true, id: newId };
}

function _apiDeleteTodo(p) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_TODO);
  var ids = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().flat();
  var idx = ids.findIndex(function(v) { return String(v) === String(p.id); });
  if (idx >= 0) sheet.deleteRow(idx + 2);
  return { success: true };
}

function _apiGetCalendar(p) {
  var year  = p.year  || new Date().getFullYear();
  var month = p.month || (new Date().getMonth() + 1);
  var ym    = year + '/' + String(month).padStart(2,'0');
  var all   = getAllMgmtData().map(_rowToObject);
  var todos = getAllTodoData().map(function(r) { return { id:r[0], title:r[1], client:r[2], dueDate:r[3], priority:r[4], status:r[5], linkedMgmt:r[6], type:'todo' }; });
  var events = [];
  all.forEach(function(item) {
    if (item.orderDate    && String(item.orderDate).indexOf(ym)    === 0) events.push({ date: item.orderDate,    label: item.client || item.orderNo,           type: 'order',    status: item.status, mgmtId: item.id });
    if (item.deliveryDate && String(item.deliveryDate).indexOf(ym) === 0) events.push({ date: item.deliveryDate, label: '納期: ' + (item.client || item.orderNo), type: 'delivery', status: item.status, mgmtId: item.id });
    if (item.quoteDate    && String(item.quoteDate).indexOf(ym)    === 0) events.push({ date: item.quoteDate,    label: '見積: ' + (item.client || item.quoteNo), type: 'quote',    status: item.status, mgmtId: item.id });
  });
  todos.forEach(function(t) {
    if (t.dueDate && String(t.dueDate).indexOf(ym) === 0) events.push({ date: t.dueDate, label: '📝 ' + t.title, type: 'todo', status: t.status, todoId: t.id });
  });
  return { success: true, year: year, month: month, events: events };
}

function _toDateStr(val) {
  if (!val || val === '') return '';
  if (val instanceof Date) { if (isNaN(val.getTime())) return ''; return Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy/MM/dd'); }
  var s = String(val).trim();
  if (s === '') return '';
  if (s.indexOf('T') > 0 && s.indexOf('Z') > 0) { try { var d = new Date(s); if (!isNaN(d.getTime())) return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd'); } catch(e) {} }
  return s;
}

function _toNum(val) {
  if (val === '' || val === null || val === undefined) return '';
  var n = Number(val);
  return isNaN(n) ? '' : n;
}

// ★ 追加：_rowToObject の完全版定義
function _rowToObject(r) {
  return {
    id:             String(r[MGMT_COLS.ID - 1] || ''),
    quoteNo:        String(r[MGMT_COLS.QUOTE_NO - 1] || ''),
    orderNo:        String(r[MGMT_COLS.ORDER_NO - 1] || ''),
    subject:        String(r[MGMT_COLS.SUBJECT - 1] || ''),
    client:         String(r[MGMT_COLS.CLIENT - 1] || ''),
    status:         String(r[MGMT_COLS.STATUS - 1] || ''),
    quoteDate:      _toDateStr(r[MGMT_COLS.QUOTE_DATE - 1]),
    orderDate:      _toDateStr(r[MGMT_COLS.ORDER_DATE - 1]),
    quoteAmount:    _toNum(r[MGMT_COLS.QUOTE_AMOUNT - 1]),
    orderAmount:    _toNum(r[MGMT_COLS.ORDER_AMOUNT - 1]),
    tax:            _toNum(r[MGMT_COLS.TAX - 1]),
    total:          _toNum(r[MGMT_COLS.TOTAL - 1]),
    quotePdfUrl:    String(r[MGMT_COLS.QUOTE_PDF_URL - 1] || ''),
    orderPdfUrl:    String(r[MGMT_COLS.ORDER_PDF_URL - 1] || ''),
    driveFolderUrl: String(r[MGMT_COLS.DRIVE_FOLDER_URL - 1] || ''),
    linked:         _isLinkedVal(r[MGMT_COLS.LINKED - 1]),
    orderType:      String(r[MGMT_COLS.ORDER_TYPE - 1] || ''),
    modelCode:      String(r[MGMT_COLS.MODEL_CODE - 1] || ''),
    orderSlipNo:    String(r[MGMT_COLS.ORDER_SLIP_NO - 1] || ''),
    assignee:       String(r[MGMT_COLS.ASSIGNEE - 1] || ''),
    deliveryDate:   _toDateStr(r[MGMT_COLS.DELIVERY_DATE  - 1]),
    orderDeadline:  _toDateStr(r[MGMT_COLS.ORDER_DEADLINE - 1]),
    revisionNo:     String(r[MGMT_COLS.REVISION_NO      - 1] || ''),
    isLatest:       String(r[MGMT_COLS.IS_LATEST        - 1] || 'TRUE'),
    memo:           String(r[MGMT_COLS.MEMO - 1] || ''),
    createdAt:      _toDateStr(r[MGMT_COLS.CREATED_AT - 1]),
    updatedAt:      _toDateStr(r[MGMT_COLS.UPDATED_AT - 1]),
  };
}

function _apiSearchDetail(p) {
  var kw = normalizeText(p.keyword || '');
  if (!kw) return { success: true, total: 0, items: [] };
  var ss = getSpreadsheet();
  var qs = ss.getSheetByName(CONFIG.SHEET_QUOTES);
  var os = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  var hitIds = {};
  if (qs && qs.getLastRow() > 1) {
    qs.getRange(2,1,qs.getLastRow()-1,16).getValues().forEach(function(r) {
      var mid = String(r[0]||''); if (!mid) return;
      if (normalizeText([r[6],r[7],r[12]].join(' ')).indexOf(kw) >= 0) {
        if (!hitIds[mid]) hitIds[mid]=[];
        hitIds[mid].push({sheet:'見積書',itemName:String(r[6]||''),spec:String(r[7]||'')});
      }
    });
  }
  if (os && os.getLastRow() > 1) {
    os.getRange(2,1,os.getLastRow()-1,20).getValues().forEach(function(r) {
      var mid = String(r[0]||''); if (!mid) return;
      if (normalizeText([r[8],r[9],r[16]].join(' ')).indexOf(kw) >= 0) {
        if (!hitIds[mid]) hitIds[mid]=[];
        hitIds[mid].push({sheet:'注文書',itemName:String(r[8]||''),spec:String(r[9]||'')});
      }
    });
  }
  var ids = Object.keys(hitIds);
  if (!ids.length) return { success:true, total:0, items:[] };
  var items = getAllMgmtData()
    .filter(function(r){ return ids.indexOf(String(r[MGMT_COLS.ID-1]))>=0; })
    .map(function(r){ var o=_rowToObject(r); o.hitDetails=hitIds[o.id]||[]; return o; });
  return { success:true, total:items.length, items:items };
}

function _apiUpdateLineStatus(p) {
  try {
    if (!p.mgmtId||!p.lineNo||!p.sheetType||p.newStatus===undefined) return {success:false,error:'mgmtId/lineNo/sheetType/newStatus required'};
    var ss = getSpreadsheet();
    var sheetName = p.sheetType==='quote' ? CONFIG.SHEET_QUOTES : CONFIG.SHEET_ORDERS;
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet||sheet.getLastRow()<=1) return {success:false,error:'empty sheet'};
    var colCount = p.sheetType==='quote' ? 16 : 20;
    var lineNoCol = p.sheetType==='quote' ? 5 : 7;
    var updated = 0;
    sheet.getRange(2,1,sheet.getLastRow()-1,colCount).getValues().forEach(function(r,i){
      if (String(r[0])===String(p.mgmtId) && String(r[lineNoCol])===String(p.lineNo)) {
        sheet.getRange(i+2,colCount).setValue(p.newStatus); updated++;
      }
    });
    return updated>0 ? {success:true} : {success:false,error:'row not found'};
  } catch(e){ return {success:false,error:e.message}; }
}

function _apiCreateRevision(p) {
  try {
    if (!p.mgmtId) return {success:false,error:'mgmtId required'};
    var ss = getSpreadsheet();
    var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var last = mgmtSheet.getLastRow();
    if (last<=1) return {success:false,error:'no data'};
    var ids = mgmtSheet.getRange(2,MGMT_COLS.ID,last-1,1).getValues().flat().map(String);
    var idx = ids.indexOf(String(p.mgmtId));
    if (idx<0) return {success:false,error:'not found'};
    var srcRow = idx+2;
    var src = mgmtSheet.getRange(srcRow,1,1,27).getValues()[0];
    var curRev = String(src[MGMT_COLS.REVISION_NO-1]||'A');
    var nextRev = String.fromCharCode(Math.min(curRev.charCodeAt(0)+1,90));
    mgmtSheet.getRange(srcRow,MGMT_COLS.IS_LATEST).setValue('FALSE');
    mgmtSheet.getRange(srcRow,MGMT_COLS.UPDATED_AT).setValue(nowJST());
    var newId = generateMgmtId();
    var newRow = src.slice();
    while(newRow.length<33) newRow.push('');
    newRow[MGMT_COLS.ID-1]            = newId;
    newRow[MGMT_COLS.IS_LATEST-1]     = 'TRUE';
    newRow[MGMT_COLS.PARENT_MGMT_ID-1]= p.mgmtId;
    newRow[MGMT_COLS.REVISION_NO-1]   = nextRev;
    newRow[MGMT_COLS.STATUS-1]        = CONFIG.STATUS.PLANNED;
    newRow[MGMT_COLS.LINKED-1]        = 'FALSE';
    newRow[MGMT_COLS.CREATED_AT-1]    = nowJST();
    newRow[MGMT_COLS.UPDATED_AT-1]    = nowJST();
    newRow[MGMT_COLS.GMAIL_MSG_ID-1]  = '';
    mgmtSheet.appendRow(newRow);
    return {success:true, newMgmtId:newId, revision:nextRev};
  } catch(e){ return {success:false,error:e.message}; }
}

function _apiCheckDeadlines() { checkOrderDeadlines(); return {success:true}; }

function checkOrderDeadlines() {
  var webhookUrl = _getChatWebhookUrl();
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var last  = sheet.getLastRow();
  if (last<=1) return;
  var today = new Date(); today.setHours(0,0,0,0);
  sheet.getRange(2,1,last-1,29).getValues().forEach(function(row,i){
    var orderNo = String(row[MGMT_COLS.ORDER_NO-1]||'');
    var dlVal   = row[MGMT_COLS.ORDER_DEADLINE-1];
    if (!orderNo||!dlVal) return;
    var dl = new Date(dlVal); dl.setHours(0,0,0,0);
    var diff = Math.round((dl-today)/86400000);
    var flagStr = String(row[MGMT_COLS.DEADLINE_NOTIFIED-1]||'');
    var level = diff<0&&flagStr.indexOf('over')<0?'over':diff===0&&flagStr.indexOf('due')<0?'due':diff<=3&&flagStr.indexOf('3day')<0?'3day':'';
    if (!level) return;
    var subj = String(row[MGMT_COLS.SUBJECT-1]||orderNo);
    var label = level==='over'?('期限超過('+Math.abs(diff)+'日)'):level==='due'?'本日が期限':(diff+'日後が期限');
    if (webhookUrl) {
      try {
        UrlFetchApp.fetch(webhookUrl,{method:'post',contentType:'application/json',
          payload:JSON.stringify({text:'⚠️ *注文書期限アラート*\n• '+subj+'\n• '+label+'\n• 期限: '+Utilities.formatDate(dl,'Asia/Tokyo','yyyy/MM/dd')}),
          muteHttpExceptions:true});
      } catch(e){}
    }
    var newFlag = flagStr ? flagStr+','+level : level;
    sheet.getRange(i+2,MGMT_COLS.DEADLINE_NOTIFIED).setValue(newFlag);
    sheet.getRange(i+2,MGMT_COLS.UPDATED_AT).setValue(nowJST());
  });
}

var SETTINGS_KEY = 'SYS_SETTINGS';
function _apiSaveSettings(p) {
  try {
    var props = PropertiesService.getScriptProperties();
    if (p.adminEmails !== undefined) props.setProperty('ADMIN_EMAILS', p.adminEmails);
    if (p.chatWebhook !== undefined) props.setProperty('GOOGLE_CHAT_WEBHOOK_URL', p.chatWebhook);

    // ★ Gemini APIキーの保存・反映
    if (p.geminiKey && String(p.geminiKey).trim() !== '') {
      var newKey = String(p.geminiKey).trim();
      props.setProperty('GEMINI_API_KEY', newKey);
      // CONFIGに即座反映（同じGAS実行コンテキスト内）
      CONFIG.GEMINI_API_KEY = newKey;
      Logger.log('[SETTINGS] Gemini APIキーを更新しました');
    }

    // ★ スプレッドシートIDの保存
    if (p.spreadsheetId && String(p.spreadsheetId).trim() !== '') {
      props.setProperty('SPREADSHEET_ID', String(p.spreadsheetId).trim());
      CONFIG.SPREADSHEET_ID = String(p.spreadsheetId).trim();
    }

    // ★ その他の追加設定の保存
    if (p.notifyEmails !== undefined) props.setProperty('NOTIFY_EMAILS', p.notifyEmails);
    if (p.rakurakuApiKey !== undefined && String(p.rakurakuApiKey).trim() !== '') props.setProperty('RAKURAKU_API_KEY', p.rakurakuApiKey);
    if (p.rakurakuCompany !== undefined) props.setProperty('RAKURAKU_COMPANY', p.rakurakuCompany);
    if (p.rakurakuEndpoint !== undefined) props.setProperty('RAKURAKU_ENDPOINT', p.rakurakuEndpoint);
    if (p.n8nOrderWebhook !== undefined) props.setProperty('N8N_ORDER_WEBHOOK', p.n8nOrderWebhook);
    if (p.n8nQuoteWebhook !== undefined) props.setProperty('N8N_QUOTE_WEBHOOK', p.n8nQuoteWebhook);
    if (p.customWebhook !== undefined)   props.setProperty('CUSTOM_WEBHOOK', p.customWebhook);

    var current = _loadSettingsObj();
    var merged = {
      webhookUrl:  p.chatWebhook  !== undefined ? String(p.chatWebhook)  : (p.webhookUrl !== undefined ? String(p.webhookUrl) : current.webhookUrl),
      notifyOrder: p.notifyOrder  !== undefined ? !!p.notifyOrder        : current.notifyOrder,
      notifyQuote: p.notifyQuote  !== undefined ? !!p.notifyQuote        : current.notifyQuote,
      notifyDl:    p.notifyDl     !== undefined ? !!p.notifyDl           : current.notifyDl,
      alertDays:   p.alertDays    !== undefined ? Number(p.alertDays)||3  : current.alertDays,
    };
    props.setProperty(SETTINGS_KEY, JSON.stringify(merged));
    if (merged.webhookUrl) props.setProperty('GOOGLE_CHAT_WEBHOOK_URL', merged.webhookUrl);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function _apiLoadSettings() {
  try {
    var props = PropertiesService.getScriptProperties();
    var s = _loadSettingsObj();
    return {
      success:        true,
      settings:       s,
      adminEmails:    props.getProperty('ADMIN_EMAILS')            || '',
      chatWebhook:    props.getProperty('GOOGLE_CHAT_WEBHOOK_URL') || s.webhookUrl || '',
      spreadsheetId:  props.getProperty('SPREADSHEET_ID')          || '',  // ★追加
      notifyEmails:   props.getProperty('NOTIFY_EMAILS')           || '',  // ★追加
      rakurakuCompany:  props.getProperty('RAKURAKU_COMPANY')      || '',  // ★追加
      rakurakuEndpoint: props.getProperty('RAKURAKU_ENDPOINT')     || '',  // ★追加
      n8nOrderWebhook:  props.getProperty('N8N_ORDER_WEBHOOK')     || '',  // ★追加
      n8nQuoteWebhook:  props.getProperty('N8N_QUOTE_WEBHOOK')     || '',  // ★追加
      customWebhook:    props.getProperty('CUSTOM_WEBHOOK')        || '',  // ★追加
    };
  } catch(e) { return { success: false, error: e.message }; }
}

function _getChatWebhookUrl() {
  return PropertiesService.getScriptProperties().getProperty('GOOGLE_CHAT_WEBHOOK_URL') ||
         _loadSettingsObj().webhookUrl || '';
}

// \u2605 Gemini API \u63a5\u7d9a\u30c6\u30b9\u30c8\nfunction _apiTestGeminiConnection() {\n  try {\n    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') ||\n                 CONFIG.GEMINI_API_KEY;\n    if (!apiKey) return { success: false, error: 'GEMINI_API_KEY\u304c\u672a\u767b\u9332\u3067\u3059\u3002\u7ba1\u7406\u30b3\u30f3\u30bd\u30fc\u30eb\u306b\u3066API\u30ad\u30fc\u3092\u767b\u9332\u3057\u3066\u304f\u3060\u3055\u3044\u3002' };\n\n    var model = CONFIG.GEMINI_PRIMARY_MODEL || 'gemini-1.5-flash';\n    var url   = CONFIG.GEMINI_API_ENDPOINT + model + ':generateContent?key=' + apiKey;\n    var res = UrlFetchApp.fetch(url, {\n      method: 'post',\n      contentType: 'application/json',\n      payload: JSON.stringify({\n        contents: [{ parts: [{ text: '\u30c6\u30b9\u30c8\u3002OK\u3068\u3060\u3051\u8fd4\u3057\u3066\u3002' }] }],\n        generationConfig: { maxOutputTokens: 10, temperature: 0 }\n      }),\n      muteHttpExceptions: true,\n    });\n    var code = res.getResponseCode();\n    if (code === 200) {\n      return { success: true, model: model, message: 'Gemini API\u63a5\u7d9a\u6210\u529f' };\n    } else if (code === 400) {\n      return { success: false, error: 'API\u30ea\u30af\u30a8\u30b9\u30c8\u30a8\u30e9\u30fc (400): ' + res.getContentText().substring(0, 200) };\n    } else if (code === 401 || code === 403) {\n      return { success: false, error: 'API\u30ad\u30fc\u304c\u7121\u52b9\u304b\u6a29\u9650\u306a\u3057 (' + code + ')\u3002\u6b63\u3057\u3044\u7121\u6599\u7248API\u30ad\u30fc\u3092\u767b\u9332\u3057\u3066\u304f\u3060\u3055\u3044\u3002' };\n    } else if (code === 429) {\n      return { success: false, error: 'API\u30ec\u30fc\u30c8\u30ea\u30df\u30c3\u30c8 (429): \u3057\u3070\u3089\u304f\u5f85\u3063\u3066\u518d\u5ea6\u30c6\u30b9\u30c8\u3057\u3066\u304f\u3060\u3055\u3044\u3002\u7121\u6599\u67a0\u5185\u3067\u3059\u3002' };\n    } else {\n      return { success: false, error: 'HTTP ' + code + ': ' + res.getContentText().substring(0, 200) };\n    }\n  } catch(e) {\n    return { success: false, error: e.message };\n  }\n}\n\nfunction _apiTestWebhook(p) {
  var url = String(p.url || '').trim();
  if (!url) return { success: false, error: 'URLが必要です' };
  try {
    var res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ text: '✅ 見積・注文書管理システム: Webhook接続テスト成功！ ' + nowJST() }),
      muteHttpExceptions: true
    });
    return { success: res.getResponseCode() === 200 };
  } catch(e) { return { success: false, error: e.message }; }
}

function _apiSendAnnouncement() {
  return { success: false, error: '周知メール送信は無効化されています。' };
}

function notifyQuoteImported(info) {
  try {
    var settings   = _loadSettingsObj();
    var webhookUrl = _getChatWebhookUrl();
    var amountStr  = info.amount ? '¥' + Number(info.amount).toLocaleString() : '—';
    var rowUrl     = _getMgmtRowUrl(info.mgmtId);
    var appUrl     = ScriptApp.getService().getUrl();
    var lines = [
      '【📄 見積書を登録しました】',
      '案件名: ' + (info.subject || '—'),
      '顧客名: ' + (info.client  || '—'),
      '金額: '   + amountStr,
      '',
      '🌐 システムで確認（転記された行）',
      rowUrl || appUrl,
    ];
    if (info.pdfUrl)    lines.push('📎 PDF: '     + info.pdfUrl);
    if (info.folderUrl) lines.push('📁 フォルダ: ' + info.folderUrl);

    // ① Google Chat Webhook 通知
    if (webhookUrl && settings.notifyQuote) {
      _postToChat(webhookUrl, lines.join('\n'));
    }

    // ② Gmail メール通知
    var subject  = '【見積書登録】' + (info.client || '') + ' / ' + (info.subject || '') + ' / ' + amountStr;
    var bodyText = lines.join('\n');
    var htmlBody = _buildEmailHtml('見積書を登録しました', [
      ['案件名', info.subject || '—'],
      ['顧客名', info.client  || '—'],
      ['金額',   amountStr],
      ['見積No', info.quoteNo || '—'],
      ['PDF',    info.pdfUrl  ? '<a href="' + info.pdfUrl + '">開く</a>' : '—'],
    ], appUrl);
    _sendNotifyEmail(subject, bodyText, htmlBody);

    // ③ 監査ログ
    _writeAuditLog('quote_import', '顧客:' + (info.client||'—') + ' / 金額:' + amountStr, 'success');
  } catch(e) {
    Logger.log('[NOTIFY QUOTE ERROR] ' + e.message);
  }
}

function notifyOrderImported(info) {
  try {
    var settings   = _loadSettingsObj();
    var webhookUrl = _getChatWebhookUrl();
    var amountStr  = info.amount ? '¥' + Number(info.amount).toLocaleString() : '—';
    var rowUrl     = _getMgmtRowUrl(info.mgmtId);
    var appUrl     = ScriptApp.getService().getUrl();
    var lr         = info.linkResult || {};
    var lines = [
      '【📦 注文書を登録しました】',
      '発注書番号: ' + (info.orderNo || '—'),
      '顧客名: '    + (info.client  || '—'),
      '金額: '      + amountStr,
      '',
      '🌐 システムで確認（転記された行）',
      rowUrl || appUrl,
    ];
    if (info.pdfUrl)    lines.push('📎 PDF: '     + info.pdfUrl);
    if (info.folderUrl) lines.push('📁 フォルダ: ' + info.folderUrl);
    lines.push('');

    if (lr.status === 'auto_linked') {
      lines.push('✅ 見積書と自動紐づけ済み（スコア: ' + (lr.score || '—') + '点）');
      lines.push('見積No: ' + (lr.quoteNo || '—'));
      if (lr.quoteUrl) lines.push('📄 紐づく見積書PDF: ' + lr.quoteUrl);
    } else if (lr.status === 'candidates_found' || lr.status === 'forced_candidate') {
      lines.push('⚠️ 紐づく見積書の候補があります（要確認）');
      var candidates = lr.candidates || [];
      candidates.slice(0, 3).forEach(function(c, i) {
        lines.push('');
        lines.push('候補' + (i + 1) + ': 見積No.' + (c.quoteNo || '—') +
                   '  スコア: ' + (c.score || '—') + '点' +
                   (c.keywords ? '  品名: ' + String(c.keywords).substring(0, 20) : ''));
        if (c.quoteUrl) lines.push('  📄 見積書PDF: ' + c.quoteUrl);
      });
      lines.push('');
      lines.push('▶ 紐づけ確定はシステムから行ってください');
      lines.push(appUrl);
    }

    // ① Google Chat Webhook 通知
    if (webhookUrl && settings.notifyOrder) {
      _postToChat(webhookUrl, lines.join('\n'));
    }

    // ② Gmail メール通知
    var linkStatus = lr.status === 'auto_linked' ? '✅ 自動紐づけ済み' :
                     (lr.status === 'candidates_found' ? '⚠️ 候補あり（要確認）' : '—');
    var subject  = '【注文書登録】' + (info.client || '') + ' / ' + (info.orderNo || '') + ' / ' + amountStr;
    var bodyText = lines.join('\n');
    var htmlBody = _buildEmailHtml('注文書を登録しました', [
      ['発注書番号', info.orderNo  || '—'],
      ['顧客名',     info.client   || '—'],
      ['金額',       amountStr],
      ['紐づけ状況', linkStatus],
    ], appUrl);
    _sendNotifyEmail(subject, bodyText, htmlBody);

    // ③ 監査ログ
    _writeAuditLog('order_import', '顧客:' + (info.client||'—') + ' / 金額:' + amountStr + ' / 紐づけ:' + linkStatus, 'success');
  } catch(e) {
    Logger.log('[NOTIFY ORDER ERROR] ' + e.message);
  }
}


function _getMgmtRowUrl(mgmtId) {
  try {
    if (!mgmtId) return '';
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    if (!sheet || sheet.getLastRow() <= 1) return '';
    var ids = sheet.getRange(2, MGMT_COLS.ID, sheet.getLastRow() - 1, 1)
                   .getValues().flat();
    var idx = ids.map(String).indexOf(String(mgmtId));
    if (idx < 0) return '';
    return ss.getUrl() + '&gid=' + sheet.getSheetId() + '&range=A' + (idx + 2);
  } catch(e) {
    Logger.log('[GET ROW URL ERROR] ' + e.message);
    return '';
  }
}

function _postToChat(webhookUrl, text) {
  try {
    UrlFetchApp.fetch(webhookUrl, {
      method:          'post',
      contentType:     'application/json',
      payload:         JSON.stringify({ text: text }),
      muteHttpExceptions: true,
    });
    Logger.log('[CHAT NOTIFY] 送信: ' + text.substring(0, 80));
  } catch(e) {
    Logger.log('[CHAT POST ERROR] ' + e.message);
  }
}

function _apiGetMatchingCandidates(p) {
  return getMatchingCandidates();
}

function _apiConfirmOrderLink(p) {
  if (!p.orderMgmtId || !p.quoteMgmtId) return { success: false, error: 'IDが不足しています' };
  return confirmManualLink(p.orderMgmtId, p.quoteMgmtId);
}

function _apiRunBatchMatching(p) {
  return runBatchMatching();
}

function _apiUploadOrderWithLink(p) {
  const res = _apiUploadPdf(p);
  if (res.success && res.mgmtId) {
    const aiRes = aiLinkOrderToQuote(res.mgmtId);
    res.linkResult = aiRes;
  }
  return res;
}

// ============================================================
// ★ 監査ログ取得 API
// ============================================================
function _apiGetAuditLog(p) {
  try {
    var limit = Math.min(Number(p.limit || 100), 500);
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName('監査ログ');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, logs: [] };
    var lastRow  = sheet.getLastRow();
    var startRow = Math.max(2, lastRow - limit + 1);
    var numRows  = lastRow - startRow + 1;
    var data     = sheet.getRange(startRow, 1, numRows, 5).getValues();
    // 新しい順に並び替え
    data.reverse();
    var logs = data.map(function(r) {
      return [
        r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss') : String(r[0]),
        String(r[1] || ''), String(r[2] || ''), String(r[3] || ''), String(r[4] || '')
      ];
    });
    return { success: true, logs: logs, total: data.length };
  } catch(e) {
    return _errorResponse(e, '監査ログ取得');
  }
}