// ============================================================
// 見積書・注文書管理システム
// ファイル 3/4: Webアプリ doGet / API ルーター (完全版・超軽量通信)
// ============================================================

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

  var dashboardTmpl = HtmlService.createTemplateFromFile('Dashboard');
  dashboardTmpl.isAdmin = isAdmin;
  dashboardTmpl.userEmail = userEmail;

  return dashboardTmpl.evaluate()
    .setTitle('見積・注文 管理システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// ★ 安全な通信ルーター（エラーの元凶を排除）
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
        try {
          _registerTriggers();
          try { setupMonitoringTriggers(); } catch(e2) { Logger.log('[reregisterTriggers] monitoring: ' + e2.message); }
          try { setupModelMasterTriggers(); } catch(e3) { Logger.log('[reregisterTriggers] modelMaster: ' + e3.message); }
          return { success: true };
        } catch(e) { return { success: false, error: e.message }; }
      case 'orderPriceCompare':      res = apiOrderPriceCompare(payload); break;
      case 'batchOrderPriceCompare': res = apiBatchOrderPriceCompare(payload); break;
      case 'searchUnitPrice':
        res = { success: true, results: searchUnitPrice(payload.itemName, payload.spec, payload.client) }; break;
      case 'searchCaseSummary':
        var _kws = String(payload.keywords || '').split(/[\s,　]+/).filter(Boolean);
        res = { success: true, results: searchCaseSummary(_kws) }; break;
      case 'syncDetailDB':
        syncAllDetailDB(); res = { success: true, message: '明細DB同期完了' }; break;
      case 'rebuildUnitPrice':
        rebuildUnitPriceMaster(); res = { success: true, message: '単価マスタ再構築完了' }; break;
      case 'ocrPreview':  res = apiOcrPreview(payload);  break;
      case 'ocrApprove':  res = apiOcrApprove(payload);  break;
      case 'ocrDiscard':  res = apiOcrDiscard(payload);  break;
      // ── メニュー設定（管理コンソール） ──
      case 'menuConfigLoad':        res = _apiMenuConfigLoad(); break;
      case 'menuConfigSave':        res = _apiMenuConfigSave(payload); break;
      // ── システム設定（管理コンソール） ──
      case 'loadSettings':          res = _apiLoadSettings(); break;
      case 'saveSettings':          res = _apiSaveSettings(payload); break;
      // ── 注文書↔見積書 紐づけ ──
      case 'confirmOrderLink':      res = _apiConfirmOrderLink(payload); break;
      case 'getOrderLinkCandidates': res = _apiGetOrderLinkCandidates(payload); break;
      // ── OCR使用量 ──
      case 'getOcrUsage':           res = _apiGetOcrUsage(); break;
      // ── 機種マスタ（21_model_master.gs）──
      case 'modelMasterList':        res = apiModelMasterList(); break;
      case 'modelMasterGet':         res = apiModelMasterGet(payload); break;
      case 'modelMasterSave':        res = apiModelMasterSave(payload); break;
      case 'modelMasterDelete':      res = apiModelMasterDelete(payload); break;
      case 'rebuildModelMaster':     res = apiRebuildModelMaster(); break;
      case 'syncLedgerModelCodes':   res = syncLedgerModelCodes(); break;
      case 'syncModelNameToMgmt':    res = syncModelNameToMgmt(payload.modelCode, payload.modelName); break;
      case 'initModelMasterSheet':   res = initModelMasterSheet(); break;
      // ── 外部API連携管理 ──
      case 'externalApiList':    res = _apiExternalApiList(); break;
      case 'externalApiSave':    res = _apiExternalApiSave(payload); break;
      case 'externalApiDelete':  res = _apiExternalApiDelete(payload); break;
      case 'externalApiTest':    res = _apiExternalApiTest(payload); break;
      case 'externalApiTrigger': res = _apiExternalApiTrigger(payload); break;
      // ── 発注期限アラート（10_alert_monitor.gs）──
      case 'checkDeadlines': {
        var dlAlerts = checkOrderDeadlines();
        res = { success: true, alertCount: dlAlerts.length, alerts: dlAlerts }; break;
      }
      case 'runAllMonitoring': {
        var monRes = runAllMonitoring();
        res = { success: true, alertCount: monRes.alertCount, alerts: monRes.alerts }; break;
      }
      case 'sendTestAlert': res = _apiSendTestAlert(); break;
      // ── 明細ステータス・改訂版 ──
      case 'updateLineStatus': res = _apiUpdateLineStatus(payload); break;
      case 'createRevision':   res = _apiCreateRevision(payload); break;
      // ── 見積セット チェックリスト（11_checklist_api.gs）──
      case 'checklistGetByModel':   return _apiChecklistGetByModel(payload);
      case 'checklistSaveItem':     return _apiChecklistSaveItem(payload);
      case 'checklistDeleteItem':   return _apiChecklistDeleteItem(payload);
      case 'checklistReorder':      return _apiChecklistReorder(payload);
      case 'checklistGetSections':  return _apiChecklistGetSections(payload);
      case 'checklistCopyTemplate': return _apiChecklistCopyTemplate(payload);
      // ── カスタム書類テンプレート（22_custom_doc.gs）──
      case 'customDocTypeList':   return apiCustomDocTypeList();
      case 'customDocTypeSave':   return apiCustomDocTypeSave(payload);
      case 'customDocTypeDelete': return apiCustomDocTypeDelete(payload);
      case 'customDocGetAll':     return apiCustomDocGetAll(payload);
      case 'customDocSave':       return apiCustomDocSave(payload);
      case 'customDocDelete':     return apiCustomDocDelete(payload);
      case 'customDocOcr':        return apiCustomDocOcr(payload);
      case 'customDocPresets':    return apiCustomDocPresets();
      default: return { success: false, error: '不明なアクション: ' + action };
    }
    
    // ★ エラーの元凶だった「JSON.parse(JSON.stringify(res))」を完全削除し、そのまま返す
    return res;

  } catch(e) {
    Logger.log('[API ERROR] ' + action + ': ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

// ============================================================
// ★ 超高速＆軽量版：案件データ取得
// ============================================================
function _apiGetAll() {
  try {
    var rows = getAllMgmtData();

    rows = rows.filter(function(r) {
      var colIdx = (MGMT_COLS.IS_LATEST || 0) - 1;
      if (colIdx < 0 || colIdx >= r.length) return true;
      var v = String(r[colIdx] || '');
      return v === '' || v.toUpperCase() === 'TRUE';
    });

    rows = rows.filter(function(r) {
      var hidden = CONFIG.STATUS_HIDDEN || [];
      return hidden.indexOf(String(r[MGMT_COLS.STATUS - 1] || '')) < 0;
    });

    var orderRows = rows.filter(function(r) {
      return String(r[MGMT_COLS.ORDER_NO - 1]).trim() !== '';
    });

    var items = _deduplicateMgmtRows(orderRows);

    var ss = getSpreadsheet();
    var os = ss.getSheetByName(CONFIG.SHEET_ORDERS);
    if (os && os.getLastRow() > 1) {
      var orderData = os.getDataRange().getValues();
      orderData.shift(); 

      // 検索を高速化するための一括マッピング
      var orderLinesMap = {};
      orderData.forEach(function(r) {
        var mgmtId = String(r[0]);
        if (!mgmtId) return;
        if (!orderLinesMap[mgmtId]) orderLinesMap[mgmtId] = [];
        // 品名(8列目), 仕様(9列目)の文字だけを抽出
        orderLinesMap[mgmtId].push(String(r[8] || '') + ' ' + String(r[9] || ''));
      });

      items.forEach(function(item) {
        if (orderLinesMap[item.id]) {
          item.detailText = orderLinesMap[item.id].join(' ').substring(0, 300); // 通信制限に引っかからないよう文字数カット
        }
      });
    }

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

// 案件重複排除ロジック
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

// ============================================================
// ★ 超高速＆軽量版：見積書一覧 API
// ============================================================
function _apiQuoteListGetAll() {
  try {
    var ss         = getSpreadsheet();
    var quoteSheet = ss.getSheetByName(CONFIG.SHEET_QUOTES);
    var mgmtData  = getAllMgmtData();
    
    var allRows   = mgmtData.filter(function(r) { return String(r[MGMT_COLS.QUOTE_NO - 1]).trim() !== ''; });
    
    var seenQNo  = {};
    var quoteRows = allRows.filter(function(r) {
      var qNo = String(r[MGMT_COLS.QUOTE_NO - 1]).trim();
      if (seenQNo[qNo]) return false;
      seenQNo[qNo] = true;
      return true;
    });

    var quoteLineMap = {};
    if (quoteSheet && quoteSheet.getLastRow() > 1) {
      var qData = quoteSheet.getDataRange().getValues();
      qData.shift(); 

      qData.forEach(function(r) {
        var mgmtId = String(r[QUOTE_COLS.MGMT_ID - 1] || '');
        if (!mgmtId) return;
        if (!quoteLineMap[mgmtId]) {
          quoteLineMap[mgmtId] = {
            issueDate:   _toDateStr(r[QUOTE_COLS.ISSUE_DATE  - 1]),
            destCompany: String(r[QUOTE_COLS.DEST_COMPANY - 1] || ''),
            destPerson:  String(r[QUOTE_COLS.DEST_PERSON  - 1] || ''),
            linesText:   '' 
          };
        }
        // 品名(6列目), 仕様(7列目) のみ抽出
        quoteLineMap[mgmtId].linesText += String(r[6] || '') + ' ' + String(r[7] || '') + ' ';
      });
    }

    var items = quoteRows.map(function(r) {
      var mgmtId  = String(r[MGMT_COLS.ID - 1] || '');
      var lineInfo = quoteLineMap[mgmtId] || { linesText: '' };
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
        subject:     String(r[MGMT_COLS.SUBJECT - 1]       || ''),
        detailText:  String(lineInfo.linesText).substring(0, 300) // 軽量化
      };
    });

    // 発行日の降順（新しい順）ソート — yyyy/MM/dd 形式で統一されているため直接比較可
    items.sort(function(a, b) {
      var da = String(a.issueDate || '').replace(/\//g, '-');
      var db = String(b.issueDate || '').replace(/\//g, '-');
      if (db > da) return 1;
      if (db < da) return -1;
      // 発行日が同じ場合は見積Noの降順
      return String(b.quoteNo || '').localeCompare(String(a.quoteNo || ''));
    });

    return { success: true, total: items.length, items: items };
  } catch(e) { return { success: false, error: e.message }; }
}

function _isLinkedVal(val) { return val === true || val === 'TRUE' || val === 'true'; }

// ──────────────────────────────────────────
// その他 API機能（既存維持）
// ──────────────────────────────────────────
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
  if (!p.base64Data || !p.fileName || !p.docType) return { success: false, error: '必須パラメータ不足' };
  return processUploadedPdf(p.base64Data, p.fileName, p.docType, p.orderType || '');
}

// ============================================================
// ステータス同期: 案件管理 ↔ 見積台帳（双方向）
// ============================================================

/** 案件管理ステータス → 見積台帳ステータス への変換マップ */
function _mgmtStatusToLedgerStatus(mgmtStatus) {
  var map = {
    '作成予定':  '作成予定',
    '送信済み':  '送信済み',
    '受領':      '受注済み',
    '受注済み':  '受注済み',
    '納品済み':  '受注済み',
    'キャンセル':'キャンセル',
    '失注':      'キャンセル',
    '保留':      null,   // 台帳に相当するステータスなし → 変更しない
  };
  return map.hasOwnProperty(mgmtStatus) ? map[mgmtStatus] : null;
}

/** 見積台帳ステータス → 案件管理ステータス への変換マップ */
function _ledgerStatusToMgmtStatus(ledgerStatus) {
  var map = {
    '作成予定':  '作成予定',
    '作成中':    '作成予定',
    '送信済み':  '送信済み',
    '受注済み':  '受注済み',
    'キャンセル':'キャンセル',
  };
  return map.hasOwnProperty(ledgerStatus) ? map[ledgerStatus] : null;
}

/**
 * quoteNo で見積台帳を検索しステータスを更新する
 * @param {Spreadsheet} ss
 * @param {string} quoteNo
 * @param {string} mgmtStatus  案件管理側の新ステータス
 */
function _syncLedgerStatus(ss, quoteNo, mgmtStatus) {
  if (!quoteNo) return;
  var ledgerStatus = _mgmtStatusToLedgerStatus(mgmtStatus);
  if (!ledgerStatus) return;                        // マッピング対象外はスキップ
  try {
    var sheet = ss.getSheetByName(CONFIG.SHEET_LEDGER);
    if (!sheet || sheet.getLastRow() <= 1) return;
    var qNos = sheet.getRange(2, LEDGER_COLS.QUOTE_NO,
                              sheet.getLastRow() - 1, 1).getValues().flat();
    qNos.forEach(function(qno, i) {
      if (String(qno).trim() === String(quoteNo).trim()) {
        sheet.getRange(i + 2, LEDGER_COLS.STATUS).setValue(ledgerStatus);
      }
    });
  } catch(e) {
    Logger.log('[_syncLedgerStatus ERROR] ' + e.message);
  }
}

/**
 * ledgerId で案件管理シートを検索しステータスを更新する
 * （台帳→管理の逆同期）
 * @param {Spreadsheet} ss
 * @param {string} quoteNo  台帳の見積番号
 * @param {string} ledgerStatus  台帳側の新ステータス
 */
function _syncMgmtStatusFromLedger(ss, quoteNo, ledgerStatus) {
  if (!quoteNo) return;
  var mgmtStatus = _ledgerStatusToMgmtStatus(ledgerStatus);
  if (!mgmtStatus) return;
  try {
    var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    if (!sheet || sheet.getLastRow() <= 1) return;
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1,
                              Math.max(MGMT_COLS.QUOTE_NO, MGMT_COLS.STATUS)).getValues();
    data.forEach(function(row, i) {
      if (String(row[MGMT_COLS.QUOTE_NO - 1]).trim() === String(quoteNo).trim()) {
        sheet.getRange(i + 2, MGMT_COLS.STATUS).setValue(mgmtStatus);
        sheet.getRange(i + 2, MGMT_COLS.UPDATED_AT).setValue(nowJST());
      }
    });
  } catch(e) {
    Logger.log('[_syncMgmtStatusFromLedger ERROR] ' + e.message);
  }
}

// ============================================================
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
  // ★ 見積台帳へ同期
  var quoteNo = String(sheet.getRange(idx+2, MGMT_COLS.QUOTE_NO).getValue() || '').trim();
  _syncLedgerStatus(ss, quoteNo, p.newStatus);
  return { success: true };
}

function _apiAppendBoardRow(p) {
  try {
    if (!p.sheetName || !p.rowData) return { success: false, error: 'パラメータ不足' };
    var ss    = _getBoardSS();
    var sheet = ss.getSheetByName(p.sheetName);
    if (!sheet) return { success: false, error: 'シート「' + p.sheetName + '」が見つかりません' };
    sheet.appendRow(p.rowData);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
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
    // ③ データ連携: modelCode が更新された場合は機種マスタへ自動登録
    if (p.modelCode) {
      try { _ensureModelCode(p.modelCode, ''); } catch(e2) {}
    }
    if (p.linkedQuoteNo !== undefined && p.orderNo) {
      _updateLinkedSheetByMgmtId(ss, CONFIG.SHEET_ORDERS, p.mgmtId, 3, p.linkedQuoteNo);
    }
    if (p.linkedOrderNo !== undefined && p.quoteNo) {
      if (p.linkedOrderNo) sheet.getRange(row, MGMT_COLS.ORDER_NO).setValue(p.linkedOrderNo);
      sheet.getRange(row, MGMT_COLS.LINKED).setValue('TRUE');
      sheet.getRange(row, MGMT_COLS.STATUS).setValue(CONFIG.STATUS.RECEIVED);
    }
    // ★ ステータスが更新された場合は見積台帳へ同期
    if (p.status !== undefined) {
      var curQuoteNo = String(sheet.getRange(row, MGMT_COLS.QUOTE_NO).getValue() || '').trim();
      _syncLedgerStatus(ss, curQuoteNo, p.status);
    }
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
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
  } catch(e) { return { success: false, error: e.message }; }
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
  } catch(ex) {}

  return { success: true, mgmtId: p.mgmtId, quoteLines: quoteLines, orderLines: orderLines, allQuoteLines: allQuoteLinesForMatch };
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
  } catch(e) { return { success: false, error: e.message }; }
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
  } catch(e) { return { success: false, error: e.message }; }
}

function _apiModelInfoUpload(p) {
  try {
    if (!p.base64Data || !p.fileName) return { success: false, error: 'ファイルデータが必要です' };
    var folder   = DriveApp.getFolderById(CONFIG.WEB_UPLOAD_FOLDER_ID);
    var blob     = Utilities.newBlob(Utilities.base64Decode(p.base64Data), 'application/pdf', p.fileName);
    var file     = folder.createFile(blob);
    return { success: true, url: file.getUrl(), fileName: p.fileName };
  } catch(e) { return { success: false, error: e.message }; }
}

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
  } catch(e) { return { success: false, error: e.message }; }
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

function _apiDriveSearchFallback(p) { return { success: false, error: "Not supported" }; }

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
    if (!targetRow) return { success: false, error: '管理IDが見つかりません' };
    
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
    return { success: true, mgmt: mgmt, quoteLines: quoteLines };
  } catch(e) { return { success: false, error: e.message }; }
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

function _apiLedgerGetAll() { return { success: true, items: getAllLedgerData().map(_ledgerRowToObject) }; }

function _apiLedgerSave(p) {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_LEDGER);
    if (!sheet) return { success: false, error: '見積台帳シートが存在しません。' };
    var isNew    = !p.ledgerId;
    var ledgerId = isNew ? generateLedgerId() : p.ledgerId;
    // ③ データ連携: machineCode が設定されていれば機種マスタへ自動登録
    if (p.machineCode) {
      try { _ensureModelCode(p.machineCode, p.boardName || ''); } catch(e2) {}
    }

    if (isNew) {
      sheet.appendRow([
        ledgerId, p.quoteNo || '', p.issueDate || '', p.dest || '',
        p.category || '', p.subject || '', p.status || LEDGER_STATUS.PENDING,
        p.saveUrl || '', p.machineCode || '', p.boardName || '',
        p.modelNo || '', p.amount !== undefined && p.amount !== '' ? Number(p.amount) : '',
        p.submitTo || '', p.remarks || '', p.sentDate || '',
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
        amount: LEDGER_COLS.AMOUNT, submitTo: LEDGER_COLS.SUBMIT_TO,
        remarks: LEDGER_COLS.REMARKS,
        sentDate: LEDGER_COLS.SENT_DATE,
      };
      Object.keys(fields).forEach(function(key) {
        if (p[key] !== undefined) sheet.getRange(row, fields[key]).setValue(p[key]);
      });
      // ★ ステータスが更新された場合は案件管理シートへ逆同期
      if (p.status !== undefined) {
        var ledgerQuoteNo = String(sheet.getRange(row, LEDGER_COLS.QUOTE_NO).getValue() || '').trim();
        _syncMgmtStatusFromLedger(getSpreadsheet(), ledgerQuoteNo, p.status);
      }
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
    if (item.orderDate    && String(item.orderDate).indexOf(ym)    === 0) events.push({ date: item.orderDate,    label: item.client || item.orderNo,            type: 'order',    status: item.status, mgmtId: item.id });
    if (item.deliveryDate && String(item.deliveryDate).indexOf(ym) === 0) events.push({ date: item.deliveryDate, label: '納期: ' + (item.client || item.orderNo), type: 'delivery', status: item.status, mgmtId: item.id });
    if (item.quoteDate    && String(item.quoteDate).indexOf(ym)    === 0) events.push({ date: item.quoteDate,    label: '見積: ' + (item.client || item.quoteNo), type: 'quote',    status: item.status, mgmtId: item.id });
  });
  todos.forEach(function(t) {
    if (t.dueDate && String(t.dueDate).indexOf(ym) === 0) events.push({ date: t.dueDate, label: '📝 ' + t.title, type: 'todo', status: t.status, todoId: t.id });
  });
  return { success: true, year: year, month: month, events: events };
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

// _apiCheckDeadlines は handleApiRequest の 'checkDeadlines' ケースに統合済み

function _apiUploadOrderWithLink(p) {
  var res = _apiUploadPdf(p);
  if (res.success && res.mgmtId) {
    var aiRes = aiLinkOrderToQuote(res.mgmtId);
    res.linkResult = aiRes;
  }
  return res;
}

// ── 外部API連携管理 ──
var EXTERNAL_API_KEY = 'EXTERNAL_API_CONNECTIONS';

function _apiExternalApiList() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(EXTERNAL_API_KEY);
    var items = raw ? JSON.parse(raw) : [];
    // APIキーはマスキングして返す
    var masked = items.map(function(c) {
      return { id:c.id, name:c.name, type:c.type, endpoint:c.endpoint, trigger:c.trigger, fields:c.fields, desc:c.desc, enabled:c.enabled, hasApiKey:!!(c.apikey) };
    });
    return { success: true, items: masked };
  } catch(e) { return { success: false, error: e.message }; }
}

function _apiExternalApiSave(p) {
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(EXTERNAL_API_KEY);
    var items = raw ? JSON.parse(raw) : [];
    if (p.id) {
      var idx = items.findIndex(function(c){ return c.id === p.id; });
      if (idx >= 0) {
        // apikey が空の場合は既存のものを保持
        if (!p.apikey) p.apikey = items[idx].apikey;
        items[idx] = Object.assign(items[idx], p);
      } else {
        items.push(p);
      }
    } else {
      p.id = 'api_' + new Date().getTime().toString(36);
      p.createdAt = new Date().toISOString();
      items.push(p);
    }
    props.setProperty(EXTERNAL_API_KEY, JSON.stringify(items));
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function _apiExternalApiDelete(p) {
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(EXTERNAL_API_KEY);
    var items = raw ? JSON.parse(raw) : [];
    items = items.filter(function(c){ return c.id !== p.id; });
    props.setProperty(EXTERNAL_API_KEY, JSON.stringify(items));
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function _apiExternalApiTest(p) {
  try {
    var endpoint = p.endpoint || '';
    if (!endpoint) return { success: false, error: 'エンドポイントが指定されていません' };
    var type = p.type || 'gas_webhook';
    if (type === 'spreadsheet') {
      // スプレッドシートIDの場合は開けるか確認
      var ss = SpreadsheetApp.openById(endpoint);
      return { success: true, statusCode: 200, message: 'スプレッドシート接続OK: ' + ss.getName() };
    }
    // Webhook / REST API
    var options = { method: 'get', muteHttpExceptions: true };
    if (p.apikey) options.headers = { 'Authorization': 'Bearer ' + p.apikey };
    var response = UrlFetchApp.fetch(endpoint, options);
    var code = response.getResponseCode();
    return { success: code >= 200 && code < 400, statusCode: code };
  } catch(e) { return { success: false, error: e.message }; }
}

function _apiExternalApiTrigger(p) {
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(EXTERNAL_API_KEY);
    var items = raw ? JSON.parse(raw) : [];
    var conn = items.find(function(c){ return c.id === p.id; });
    if (!conn) return { success: false, error: '連携設定が見つかりません' };
    if (!conn.enabled) return { success: false, error: 'この連携は無効です' };
    return _sendExternalApi(conn, { trigger: 'manual', timestamp: new Date().toISOString() });
  } catch(e) { return { success: false, error: e.message }; }
}

function _sendExternalApi(conn, data) {
  try {
    var endpoint = conn.endpoint || '';
    if (!endpoint) return { success: false, error: 'エンドポイント未設定' };
    if (conn.type === 'spreadsheet') {
      // スプレッドシートへの書き込みは別途実装
      return { success: true, message: 'スプレッドシート連携は手動操作が必要です' };
    }
    var payload = JSON.stringify(data);
    var options = { method: 'post', contentType: 'application/json', payload: payload, muteHttpExceptions: true };
    if (conn.apikey) options.headers = { 'Authorization': 'Bearer ' + conn.apikey };
    var response = UrlFetchApp.fetch(endpoint, options);
    var code = response.getResponseCode();
    return { success: code >= 200 && code < 400, statusCode: code };
  } catch(e) { return { success: false, error: e.message }; }
}