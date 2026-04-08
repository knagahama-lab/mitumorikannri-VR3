// ============================================================
// 見積書・注文書管理システム
// ファイル 3/4: Webアプリ doGet / API ルーター
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
// dashboard.html 内で <?!= include("dashboard_js") ?> のように使用
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function handleApiRequest(action, payload) {
  try {
    payload = payload || {};
    switch (action) {
      case 'getAll':        return _apiGetAll();
      case 'search':        return _apiSearch(payload);
      case 'uploadPdf':     return _apiUploadPdf(payload);
      case 'updateStatus':  return _apiUpdateStatus(payload);
      case 'getDetail':     return _apiGetDetail(payload);
      case 'getTodos':      return _apiGetTodos();
      case 'saveTodo':      return _apiSaveTodo(payload);
      case 'deleteTodo':    return _apiDeleteTodo(payload);
      case 'getCalendar':   return _apiGetCalendar(payload);
      case 'getCandidates': return _apiGetCandidates();
      case 'confirmLink':   return _apiConfirmLink(payload);
      case 'runMatching':   return _apiRunMatching(payload);
      // ===== 見積台帳API =====
      case 'quoteListGetAll':  return _apiQuoteListGetAll();
      case 'getQuoteDetail':   return _apiGetQuoteDetail(payload);
      case 'ledgerGetAll':    return _apiLedgerGetAll();
      case 'ledgerSave':      return _apiLedgerSave(payload);
      case 'ledgerDelete':    return _apiLedgerDelete(payload);
      case 'ledgerUpdateUrl':   return _apiLedgerUpdateUrl(payload);
      case 'ledgerUploadFile':  return _apiLedgerUploadFile(payload);
      case 'ledgerGetMachines': return _apiLedgerGetMachines();
      case 'ledgerCreateMachine': return _apiLedgerCreateMachine(payload);
      // ===== 基板マスタ・機種マスタ新規登録 =====
      case 'appendBoardRow':   return _apiAppendBoardRow(payload);
      // ===== 編集・削除API =====
      case 'updateMgmt':       return _apiUpdateMgmt(payload);
      case 'deleteMgmt':       return _apiDeleteMgmt(payload);
      // ===== チャットボットAPI =====
      case 'chatbotQuery':    return apiChatbotQuery(payload);
      // ===== 機種情報管理API =====
      case 'modelInfoGet':    return _apiModelInfoGet(payload);
      case 'modelInfoSave':   return _apiModelInfoSave(payload);
      case 'modelInfoUpload': return _apiModelInfoUpload(payload);
      // ===== Drive検索API =====
      case 'driveSearch':       return _apiDriveSearch(payload);
      case 'driveRefreshCache': return { success: true, count: refreshDrivePdfCache() };
      case 'driveDeleteFile':   return _apiDriveDeleteFile(payload);
      // ===== 見積提出管理API =====
      case 'qsGetAll':      return _apiQsGetAll(payload);
      case 'qsSave':        return _apiQsSave(payload);
      case 'qsDelete':      return _apiQsDelete(payload);
      case 'qsUploadFile':  return _apiQsUploadFile(payload);
      case 'qsGetMachines': return _apiQsGetMachines();
      // ===== 部品・PCB 原価管理API =====
      case 'partsGetAll':      return _apiPartsGetAll(payload);
      case 'partsSave':        return _apiPartsSave(payload);
      case 'partsDelete':      return _apiPartsDelete(payload);
      case 'partsImportCSV':   return _apiPartsImportCSV(payload);
      case 'partsExportCSV':   return _apiPartsExportCSV(payload);
      case 'pcbGetAll':        return _apiPcbGetAll(payload);
      case 'pcbSave':          return _apiPcbSave(payload);
      case 'pcbDelete':        return _apiPcbDelete(payload);
      case 'pcbImportCSV':     return _apiPcbImportCSV(payload);
      case 'pcbExportCSV':     return _apiPcbExportCSV(payload);
      case 'ensurePartsSheets': return ensurePartsSheets() || { success: true };
      // ===== 基板管理API =====
      case 'boardGetAll':       return apiBoardGetAll();
      case 'boardAddNew':       return _apiBoardAddNew(payload);
      case 'machineAddNew':     return _apiMachineAddNew(payload);
      case 'boardGetDetail':    return apiBoardGetDetail(payload.boardId, payload.boardName);
      case 'boardGetAnalysis':  return apiGetBoardAnalysis();
      case 'boardGetOrders':    return apiGetOrdersWithBoardInfo();
      case 'boardComparePrice': return apiComparePriceToBOM(payload.mgmtId);
      // ===== 注文書PDF登録＋AI紐づけ＋Chat通知 =====
      case 'searchDetail':           return _apiSearchDetail(payload);
      case 'updateLineStatus':       return _apiUpdateLineStatus(payload);
      case 'createRevision':         return _apiCreateRevision(payload);
      case 'checkDeadlines':         return _apiCheckDeadlines();
      case 'saveSettings':           return _apiSaveSettings(payload);
      case 'loadSettings':           return _apiLoadSettings();
      case 'testWebhook':            return _apiTestWebhook(payload);
      case 'sendAnnouncement':       return _apiSendAnnouncement();
      case 'uploadOrderWithLink':    return _apiUploadOrderWithLink(payload);
      case 'confirmOrderLink':         return _apiConfirmOrderLink(payload);
      case 'getOrderLinkCandidates':   return _apiGetOrderLinkCandidates(payload);
      default: return { success: false, error: '不明なアクション: ' + action };
    }
  } catch(e) {
    Logger.log('[API ERROR] ' + action + ': ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

// ===== 案件データ取得 =====

function _apiGetAll() {
  var rows = getAllMgmtData();
  rows = rows.filter(function(r) {
    var v = String(r[MGMT_COLS.IS_LATEST - 1] || '');
    return v === '' || v === 'TRUE' || v === true;
  });
  rows = rows.filter(function(r) {
    var hidden = CONFIG.STATUS_HIDDEN || [];
    return hidden.indexOf(String(r[MGMT_COLS.STATUS - 1] || '')) < 0;
  });
  // 注文書一覧：注文番号がある行のみ
  var orderRows = rows.filter(function(r) {
    return String(r[MGMT_COLS.ORDER_NO - 1]).trim() !== '';
  });
  var items = _deduplicateMgmtRows(orderRows);
  // 発注日の新しい順にソート
  items.sort(function(a, b) {
    var da = String(a.orderDate || a.quoteDate || '');
    var db = String(b.orderDate || b.quoteDate || '');
    return db.localeCompare(da);
  });
  return { success: true, total: items.length, items: items };
}

/**
 * 管理シートの重複行を除去してグループ化する
 */
function _deduplicateMgmtRows(rows) {
  var seen   = {};   // key → オブジェクト
  var result = [];

  rows.forEach(function(r) {
    var obj     = _rowToObject(r);
    var orderNo = String(obj.orderNo  || '').trim();
    var quoteNo = String(obj.quoteNo  || '').trim();

    // グループキー：注文番号優先、なければ見積番号、両方なければ管理ID
    var key = orderNo || quoteNo || obj.id;
    if (!key) return;

    if (seen[key]) {
      // 既存行を更新（空欄の項目を補完）
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
      .some(function(col) { return normalizeText(String(r[col-1])).indexOf(kw) >= 0; });
  });
  if (st)    data = data.filter(function(r) { return String(r[MGMT_COLS.STATUS-1])      === st; });
  if (ot)    data = data.filter(function(r) { return String(r[MGMT_COLS.ORDER_TYPE-1]) === ot; });
  if (month) data = data.filter(function(r) {
    return String(r[MGMT_COLS.QUOTE_DATE-1]).indexOf(month) === 0 ||
           String(r[MGMT_COLS.ORDER_DATE-1]).indexOf(month) === 0;
  });
  // 注文書のみ
  data = data.filter(function(r) {
    return String(r[MGMT_COLS.ORDER_NO - 1]).trim() !== '';
  });
  var items = _deduplicateMgmtRows(data);
  // 発注日降順
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

// ============================================================
// ⑤ 基板マスタ・機種マスタ 新規行追加API
// ============================================================

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

// ============================================================
// ① 管理シート 編集API
// ============================================================

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

    // 更新可能フィールドのみ更新（undefined は無視）
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

    // 見積書シート・注文書シートの紐づけ番号も更新
    if (p.linkedQuoteNo !== undefined && p.orderNo) {
      _updateLinkedSheetByMgmtId(ss, CONFIG.SHEET_ORDERS, p.mgmtId, 3, p.linkedQuoteNo); // 注文書シート 3列目=紐づけ見積番号
    }
    if (p.linkedOrderNo !== undefined && p.quoteNo) {
      // 管理シートの ORDER_NO を更新（見積書に注文書を紐づけ）
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

// ============================================================
// ① 管理シート 削除API
// ============================================================

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

    // 見積書シート・注文書シートの関連行も削除
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
  // 後ろから削除しないとインデックスがずれる
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

  // 管理シートから同じ注文番号・見積番号を持つ全管理IDを収集
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

  // 見積書明細
  var qs    = ss.getSheetByName(CONFIG.SHEET_QUOTES);
  var qLast = qs.getLastRow();
  var quoteLines = [];
  if (qLast > 1) {
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

  // 注文書明細
  var os    = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  var oLast = os.getLastRow();
  var orderLines = [];
  if (oLast > 1) {
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

  return { success: true, mgmtId: p.mgmtId, quoteLines: quoteLines, orderLines: orderLines };
}

// ===== 機種情報管理 API =====

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

// Drive検索関連の変数は省略せず記載
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
  // すでに実装済み
}

// ===== 紐づけ API =====
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
  return { success: true, result: batchMatchAllUnlinked() };
}

// ===== 見積書一覧 API =====
function _apiQuoteListGetAll() {
  try {
    var ss         = getSpreadsheet();
    var mgmtSheet  = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
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

// ===== 見積台帳 API =====
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

// ===== Todo API =====
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

// ===== カレンダーAPI =====
function _apiGetCalendar(p) {
  var year  = p.year  || new Date().getFullYear();
  var month = p.month || (new Date().getMonth() + 1);
  var ym    = year + '/' + String(month).padStart(2,'0');

  var all   = getAllMgmtData().map(_rowToObject);
  var todos = getAllTodoData().map(function(r) { return { id:r[0], title:r[1], client:r[2], dueDate:r[3], priority:r[4], status:r[5], linkedMgmt:r[6], type:'todo' }; });
  var events = [];

  all.forEach(function(item) {
    if (item.orderDate && String(item.orderDate).indexOf(ym) === 0) events.push({ date: item.orderDate, label: item.client || item.orderNo, type: 'order', status: item.status, mgmtId: item.id });
    if (item.deliveryDate && String(item.deliveryDate).indexOf(ym) === 0) events.push({ date: item.deliveryDate, label: '納期: ' + (item.client || item.orderNo), type: 'delivery', status: item.status, mgmtId: item.id });
    if (item.quoteDate && String(item.quoteDate).indexOf(ym) === 0) events.push({ date: item.quoteDate, label: '見積: ' + (item.client || item.quoteNo), type: 'quote', status: item.status, mgmtId: item.id });
  });

  todos.forEach(function(t) {
    if (t.dueDate && String(t.dueDate).indexOf(ym) === 0) events.push({ date: t.dueDate, label: '📝 ' + t.title, type: 'todo', status: t.status, todoId: t.id });
  });

  return { success: true, year: year, month: month, events: events };
}

// ===== 行→オブジェクト変換 =====
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

function _rowToObject(row) {
  return {
    id:             String(row[MGMT_COLS.ID - 1] || ''),
    quoteNo:        String(row[MGMT_COLS.QUOTE_NO - 1] || ''),
    orderNo:        String(row[MGMT_COLS.ORDER_NO - 1] || ''),
    subject:        String(row[MGMT_COLS.SUBJECT - 1] || ''),
    client:         String(row[MGMT_COLS.CLIENT - 1] || ''),
    status:         String(row[MGMT_COLS.STATUS - 1] || ''),
    quoteDate:      _toDateStr(row[MGMT_COLS.QUOTE_DATE - 1]),
    orderDate:      _toDateStr(row[MGMT_COLS.ORDER_DATE - 1]),
    quoteAmount:    _toNum(row[MGMT_COLS.QUOTE_AMOUNT - 1]),
    orderAmount:    _toNum(row[MGMT_COLS.ORDER_AMOUNT - 1]),
    tax:            _toNum(row[MGMT_COLS.TAX - 1]),
    total:          _toNum(row[MGMT_COLS.TOTAL - 1]),
    quotePdfUrl:    String(row[MGMT_COLS.QUOTE_PDF_URL - 1] || ''),
    orderPdfUrl:    String(row[MGMT_COLS.ORDER_PDF_URL - 1] || ''),
    driveFolderUrl: String(row[MGMT_COLS.DRIVE_FOLDER_URL - 1] || ''),
    linked:         String(row[MGMT_COLS.LINKED - 1] || ''),
    orderType:      String(row[MGMT_COLS.ORDER_TYPE - 1] || ''),
    modelCode:      String(row[MGMT_COLS.MODEL_CODE - 1] || ''),
    orderSlipNo:    String(row[MGMT_COLS.ORDER_SLIP_NO - 1] || ''),
    assignee:       String(row[MGMT_COLS.ASSIGNEE - 1] || ''),
    deliveryDate:   _toDateStr(row[MGMT_COLS.DELIVERY_DATE  - 1]),
    orderDeadline:  _toDateStr(row[MGMT_COLS.ORDER_DEADLINE - 1]),
    revisionNo:     String(row[MGMT_COLS.REVISION_NO      - 1] || ''),
    isLatest:       String(row[MGMT_COLS.IS_LATEST        - 1] || 'TRUE'),
    memo:           String(row[MGMT_COLS.MEMO - 1] || ''),
    createdAt:      _toDateStr(row[MGMT_COLS.CREATED_AT - 1]),
    updatedAt:      _toDateStr(row[MGMT_COLS.UPDATED_AT - 1]),
  };
}

// ============================================================
// ★⑩ 明細横断検索
// ============================================================
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

// ============================================================
// ★⑥ 明細ステータス更新
// ============================================================
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

// ============================================================
// ★⑨ 改版
// ============================================================
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

// ============================================================
// ★③ デッドラインアラート
// ============================================================
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
    var label = level==='over'?('\u671F\u9650\u8D85\u904E('+Math.abs(diff)+'\u65E5)'):level==='due'?'\u672C\u65E5\u304C\u671F\u9650':(diff+'\u65E5\u5F8C\u304C\u671F\u9650');
    if (webhookUrl) {
      try {
        UrlFetchApp.fetch(webhookUrl,{method:'post',contentType:'application/json',
          payload:JSON.stringify({text:'\u26A0\uFE0F *\u6CE8\u6587\u66F8\u671F\u9650\u30A2\u30E9\u30FC\u30C8*\n\u2022 '+subj+'\n\u2022 '+label+'\n\u2022 \u671F\u9650: '+Utilities.formatDate(dl,'Asia/Tokyo','yyyy/MM/dd')}),
          muteHttpExceptions:true});
      } catch(e){}
    }
    var newFlag = flagStr ? flagStr+','+level : level;
    sheet.getRange(i+2,MGMT_COLS.DEADLINE_NOTIFIED).setValue(newFlag);
    sheet.getRange(i+2,MGMT_COLS.UPDATED_AT).setValue(nowJST());
  });
}

// ============================================================
// ★④ 通知設定 保存・読み込み API
// ============================================================
var SETTINGS_KEY = 'SYS_SETTINGS';
function _apiSaveSettings(p) {
  try {
    var current = _loadSettingsObj();
    var merged = {
      webhookUrl:   p.webhookUrl   !== undefined ? String(p.webhookUrl)   : current.webhookUrl,
      notifyOrder:  p.notifyOrder  !== undefined ? !!p.notifyOrder        : current.notifyOrder,
      notifyQuote:  p.notifyQuote  !== undefined ? !!p.notifyQuote        : current.notifyQuote,
      notifyDl:     p.notifyDl     !== undefined ? !!p.notifyDl           : current.notifyDl,
      alertDays:    p.alertDays    !== undefined ? Number(p.alertDays)||3  : current.alertDays,
    };
    PropertiesService.getScriptProperties().setProperty(SETTINGS_KEY, JSON.stringify(merged));
    if (merged.webhookUrl) PropertiesService.getScriptProperties().setProperty('GOOGLE_CHAT_WEBHOOK_URL', merged.webhookUrl);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function _apiLoadSettings() {
  try { return { success: true, settings: _loadSettingsObj() }; } catch(e) { return { success: false, error: e.message }; }
}

function _loadSettingsObj() {
  var raw = PropertiesService.getScriptProperties().getProperty(SETTINGS_KEY);
  var defaults = { webhookUrl: CONFIG.GOOGLE_CHAT_WEBHOOK_URL || '', notifyOrder: true, notifyQuote: true, notifyDl: true, alertDays: 3 };
  if (!raw) return defaults;
  try { return Object.assign(defaults, JSON.parse(raw)); } catch(e) { return defaults; }
}

function _apiTestWebhook(p) {
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

// ============================================================
// ★⓪ 周知メール一括送信
// ============================================================
function _apiSendAnnouncement() {
  return { success: false, error: '周知メール送信は無効化されています。' };
}

// ============================================================
// ★システム設定 (GUIから)
// ============================================================
function _apiLoadSettings() {
  var props = PropertiesService.getScriptProperties();
  return {
    success: true,
    adminEmails: props.getProperty('ADMIN_EMAILS') || '',
    chatWebhook: props.getProperty('GOOGLE_CHAT_WEBHOOK_URL') || ''
  };
}

function _apiSaveSettings(p) {
  try {
    var props = PropertiesService.getScriptProperties();
    if (p.adminEmails !== undefined) props.setProperty('ADMIN_EMAILS', p.adminEmails);
    if (p.chatWebhook !== undefined) props.setProperty('GOOGLE_CHAT_WEBHOOK_URL', p.chatWebhook);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}