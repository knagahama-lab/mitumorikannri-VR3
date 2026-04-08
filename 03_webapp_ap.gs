// ============================================================
// 見積書・注文書管理システム
// ファイル 3/4: Webアプリ doGet / API ルーター
// ============================================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('Dashboard').evaluate()
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
      case 'uploadOrderWithLink':      return _apiUploadOrderWithLink(payload);
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
  // 注文書一覧：注文番号がある行のみ（見積書のみの案件は除外）
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
 * 同じ「注文番号」または「見積番号」を持つ行は1件にまとめる
 * 金額・日付などは最初に登録された行の値を使用
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
  var kw    = (p.keyword   || '').toLowerCase();
  var st    = p.status    || '';
  var ot    = p.orderType || '';
  var month = p.yearMonth || '';

  var data = getAllMgmtData();
  if (kw) data = data.filter(function(r) {
    return [MGMT_COLS.QUOTE_NO,MGMT_COLS.ORDER_NO,MGMT_COLS.ORDER_SLIP_NO,
            MGMT_COLS.MODEL_CODE,MGMT_COLS.SUBJECT,MGMT_COLS.CLIENT]
      .some(function(col) { return String(r[col-1]).toLowerCase().indexOf(kw) >= 0; });
  });
  if (st)    data = data.filter(function(r) { return String(r[MGMT_COLS.STATUS-1])     === st; });
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
  var valid = ['作成予定','送信済み','受領','受注済み','納品済み'];
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
      deliveryDate: MGMT_COLS.DELIVERY_DATE,
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
  // （重複除去で代表IDのみ残っているため、全関連IDの明細を引く必要がある）
  var allMgmt    = getAllMgmtData();
  var targetRow  = allMgmt.find(function(r) { return String(r[MGMT_COLS.ID-1]) === String(p.mgmtId); });
  var relatedIds = [String(p.mgmtId)];

  if (targetRow) {
    var orderNo = String(targetRow[MGMT_COLS.ORDER_NO-1] || '').trim();
    var quoteNo = String(targetRow[MGMT_COLS.QUOTE_NO-1] || '').trim();
    // 同じ注文番号または見積番号を持つ全管理IDを収集
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

  // 見積書明細（関連する全管理IDの明細を取得）
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

  // 注文書明細（関連する全管理IDの明細を取得・重複除去）
  var os    = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  var oLast = os.getLastRow();
  var orderLines = [];
  if (oLast > 1) {
    var seenLines = {};
    os.getRange(2,1,oLast-1,19).getValues()
      .filter(function(r) { return relatedIds.indexOf(String(r[0])) >= 0; })
      .forEach(function(r) {
        // 品名＋仕様＋数量＋単価 をキーにして重複除去
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

// ===== Drive検索 API =====

// ===== 機種情報管理 API =====

function _apiModelInfoGet(p) {
  try {
    var all = getAllModelInfoData().map(_modelInfoRowToObject);

    // 基板IDで検索
    if (p.boardId) {
      var found = all.find(function(r) { return r.boardId === String(p.boardId).trim(); });
      return { success: true, item: found || null };
    }
    // 機種コードで検索（複数の基板が返る）
    if (p.modelCode) {
      var items = all.filter(function(r) { return r.modelCode === String(p.modelCode).trim(); });
      return { success: true, items: items };
    }
    // 全件返す
    return { success: true, items: all };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _apiModelInfoSave(p) {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_MODEL_INFO);
    if (!sheet) return { success: false, error: '基板情報管理シートが見つかりません。initialSetupを実行してください。' };

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
      boardId,
      p.modelCode      || '',
      p.quoteUrl       || '',
      p.orderUrl       || '',
      p.purchaseUrl1   || '',
      p.purchaseUrl2   || '',
      p.purchaseUrl3   || '',
      p.localServerUrl || '',
      p.comment        || '',
      nowJST(),
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
  // PDFをDriveにアップロードしてURLを返す
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

var DRIVE_SEARCH_FOLDER_ID = '1oAkPV-O4FZezbGmv7sTlNA4wFljMOyv6';

/**
 * Drive検索 - キャッシュ方式（高速）
 * 全件インデックスをスクリプトプロパティにキャッシュ
 * 検索はキャッシュから行うため1秒以内に完了
 */
var DRIVE_CACHE_KEY     = 'DRIVE_PDF_CACHE';
var DRIVE_CACHE_TS_KEY  = 'DRIVE_PDF_CACHE_TS';
var DRIVE_CACHE_TTL_MS  = 60 * 60 * 1000; // 1時間

function _apiDriveSearch(p) {
  try {
    var keyword  = String(p.keyword  || '').trim().toLowerCase();
    var dateFrom = String(p.dateFrom || '').trim();
    var dateTo   = String(p.dateTo   || '').trim();
    var forceRefresh = p.forceRefresh === true;

    // キャッシュから全件リストを取得
    var allFiles = _getDrivePdfCache(forceRefresh);

    // フィルタリング
    var filtered = allFiles.filter(function(f) {
      if (keyword && f.name.toLowerCase().indexOf(keyword) < 0) return false;
      if (dateFrom && f.updatedAt < dateFrom) return false;
      if (dateTo   && f.updatedAt > dateTo + 'z') return false;
      return true;
    });

    // 更新日降順
    filtered.sort(function(a, b) {
      return String(b.updatedAt).localeCompare(String(a.updatedAt));
    });

    Logger.log('[DRIVE SEARCH] キャッシュ件数:' + allFiles.length + ' フィルタ後:' + filtered.length);
    return {
      success:     true,
      total:       filtered.length,
      items:       filtered.slice(0, 200),
      cacheTotal:  allFiles.length,
      cached:      true,
    };
  } catch(e) {
    Logger.log('[DRIVE SEARCH ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * キャッシュを取得（TTL切れまたは強制更新の場合は再構築）
 */
function _getDrivePdfCache(forceRefresh) {
  var props = PropertiesService.getScriptProperties();

  if (!forceRefresh) {
    var ts  = props.getProperty(DRIVE_CACHE_TS_KEY);
    var now = new Date().getTime();
    if (ts && (now - Number(ts)) < DRIVE_CACHE_TTL_MS) {
      try {
        var cached = props.getProperty(DRIVE_CACHE_KEY);
        if (cached) {
          var parsed = JSON.parse(cached);
          Logger.log('[CACHE HIT] ' + parsed.length + '件');
          return parsed;
        }
      } catch(e) { /* キャッシュ破損時は再構築 */ }
    }
  }

  Logger.log('[CACHE MISS] インデックス再構築開始...');
  var files = _buildDrivePdfIndex();
  
  // スクリプトプロパティは500KBまでなのでURLなしで保存
  var slim = files.map(function(f) {
    return { id:f.id, name:f.name, url:f.url, updatedAt:f.updatedAt, createdAt:f.createdAt, size:f.size };
  });
  
  try {
    props.setProperty(DRIVE_CACHE_KEY, JSON.stringify(slim));
    props.setProperty(DRIVE_CACHE_TS_KEY, String(new Date().getTime()));
    Logger.log('[CACHE SET] ' + slim.length + '件を保存');
  } catch(e) {
    // 容量オーバーの場合は先頭2000件のみ保存
    var trimmed = slim.slice(0, 2000);
    props.setProperty(DRIVE_CACHE_KEY, JSON.stringify(trimmed));
    props.setProperty(DRIVE_CACHE_TS_KEY, String(new Date().getTime()));
    Logger.log('[CACHE SET] 容量制限により' + trimmed.length + '件を保存');
  }
  return slim;
}

/**
 * Drive PDF インデックスを全件構築
 * バックグラウンドで実行（初回のみ時間がかかる）
 */
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
        var params = {
          q: q, pageSize: 200,
          fields: 'nextPageToken, files(id, name, webViewLink, size, modifiedTime, createdTime)',
          orderBy: 'modifiedTime desc',
        };
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
    } catch(fe) {
      Logger.log('[INDEX BUILD] スキップ: ' + fe.message);
    }
  }
  Logger.log('[INDEX BUILD] 完了: ' + results.length + '件');
  return results;
}

/**
 * キャッシュを手動更新するトリガー関数（1時間ごとに自動実行も可）
 */
function refreshDrivePdfCache() {
  Logger.log('[CACHE REFRESH] 開始');
  var files = _buildDrivePdfIndex();
  var slim  = files.map(function(f) {
    return { id:f.id, name:f.name, url:f.url, updatedAt:f.updatedAt, createdAt:f.createdAt, size:f.size };
  });
  var props = PropertiesService.getScriptProperties();
  try {
    props.setProperty(DRIVE_CACHE_KEY, JSON.stringify(slim));
  } catch(e) {
    props.setProperty(DRIVE_CACHE_KEY, JSON.stringify(slim.slice(0, 2000)));
  }
  props.setProperty(DRIVE_CACHE_TS_KEY, String(new Date().getTime()));
  Logger.log('[CACHE REFRESH] 完了: ' + slim.length + '件');
  return slim.length;
}

function _apiDriveSearchFallback(p) {
  try {
    var keyword  = String(p.keyword  || '').trim();
    var dateFrom = String(p.dateFrom || '').trim();
    var dateTo   = String(p.dateTo   || '').trim();
    var folder   = DriveApp.getFolderById(DRIVE_SEARCH_FOLDER_ID);
    var results  = [];
    _searchFilesInFolder(folder, keyword, dateFrom, dateTo, results, 0, '');
    results.sort(function(a, b) { return new Date(b.updatedAtRaw) - new Date(a.updatedAtRaw); });
    return { success: true, total: results.length, items: results.slice(0, 200) };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _searchFilesInFolder(folder, keyword, dateFrom, dateTo, results, depth, folderPath) {
  if (depth > 6) return;
  var currentPath = folderPath ? folderPath + ' / ' + folder.getName() : folder.getName();
  var files = folder.getFilesByType(MimeType.PDF);
  while (files.hasNext()) {
    var file    = files.next();
    var name    = file.getName();
    var updated = file.getLastUpdated();
    var created = file.getDateCreated();
    if (keyword && name.toLowerCase().indexOf(keyword.toLowerCase()) < 0) continue;
    if (dateFrom) { var from = new Date(dateFrom.replace(/\//g,'-')); if (updated < from && created < from) continue; }
    if (dateTo)   { var to   = new Date(dateTo.replace(/\//g,'-')+'T23:59:59'); if (updated > to && created > to) continue; }
    results.push({
      id: file.getId(), name: name, url: file.getUrl(),
      size: Math.round(file.getSize()/1024),
      updatedAt:    Utilities.formatDate(updated,'Asia/Tokyo','yyyy/MM/dd HH:mm'),
      createdAt:    Utilities.formatDate(created,'Asia/Tokyo','yyyy/MM/dd HH:mm'),
      updatedAtRaw: updated.getTime(), folderPath: currentPath,
    });
  }
  var subs = folder.getFolders();
  while (subs.hasNext()) { _searchFilesInFolder(subs.next(), keyword, dateFrom, dateTo, results, depth+1, currentPath); }
}

/**
 * Drive APIで指定フォルダ配下の全サブフォルダIDを高速取得
 * 再帰なし・APIクエリで1〜2回のAPI呼び出しで完結
 */
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

    // 現在のバッチのフォルダを一括クエリ（OR条件で複数フォルダを一度に検索）
    var orParts = currentBatch.map(function(id) { return "'" + id + "' in parents"; });
    // Drive APIのOR条件は "A or B" 形式（括弧なし）
    var batchSize = 10; // OR条件が多すぎるとエラーになるため分割
    for (var i = 0; i < orParts.length; i += batchSize) {
      var batch = orParts.slice(i, i + batchSize);
      var q = '(' + batch.join(' or ') + ') and mimeType = "application/vnd.google-apps.folder" and trashed = false';
      try {
        var resp = Drive.Files.list({
          q: q,
          fields: 'files(id, name)',
          pageSize: 200,
        });
        (resp.files || []).forEach(function(f) {
          if (!visited[f.id]) {
            visited[f.id] = true;
            allIds.push(f.id);
            queue.push(f.id);
          }
        });
      } catch(e) {
        Logger.log('[GET SUBFOLDERS] バッチエラー: ' + e.message);
      }
    }
  }

  Logger.log('[GET SUBFOLDERS] 取得フォルダ数: ' + allIds.length);
  return allIds;
}

function _getFolderList(folder, depth) {
  if (depth > 3) return [];
  var list = [{ name: folder.getName(), id: folder.getId() }];
  var subs = folder.getFolders();
  while (subs.hasNext()) { list = list.concat(_getFolderList(subs.next(), depth+1)); }
  return list;
}

function _apiDriveDeleteFile(p) {
  try {
    if (!p.fileId) return { success: false, error: 'fileIdが必要です' };
    // Driveからファイルを削除（ゴミ箱へ）
    DriveApp.getFileById(p.fileId).setTrashed(true);
    // キャッシュからも削除
    var props  = PropertiesService.getScriptProperties();
    var cached = props.getProperty(DRIVE_CACHE_KEY);
    if (cached) {
      var files   = JSON.parse(cached);
      var updated = files.filter(function(f) { return f.id !== p.fileId; });
      props.setProperty(DRIVE_CACHE_KEY, JSON.stringify(updated));
    }
    Logger.log('[DRIVE DELETE] ' + p.fileId);
    return { success: true };
  } catch(e) {
    Logger.log('[DRIVE DELETE ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ===== 見積書詳細 API =====

function _apiGetQuoteDetail(p) {
  if (!p.mgmtId) return { success: false, error: '管理IDが必要' };
  try {
    var ss        = getSpreadsheet();
    var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var mgmtRow   = _getMgmtRowById(mgmtSheet, p.mgmtId);
    if (!mgmtRow) return { success: false, error: '管理ID未発見' };

    // 同じ見積番号を持つ全管理IDを収集
    var quoteNo   = String(mgmtRow[MGMT_COLS.QUOTE_NO - 1] || '').trim();
    var allMgmt   = getAllMgmtData();
    var relatedIds = [String(p.mgmtId)];
    if (quoteNo) {
      allMgmt.forEach(function(r) {
        var id  = String(r[MGMT_COLS.ID - 1] || '');
        var qNo = String(r[MGMT_COLS.QUOTE_NO - 1] || '').trim();
        if (id !== p.mgmtId && qNo === quoteNo) relatedIds.push(id);
      });
    }

    // 見積書明細を全関連IDから取得
    var qs    = ss.getSheetByName(CONFIG.SHEET_QUOTES);
    var qLast = qs.getLastRow();
    var lines = [];
    if (qLast > 1) {
      lines = qs.getRange(2, 1, qLast - 1, 15).getValues()
        .filter(function(r) { return relatedIds.indexOf(String(r[0])) >= 0; })
        .map(function(r) {
          return {
            lineNo:    r[5], itemName:  r[6], spec:      r[7],
            qty:       r[8], unit:      r[9], unitPrice: r[10],
            amount:    r[11], remarks:  r[12], pdfUrl:   r[13],
            issueDate: _toDateStr(r[2]), destCompany: r[3], destPerson: r[4],
          };
        });
    }

    var obj = _rowToObject(mgmtRow);
    return {
      success:    true,
      mgmt:       obj,
      quoteLines: lines,
    };
  } catch(e) {
    Logger.log('[QUOTE DETAIL ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
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
  if (!p.orderMgmtId || !p.quoteMgmtId)
    return { success: false, error: '管理IDが必要です' };
  return confirmManualLink(p.orderMgmtId, p.quoteMgmtId);
}

function _apiRunMatching(p) {
  var mgmtId = p && p.mgmtId;
  if (mgmtId) {
    var result = matchOrderToQuote(mgmtId);
    return { success: true, result: result };
  }
  // 全件バッチ
  var result = batchMatchAllUnlinked();
  return { success: true, result: result };
}

// ===== 見積書一覧 API =====

function _apiQuoteListGetAll() {
  try {
    var ss         = getSpreadsheet();
    var mgmtSheet  = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var quoteSheet = ss.getSheetByName(CONFIG.SHEET_QUOTES);

    // 管理シートから見積番号がある行を取得（重複除去済み）
    var mgmtData  = getAllMgmtData();
    var allRows   = mgmtData.filter(function(r) {
      return String(r[MGMT_COLS.QUOTE_NO - 1]).trim() !== '';
    });
    // 見積番号で重複除去
    var seenQNo  = {};
    var quoteRows = allRows.filter(function(r) {
      var qNo = String(r[MGMT_COLS.QUOTE_NO - 1]).trim();
      if (seenQNo[qNo]) return false;
      seenQNo[qNo] = true;
      return true;
    });

    // 見積書シートの先頭行（発行日・送り先・担当者）を取得するためのマップ
    var quoteLineMap = {};
    if (quoteSheet && quoteSheet.getLastRow() > 1) {
      var qData = quoteSheet.getRange(2, 1, quoteSheet.getLastRow() - 1, 15).getValues();
      qData.forEach(function(r) {
        var mgmtId = String(r[QUOTE_COLS.MGMT_ID - 1] || '');
        if (!mgmtId || quoteLineMap[mgmtId]) return; // 先頭行のみ取得
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

    // 発行日降順でソート（新しい順）
    items.sort(function(a, b) {
      var da = String(a.issueDate || a.quoteDate || '');
      var db = String(b.issueDate || b.quoteDate || '');
      return db.localeCompare(da);
    });

    return { success: true, total: items.length, items: items };
  } catch(e) {
    Logger.log('[QUOTE LIST ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// linked値の正規化（TRUE文字列・真偽値どちらにも対応）
function _isLinkedVal(val) {
  return val === true || val === 'TRUE' || val === 'true';
}

// ===== 見積台帳 API =====

function _apiLedgerGetAll() {
  var rows = getAllLedgerData();
  return { success: true, items: rows.map(_ledgerRowToObject) };
}

function _apiLedgerSave(p) {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_LEDGER);
    if (!sheet) return { success: false, error: '見積台帳シートが存在しません。initialSetupを実行してください。' };

    var isNew    = !p.ledgerId;
    var ledgerId = isNew ? generateLedgerId() : p.ledgerId;

    if (isNew) {
      sheet.appendRow([
        ledgerId,
        p.quoteNo     || '',
        p.issueDate   || '',
        p.dest        || '',
        p.category    || '',
        p.subject     || '',
        p.status      || LEDGER_STATUS.PENDING,
        p.saveUrl     || '',
        p.machineCode || '',
        p.boardName   || '',
        p.modelNo     || '',   // ★型番
        p.amount !== undefined && p.amount !== '' ? Number(p.amount) : '',  // ★金額
        p.submitTo    || '',   // ★提出先担当者
        p.remarks     || '',   // ★備考
      ]);
    } else {
      var last = sheet.getLastRow();
      if (last <= 1) return { success: false, error: '対象行が見つかりません' };
      var ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
      var idx = ids.indexOf(String(ledgerId));
      if (idx < 0) return { success: false, error: '台帳IDが見つかりません: ' + ledgerId };
      var row = idx + 2;
      var fields = {
        quoteNo:     LEDGER_COLS.QUOTE_NO,
        issueDate:   LEDGER_COLS.ISSUE_DATE,
        dest:        LEDGER_COLS.DEST,
        category:    LEDGER_COLS.CATEGORY,
        subject:     LEDGER_COLS.SUBJECT,
        status:      LEDGER_COLS.STATUS,
        saveUrl:     LEDGER_COLS.SAVE_URL,
        machineCode: LEDGER_COLS.MACHINE_CODE,
        boardName:   LEDGER_COLS.BOARD_NAME,
        modelNo:     LEDGER_COLS.MODEL_NO,
        amount:      LEDGER_COLS.AMOUNT,
        submitTo:    LEDGER_COLS.SUBMIT_TO,
        remarks:     LEDGER_COLS.REMARKS,
      };
      Object.keys(fields).forEach(function(key) {
        if (p[key] !== undefined) {
          sheet.getRange(row, fields[key]).setValue(p[key]);
        }
      });
    }
    return { success: true, ledgerId: ledgerId };
  } catch(e) {
    Logger.log('[LEDGER SAVE ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}


function _apiLedgerDelete(p) {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_LEDGER);
    if (!sheet || !p.ledgerId) return { success: false, error: 'パラメータ不足' };
    var last = sheet.getLastRow();
    if (last <= 1) return { success: false, error: '対象なし' };
    var ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
    var idx = ids.indexOf(String(p.ledgerId));
    if (idx < 0) return { success: false, error: '対象行が見つかりません' };
    sheet.deleteRow(idx + 2);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

/**
 * メール送信検知後に保存先URLとステータスを自動更新
 * 件名・宛先名でファジーマッチングして対象行を特定
 */
function _apiLedgerUpdateUrl(p) {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_LEDGER);
    if (!sheet) return { success: false, error: 'シートなし' };
    var last = sheet.getLastRow();
    if (last <= 1) return { success: true, matched: false };

    var data = sheet.getRange(2, 1, last - 1, 10).getValues();
    var bestIdx = -1;
    var bestScore = 0;

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

    if (bestIdx < 0 || bestScore < 40) {
      Logger.log('[LEDGER UPDATE] マッチなし subject=' + p.subject + ' dest=' + p.dest);
      return { success: true, matched: false };
    }

    var targetRow = bestIdx + 2;
    sheet.getRange(targetRow, LEDGER_COLS.SAVE_URL).setValue(p.saveUrl || '');
    sheet.getRange(targetRow, LEDGER_COLS.STATUS).setValue(LEDGER_STATUS.SENT);
    if (p.issueDate) sheet.getRange(targetRow, LEDGER_COLS.ISSUE_DATE).setValue(p.issueDate);
    Logger.log('[LEDGER UPDATE] 自動更新 行=' + targetRow + ' スコア=' + bestScore);
    return { success: true, matched: true, row: targetRow, score: bestScore };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _normStr(s) {
  return String(s).replace(/\s+/g,'').replace(/　/g,'').toLowerCase()
    .replace(/株式会社|有限会社|（株）|\(株\)/g,'');
}
function _strIncludes(a, b) {
  return a.indexOf(b) >= 0 || b.indexOf(a) >= 0;
}


// ===== 見積台帳 ファイルアップロードAPI =====

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
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _apiLedgerGetMachines() {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_LEDGER);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, machines: [] };
    var col = sheet.getRange(2, LEDGER_COLS.MACHINE_CODE, sheet.getLastRow() - 1, 1).getValues().flat();
    var seen = {}, machines = [];
    col.forEach(function(v) {
      var m = String(v||'').trim();
      if (m && !seen[m]) { seen[m] = true; machines.push(m); }
    });
    return { success: true, machines: machines.sort() };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// 機種のみ新規作成（見積なしのプレースホルダー行）
function _apiLedgerCreateMachine(p) {
  try {
    var machineCode = String(p.machineCode || '').trim();
    if (!machineCode) return { success: false, error: '機種コードは必須です' };

    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_LEDGER);
    if (!sheet) return { success: false, error: '見積台帳シートが存在しません' };

    // 既存チェック：同機種コード・件名が「（機種フォルダ）」の行があれば重複
    if (sheet.getLastRow() > 1) {
      var codes = sheet.getRange(2, LEDGER_COLS.MACHINE_CODE, sheet.getLastRow() - 1, 2).getValues();
      for (var i = 0; i < codes.length; i++) {
        if (String(codes[i][0]).trim() === machineCode && String(codes[i][1]).trim() === '（機種フォルダ）') {
          return { success: false, error: '機種「' + machineCode + '」はすでに登録されています' };
        }
      }
    }

    var ledgerId = generateLedgerId();
    // プレースホルダー行：件名を「（機種フォルダ）」として識別
    sheet.appendRow([
      ledgerId,
      '',           // 見積No.
      '',           // 発行日
      '',           // 宛先
      '',           // 分類
      '（機種フォルダ）',  // 件名 ← プレースホルダー識別子
      '__MACHINE_FOLDER__',  // ステータス ← 特殊値
      '',           // 保存先URL
      machineCode,  // 機種コード ★
      '',           // 基板名
      '',           // 型番
      '',           // 金額
      '',           // 提出先
      p.remarks || '',  // 備考
    ]);
    return { success: true, ledgerId: ledgerId, machineCode: machineCode };
  } catch(e) {
    Logger.log('[LEDGER CREATE MACHINE ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}



// ===== Todo API =====

function _apiGetTodos() {
  var items = getAllTodoData().map(function(r) {
    return { id:r[0], title:r[1], client:r[2], dueDate:r[3],
             priority:r[4], status:r[5], linkedMgmt:r[6], memo:r[7] };
  });
  return { success: true, items: items };
}

function _apiSaveTodo(p) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_TODO);
  if (!sheet) return { success: false, error: 'Todoシートなし' };

  if (p.id) {
    // 更新
    var last = sheet.getLastRow();
    if (last > 1) {
      var ids = sheet.getRange(2,1,last-1,1).getValues().flat();
      var idx = ids.findIndex(function(v) { return String(v) === String(p.id); });
      if (idx >= 0) {
        sheet.getRange(idx+2, 1, 1, 8).setValues([[
          p.id, p.title||'', p.client||'', p.dueDate||'',
          p.priority||'中', p.status||'未着手', p.linkedMgmt||'', p.memo||''
        ]]);
        return { success: true, id: p.id };
      }
    }
  }
  // 新規
  var newId = generateTodoId();
  sheet.appendRow([newId, p.title||'', p.client||'', p.dueDate||'',
    p.priority||'中', p.status||'未着手', p.linkedMgmt||'', p.memo||'']);
  return { success: true, id: newId };
}

function _apiDeleteTodo(p) {
  if (!p.id) return { success: false, error: 'IDが必要' };
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_TODO);
  var last  = sheet.getLastRow();
  if (last <= 1) return { success: false, error: 'データなし' };

  var ids = sheet.getRange(2,1,last-1,1).getValues().flat();
  var idx = ids.findIndex(function(v) { return String(v) === String(p.id); });
  if (idx < 0) return { success: false, error: '未発見' };
  sheet.deleteRow(idx + 2);
  return { success: true };
}

// ===== カレンダーAPI =====

function _apiGetCalendar(p) {
  var year  = p.year  || new Date().getFullYear();
  var month = p.month || (new Date().getMonth() + 1);
  var ym    = year + '/' + String(month).padStart(2,'0');

  var all   = getAllMgmtData().map(_rowToObject);
  var todos = getAllTodoData().map(function(r) {
    return { id:r[0], title:r[1], client:r[2], dueDate:r[3],
             priority:r[4], status:r[5], linkedMgmt:r[6], type:'todo' };
  });

  // 当月に関係するイベントを収集
  var events = [];

  all.forEach(function(item) {
    // 発注日
    if (item.orderDate && String(item.orderDate).indexOf(ym) === 0) {
      events.push({ date: item.orderDate, label: item.client || item.orderNo,
                    type: 'order', status: item.status, mgmtId: item.id });
    }
    // 納期
    if (item.deliveryDate && String(item.deliveryDate).indexOf(ym) === 0) {
      events.push({ date: item.deliveryDate, label: '納期: ' + (item.client || item.orderNo),
                    type: 'delivery', status: item.status, mgmtId: item.id });
    }
    // 見積日
    if (item.quoteDate && String(item.quoteDate).indexOf(ym) === 0) {
      events.push({ date: item.quoteDate, label: '見積: ' + (item.client || item.quoteNo),
                    type: 'quote', status: item.status, mgmtId: item.id });
    }
  });

  // Todo期限
  todos.forEach(function(t) {
    if (t.dueDate && String(t.dueDate).indexOf(ym) === 0) {
      events.push({ date: t.dueDate, label: '📝 ' + t.title,
                    type: 'todo', status: t.status, todoId: t.id });
    }
  });

  return { success: true, year: year, month: month, events: events };
}

// ===== 行→オブジェクト変換 =====

/**
 * スプレッドシートのDate型を日本時間の文字列に変換
 * 例: Date オブジェクト → "2026/01/09"
 */
function _toDateStr(val) {
  if (!val || val === '') return '';
  if (val instanceof Date) {
    // 有効な日付かチェック
    if (isNaN(val.getTime())) return '';
    return Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy/MM/dd');
  }
  // 文字列の場合はそのまま返す（ただしISO形式を変換）
  var s = String(val).trim();
  if (s === '') return '';
  // ISO 8601形式 (2026-01-08T15:00:00.000Z) を変換
  if (s.indexOf('T') > 0 && s.indexOf('Z') > 0) {
    try {
      var d = new Date(s);
      if (!isNaN(d.getTime())) {
        return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
      }
    } catch(e) {}
  }
  return s;
}

/**
 * 数値・空文字の安全な変換
 */
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
    deliveryDate:   _toDateStr(row[MGMT_COLS.DELIVERY_DATE - 1]),
    memo:           String(row[MGMT_COLS.MEMO - 1] || ''),
    createdAt:      _toDateStr(row[MGMT_COLS.CREATED_AT - 1]),
    updatedAt:      _toDateStr(row[MGMT_COLS.UPDATED_AT - 1]),
  };
}

// ===== デバッグ用（確認後削除可） =====
function debugGetAll() {
  try {
    var ss = getSpreadsheet();
    Logger.log('SpreadsheetID: ' + ss.getId());
    var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    Logger.log('シート存在: ' + (sheet ? 'YES' : 'NO'));
    Logger.log('最終行: ' + (sheet ? sheet.getLastRow() : 'N/A'));
    Logger.log('最終列: ' + (sheet ? sheet.getLastColumn() : 'N/A'));
    var result = _apiGetAll();
    Logger.log('success: ' + result.success);
    Logger.log('total: ' + result.total);
    Logger.log('error: ' + (result.error || 'なし'));
    if (result.items && result.items.length > 0) {
      Logger.log('1件目: ' + JSON.stringify(result.items[0]));
    }
  } catch(e) {
    Logger.log('ERROR: ' + e.message + '\n' + e.stack);
  }
}

// ===== Drive構造確認用テスト関数 =====
function testDriveStructure() {
  var root = DriveApp.getFolderById(DRIVE_SEARCH_FOLDER_ID);
  Logger.log('ルート: ' + root.getName());
  _logFolderStructure(root, 0, '');
}

function _logFolderStructure(folder, depth, path) {
  if (depth > 5) return;
  var indent = '';
  for (var i = 0; i < depth; i++) indent += '  ';
  var currentPath = path ? path + '/' + folder.getName() : folder.getName();

  // 全ファイル数カウント
  var allFiles = folder.getFiles();
  var fileCount = 0;
  var pdfCount  = 0;
  var sampleName = '';
  while (allFiles.hasNext()) {
    var f = allFiles.next();
    fileCount++;
    if (f.getMimeType() === MimeType.PDF) {
      pdfCount++;
      if (!sampleName) sampleName = f.getName();
    }
  }
  Logger.log(indent + '📁 ' + folder.getName() + ' (全' + fileCount + '件, PDF' + pdfCount + '件)' + (sampleName ? ' 例:' + sampleName : ''));

  var subs = folder.getFolders();
  while (subs.hasNext()) {
    _logFolderStructure(subs.next(), depth + 1, currentPath);
  }
}

function testDriveSearchDirect() {
  var result = _apiDriveSearch({ keyword: '', dateFrom: '', dateTo: '' });
  Logger.log('success: ' + result.success);
  Logger.log('total: '   + result.total);
  Logger.log('error: '   + (result.error || 'なし'));
  if (result.items && result.items[0]) {
    Logger.log('1件目: ' + JSON.stringify(result.items[0]));
  }
}
