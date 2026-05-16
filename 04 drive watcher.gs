// ============================================================
// 見積書・注文書管理システム
// ファイル 4/4: Driveフォルダ自動監視
// ============================================================
//
// 【動作の流れ】
// 5分ごとにインポート用フォルダを監視
//   ↓ 未処理のPDFを検知
//   ↓ Gemini OCRで解析・スプレッドシートに転記
//   ↓ 保存先フォルダへ移動
//   ↓ 処理済みファイルIDをスクリプトプロパティに記録（重複防止）
//
// インポート用フォルダ:
//   見積書        → 1Y66PDSi35ScuIyS0Jgm0l3p2l7MEM2Jk
//   注文書（試作） → 1Ufq4xMjOmZvUQLC_Zp0EAWlHF0mYAGDM
//   注文書（量産） → 1ujzCtYzOqU9_a6tiEXOHhDWRv15a0p0k
//
// 保存先フォルダ:
//   見積書        → 1sB42xntGKL31GeT9OjOKTxVJwj9IQz-h
//   注文書（試作） → 1wVeYlt-9GsortfOsUggBsWta8GtXIRvS
//   注文書（量産） → 1ASyV7PmhYQVH-72rVD3evToYJWxGhMbA
// ============================================================

/**
 * メイントリガー関数（5分ごとに自動実行）
 */
function processDriveImports() {
  try {
    Logger.log('[DRIVE WATCH] 監視開始 ' + nowJST());
    try { _updateTriggerHealth(); } catch(e) {}

    var processed = 0;
    processed += _watchFolder(CONFIG.IMPORT_QUOTE_FOLDER_ID, CONFIG.QUOTE_FOLDER_ID, 'quote', '');
    processed += _watchFolder(CONFIG.IMPORT_ORDER_TRIAL_FOLDER_ID, CONFIG.ORDER_TRIAL_FOLDER_ID, 'order', CONFIG.ORDER_TYPE.TRIAL);
    processed += _watchFolder(CONFIG.IMPORT_ORDER_MASS_FOLDER_ID, CONFIG.ORDER_MASS_FOLDER_ID, 'order', CONFIG.ORDER_TYPE.MASS);

    Logger.log('[DRIVE WATCH] 完了。処理ファイル数: ' + processed);
    return processed;
  } catch(e) {
    Logger.log('[DRIVE WATCH ERROR] ' + e.message);
    return 0;
  }
}

/**
 * 指定フォルダを監視し、未処理PDFを処理して保存先へ移動する
 */
function _watchFolder(importFolderId, saveFolderId, docType, orderType) {
  var processedIds = _getProcessedFileIds();
  var count = 0;
  try {
    var folder = DriveApp.getFolderById(importFolderId);
    var files  = folder.getFilesByType(MimeType.PDF);
    while (files.hasNext()) {
      var file   = files.next();
      var fileId = file.getId();
      if (processedIds[fileId]) { Logger.log('[DRIVE WATCH] スキップ（処理済み）: ' + file.getName()); continue; }
      Logger.log('[DRIVE WATCH] 処理開始: ' + file.getName());
      try {
        // OCR解析（17_ocr_hybrid.gs の extractPdfData が優先使用される）
        var ocr = extractPdfData(file, docType);
        if (!ocr) {
          Logger.log('[DRIVE WATCH] OCR失敗: ' + file.getName());
          ocr = { documentNo: file.getName().replace(/\.pdf$/i,''), issueDate: '', documentDate: '',
                  destCompany: '', clientName: '', subject: file.getName(),
                  subtotal: 0, tax: 0, totalAmount: 0, lineItems: [], actionType: 'new' };
          try { _logOcrResult(file.getName(), 'ocr_failed', null, 'Driveインポート：OCR失敗'); } catch(e) {}
        }

        // 保存先フォルダへコピー
        var clientName    = ocr.destCompany || ocr.clientName || '未分類';
        var finalFolderId = _getOrCreateSubFolder(saveFolderId, clientName);
        var saveFolder    = DriveApp.getFolderById(finalFolderId);
        var newName       = nowJST().replace(/[\/: ]/g,'') + '_' + file.getName();
        var savedFile     = file.makeCopy(newName, saveFolder);
        var pdfUrl        = savedFile.getUrl();
        var folderUrl     = getFolderUrl(finalFolderId);
        Logger.log('[DRIVE WATCH] 保存先へコピー完了: ' + newName);

        // スプレッドシートへ転記
        var mockMsgId = 'DRIVE_' + fileId;
        var finalMgmtId;
        if (docType === 'quote') {
          finalMgmtId = _processQuotePdfFromFile(pdfUrl, folderUrl, ocr, mockMsgId);
          // 見積台帳へも自動登録
          try {
            _apiLedgerUpdateUrl({
              quoteNo: ocr.documentNo || '', dest: ocr.destCompany || ocr.clientName || '',
              subject: ocr.subject || file.getName(),
              issueDate: ocr.issueDate || ocr.documentDate || nowJST().substring(0,10),
              saveUrl: pdfUrl, sentDate: nowJST().substring(0,10), status: '送信済み',
            });
          } catch(le) { Logger.log('[DRIVE WATCH] 台帳登録エラー: ' + le.message); }
        } else {
          var finalType = orderType || ocr.orderType || '';
          finalMgmtId = _saveOrderData(ocr, finalType, pdfUrl, folderUrl, mockMsgId, file.getName());
          try {
            if (finalMgmtId) {
              var lr = aiLinkOrderToQuote(finalMgmtId);
              _sendOrderRegistrationToChat(finalMgmtId, { documentNo: ocr.documentNo || file.getName(), orderType: finalType }, lr);
            }
          } catch(le) { Logger.log('[DRIVE WATCH LINK ERROR] ' + le.message); }
        }

        _markFileAsProcessed(fileId);
        _moveToProcessedSubfolder(file, folder);
        count++;
        Logger.log('[DRIVE WATCH] 転記完了: ' + file.getName() + ' mgmtId=' + (finalMgmtId||'?'));
        Utilities.sleep(2000);

      } catch(fileErr) {
        Logger.log('[DRIVE WATCH FILE ERROR] ' + file.getName() + ': ' + fileErr.message);
        _markFileAsProcessed(fileId);
        try { _logOcrResult(file.getName(), 'error', null, fileErr.message); } catch(e) {}
      }
    }
  } catch(folderErr) {
    Logger.log('[DRIVE WATCH FOLDER ERROR] FolderID=' + importFolderId + ': ' + folderErr.message);
  }
  return count;
}

/**
 * 処理済みファイルをインポートフォルダ内の「処理済み」サブフォルダへ移動
 * サブフォルダが存在しない場合は自動作成
 */
function _moveToProcessedSubfolder(file, parentFolder) {
  try {
    var subFolderName = '処理済み';
    var subFolders = parentFolder.getFoldersByName(subFolderName);
    var subFolder;

    if (subFolders.hasNext()) {
      subFolder = subFolders.next();
    } else {
      subFolder = parentFolder.createFolder(subFolderName);
      Logger.log('[DRIVE WATCH] 「処理済み」サブフォルダを作成: ' + parentFolder.getName());
    }

    // サブフォルダへ移動
    subFolder.addFile(file);
    parentFolder.removeFile(file);
    Logger.log('[DRIVE WATCH] 「処理済み」フォルダへ移動: ' + file.getName());
  } catch(e) {
    Logger.log('[DRIVE WATCH MOVE ERROR] ' + e.message);
    // 移動失敗でも処理は続行（スプレッドシートへの転記は完了済み）
  }
}


// ============================================================
// 処理済みファイルID管理（スクリプトプロパティを使用）
// ============================================================

var PROCESSED_IDS_KEY = 'DRIVE_PROCESSED_FILE_IDS';
var MAX_IDS_STORED    = 500; // 保持する最大ID数（古いものから削除）

/**
 * 処理済みファイルIDの辞書を取得
 * @returns {object} {fileId: timestamp} の辞書
 */
function _getProcessedFileIds() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(PROCESSED_IDS_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch(e) {
    return {};
  }
}

/**
 * ファイルIDを処理済みとして記録
 */
function _markFileAsProcessed(fileId) {
  try {
    var ids = _getProcessedFileIds();
    ids[fileId] = nowJST();

    // 上限を超えたら古いものを削除
    var keys = Object.keys(ids);
    if (keys.length > MAX_IDS_STORED) {
      var sorted = keys.sort(function(a,b) {
        return ids[a] < ids[b] ? -1 : 1;
      });
      // 古い50件を削除
      sorted.slice(0, 50).forEach(function(k) { delete ids[k]; });
    }

    PropertiesService.getScriptProperties().setProperty(
      PROCESSED_IDS_KEY, JSON.stringify(ids)
    );
  } catch(e) {
    Logger.log('[MARK PROCESSED ERROR] ' + e.message);
  }
}


// ============================================================
// 手動実行・テスト用
// ============================================================

/**
 * 処理済みIDを全削除（再処理したい場合に手動実行）
 */
function clearProcessedFileIds() {
  PropertiesService.getScriptProperties().deleteProperty(PROCESSED_IDS_KEY);
  Logger.log('処理済みファイルIDをリセットしました。次回の監視で全ファイルが再処理されます。');
}

/**
 * 各インポートフォルダの状況を確認（手動実行でログ確認）
 */
function checkImportFolders() {
  var folders = [
    { label: '見積書インポート',        id: CONFIG.IMPORT_QUOTE_FOLDER_ID },
    { label: '注文書（試作）インポート', id: CONFIG.IMPORT_ORDER_TRIAL_FOLDER_ID },
    { label: '注文書（量産）インポート', id: CONFIG.IMPORT_ORDER_MASS_FOLDER_ID },
  ];
  var processedIds = _getProcessedFileIds();

  folders.forEach(function(f) {
    try {
      var folder = DriveApp.getFolderById(f.id);
      var files  = folder.getFilesByType(MimeType.PDF);
      var pending = [], done = [];

      while (files.hasNext()) {
        var file = files.next();
        if (processedIds[file.getId()]) {
          done.push(file.getName());
        } else {
          pending.push(file.getName());
        }
      }

      Logger.log('\n===== ' + f.label + ' =====');
      Logger.log('未処理: ' + pending.length + '件');
      pending.forEach(function(n) { Logger.log('  → ' + n); });
      Logger.log('処理済み（フォルダ残り）: ' + done.length + '件');
    } catch(e) {
      Logger.log('[ERROR] ' + f.label + ': ' + e.message);
    }
  });
}

/**
 * 顧客名と年月からサブフォルダを自動生成または取得
 */
function _getOrCreateSubFolder(parentFolderId, clientName) {
  var parent = DriveApp.getFolderById(parentFolderId);
  var now = new Date();
  var monthStr = now.getFullYear() + "-" + ("0" + (now.getMonth() + 1)).slice(-2);
  
  // 1. 月フォルダの確認・作成
  var monthFolders = parent.getFoldersByName(monthStr);
  var monthFolder = monthFolders.hasNext() ? monthFolders.next() : parent.createFolder(monthStr);

  // 2. 顧客フォルダの確認・作成
  var safeClientName = String(clientName).replace(/[\/:*?"<>|]/g, "_");
  var clientFolders = monthFolder.getFoldersByName(safeClientName);
  var clientFolder = clientFolders.hasNext() ? clientFolders.next() : monthFolder.createFolder(safeClientName);
  
  return clientFolder.getId();
}
