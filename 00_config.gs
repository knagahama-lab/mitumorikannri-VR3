// ============================================================
// 00_config.gs
// 設定・初期化・共通ユーティリティ（統合版）
//
// 統合元:
//   - 01 config and setup.gs  （CONFIG・列定義・セットアップ）
//   - 00 missing utils.gs     （補完ユーティリティ）
// ============================================================

// ============================================================
// システム設定
// ============================================================

var CONFIG = {
  SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '',
  GEMINI_API_KEY: '',  // 必ず getGeminiApiKey() を使うこと

  GEMINI_PRIMARY_MODEL:  'gemini-1.5-flash-latest',
  GEMINI_FALLBACK_MODEL: 'gemini-1.5-flash-8b',
  GEMINI_API_ENDPOINT:   'https://generativelanguage.googleapis.com/v1beta/models/',

  WEB_UPLOAD_FOLDER_ID:  '1sB42xntGKL31GeT9OjOKTxVJwj9IQz-h',
  ORDER_TRIAL_FOLDER_ID: '1wVeYlt-9GsortfOsUggBsWta8GtXIRvS',
  ORDER_MASS_FOLDER_ID:  '1ASyV7PmhYQVH-72rVD3evToYJWxGhMbA',
  QUOTE_FOLDER_ID:       '1sB42xntGKL31GeT9OjOKTxVJwj9IQz-h',

  IMPORT_QUOTE_FOLDER_ID:       '1Y66PDSi35ScuIyS0Jgm0l3p2l7MEM2Jk',
  IMPORT_ORDER_TRIAL_FOLDER_ID: '1Ufq4xMjOmZvUQLC_Zp0EAWlHF0mYAGDM',
  IMPORT_ORDER_MASS_FOLDER_ID:  '1ujzCtYzOqU9_a6tiEXOHhDWRv15a0p0k',

  // 処理済み移動先フォルダ（未設定の場合はインポートフォルダ内に「処理済み」サブフォルダを自動作成）
  PROCESSED_ORDER_TRIAL_FOLDER_ID: '1xSZfQulz5zseOtKYNx8QvuS7_R8Q1FJ5',

  SHEET_MANAGEMENT: '管理シート',
  SHEET_QUOTES:     '見積書シート',
  SHEET_ORDERS:     '注文書シート',
  SHEET_EMAIL_CFG:  'メール監視設定',
  SHEET_TODO:       'Todoリスト',
  SHEET_LEDGER:     '見積台帳',
  SHEET_MODEL_INFO:   '基板情報管理',
  SHEET_MODEL_MASTER: '機種マスタ',     // 追加: 機種マスタシート（21_model_master.gs）
  SHEET_QUOTE_SET:   '見積セット管理',  // 追加: 見積セット管理（25_quoteset.gs）

  RATE_LIMIT_WAIT_MS: 20000,
  RATE_LIMIT_RETRIES: 2,

  GOOGLE_CHAT_WEBHOOK_URL: PropertiesService.getScriptProperties().getProperty('GOOGLE_CHAT_WEBHOOK_URL') || '',

  STATUS: {
    PLANNED:   '作成予定',
    SENT:      '送信済み',
    RECEIVED:  '受領',
    ORDERED:   '受注済み',
    DELIVERED: '納品済み',
    CANCELLED: 'キャンセル',
    REVISED:   '受領（差し替え）',
  },
  ORDER_TYPE: { TRIAL: '試作', MASS: '量産' },
};

// ============================================================
// 列定義
// ============================================================

var MGMT_COLS = {
  ID:                1,
  QUOTE_NO:          2,
  ORDER_NO:          3,
  SUBJECT:           4,
  CLIENT:            5,
  STATUS:            6,
  QUOTE_DATE:        7,
  ORDER_DATE:        8,
  QUOTE_AMOUNT:      9,
  ORDER_AMOUNT:      10,
  TAX:               11,
  TOTAL:             12,
  QUOTE_PDF_URL:     13,
  ORDER_PDF_URL:     14,
  DRIVE_FOLDER_URL:  15,
  QUOTE_SHEET_ROW:   16,
  ORDER_SHEET_ROW:   17,
  LINKED:            18,
  ORDER_TYPE:        19,
  MODEL_CODE:        20,
  ORDER_SLIP_NO:     21,
  ASSIGNEE:          22,
  DELIVERY_DATE:     23,
  MEMO:              24,
  CREATED_AT:        25,
  UPDATED_AT:        26,
  GMAIL_MSG_ID:      27,
  IS_LATEST:         28,
  ORDER_DEADLINE:    29,
  REVISION_NO:       30,
  PARENT_MGMT_ID:    31,
  DEADLINE_NOTIFIED: 32,
  BOARD_NAME:        33,
};

var QUOTE_COLS = {
  MGMT_ID:      1,
  QUOTE_NO:     2,
  ISSUE_DATE:   3,
  DEST_COMPANY: 4,
  DEST_PERSON:  5,
  LINE_NO:      6,
  ITEM_NAME:    7,
  SPEC:         8,
  QTY:          9,
  UNIT:         10,
  UNIT_PRICE:   11,
  AMOUNT:       12,
  REMARKS:      13,
  PDF_URL:      14,
  FOLDER_URL:   15,
};

var ORDER_COLS = {
  MGMT_ID:        1,
  ORDER_NO:       2,
  LINKED_QUOTE:   3,
  ORDER_TYPE:     4,
  ORDER_DATE:     5,
  MODEL_CODE:     6,
  ORDER_SLIP_NO:  7,
  LINE_NO:        8,
  ITEM_NAME:      9,
  SPEC:           10,
  FIRST_DELIVERY: 11,
  DELIVERY_DEST:  12,
  QTY:            13,
  UNIT:           14,
  UNIT_PRICE:     15,
  AMOUNT:         16,
  REMARKS:        17,
  PDF_URL:        18,
  FOLDER_URL:     19,
};

var TODO_COLS = {
  ID:          1,
  TITLE:       2,
  CLIENT:      3,
  DUE_DATE:    4,
  PRIORITY:    5,
  STATUS:      6,
  LINKED_MGMT: 7,
  MEMO:        8,
};

var MODEL_INFO_COLS = {
  BOARD_ID:         1,
  MODEL_CODE:       2,
  QUOTE_URL:        3,
  ORDER_URL:        4,
  PURCHASE_URL1:    5,
  PURCHASE_URL2:    6,
  PURCHASE_URL3:    7,
  LOCAL_SERVER_URL: 8,
  COMMENT:          9,
  UPDATED_AT:       10,
};

var LEDGER_COLS = {
  LEDGER_ID:        1,
  QUOTE_NO:         2,
  ISSUE_DATE:       3,
  DEST:             4,
  CATEGORY:         5,
  SUBJECT:          6,
  STATUS:           7,
  SAVE_URL:         8,
  MACHINE_CODE:     9,
  BOARD_NAME:       10,
  MODEL_NO:         11,
  AMOUNT:           12,
  SUBMIT_TO:        13,
  REMARKS:          14,
  SENT_DATE:        15,
  COMPOSITION_TYPE: 16,
};

var LEDGER_STATUS = {
  DRAFT:   '作成中',
  PENDING: '作成予定',
  SENT:    '送信済み',
  CANCEL:  'キャンセル',
};

// ============================================================
// 初期セットアップ
// ============================================================

function initialSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('スプレッドシートにバインドして実行してください。');
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  _createOrSetupSheet(ss, CONFIG.SHEET_MANAGEMENT, _getMgmtHeaders(),  '#E8F0FE');
  _createOrSetupSheet(ss, CONFIG.SHEET_QUOTES,     _getQuoteHeaders(), '#E6F4EA');
  _createOrSetupSheet(ss, CONFIG.SHEET_ORDERS,     _getOrderHeaders(), '#FEF7E0');
  _setupEmailConfigSheet(ss);
  _setupTodoSheet(ss);
  _setupLedgerSheet(ss);
  _setupModelInfoSheet(ss);
  // 機種マスタシート（21_model_master.gs）
  try { initModelMasterSheet(); } catch(e) { Logger.log('機種マスタ初期化スキップ: ' + e.message); }
  _registerTriggers();
  // 機種マスタ・発注期限のトリガーも登録
  try { setupMonitoringTriggers(); } catch(e) {}
  try { setupModelMasterTriggers(); } catch(e) {}

  SpreadsheetApp.getUi().alert(
    '初期化完了！\n\n' +
    '「メール監視設定」シートにキーワード・メールアドレスを入力してください。\n' +
    'その後、新しいバージョンでWebアプリをデプロイしてください。'
  );
}

function _createOrSetupSheet(ss, name, headers, color) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  var r = sheet.getRange(1, 1, 1, headers.length);
  r.setValues([headers]);
  r.setBackground(color);
  r.setFontWeight('bold');
  r.setFontSize(10);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

function _setupEmailConfigSheet(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_EMAIL_CFG);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_EMAIL_CFG);
  var headers = ['有効','種別','キーワード（ファイル名・件名）','送信元メールアドレス','宛先メールアドレス（自社）','注文種別','備考'];
  var hr = sheet.getRange(1, 1, 1, headers.length);
  hr.setValues([headers]).setBackground('#FCE8B2').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.getRange(2, 1, 1000, 1).insertCheckboxes();
}

function _setupTodoSheet(ss) {
  _createOrSetupSheet(ss, CONFIG.SHEET_TODO,
    ['Todo ID','タイトル','顧客名','期限日','優先度','ステータス','関連管理ID','メモ'], '#F3E8FD');
}

function _setupLedgerSheet(ss) {
  var headers = ['台帳ID','見積No.','発行日','宛先（企業名）','分類','件名','ステータス',
                 '保存先URL','機種コード','基板名','型番','見積金額','提出先担当者','備考','メール送信日'];
  var sheet = _createOrSetupSheet(ss, CONFIG.SHEET_LEDGER, headers, '#FFF3E0');
  sheet.getRange(2,5,1000,1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['試作','量産','修理','その他'],true).build());
  sheet.getRange(2,7,1000,1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['作成予定','作成中','送信済み','キャンセル'],true).build());
}

function _setupModelInfoSheet(ss) {
  var headers = ['基板ID','機種コード','関連見積書URL','関連注文書URL',
                 '仕入れ見積URL1','仕入れ見積URL2','仕入れ見積URL3','ローカルサーバーURL','コメント','最終更新日'];
  _createOrSetupSheet(ss, CONFIG.SHEET_MODEL_INFO, headers, '#E8F5E9');
}

function _getMgmtHeaders() {
  return ['管理ID','見積番号','注文番号','件名','顧客名','ステータス','見積日','発注日',
          '見積金額','注文金額','消費税','合計金額','見積書PDF URL','注文書PDF URL',
          '保存先フォルダURL','見積書シート行','注文書シート行','紐づけ済み','注文種別',
          '機種コード','発注伝票番号','担当者','納期','メモ','登録日時','更新日時','GmailID'];
}

function _getQuoteHeaders() {
  return ['管理ID','見積番号','発行日','送り先会社名','送り先担当者名',
          '行No','品名','仕様','数量','単位','単価','金額','備考','PDF URL','フォルダURL'];
}

function _getOrderHeaders() {
  return ['管理ID','注文番号','見積番号(紐づけ)','注文種別','発注日','機種コード','発注伝票番号',
          '行No','品名','仕様','初回納品日','納品先','数量','単位','単価','金額','備考','PDF URL','フォルダURL'];
}

function _registerTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('processNewEmails').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('processDriveImports').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('autoMatchNewOrders').timeBased().everyHours(1).create();
  Logger.log('トリガー登録完了');
}

// ============================================================
// メール監視設定
// ============================================================

function getEmailConfigs() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_EMAIL_CFG);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var configs = [];
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues().forEach(function(row) {
    var enabled  = row[0] === true || row[0] === 'TRUE';
    var keywords = String(row[2]).split(',').map(function(k){ return k.trim().toLowerCase(); }).filter(Boolean);
    if (enabled && keywords.length > 0) {
      configs.push({
        docType:   String(row[1]).trim() === '見積書' ? 'quote' : 'order',
        keywords:  keywords,
        fromEmail: String(row[3]).trim().toLowerCase(),
        toEmail:   String(row[4]).trim().toLowerCase(),
        orderType: String(row[5]).trim(),
      });
    }
  });
  return configs;
}

function matchEmailConfig(filename, subject, fromAddr, toAddr) {
  var text = (filename + ' ' + subject).toLowerCase();
  var from = (fromAddr || '').toLowerCase();
  var to   = (toAddr   || '').toLowerCase();
  var configs = getEmailConfigs();
  for (var i = 0; i < configs.length; i++) {
    var cfg = configs[i];
    if (!cfg.keywords.some(function(k){ return text.indexOf(k) >= 0; })) continue;
    if (cfg.fromEmail && from.indexOf(cfg.fromEmail) < 0) continue;
    if (cfg.toEmail   && to.indexOf(cfg.toEmail)     < 0) continue;
    return cfg;
  }
  return null;
}

// ============================================================
// 基本ユーティリティ
// ============================================================

function getGeminiApiKey() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
}

function getSpreadsheet() {
  var id = CONFIG.SPREADSHEET_ID ||
           PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id || String(id).trim() === '') {
    throw new Error('スプレッドシートIDが未設定です。システム管理 → 設定で SPREADSHEET_ID を登録してください。');
  }
  return SpreadsheetApp.openById(id);
}

function generateMgmtId() {
  return 'QM-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd') + '-' +
    (Math.floor(Math.random() * 9000) + 1000);
}

function generateTodoId() {
  return 'TD-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');
}

function generateLedgerId() {
  return 'LQ-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd') + '-' +
    (Math.floor(Math.random() * 9000) + 1000);
}

function nowJST() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}

function getFolderUrl(folderId) {
  return 'https://drive.google.com/drive/folders/' + folderId;
}

function _getOrderFolderId(orderType) {
  if (orderType === CONFIG.ORDER_TYPE.TRIAL) return CONFIG.ORDER_TRIAL_FOLDER_ID;
  if (orderType === CONFIG.ORDER_TYPE.MASS)  return CONFIG.ORDER_MASS_FOLDER_ID;
  return CONFIG.WEB_UPLOAD_FOLDER_ID;
}

function normalizeText(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .toLowerCase()
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) { return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); })
    .replace(/[　]/g, ' ')
    .trim();
}

function _loadSettingsObj() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('SYS_SETTINGS');
    if (raw) return JSON.parse(raw);
  } catch (e) {}
  return { webhookUrl: '', notifyOrder: true, notifyQuote: true, notifyDl: true, alertDays: 3 };
}

function _isLinkedVal(val) {
  return val === true || String(val).toUpperCase() === 'TRUE';
}

function _toDateStr(v) {
  if (!v) return '';
  try {
    var d = (v instanceof Date) ? v : new Date(v);
    if (isNaN(d.getTime())) return String(v);
    return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
  } catch (e) { return String(v); }
}

function _toNum(v) {
  if (v === '' || v === null || v === undefined) return 0;
  var n = Number(String(v).replace(/[,¥￥\s]/g, ''));
  return isNaN(n) ? 0 : n;
}

// ============================================================
// データ取得
// ============================================================

function getAllMgmtData() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var last  = sheet.getLastRow();
  if (last <= 1) return [];
  var readCols = Math.max(sheet.getLastColumn(), 33);
  return sheet.getRange(2, 1, last - 1, readCols).getValues()
    .filter(function(r) { return r[0] !== ''; });
}

function getAllTodoData() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_TODO);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues()
    .filter(function(r) { return r[0] !== ''; });
}

function getAllLedgerData() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_LEDGER);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).getValues();
  data.forEach(function(row, i) {
    var hasContent = row.slice(1).some(function(v) { return String(v).trim() !== ''; });
    if (!hasContent) return;
    if (row[0] === '' || row[0] === null || row[0] === undefined) {
      row[0] = generateLedgerId();
      sheet.getRange(i + 2, 1).setValue(row[0]);
    }
  });
  return data.filter(function(r) { return r.slice(1).some(function(v){ return String(v).trim() !== ''; }); });
}

function getAllModelInfoData() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_MODEL_INFO);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues()
    .filter(function(r) { return String(r[0]).trim() !== ''; });
}

function findMgmtRowByQuoteNo(quoteNo) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var last  = sheet.getLastRow();
  if (last <= 1) return -1;
  var vals = sheet.getRange(2, MGMT_COLS.QUOTE_NO, last - 1, 1).getValues().flat();
  var idx  = vals.findIndex(function(v) { return String(v).trim() === String(quoteNo).trim(); });
  return idx >= 0 ? idx + 2 : -1;
}

function isMessageAlreadyProcessed(msgId) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var last  = sheet.getLastRow();
  if (last <= 1) return false;
  var ids = sheet.getRange(2, MGMT_COLS.GMAIL_MSG_ID, last - 1, 1).getValues().flat();
  return ids.some(function(v) { return String(v).trim() === String(msgId).trim(); });
}

function _ledgerRowToObject(row) {
  return {
    ledgerId:        String(row[0]  || ''),
    quoteNo:         String(row[1]  || ''),
    issueDate:       _toDateStr(row[2]),
    dest:            String(row[3]  || ''),
    category:        String(row[4]  || ''),
    subject:         String(row[5]  || ''),
    status:          String(row[6]  || ''),
    saveUrl:         String(row[7]  || ''),
    machineCode:     String(row[8]  || ''),
    boardName:       String(row[9]  || ''),
    modelNo:         String(row[10] || ''),
    amount:          (row[11] !== '' && row[11] !== null) ? Number(row[11]) : null,
    submitTo:        String(row[12] || ''),
    remarks:         String(row[13] || ''),
    sentDate:        _toDateStr(row[14]),
    compositionType: String(row[15] || ''),
  };
}

function _modelInfoRowToObject(row) {
  return {
    boardId:        String(row[0] || ''),
    modelCode:      String(row[1] || ''),
    quoteUrl:       String(row[2] || ''),
    orderUrl:       String(row[3] || ''),
    purchaseUrl1:   String(row[4] || ''),
    purchaseUrl2:   String(row[5] || ''),
    purchaseUrl3:   String(row[6] || ''),
    localServerUrl: String(row[7] || ''),
    comment:        String(row[8] || ''),
    updatedAt:      _toDateStr(row[9]),
  };
}

// ============================================================
// 通信ユーティリティ
// ============================================================

function fetchWithRetry(url, options, retries) {
  retries = retries || CONFIG.RATE_LIMIT_RETRIES;
  for (var i = 0; i < retries; i++) {
    try {
      var res  = UrlFetchApp.fetch(url, options);
      var code = res.getResponseCode();
      if (code === 429) { Utilities.sleep(CONFIG.RATE_LIMIT_WAIT_MS); continue; }
      if (code === 404) throw new Error('HTTP 404: ' + res.getContentText().substring(0, 200));
      if (code >= 200 && code < 300) return res;
      throw new Error('HTTP ' + code + ': ' + res.getContentText().substring(0, 200));
    } catch (e) {
      if (i === retries - 1) throw e;
      Utilities.sleep(3000);
    }
  }
}

// ============================================================
// マッチング候補の取得・手動確定（旧 00 missing utils.gs）
// ============================================================

function getMatchingCandidates() {
  try {
    var raw    = PropertiesService.getScriptProperties().getProperty('AI_MATCHING_CANDIDATES');
    var stored = raw ? JSON.parse(raw) : [];
    var allMgmt = getAllMgmtData();

    var unlinkedOrders = allMgmt.filter(function(r) {
      return String(r[MGMT_COLS.ORDER_NO - 1]).trim() !== '' &&
             !_isLinkedVal(r[MGMT_COLS.LINKED - 1]);
    });

    var ss = getSpreadsheet();
    var qs = ss.getSheetByName(CONFIG.SHEET_QUOTES);
    var quoteLineMap = {};
    if (qs && qs.getLastRow() > 1) {
      qs.getRange(2, 1, qs.getLastRow() - 1, 15).getValues()
        .filter(function(r) { return r[0] && r[6]; })
        .forEach(function(r) {
          var mid = String(r[0]);
          if (!quoteLineMap[mid]) quoteLineMap[mid] = [];
          quoteLineMap[mid].push({ itemName: String(r[6]||''), spec: String(r[7]||''), unitPrice: r[10], pdfUrl: String(r[13]||'') });
        });
    }

    var unlinkedQuoteObjs = allMgmt.filter(function(r) {
      return String(r[MGMT_COLS.QUOTE_NO - 1]).trim() !== '' && !_isLinkedVal(r[MGMT_COLS.LINKED - 1]);
    }).map(function(r) {
      var mid = String(r[MGMT_COLS.ID - 1]);
      return {
        quoteId:   mid,
        quoteNo:   String(r[MGMT_COLS.QUOTE_NO - 1]     || ''),
        client:    String(r[MGMT_COLS.CLIENT - 1]        || ''),
        issueDate: _toDateStr(r[MGMT_COLS.QUOTE_DATE - 1]),
        amount:    _toNum(r[MGMT_COLS.QUOTE_AMOUNT - 1]),
        subject:   String(r[MGMT_COLS.SUBJECT - 1]       || ''),
        quoteUrl:  String(r[MGMT_COLS.QUOTE_PDF_URL - 1] || ''),
        items:     quoteLineMap[mid] || [],
      };
    });

    var storedIds = {};
    stored.forEach(function(s) { storedIds[s.orderMgmtId] = true; });

    var FALLBACK_REASON = '類似候補が見つからなかったため、未紐づけ見積書を参考表示しています。';

    function _makeFallbackCandidates(orderClient) {
      var fb = unlinkedQuoteObjs.filter(function(q) {
        return !orderClient || q.client.indexOf(orderClient) >= 0 || orderClient.indexOf(q.client) >= 0;
      });
      if (fb.length === 0) fb = unlinkedQuoteObjs.slice(0, 3);
      return fb.slice(0, 3).map(function(q) {
        return { quoteId: q.quoteId, quoteNo: q.quoteNo, client: q.client,
                 issueDate: q.issueDate, amount: q.amount, subject: q.subject,
                 quoteUrl: q.quoteUrl, items: q.items, score: 0, reason: FALLBACK_REASON };
      });
    }

    unlinkedOrders.forEach(function(r) {
      var mid         = String(r[MGMT_COLS.ID - 1]);
      var orderClient = String(r[MGMT_COLS.CLIENT - 1] || '');
      if (storedIds[mid]) return;
      stored.push({
        orderMgmtId: mid,
        orderNo:     String(r[MGMT_COLS.ORDER_NO - 1]     || ''),
        orderClient: orderClient,
        orderDate:   _toDateStr(r[MGMT_COLS.ORDER_DATE - 1]),
        orderAmount: _toNum(r[MGMT_COLS.ORDER_AMOUNT - 1]),
        orderPdfUrl: String(r[MGMT_COLS.ORDER_PDF_URL - 1] || ''),
        candidates:  _makeFallbackCandidates(orderClient),
      });
    });

    stored = stored.map(function(item) {
      if (item.candidates && item.candidates.length > 0) return item;
      item.candidates = _makeFallbackCandidates(item.orderClient || '');
      return item;
    });

    return { success: true, items: stored };
  } catch (e) {
    Logger.log('[getMatchingCandidates ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

function confirmManualLink(orderMgmtId, quoteMgmtId) {
  try {
    if (!orderMgmtId || !quoteMgmtId) return { success: false, error: '管理IDが不足しています' };

    var ss        = getSpreadsheet();
    var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var last      = mgmtSheet.getLastRow();
    if (last <= 1) return { success: false, error: 'データなし' };

    var allData = mgmtSheet.getRange(2, 1, last - 1, 27).getValues();
    var ids     = allData.map(function(r) { return String(r[MGMT_COLS.ID - 1]); });

    var oIdx = ids.indexOf(String(orderMgmtId));
    if (oIdx < 0) return { success: false, error: '注文書管理IDが見つかりません: ' + orderMgmtId };
    var qIdx = ids.indexOf(String(quoteMgmtId));
    if (qIdx < 0) return { success: false, error: '見積書管理IDが見つかりません: ' + quoteMgmtId };

    var quoteNo     = String(allData[qIdx][MGMT_COLS.QUOTE_NO - 1]      || '');
    var quotePdfUrl = String(allData[qIdx][MGMT_COLS.QUOTE_PDF_URL - 1] || '');
    var orderNo     = String(allData[oIdx][MGMT_COLS.ORDER_NO - 1]      || '');
    var orderPdfUrl = String(allData[oIdx][MGMT_COLS.ORDER_PDF_URL - 1] || '');

    mgmtSheet.getRange(oIdx+2, MGMT_COLS.QUOTE_NO).setValue(quoteNo);
    mgmtSheet.getRange(oIdx+2, MGMT_COLS.QUOTE_PDF_URL).setValue(quotePdfUrl);
    mgmtSheet.getRange(oIdx+2, MGMT_COLS.LINKED).setValue('TRUE');
    mgmtSheet.getRange(oIdx+2, MGMT_COLS.STATUS).setValue(CONFIG.STATUS.RECEIVED);
    mgmtSheet.getRange(oIdx+2, MGMT_COLS.UPDATED_AT).setValue(nowJST());

    mgmtSheet.getRange(qIdx+2, MGMT_COLS.ORDER_NO).setValue(orderNo);
    mgmtSheet.getRange(qIdx+2, MGMT_COLS.ORDER_PDF_URL).setValue(orderPdfUrl);
    mgmtSheet.getRange(qIdx+2, MGMT_COLS.LINKED).setValue('TRUE');
    mgmtSheet.getRange(qIdx+2, MGMT_COLS.STATUS).setValue(CONFIG.STATUS.RECEIVED);
    mgmtSheet.getRange(qIdx+2, MGMT_COLS.UPDATED_AT).setValue(nowJST());

    try {
      var raw = PropertiesService.getScriptProperties().getProperty('AI_MATCHING_CANDIDATES');
      if (raw) {
        var stored = JSON.parse(raw).filter(function(item) {
          return String(item.orderMgmtId) !== String(orderMgmtId);
        });
        PropertiesService.getScriptProperties().setProperty('AI_MATCHING_CANDIDATES', JSON.stringify(stored));
      }
    } catch (e2) {}

    return { success: true, quoteNo: quoteNo, orderNo: orderNo };
  } catch (e) {
    Logger.log('[confirmManualLink ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

function _findLatestOrderMgmtId() {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var last  = sheet.getLastRow();
    if (last <= 1) return null;
    var row = sheet.getRange(last, MGMT_COLS.ID).getValue();
    return row ? String(row) : null;
  } catch (e) { return null; }
}

function _sendOrderRegistrationToChat(mgmtId, uploadInfo, linkResult) {
  try {
    var webhookUrl = CONFIG.GOOGLE_CHAT_WEBHOOK_URL ||
      PropertiesService.getScriptProperties().getProperty('GOOGLE_CHAT_WEBHOOK_URL');
    if (!webhookUrl) return;
    var lr    = linkResult || {};
    var lines = ['【📦 注文書を登録（Drive自動取込）】',
                 '発注書番号: ' + (uploadInfo.documentNo || '—'),
                 '種別: '      + (uploadInfo.orderType   || '—')];
    if (lr.status === 'auto_linked')      lines.push('✅ AI自動紐づけ完了');
    else if (lr.status === 'candidates_found') lines.push('⚠️ 紐づけ候補あり（要確認）');
    else                                   lines.push('❌ 紐づく見積書が見つかりませんでした');
    try { lines.push('▶ ' + ScriptApp.getService().getUrl()); } catch (e2) {}
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ text: lines.join('\n') }),
      muteHttpExceptions: true,
    });
  } catch (e) {
    Logger.log('[_sendOrderRegistrationToChat ERROR] ' + e.message);
  }
}

// ============================================================
// 診断・テスト用
// ============================================================

function testGeminiConnection() {
  var apiKey = getGeminiApiKey();
  Logger.log('APIキー先頭8文字: ' + (apiKey ? apiKey.substring(0,8) : 'NULL'));
  var url = CONFIG.GEMINI_API_ENDPOINT + CONFIG.GEMINI_PRIMARY_MODEL + ':generateContent?key=' + apiKey;
  var res = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({contents:[{parts:[{text:'テスト。OKとだけ返して。'}]}]}),
    muteHttpExceptions: true,
  });
  Logger.log('Status: ' + res.getResponseCode());
  Logger.log('Body: '   + res.getContentText().substring(0,400));
}

function checkAvailableModels() {
  var res  = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?key=' + getGeminiApiKey(), {muteHttpExceptions: true});
  var data = JSON.parse(res.getContentText());
  Logger.log(data.models ? data.models.map(function(m){ return m.name; }).join('\n') : res.getContentText());
}

function checkMgmtColumns() {
  var sheet   = getSpreadsheet().getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log('管理シート 実際の列数: ' + headers.length);
  Logger.log('ヘッダー一覧: ' + JSON.stringify(headers));
}
