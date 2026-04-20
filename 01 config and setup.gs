// ============================================================
// 見積書・注文書管理システム
// ファイル 1/4: 設定・初期化・シートセットアップ
// ============================================================

var CONFIG = {
  SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '',
  GEMINI_API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '',

GEMINI_PRIMARY_MODEL:  'gemini-2.0-flash-lite',
GEMINI_FALLBACK_MODEL: 'gemini-2.0-flash',
GEMINI_API_ENDPOINT:   'https://generativelanguage.googleapis.com/v1beta/models/',

  WEB_UPLOAD_FOLDER_ID:  '1sB42xntGKL31GeT9OjOKTxVJwj9IQz-h',
  ORDER_TRIAL_FOLDER_ID: '1wVeYlt-9GsortfOsUggBsWta8GtXIRvS',
  ORDER_MASS_FOLDER_ID:  '1ASyV7PmhYQVH-72rVD3evToYJWxGhMbA',
  QUOTE_FOLDER_ID:       '1sB42xntGKL31GeT9OjOKTxVJwj9IQz-h',

  IMPORT_QUOTE_FOLDER_ID:       '1Y66PDSi35ScuIyS0Jgm0l3p2l7MEM2Jk',
  IMPORT_ORDER_TRIAL_FOLDER_ID: '1Ufq4xMjOmZvUQLC_Zp0EAWlHF0mYAGDM',
  IMPORT_ORDER_MASS_FOLDER_ID:  '1ujzCtYzOqU9_a6tiEXOHhDWRv15a0p0k',

  SHEET_MANAGEMENT: '管理シート',
  SHEET_QUOTES:     '見積書シート',
  SHEET_ORDERS:     '注文書シート',
  SHEET_EMAIL_CFG:  'メール監視設定',
  SHEET_TODO:       'Todoリスト',
  SHEET_LEDGER:     '見積台帳',
  SHEET_MODEL_INFO: '基板情報管理',

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

// ===== 管理シート 列定義（33列）=====
// ★実際のシートが27列でも、30列以上読み込んでも空欄として処理される
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
  // ★追加：参照されているが未定義だった列
  IS_LATEST:         28,
  ORDER_DEADLINE:    29,
  REVISION_NO:       30,
  PARENT_MGMT_ID:    31,  // ★ _apiCreateRevision で参照
  DEADLINE_NOTIFIED: 32,  // ★ checkOrderDeadlines で参照
};

// ===== 見積書シート 列定義（15列）=====
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

// ===== 注文書シート 列定義（19列）=====
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

// ===== Todoシート 列定義（8列）=====
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

// ===== 基板情報管理 列定義（10列）=====
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

// ===== 見積台帳 列定義（14列）=====
var LEDGER_COLS = {
  LEDGER_ID:    1,
  QUOTE_NO:     2,
  ISSUE_DATE:   3,
  DEST:         4,
  CATEGORY:     5,
  SUBJECT:      6,
  STATUS:       7,
  SAVE_URL:     8,
  MACHINE_CODE: 9,
  BOARD_NAME:   10,
  MODEL_NO:     11,
  AMOUNT:       12,
  SUBMIT_TO:    13,
  REMARKS:      14,
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
  _registerTriggers();

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
  var headers = [
    '有効', '種別', 'キーワード（ファイル名・件名）', '送信元メールアドレス',
    '宛先メールアドレス（自社）', '注文種別', '備考'
  ];
  var hr = sheet.getRange(1, 1, 1, headers.length);
  hr.setValues([headers]);
  hr.setBackground('#FCE8B2');
  hr.setFontWeight('bold');
  sheet.setFrozenRows(1);
  var samples = [
    ['TRUE',  '見積書', '見積,mitsumori,quote,見積書', '',                       'yourcompany@gmail.com', '',    '自社送信済み見積書の自動検知'],
    ['TRUE',  '注文書', '注文,発注,order,発注書,purchase', 'client@example.com', '',                       '試作', '得意先Aからの試作注文'],
    ['FALSE', '注文書', '注文,量産',                 'client2@example.com',  '',                       '量産', '※無効化サンプル'],
  ];
  sheet.getRange(2, 1, samples.length, headers.length).setValues(samples);
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 80);
  sheet.setColumnWidth(7, 200);
  var lastRow = Math.max(sheet.getLastRow(), 10);
  sheet.getRange(2, 1, lastRow, 1).insertCheckboxes();
}

function _setupTodoSheet(ss) {
  var headers = ['Todo ID', 'タイトル', '顧客名', '期限日', '優先度', 'ステータス', '関連管理ID', 'メモ'];
  _createOrSetupSheet(ss, CONFIG.SHEET_TODO, headers, '#F3E8FD');
}

function _setupLedgerSheet(ss) {
  var headers = ['台帳ID','見積No.','発行日','宛先（企業名）','分類','件名','ステータス','保存先URL','機種コード','基板名','型番','見積金額','提出先担当者','備考'];
  var sheet = _createOrSetupSheet(ss, CONFIG.SHEET_LEDGER, headers, '#FFF3E0');
  var catRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['試作','量産','修理','その他'], true).build();
  sheet.getRange(2, 5, 1000, 1).setDataValidation(catRule);
  var stRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['作成予定','作成中','送信済み','キャンセル'], true).build();
  sheet.getRange(2, 7, 1000, 1).setDataValidation(stRule);
  sheet.setColumnWidth(1,  140);
  sheet.setColumnWidth(2,  110);
  sheet.setColumnWidth(3,  100);
  sheet.setColumnWidth(4,  180);
  sheet.setColumnWidth(5,  80);
  sheet.setColumnWidth(6,  220);
  sheet.setColumnWidth(7,  90);
  sheet.setColumnWidth(8,  300);
  sheet.setColumnWidth(9,  120);
  sheet.setColumnWidth(10, 180);
  sheet.setColumnWidth(11, 130);
  sheet.setColumnWidth(12, 100);
  sheet.setColumnWidth(13, 130);
  sheet.setColumnWidth(14, 250);
}

function _setupModelInfoSheet(ss) {
  var headers = [
    '基板ID', '機種コード',
    '関連見積書URL', '関連注文書URL',
    '仕入れ見積URL1', '仕入れ見積URL2', '仕入れ見積URL3',
    'ローカルサーバーURL', 'コメント', '最終更新日'
  ];
  var sheet = _createOrSetupSheet(ss, CONFIG.SHEET_MODEL_INFO, headers, '#E8F5E9');
  sheet.setColumnWidth(1,  120);
  sheet.setColumnWidth(2,  120);
  sheet.setColumnWidth(3,  280);
  sheet.setColumnWidth(4,  280);
  sheet.setColumnWidth(5,  280);
  sheet.setColumnWidth(6,  280);
  sheet.setColumnWidth(7,  280);
  sheet.setColumnWidth(8,  280);
  sheet.setColumnWidth(9,  300);
  sheet.setColumnWidth(10, 140);
}

function getAllModelInfoData() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_MODEL_INFO);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues()
    .filter(function(r) { return String(r[0]).trim() !== ''; });
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

function generateLedgerId() {
  return 'LQ-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd') + '-' +
    (Math.floor(Math.random() * 9000) + 1000);
}

function getAllLedgerData() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_LEDGER);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
  data.forEach(function(row, i) {
    var hasContent = row.slice(1).some(function(v) { return String(v).trim() !== ''; });
    if (!hasContent) return;
    if (row[0] === '' || row[0] === null || row[0] === undefined) {
      row[0] = generateLedgerId();
      sheet.getRange(i + 2, 1).setValue(row[0]);
    }
  });
  return data.filter(function(r) {
    return r.slice(1).some(function(v) { return String(v).trim() !== ''; });
  });
}

function _ledgerRowToObject(row) {
  return {
    ledgerId:    String(row[0]  || ''),
    quoteNo:     String(row[1]  || ''),
    issueDate:   _toDateStr(row[2]),
    dest:        String(row[3]  || ''),
    category:    String(row[4]  || ''),
    subject:     String(row[5]  || ''),
    status:      String(row[6]  || ''),
    saveUrl:     String(row[7]  || ''),
    machineCode: String(row[8]  || ''),
    boardName:   String(row[9]  || ''),
    modelNo:     String(row[10] || ''),
    amount:      (row[11] !== '' && row[11] !== null && row[11] !== undefined) ? Number(row[11]) : null,
    submitTo:    String(row[12] || ''),
    remarks:     String(row[13] || ''),
  };
}

function _getMgmtHeaders() {
  return [
    '管理ID','見積番号','注文番号','件名','顧客名',
    'ステータス','見積日','発注日','見積金額','注文金額',
    '消費税','合計金額','見積書PDF URL','注文書PDF URL','保存先フォルダURL',
    '見積書シート行','注文書シート行','紐づけ済み','注文種別',
    '機種コード','発注伝票番号','担当者','納期','メモ',
    '登録日時','更新日時','GmailID',
  ];
}

function _getQuoteHeaders() {
  return [
    '管理ID','見積番号',
    '発行日','送り先会社名','送り先担当者名',
    '行No','品名','仕様','数量','単位','単価','金額','備考','PDF URL','フォルダURL'
  ];
}

function _getOrderHeaders() {
  return [
    '管理ID','注文番号','見積番号(紐づけ)','注文種別',
    '発注日','機種コード','発注伝票番号',
    '行No','品名','仕様','初回納品日','納品先',
    '数量','単位','単価','金額','備考','PDF URL','フォルダURL',
  ];
}

function _registerTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('processNewEmails').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('processDriveImports').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('autoMatchNewOrders').timeBased().everyHours(1).create();
  Logger.log('トリガー登録完了');
}

// ============================================================
// メール監視設定シート読み込み
// ============================================================

function getEmailConfigs() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_EMAIL_CFG);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  var configs = [];
  data.forEach(function(row) {
    var enabled   = row[0] === true || row[0] === 'TRUE';
    var docType   = String(row[1]).trim();
    var keywords  = String(row[2]).split(',').map(function(k) { return k.trim().toLowerCase(); }).filter(Boolean);
    var fromEmail = String(row[3]).trim().toLowerCase();
    var toEmail   = String(row[4]).trim().toLowerCase();
    var orderType = String(row[5]).trim();
    if (enabled && keywords.length > 0) {
      configs.push({
        docType:   docType === '見積書' ? 'quote' : 'order',
        keywords:  keywords,
        fromEmail: fromEmail,
        toEmail:   toEmail,
        orderType: orderType,
      });
    }
  });
  return configs;
}

function matchEmailConfig(filename, subject, fromAddr, toAddr) {
  var configs = getEmailConfigs();
  var text = (filename + ' ' + subject).toLowerCase();
  var from = (fromAddr || '').toLowerCase();
  var to   = (toAddr   || '').toLowerCase();
  for (var i = 0; i < configs.length; i++) {
    var cfg = configs[i];
    var kwMatch = cfg.keywords.some(function(k) { return text.indexOf(k) >= 0; });
    if (!kwMatch) continue;
    if (cfg.fromEmail && from.indexOf(cfg.fromEmail) < 0) continue;
    if (cfg.toEmail   && to.indexOf(cfg.toEmail) < 0)     continue;
    return cfg;
  }
  return null;
}

// ============================================================
// ユーティリティ
// ============================================================

function getSpreadsheet() {
  var id = CONFIG.SPREADSHEET_ID ||
           PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return SpreadsheetApp.openById(id);
}

function generateMgmtId() {
  return 'QM-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd') + '-' +
    (Math.floor(Math.random() * 9000) + 1000);
}

function generateTodoId() {
  return 'TD-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');
}

function nowJST() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}

function getFolderUrl(folderId) {
  return 'https://drive.google.com/drive/folders/' + folderId;
}

function getAllMgmtData() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var last  = sheet.getLastRow();
  if (last <= 1) return [];
  // ★実際の列数より多く指定してもGASは空欄として返すので安全
  var actualCols = sheet.getLastColumn();
  var readCols   = Math.max(actualCols, 27); // 最低27列は読む
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

function fetchWithRetry(url, options, retries) {
  retries = retries || CONFIG.RATE_LIMIT_RETRIES;
  for (var i = 0; i < retries; i++) {
    try {
      var res  = UrlFetchApp.fetch(url, options);
      var code = res.getResponseCode();
      if (code === 429) {
        Logger.log('[RATE_LIMIT] 429 → ' + CONFIG.RATE_LIMIT_WAIT_MS/1000 + 's待機');
        Utilities.sleep(CONFIG.RATE_LIMIT_WAIT_MS);
        continue;
      }
      if (code === 404) throw new Error('HTTP 404: ' + res.getContentText().substring(0, 200));
      if (code >= 200 && code < 300) return res;
      throw new Error('HTTP ' + code + ': ' + res.getContentText().substring(0, 200));
    } catch(e) {
      if (i === retries - 1) throw e;
      Logger.log('[RETRY ' + (i+1) + '/' + retries + '] ' + e.message);
      Utilities.sleep(3000);
    }
  }
}

function _getOrderFolderId(orderType) {
  if (orderType === CONFIG.ORDER_TYPE.TRIAL) return CONFIG.ORDER_TRIAL_FOLDER_ID;
  if (orderType === CONFIG.ORDER_TYPE.MASS)  return CONFIG.ORDER_MASS_FOLDER_ID;
  return CONFIG.WEB_UPLOAD_FOLDER_ID;
}

function testGeminiConnection() {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
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
  var apiKey = CONFIG.GEMINI_API_KEY;
  var url    = 'https://generativelanguage.googleapis.com/v1beta/models?key=' + apiKey;
  var res    = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  var data   = JSON.parse(res.getContentText());
  if (data.models) {
    Logger.log('【使えるモデル一覧】\n' + data.models.map(function(m){ return m.name; }).join('\n'));
  } else {
    Logger.log('エラー: ' + res.getContentText());
  }
}

// ★診断用：実際のシート列数を確認する関数
function checkMgmtColumns() {
  var ss      = getSpreadsheet();
  var sheet   = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log('管理シート 実際の列数: ' + headers.length);
  Logger.log('ヘッダー一覧: ' + JSON.stringify(headers));
}