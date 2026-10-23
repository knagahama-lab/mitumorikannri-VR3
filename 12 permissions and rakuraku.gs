// ============================================================
// 見積書・注文書管理システム
// ファイル 12: 権限管理・メール配信・楽楽販売連携
// ============================================================

var VIEWER_SHEET = '閲覧権限管理';

/** 楽楽販売からのWebhook受信（doGet側でrouting） */
function handleRakurakuWebhook(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    Logger.log('[RAKU WEBHOOK] ' + JSON.stringify(body).substring(0, 200));
    
    var eventType = body.event_type || body.event || '';
    var est       = body.estimate || body.data || {};
    
    if (eventType.indexOf('estimate') >= 0) {
      _apiSyncRakurakuNow({});
    }
    
    return ContentService.createTextOutput(JSON.stringify({ received: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(e2) {
    return ContentService.createTextOutput(JSON.stringify({ error: e2.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
// ============================================================
// メニュー設定の全ユーザー共有保存
// ============================================================
function _apiMenuConfigLoad() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('SHARED_MENU_CONFIG');
    return { success: true, menuConfig: raw ? JSON.parse(raw) : null };
  } catch(e) { return { success: false, error: e.message }; }
}

function _apiMenuConfigSave(p) {
  try {
    PropertiesService.getScriptProperties().setProperty('SHARED_MENU_CONFIG', JSON.stringify(p.menuConfig));
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ★ 楽楽販売Webhook（handleRakurakuWebhook）から呼ばれるため残す
var RAKU_SETTINGS_KEY = 'RAKURAKU_SETTINGS';
var RAKU_LOG_KEY      = 'RAKURAKU_SYNC_LOG';

function _apiSyncRakurakuNow(p) {
  try {
    var props  = PropertiesService.getScriptProperties();
    var saved  = JSON.parse(props.getProperty(RAKU_SETTINGS_KEY) || '{}');
    var tenant = saved.tenant;
    var apiKey = saved.apiKey;
    var mapping = saved.mapping || {};
    var syncStatuses = (saved.syncStatus || '承認済み,送付済み').split(',').map(function(s) { return s.trim(); });
    
    if (!tenant || !apiKey) return { success: false, error: 'テナントIDとAPIキーを設定してください' };
    
    var log = [];
    var now = nowJST();
    log.push('[' + now + '] 同期開始');
    
    // 楽楽販売APIから見積書一覧取得
    var url = 'https://' + tenant + '.raku-raku.jp/api/v1/estimates?limit=100';
    var res = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + apiKey },
      muteHttpExceptions: true,
    });
    
    if (res.getResponseCode() !== 200) {
      var errMsg = 'API取得エラー: ' + res.getResponseCode();
      log.push('[ERROR] ' + errMsg);
      _saveRakuLog(log);
      return { success: false, error: errMsg };
    }
    
    var apiData = JSON.parse(res.getContentText());
    var estimates = apiData.data || apiData.estimates || [];
    
    log.push('[INFO] 取得件数: ' + estimates.length + '件');
    
    // フィールドマッピング
    var qNoField     = mapping.quoteno || '見積番号';
    var clientField  = mapping.client  || '取引先名';
    var amountField  = mapping.amount  || '見積金額';
    var dateField    = mapping.date    || '見積日';
    var statusField  = mapping.status  || 'ステータス';
    
    // 管理シートに取込み
    var ss = getSpreadsheet();
    var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var mgmtData  = mgmtSheet.getDataRange().getValues();
    var existingQuoteNos = mgmtData.slice(1).map(function(r) { return String(r[MGMT_COLS.QUOTE_NO - 1] || '').trim(); });
    
    var imported = 0;
    estimates.forEach(function(est) {
      var quoteNo = String(est[qNoField] || est.estimate_no || '').trim();
      var status  = String(est[statusField] || est.status || '').trim();
      
      if (!quoteNo) return;
      
      // 対象ステータスチェック
      if (syncStatuses.length > 0 && !syncStatuses.some(function(s) { return status.indexOf(s) >= 0; })) {
        return; // 対象外ステータス
      }
      
      // 重複チェック
      if (existingQuoteNos.indexOf(quoteNo) >= 0) {
        log.push('[SKIP] 既存: ' + quoteNo);
        return;
      }
      
      // 管理シートに新規追加
      var mgmtId  = generateMgmtId();
      var client  = String(est[clientField]  || est.customer_name || '').trim();
      var amount  = Number(est[amountField]  || est.total_amount  || 0);
      var issDate = String(est[dateField]    || est.estimate_date || '').trim();
      
      var row = new Array(27).fill('');
      row[MGMT_COLS.ID          - 1] = mgmtId;
      row[MGMT_COLS.QUOTE_NO    - 1] = quoteNo;
      row[MGMT_COLS.CLIENT      - 1] = client;
      row[MGMT_COLS.STATUS      - 1] = CONFIG.STATUS.SENT;
      row[MGMT_COLS.QUOTE_DATE  - 1] = issDate;
      row[MGMT_COLS.QUOTE_AMOUNT- 1] = amount;
      row[MGMT_COLS.MEMO        - 1] = '楽楽販売から同期 (' + now + ')';
      row[MGMT_COLS.CREATED_AT  - 1] = now;
      row[MGMT_COLS.UPDATED_AT  - 1] = now;
      
      mgmtSheet.appendRow(row);
      existingQuoteNos.push(quoteNo);
      imported++;
      log.push('[OK] 取込: ' + quoteNo + ' | ' + client + ' | ¥' + amount.toLocaleString());
    });
    
    log.push('[' + nowJST() + '] 同期完了: ' + imported + '件 取込');
    _saveRakuLog(log);
    
    return { success: true, imported: imported, log: log };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _saveRakuLog(logLines) {
  try {
    var maxLines = 50;
    var existing = JSON.parse(PropertiesService.getScriptProperties().getProperty(RAKU_LOG_KEY) || '[]');
    var merged = logLines.concat(existing).slice(0, maxLines);
    PropertiesService.getScriptProperties().setProperty(RAKU_LOG_KEY, JSON.stringify(merged));
  } catch(e) {}
}
