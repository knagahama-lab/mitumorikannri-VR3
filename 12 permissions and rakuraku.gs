// ============================================================
// 見積書・注文書管理システム
// ファイル 12: 権限管理・メール配信・楽楽販売連携
// ============================================================

var VIEWER_SHEET = '閲覧権限管理';

// ============================================================
// ① 閲覧権限管理
// ============================================================

function _getViewerSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(VIEWER_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(VIEWER_SHEET);
    var headers = ['行番号', 'メールアドレス', '氏名', '部署', '権限レベル', '備考', '登録日時'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#7c3aed').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(5, 100);
  }
  return sheet;
}

function _apiGetViewerPermissions() {
  try {
    var sheet = _getViewerSheet();
    if (sheet.getLastRow() <= 1) return { success: true, viewers: [] };
    var data = sheet.getDataRange().getValues();
    var heads = data[0];
    var viewers = data.slice(1).map(function(row, i) {
      return {
        rowIndex: i + 2,
        email:    String(row[1] || ''),
        name:     String(row[2] || ''),
        dept:     String(row[3] || ''),
        role:     String(row[4] || 'viewer'),
        memo:     String(row[5] || ''),
        createdAt:String(row[6] || ''),
      };
    }).filter(function(v) { return v.email; });
    return { success: true, viewers: viewers };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _apiSaveViewerPermission(p) {
  try {
    var sheet = _getViewerSheet();
    var row = [
      '', // 行番号（自動）
      p.email || '', p.name || '', p.dept || '',
      p.role || 'viewer', p.memo || '', nowJST()
    ];
    if (p.rowIndex && Number(p.rowIndex) >= 2) {
      row[0] = Number(p.rowIndex); // 行番号セット
      sheet.getRange(Number(p.rowIndex), 1, 1, row.length).setValues([row]);
    } else {
      var newRow = sheet.getLastRow() + 1;
      row[0] = newRow;
      sheet.appendRow(row);
    }
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _apiDeleteViewerPermission(p) {
  try {
    if (!p.rowIndex || Number(p.rowIndex) < 2) return { success: false, error: 'rowIndex不足' };
    _getViewerSheet().deleteRow(Number(p.rowIndex));
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

/**
 * doGet の権限チェック（権限管理シートを参照）
 * 既存の isAdmin に加えてビューワー権限レベルを返す
 */
function getUserRole(userEmail) {
  if (!userEmail) return 'none';
  
  // 管理者メールチェック（既存ロジック）
  var adminEmails = (PropertiesService.getScriptProperties().getProperty('ADMIN_EMAILS') || '').split(',').map(function(s) { return s.trim().toLowerCase(); });
  if (!adminEmails.join('') || adminEmails.indexOf(userEmail.toLowerCase()) >= 0) return 'admin';
  
  // 権限管理シートを参照
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(VIEWER_SHEET);
    if (!sheet || sheet.getLastRow() <= 1) return 'viewer'; // シートなし = 全員閲覧可
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === userEmail.toLowerCase()) {
        return String(data[i][4] || 'viewer');
      }
    }
  } catch(e) {
    Logger.log('[ROLE] ' + e.message);
  }
  return 'none'; // 未登録ユーザー
}

// ============================================================
// ② メール配信設定
// ============================================================

var MAIL_SEND_KEY = 'MAIL_SEND_SETTINGS';

function _apiGetMailSendSettings() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(MAIL_SEND_KEY);
    var settings = raw ? JSON.parse(raw) : {};
    return { success: true, mailSendSettings: settings };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _apiSaveMailSendSettings(p) {
  try {
    PropertiesService.getScriptProperties().setProperty(MAIL_SEND_KEY, JSON.stringify(p));
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

/**
 * イベント発生時にグループメール配信する
 * eventKey: 'quote_created' | 'order_received' | 'link_confirmed' | 'link_failed' |
 *           'approval_required' | 'approval_done' | 'approval_rejected' | 'delivery_alert'
 */
function sendEventMail(eventKey, data) {
  try {
    var settingsRaw = PropertiesService.getScriptProperties().getProperty(MAIL_SEND_KEY);
    if (!settingsRaw) return;
    var settings = JSON.parse(settingsRaw);
    var evtCfg = (settings.events || {})[eventKey];
    if (!evtCfg || !evtCfg.enabled) return;
    
    var groups = settings.groups || {};
    var templates = settings.templates || {};
    
    // 配信先グループを解決
    var toGroups = String(evtCfg.to || 'sales').split(',');
    var toEmails = [];
    toGroups.forEach(function(g) {
      g = g.trim();
      if (g === 'all') {
        Object.values(groups).forEach(function(gl) {
          gl.split(',').forEach(function(e) { var em = e.trim(); if (em) toEmails.push(em); });
        });
      } else if (groups[g]) {
        groups[g].split(',').forEach(function(e) { var em = e.trim(); if (em) toEmails.push(em); });
      }
    });
    
    // 重複排除
    toEmails = toEmails.filter(function(e, i, a) { return e && a.indexOf(e) === i; });
    if (!toEmails.length) return;
    
    // 件名テンプレート適用
    var subjectTpl = templates[eventKey.split('_')[0]] || '【見積管理】' + eventKey;
    var subject = subjectTpl
      .replace('{orderNo}',  data.orderNo  || '')
      .replace('{quoteNo}',  data.quoteNo  || '')
      .replace('{client}',   data.client   || '')
      .replace('{amount}',   (data.amount ? '¥' + Number(data.amount).toLocaleString() : ''))
      .replace('{assignee}', data.assignee || '')
      .replace('{date}',     nowJST().substring(0, 10));
    
    // 本文
    var body = buildMailBody(eventKey, data);
    
    // 送信
    toEmails.forEach(function(email) {
      try {
        GmailApp.sendEmail(email, subject, body, { name: '見積・注文管理システム' });
        Logger.log('[MAIL] 送信: ' + email + ' | ' + subject);
      } catch(e2) {
        Logger.log('[MAIL] 送信失敗: ' + email + ' | ' + e2.message);
      }
    });
    
  } catch(e) {
    Logger.log('[MAIL_EVENT] エラー: ' + e.message);
  }
}

function buildMailBody(eventKey, data) {
  var lines = [
    '自動送信メッセージ（見積・注文管理システム）',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    '',
  ];
  
  if (data.orderNo)  lines.push('注文書番号: ' + data.orderNo);
  if (data.quoteNo)  lines.push('見積書番号: ' + data.quoteNo);
  if (data.client)   lines.push('顧客名:     ' + data.client);
  if (data.amount)   lines.push('金額:       ¥' + Number(data.amount).toLocaleString());
  if (data.assignee) lines.push('担当者:     ' + data.assignee);
  if (data.status)   lines.push('ステータス: ' + data.status);
  if (data.reason)   lines.push('理由:       ' + data.reason);
  if (data.pdfUrl)   lines.push('PDFリンク:  ' + data.pdfUrl);
  
  lines.push('');
  lines.push('━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  // システムURL
  var webUrl = ScriptApp.getService().getUrl();
  if (webUrl) lines.push('システムURL: ' + webUrl);
  
  return lines.join('\n');
}

function _apiSendTestMailAll(p) {
  try {
    var settingsRaw = PropertiesService.getScriptProperties().getProperty(MAIL_SEND_KEY);
    if (!settingsRaw) return { success: false, error: 'メール配信設定が未保存です' };
    var settings = JSON.parse(settingsRaw);
    var groups = settings.groups || {};
    var allEmails = [];
    Object.values(groups).forEach(function(gl) {
      (gl||'').split(',').forEach(function(e) { var em = e.trim(); if (em) allEmails.push(em); });
    });
    allEmails = allEmails.filter(function(e, i, a) { return e && a.indexOf(e) === i; });
    
    if (!allEmails.length) return { success: false, error: '配信先グループにメールアドレスが設定されていません' };
    
    var subject = '【テスト】見積・注文管理システム メール配信テスト';
    var body = ['これはテストメールです。', '見積・注文管理システムから自動送信されました。', '', '送信日時: ' + nowJST()].join('\n');
    
    var sent = 0;
    allEmails.forEach(function(email) {
      try {
        GmailApp.sendEmail(email, subject, body, { name: '見積・注文管理システム' });
        sent++;
      } catch(e2) {}
    });
    return { success: true, sentCount: sent };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// ③ 楽楽販売 連携
// ============================================================

var RAKU_SETTINGS_KEY = 'RAKURAKU_SETTINGS';
var RAKU_LOG_KEY      = 'RAKURAKU_SYNC_LOG';

function _apiGetRakurakuSettings() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(RAKU_SETTINGS_KEY);
    var settings = raw ? JSON.parse(raw) : {};
    
    // Webhook URL（このGASのURL）
    var webhookUrl = ScriptApp.getService().getUrl() + '?action=rakurakuWebhook';
    settings.webhookUrl = webhookUrl;
    
    // 同期ログ
    var logRaw = PropertiesService.getScriptProperties().getProperty(RAKU_LOG_KEY);
    settings.syncLog = logRaw ? JSON.parse(logRaw) : ['同期ログなし'];
    
    // APIキーはマスク
    if (settings.apiKey) settings.apiKey = '（設定済み）';
    
    return { success: true, settings: settings };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _apiSaveRakurakuSettings(p) {
  try {
    var props = PropertiesService.getScriptProperties();
    var existing = JSON.parse(props.getProperty(RAKU_SETTINGS_KEY) || '{}');
    var toSave = Object.assign(existing, {
      mode:       p.mode || 'api',
      tenant:     p.tenant || existing.tenant,
      syncStatus: p.syncStatus,
      interval:   p.interval,
      mapping:    p.mapping || existing.mapping,
    });
    // APIキーは「（設定済み）」以外のときだけ更新
    if (p.apiKey && p.apiKey !== '（設定済み）') {
      toSave.apiKey = p.apiKey;
    }
    props.setProperty(RAKU_SETTINGS_KEY, JSON.stringify(toSave));
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _apiTestRakurakuConnection(p) {
  try {
    var props  = PropertiesService.getScriptProperties();
    var saved  = JSON.parse(props.getProperty(RAKU_SETTINGS_KEY) || '{}');
    var tenant = p.tenant || saved.tenant;
    var apiKey = (p.apiKey && p.apiKey !== '（設定済み）') ? p.apiKey : saved.apiKey;
    
    if (!tenant || !apiKey) return { success: false, error: 'テナントIDとAPIキーが必要です' };
    
    // 楽楽販売 REST API テスト（見積書一覧を1件だけ取得）
    var url = 'https://' + tenant + '.raku-raku.jp/api/v1/estimates?limit=1';
    var res = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + apiKey, 'Content-Type': 'application/json' },
      muteHttpExceptions: true,
    });
    
    if (res.getResponseCode() === 200) {
      var data = JSON.parse(res.getContentText());
      return { success: true, count: (data.total || data.data?.length || 0) };
    } else if (res.getResponseCode() === 401) {
      return { success: false, error: 'APIキーが無効です（401）' };
    } else if (res.getResponseCode() === 404) {
      return { success: false, error: 'テナントIDが見つかりません（404）。サブドメインを確認してください' };
    } else {
      return { success: false, error: 'HTTPエラー: ' + res.getResponseCode() + ' / ' + res.getContentText().substring(0,200) };
    }
  } catch(e) {
    return { success: false, error: e.message };
  }
}

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

/** CSV取込（楽楽販売エクスポートCSV） */
function _apiImportRakurakuCsv(p) {
  try {
    if (!p.csvData || !p.headers) return { success: false, error: 'CSVデータがありません' };
    
    var lines = p.csvData.split('\n').filter(function(l) { return l.trim(); });
    if (lines.length < 2) return { success: false, error: 'データ行がありません' };
    
    var mapping = p.mapping || {};
    var qNoField = mapping.quoteno || '見積番号';
    
    // ヘッダーからフィールドインデックスを特定
    var csvHeaders = p.headers;
    var getIdx = function(fieldName) { return csvHeaders.indexOf(fieldName); };
    
    var qNoIdx    = getIdx(qNoField);
    var clientIdx = getIdx(mapping.client  || '取引先名');
    var amountIdx = getIdx(mapping.amount  || '見積金額');
    var dateIdx   = getIdx(mapping.date    || '見積日');
    var statusIdx = getIdx(mapping.status  || 'ステータス');
    
    var ss        = getSpreadsheet();
    var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var mgmtData  = mgmtSheet.getDataRange().getValues();
    var existingQuoteNos = mgmtData.slice(1).map(function(r) { return String(r[MGMT_COLS.QUOTE_NO - 1]||'').trim(); });
    
    var imported = 0;
    var now = nowJST();
    
    lines.slice(1).forEach(function(line) {
      var cells = line.split(',').map(function(c) { return c.trim().replace(/^"|"$/g, ''); });
      var quoteNo = qNoIdx >= 0 ? cells[qNoIdx] : '';
      if (!quoteNo || existingQuoteNos.indexOf(quoteNo) >= 0) return;
      
      var mgmtId  = generateMgmtId();
      var client  = clientIdx  >= 0 ? cells[clientIdx]  : '';
      var amount  = amountIdx  >= 0 ? Number(String(cells[amountIdx]).replace(/[,，]/g, '')) : 0;
      var issDate = dateIdx    >= 0 ? cells[dateIdx]    : '';
      
      var row = new Array(27).fill('');
      row[MGMT_COLS.ID          - 1] = mgmtId;
      row[MGMT_COLS.QUOTE_NO    - 1] = quoteNo;
      row[MGMT_COLS.CLIENT      - 1] = client;
      row[MGMT_COLS.STATUS      - 1] = CONFIG.STATUS.SENT;
      row[MGMT_COLS.QUOTE_DATE  - 1] = issDate;
      row[MGMT_COLS.QUOTE_AMOUNT- 1] = amount;
      row[MGMT_COLS.MEMO        - 1] = '楽楽販売CSV取込 (' + now + ')';
      row[MGMT_COLS.CREATED_AT  - 1] = now;
      row[MGMT_COLS.UPDATED_AT  - 1] = now;
      
      mgmtSheet.appendRow(row);
      existingQuoteNos.push(quoteNo);
      imported++;
    });
    
    return { success: true, imported: imported };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

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
