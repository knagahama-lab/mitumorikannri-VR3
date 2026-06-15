// ============================================================
// 見積書・注文書管理システム
// このファイルは設定・ユーティリティ専用（ルーターは 06 board api.gs）
// ルーター（doGet / handleApiRequest）は 06 board api.gs にのみ存在する。
// ここに定義するのは 06 board api.gs に存在しない固有の関数のみ。
// ============================================================

// ============================================================
// ★ _rowToObject — 管理シート1行 → オブジェクト変換
// ============================================================
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

// ============================================================
// ★ 管理コンソール設定 API
// ============================================================
function _apiLoadSettings() {
  try {
    var props = PropertiesService.getScriptProperties();
    var raw   = props.getProperty('SYS_SETTINGS');
    var s     = raw ? JSON.parse(raw) : {};
    var geminiKey     = props.getProperty('GEMINI_API_KEY') || '';
    var geminiKeyIsSet = !!geminiKey;
    var geminiKeyHint  = geminiKey ? geminiKey.substring(0, 8) + '...' : '';
    return {
      success:         true,
      settings:        s,
      adminEmails:     props.getProperty('ADMIN_EMAILS')      || '',
      notifyEmails:    props.getProperty('NOTIFY_EMAILS')     || '',
      salesEmails:     props.getProperty('SALES_EMAILS')      || '',
      procurementEmail:props.getProperty('PROCUREMENT_EMAIL') || '',
      approvalButtonEnabled: props.getProperty('APPROVAL_BUTTON_ENABLED') !== 'false',
      spreadsheetId:   props.getProperty('SPREADSHEET_ID')    || '',
      chatWebhook:     props.getProperty('CHAT_WEBHOOK_URL')  || s.webhookUrl || '',
      rakurakuCompany: props.getProperty('RAKURAKU_COMPANY')  || '',
      rakurakuEndpoint:props.getProperty('RAKURAKU_ENDPOINT') || '',
      n8nOrderWebhook: props.getProperty('N8N_ORDER_WEBHOOK') || '',
      n8nQuoteWebhook: props.getProperty('N8N_QUOTE_WEBHOOK') || '',
      customWebhook:   props.getProperty('CUSTOM_WEBHOOK')    || '',
      geminiKeyIsSet:  geminiKeyIsSet,
      geminiKeyHint:   geminiKeyHint,
    };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _apiSaveSettings(p) {
  try {
    var props = PropertiesService.getScriptProperties();
    if (p.adminEmails        !== undefined) props.setProperty('ADMIN_EMAILS',        String(p.adminEmails));
    if (p.notifyEmails       !== undefined) props.setProperty('NOTIFY_EMAILS',       String(p.notifyEmails));
    if (p.salesEmails        !== undefined) props.setProperty('SALES_EMAILS',        String(p.salesEmails));
    if (p.procurementEmail   !== undefined) props.setProperty('PROCUREMENT_EMAIL',   String(p.procurementEmail));
    if (p.approvalButtonEnabled !== undefined) props.setProperty('APPROVAL_BUTTON_ENABLED', p.approvalButtonEnabled ? 'true' : 'false');
    if (p.spreadsheetId   !== undefined && p.spreadsheetId) props.setProperty('SPREADSHEET_ID', String(p.spreadsheetId));
    if (p.chatWebhook     !== undefined) props.setProperty('CHAT_WEBHOOK_URL',   String(p.chatWebhook));
    if (p.geminiKey       !== undefined && p.geminiKey) props.setProperty('GEMINI_API_KEY',   String(p.geminiKey));
    if (p.rakurakuCompany !== undefined) props.setProperty('RAKURAKU_COMPANY',   String(p.rakurakuCompany));
    if (p.rakurakuEndpoint!== undefined) props.setProperty('RAKURAKU_ENDPOINT',  String(p.rakurakuEndpoint));
    if (p.n8nOrderWebhook !== undefined) props.setProperty('N8N_ORDER_WEBHOOK',  String(p.n8nOrderWebhook));
    if (p.n8nQuoteWebhook !== undefined) props.setProperty('N8N_QUOTE_WEBHOOK',  String(p.n8nQuoteWebhook));
    if (p.customWebhook   !== undefined) props.setProperty('CUSTOM_WEBHOOK',     String(p.customWebhook));
    var raw = props.getProperty('SYS_SETTINGS');
    var s   = raw ? JSON.parse(raw) : {};
    if (p.notifyOrder  !== undefined) s.notifyOrder = p.notifyOrder;
    if (p.notifyQuote  !== undefined) s.notifyQuote  = p.notifyQuote;
    if (p.notifyDl     !== undefined) s.notifyDl     = p.notifyDl;
    if (p.alertDays    !== undefined) s.alertDays    = p.alertDays;
    if (p.webhookUrl   !== undefined) s.webhookUrl   = p.webhookUrl || p.chatWebhook;
    if (p.deliveryRemindDays !== undefined) s.deliveryRemindDays = Number(p.deliveryRemindDays) || 7;
    if (p.deliveryUrgentDays !== undefined) s.deliveryUrgentDays = Number(p.deliveryUrgentDays) || 1;
    if (p.deadlineRemindDays !== undefined) s.deadlineRemindDays = Number(p.deadlineRemindDays) || 7;
    if (p.deadlineUrgentDays !== undefined) s.deadlineUrgentDays = Number(p.deadlineUrgentDays) || 1;
    if (p.alertDelivery      !== undefined) s.alertDelivery      = !!p.alertDelivery;
    if (p.alertDeadline      !== undefined) s.alertDeadline      = !!p.alertDeadline;
    if (p.alertOverdue       !== undefined) s.alertOverdue       = !!p.alertOverdue;
    if (p.alertUnlinked      !== undefined) s.alertUnlinked      = !!p.alertUnlinked;
    if (p.stagnantWarnDays  !== undefined) s.stagnantWarnDays  = Number(p.stagnantWarnDays)  || 7;
    if (p.stagnantAlertDays !== undefined) s.stagnantAlertDays = Number(p.stagnantAlertDays) || 14;
    if (p.unlinkedOrderDays !== undefined) s.unlinkedOrderDays = Number(p.unlinkedOrderDays) || 2;
    if (p.ocrFailThreshold  !== undefined) s.ocrFailThreshold  = Number(p.ocrFailThreshold)  || 3;
    props.setProperty('SYS_SETTINGS', JSON.stringify(s));
    try { _applyMonitorConfig(s); } catch(e2) {}
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _applyMonitorConfig(s) {
  if (typeof MONITOR_CONFIG === 'undefined') return;
  if (s.deliveryRemindDays) MONITOR_CONFIG.DELIVERY_REMIND_DAYS  = s.deliveryRemindDays;
  if (s.deliveryUrgentDays) MONITOR_CONFIG.DELIVERY_URGENT_DAYS  = s.deliveryUrgentDays;
  if (s.stagnantWarnDays)   MONITOR_CONFIG.STAGNANT_WARN_DAYS    = s.stagnantWarnDays;
  if (s.stagnantAlertDays)  MONITOR_CONFIG.STAGNANT_ALERT_DAYS   = s.stagnantAlertDays;
  if (s.unlinkedOrderDays)  MONITOR_CONFIG.UNLINKED_ORDER_DAYS   = s.unlinkedOrderDays;
  if (s.ocrFailThreshold)   MONITOR_CONFIG.OCR_FAIL_THRESHOLD    = s.ocrFailThreshold;
}

function _apiSendTestAlert() {
  try {
    var props       = PropertiesService.getScriptProperties();
    var notifyEmails = props.getProperty('NOTIFY_EMAILS') || '';
    var webhookUrl   = props.getProperty('CHAT_WEBHOOK_URL') || props.getProperty('GOOGLE_CHAT_WEBHOOK_URL') || '';
    var msg = '【✅ 見積管理システム テスト通知】\n送信日時: ' + nowJST() + '\n通知設定が正常に動作しています。';
    var sent = false;
    if (notifyEmails) {
      var emails = notifyEmails.split(',').map(function(e){ return e.trim(); }).filter(Boolean);
      emails.forEach(function(email) {
        try { GmailApp.sendEmail(email, '【見積管理システム】テスト通知', msg); sent = true; } catch(e2) {}
      });
    }
    if (webhookUrl) {
      try {
        UrlFetchApp.fetch(webhookUrl, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ text: msg }), muteHttpExceptions: true });
        sent = true;
      } catch(e3) {}
    }
    if (!sent) return { success: false, error: '通知先が設定されていません' };
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function _apiGetOcrUsage() {
  try {
    if (typeof getOcrUsageInfo === 'function') return getOcrUsageInfo();
    var raw = PropertiesService.getScriptProperties().getProperty('OCR_USAGE_LOG');
    var log = raw ? JSON.parse(raw) : [];
    return { success: true, total: log.length, items: log.slice(0, 50) };
  } catch(e) { return { success: false, error: e.message }; }
}
