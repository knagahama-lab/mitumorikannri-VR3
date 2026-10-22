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