// ============================================================
// 見積書・注文書管理システム
// ファイル 17: Gemini APIキー複数管理・レート制限対策
// ============================================================
//
// 【このファイルの役割】
//   Gemini APIの無料枠（15req/分, 1500req/日）を使い切った場合でも
//   複数のAPIキーをローテーションして処理を継続させる。
//
// 【セットアップ手順】
//   1. Google AI Studio (https://aistudio.google.com/) で複数の
//      Googleアカウントを使ってAPIキーを発行する（無料枠でOK）
//
//   2. GASエディタのスクリプトプロパティに以下を登録：
//      GEMINI_API_KEY_1 = AIzaSy... （メインキー）
//      GEMINI_API_KEY_2 = AIzaSy... （サブキー1）
//      GEMINI_API_KEY_3 = AIzaSy... （サブキー3 ※任意）
//      GEMINI_API_KEY   = （空でもOK。自動で上記から選ばれる）
//
//   3. このファイルを追加するだけで既存コードに自動適用される。
//      （_callGeminiApi をオーバーライドするため）
//
// 【動作の仕組み】
//   ① キー1でリクエスト → 429エラーなら自動でキー2に切り替え
//   ② キー2も429なら → キー3へ、それも駄目なら60秒待機してリトライ
//   ③ 使用回数をスクリプトプロパティに記録して日次リセット
//   ④ 残り使用可能回数をダッシュボードに表示
//
// 【無料枠の目安】
//   gemini-1.5-flash: 15リクエスト/分、1,500リクエスト/日
//   キーを3本使えば → 45リクエスト/分、4,500リクエスト/日
// ============================================================

// ============================================================
// APIキー管理設定
// ============================================================

var API_KEY_MANAGER_CONFIG = {
  // スクリプトプロパティのキー名（複数登録可）
  KEY_PROP_NAMES: ['GEMINI_API_KEY_1', 'GEMINI_API_KEY_2', 'GEMINI_API_KEY_3', 'GEMINI_API_KEY'],

  // 使用回数カウンタのキー
  USAGE_COUNT_PREFIX: 'GEMINI_USAGE_',
  USAGE_DATE_KEY:     'GEMINI_USAGE_DATE',

  // 無料枠の上限（保守的に設定）
  DAILY_LIMIT_PER_KEY:  1400,   // 1日1500回の上限より少し低く設定
  MINUTE_LIMIT_PER_KEY: 14,     // 1分15回の上限より少し低く設定

  // レート制限時の待機時間
  WAIT_ON_429_MS:       65000,  // 429エラー時：65秒待機（1分ルールより余裕を持って）
  WAIT_BETWEEN_KEYS_MS: 1000,   // キー切り替え時のウェイト

  // リトライ設定
  MAX_RETRIES_PER_KEY: 2,
};

// ============================================================
// メインの _callGeminiApi を上書き（02_ocr_and_processing.gs より後に読み込まれる想定）
// GASは同名関数が複数あると後に定義されたものが優先される
// ============================================================

/**
 * Gemini APIを呼び出す（複数キー対応・レート制限自動回避版）
 * 02_ocr_and_processing.gs の同名関数をオーバーライドする
 */
function _callGeminiApi(model, body) {
  var keys = _getAvailableApiKeys();

  if (keys.length === 0) {
    Logger.log('[API_MANAGER] 利用可能なAPIキーがありません。スクリプトプロパティを確認してください。');
    return null;
  }

  // 各キーで順番に試す
  for (var keyIdx = 0; keyIdx < keys.length; keyIdx++) {
    var apiKey = keys[keyIdx];
    var keyName = 'KEY_' + (keyIdx + 1);

    // 日次使用制限チェック
    if (_isKeyExhaustedToday(keyIdx)) {
      Logger.log('[API_MANAGER] ' + keyName + ' は本日の上限に達しています。次のキーへ。');
      continue;
    }

    var url = CONFIG.GEMINI_API_ENDPOINT + model + ':generateContent?key=' + apiKey;

    for (var retry = 0; retry < API_KEY_MANAGER_CONFIG.MAX_RETRIES_PER_KEY; retry++) {
      try {
        var res = UrlFetchApp.fetch(url, {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify(body),
          muteHttpExceptions: true,
        });

        var code = res.getResponseCode();

        if (code === 200) {
          // 成功：使用回数を記録して結果を返す
          _incrementUsageCount(keyIdx);
          Logger.log('[API_MANAGER] ' + keyName + ' 成功 (model: ' + model + ', 本日使用: ' + _getTodayUsage(keyIdx) + '回)');
          return JSON.parse(res.getContentText());
        }

        if (code === 429) {
          // レート制限：このキーを一時停止して次のキーへ
          Logger.log('[API_MANAGER] ' + keyName + ' レート制限(429)。次のキーへ切り替えます。');
          _markKeyRateLimited(keyIdx);
          Utilities.sleep(API_KEY_MANAGER_CONFIG.WAIT_BETWEEN_KEYS_MS);
          break; // 内側のリトライループを抜けて次のキーへ
        }

        if (code === 401 || code === 403) {
          // 認証エラー：このキーをスキップ
          Logger.log('[API_MANAGER] ' + keyName + ' 認証エラー(' + code + ')。このキーをスキップします。');
          break;
        }

        if (code === 503 || code === 500) {
          // サーバーエラー：少し待ってリトライ
          Logger.log('[API_MANAGER] ' + keyName + ' サーバーエラー(' + code + ')。' + (retry + 1) + '回目リトライ...');
          Utilities.sleep(5000 * (retry + 1));
          continue;
        }

        // その他のエラー
        Logger.log('[API_MANAGER] ' + keyName + ' エラー(' + code + '): ' + res.getContentText().substring(0, 200));
        break;

      } catch (e) {
        Logger.log('[API_MANAGER] ' + keyName + ' 例外: ' + e.message);
        if (retry < API_KEY_MANAGER_CONFIG.MAX_RETRIES_PER_KEY - 1) {
          Utilities.sleep(3000);
        }
      }
    }
  }

  // 全キーが失敗した場合：全キーがレート制限中なら待機して1回だけ再試行
  if (_allKeysRateLimited(keys.length)) {
    Logger.log('[API_MANAGER] 全キーがレート制限中。' + (API_KEY_MANAGER_CONFIG.WAIT_ON_429_MS / 1000) + '秒待機後にリトライします...');
    Utilities.sleep(API_KEY_MANAGER_CONFIG.WAIT_ON_429_MS);
    _clearRateLimitFlags();

    // 待機後に最初のキーで1回だけ再試行
    var fallbackKey = keys[0];
    var fallbackUrl = CONFIG.GEMINI_API_ENDPOINT + model + ':generateContent?key=' + fallbackKey;
    try {
      var fallbackRes = UrlFetchApp.fetch(fallbackUrl, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(body),
        muteHttpExceptions: true,
      });
      if (fallbackRes.getResponseCode() === 200) {
        _incrementUsageCount(0);
        Logger.log('[API_MANAGER] 待機後リトライ成功');
        return JSON.parse(fallbackRes.getContentText());
      }
    } catch (e) {
      Logger.log('[API_MANAGER] 待機後リトライも失敗: ' + e.message);
    }
  }

  Logger.log('[API_MANAGER] 全キーで失敗しました。処理をスキップします。');
  return null;
}

// ============================================================
// APIキー管理ユーティリティ
// ============================================================

/**
 * 利用可能なAPIキーリストを取得（重複除去・空除去）
 */
function _getAvailableApiKeys() {
  var props = PropertiesService.getScriptProperties();
  var keys = [];
  var seen = {};

  API_KEY_MANAGER_CONFIG.KEY_PROP_NAMES.forEach(function(propName) {
    var key = (props.getProperty(propName) || '').trim();
    if (key && key.length > 10 && !seen[key]) {
      seen[key] = true;
      keys.push(key);
    }
  });

  return keys;
}

/**
 * 今日の日付文字列を取得（日次リセット用）
 */
function _getTodayStr() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
}

/**
 * 今日の使用回数を取得
 */
function _getTodayUsage(keyIdx) {
  var props = PropertiesService.getScriptProperties();
  var today = _getTodayStr();
  var storedDate = props.getProperty(API_KEY_MANAGER_CONFIG.USAGE_DATE_KEY) || '';

  // 日付が変わっていたらリセット
  if (storedDate !== today) {
    _resetAllUsageCounts();
    return 0;
  }

  var count = props.getProperty(API_KEY_MANAGER_CONFIG.USAGE_COUNT_PREFIX + keyIdx);
  return count ? parseInt(count) : 0;
}

/**
 * 使用回数を1増加
 */
function _incrementUsageCount(keyIdx) {
  try {
    var props = PropertiesService.getScriptProperties();
    var today = _getTodayStr();
    props.setProperty(API_KEY_MANAGER_CONFIG.USAGE_DATE_KEY, today);
    var current = _getTodayUsage(keyIdx);
    props.setProperty(API_KEY_MANAGER_CONFIG.USAGE_COUNT_PREFIX + keyIdx, String(current + 1));
  } catch(e) {
    Logger.log('[API_MANAGER] 使用回数記録エラー: ' + e.message);
  }
}

/**
 * 今日の使用制限に達しているか確認
 */
function _isKeyExhaustedToday(keyIdx) {
  return _getTodayUsage(keyIdx) >= API_KEY_MANAGER_CONFIG.DAILY_LIMIT_PER_KEY;
}

/**
 * キーをレート制限中としてマーク（5分間）
 */
function _markKeyRateLimited(keyIdx) {
  try {
    var cache = CacheService.getScriptCache();
    cache.put('RATE_LIMITED_KEY_' + keyIdx, '1', 300); // 5分間キャッシュ
  } catch(e) {}
}

/**
 * 全キーがレート制限中か確認
 */
function _allKeysRateLimited(keyCount) {
  try {
    var cache = CacheService.getScriptCache();
    for (var i = 0; i < keyCount; i++) {
      if (!cache.get('RATE_LIMITED_KEY_' + i)) return false;
    }
    return true;
  } catch(e) {
    return false;
  }
}

/**
 * レート制限フラグをクリア（待機後）
 */
function _clearRateLimitFlags() {
  try {
    var cache = CacheService.getScriptCache();
    var keys = [];
    for (var i = 0; i < 5; i++) keys.push('RATE_LIMITED_KEY_' + i);
    cache.removeAll(keys);
  } catch(e) {}
}

/**
 * 全キーの使用回数をリセット（日次）
 */
function _resetAllUsageCounts() {
  try {
    var props = PropertiesService.getScriptProperties();
    for (var i = 0; i < 5; i++) {
      props.deleteProperty(API_KEY_MANAGER_CONFIG.USAGE_COUNT_PREFIX + i);
    }
    props.setProperty(API_KEY_MANAGER_CONFIG.USAGE_DATE_KEY, _getTodayStr());
    Logger.log('[API_MANAGER] 使用回数をリセットしました（日次）');
  } catch(e) {}
}

// ============================================================
// API使用状況レポート（管理コンソール用）
// ============================================================

/**
 * 全キーの使用状況をオブジェクトで返す
 * 03_webapp_ap.gs の handleApiRequest から 'getApiUsage' アクションで呼ぶ
 */
function getApiUsageReport() {
  var keys = _getAvailableApiKeys();
  var today = _getTodayStr();
  var storedDate = PropertiesService.getScriptProperties().getProperty(API_KEY_MANAGER_CONFIG.USAGE_DATE_KEY) || '';
  var isToday = (storedDate === today);

  var report = keys.map(function(key, idx) {
    var usage = isToday ? _getTodayUsage(idx) : 0;
    var limited = false;
    try {
      limited = !!CacheService.getScriptCache().get('RATE_LIMITED_KEY_' + idx);
    } catch(e) {}
    return {
      keyIndex:   idx + 1,
      keyMasked:  key.substring(0, 8) + '...' + key.substring(key.length - 4),
      usedToday:  usage,
      dailyLimit: API_KEY_MANAGER_CONFIG.DAILY_LIMIT_PER_KEY,
      remaining:  Math.max(0, API_KEY_MANAGER_CONFIG.DAILY_LIMIT_PER_KEY - usage),
      percent:    Math.min(100, Math.round(usage / API_KEY_MANAGER_CONFIG.DAILY_LIMIT_PER_KEY * 100)),
      rateLimited: limited,
      status:     limited ? '⚠️ 一時制限中' : (usage >= API_KEY_MANAGER_CONFIG.DAILY_LIMIT_PER_KEY ? '❌ 本日上限' : '✅ 利用可能'),
    };
  });

  var totalUsed      = report.reduce(function(s, r) { return s + r.usedToday; }, 0);
  var totalRemaining = report.reduce(function(s, r) { return s + r.remaining; }, 0);

  return {
    success:        true,
    keyCount:       keys.length,
    keys:           report,
    totalUsedToday: totalUsed,
    totalRemaining: totalRemaining,
    resetAt:        '毎日0:00（日本時間）',
    tips: keys.length === 1
      ? '💡 APIキーを複数登録するとレート制限を回避しやすくなります（GEMINI_API_KEY_1, _2, _3）'
      : '',
  };
}

// ============================================================
// 手動操作用ユーティリティ（GASエディタから実行）
// ============================================================

/**
 * 現在のAPI使用状況をログに出力（手動確認用）
 */
function checkApiUsageStatus() {
  var report = getApiUsageReport();
  Logger.log('=== Gemini API 使用状況 ===');
  Logger.log('登録キー数: ' + report.keyCount + '本');
  Logger.log('本日合計使用: ' + report.totalUsedToday + '回');
  Logger.log('残り合計: ' + report.totalRemaining + '回');
  report.keys.forEach(function(k) {
    Logger.log('[KEY ' + k.keyIndex + '] ' + k.keyMasked +
      ' → 使用: ' + k.usedToday + '/' + k.dailyLimit +
      ' (' + k.percent + '%) ' + k.status);
  });
}

/**
 * APIキーが正しく設定されているか一括テスト（手動実行用）
 */
function testAllApiKeys() {
  var keys = _getAvailableApiKeys();
  Logger.log('=== APIキー一括テスト（' + keys.length + '本）===');

  if (keys.length === 0) {
    Logger.log('❌ APIキーが1本も設定されていません。');
    Logger.log('スクリプトプロパティに GEMINI_API_KEY_1 を設定してください。');
    return;
  }

  var model = CONFIG.GEMINI_PRIMARY_MODEL || 'gemini-1.5-flash';
  keys.forEach(function(key, idx) {
    var url = CONFIG.GEMINI_API_ENDPOINT + model + ':generateContent?key=' + key;
    try {
      var res = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          contents: [{ parts: [{ text: 'OK' }] }],
          generationConfig: { maxOutputTokens: 5, temperature: 0 },
        }),
        muteHttpExceptions: true,
      });
      var code = res.getResponseCode();
      var masked = key.substring(0, 8) + '...' + key.substring(key.length - 4);
      if (code === 200) {
        Logger.log('✅ KEY_' + (idx + 1) + ' [' + masked + '] 接続OK');
      } else if (code === 429) {
        Logger.log('⚠️ KEY_' + (idx + 1) + ' [' + masked + '] レート制限中（キーは有効）');
      } else if (code === 401 || code === 403) {
        Logger.log('❌ KEY_' + (idx + 1) + ' [' + masked + '] 認証失敗（無効なキー）');
      } else {
        Logger.log('❓ KEY_' + (idx + 1) + ' [' + masked + '] HTTP ' + code);
      }
    } catch(e) {
      Logger.log('❌ KEY_' + (idx + 1) + ' 例外: ' + e.message);
    }
    Utilities.sleep(500);
  });
}

/**
 * 使用回数カウンターを手動リセット（月次メンテ用）
 */
function manualResetApiUsage() {
  _resetAllUsageCounts();
  _clearRateLimitFlags();
  Logger.log('✅ API使用回数とレート制限フラグをリセットしました。');
}