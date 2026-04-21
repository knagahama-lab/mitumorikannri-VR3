// ============================================================
// ファイル: 14_ocr_retry_batch.gs
// 目的: 夜間バッチによるOCR情報の自己修復（自己補完）機能
// ============================================================

const OCR_RETRY_CONFIG = {
  SHEET_NAME: '案件一覧', // 対象シート名（仮定情報に基づく）
  MODEL: 'gemini-3.1-flash-lite',
  SLEEP_MS: 5000,         // 15回/分（RPM）の制限回避のため5秒待機
  MAX_EXEC_SECONDS: 270,  // GASタイムアウトエラー回避（開始から4.5分=270秒で強制終了）
  MAX_RETRY: 3            // 最大リトライ回数
};

/**
 * 【Time-driven Trigger】夜間に定期実行されるバッチのメイン関数
 */
function runOcrRetryBatch() {
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(OCR_RETRY_CONFIG.SHEET_NAME);
  
  if (!sheet) {
    Logger.log('[OCR RETRY] シート「' + OCR_RETRY_CONFIG.SHEET_NAME + '」が見つかりません。作成または名前を確認してください。');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // データなし

  // 全データ取得 (A~I列: 1~9)
  var range = sheet.getRange(1, 1, lastRow, 9);
  var values = range.getValues();
  
  // 1行目のヘッダー情報を取得（カラム名の特定用）
  var headers = values[0];

  // OCR対象となる情報列(C〜G列：インデックス2〜6)
  var targetColIndexes = [2, 3, 4, 5, 6];

  for (var i = 1; i < values.length; i++) {
    // タイムアウト検証（もし処理開始から4分30秒経過していたらループを抜ける）
    var elapsedSeconds = (new Date().getTime() - startTime) / 1000;
    if (elapsedSeconds > OCR_RETRY_CONFIG.MAX_EXEC_SECONDS) {
      Logger.log('[OCR RETRY] タイムアウト回避規定時間に達したため、処理を中断・次回に持ち越します。完了行: ' + i);
      break;
    }

    var rowId    = values[i][0]; // A列: 案件ID
    var fileId   = values[i][1]; // B列: ドライブ上のファイルID
    var status   = values[i][7]; // H列: OCRステータス
    var retryCnt = Number(values[i][8]) || 0; // I列: リトライ回数

    // 抽出条件：ステータスが「補完待ち」かつ リトライ回数が3未満
    if (status === '補完待ち' && retryCnt < OCR_RETRY_CONFIG.MAX_RETRY) {
      
      // 不足している項目（空文字の列）のカラム名を特定
      var missingHeaders = [];
      var missingIndexes = [];
      targetColIndexes.forEach(function(colIdx) {
        if (!values[i][colIdx] || String(values[i][colIdx]).trim() === '') {
          missingHeaders.push(headers[colIdx]);
          missingIndexes.push(colIdx);
        }
      });

      // もしすべて埋まっている（不足なし）場合はステータスを完了にしてスキップ
      if (missingHeaders.length === 0) {
        sheet.getRange(i + 1, 8).setValue('完了'); // H列
        continue;
      }

      Logger.log('[OCR RETRY] 対象ID: ' + rowId + ' 不足項目: ' + missingHeaders.join(', '));

      try {
        // ドライブからファイル情報を取得・Base64化
        var file = DriveApp.getFileById(fileId);
        var mimeType = file.getMimeType() || 'application/pdf';
        var base64 = Utilities.base64Encode(file.getBlob().getBytes());

        // プロンプト生成（不足しているカラム名リストを直接指定して厳格な推測を依頼）
        var prompt = 'この画像は見積書です。前回の読み取りで以下の項目が取得できませんでした。\n' + 
                     '画像から [' + missingHeaders.join(', ') + '] のみを推測・抽出し、厳格なJSON形式で返してください。\n' +
                     'JSONキーには、こちらの指定した日本語名をそのまま利用してください。値が取得できない場合は空文字としてください。\n' +
                     'ルール: 有効なJSONのみ。マークダウンや説明は一切記載しないこと。';

        var body = {
          contents: [{ parts: [
            { text: prompt },
            { inline_data: { mime_type: mimeType, data: base64 } }
          ]}],
          generationConfig: { temperature: 0.1, responseMimeType: 'application/json' }
        };

        // API実行（gemini-3.1-flash-lite）
        var apiRes = _callGeminiApiOcrRetry(OCR_RETRY_CONFIG.MODEL, body);
        
        // 取得した結果をパース
        var extractedData = {};
        if (apiRes) extractedData = _parseGeminiJsonResponse(apiRes);

        // 取得データの反映検証と更新
        var isAllFilled = true;
        missingIndexes.forEach(function(colIdx) {
          var hName = headers[colIdx];
          var extVal = extractedData[hName];
          if (extVal && String(extVal).trim() !== '') {
            // スプレッドシートへ書き込み
            sheet.getRange(i + 1, colIdx + 1).setValue(extVal);
          } else {
             // 依然として取得できない項目がある
             isAllFilled = false;
          }
        });

        // リトライ回数 + 1
        var nextRetry = retryCnt + 1;
        sheet.getRange(i + 1, 9).setValue(nextRetry); // I列を更新

        // ステータス判断
        if (isAllFilled) {
          sheet.getRange(i + 1, 8).setValue('完了');
        } else if (nextRetry >= OCR_RETRY_CONFIG.MAX_RETRY) {
          sheet.getRange(i + 1, 8).setValue('手動確認');
        }

      } catch(e) {
        Logger.log('[OCR RETRY ERROR] ' + rowId + ': ' + e.message);
      }
      
      // ★ 最重要：RPM(15回/分)制限回避。API呼び出しの有無に関わらず安全のため待機
      Utilities.sleep(OCR_RETRY_CONFIG.SLEEP_MS);
    }
  }
  
  Logger.log('[OCR RETRY BATCH] 処理完了');
}

/**
 * Gemini APIへの直接リクエスト関数（バッチ処理専用）
 */
function _callGeminiApiOcrRetry(model, body) {
  var key = typeof CONFIG !== 'undefined' && CONFIG.GEMINI_API_KEY 
            ? CONFIG.GEMINI_API_KEY 
            : PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
            
  if (!key) throw new Error('GEMINI_API_KEY未設定');
  
  var endpoint = typeof CONFIG !== 'undefined' && CONFIG.GEMINI_API_ENDPOINT
                 ? CONFIG.GEMINI_API_ENDPOINT
                 : 'https://generativelanguage.googleapis.com/v1beta/models/';
                 
  var url = endpoint + model + ':generateContent?key=' + key;
  
  var res = UrlFetchApp.fetch(url, {
    method: 'post', 
    contentType: 'application/json',
    payload: JSON.stringify(body), 
    muteHttpExceptions: true
  });
  
  if (res.getResponseCode() !== 200) {
    Logger.log('[API NON-200 ERROR] ' + res.getContentText());
    return null;
  }
  return JSON.parse(res.getContentText());
}

/**
 * JSON解析のヘルパー関数（Markdown除去）
 */
function _parseGeminiJsonResponse(result) {
  try {
    var text = '';
    if (result.candidates && result.candidates[0] &&
        result.candidates[0].content && result.candidates[0].content.parts) {
      text = result.candidates[0].content.parts[0].text || '';
    }
    // Markdownの装飾ブロックを除去
    text = text.replace(/```json|```/gi, '').trim();
    if (!text) return {};
    return JSON.parse(text);
  } catch(e) {
    Logger.log('[JSON PARSE ERROR RETRY BATCH] ' + e.message);
    return {};
  }
}
