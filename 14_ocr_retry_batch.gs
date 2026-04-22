// ============================================================
// ファイル: 14_ocr_retry_batch.gs
// 目的: 夜間バッチによるOCR情報の自己修復（パターンB: 管理シート完全統合版）
// ============================================================

const OCR_RETRY_CONFIG = {
  // 確実に枠(RPD 20)があるモデルを指定
  MODEL: 'gemini-2.5-flash', 
  // 5 RPM（1分間に5回）制限を回避するため15秒待機にする
  SLEEP_MS: 15000,            
  MAX_EXEC_SECONDS: 270,      // GASタイムアウトエラー回避（4.5分で強制終了）
  MAX_RETRY: 3,               // 最大リトライ回数
  STATUS_COL_NAME: 'OCRステータス',
  RETRY_COL_NAME: 'リトライ回数',
  // スクリプトプロパティに保存した夜間用APIキーの名称
  BATCH_API_KEY_NAME: 'GEMINI_API_KEY_BATCH' 
};

/**
 * 【Time-driven Trigger】夜間に定期実行されるバッチのメイン関数
 * トリガー設定から「1日1回 深夜」に設定して使用します
 */
function runOcrRetryBatch() {
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 実運用されている管理シートを取得
  var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  if (!sheet) {
    Logger.log('[OCR RETRY] 管理シートが見つかりません。');
    return;
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1) return; // データなし

  var range = sheet.getRange(1, 1, lastRow, lastCol);
  var values = range.getValues();
  var headers = values[0];

  // OCRステータス、リトライ回数の列が存在するか確認し、なければ末尾に追加
  var statusColIdx = headers.indexOf(OCR_RETRY_CONFIG.STATUS_COL_NAME);
  var retryColIdx = headers.indexOf(OCR_RETRY_CONFIG.RETRY_COL_NAME);

  if (statusColIdx === -1) {
    statusColIdx = lastCol;
    sheet.getRange(1, statusColIdx + 1).setValue(OCR_RETRY_CONFIG.STATUS_COL_NAME);
    headers.push(OCR_RETRY_CONFIG.STATUS_COL_NAME);
    lastCol++;
  }
  if (retryColIdx === -1) {
    retryColIdx = lastCol;
    sheet.getRange(1, retryColIdx + 1).setValue(OCR_RETRY_CONFIG.RETRY_COL_NAME);
    headers.push(OCR_RETRY_CONFIG.RETRY_COL_NAME);
    lastCol++;
  }

  // ★ 補完対象とする業務データ（空欄であれば修復を試みる主要なカラム）
  var targetColIndexes = [
    MGMT_COLS.QUOTE_NO - 1,       // 見積書番号
    MGMT_COLS.ORDER_NO - 1,       // 注文書番号
    MGMT_COLS.SUBJECT - 1,        // 件名
    MGMT_COLS.CLIENT - 1,         // 顧客名
    MGMT_COLS.QUOTE_DATE - 1,     // 見積日
    MGMT_COLS.ORDER_DATE - 1,     // 注文日
    MGMT_COLS.QUOTE_AMOUNT - 1,   // 見積金額
    MGMT_COLS.ORDER_AMOUNT - 1    // 注文金額
  ];

  for (var i = 1; i < values.length; i++) {
    // タイムアウト検証
    var elapsedSeconds = (new Date().getTime() - startTime) / 1000;
    if (elapsedSeconds > OCR_RETRY_CONFIG.MAX_EXEC_SECONDS) {
      Logger.log('[OCR RETRY] タイムアウト回避規定時間に達したため中断。完了行: ' + i);
      break;
    }

    var rowId    = values[i][MGMT_COLS.ID - 1] || 'Unknown';
    var pdfUrl   = values[i][MGMT_COLS.QUOTE_PDF_URL - 1] || values[i][MGMT_COLS.ORDER_PDF_URL - 1]; 
    var status   = values[i][statusColIdx] || '';
    var retryCnt = Number(values[i][retryColIdx]) || 0;

    // 抽出条件：ステータスが「補完待ち」かつ リトライ回数が上限未満
    if (status === '補完待ち' && retryCnt < OCR_RETRY_CONFIG.MAX_RETRY) {
      
      if (!pdfUrl) {
         sheet.getRange(i + 1, statusColIdx + 1).setValue('手動確認');
         continue;
      }

      // 不足項目の特定
      var missingHeaders = [];
      var missingIndexes = [];
      targetColIndexes.forEach(function(colIdx) {
        var val = String(values[i][colIdx] || '').trim();
        if (val === '' || val === '—' || val === '-') {
          missingHeaders.push(headers[colIdx]);
          missingIndexes.push(colIdx);
        }
      });

      if (missingHeaders.length === 0) {
        sheet.getRange(i + 1, statusColIdx + 1).setValue('完了');
        continue;
      }

      Logger.log('[OCR RETRY] 対象ID: ' + rowId + ' 不足項目: ' + missingHeaders.join(', '));

      try {
        var fileId = _extractFileIdFromUrl(pdfUrl);
        if (!fileId) throw new Error("Invalid PDF URL");
        
        var file = DriveApp.getFileById(fileId);
        var mimeType = file.getMimeType() || 'application/pdf';
        var base64 = Utilities.base64Encode(file.getBlob().getBytes());

        var prompt = 'この画像は書類（見積書または注文書）です。\n' + 
                     '前回のデータベース登録時に以下の項目が読み取りできませんでした。\n' +
                     '画像から [' + missingHeaders.join(', ') + '] のみを推測・抽出し、厳格なJSON形式で返してください。\n' +
                     'JSONキーには日本語のカラム名をそのまま利用してください。値が取得できない場合は空文字としてください。\n' +
                     'ルール: 有効なJSONのみ。マークダウンや説明は一切記載しないこと。金額などは数字のみ。';

        var body = {
          contents: [{ parts: [
            { text: prompt },
            { inline_data: { mime_type: mimeType, data: base64 } }
          ]}],
          generationConfig: { temperature: 0.1, responseMimeType: 'application/json' }
        };

        var apiRes = _callGeminiApiOcrRetry(OCR_RETRY_CONFIG.MODEL, body);
        
        var extractedData = {};
        if (apiRes) extractedData = _parseGeminiJsonResponse(apiRes);

        var isAllFilled = true;
        missingIndexes.forEach(function(colIdx) {
          var hName = headers[colIdx];
          var extVal = extractedData[hName];
          if (extVal && String(extVal).trim() !== '') {
            sheet.getRange(i + 1, colIdx + 1).setValue(extVal);
          } else {
             isAllFilled = false;
          }
        });

        var nextRetry = retryCnt + 1;
        sheet.getRange(i + 1, retryColIdx + 1).setValue(nextRetry);

        if (isAllFilled) {
          sheet.getRange(i + 1, statusColIdx + 1).setValue('完了');
        } else if (nextRetry >= OCR_RETRY_CONFIG.MAX_RETRY) {
          sheet.getRange(i + 1, statusColIdx + 1).setValue('手動確認');
        }

      } catch(e) {
        Logger.log('[OCR RETRY ERROR] ' + rowId + ': ' + e.message);
      }
      
      // API制限回避の待機
      Utilities.sleep(OCR_RETRY_CONFIG.SLEEP_MS);
    }
  }
}

/**
 * URLからGoogle DriveのFile IDを抽出
 */
function _extractFileIdFromUrl(url) {
  var match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

/**
 * Gemini APIへの直接リクエスト（バッチ専用APIキー対応）
 */
function _callGeminiApiOcrRetry(model, body) {
  // 1. スクリプトプロパティから直接取得を試みる
  var batchKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY_BATCH');
  
  var key = "";
  if (batchKey) {
    Logger.log("[DEBUG] 夜間用キー(BATCH)を使用します: " + batchKey.substring(0, 5) + "...");
    key = batchKey;
  } else {
    Logger.log("[DEBUG] 夜間用キーが見つからないため、通常キーを使用します。");
    key = typeof CONFIG !== 'undefined' && CONFIG.GEMINI_API_KEY 
          ? CONFIG.GEMINI_API_KEY 
          : PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  }
            
  if (!key) throw new Error('APIキーがどこにも設定されていません');
  
  // 以降の通信処理はそのまま
  var endpoint = 'https://generativelanguage.googleapis.com/v1beta/models/';
  var url = endpoint + model + ':generateContent?key=' + key;
  
  var res = UrlFetchApp.fetch(url, {
    method: 'post', 
    contentType: 'application/json',
    payload: JSON.stringify(body), 
    muteHttpExceptions: true
  });
  
  return JSON.parse(res.getContentText());
}

/**
 * JSON解析
 */
function _parseGeminiJsonResponse(result) {
  try {
    var text = '';
    if (result.candidates && result.candidates[0] &&
        result.candidates[0].content && result.candidates[0].content.parts) {
      text = result.candidates[0].content.parts[0].text || '';
    }
    text = text.replace(/```json|```/gi, '').trim();
    if (!text) return {};
    return JSON.parse(text);
  } catch(e) {
    Logger.log('[JSON PARSE ERROR] ' + e.message);
    return {};
  }
}

/**
 * 空欄データを「補完待ち」にセットする（手動リセット用）
 */
function resetEmptyFieldsForOcrRetry() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
  if (!sheet) return;

  var values = sheet.getDataRange().getValues();
  var headers = values[0];

  var statusColIdx = headers.indexOf(OCR_RETRY_CONFIG.STATUS_COL_NAME);
  var retryColIdx = headers.indexOf(OCR_RETRY_CONFIG.RETRY_COL_NAME);

  if (statusColIdx === -1) {
    Logger.log('OCRステータス列がありません。runOcrRetryBatchを一度実行してください。');
    return;
  }

  var resetCount = 0;
  for (var i = 1; i < values.length; i++) {
    var subject = String(values[i][MGMT_COLS.SUBJECT - 1] || '').trim();
    var qDate = String(values[i][MGMT_COLS.QUOTE_DATE - 1] || '').trim();
    
    // 件名が空欄・「—」・「-」の場合、または日付が空の場合
    if (subject === '' || subject === '—' || subject === '-' || qDate === '') {
      sheet.getRange(i + 1, statusColIdx + 1).setValue('補完待ち');
      sheet.getRange(i + 1, retryColIdx + 1).setValue(0);
      resetCount++;
    }
  }
  Logger.log(resetCount + ' 件を「補完待ち」にセットしました。');
}