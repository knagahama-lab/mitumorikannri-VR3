// ============================================================
// BOM管理システム - スプレッドシート同期 API
// ============================================================

// ★連携するスプレッドシートのID
var BOM_SS_ID = '1Fo-zaz1fEbo52yaPBj_lXzCC2yaXCxwCF9P_5Tpe8S0';

function apiBomSave(dbPayload) {
  try {
    var ss = SpreadsheetApp.openById(BOM_SS_ID);
    
    // それぞれの配列を別々のシートに「表」として書き込む（洗い替え）
    _writeTableToSheet(ss, '部品マスタ', dbPayload.parts, ['id','name','category','maker','cost','price','unit','spec','stock']);
    _writeTableToSheet(ss, '機種マスタ', dbPayload.products, ['id','name','code','desc']);
    _writeTableToSheet(ss, '基板マスタ', dbPayload.boards, ['id','productId','name','code','desc','version']);
    _writeTableToSheet(ss, 'BOM', dbPayload.bom, ['id','boardId','partId','qty','note']);
    _writeTableToSheet(ss, '価格履歴', dbPayload.priceHistory, ['id','partId','date','oldPrice','newPrice','reason']);
    
    return { success: true };
  } catch(e) {
    Logger.log('[BOM SAVE ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

function apiBomLoad() {
  try {
    var ss = SpreadsheetApp.openById(BOM_SS_ID);
    var db = { parts: [], products: [], boards: [], bom: [], priceHistory: [] };
    
    db.parts = _readTableFromSheet(ss, '部品マスタ', ['id','name','category','maker','cost','price','unit','spec','stock']);
    db.products = _readTableFromSheet(ss, '機種マスタ', ['id','name','code','desc']);
    db.boards = _readTableFromSheet(ss, '基板マスタ', ['id','productId','name','code','desc','version']);
    db.bom = _readTableFromSheet(ss, 'BOM', ['id','boardId','partId','qty','note']);
    db.priceHistory = _readTableFromSheet(ss, '価格履歴', ['id','partId','date','oldPrice','newPrice','reason']);
    
    return { success: true, db: db };
  } catch(e) {
    Logger.log('[BOM LOAD ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// データを表形式でシートに書き込む共通関数
function _writeTableToSheet(ss, sheetName, dataArray, keys) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  sheet.clearContents(); // 一旦全クリア
  
  // 見やすいように日本語のヘッダー名を設定
  var headers = keys;
  if(sheetName === '部品マスタ') headers = ['部品ID','部品名','カテゴリ','メーカー','原価','売値','単位','仕様','在庫'];
  if(sheetName === '機種マスタ') headers = ['機種ID','機種名','機種コード','説明'];
  if(sheetName === '基板マスタ') headers = ['基板ID','紐づく機種ID','基板名','基板コード','説明','バージョン'];
  if(sheetName === 'BOM') headers = ['BOM_ID','基板ID','部品ID','数量','備考'];
  if(sheetName === '価格履歴') headers = ['履歴ID','部品ID','改定日','旧価格','新価格','理由'];
  
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#e8f5ed').setFontWeight('bold');
  
  if (dataArray && dataArray.length > 0) {
    var rows = dataArray.map(function(obj) {
      return keys.map(function(k) { return obj[k] !== undefined ? obj[k] : ''; });
    });
    sheet.getRange(2, 1, rows.length, keys.length).setValues(rows);
  }
}

// シートからデータを読み込んでオブジェクト配列に戻す共通関数
function _readTableFromSheet(ss, sheetName, keys) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, keys.length).getValues();
  return data.map(function(row) {
    var obj = {};
    keys.forEach(function(k, i) { obj[k] = row[i]; });
    return obj;
  });
}