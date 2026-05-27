// ============================================================
// BOM管理システム — スプレッドシートAPI（統合版）
// ============================================================
//
// 【統合元】
//   08 parts api.gs — バルク保存・読み込み型（シンプル）
//   bom_api.gs      — 個別CRUD型（高機能・メイン）
//
// 【重要: BOM_SS_ID について】
//   08 parts api.gs の BOM_SS_ID: '1Fo-zaz1fEbo52yaPBj_lXzCC2yaXCxwCF9P_5Tpe8S0'
//   bom_api.gs の BOM_SS_ID     : '1OEpET_rvYRFVClVuh9VzRbfcEbnpSFfSZ0ZpM8yybUQ'
//   → 2つのIDが異なるため、正しいIDに合わせて BOM_SS_ID を設定してください。
//   → デフォルトは bom_api.gs のIDを使用（CRUD機能が完全なため）。
// ============================================================

// ★ 使用するスプレッドシートのIDをここで設定
var BOM_SS_ID = '1OEpET_rvYRFVClVuh9VzRbfcEbnpSFfSZ0ZpM8yybUQ';

// 旧 08 parts api.gs で使用していたIDが必要な場合はこちら:
// var BOM_SS_ID_LEGACY = '1Fo-zaz1fEbo52yaPBj_lXzCC2yaXCxwCF9P_5Tpe8S0';

var BOM_SHEET = {
  PARTS   : '部品マスタ',
  PRODUCTS: '機種マスタ',
  BOARDS  : '基板マスタ',
  BOM     : 'BOM',
};

// ============================================================
// スプレッドシート取得
// ============================================================
function _getBomSS() {
  return SpreadsheetApp.openById(BOM_SS_ID);
}

// ============================================================
// シートが無ければ作成してヘッダーをセット
// ============================================================
function _ensureBomSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E8F0FE');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ============================================================
// 初期化（シートが無い場合に作成）
// ============================================================
function initBomSheets() {
  var ss = _getBomSS();
  _ensureBomSheet(ss, BOM_SHEET.PARTS,    ['部品ID','部品名','カテゴリ','メーカー','単位','仕様','原価','売値','在庫']);
  _ensureBomSheet(ss, BOM_SHEET.PRODUCTS, ['機種ID','機種名','機種コード','説明']);
  _ensureBomSheet(ss, BOM_SHEET.BOARDS,   ['基板ID','機種ID','基板名','コード','説明','バージョン']);
  _ensureBomSheet(ss, BOM_SHEET.BOM,      ['BOMID','基板ID','部品ID','数量','備考']);
  Logger.log('BOMシート初期化完了');
}

// ============================================================
// 全データ一括取得（apiBomGetAll）
// ============================================================
function apiBomGetAll() {
  try {
    var ss = _getBomSS();

    // 基板マスタ — 日本語ヘッダー → 英語キーへ正規化
    var rawBoards = _sheetToObjects(ss, BOM_SHEET.BOARDS);
    var boards = rawBoards.map(function(r) {
      return {
        id       : String(r['基板ID']      || r['id']        || ''),
        productId: String(r['機種ID']      || r['紐づく機種ID'] || r['productId'] || ''),
        name     : String(r['基板名']      || r['name']      || ''),
        code     : String(r['コード']      || r['基板コード'] || r['code']       || ''),
        desc     : String(r['説明']        || r['desc']      || ''),
        version  : String(r['バージョン']  || r['version']   || ''),
      };
    });

    // 機種マスタ（BOM SS）— 正規化
    var rawProducts = _sheetToObjects(ss, BOM_SHEET.PRODUCTS);
    var products = rawProducts.map(function(r) {
      return {
        id  : String(r['機種ID']    || r['id']   || ''),
        name: String(r['機種名']    || r['name'] || ''),
        code: String(r['機種コード']|| r['code'] || ''),
        desc: String(r['説明']      || r['desc'] || ''),
      };
    });

    // 部品マスタ — 正規化
    var rawParts = _sheetToObjects(ss, BOM_SHEET.PARTS);
    var parts = rawParts.map(function(r) {
      return {
        id      : String(r['部品ID']   || r['id']       || ''),
        name    : String(r['部品名']   || r['name']     || ''),
        category: String(r['カテゴリ'] || r['category'] || ''),
        maker   : String(r['メーカー'] || r['maker']    || ''),
        unit    : String(r['単位']     || r['unit']     || ''),
        spec    : String(r['仕様']     || r['spec']     || ''),
        cost    : Number(r['原価']     || r['cost']     || 0),
        price   : Number(r['売値']     || r['price']    || 0),
        stock   : Number(r['在庫']     || r['stock']    || 0),
      };
    });

    // BOM — 正規化
    var rawBom = _sheetToObjects(ss, BOM_SHEET.BOM);
    var bom = rawBom.map(function(r) {
      return {
        id     : String(r['BOM_ID'] || r['BOMID'] || r['id']      || ''),
        boardId: String(r['基板ID'] || r['boardId']               || ''),
        partId : String(r['部品ID'] || r['partId']                || ''),
        qty    : Number(r['数量']   || r['qty']                   || 0),
        note   : String(r['備考']   || r['note']                  || ''),
      };
    });

    return {
      success : true,
      parts   : parts,
      products: products,
      boards  : boards,
      bom     : bom,
    };
  } catch(e) {
    Logger.log('[apiBomGetAll ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// バルク保存（08 parts api.gs 互換）
// ============================================================
function apiBomSave(dbPayload) {
  try {
    var ss = _getBomSS();
    _writeTableToSheet(ss, BOM_SHEET.PARTS,    dbPayload.parts,        ['id','name','category','maker','cost','price','unit','spec','stock']);
    _writeTableToSheet(ss, BOM_SHEET.PRODUCTS, dbPayload.products,     ['id','name','code','desc']);
    _writeTableToSheet(ss, BOM_SHEET.BOARDS,   dbPayload.boards,       ['id','productId','name','code','desc','version']);
    _writeTableToSheet(ss, BOM_SHEET.BOM,      dbPayload.bom,          ['id','boardId','partId','qty','note']);
    _writeTableToSheet(ss, '価格履歴',         dbPayload.priceHistory, ['id','partId','date','oldPrice','newPrice','reason']);
    return { success: true };
  } catch(e) {
    Logger.log('[BOM SAVE ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// バルク読み込み（08 parts api.gs 互換）
// ============================================================
function apiBomLoad() {
  try {
    var ss = _getBomSS();
    var db = { parts: [], products: [], boards: [], bom: [], priceHistory: [] };
    db.parts        = _readTableFromSheet(ss, BOM_SHEET.PARTS,    ['id','name','category','maker','cost','price','unit','spec','stock']);
    db.products     = _readTableFromSheet(ss, BOM_SHEET.PRODUCTS, ['id','name','code','desc']);
    db.boards       = _readTableFromSheet(ss, BOM_SHEET.BOARDS,   ['id','productId','name','code','desc','version']);
    db.bom          = _readTableFromSheet(ss, BOM_SHEET.BOM,      ['id','boardId','partId','qty','note']);
    db.priceHistory = _readTableFromSheet(ss, '価格履歴',         ['id','partId','date','oldPrice','newPrice','reason']);
    return { success: true, db: db };
  } catch(e) {
    Logger.log('[BOM LOAD ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// 個別 CRUD — 部品
// ============================================================
function apiBomSavePart(p) {
  try {
    var ss    = _getBomSS();
    var sheet = _ensureBomSheet(ss, BOM_SHEET.PARTS,
      ['部品ID','部品名','カテゴリ','メーカー','単位','仕様','原価','売値','在庫']);
    var row = [
      p.id, p.name, p.category || '', p.maker || '',
      p.unit || '個', p.spec || '',
      Number(p.cost) || 0, Number(p.price) || 0, Number(p.stock) || 0,
    ];
    return _upsertRow(sheet, p.id, row);
  } catch(e) { return { success: false, error: e.message }; }
}

function apiBomDeletePart(id) {
  try {
    var ss = _getBomSS();
    _deleteRowById(ss.getSheetByName(BOM_SHEET.PARTS), id);
    _deleteRowsByCol(ss.getSheetByName(BOM_SHEET.BOM), 3, id);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ============================================================
// 個別 CRUD — 機種
// ============================================================
function apiBomSaveProduct(p) {
  try {
    var ss    = _getBomSS();
    var sheet = _ensureBomSheet(ss, BOM_SHEET.PRODUCTS, ['機種ID','機種名','機種コード','説明']);
    return _upsertRow(sheet, p.id, [p.id, p.name, p.code || '', p.desc || '']);
  } catch(e) { return { success: false, error: e.message }; }
}

function apiBomDeleteProduct(id) {
  try {
    var ss = _getBomSS();
    _deleteRowById(ss.getSheetByName(BOM_SHEET.PRODUCTS), id);
    _deleteRowsByCol(ss.getSheetByName(BOM_SHEET.BOARDS), 2, id);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ============================================================
// 個別 CRUD — 基板
// ============================================================
function apiBomSaveBoard(b) {
  try {
    var ss    = _getBomSS();
    var sheet = _ensureBomSheet(ss, BOM_SHEET.BOARDS,
      ['基板ID','機種ID','基板名','コード','説明','バージョン']);
    return _upsertRow(sheet, b.id, [b.id, b.productId, b.name, b.code || '', b.desc || '', b.version || '']);
  } catch(e) { return { success: false, error: e.message }; }
}

function apiBomDeleteBoard(id) {
  try {
    var ss = _getBomSS();
    _deleteRowById(ss.getSheetByName(BOM_SHEET.BOARDS), id);
    _deleteRowsByCol(ss.getSheetByName(BOM_SHEET.BOM), 2, id);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ============================================================
// 基板詳細取得（関連見積・注文・機種情報付き）
// ============================================================

function apiBoardMasterGet(payload) {
  try {
    var boardId = String((payload || {}).boardId || '').trim();
    if (!boardId) return { success: false, error: '基板IDが必要です' };

    var ss = _getBomSS();

    // ── 基板情報 ──
    var rawBoards = _sheetToObjects(ss, BOM_SHEET.BOARDS);
    var rawBoard  = rawBoards.filter(function(r) {
      return String(r['基板ID'] || r['id'] || '').trim() === boardId;
    })[0];
    if (!rawBoard) return { success: false, error: '基板が見つかりません: ' + boardId };

    var boardInfo = {
      id       : boardId,
      productId: String(rawBoard['機種ID']     || rawBoard['productId'] || ''),
      name     : String(rawBoard['基板名']     || rawBoard['name']      || ''),
      code     : String(rawBoard['コード']     || rawBoard['code']      || ''),
      desc     : String(rawBoard['説明']       || rawBoard['desc']      || ''),
      version  : String(rawBoard['バージョン'] || rawBoard['version']  || ''),
    };

    // ── 親機種（機種マスタ）情報 ──
    var parentProduct = null;
    var modelCode     = '';
    if (boardInfo.productId) {
      var rawProducts = _sheetToObjects(ss, BOM_SHEET.PRODUCTS);
      var rawProduct  = rawProducts.filter(function(p) {
        return String(p['機種ID'] || p['id'] || '').trim() === boardInfo.productId;
      })[0];
      if (rawProduct) {
        modelCode    = String(rawProduct['機種コード'] || rawProduct['code'] || '');
        parentProduct = {
          id  : boardInfo.productId,
          name: String(rawProduct['機種名'] || rawProduct['name'] || ''),
          code: modelCode,
          desc: String(rawProduct['説明']   || rawProduct['desc'] || ''),
        };
      }
    }

    // ── 関連見積書・注文書（機種コードで管理シートを検索）──
    var relatedQuotes = [];
    var relatedOrders = [];

    if (modelCode) {
      var mgmtData  = getAllMgmtData();
      var seenQuotes = {};
      var seenOrders = {};

      mgmtData.forEach(function(r) {
        var mc = String(r[MGMT_COLS.MODEL_CODE - 1] || '').trim();
        if (mc !== modelCode) return;

        var qNo = String(r[MGMT_COLS.QUOTE_NO - 1] || '').trim();
        var oNo = String(r[MGMT_COLS.ORDER_NO  - 1] || '').trim();

        if (qNo && !seenQuotes[qNo]) {
          seenQuotes[qNo] = true;
          relatedQuotes.push({
            mgmtId   : String(r[MGMT_COLS.ID           - 1] || ''),
            quoteNo  : qNo,
            client   : String(r[MGMT_COLS.CLIENT        - 1] || ''),
            quoteDate: _toDateStr(r[MGMT_COLS.QUOTE_DATE - 1]),
            amount   : _toNum(r[MGMT_COLS.QUOTE_AMOUNT  - 1]),
            status   : String(r[MGMT_COLS.STATUS        - 1] || ''),
            pdfUrl   : String(r[MGMT_COLS.QUOTE_PDF_URL - 1] || ''),
            subject  : String(r[MGMT_COLS.SUBJECT       - 1] || ''),
            linked   : _isLinkedVal(r[MGMT_COLS.LINKED  - 1]),
          });
        }
        if (oNo && !seenOrders[oNo]) {
          seenOrders[oNo] = true;
          relatedOrders.push({
            mgmtId      : String(r[MGMT_COLS.ID              - 1] || ''),
            orderNo     : oNo,
            client      : String(r[MGMT_COLS.CLIENT           - 1] || ''),
            orderDate   : _toDateStr(r[MGMT_COLS.ORDER_DATE   - 1]),
            amount      : _toNum(r[MGMT_COLS.ORDER_AMOUNT     - 1]),
            status      : String(r[MGMT_COLS.STATUS           - 1] || ''),
            pdfUrl      : String(r[MGMT_COLS.ORDER_PDF_URL    - 1] || ''),
            deliveryDate: _toDateStr(r[MGMT_COLS.DELIVERY_DATE - 1]),
            orderType   : String(r[MGMT_COLS.ORDER_TYPE       - 1] || ''),
            orderSlipNo : String(r[MGMT_COLS.ORDER_SLIP_NO    - 1] || ''),
            linked      : _isLinkedVal(r[MGMT_COLS.LINKED     - 1]),
          });
        }
      });

      relatedQuotes.sort(function(a,b){ return (b.quoteDate||'').localeCompare(a.quoteDate||''); });
      relatedOrders.sort(function(a,b){ return (b.orderDate||'').localeCompare(a.orderDate||''); });
    }

    return JSON.parse(JSON.stringify({
      success      : true,
      boardInfo    : boardInfo,
      parentProduct: parentProduct,
      modelCode    : modelCode,
      relatedQuotes: relatedQuotes,
      relatedOrders: relatedOrders,
    }));
  } catch(e) {
    Logger.log('[apiBoardMasterGet ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// 個別 CRUD — BOM行
// ============================================================
function apiBomSaveBomRow(x) {
  try {
    var ss    = _getBomSS();
    var sheet = _ensureBomSheet(ss, BOM_SHEET.BOM, ['BOMID','基板ID','部品ID','数量','備考']);
    return _upsertRow(sheet, x.id, [x.id, x.boardId, x.partId, Number(x.qty) || 1, x.note || '']);
  } catch(e) { return { success: false, error: e.message }; }
}

function apiBomDeleteBomRow(id) {
  try {
    _deleteRowById(_getBomSS().getSheetByName(BOM_SHEET.BOM), id);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ============================================================
// CSV一括インポート（最終リスト）
// ============================================================
function apiBomImportFinalList(rows) {
  try {
    var ss         = _getBomSS();
    var partSheet  = _ensureBomSheet(ss, BOM_SHEET.PARTS,    ['部品ID','部品名','カテゴリ','メーカー','単位','仕様','原価','売値','在庫']);
    var prodSheet  = _ensureBomSheet(ss, BOM_SHEET.PRODUCTS, ['機種ID','機種名','機種コード','説明']);
    var boardSheet = _ensureBomSheet(ss, BOM_SHEET.BOARDS,   ['基板ID','機種ID','基板名','コード','説明','バージョン']);
    var bomSheet   = _ensureBomSheet(ss, BOM_SHEET.BOM,      ['BOMID','基板ID','部品ID','数量','備考']);

    var addedProducts = 0, addedBoards = 0, addedParts = 0, updatedBom = 0;

    var existProds  = _sheetToObjects(ss, BOM_SHEET.PRODUCTS);
    var existBoards = _sheetToObjects(ss, BOM_SHEET.BOARDS);
    var existParts  = _sheetToObjects(ss, BOM_SHEET.PARTS);
    var existBom    = _sheetToObjects(ss, BOM_SHEET.BOM);

    var prodMap  = {}; existProds.forEach(function(p)  { prodMap[p['機種コード']]                 = p; });
    var boardMap = {}; existBoards.forEach(function(b) { boardMap[b['機種ID'] + '_' + b['基板名']] = b; });
    var partMap  = {}; existParts.forEach(function(p)  { partMap[p['部品ID']]                     = p; });
    var bomMap   = {}; existBom.forEach(function(x)    { bomMap[x['基板ID'] + '_' + x['部品ID']]  = x; });

    var newProds = [], newBoards = [], newParts = [], newBom = [], updateBom = [];

    rows.forEach(function(row) {
      if (row.length < 11) return;
      var prodCode  = String(row[0]).trim();
      var prodName  = String(row[1]).trim();
      var boardCode = String(row[3]).trim();
      var boardName = String(row[4]).trim();
      var partCode  = String(row[5]).trim();
      var partName  = String(row[6]).trim();
      var maker     = String(row[9]).trim();
      var qty       = Number(row[10]) || 0;
      var note      = row.length > 15 ? String(row[15]).trim() : '';

      if (!prodCode || !boardName || (!partCode && !partName)) return;

      var prod = prodMap[prodCode];
      if (!prod) {
        var newProdId = 'M' + Utilities.formatDate(new Date(),'Asia/Tokyo','yyyyMMddHHmmss') + Math.floor(Math.random()*100);
        prod = { '機種ID': newProdId, '機種名': prodName||prodCode, '機種コード': prodCode, '説明': 'CSV自動作成' };
        prodMap[prodCode] = prod;
        newProds.push([newProdId, prodName||prodCode, prodCode, 'CSV自動作成']);
        addedProducts++;
      }

      var boardKey = prod['機種ID'] + '_' + boardName;
      var board = boardMap[boardKey];
      if (!board) {
        var newBoardId = 'B' + Utilities.formatDate(new Date(),'Asia/Tokyo','yyyyMMddHHmmss') + Math.floor(Math.random()*100);
        board = { '基板ID': newBoardId, '機種ID': prod['機種ID'], '基板名': boardName, 'コード': boardCode };
        boardMap[boardKey] = board;
        newBoards.push([newBoardId, prod['機種ID'], boardName, boardCode, '', '']);
        addedBoards++;
      }

      var pId = partCode || ('P' + Utilities.formatDate(new Date(),'Asia/Tokyo','yyyyMMddHHmmss') + Math.floor(Math.random()*100));
      var part = partMap[pId];
      if (!part) {
        part = { '部品ID': pId, '部品名': partName };
        partMap[pId] = part;
        newParts.push([pId, partName, '電子部品', maker, '個', '', 0, 0, 0]);
        addedParts++;
      }

      var bomKey  = board['基板ID'] + '_' + pId;
      var existing = bomMap[bomKey];
      if (existing) {
        updateBom.push({ id: existing['BOMID'], boardId: board['基板ID'], partId: pId, qty: qty, note: note });
      } else {
        var newBomId = 'BOM' + Utilities.formatDate(new Date(),'Asia/Tokyo','yyyyMMddHHmmss') + Math.floor(Math.random()*100);
        bomMap[bomKey] = { 'BOMID': newBomId };
        newBom.push([newBomId, board['基板ID'], pId, qty, note]);
      }
      updatedBom++;
    });

    if (newProds.length  > 0) prodSheet.getRange(prodSheet.getLastRow()+1,   1, newProds.length,  4).setValues(newProds);
    if (newBoards.length > 0) boardSheet.getRange(boardSheet.getLastRow()+1,  1, newBoards.length, 6).setValues(newBoards);
    if (newParts.length  > 0) partSheet.getRange(partSheet.getLastRow()+1,   1, newParts.length,  9).setValues(newParts);
    if (newBom.length    > 0) bomSheet.getRange(bomSheet.getLastRow()+1,     1, newBom.length,    5).setValues(newBom);
    updateBom.forEach(function(u){ _upsertRow(bomSheet, u.id, [u.id, u.boardId, u.partId, u.qty, u.note]); });

    return { success: true, addedProducts: addedProducts, addedBoards: addedBoards, addedParts: addedParts, updatedBom: updatedBom };
  } catch(e) {
    Logger.log('[BOM IMPORT ERROR] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

// ============================================================
// CSV一括インポート（k10部品表）
// ============================================================
function apiBomImportK10Parts(rows) {
  try {
    var ss        = _getBomSS();
    var partSheet = _ensureBomSheet(ss, BOM_SHEET.PARTS,
      ['部品ID','部品名','カテゴリ','メーカー','単位','仕様','原価','売値','在庫']);
    var added = 0, updated = 0;
    rows.forEach(function(row) {
      if (row.length < 18) return;
      var id    = String(row[2]).trim();
      var name  = String(row[3]).trim();
      var maker = String(row[10]).trim();
      var cost  = parseFloat(row[12]) || 0;
      var spec  = String(row[17]).trim();
      if (!id || !name) return;
      var result = _upsertRow(partSheet, id, [id, name, '電子部品', maker === '無し' ? '' : maker, '個', spec, cost, 0, 0]);
      if (result.isNew) added++; else updated++;
    });
    return { success: true, added: added, updated: updated };
  } catch(e) { return { success: false, error: e.message }; }
}

// ============================================================
// ユーティリティ
// ============================================================

function _sheetToObjects(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data    = sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).getValues();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return data
    .filter(function(r){ return String(r[0]).trim() !== ''; })
    .map(function(r) {
      var obj = {};
      headers.forEach(function(h, i){ obj[String(h)] = r[i] !== undefined ? r[i] : ''; });
      return obj;
    });
}

function _writeTableToSheet(ss, sheetName, dataArray, keys) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clearContents();

  var headers = keys;
  if (sheetName === BOM_SHEET.PARTS)    headers = ['部品ID','部品名','カテゴリ','メーカー','原価','売値','単位','仕様','在庫'];
  if (sheetName === BOM_SHEET.PRODUCTS) headers = ['機種ID','機種名','機種コード','説明'];
  if (sheetName === BOM_SHEET.BOARDS)   headers = ['基板ID','紐づく機種ID','基板名','基板コード','説明','バージョン'];
  if (sheetName === BOM_SHEET.BOM)      headers = ['BOM_ID','基板ID','部品ID','数量','備考'];
  if (sheetName === '価格履歴')          headers = ['履歴ID','部品ID','改定日','旧価格','新価格','理由'];

  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#e8f5ed').setFontWeight('bold');

  if (dataArray && dataArray.length > 0) {
    var rows = dataArray.map(function(obj) {
      return keys.map(function(k){ return obj[k] !== undefined ? obj[k] : ''; });
    });
    sheet.getRange(2, 1, rows.length, keys.length).setValues(rows);
  }
}

function _readTableFromSheet(ss, sheetName, keys) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, keys.length).getValues();
  return data.map(function(row) {
    var obj = {};
    keys.forEach(function(k, i){ obj[k] = row[i]; });
    return obj;
  });
}

function _upsertRow(sheet, id, rowData) {
  if (!sheet || !id) return { success: false, error: 'パラメータ不足' };
  var last = sheet.getLastRow();
  if (last > 1) {
    var ids = sheet.getRange(2, 1, last-1, 1).getValues().flat().map(String);
    var idx = ids.indexOf(String(id));
    if (idx >= 0) {
      sheet.getRange(idx+2, 1, 1, rowData.length).setValues([rowData]);
      return { success: true, isNew: false };
    }
  }
  sheet.appendRow(rowData);
  return { success: true, isNew: true };
}

function _deleteRowById(sheet, id) {
  if (!sheet || sheet.getLastRow() <= 1) return;
  var ids = sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues().flat().map(String);
  for (var i = ids.length-1; i >= 0; i--) {
    if (ids[i] === String(id)) sheet.deleteRow(i+2);
  }
}

function _deleteRowsByCol(sheet, colNum, value) {
  if (!sheet || sheet.getLastRow() <= 1) return;
  var vals = sheet.getRange(2, colNum, sheet.getLastRow()-1, 1).getValues().flat().map(String);
  for (var i = vals.length-1; i >= 0; i--) {
    if (vals[i] === String(value)) sheet.deleteRow(i+2);
  }
}
