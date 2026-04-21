// ============================================================
// BOM管理システム用 スプレッドシートAPI
// ファイル: bom_api.gs
// ============================================================

var BOM_SS_ID = '1OEpET_rvYRFVClVuh9VzRbfcEbnpSFfSZ0ZpM8yybUQ'; // 既存の基板SSと同じ

var BOM_SHEET = {
  PARTS:    '部品マスタ',
  PRODUCTS: '機種マスタ',
  BOARDS:   '基板マスタ',
  BOM:      'BOM',
};

// ===== スプレッドシート取得 =====
function _getBomSS() {
  return SpreadsheetApp.openById(BOM_SS_ID);
}

// ===== シートが無ければ作成し、ヘッダーをセット =====
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

// ===== 初期化（シートが無い場合に作成） =====
function initBomSheets() {
  var ss = _getBomSS();
  _ensureBomSheet(ss, BOM_SHEET.PARTS,    ['部品ID','部品名','カテゴリ','メーカー','単位','仕様','原価','売値','在庫']);
  _ensureBomSheet(ss, BOM_SHEET.PRODUCTS, ['機種ID','機種名','機種コード','説明']);
  _ensureBomSheet(ss, BOM_SHEET.BOARDS,   ['基板ID','機種ID','基板名','コード','説明','バージョン']);
  _ensureBomSheet(ss, BOM_SHEET.BOM,      ['BOMID','基板ID','部品ID','数量','備考']);
  Logger.log('BOMシート初期化完了');
}

// ===== 全データ一括取得 =====
function apiBomGetAll() {
  try {
    var ss = _getBomSS();
    return {
      success:  true,
      parts:    _sheetToObjects(ss, BOM_SHEET.PARTS),
      products: _sheetToObjects(ss, BOM_SHEET.PRODUCTS),
      boards:   _sheetToObjects(ss, BOM_SHEET.BOARDS),
      bom:      _sheetToObjects(ss, BOM_SHEET.BOM),
    };
  } catch(e) {
    Logger.log('[apiBomGetAll ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ===== シートデータをオブジェクト配列に変換 =====
function _sheetToObjects(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return data
    .filter(function(r) { return String(r[0]).trim() !== ''; })
    .map(function(r) {
      var obj = {};
      headers.forEach(function(h, i) { obj[String(h)] = r[i] !== undefined ? r[i] : ''; });
      return obj;
    });
}

// ===== 部品保存（追加 or 上書き） =====
function apiBomSavePart(p) {
  try {
    var ss    = _getBomSS();
    var sheet = _ensureBomSheet(ss, BOM_SHEET.PARTS,
      ['部品ID','部品名','カテゴリ','メーカー','単位','仕様','原価','売値','在庫']);
    var row = [
      p.id, p.name, p.category || '', p.maker || '',
      p.unit || '個', p.spec || '',
      Number(p.cost) || 0, Number(p.price) || 0, Number(p.stock) || 0
    ];
    return _upsertRow(sheet, p.id, row);
  } catch(e) { return { success: false, error: e.message }; }
}

// ===== 部品削除 =====
function apiBomDeletePart(id) {
  try {
    var ss    = _getBomSS();
    var sheet = ss.getSheetByName(BOM_SHEET.PARTS);
    _deleteRowById(sheet, id);
    // BOMからも削除
    var bomSheet = ss.getSheetByName(BOM_SHEET.BOM);
    _deleteRowsByCol(bomSheet, 3, id); // 3列目=部品ID
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ===== 機種保存 =====
function apiBomSaveProduct(p) {
  try {
    var ss    = _getBomSS();
    var sheet = _ensureBomSheet(ss, BOM_SHEET.PRODUCTS, ['機種ID','機種名','機種コード','説明']);
    var row   = [p.id, p.name, p.code || '', p.desc || ''];
    return _upsertRow(sheet, p.id, row);
  } catch(e) { return { success: false, error: e.message }; }
}

// ===== 機種削除 =====
function apiBomDeleteProduct(id) {
  try {
    var ss    = _getBomSS();
    _deleteRowById(ss.getSheetByName(BOM_SHEET.PRODUCTS), id);
    // 配下の基板も削除
    var boardSheet = ss.getSheetByName(BOM_SHEET.BOARDS);
    _deleteRowsByCol(boardSheet, 2, id); // 2列目=機種ID
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ===== 基板保存 =====
function apiBomSaveBoard(b) {
  try {
    var ss    = _getBomSS();
    var sheet = _ensureBomSheet(ss, BOM_SHEET.BOARDS,
      ['基板ID','機種ID','基板名','コード','説明','バージョン']);
    var row = [b.id, b.productId, b.name, b.code || '', b.desc || '', b.version || ''];
    return _upsertRow(sheet, b.id, row);
  } catch(e) { return { success: false, error: e.message }; }
}

// ===== 基板削除 =====
function apiBomDeleteBoard(id) {
  try {
    var ss = _getBomSS();
    _deleteRowById(ss.getSheetByName(BOM_SHEET.BOARDS), id);
    _deleteRowsByCol(ss.getSheetByName(BOM_SHEET.BOM), 2, id); // 2列目=基板ID
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ===== BOM行保存 =====
function apiBomSaveBomRow(x) {
  try {
    var ss    = _getBomSS();
    var sheet = _ensureBomSheet(ss, BOM_SHEET.BOM,
      ['BOMID','基板ID','部品ID','数量','備考']);
    var row = [x.id, x.boardId, x.partId, Number(x.qty) || 1, x.note || ''];
    return _upsertRow(sheet, x.id, row);
  } catch(e) { return { success: false, error: e.message }; }
}

// ===== BOM行削除 =====
function apiBomDeleteBomRow(id) {
  try {
    var ss    = _getBomSS();
    _deleteRowById(ss.getSheetByName(BOM_SHEET.BOM), id);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ===== CSV一括インポート（最終リスト） =====
function apiBomImportFinalList(rows) {
  try {
    var ss           = _getBomSS();
    var partSheet    = _ensureBomSheet(ss, BOM_SHEET.PARTS,    ['部品ID','部品名','カテゴリ','メーカー','単位','仕様','原価','売値','在庫']);
    var prodSheet    = _ensureBomSheet(ss, BOM_SHEET.PRODUCTS, ['機種ID','機種名','機種コード','説明']);
    var boardSheet   = _ensureBomSheet(ss, BOM_SHEET.BOARDS,   ['基板ID','機種ID','基板名','コード','説明','バージョン']);
    var bomSheet     = _ensureBomSheet(ss, BOM_SHEET.BOM,      ['BOMID','基板ID','部品ID','数量','備考']);

    var addedProducts = 0, addedBoards = 0, addedParts = 0, updatedBom = 0;

    // 既存データをメモリに読み込んでキャッシュ（高速化）
    var existProds  = _sheetToObjects(ss, BOM_SHEET.PRODUCTS);
    var existBoards = _sheetToObjects(ss, BOM_SHEET.BOARDS);
    var existParts  = _sheetToObjects(ss, BOM_SHEET.PARTS);
    var existBom    = _sheetToObjects(ss, BOM_SHEET.BOM);

    var prodMap   = {}; existProds.forEach(function(p)  { prodMap[p['機種コード']]              = p; });
    var boardMap  = {}; existBoards.forEach(function(b) { boardMap[b['機種ID'] + '_' + b['基板名']] = b; });
    var partMap   = {}; existParts.forEach(function(p)  { partMap[p['部品ID']]                  = p; });
    var bomMap    = {}; existBom.forEach(function(x)    { bomMap[x['基板ID'] + '_' + x['部品ID']] = x; });

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

      // 機種
      var prod = prodMap[prodCode];
      if (!prod) {
        var newProdId = 'M' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + Math.floor(Math.random()*100);
        prod = { '機種ID': newProdId, '機種名': prodName || prodCode, '機種コード': prodCode, '説明': 'CSV自動作成' };
        prodMap[prodCode] = prod;
        newProds.push([newProdId, prodName || prodCode, prodCode, 'CSV自動作成']);
        addedProducts++;
      }

      // 基板
      var boardKey = prod['機種ID'] + '_' + boardName;
      var board = boardMap[boardKey];
      if (!board) {
        var newBoardId = 'B' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + Math.floor(Math.random()*100);
        board = { '基板ID': newBoardId, '機種ID': prod['機種ID'], '基板名': boardName, 'コード': boardCode };
        boardMap[boardKey] = board;
        newBoards.push([newBoardId, prod['機種ID'], boardName, boardCode, '', '']);
        addedBoards++;
      }

      // 部品
      var pId = partCode || ('P' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + Math.floor(Math.random()*100));
      var part = partMap[pId];
      if (!part) {
        part = { '部品ID': pId, '部品名': partName };
        partMap[pId] = part;
        newParts.push([pId, partName, '電子部品', maker, '個', '', 0, 0, 0]);
        addedParts++;
      }

      // BOM
      var bomKey = board['基板ID'] + '_' + pId;
      var existing = bomMap[bomKey];
      if (existing) {
        // 数量更新
        updateBom.push({ id: existing['BOMID'], boardId: board['基板ID'], partId: pId, qty: qty, note: note });
      } else {
        var newBomId = 'BOM' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + Math.floor(Math.random()*100);
        bomMap[bomKey] = { 'BOMID': newBomId };
        newBom.push([newBomId, board['基板ID'], pId, qty, note]);
      }
      updatedBom++;
    });

    // バッチ書き込み（高速）
    if (newProds.length  > 0) prodSheet.getRange(prodSheet.getLastRow()+1,  1, newProds.length,  4).setValues(newProds);
    if (newBoards.length > 0) boardSheet.getRange(boardSheet.getLastRow()+1, 1, newBoards.length, 6).setValues(newBoards);
    if (newParts.length  > 0) partSheet.getRange(partSheet.getLastRow()+1,  1, newParts.length,  9).setValues(newParts);
    if (newBom.length    > 0) bomSheet.getRange(bomSheet.getLastRow()+1,    1, newBom.length,    5).setValues(newBom);

    // BOM数量更新
    updateBom.forEach(function(u) { _upsertRow(bomSheet, u.id, [u.id, u.boardId, u.partId, u.qty, u.note]); });

    return { success: true, addedProducts: addedProducts, addedBoards: addedBoards, addedParts: addedParts, updatedBom: updatedBom };
  } catch(e) {
    Logger.log('[BOM IMPORT ERROR] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

// ===== CSV一括インポート（k10部品表） =====
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

// ID（1列目）でupsert（なければappend、あればupdate）
function _upsertRow(sheet, id, rowData) {
  if (!sheet || !id) return { success: false, error: 'パラメータ不足' };
  var last = sheet.getLastRow();
  if (last > 1) {
    var ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat().map(String);
    var idx = ids.indexOf(String(id));
    if (idx >= 0) {
      sheet.getRange(idx + 2, 1, 1, rowData.length).setValues([rowData]);
      return { success: true, isNew: false };
    }
  }
  sheet.appendRow(rowData);
  return { success: true, isNew: true };
}

// 1列目のIDで行削除
function _deleteRowById(sheet, id) {
  if (!sheet || sheet.getLastRow() <= 1) return;
  var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String);
  for (var i = ids.length - 1; i >= 0; i--) {
    if (ids[i] === String(id)) sheet.deleteRow(i + 2);
  }
}

// 指定列の値で行を複数削除
function _deleteRowsByCol(sheet, colNum, value) {
  if (!sheet || sheet.getLastRow() <= 1) return;
  var vals = sheet.getRange(2, colNum, sheet.getLastRow() - 1, 1).getValues().flat().map(String);
  for (var i = vals.length - 1; i >= 0; i--) {
    if (vals[i] === String(value)) sheet.deleteRow(i + 2);
  }
}