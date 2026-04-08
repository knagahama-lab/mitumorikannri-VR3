// ============================================================
// 統合管理システム
// ファイル 6/6: 基板・部品管理API
// ============================================================

var BOARD_SS_ID = '1OEpET_rvYRFVClVuh9VzRbfcEbnpSFfSZ0ZpM8yybUQ';

var BOARD_CONFIG = {
  SHEET_BOARD:      '基板マスタ',
  SHEET_MACHINE:    '機種マスタ',
  SHEET_COMPONENTS: '部品マスタ',
  SHEET_BOM:        'BOM',
  SHEET_ESTIMATES:  '見積管理',
};

var BOARD_HEADERS = {
  BOARD:      ['基板分類','基板ID','基板名','バージョン','作成日','備考','画像URL','部品表URL'],
  MACHINE:    ['機種コード','機種名','種類','ブランド','発売日','販売台数','M基板','D基板','DE基板','E基板','C基板','S基板','特徴','フォルダURL','備考'],
  COMPONENTS: ['部品コード','部品名','メーカー名','公表単価','仕入先','備考'],
  BOM:        ['基板ID','部品コード','使用数量','実装位置','備考'],
  ESTIMATES:  ['見積No','発行日','対象基板ID','見積種別','見積金額','ステータス','PDFリンク1','PDFリンク2','PDFリンク3','Excelリンク','備考'],
};

function _getBoardSS() {
  return SpreadsheetApp.openById(BOARD_SS_ID);
}

function _getBoardSheetData(sheetName, headers) {
  try {
    var ss     = _getBoardSS();
    var sheet  = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    var values = sheet.getDataRange().getValues();
    if (values.length < 2) return [];
    return values.slice(1).map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) { if (h) obj[h] = row[i] !== undefined ? row[i] : ''; });
      return obj;
    });
  } catch(e) {
    Logger.log('[BOARD SS ERROR] ' + sheetName + ': ' + e.message);
    return [];
  }
}

function apiBoardGetAll() {
  var boards = _getBoardSheetData(BOARD_CONFIG.SHEET_BOARD, BOARD_HEADERS.BOARD);
  return { success: true, items: boards };
}

function apiBoardGetDetail(boardId, boardName) {
  try {
    var bom       = _getBoardBOM(boardId, boardName);
    var estimates = _getBoardSheetData(BOARD_CONFIG.SHEET_ESTIMATES, BOARD_HEADERS.ESTIMATES)
      .filter(function(e) { return String(e['対象基板ID']) === String(boardId); });
    var machine   = _getMachineByBoard(boardId, boardName);
    var totalCost = bom.reduce(function(s, item) { return s + (item['小計'] || 0); }, 0);
    return { success: true, bom: bom, totalCost: totalCost, estimates: estimates, machine: machine };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _getBoardBOM(boardId, boardName) {
  var bomData  = _getBoardSheetData(BOARD_CONFIG.SHEET_BOM, BOARD_HEADERS.BOM);
  var compData = _getBoardSheetData(BOARD_CONFIG.SHEET_COMPONENTS, BOARD_HEADERS.COMPONENTS);
  var bId      = String(boardId   || '').trim();
  var bName    = String(boardName || '').trim();
  var compMap  = {};
  compData.forEach(function(c) {
    var code = String(c['部品コード'] || '').trim();
    if (code) compMap[code] = c;
  });
  return bomData.filter(function(b) {
    var bid = String(b['基板ID'] || '').trim();
    return bid && (bid === bId || bid === bName || bName.indexOf(bid) >= 0 || bId.indexOf(bid) >= 0);
  }).map(function(b) {
    var code = String(b['部品コード'] || '').trim();
    var comp = compMap[code] || {};
    var up   = parseFloat(comp['公表単価']) || 0;
    var qty  = parseFloat(b['使用数量'])   || 0;
    return { '部品コード': code||'コードなし', '部品名': comp['部品名']||'未登録', '使用数量': qty, '公表単価': up, '小計': up * qty };
  });
}

function _getMachineByBoard(boardId, boardName) {
  var machines  = _getBoardSheetData(BOARD_CONFIG.SHEET_MACHINE, BOARD_HEADERS.MACHINE);
  var bId       = String(boardId   || '').trim();
  var bName     = String(boardName || '').trim();
  var boardCols = ['M基板','D基板','DE基板','E基板','C基板','S基板'];
  for (var i = 0; i < machines.length; i++) {
    var m = machines[i];
    for (var j = 0; j < boardCols.length; j++) {
      var val = String(m[boardCols[j]] || '').trim();
      if (val && (val === bId || val === bName || bName.indexOf(val) >= 0 || bId.indexOf(val) >= 0)) return m;
    }
  }
  return null;
}

function apiGetOrdersWithBoardInfo() {
  try {
    var orders   = getAllMgmtData().map(_rowToObject);
    var machines = _getBoardSheetData(BOARD_CONFIG.SHEET_MACHINE, BOARD_HEADERS.MACHINE);
    var boards   = _getBoardSheetData(BOARD_CONFIG.SHEET_BOARD,   BOARD_HEADERS.BOARD);
    var machineMap = {};
    machines.forEach(function(m) { var c = String(m['機種コード']||'').trim(); if(c) machineMap[c]=m; });
    var boardMap = {};
    boards.forEach(function(b) { var id = String(b['基板ID']||'').trim(); if(id) boardMap[id]=b; });
    var enriched = orders.filter(function(o) { return o.orderNo; }).map(function(o) {
      var modelCode = String(o.modelCode||'').trim();
      var machine   = modelCode ? machineMap[modelCode] : null;
      var boardIds  = [];
      if (machine) {
        ['M基板','D基板','DE基板','E基板','C基板','S基板'].forEach(function(col) {
          var bid = String(machine[col]||'').trim();
          if (bid) boardIds.push({ type:col, id:bid, name: boardMap[bid] ? boardMap[bid]['基板名'] : bid });
        });
      }
      return { mgmtId:o.id, orderNo:o.orderNo, orderSlipNo:o.orderSlipNo, client:o.client,
               orderDate:o.orderDate, orderAmount:o.orderAmount, status:o.status,
               orderType:o.orderType, modelCode:modelCode,
               machineName: machine ? String(machine['機種名']||'') : '', boards:boardIds };
    });
    return { success: true, items: enriched };
  } catch(e) {
    Logger.log('[ORDER BOARD ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

function apiGetBoardAnalysis() {
  try {
    var allRows = getAllMgmtData().map(_rowToObject).filter(function(o) { return o.orderNo && o.modelCode; });
    var seenNos = {};
    var orders  = allRows.filter(function(o) {
      var k = String(o.orderNo).trim();
      if (seenNos[k]) return false;
      seenNos[k] = true; return true;
    });
    var machines = _getBoardSheetData(BOARD_CONFIG.SHEET_MACHINE, BOARD_HEADERS.MACHINE);
    var bom      = _getBoardSheetData(BOARD_CONFIG.SHEET_BOM,        BOARD_HEADERS.BOM);
    var comps    = _getBoardSheetData(BOARD_CONFIG.SHEET_COMPONENTS,  BOARD_HEADERS.COMPONENTS);
    var compMap  = {};
    comps.forEach(function(c) { var cd = String(c['部品コード']||'').trim(); if(cd) compMap[cd]=parseFloat(c['公表単価'])||0; });
    var boardCostMap = {};
    bom.forEach(function(b) {
      var bid=String(b['基板ID']||'').trim(), cd=String(b['部品コード']||'').trim();
      boardCostMap[bid]=(boardCostMap[bid]||0)+(compMap[cd]||0)*(parseFloat(b['使用数量'])||0);
    });
    var machineMap = {};
    machines.forEach(function(m) { var c=String(m['機種コード']||'').trim(); if(c) machineMap[c]=m; });
    var summary = {};
    orders.forEach(function(o) {
      var code = String(o.modelCode||'').trim(); if(!code) return;
      if (!summary[code]) {
        var m=machineMap[code]||{}, totalBom=0, boardNames=[];
        ['M基板','D基板','DE基板','E基板','C基板','S基板'].forEach(function(col) {
          var bid=String(m[col]||'').trim(); if(bid){ totalBom+=boardCostMap[bid]||0; boardNames.push(bid); }
        });
        summary[code]={ modelCode:code, machineName:String(m['機種名']||''), orderCount:0, totalAmount:0,
                        totalBomCost:totalBom, boards:boardNames, orders:[] };
      }
      summary[code].orderCount++;
      summary[code].totalAmount += Number(o.orderAmount)||0;
      summary[code].orders.push({ orderNo:o.orderNo, orderDate:o.orderDate, amount:o.orderAmount, status:o.status, orderType:o.orderType });
    });
    var result = Object.values(summary).sort(function(a,b){ return b.totalAmount-a.totalAmount; });
    result.forEach(function(item) {
      item.unitBomCost = item.totalBomCost;
      if (item.totalBomCost>0 && item.totalAmount>0) {
        var est=item.totalBomCost*item.orderCount;
        item.grossProfit=item.totalAmount-est;
        item.grossMarginRate=Math.round((item.grossProfit/item.totalAmount)*100);
        item.bomCoverageNote=item.grossMarginRate>95?'BOMコスト要確認':'';
      } else { item.grossProfit=null; item.grossMarginRate=null; item.bomCoverageNote=''; }
    });
    return { success: true, items: result };
  } catch(e) {
    Logger.log('[BOARD ANALYSIS ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

function apiComparePriceToBOM(mgmtId) {
  try {
    var ss        = getSpreadsheet();
    var qs        = ss.getSheetByName(CONFIG.SHEET_QUOTES);
    var qLast     = qs.getLastRow();
    var quoteLines = [];
    if (qLast > 1) {
      quoteLines = qs.getRange(2,1,qLast-1,15).getValues()
        .filter(function(r){ return String(r[0])===String(mgmtId); })
        .map(function(r){ return { itemName:r[6], spec:r[7], qty:r[8], unit:r[9], unitPrice:r[10], amount:r[11] }; });
    }
    var mgmtSheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    var mgmtRow   = _getMgmtRowById(mgmtSheet, mgmtId);
    if (!mgmtRow) return { success: false, error: '管理IDが見つかりません' };
    var modelCode = String(mgmtRow[MGMT_COLS.MODEL_CODE-1]||'').trim();
    if (!modelCode) return { success: true, comparison: [], note: '機種コードが未設定のため比較不可' };
    var machines = _getBoardSheetData(BOARD_CONFIG.SHEET_MACHINE, BOARD_HEADERS.MACHINE);
    var machine  = machines.find(function(m){ return String(m['機種コード']||'').trim()===modelCode; });
    if (!machine) return { success: true, comparison: [], note: '機種マスタに機種コード「'+modelCode+'」が見つかりません' };
    var comps   = _getBoardSheetData(BOARD_CONFIG.SHEET_COMPONENTS, BOARD_HEADERS.COMPONENTS);
    var compMap = {};
    comps.forEach(function(c){ var cd=String(c['部品コード']||'').trim(); if(cd) compMap[cd]=c; });
    var bomAll      = _getBoardSheetData(BOARD_CONFIG.SHEET_BOM, BOARD_HEADERS.BOM);
    var allBomItems = [];
    ['M基板','D基板','DE基板','E基板','C基板','S基板'].forEach(function(col) {
      var bid=String(machine[col]||'').trim(); if(!bid) return;
      bomAll.filter(function(b){ return String(b['基板ID']||'').trim()===bid; }).forEach(function(b) {
        var cd=String(b['部品コード']||'').trim(), comp=compMap[cd]||{};
        allBomItems.push({ boardType:col, boardId:bid, partCode:cd, partName:comp['部品名']||'未登録',
                           bomQty:parseFloat(b['使用数量'])||0, bomUnitCost:parseFloat(comp['公表単価'])||0 });
      });
    });
    var comparison = quoteLines.map(function(ql) {
      var matched = allBomItems.find(function(b) {
        return b.partName.indexOf(ql.itemName)>=0 || ql.itemName.indexOf(b.partName)>=0 ||
               (b.partCode && ql.spec && ql.spec.indexOf(b.partCode)>=0);
      });
      return { itemName:ql.itemName, spec:ql.spec, quoteQty:ql.qty, quoteUnitPrice:ql.unitPrice,
               bomUnitCost:matched?matched.bomUnitCost:null, bomQty:matched?matched.bomQty:null,
               partCode:matched?matched.partCode:'',
               diff:matched?(ql.unitPrice-matched.bomUnitCost):null,
               diffRate:(matched&&matched.bomUnitCost>0)?Math.round(((ql.unitPrice-matched.bomUnitCost)/matched.bomUnitCost)*100):null };
    });
    return { success: true, modelCode:modelCode, machineName:machine['機種名']||'', comparison:comparison };
  } catch(e) {
    Logger.log('[PRICE COMPARE ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}