// ============================================================
// 統合管理システム
// ファイル 6/6: 基板・部品管理API & AIマッチング統合版
// ============================================================
//
// 【BOM移行について】
//   BOM専用GASプロジェクト（BOM-buhinnhyou）を独立させた場合、
//   以下の BOARD_SS_ID をBOM専用SSのIDに変更するだけで
//   このファイルの他のコードは一切変更不要です。
//
//   変更前: var BOARD_SS_ID = '1OEpET_rvYRFVClVuh9VzRbfcEbnpSFfSZ0ZpM8yybUQ';
//   変更後: var BOARD_SS_ID = '★BOM専用SSのIDをここに設定★';
//
// 【_getMgmtRowById について】
//   この関数は 12_price_compare_v2.gs で定義されています。
//   このファイル内の apiComparePriceToBOM() から参照できます。
// ============================================================

var BOARD_SS_ID = '1OEpET_rvYRFVClVuh9VzRbfcEbnpSFfSZ0ZpM8yybUQ';
// ↑ BOM専用GASを独立させたら、ここをBOM専用SSのIDに変更してください

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

// ============================================================
// BOM・部品管理 API (CRUD)
// ============================================================

function _saveRowToBoardSS(sheetName, headers, uniqueCol, dataObj) {
  try {
    const ss = _getBoardSS();
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const values = sheet.getDataRange().getValues();
    const colIdx = headers.indexOf(uniqueCol);
    const uniqueVal = String(dataObj[uniqueCol] || '');
    let rowIdx = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][colIdx]) === uniqueVal) { rowIdx = i + 1; break; }
    }
    const row = headers.map(h => dataObj[h] !== undefined ? dataObj[h] : '');
    if (rowIdx !== -1) sheet.getRange(rowIdx, 1, 1, headers.length).setValues([row]);
    else sheet.appendRow(row);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function _deleteRowFromBoardSS(sheetName, headers, uniqueCol, uniqueVal) {
  try {
    const ss = _getBoardSS();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: true };
    const values = sheet.getDataRange().getValues();
    const colIdx = headers.indexOf(uniqueCol);
    for (let i = values.length - 1; i >= 1; i--) {
      if (String(values[i][colIdx]) === String(uniqueVal)) sheet.deleteRow(i + 1);
    }
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function apiBoardSavePart(p) { return _saveRowToBoardSS(BOARD_CONFIG.SHEET_COMPONENTS, BOARD_HEADERS.COMPONENTS, '部品コード', p); }
function apiBoardDeletePart(id) { return _deleteRowFromBoardSS(BOARD_CONFIG.SHEET_COMPONENTS, BOARD_HEADERS.COMPONENTS, '部品コード', id); }
function apiBoardGetParts() { return { success: true, items: _getBoardSheetData(BOARD_CONFIG.SHEET_COMPONENTS, BOARD_HEADERS.COMPONENTS) }; }

function apiBoardSaveMachine(m) { return _saveRowToBoardSS(BOARD_CONFIG.SHEET_MACHINE, BOARD_HEADERS.MACHINE, '機種コード', m); }
function apiBoardDeleteMachine(id) { return _deleteRowFromBoardSS(BOARD_CONFIG.SHEET_MACHINE, BOARD_HEADERS.MACHINE, '機種コード', id); }
function apiBoardGetMachines() { return { success: true, items: _getBoardSheetData(BOARD_CONFIG.SHEET_MACHINE, BOARD_HEADERS.MACHINE) }; }

function apiBoardSaveBoard(b) { return _saveRowToBoardSS(BOARD_CONFIG.SHEET_BOARD, BOARD_HEADERS.BOARD, '基板ID', b); }
function apiBoardDeleteBoard(id) { return _deleteRowFromBoardSS(BOARD_CONFIG.SHEET_BOARD, BOARD_HEADERS.BOARD, '基板ID', id); }

function apiBoardSaveBOM(boardId, lines) {
  try {
    const ss = _getBoardSS();
    const sheet = ss.getSheetByName(BOARD_CONFIG.SHEET_BOM);
    const values = sheet.getDataRange().getValues();
    for (let i = values.length - 1; i >= 1; i--) {
      if (String(values[i][0]) === String(boardId)) sheet.deleteRow(i + 1);
    }
    if (lines.length > 0) {
      const headers = BOARD_HEADERS.BOM;
      const data = lines.map(line => headers.map(h => line[h] !== undefined ? line[h] : ''));
      sheet.getRange(sheet.getLastRow() + 1, 1, data.length, headers.length).setValues(data);
    }
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ============================================================
// 🚀 CSVの一括インポート (配列処理による高速化版)
// ============================================================
function apiBoardImportBOMCSV(csvText) {
  try {
    if (!csvText) return { success: false, error: 'データがありません' };
    const lines = String(csvText).split(/\r\n|\n/).filter(l => l.trim() !== '');
    if (lines.length < 2) return { success: false, error: 'データ行がありません' };

    let addedProducts = 0, addedBoards = 0, addedParts = 0, updatedBOM = 0;

    const ss = _getBoardSS();
    const machineSheet = ss.getSheetByName(BOARD_CONFIG.SHEET_MACHINE);
    const boardSheet   = ss.getSheetByName(BOARD_CONFIG.SHEET_BOARD);
    const compSheet    = ss.getSheetByName(BOARD_CONFIG.SHEET_COMPONENTS);
    const bomSheet     = ss.getSheetByName(BOARD_CONFIG.SHEET_BOM);

    const getMap = (sheet, keyCol) => {
      const m = {};
      if (sheet && sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
             .getValues().forEach(row => m[String(row[keyCol]).trim()] = true);
      }
      return m;
    };
    
    const existingMachines = getMap(machineSheet, 0); 
    const existingBoards   = getMap(boardSheet, 1);   
    const existingParts    = getMap(compSheet, 0);    

    const newMachineRows = [];
    const newBoardRows   = [];
    const newPartRows    = [];
    const newBOMRows     = [];
    const existingBomSet = new Set(); 

    for (let i = 1; i < lines.length; i++) {
      const row = _parseCSVLineFast(lines[i]);
      if (row.length < 11) continue;

      const prodCode    = row[0]; 
      const prodName    = row[1]; 
      const boardCode   = row[3]; 
      const boardName   = row[4]; 
      const partCode    = row[5]; 
      const partName    = row[6]; 
      const makerPartNo = row[7]; 
      const maker       = row[9]; 
      const qty         = Number(row[10]) || 0; 
      const note        = row.length > 15 ? row[15] : ''; 

      if (!prodCode || !boardName || (!partCode && !partName)) continue;

      if (!existingMachines[prodCode]) {
        const mRow = BOARD_HEADERS.MACHINE.map(() => '');
        mRow[0] = prodCode; mRow[1] = prodName;
        newMachineRows.push(mRow);
        existingMachines[prodCode] = true;
        addedProducts++;
      }

      const bId = boardCode || (prodCode + '-' + boardName.substring(0,4));
      if (!existingBoards[bId]) {
        const bRow = BOARD_HEADERS.BOARD.map(() => '');
        bRow[1] = bId; bRow[2] = boardName; bRow[5] = 'CSV自動作成';
        newBoardRows.push(bRow);
        existingBoards[bId] = true;
        addedBoards++;
      }

      const pId = partCode || makerPartNo;
      if (pId && !existingParts[pId]) {
        const pRow = BOARD_HEADERS.COMPONENTS.map(() => '');
        pRow[0] = pId; pRow[1] = partName; pRow[2] = maker; pRow[3] = 0; pRow[5] = makerPartNo;
        newPartRows.push(pRow);
        existingParts[pId] = true;
        addedParts++;
      }

      const bomKey = bId + '|' + pId;
      if (!existingBomSet.has(bomKey)) {
        newBOMRows.push([bId, pId, qty, '', note]);
        existingBomSet.add(bomKey);
        updatedBOM++;
      }
    }

    if (newMachineRows.length > 0) machineSheet.getRange(machineSheet.getLastRow() + 1, 1, newMachineRows.length, newMachineRows[0].length).setValues(newMachineRows);
    if (newBoardRows.length > 0)   boardSheet.getRange(boardSheet.getLastRow() + 1, 1, newBoardRows.length, newBoardRows[0].length).setValues(newBoardRows);
    if (newPartRows.length > 0)    compSheet.getRange(compSheet.getLastRow() + 1, 1, newPartRows.length, newPartRows[0].length).setValues(newPartRows);
    if (newBOMRows.length > 0)     bomSheet.getRange(bomSheet.getLastRow() + 1, 1, newBOMRows.length, newBOMRows[0].length).setValues(newBOMRows);

    return { success: true, addedProducts, addedBoards, addedParts, updatedBOM };

  } catch(e) {
    Logger.log("BOMインポートエラー: " + e.message);
    return { success: false, error: e.message };
  }
}

function _parseCSVLineFast(line) {
  const result = [];
  let current = '', inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    if (char === '"' && line[i+1] === '"') { current += '"'; i++; } 
    else if (char === '"') { inQuotes = !inQuotes; } 
    else if (char === ',' && !inQuotes) { result.push(current.trim()); current = ''; } 
    else { current += char; }
  }
  result.push(current.trim());
  return result;
}


// ============================================================
// 🤖 Gemini API を使用した高精度マッチング (安全装置つき・2.5Flash対応)
// ============================================================
function matchWithGeminiAPI(orderData, quotesList) {
  let apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) {
    Logger.log("GEMINI_API_KEY が設定されていません。");
    return [];
  }
  
  apiKey = apiKey.trim();
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;

  const prompt = `
あなたは企業の優秀な営業事務アシスタントです。
新たに受領した以下の「注文書データ」に最も合致する「見積書」を、後述の「見積書候補リスト」から見つけ出し、それぞれにスコア（0〜100点）をつけてください。

【判定基準】
・会社名の表記ゆれ（株式会社の有無、前株/後株、略称）は同一とみなす。
・金額のズレ（税抜と税込の10%の違い、数円の端数）は同一とみなす。
・OCR特有の誤字（0とO、1とIなど）を推測してカバーする。
・品名や機種コードが部分一致していれば加点する。
・見積番号が注文書の備考等に含まれていれば100点とする。

【重要】
合致しそうな見積書が1つもない場合は、必ず空の配列 [] のみを出力してください。
出力は必ず以下のフォーマットのJSON配列のみとし、それ以外の文章やマークダウン記法(\`\`\`json など)は一切含めないでください。

[
  {
    "quoteId": "見積書の管理ID",
    "score": 95,
    "reason": "金額（税抜）が一致し、宛先の表記ゆれも同一企業とみなせるため。"
  }
]

【注文書データ】
${JSON.stringify(orderData, null, 2)}

【見積書候補リスト】
${JSON.stringify(quotesList, null, 2)}
  `;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.1,
      response_mime_type: "application/json"
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    
    // API自体がエラーを返した場合の処理
    if (!responseText) return [];
    const json = JSON.parse(responseText);
    
    if (json.error) {
      Logger.log("Gemini API Error: " + json.error.message);
      return [];
    }
    
    if (!json.candidates || !json.candidates[0] || !json.candidates[0].content) {
      return [];
    }
    
    let resultText = json.candidates[0].content.parts[0].text;
    
    // 万が一空っぽの場合は空配列を返す
    if (!resultText || resultText.trim() === "") {
      return [];
    }
    
    // マークダウン記号を取り除く
    resultText = resultText.replace(/```json/gi, '').replace(/```/g, '').trim();
    
    // パース（データ化）で失敗してもシステムを止めない安全装置
    try {
      return JSON.parse(resultText);
    } catch (parseError) {
      Logger.log("JSONパースエラー発生。AIの生の出力: " + resultText);
      return []; // エラーになっても空として扱い、処理を続行する
    }

  } catch (e) {
    Logger.log("AIマッチング実行失敗: " + e.message);
    return [];
  }
}