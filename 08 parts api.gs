// ============================================================
// 見積・注文 管理システム
// ファイル 8/8: 部品・PCB 原価管理 API
// ============================================================
//
// 対象スプレッドシート: 基板管理SS (BOARD_SS_ID)
//   シート「部品マスタ」  … 部品コード / 部品名 / メーカー / 原価 / 公表単価 / 備考
//   シート「PCBマスタ」   … PCB品番 / PCB名 / 基板ID / 原価 / 公表単価 / 備考
//
// API アクション一覧
//   partsGetAll        全部品取得（検索フィルタ対応）
//   partsSave          1件保存（新規 or 上書き）
//   partsDelete        1件削除
//   partsImportCSV     CSVデータを一括インポート（キーで upsert）
//   partsExportCSV     全データをCSV文字列で返す
//   pcbGetAll          全PCB取得
//   pcbSave            PCB 1件保存
//   pcbDelete          PCB 1件削除
//   pcbImportCSV       PCB CSV一括インポート
//   pcbExportCSV       PCB CSVエクスポート
// ============================================================

// ===== シート名・列定義 =====

var PARTS_SHEET_NAME = '部品マスタ';
var PCB_SHEET_NAME   = 'PCBマスタ';

// 部品マスタ列（1-indexed）
var PARTS_COL = {
  CODE:       1,  // 部品コード（主キー）
  NAME:       2,  // 部品名
  MAKER:      3,  // メーカー名
  COST:       4,  // 原価（仕入れ単価）
  LIST_PRICE: 5,  // 公表単価（売値）
  UNIT:       6,  // 単位
  REMARKS:    7,  // 備考
  UPDATED_AT: 8,  // 最終更新日
};
var PARTS_HEADERS = ['部品コード','部品名','メーカー名','原価','公表単価','単位','備考','最終更新日'];

// PCBマスタ列（1-indexed）
var PCB_COL = {
  CODE:       1,  // PCB品番（主キー）
  NAME:       2,  // PCB名
  BOARD_ID:   3,  // 対応基板ID
  COST:       4,  // 原価
  LIST_PRICE: 5,  // 公表単価
  MAKER:      6,  // 製造メーカー
  UNIT:       7,  // 単位
  REMARKS:    8,  // 備考
  UPDATED_AT: 9,  // 最終更新日
};
var PCB_HEADERS = ['PCB品番','PCB名','対応基板ID','原価','公表単価','製造メーカー','単位','備考','最終更新日'];

// ============================================================
// シート初期化
// ============================================================

function ensurePartsSheets() {
  var ss = _getBoardSS();

  // 部品マスタ
  if (!ss.getSheetByName(PARTS_SHEET_NAME)) {
    var s = ss.insertSheet(PARTS_SHEET_NAME);
    var hr = s.getRange(1, 1, 1, PARTS_HEADERS.length);
    hr.setValues([PARTS_HEADERS]);
    hr.setBackground('#E8F0FE');
    hr.setFontWeight('bold');
    s.setFrozenRows(1);
    s.setColumnWidth(1, 160);
    s.setColumnWidth(2, 220);
    s.setColumnWidth(3, 160);
    s.setColumnWidth(4, 100);
    s.setColumnWidth(5, 100);
    s.setColumnWidth(6, 70);
    s.setColumnWidth(7, 200);
    s.setColumnWidth(8, 140);
    Logger.log('[PARTS] 部品マスタシート作成');
  }

  // PCBマスタ
  if (!ss.getSheetByName(PCB_SHEET_NAME)) {
    var p = ss.insertSheet(PCB_SHEET_NAME);
    var pr = p.getRange(1, 1, 1, PCB_HEADERS.length);
    pr.setValues([PCB_HEADERS]);
    pr.setBackground('#FEF7E0');
    pr.setFontWeight('bold');
    p.setFrozenRows(1);
    p.setColumnWidth(1, 160);
    p.setColumnWidth(2, 220);
    p.setColumnWidth(3, 120);
    p.setColumnWidth(4, 100);
    p.setColumnWidth(5, 100);
    p.setColumnWidth(6, 160);
    p.setColumnWidth(7, 70);
    p.setColumnWidth(8, 200);
    p.setColumnWidth(9, 140);
    Logger.log('[PARTS] PCBマスタシート作成');
  }
}

// ============================================================
// 部品マスタ API
// ============================================================

function _apiPartsGetAll(p) {
  try {
    ensurePartsSheets();
    var ss    = _getBoardSS();
    var sheet = ss.getSheetByName(PARTS_SHEET_NAME);
    var items = _readPartsSheet(sheet, PARTS_COL, PARTS_HEADERS.length);
    var kw    = String(p.keyword || '').toLowerCase().trim();
    if (kw) {
      items = items.filter(function(r) {
        return [r.code, r.name, r.maker, r.remarks].some(function(v) {
          return String(v || '').toLowerCase().indexOf(kw) >= 0;
        });
      });
    }
    return { success: true, total: items.length, items: items };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _apiPartsSave(p) {
  try {
    ensurePartsSheets();
    var ss    = _getBoardSS();
    var sheet = ss.getSheetByName(PARTS_SHEET_NAME);
    var code  = String(p.code || '').trim();
    if (!code) return { success: false, error: '部品コードは必須です' };

    var rowData = [
      code,
      p.name       || '',
      p.maker      || '',
      isNaN(Number(p.cost))      ? '' : Number(p.cost),
      isNaN(Number(p.listPrice)) ? '' : Number(p.listPrice),
      p.unit       || '',
      p.remarks    || '',
      nowJST(),
    ];

    var existRow = _findRowByKey(sheet, code);
    if (existRow > 0) {
      sheet.getRange(existRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    return { success: true, code: code };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _apiPartsDelete(p) {
  try {
    var ss    = _getBoardSS();
    var sheet = ss.getSheetByName(PARTS_SHEET_NAME);
    if (!sheet) return { success: false, error: 'シートなし' };
    var code = String(p.code || '').trim();
    var row  = _findRowByKey(sheet, code);
    if (row < 0) return { success: false, error: '部品コードが見つかりません' };
    sheet.deleteRow(row);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _apiPartsImportCSV(p) {
  try {
    ensurePartsSheets();
    if (!p.csvText) return { success: false, error: 'CSVデータが空です' };
    var rows   = _parseCSV(p.csvText);
    if (rows.length === 0) return { success: false, error: '有効な行がありません' };

    var ss    = _getBoardSS();
    var sheet = ss.getSheetByName(PARTS_SHEET_NAME);

    // ヘッダー行を除去（1行目が部品コードでなければヘッダーとみなす）
    var dataRows = rows;
    if (rows.length > 0 && String(rows[0][0]).replace(/\s/g,'') === '部品コード') {
      dataRows = rows.slice(1);
    }

    var updated = 0; var added = 0; var skipped = 0;
    var now = nowJST();

    dataRows.forEach(function(cols) {
      var code = String(cols[0] || '').trim();
      if (!code) { skipped++; return; }

      var rowData = [
        code,
        cols[1] || '',                                   // 部品名
        cols[2] || '',                                   // メーカー
        isNaN(Number(cols[3])) ? '' : Number(cols[3]),   // 原価
        isNaN(Number(cols[4])) ? '' : Number(cols[4]),   // 公表単価
        cols[5] || '',                                   // 単位
        cols[6] || '',                                   // 備考
        now,
      ];

      var existRow = _findRowByKey(sheet, code);
      if (existRow > 0) {
        sheet.getRange(existRow, 1, 1, rowData.length).setValues([rowData]);
        updated++;
      } else {
        sheet.appendRow(rowData);
        added++;
      }
    });

    return { success: true, added: added, updated: updated, skipped: skipped };
  } catch(e) {
    Logger.log('[PARTS IMPORT ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

function _apiPartsExportCSV(p) {
  try {
    var ss    = _getBoardSS();
    var sheet = ss.getSheetByName(PARTS_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, csv: PARTS_HEADERS.join(',') + '\n' };
    var data = sheet.getRange(1, 1, sheet.getLastRow(), PARTS_HEADERS.length).getValues();
    var csv  = data.map(function(row) {
      return row.map(function(v) {
        var s = String(v === null || v === undefined ? '' : v);
        if (s.indexOf(',') >= 0 || s.indexOf('"') >= 0 || s.indexOf('\n') >= 0) {
          s = '"' + s.replace(/"/g, '""') + '"';
        }
        return s;
      }).join(',');
    }).join('\n');
    return { success: true, csv: csv, filename: '部品マスタ_' + _dateStamp() + '.csv' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// PCBマスタ API
// ============================================================

function _apiPcbGetAll(p) {
  try {
    ensurePartsSheets();
    var ss    = _getBoardSS();
    var sheet = ss.getSheetByName(PCB_SHEET_NAME);
    var items = _readPcbSheet(sheet);
    var kw    = String(p.keyword || '').toLowerCase().trim();
    if (kw) {
      items = items.filter(function(r) {
        return [r.code, r.name, r.boardId, r.maker, r.remarks].some(function(v) {
          return String(v || '').toLowerCase().indexOf(kw) >= 0;
        });
      });
    }
    return { success: true, total: items.length, items: items };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _apiPcbSave(p) {
  try {
    ensurePartsSheets();
    var ss    = _getBoardSS();
    var sheet = ss.getSheetByName(PCB_SHEET_NAME);
    var code  = String(p.code || '').trim();
    if (!code) return { success: false, error: 'PCB品番は必須です' };

    var rowData = [
      code,
      p.name      || '',
      p.boardId   || '',
      isNaN(Number(p.cost))      ? '' : Number(p.cost),
      isNaN(Number(p.listPrice)) ? '' : Number(p.listPrice),
      p.maker     || '',
      p.unit      || '',
      p.remarks   || '',
      nowJST(),
    ];

    var existRow = _findRowByKey(sheet, code);
    if (existRow > 0) {
      sheet.getRange(existRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    return { success: true, code: code };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _apiPcbDelete(p) {
  try {
    var ss    = _getBoardSS();
    var sheet = ss.getSheetByName(PCB_SHEET_NAME);
    if (!sheet) return { success: false, error: 'シートなし' };
    var row = _findRowByKey(sheet, String(p.code || '').trim());
    if (row < 0) return { success: false, error: 'PCB品番が見つかりません' };
    sheet.deleteRow(row);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _apiPcbImportCSV(p) {
  try {
    ensurePartsSheets();
    if (!p.csvText) return { success: false, error: 'CSVデータが空です' };
    var rows = _parseCSV(p.csvText);
    if (rows.length === 0) return { success: false, error: '有効な行がありません' };

    var ss    = _getBoardSS();
    var sheet = ss.getSheetByName(PCB_SHEET_NAME);

    var dataRows = rows;
    if (rows.length > 0 && String(rows[0][0]).replace(/\s/g,'') === 'PCB品番') {
      dataRows = rows.slice(1);
    }

    var updated = 0; var added = 0; var skipped = 0;
    var now = nowJST();

    dataRows.forEach(function(cols) {
      var code = String(cols[0] || '').trim();
      if (!code) { skipped++; return; }

      var rowData = [
        code,
        cols[1] || '',
        cols[2] || '',
        isNaN(Number(cols[3])) ? '' : Number(cols[3]),
        isNaN(Number(cols[4])) ? '' : Number(cols[4]),
        cols[5] || '',
        cols[6] || '',
        cols[7] || '',
        now,
      ];

      var existRow = _findRowByKey(sheet, code);
      if (existRow > 0) {
        sheet.getRange(existRow, 1, 1, rowData.length).setValues([rowData]);
        updated++;
      } else {
        sheet.appendRow(rowData);
        added++;
      }
    });

    return { success: true, added: added, updated: updated, skipped: skipped };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _apiPcbExportCSV(p) {
  try {
    var ss    = _getBoardSS();
    var sheet = ss.getSheetByName(PCB_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, csv: PCB_HEADERS.join(',') + '\n' };
    var data = sheet.getRange(1, 1, sheet.getLastRow(), PCB_HEADERS.length).getValues();
    var csv  = data.map(function(row) {
      return row.map(function(v) {
        var s = String(v === null || v === undefined ? '' : v);
        if (s.indexOf(',') >= 0 || s.indexOf('"') >= 0 || s.indexOf('\n') >= 0) {
          s = '"' + s.replace(/"/g, '""') + '"';
        }
        return s;
      }).join(',');
    }).join('\n');
    return { success: true, csv: csv, filename: 'PCBマスタ_' + _dateStamp() + '.csv' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// ユーティリティ
// ============================================================

// シートの1列目（主キー列）からキーで行番号を返す（なければ -1）
function _findRowByKey(sheet, key) {
  var last = sheet.getLastRow();
  if (last <= 1) return -1;
  var keys = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
  var idx  = keys.map(String).indexOf(String(key));
  return idx >= 0 ? idx + 2 : -1;
}

// 部品マスタ行 → オブジェクト
function _readPartsSheet(sheet, colDef, colCount) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, colCount).getValues()
    .filter(function(r) { return String(r[0]).trim() !== ''; })
    .map(function(r) {
      var cost      = parseFloat(r[colDef.COST - 1])       || 0;
      var listPrice = parseFloat(r[colDef.LIST_PRICE - 1]) || 0;
      return {
        code:       String(r[0] || ''),
        name:       String(r[1] || ''),
        maker:      String(r[2] || ''),
        cost:       cost,
        listPrice:  listPrice,
        unit:       String(r[5] || ''),
        remarks:    String(r[6] || ''),
        updatedAt:  String(r[7] || ''),
        margin:     (cost > 0 && listPrice > 0)
                      ? Math.round((listPrice - cost) / listPrice * 100)
                      : null,
      };
    });
}

// PCBマスタ行 → オブジェクト
function _readPcbSheet(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, PCB_HEADERS.length).getValues()
    .filter(function(r) { return String(r[0]).trim() !== ''; })
    .map(function(r) {
      var cost      = parseFloat(r[3]) || 0;
      var listPrice = parseFloat(r[4]) || 0;
      return {
        code:       String(r[0] || ''),
        name:       String(r[1] || ''),
        boardId:    String(r[2] || ''),
        cost:       cost,
        listPrice:  listPrice,
        maker:      String(r[5] || ''),
        unit:       String(r[6] || ''),
        remarks:    String(r[7] || ''),
        updatedAt:  String(r[8] || ''),
        margin:     (cost > 0 && listPrice > 0)
                      ? Math.round((listPrice - cost) / listPrice * 100)
                      : null,
      };
    });
}

// CSVテキスト → 2次元配列（ダブルクォート対応）
function _parseCSV(text) {
  // BOM除去
  text = text.replace(/^\uFEFF/, '');
  var lines  = [];
  var line   = [];
  var field  = '';
  var inQ    = false;
  var i      = 0;
  var chars  = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

  while (i < chars.length) {
    var c = chars[i];
    if (inQ) {
      if (c === '"') {
        if (chars[i + 1] === '"') { field += '"'; i += 2; continue; }
        inQ = false;
      } else {
        field += c;
      }
    } else {
      if (c === '"') {
        inQ = true;
      } else if (c === ',') {
        line.push(field); field = '';
      } else if (c === '\n') {
        line.push(field); field = '';
        if (line.some(function(v) { return String(v).trim() !== ''; })) {
          lines.push(line);
        }
        line = [];
      } else {
        field += c;
      }
    }
    i++;
  }
  // 最後の行
  if (field !== '' || line.length > 0) {
    line.push(field);
    if (line.some(function(v) { return String(v).trim() !== ''; })) {
      lines.push(line);
    }
  }
  return lines;
}

function _dateStamp() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmm');
}