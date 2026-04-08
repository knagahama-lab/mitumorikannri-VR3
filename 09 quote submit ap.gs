// ============================================================
// 見積・注文 管理システム
// ファイル 9/9: 見積提出管理 API（機種別・基板別 見積書追跡）
// ============================================================
//
// シート「見積提出管理」の列定義:
//   A: ID  B: 機種コード  C: ブランド  D: 基板名
//   E: フォーマット  F: 金額  G: 備考  H: 見積No
//   I: 提出先  J: 提出日  K: ファイルURL  L: 最終更新日
// ============================================================

var QS_SHEET_NAME = '見積提出管理';

var QS_COL = {
  ID:          1,
  MACHINE:     2,
  BRAND:       3,
  BOARD:       4,
  FORMAT:      5,
  AMOUNT:      6,
  REMARKS:     7,
  QUOTE_NO:    8,
  SUBMIT_TO:   9,
  SUBMIT_DATE: 10,
  FILE_URL:    11,
  UPDATED_AT:  12,
  PARENT_ID:   13,
  IS_LATEST:   14,
  STATUS:      15
};

var QS_HEADERS = [
  'ID','機種コード','ブランド','基板名','フォーマット',
  '金額','備考','見積No','提出先','提出日','ファイルURL','最終更新日',
  '親ID','最新フラグ','ステータス'
];

// ============================================================
// シート初期化 + 初期データ投入
// ============================================================

function initQuoteSubmitSheet() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(QS_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(QS_SHEET_NAME);
    Logger.log('[QS] シート新規作成');
  }

  // ヘッダー設定
  var hr = sheet.getRange(1, 1, 1, QS_HEADERS.length);
  hr.setValues([QS_HEADERS]);
  hr.setBackground('#E8F0FE');
  hr.setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 90);
  sheet.setColumnWidth(7, 200);
  sheet.setColumnWidth(8, 160);
  sheet.setColumnWidth(9, 100);
  sheet.setColumnWidth(10, 110);
  sheet.setColumnWidth(11, 300);
  sheet.setColumnWidth(12, 140);
  sheet.setColumnWidth(13, 140);
  sheet.setColumnWidth(14, 100);
  sheet.setColumnWidth(15, 100);

  Logger.log('[QS] シート初期化完了');
  return sheet;
}

// 初期データ投入（Excelから抽出した355件）
function importQuoteSubmitInitialData() {
  var sheet = initQuoteSubmitSheet();
  if (sheet.getLastRow() > 1) {
    Logger.log('[QS] データ既存のため投入スキップ（行数: ' + sheet.getLastRow() + '）');
    SpreadsheetApp.getUi().alert('データが既に登録されています（' + (sheet.getLastRow()-1) + '件）。\n再投入する場合はシートのデータを手動で削除してから実行してください。');
    return;
  }

  var data = _getInitialData();
  var now  = nowJST();
  var rows = data.map(function(d) {
    return [
      d.id, d.machine_code, d.brand, d.board_name, d.format,
      d.amount || '', d.remarks, d.quote_no, d.submit_to, d.submit_date,
      d.file_url, now
    ];
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, QS_HEADERS.length).setValues(rows);
  }
  Logger.log('[QS] 初期データ投入完了: ' + rows.length + '件');
  SpreadsheetApp.getUi().alert('見積提出管理データを ' + rows.length + '件 投入しました！');
}

// ============================================================
// API
// ============================================================

function _apiQsGetAll(p) {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(QS_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, items: [] };

    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, QS_HEADERS.length).getValues()
      .filter(function(r) { return String(r[0]).trim() !== ''; });

    var items = rows.map(_qsRowToObj);

    // フィルタ
    var kw      = String(p.keyword  || '').toLowerCase().trim();
    var machine = String(p.machine  || '').trim();
    var board   = String(p.board    || '').trim();

    if (kw) {
      items = items.filter(function(r) {
        return [r.machineCode, r.boardName, r.quoteNo, r.submitTo, r.remarks, r.brand].some(function(v) {
          return String(v||'').toLowerCase().indexOf(kw) >= 0;
        });
      });
    }
    if (machine) {
      items = items.filter(function(r) { return r.machineCode === machine; });
    }
    if (board) {
      items = items.filter(function(r) { return String(r.boardName||'').indexOf(board) >= 0; });
    }

    // 提出日の新しい順
    items.sort(function(a, b) {
      return String(b.submitDate||'').localeCompare(String(a.submitDate||''));
    });

    // 機種コード一覧も返す
    var allRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    var machines = [];
    var seen = {};
    allRows.forEach(function(r) {
      var m = String(r[1]||'').trim();
      if (m && !seen[m]) { seen[m] = true; machines.push(m); }
    });
    machines.sort();

    return { success: true, total: items.length, items: items, machines: machines };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _apiQsSave(p) {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(QS_SHEET_NAME);
    if (!sheet) {
      initQuoteSubmitSheet();
      sheet = ss.getSheetByName(QS_SHEET_NAME);
    }

    var isNew = !p.id || p.id === '';
    var id    = isNew ? _generateQsId(p.machineCode) : p.id;
    var now   = nowJST();

    var parentId = p.parentId || id; // 新規の場合は自分自身を親とする
    
    // 【版管理: 旧リビジョンをFALSEにする】
    var currentData = sheet.getDataRange().getValues();
    if (currentData.length > 1) {
      for (var i = 1; i < currentData.length; i++) {
        var existingParent = String(currentData[i][QS_COL.PARENT_ID - 1] || currentData[i][QS_COL.ID - 1]);
        if (existingParent === parentId && currentData[i][QS_COL.IS_LATEST - 1] === true) {
          sheet.getRange(i + 1, QS_COL.IS_LATEST).setValue(false);
          sheet.getRange(i + 1, QS_COL.STATUS).setValue('改定');
        }
      }
    }

    var rowData = [
      id,
      p.machineCode  || '',
      p.brand        || '',
      p.boardName    || '',
      p.format       || '',
      (p.amount !== undefined && p.amount !== '' && !isNaN(Number(p.amount))) ? Number(p.amount) : '',
      p.remarks      || '',
      p.quoteNo      || '',
      p.submitTo     || '',
      p.submitDate   || '',
      p.fileUrl      || '',
      now,
      parentId,
      true,
      p.status || '有効'
    ];

    if (isNew) {
      sheet.appendRow(rowData);
    } else {
      var row = _qsFindRowById(sheet, id);
      if (row < 0) return { success: false, error: 'IDが見つかりません: ' + id };
      sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
    }
    return { success: true, id: id };
  } catch(e) {
    Logger.log('[QS SAVE ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

function _apiQsDelete(p) {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(QS_SHEET_NAME);
    if (!sheet || !p.id) return { success: false, error: 'パラメータ不足' };
    var row = _qsFindRowById(sheet, p.id);
    if (row < 0) return { success: false, error: 'IDが見つかりません' };
    sheet.deleteRow(row);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _apiQsUploadFile(p) {
  try {
    if (!p.base64Data || !p.fileName) return { success: false, error: 'ファイルデータ不足' };
    var folder   = DriveApp.getFolderById(CONFIG.WEB_UPLOAD_FOLDER_ID);
    var mimeType = p.mimeType || 'application/pdf';
    var prefix   = p.machineCode ? p.machineCode + '_' : '';
    var blob     = Utilities.newBlob(Utilities.base64Decode(p.base64Data), mimeType, prefix + p.fileName);
    var file     = folder.createFile(blob);
    return { success: true, url: file.getUrl(), fileId: file.getId(), fileName: p.fileName };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _apiQsGetMachines() {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(QS_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, machines: [] };
    var col = sheet.getRange(2, QS_COL.MACHINE, sheet.getLastRow() - 1, 1).getValues().flat();
    var seen = {};
    var machines = [];
    col.forEach(function(v) {
      var m = String(v||'').trim();
      if (m && !seen[m]) { seen[m] = true; machines.push(m); }
    });
    return { success: true, machines: machines.sort() };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// ユーティリティ
// ============================================================

function _qsRowToObj(row) {
  return {
    id:          String(row[0]  || ''),
    machineCode: String(row[1]  || ''),
    brand:       String(row[2]  || ''),
    boardName:   String(row[3]  || ''),
    format:      String(row[4]  || ''),
    amount:      row[5] !== '' && row[5] !== null ? Number(row[5]) : null,
    remarks:     String(row[6]  || ''),
    quoteNo:     String(row[7]  || ''),
    submitTo:    String(row[8]  || ''),
    submitDate:  String(row[9]  || ''),
    fileUrl:     String(row[10] || ''),
    updatedAt:   String(row[11] || ''),
    parentId:    String(row[12] || ''),
    isLatest:    row[13] !== undefined ? row[13] : true,
    status:      String(row[14] || '')
  };
}

function _qsFindRowById(sheet, id) {
  var last = sheet.getLastRow();
  if (last <= 1) return -1;
  var ids = sheet.getRange(2, QS_COL.ID, last - 1, 1).getValues().flat();
  var idx = ids.map(String).indexOf(String(id));
  return idx >= 0 ? idx + 2 : -1;
}

function _generateQsId(machineCode) {
  var code = (machineCode || 'QS').replace(/[^A-Za-z0-9]/g, '').substring(0, 6);
  return 'QS-' + code + '-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');
}

function _getInitialData() {
  return [
      {id:'QS-E63-0001',machine_code:'E63',brand:'FJ（）',board_name:'FJ+M2402A PCB',format:'新提案',amount:659.0,remarks:'',quote_no:'製見2024-0154',submit_to:'大山さん',submit_date:'2025/01/10',file_url:''},
      {id:'QS-E63-0002',machine_code:'E63',brand:'FJ（）',board_name:'FJ+M2402A仕掛基板',format:'新提案',amount:3452.0,remarks:'',quote_no:'No.2025-013',submit_to:'鎌田さん',submit_date:'',file_url:''},
      {id:'QS-E63-0003',machine_code:'E63',brand:'FJ（）',board_name:'FJ+M2402A（完）',format:'新提案',amount:708.0,remarks:'',quote_no:'No.2025-137',submit_to:'大山さん',submit_date:'2026/02/13',file_url:''},
      {id:'QS-E63-0004',machine_code:'E63',brand:'FJ（）',board_name:'オール新品',format:'新提案',amount:1244.0,remarks:'各仕掛基板、PCB含まず',quote_no:'No.2025-138',submit_to:'大山さん',submit_date:'2026/02/13',file_url:''},
      {id:'QS-E62-0005',machine_code:'E62',brand:'FJ（）',board_name:'FJ+M2402A PCB',format:'新提案',amount:659.0,remarks:'',quote_no:'製見2024-0154',submit_to:'大山さん',submit_date:'2025/01/10',file_url:''},
      {id:'QS-E62-0006',machine_code:'E62',brand:'FJ（）',board_name:'FJ+M2402A仕掛基板',format:'新提案',amount:3452.0,remarks:'',quote_no:'No.2025-013',submit_to:'鎌田さん',submit_date:'',file_url:''},
      {id:'QS-E62-0007',machine_code:'E62',brand:'FJ（）',board_name:'FJ+M2402A（完）',format:'新提案',amount:708.0,remarks:'',quote_no:'No.2025-020',submit_to:'大山さん',submit_date:'2025/05/28',file_url:''},
      {id:'QS-E62-0008',machine_code:'E62',brand:'FJ（）',board_name:'オール新品',format:'新提案',amount:16942.0,remarks:'E2102B仕掛、PCB含まず',quote_no:'No.2025-077',submit_to:'大山さん',submit_date:'2025/10/17',file_url:''},
      {id:'QS-E62-0009',machine_code:'E62',brand:'FJ（）',board_name:'見本機（D,DE中古）',format:'新提案',amount:1444.0,remarks:'E2102B仕掛、PCB含まず',quote_no:'No.2025-127',submit_to:'大山さん',submit_date:'2026/02/05',file_url:''},
      {id:'QS-PCB-0010',machine_code:'PCB',brand:'新提案　→管理費を固定、利益を12％に設定した見積書',board_name:'FJ+E2101B PCB',format:'',amount:1098.0,remarks:'',quote_no:'2025-059',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-PCB-0011',machine_code:'PCB',brand:'新提案　→管理費を固定、利益を12％に設定した見積書',board_name:'FJ+DE2101A1 PCB',format:'',amount:480.0,remarks:'',quote_no:'2025-060',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-PCB-0012',machine_code:'PCB',brand:'新提案　→管理費を固定、利益を12％に設定した見積書',board_name:'FJ+D2101A PCB',format:'',amount:537.6,remarks:'',quote_no:'2025-061',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-PCB-0013',machine_code:'PCB',brand:'新提案　→管理費を固定、利益を12％に設定した見積書',board_name:'FJ+DE1802A PCB',format:'',amount:460.0,remarks:'',quote_no:'2025-126',submit_to:'',submit_date:'2026/02/05',file_url:''},
      {id:'QS-PCB-0014',machine_code:'PCB',brand:'新提案　→管理費を固定、利益を12％に設定した見積書',board_name:'FJ+E2102B PCB',format:'',amount:1097.0,remarks:'',quote_no:'2025-047',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-PCB-0015',machine_code:'PCB',brand:'新提案　→管理費を固定、利益を12％に設定した見積書',board_name:'FJ+D1401C PCB',format:'',amount:633.0,remarks:'',quote_no:'2025-109',submit_to:'',submit_date:'2026/01/16',file_url:''},
      {id:'QS-PCB-0016',machine_code:'PCB',brand:'新提案　→管理費を固定、利益を12％に設定した見積書',board_name:'FJ+E2501B PCB',format:'',amount:974.0,remarks:'提出済み',quote_no:'2026-01-152',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-PCB-0017',machine_code:'PCB',brand:'新提案　→管理費を固定、利益を12％に設定した見積書',board_name:'FJ+C2501C PCB',format:'',amount:854.0,remarks:'提出済み',quote_no:'2026-01-154',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-PCB-0018',machine_code:'PCB',brand:'新提案　→管理費を固定、利益を12％に設定した見積書',board_name:'FJ+M2503A PCB',format:'',amount:388.0,remarks:'提出済み',quote_no:'2026-121',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-PCB-0019',machine_code:'PCB',brand:'新提案　→管理費を固定、利益を12％に設定した見積書',board_name:'FJ+C2502A PCB',format:'',amount:638.0,remarks:'',quote_no:'2025-134',submit_to:'',submit_date:'2026/02/09',file_url:''},
      {id:'QS-PFK30-0020',machine_code:'PFK30',brand:'FJ（）',board_name:'PFK30 枠制御基板（C2501C）見本機検査',format:'',amount:null,remarks:'PCB、ケース別途 R89外し、検査含む',quote_no:'2025-117',submit_to:'井野さん',submit_date:'',file_url:''},
      {id:'QS-PFK30-0021',machine_code:'PFK30',brand:'FJ（）',board_name:'PFK30 枠制御基板仕掛基板',format:'',amount:3150.0,remarks:'',quote_no:'2025-117',submit_to:'井野さん',submit_date:'',file_url:''},
      {id:'QS-PFK30-0022',machine_code:'PFK30',brand:'FJ（）',board_name:'PCB',format:'',amount:854.0,remarks:'',quote_no:'2025-117',submit_to:'井野さん',submit_date:'',file_url:''},
      {id:'QS-PFK30-0023',machine_code:'PFK30',brand:'FJ（）',board_name:'ケース上蓋',format:'',amount:335.0,remarks:'',quote_no:'2025-142',submit_to:'井野さん',submit_date:'2026/03/06',file_url:''},
      {id:'QS-PFK30-0024',machine_code:'PFK30',brand:'FJ（）',board_name:'ケース下蓋',format:'',amount:306.0,remarks:'',quote_no:'2025-142',submit_to:'井野さん',submit_date:'2026/03/06',file_url:''},
      {id:'QS-PFK30-0025',machine_code:'PFK30',brand:'FJ（）',board_name:'PFK30 枠制御基板（完）',format:'',amount:1643.0,remarks:'ケース上下含む',quote_no:'2025-142',submit_to:'井野さん',submit_date:'2026/03/06',file_url:''},
      {id:'QS-PFK30-0026',machine_code:'PFK30',brand:'FJ（）',board_name:'アンプ基板',format:'',amount:872.0,remarks:'',quote_no:'2025-148',submit_to:'井野さん',submit_date:'2026/03/18',file_url:''},
      {id:'QS-PFK30-0027',machine_code:'PFK30',brand:'FJ（）',board_name:'ジャック基板',format:'',amount:244.0,remarks:'',quote_no:'2025-116',submit_to:'井野さん',submit_date:'2026/02/03',file_url:''},
      {id:'QS-C46B-0028',machine_code:'C46B',brand:'FJ（）',board_name:'主制御基板（完）',format:'新提案',amount:4230.0,remarks:'',quote_no:'2025-099',submit_to:'金沢さん',submit_date:'2025/12/15',file_url:''},
      {id:'QS-C46B-0029',machine_code:'C46B',brand:'FJ（）',board_name:'液晶演出制御（完）ALL新品',format:'新提案',amount:22552.0,remarks:'DE2101A1  仕掛、PCB含まず',quote_no:'2025-098',submit_to:'金沢さん',submit_date:'2025/12/15',file_url:''},
      {id:'QS-C46-0030',machine_code:'C46',brand:'FJ（）',board_name:'FJ+M2401A PCB',format:'新提案',amount:655.0,remarks:'PCBのみ',quote_no:'製見-2024-0135',submit_to:'大山さん',submit_date:'2024/11/20',file_url:''},
      {id:'QS-C46-0031',machine_code:'C46',brand:'FJ（）',board_name:'FJ+M2401A',format:'新提案',amount:2457.0,remarks:'PCB除く',quote_no:'24-MP-038',submit_to:'大山さん',submit_date:'2025/04/02',file_url:''},
      {id:'QS-C46-0032',machine_code:'C46',brand:'FJ（）',board_name:'FJ+M2401A （完）',format:'新提案',amount:562.0,remarks:'',quote_no:'2025-032',submit_to:'大山さん',submit_date:'2025/06/27',file_url:''},
      {id:'QS-C46-0033',machine_code:'C46',brand:'FJ（）',board_name:'液晶演出制御（完）ALL新品',format:'新提案',amount:25437.0,remarks:'DE2101A1 PCB含まず',quote_no:'2025-031',submit_to:'',submit_date:'2025/06/27',file_url:''},
      {id:'QS-C46-0034',machine_code:'C46',brand:'FJ（）',board_name:'液晶演出制御（完）ALL新品',format:'新提案',amount:14281.0,remarks:'E、DE仕掛、PCB別途',quote_no:'2025-062',submit_to:'',submit_date:'2025/09/11',file_url:''},
      {id:'QS-C46-0035',machine_code:'C46',brand:'FJ（）',board_name:'液晶演出制御（完）アフター',format:'新提案',amount:14274.0,remarks:'E、DE仕掛、PCB別途',quote_no:'2025-067-1',submit_to:'',submit_date:'2025/10/07',file_url:''},
      {id:'QS-C46-0036',machine_code:'C46',brand:'FJ（）',board_name:'液晶演出制御（完）',format:'',amount:1288.0,remarks:'Dのみリユース',quote_no:'2025-070',submit_to:'大山さん',submit_date:'2025/10/15',file_url:''},
      {id:'QS-C46-0037',machine_code:'C46',brand:'FJ（）',board_name:'液晶演出制御（完）',format:'',amount:1488.0,remarks:'ALLリユース',quote_no:'2025-072',submit_to:'大山さん',submit_date:'2025/10/15',file_url:''},
      {id:'QS-C46-0038',machine_code:'C46',brand:'FJ（）',board_name:'液晶演出制御（完）',format:'',amount:1388.0,remarks:'D、Eリユース',quote_no:'2025-071',submit_to:'大山さん',submit_date:'2025/10/15',file_url:''},
      {id:'QS-A87（AG6R）-0039',machine_code:'A87（AG6R）',brand:'FJ（）',board_name:'FJ+M2503A PCB',format:'新提案',amount:655.0,remarks:'PCBのみ',quote_no:'',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-A87（AG6R）-0040',machine_code:'A87（AG6R）',brand:'FJ（）',board_name:'FJ+M2503A （仕掛）',format:'新提案',amount:2457.0,remarks:'PCB除く（仕掛）',quote_no:'',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-A87（AG6R）-0041',machine_code:'A87（AG6R）',brand:'FJ（）',board_name:'主基板（完）',format:'',amount:708.0,remarks:'',quote_no:'',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-A87（AG6R）-0042',machine_code:'A87（AG6R）',brand:'FJ（）',board_name:'D60 液晶演出制御基板（完）',format:'新提案',amount:1679.0,remarks:'VDP、E基板PCB除く',quote_no:'製見-2024-0174',submit_to:'大山さん',submit_date:'2025/03/30',file_url:''},
      {id:'QS-A87（AG6R）-0043',machine_code:'A87（AG6R）',brand:'FJ（）',board_name:'SNB52163A',format:'新提案',amount:7837.0,remarks:'',quote_no:'製見-2024-0170',submit_to:'大山さん',submit_date:'2025/03/30',file_url:''},
      {id:'QS-A87（AG6R）-0044',machine_code:'A87（AG6R）',brand:'FJ（）',board_name:'FJ+E2501B（仕掛）',format:'新提案',amount:4736.0,remarks:'',quote_no:'製見-2024-0170',submit_to:'大山さん',submit_date:'2025/03/30',file_url:''},
      {id:'QS-A87（AG6R）-0045',machine_code:'A87（AG6R）',brand:'FJ（）',board_name:'FJ+E2501B PCB',format:'新提案',amount:956.0,remarks:'PCBのみ',quote_no:'',submit_to:'大山さん',submit_date:'2025/03/30',file_url:''},
      {id:'QS-A87（AG6R）-0046',machine_code:'A87（AG6R）',brand:'FJ（）',board_name:'AX51611（VDP）',format:'新提案',amount:16659.0,remarks:'VDP単体',quote_no:'',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-D61-0047',machine_code:'D61',brand:'FJ（）',board_name:'液晶演出制御（完）',format:'新提案',amount:1188.0,remarks:'ALL仕掛基板、PCB別途',quote_no:'No.2025-108',submit_to:'金沢さん',submit_date:'2026/02/03',file_url:''},
      {id:'QS-A86-0048',machine_code:'A86',brand:'FJ（）',board_name:'液晶演出制御（完）ALL新品',format:'新提案',amount:14281.0,remarks:'DE/E仕掛、PCB別',quote_no:'No.2025-076',submit_to:'井野さん',submit_date:'',file_url:''},
      {id:'QS-A86-0049',machine_code:'A86',brand:'FJ（）',board_name:'液晶演出制御（完）見本機',format:'新提案',amount:1276.0,remarks:'見本機（Dのみリユース） 見本機シール無',quote_no:'No.2025-079',submit_to:'井野さん',submit_date:'2025/11/14',file_url:''},
      {id:'QS-A86-0050',machine_code:'A86',brand:'FJ（）',board_name:'液晶演出制御（完）',format:'新提案',amount:1188.0,remarks:'ALL仕掛基板、PCB別途',quote_no:'No.2025-108',submit_to:'井野さん',submit_date:'2026/01/16',file_url:''},
      {id:'QS-A85-0051',machine_code:'A85',brand:'FJ（）',board_name:'FJ+M2401A PCB',format:'新提案',amount:655.0,remarks:'PCBのみ',quote_no:'製見-2024-0135',submit_to:'大山さん',submit_date:'2024/11/20',file_url:''},
      {id:'QS-A85-0052',machine_code:'A85',brand:'FJ（）',board_name:'FJ+M2401A',format:'新提案',amount:2457.0,remarks:'PCB除く',quote_no:'24-MP-038',submit_to:'大山さん',submit_date:'2025/04/02',file_url:''},
      {id:'QS-A85-0053',machine_code:'A85',brand:'FJ（）',board_name:'FJ+M2401A （完）',format:'新提案',amount:562.0,remarks:'',quote_no:'',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-A85-0054',machine_code:'A85',brand:'FJ（）',board_name:'液晶演出制御（完）',format:'新提案',amount:1188.0,remarks:'ALL仕掛基板、PCB別途',quote_no:'2025-107',submit_to:'井野さん',submit_date:'2026/01/16',file_url:''},
      {id:'QS-A85-0055',machine_code:'A85',brand:'FJ（）',board_name:'液晶演出制御（完）アフター',format:'新提案',amount:1181.0,remarks:'ALL仕掛基板、PCB別途',quote_no:'2025-132',submit_to:'井野さん',submit_date:'2026/02/13',file_url:''},
      {id:'QS-A85-0056',machine_code:'A85',brand:'FJ（）',board_name:'液晶演出制御（完）ALL中古',format:'新提案',amount:1488.0,remarks:'ALL仕掛基板、PCB別途',quote_no:'2025-131',submit_to:'井野さん',submit_date:'2026/02/13',file_url:''},
      {id:'QS-D59B-0057',machine_code:'D59B',brand:'FJ（）',board_name:'FJ+M2401A PCB',format:'新提案',amount:655.0,remarks:'PCBのみ',quote_no:'製見-2024-0135',submit_to:'大山さん',submit_date:'2024/11/20',file_url:''},
      {id:'QS-D59B-0058',machine_code:'D59B',brand:'FJ（）',board_name:'FJ+M2401A',format:'新提案',amount:2457.0,remarks:'PCB除く',quote_no:'24-MP-038',submit_to:'大山さん',submit_date:'2025/04/02',file_url:''},
      {id:'QS-D59B-0059',machine_code:'D59B',brand:'FJ（）',board_name:'D59B 主制御基板（完）M2401A',format:'新提案',amount:562.0,remarks:'',quote_no:'2025-001',submit_to:'大山さん',submit_date:'2025/04/02',file_url:''},
      {id:'QS-D59B-0060',machine_code:'D59B',brand:'FJ（）',board_name:'D59B 主制御基板（完）M2401A',format:'新提案',amount:612.0,remarks:'中古基板使用',quote_no:'2025-151',submit_to:'大山さん',submit_date:'2026/03/18',file_url:''},
      {id:'QS-D59B-0061',machine_code:'D59B',brand:'FJ（）',board_name:'見本機',format:'新提案',amount:1288.0,remarks:'Dのみ中古',quote_no:'NO.2025-122',submit_to:'大山さん',submit_date:'2026/02/05',file_url:''},
      {id:'QS-D59B-0062',machine_code:'D59B',brand:'FJ（）',board_name:'アフター用',format:'新提案',amount:1181.0,remarks:'',quote_no:'NO.2025-119',submit_to:'大山さん',submit_date:'2026/02/05',file_url:''},
      {id:'QS-D59B-0063',machine_code:'D59B',brand:'FJ（）',board_name:'ALL中古基板',format:'新提案',amount:1488.0,remarks:'',quote_no:'NO.2025-123',submit_to:'大山さん',submit_date:'2026/02/05',file_url:''},
      {id:'QS-NA10-0064',machine_code:'NA10',brand:'FJ（）ZEEG筐体　主メダル数制御一体型',board_name:'NA10主メダル数制御見本機',format:'',amount:1404.0,remarks:'見本機',quote_no:'No.2025-068',submit_to:'大山さん',submit_date:'2025/10/07',file_url:''},
      {id:'QS-NA10-0065',machine_code:'NA10',brand:'FJ（）ZEEG筐体　主メダル数制御一体型',board_name:'NA10主メダル数制御見本機',format:'',amount:1670.0,remarks:'量産Lot5,000台以上',quote_no:'No.2025-075',submit_to:'大山さん',submit_date:'2025/10/15',file_url:''},
      {id:'QS-NA10-0066',machine_code:'NA10',brand:'FJ（）ZEEG筐体　主メダル数制御一体型',board_name:'治具費一式',format:'',amount:956000.0,remarks:'',quote_no:'No.2025-068',submit_to:'大山さん',submit_date:'2025/10/07',file_url:''},
      {id:'QS-NA10-0067',machine_code:'NA10',brand:'FJ（）ZEEG筐体　主メダル数制御一体型',board_name:'1. ROM圧入治具',format:'',amount:390000.0,remarks:'',quote_no:'No.2025-068',submit_to:'大山さん',submit_date:'2025/10/07',file_url:''},
      {id:'QS-NA10-0068',machine_code:'NA10',brand:'FJ（）ZEEG筐体　主メダル数制御一体型',board_name:'2. QRコードシール貼付け受け治具',format:'',amount:88000.0,remarks:'',quote_no:'No.2025-068',submit_to:'大山さん',submit_date:'2025/10/08',file_url:''},
      {id:'QS-NA10-0069',machine_code:'NA10',brand:'FJ（）ZEEG筐体　主メダル数制御一体型',board_name:'3. 主制御基板ネジ止め受け治具',format:'',amount:70000.0,remarks:'',quote_no:'No.2025-068',submit_to:'大山さん',submit_date:'2025/10/09',file_url:''},
      {id:'QS-NA10-0070',machine_code:'NA10',brand:'FJ（）ZEEG筐体　主メダル数制御一体型',board_name:'4. 下蓋ネジ止め受け治具',format:'',amount:70000.0,remarks:'',quote_no:'No.2025-068',submit_to:'大山さん',submit_date:'2025/10/10',file_url:''},
      {id:'QS-NA10-0071',machine_code:'NA10',brand:'FJ（）ZEEG筐体　主メダル数制御一体型',board_name:'5. かしめブロック圧入治具',format:'',amount:80000.0,remarks:'',quote_no:'No.2025-068',submit_to:'大山さん',submit_date:'2025/10/11',file_url:''},
      {id:'QS-NA10-0072',machine_code:'NA10',brand:'FJ（）ZEEG筐体　主メダル数制御一体型',board_name:'6. 勘合確認治具',format:'',amount:90000.0,remarks:'',quote_no:'No.2025-068',submit_to:'大山さん',submit_date:'2025/10/12',file_url:''},
      {id:'QS-NA10-0073',machine_code:'NA10',brand:'FJ（）ZEEG筐体　主メダル数制御一体型',board_name:'7. 封印シール貼付け治具',format:'',amount:78000.0,remarks:'',quote_no:'No.2025-068',submit_to:'大山さん',submit_date:'2025/10/13',file_url:''},
      {id:'QS-NA10-0074',machine_code:'NA10',brand:'FJ（）ZEEG筐体　主メダル数制御一体型',board_name:'8. 下蓋勘合治具',format:'',amount:90000.0,remarks:'',quote_no:'No.2025-068',submit_to:'大山さん',submit_date:'2025/10/14',file_url:''},
      {id:'QS-E49-0075',machine_code:'E49',brand:'FJ（）',board_name:'FJ+M2402A PCB',format:'新提案',amount:659.0,remarks:'',quote_no:'製見2024-0154',submit_to:'大山さん',submit_date:'2025/01/10',file_url:''},
      {id:'QS-E49-0076',machine_code:'E49',brand:'FJ（）',board_name:'FJ+M2402A仕掛基板',format:'新提案',amount:3452.0,remarks:'',quote_no:'No.2025-013',submit_to:'鎌田さん',submit_date:'',file_url:''},
      {id:'QS-E49-0077',machine_code:'E49',brand:'FJ（）',board_name:'FJ+M2402A（完）',format:'新提案',amount:708.0,remarks:'',quote_no:'No.2025-020',submit_to:'大山さん',submit_date:'2025/05/28',file_url:''},
      {id:'QS-E49-0078',machine_code:'E49',brand:'FJ（）',board_name:'オール新品',format:'新提案',amount:26129.0,remarks:'2024年価格改定',quote_no:'製見2024-0153',submit_to:'大山さん',submit_date:'2025/01/10',file_url:''},
      {id:'QS-E49-0079',machine_code:'E49',brand:'FJ（）',board_name:'D/DEリユース',format:'新提案',amount:1444.0,remarks:'',quote_no:'2025-049',submit_to:'大山さん',submit_date:'2025/09/19',file_url:''},
      {id:'QS-E49-0080',machine_code:'E49',brand:'FJ（）',board_name:'ALL中古',format:'新提案',amount:1544.0,remarks:'',quote_no:'2025-052',submit_to:'大山さん',submit_date:'2025/09/19',file_url:''},
      {id:'QS-A83B-0081',machine_code:'A83B',brand:'FJ（）',board_name:'FJ+M2003A',format:'新提案',amount:4230.0,remarks:'',quote_no:'2025-054',submit_to:'井野さん',submit_date:'2025/08/22',file_url:''},
      {id:'QS-A83B-0082',machine_code:'A83B',brand:'FJ（）',board_name:'FJ+M2003A',format:'新提案',amount:612.0,remarks:'中古基板使用',quote_no:'2025-114',submit_to:'井野さん',submit_date:'2026/01/22',file_url:''},
      {id:'QS-A83B-0083',machine_code:'A83B',brand:'FJ（）',board_name:'液晶演出制御（完）ALL新品',format:'新提案',amount:24851.0,remarks:'DE2101A1 PCB含まず',quote_no:'2025-055',submit_to:'井野さん',submit_date:'2025/08/22',file_url:''},
      {id:'QS-A83B-0084',machine_code:'A83B',brand:'FJ（）',board_name:'液晶演出制御（完）ALL中古基板使用',format:'新提案',amount:1488.0,remarks:'中古基板使用',quote_no:'2025-097',submit_to:'井野さん',submit_date:'2025/12/12',file_url:''},
      {id:'QS-A83B-0085',machine_code:'A83B',brand:'FJ（）',board_name:'液晶演出制御（完）アフター',format:'新提案',amount:9452.0,remarks:'D,DE仕掛基板別、E基板含む',quote_no:'2025-113',submit_to:'井野さん',submit_date:'2026/01/22',file_url:''},
      {id:'QS-A83B-0086',machine_code:'A83B',brand:'FJ（）',board_name:'液晶演出制御（完）D,DE中古',format:'新提案',amount:9658.0,remarks:'D,DE中古基板、E基板含む',quote_no:'2025-112',submit_to:'井野さん',submit_date:'2026/01/21',file_url:''},
      {id:'QS-D60（AG6R）-0087',machine_code:'D60（AG6R）',brand:'FJ（）',board_name:'FJ+M2401A PCB',format:'新提案',amount:655.0,remarks:'PCBのみ',quote_no:'No.24-MP036',submit_to:'大山さん',submit_date:'2025/03/30',file_url:''},
      {id:'QS-D60（AG6R）-0088',machine_code:'D60（AG6R）',brand:'FJ（）',board_name:'FJ+M2401A （仕掛）',format:'新提案',amount:2457.0,remarks:'PCB除く（仕掛）',quote_no:'No.24-MP038',submit_to:'大山さん',submit_date:'2025/03/30',file_url:''},
      {id:'QS-D60（AG6R）-0089',machine_code:'D60（AG6R）',brand:'FJ（）',board_name:'D60主制御（完）',format:'新提案',amount:562.0,remarks:'',quote_no:'製見-2024-0173',submit_to:'大山さん',submit_date:'2025/03/30',file_url:''},
      {id:'QS-D60（AG6R）-0090',machine_code:'D60（AG6R）',brand:'FJ（）',board_name:'D60 液晶演出制御基板（完）',format:'新提案',amount:1679.0,remarks:'VDP、E基板PCB除く',quote_no:'製見-2024-0174',submit_to:'大山さん',submit_date:'2025/03/30',file_url:''},
      {id:'QS-D60（AG6R）-0091',machine_code:'D60（AG6R）',brand:'FJ（）',board_name:'SNB52163A',format:'新提案',amount:7837.0,remarks:'',quote_no:'製見-2024-0170',submit_to:'大山さん',submit_date:'2025/03/30',file_url:''},
      {id:'QS-D60（AG6R）-0092',machine_code:'D60（AG6R）',brand:'FJ（）',board_name:'FJ+E2301B（仕掛）',format:'新提案',amount:4736.0,remarks:'',quote_no:'製見-2024-0170',submit_to:'大山さん',submit_date:'2025/03/30',file_url:''},
      {id:'QS-D60（AG6R）-0093',machine_code:'D60（AG6R）',brand:'FJ（）',board_name:'FJ+E2301B PCB',format:'新提案',amount:956.0,remarks:'PCBのみ',quote_no:'',submit_to:'大山さん',submit_date:'2025/03/30',file_url:''},
      {id:'QS-D60（AG6R）-0094',machine_code:'D60（AG6R）',brand:'FJ（）',board_name:'AX51611（VDP）',format:'新提案',amount:16659.0,remarks:'VDP単体',quote_no:'',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-D60（AG6R）-0095',machine_code:'D60（AG6R）',brand:'FJ（）',board_name:'D60 液晶演出制御基板（完）アフター',format:'新提案',amount:1676.0,remarks:'VDP、E基板PCB除く',quote_no:'2025-017',submit_to:'大山さん',submit_date:'2025/05/21',file_url:''},
      {id:'QS-D60（AG6R）-0096',machine_code:'D60（AG6R）',brand:'FJ（）',board_name:'D60 液晶演出制御（完）',format:'',amount:1791.0,remarks:'Lot2,000台～5,000台',quote_no:'2025-093',submit_to:'大山さん',submit_date:'2025/12/12',file_url:''},
      {id:'QS-D60（AG6R）-0097',machine_code:'D60（AG6R）',brand:'FJ（）',board_name:'D60 液晶演出制御（完）中古基板使用',format:'',amount:1567.0,remarks:'中古基板使用',quote_no:'2025-094',submit_to:'大山さん',submit_date:'2025/12/12',file_url:''},
      {id:'QS-D60（AG6R）-0098',machine_code:'D60（AG6R）',brand:'FJ（）',board_name:'D60 液晶演出制御基板（完）見本機',format:'新提案',amount:2275.0,remarks:'見本機',quote_no:'2025-022',submit_to:'金沢さん',submit_date:'2025/05/29',file_url:''},
      {id:'QS-PFK25-0099',machine_code:'PFK25',brand:'FJ（）',board_name:'FJ+C2401A 仕掛基板（レーザー）',format:'新提案',amount:3621.0,remarks:'PCB含まず',quote_no:'No.2025-004',submit_to:'井野さん',submit_date:'2025/04/08',file_url:''},
      {id:'QS-PFK25-0100',machine_code:'PFK25',brand:'FJ（）',board_name:'FJ+C2401A 仕掛基板（CN支給）',format:'新提案',amount:3584.0,remarks:'PCB含まず',quote_no:'No.2025-004',submit_to:'井野さん',submit_date:'2025/04/08',file_url:''},
      {id:'QS-PFK25-0101',machine_code:'PFK25',brand:'FJ（）',board_name:'FJ+C2401A 仕掛基板（CN自達）',format:'新提案',amount:3621.0,remarks:'PCB含まず',quote_no:'No.2025-004',submit_to:'井野さん',submit_date:'2025/04/08',file_url:''},
      {id:'QS-PFK25-0102',machine_code:'PFK25',brand:'FJ（）',board_name:'FJ+C2401A PCB',format:'新提案',amount:814.0,remarks:'',quote_no:'No.24-MP036',submit_to:'井野さん',submit_date:'2025/03/19',file_url:''},
      {id:'QS-PFK25-0103',machine_code:'PFK25',brand:'FJ（）',board_name:'PFK25枠制御基板（完）',format:'新提案',amount:1307.0,remarks:'',quote_no:'製見-2024-0110',submit_to:'井野さん',submit_date:'2025/03/19',file_url:''},
      {id:'QS-PFK25-0104',machine_code:'PFK25',brand:'FJ（）',board_name:'PFK20→PFK25リユース',format:'',amount:1602.0,remarks:'',quote_no:'2025-092-1',submit_to:'山崎さん',submit_date:'2025/12/11',file_url:''},
      {id:'QS-PFK25-0105',machine_code:'PFK25',brand:'FJ（）',board_name:'PFK20　ICT検査',format:'',amount:150.0,remarks:'',quote_no:'2025-091-1',submit_to:'山崎さん',submit_date:'2025/12/11',file_url:''},
      {id:'QS-E61-0106',machine_code:'E61',brand:'FJ（）',board_name:'FJ+M2402A PCB',format:'新提案',amount:659.0,remarks:'',quote_no:'製見2024-0154',submit_to:'大山さん',submit_date:'2025/01/10',file_url:''},
      {id:'QS-E61-0107',machine_code:'E61',brand:'FJ（）',board_name:'FJ+M2402A仕掛基板',format:'新提案',amount:3452.0,remarks:'',quote_no:'No.2025-013',submit_to:'鎌田さん',submit_date:'',file_url:''},
      {id:'QS-E61-0108',machine_code:'E61',brand:'FJ（）',board_name:'FJ+M2402A（完）',format:'新提案',amount:708.0,remarks:'',quote_no:'No.2025-012',submit_to:'鎌田さん',submit_date:'',file_url:''},
      {id:'QS-E61-0109',machine_code:'E61',brand:'FJ（）',board_name:'オール新品',format:'新提案',amount:26129.0,remarks:'2024年価格改定',quote_no:'製見2024-0153',submit_to:'大山さん',submit_date:'2025/01/10',file_url:''},
      {id:'QS-E61-0110',machine_code:'E61',brand:'FJ（）',board_name:'D/DEリユース',format:'新提案',amount:1444.0,remarks:'',quote_no:'No.2025-088',submit_to:'大山さん',submit_date:'2025/11/27',file_url:''},
      {id:'QS-E61-0111',machine_code:'E61',brand:'FJ（）',board_name:'見本機',format:'新提案',amount:1544.0,remarks:'ALL中古基板使用',quote_no:'No.2025-030',submit_to:'',submit_date:'2025/06/19',file_url:''},
      {id:'QS-E61-0112',machine_code:'E61',brand:'FJ（）',board_name:'見本機',format:'新提案',amount:10630.0,remarks:'E2102Bのみ新品',quote_no:'No.2025-029',submit_to:'',submit_date:'2025/06/19',file_url:''},
      {id:'QS-E60-0113',machine_code:'E60',brand:'FJ（）',board_name:'FJ+M2402A PCB',format:'新提案',amount:659.0,remarks:'',quote_no:'製見2024-0154',submit_to:'大山さん',submit_date:'2025/01/10',file_url:''},
      {id:'QS-E60-0114',machine_code:'E60',brand:'FJ（）',board_name:'FJ+M2402A仕掛基板',format:'新提案',amount:3452.0,remarks:'',quote_no:'No.2025-013',submit_to:'鎌田さん',submit_date:'',file_url:''},
      {id:'QS-E60-0115',machine_code:'E60',brand:'FJ（）',board_name:'見本機',format:'新提案',amount:1544.0,remarks:'ALL中古基板使用',quote_no:'No.2025-028',submit_to:'',submit_date:'2025/06/19',file_url:''},
      {id:'QS-E60-0116',machine_code:'E60',brand:'FJ（）',board_name:'見本機',format:'新提案',amount:10630.0,remarks:'E2102Bのみ新品',quote_no:'No.2025-027',submit_to:'',submit_date:'2025/06/19',file_url:''},
      {id:'QS-D58B-0117',machine_code:'D58B',brand:'JFJ（確定）',board_name:'D58B 主制御基板（完）',format:'新提案',amount:562.0,remarks:'',quote_no:'製見-2024-0172',submit_to:'大山さん',submit_date:'2025/03/30',file_url:''},
      {id:'QS-D58B-0118',machine_code:'D58B',brand:'JFJ（確定）',board_name:'FJ+M2401A（仕掛）',format:'新提案',amount:2457.0,remarks:'',quote_no:'No.24-MP038',submit_to:'大山さん',submit_date:'2025/03/30',file_url:''},
      {id:'QS-D58B-0119',machine_code:'D58B',brand:'JFJ（確定）',board_name:'FJ+M2401A PCB',format:'新提案',amount:655.0,remarks:'',quote_no:'No.24-MP036',submit_to:'大山さん',submit_date:'2025/03/30',file_url:''},
      {id:'QS-D58B-0120',machine_code:'D58B',brand:'JFJ（確定）',board_name:'D58B 液晶演出制御基板（完）',format:'新提案',amount:26062.0,remarks:'',quote_no:'製見-2024-0171',submit_to:'大山さん',submit_date:'2025/03/30',file_url:''},
      {id:'QS-D58B-0121',machine_code:'D58B',brand:'JFJ（確定）',board_name:'D58B 液晶演出制御基板（完）',format:'新提案',amount:10244.0,remarks:'エコ（E基板のみ新品）',quote_no:'No.2025-035',submit_to:'大山さん',submit_date:'2025/07/17',file_url:''},
      {id:'QS-D58B-0122',machine_code:'D58B',brand:'JFJ（確定）',board_name:'アフター',format:'新提案',amount:26056.0,remarks:'',quote_no:'No.2025-044',submit_to:'大山さん',submit_date:'2025/08/08',file_url:''},
      {id:'QS-D58B-0123',machine_code:'D58B',brand:'JFJ（確定）',board_name:'ALLリユース',format:'新提案',amount:1488.0,remarks:'ALLリユース',quote_no:'No.2025-042',submit_to:'大山さん',submit_date:'2025/08/08',file_url:''},
      {id:'QS-D58B-0124',machine_code:'D58B',brand:'JFJ（確定）',board_name:'D58B 液晶演出制御基板（完）',format:'新提案',amount:13069.0,remarks:'Dのみリユース',quote_no:'No.2025-043',submit_to:'大山さん',submit_date:'2025/08/18',file_url:''},
      {id:'QS-D56C-0125',machine_code:'D56C',brand:'FJ（）',board_name:'FJ+M2003A5 PCB',format:'新提案',amount:679.0,remarks:'PCBのみ',quote_no:'製見-2024-0033',submit_to:'大山さん',submit_date:'',file_url:''},
      {id:'QS-D56C-0126',machine_code:'D56C',brand:'FJ（）',board_name:'FJ+M2003A5',format:'新提案',amount:3764.0,remarks:'PCB除く、B36B支給条件',quote_no:'製見-2024-0132',submit_to:'大山さん',submit_date:'2024/11/20',file_url:''},
      {id:'QS-D56C-0127',machine_code:'D56C',brand:'FJ（）',board_name:'FJ+M2003A5（中古基板使用）',format:'新提案',amount:612.0,remarks:'中古基板使用',quote_no:'No.2025-024',submit_to:'大山さん',submit_date:'',file_url:''},
      {id:'QS-D56C-0128',machine_code:'D56C',brand:'FJ（）',board_name:'オール新品',format:'新提案',amount:25477.0,remarks:'24年5月価格改定適用',quote_no:'製見-2024-0133',submit_to:'大山さん',submit_date:'2024/11/20',file_url:''},
      {id:'QS-D56C-0129',machine_code:'D56C',brand:'FJ（）',board_name:'アフター用',format:'新提案',amount:25470.0,remarks:'24年5月価格改定適用',quote_no:'製見-2024-0159',submit_to:'大山さん',submit_date:'2025/01/30',file_url:''},
      {id:'QS-D56C-0130',machine_code:'D56C',brand:'FJ（）',board_name:'オール中古',format:'新提案',amount:1488.0,remarks:'ALL中古基板使用',quote_no:'No.2025-023',submit_to:'大山さん',submit_date:'',file_url:''},
      {id:'QS-A84（AG6R）-0131',machine_code:'A84（AG6R）',brand:'FJ（）',board_name:'FJ+M2401A PCB',format:'新提案',amount:655.0,remarks:'PCBのみ',quote_no:'製見-2024-0135',submit_to:'大山さん',submit_date:'2024/11/20',file_url:''},
      {id:'QS-A84（AG6R）-0132',machine_code:'A84（AG6R）',brand:'FJ（）',board_name:'FJ+M2401A',format:'新提案',amount:3019.0,remarks:'PCB除く',quote_no:'製見-2024-0137',submit_to:'井野さん',submit_date:'2024/11/20',file_url:''},
      {id:'QS-A84（AG6R）-0133',machine_code:'A84（AG6R）',brand:'FJ（）',board_name:'オール新品（SNB52161B）e機 見本機',format:'新提案',amount:14849.0,remarks:'VDP、E基板PCB除く',quote_no:'製見-2024-0165',submit_to:'井野さん',submit_date:'2025/03/05',file_url:''},
      {id:'QS-A84（AG6R）-0134',machine_code:'A84（AG6R）',brand:'FJ（）',board_name:'オール新品（SNB52161B）e機',format:'新提案',amount:14253.0,remarks:'VDP、E基板PCB除く',quote_no:'製見-2024-0169',submit_to:'井野さん',submit_date:'2025/03/14',file_url:''},
      {id:'QS-A84（AG6R）-0135',machine_code:'A84（AG6R）',brand:'FJ（）',board_name:'A84 液晶演出制御基板（完）',format:'新提案',amount:1679.0,remarks:'VDP、E基板PCB除く',quote_no:'製見-2024-0174',submit_to:'井野さん',submit_date:'2025/03/30',file_url:''},
      {id:'QS-A84（AG6R）-0136',machine_code:'A84（AG6R）',brand:'FJ（）',board_name:'SNB52163A',format:'新提案',amount:7837.0,remarks:'',quote_no:'製見-2024-0170',submit_to:'井野さん',submit_date:'2025/03/30',file_url:''},
      {id:'QS-A84（AG6R）-0137',machine_code:'A84（AG6R）',brand:'FJ（）',board_name:'FJ+E2301B（仕掛）',format:'新提案',amount:4736.0,remarks:'',quote_no:'製見-2024-0170',submit_to:'井野さん',submit_date:'2025/03/30',file_url:''},
      {id:'QS-A84（AG6R）-0138',machine_code:'A84（AG6R）',brand:'FJ（）',board_name:'FJ+E2301B PCB',format:'新提案',amount:956.0,remarks:'PCBのみ',quote_no:'',submit_to:'井野さん',submit_date:'2025/03/30',file_url:''},
      {id:'QS-A84（AG6R）-0139',machine_code:'A84（AG6R）',brand:'FJ（）',board_name:'AX51611（VDP）',format:'新提案',amount:16659.0,remarks:'VDP単体',quote_no:'',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-D59-0140',machine_code:'D59',brand:'FJ（）',board_name:'FJ+M2401A PCB',format:'新提案',amount:655.0,remarks:'PCBのみ',quote_no:'製見-2024-0135',submit_to:'大山さん',submit_date:'2024/11/20',file_url:''},
      {id:'QS-D59-0141',machine_code:'D59',brand:'FJ（）',board_name:'FJ+M2401A',format:'新提案',amount:2457.0,remarks:'PCB除く',quote_no:'24-MP-038',submit_to:'大山さん',submit_date:'2025/04/02',file_url:''},
      {id:'QS-D59-0142',machine_code:'D59',brand:'FJ（）',board_name:'D59 主制御基板（完）M2401A',format:'新提案',amount:562.0,remarks:'',quote_no:'2025-001',submit_to:'大山さん',submit_date:'2025/04/02',file_url:''},
      {id:'QS-D59-0143',machine_code:'D59',brand:'FJ（）',board_name:'オール新品',format:'新提案',amount:25437.0,remarks:'DE2101A1 PCB除く',quote_no:'製見-2024-0134',submit_to:'大山さん',submit_date:'2025/02/03',file_url:''},
      {id:'QS-D59-0144',machine_code:'D59',brand:'FJ（）',board_name:'アフター用',format:'新提案',amount:25430.0,remarks:'24年5月価格改定適用',quote_no:'製見-2024-0160',submit_to:'大山さん',submit_date:'2025/02/03',file_url:''},
      {id:'QS-D59-0145',machine_code:'D59',brand:'FJ（）',board_name:'Ｄのみリユース',format:'新提案',amount:12443.0,remarks:'24年5月価格改定適用',quote_no:'No.2025-007',submit_to:'大山さん',submit_date:'2025/05/01',file_url:''},
      {id:'QS-D59-0146',machine_code:'D59',brand:'FJ（）',board_name:'ALL中古基板',format:'新提案',amount:1488.0,remarks:'24年5月価格改定適用',quote_no:'No.2025-009',submit_to:'大山さん',submit_date:'2025/05/01',file_url:''},
      {id:'QS-D59-0147',machine_code:'D59',brand:'FJ（）',board_name:'VDP他社USED使用',format:'新提案',amount:23723.0,remarks:'24年5月価格改定適用',quote_no:'No.2025-008',submit_to:'大山さん',submit_date:'2025/05/01',file_url:''},
      {id:'QS-D56B-0148',machine_code:'D56B',brand:'FJ（）',board_name:'FJ+M2003A5 PCB',format:'新提案',amount:679.0,remarks:'PCBのみ',quote_no:'製見-2024-0033',submit_to:'大山さん',submit_date:'',file_url:''},
      {id:'QS-D56B-0149',machine_code:'D56B',brand:'FJ（）',board_name:'FJ+M2003A5',format:'新提案',amount:3807.0,remarks:'PCB除く',quote_no:'製見-2024-0151',submit_to:'大山さん',submit_date:'2024/12/25',file_url:''},
      {id:'QS-D56B-0150',machine_code:'D56B',brand:'FJ（）',board_name:'FJ+M2003A5　中古',format:'新提案',amount:612.0,remarks:'中古基板使用',quote_no:'製見-2024-0147',submit_to:'大山さん',submit_date:'2024/12/25',file_url:''},
      {id:'QS-D56B-0151',machine_code:'D56B',brand:'FJ（）',board_name:'オール新品',format:'新提案',amount:25477.0,remarks:'24年5月価格改定適用',quote_no:'製見-2024-0090',submit_to:'大山さん',submit_date:'2024/10/03',file_url:''},
      {id:'QS-D56B-0152',machine_code:'D56B',brand:'FJ（）',board_name:'オール中古',format:'新提案',amount:1488.0,remarks:'ALL中古基板使用',quote_no:'製見-2024-0148',submit_to:'大山さん',submit_date:'2024/12/25',file_url:''},
      {id:'QS-D56B-0153',machine_code:'D56B',brand:'FJ（）',board_name:'アフター用',format:'新提案',amount:25470.0,remarks:'24年5月価格改定適用',quote_no:'製見-2024-0151',submit_to:'大山さん',submit_date:'2025/01/06',file_url:''},
      {id:'QS-A81 W-0154',machine_code:'A81 W',brand:'FJ（確定）',board_name:'FJ+M2003A5',format:'新提案',amount:4046.0,remarks:'PCB除く',quote_no:'製見-2024-075',submit_to:'井野さん',submit_date:'2024/05/30',file_url:''},
      {id:'QS-A81 W-0155',machine_code:'A81 W',brand:'FJ（確定）',board_name:'FJ+M2003A5（中古基板使用）',format:'新提案',amount:612.0,remarks:'中古基板使用',quote_no:'製見-2024-0149',submit_to:'井野さん',submit_date:'2024/12/25',file_url:''},
      {id:'QS-A81 W-0156',machine_code:'A81 W',brand:'FJ（確定）',board_name:'オール新品',format:'新提案',amount:24979.0,remarks:'',quote_no:'製見-2024-076',submit_to:'井野さん',submit_date:'2024/05/30',file_url:''},
      {id:'QS-A81 W-0157',machine_code:'A81 W',brand:'FJ（確定）',board_name:'オール新品',format:'新提案',amount:25477.0,remarks:'',quote_no:'製見-2024-0130',submit_to:'井野さん',submit_date:'2024/10/30',file_url:''},
      {id:'QS-A81 W-0158',machine_code:'A81 W',brand:'FJ（確定）',board_name:'オール新品（アフター用）',format:'新提案（24年5月価格改定）',amount:25470.0,remarks:'',quote_no:'製見-2024-***',submit_to:'井野さん',submit_date:'',file_url:''},
      {id:'QS-A81 W-0159',machine_code:'A81 W',brand:'FJ（確定）',board_name:'ALLリユース',format:'新提案',amount:1488.0,remarks:'ALL中古基板使用',quote_no:'製見-2024-0150',submit_to:'井野さん',submit_date:'2024/12/25',file_url:''},
      {id:'QS-A81 W-0160',machine_code:'A81 W',brand:'FJ（確定）',board_name:'アフター',format:'新提案',amount:25470.0,remarks:'アフター',quote_no:'製見-2024-0157',submit_to:'井野さん',submit_date:'2025/01/23',file_url:''},
      {id:'QS-D57D-0161',machine_code:'D57D',brand:'FJ（）',board_name:'FJ+M2003A5 PCB',format:'新提案',amount:679.0,remarks:'PCBのみ',quote_no:'製見-2024-0033',submit_to:'大山さん',submit_date:'',file_url:''},
      {id:'QS-D57D-0162',machine_code:'D57D',brand:'FJ（）',board_name:'FJ+M2003A5',format:'新提案',amount:3807.0,remarks:'PCB除く',quote_no:'製見2024-0092',submit_to:'大山さん',submit_date:'2024/06/13',file_url:''},
      {id:'QS-D57D-0163',machine_code:'D57D',brand:'FJ（）',board_name:'オール新品(AX51501新品)',format:'新提案',amount:25540.0,remarks:'AX51501新品',quote_no:'製見2024-0101',submit_to:'大山さん',submit_date:'2024/06/27',file_url:''},
      {id:'QS-D57D-0164',machine_code:'D57D',brand:'FJ（）',board_name:'オール新品(AX51501新品)',format:'新提案',amount:25540.0,remarks:'AX51501新品',quote_no:'製見2024-0102',submit_to:'大山さん',submit_date:'2024/06/27',file_url:''},
      {id:'QS-D57D-0165',machine_code:'D57D',brand:'FJ（）',board_name:'オール新品(AX51501新品)',format:'新提案（見本機）',amount:23827.0,remarks:'AX51501_USED品',quote_no:'製見2024-0122',submit_to:'大山さん',submit_date:'2024/10/16',file_url:''},
      {id:'QS-D57C-0166',machine_code:'D57C',brand:'FJ（）',board_name:'FJ+M2003A5 PCB',format:'新提案',amount:679.0,remarks:'PCBのみ',quote_no:'製見-2024-0033',submit_to:'大山さん',submit_date:'',file_url:''},
      {id:'QS-D57C-0167',machine_code:'D57C',brand:'FJ（）',board_name:'FJ+M2003A5',format:'新提案',amount:3807.0,remarks:'PCB除く',quote_no:'製見2024-0035',submit_to:'大山さん',submit_date:'2024/04/16',file_url:''},
      {id:'QS-D57C-0168',machine_code:'D57C',brand:'FJ（）',board_name:'FJ+M2003A5_リユース',format:'新提案',amount:612.0,remarks:'エコ／リユース',quote_no:'製見2024-0143',submit_to:'大山さん',submit_date:'2024/12/09',file_url:''},
      {id:'QS-D57C-0169',machine_code:'D57C',brand:'FJ（）',board_name:'オール新品(AX51501新品)',format:'新提案',amount:25043.0,remarks:'AX51501新品',quote_no:'製見2024-0036',submit_to:'大山さん',submit_date:'2024/04/16',file_url:''},
      {id:'QS-D57C-0170',machine_code:'D57C',brand:'FJ（）',board_name:'オール新品(AX51501中古)',format:'新提案',amount:25043.0,remarks:'AX51501中古',quote_no:'製見2024-0037',submit_to:'大山さん',submit_date:'2024/04/16',file_url:''},
      {id:'QS-D57C-0171',machine_code:'D57C',brand:'FJ（）',board_name:'オール新品(AX51501中古)',format:'新提案',amount:23827.0,remarks:'AX51501中古',quote_no:'製見2024-0128',submit_to:'大山さん',submit_date:'2024/10/16',file_url:''},
      {id:'QS-D57C-0172',machine_code:'D57C',brand:'FJ（）',board_name:'オール新品(AX51501中古)アフター',format:'新提案',amount:23820.0,remarks:'AX51501中古',quote_no:'製見2024-0128',submit_to:'大山さん',submit_date:'1016.0',file_url:''},
      {id:'QS-D57C-0173',machine_code:'D57C',brand:'FJ（）',board_name:'D基板リユース',format:'新提案',amount:11955.0,remarks:'D基板リユース',quote_no:'製見2024-0102',submit_to:'大山さん',submit_date:'2024/07/03',file_url:''},
      {id:'QS-D57C-0174',machine_code:'D57C',brand:'FJ（）',board_name:'ALL中古（ALLエコ／リユース）',format:'新提案',amount:1488.0,remarks:'ALLエコ／リユース',quote_no:'製見2024-0144',submit_to:'大山さん',submit_date:'2024/12/09',file_url:''},
      {id:'QS-D57C-0175',machine_code:'D57C',brand:'FJ（）',board_name:'DEのみ新品',format:'新提案',amount:4281.0,remarks:'DE基板のみ新品',quote_no:'製見2024-0145',submit_to:'大山さん',submit_date:'2024/12/09',file_url:''},
      {id:'QS-A83-0176',machine_code:'A83',brand:'FJ（確定）',board_name:'FJ+M2003A',format:'新提案',amount:4230.0,remarks:'PCB含む',quote_no:'製見-2024-0056',submit_to:'井野さん',submit_date:'2024/07/25',file_url:''},
      {id:'QS-A83-0177',machine_code:'A83',brand:'FJ（確定）',board_name:'FJ+M2003A PCB',format:'新提案',amount:637.0,remarks:'PCBのみ',quote_no:'',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-A83-0178',machine_code:'A83',brand:'FJ（確定）',board_name:'FJ+M2003A',format:'新提案',amount:612.0,remarks:'中古基板使用',quote_no:'製見-2024-0161',submit_to:'井野さん',submit_date:'2025/02/13',file_url:''},
      {id:'QS-A83-0179',machine_code:'A83',brand:'FJ（確定）',board_name:'オール新品',format:'新提案',amount:24851.0,remarks:'24年価格改定適用',quote_no:'製見-2024-0113',submit_to:'井野さん',submit_date:'2024/09/09',file_url:''},
      {id:'QS-A83-0180',machine_code:'A83',brand:'FJ（確定）',board_name:'FJ+DE2101A1 PCB',format:'新提案',amount:480.0,remarks:'',quote_no:'製見-2024-0112',submit_to:'井野さん',submit_date:'2024/09/09',file_url:''},
      {id:'QS-A83-0181',machine_code:'A83',brand:'FJ（確定）',board_name:'アフター',format:'新提案',amount:24844.0,remarks:'アフター',quote_no:'製見-2024-0156',submit_to:'井野さん',submit_date:'2025/01/23',file_url:''},
      {id:'QS-A83-0182',machine_code:'A83',brand:'FJ（確定）',board_name:'オール新品（VDP藤USED品使用）',format:'新提案',amount:22017.0,remarks:'24年価格改定適用',quote_no:'製見-2024-0162',submit_to:'井野さん',submit_date:'2025/02/13',file_url:''},
      {id:'QS-A83-0183',machine_code:'A83',brand:'FJ（確定）',board_name:'D基板のみリユース',format:'新提案',amount:11857.0,remarks:'24年価格改定適用',quote_no:'製見-2024-0163',submit_to:'井野さん',submit_date:'2025/02/13',file_url:''},
      {id:'QS-A83-0184',machine_code:'A83',brand:'FJ（確定）',board_name:'ALL中古基板使用',format:'新提案',amount:1488.0,remarks:'ALL中古基板使用',quote_no:'製見-2024-0164',submit_to:'井野さん',submit_date:'2025/02/13',file_url:''},
      {id:'QS-A83-0185',machine_code:'A83',brand:'FJ（確定）',board_name:'D／E　中古基板使用',format:'新提案',amount:3687.0,remarks:'D／E中古基板使用',quote_no:'製見-2024-0165',submit_to:'井野さん',submit_date:'2025/02/13',file_url:''},
      {id:'QS-A80B-0186',machine_code:'A80B',brand:'JFJ（確定）',board_name:'FJ+M2003A5',format:'新提案',amount:3807.0,remarks:'PCB含まず',quote_no:'製見-2024-0110',submit_to:'井野さん',submit_date:'2024/08/26',file_url:''},
      {id:'QS-A80B-0187',machine_code:'A80B',brand:'JFJ（確定）',board_name:'リユース／エコ',format:'新提案',amount:612.0,remarks:'中古基板使用',quote_no:'製見-2024-0140',submit_to:'井野さん',submit_date:'2024/12/02',file_url:''},
      {id:'QS-A80B-0188',machine_code:'A80B',brand:'JFJ（確定）',board_name:'オール新品',format:'新提案',amount:25477.0,remarks:'24年価格改定適用',quote_no:'製見-2024-0111',submit_to:'井野さん',submit_date:'2024/08/26',file_url:''},
      {id:'QS-A80B-0189',machine_code:'A80B',brand:'JFJ（確定）',board_name:'アフター用（中古基板使用）',format:'新提案',amount:1481.0,remarks:'アフター用',quote_no:'製見-2024-0145',submit_to:'井野さん',submit_date:'2024/12/10',file_url:''},
      {id:'QS-A80B-0190',machine_code:'A80B',brand:'JFJ（確定）',board_name:'オールリユース',format:'新提案',amount:1488.0,remarks:'ALL中古基板使用',quote_no:'製見-2024-0141',submit_to:'井野さん',submit_date:'2024/12/02',file_url:''},
      {id:'QS-A80B-0191',machine_code:'A80B',brand:'JFJ（確定）',board_name:'オール新品（VDP_FUSED使用）',format:'新提案',amount:22636.0,remarks:'藤商事USED品使用 24年価格改定適用',quote_no:'製見-2024-0141',submit_to:'井野さん',submit_date:'2024/12/05',file_url:''},
      {id:'QS-A82-0192',machine_code:'A82',brand:'JJ（確定）',board_name:'FJ+M2003A',format:'新提案',amount:4230.0,remarks:'PCB含む',quote_no:'製見-2024-0056',submit_to:'井野さん',submit_date:'2024/05/23',file_url:''},
      {id:'QS-A82-0193',machine_code:'A82',brand:'JJ（確定）',board_name:'FJ+M2003A（B36B-PUDSS-1支給）',format:'新提案',amount:4187.0,remarks:'PCB含む／CN支給',quote_no:'製見-2024-0056-1',submit_to:'井野さん',submit_date:'2024/07/31',file_url:''},
      {id:'QS-A82-0194',machine_code:'A82',brand:'JJ（確定）',board_name:'JJ+M2003A（PCB／B36B-PUDSS-1支給）',format:'新提案',amount:3550.0,remarks:'PCB含まず／CN支給',quote_no:'製見-2024-0056-2',submit_to:'井野さん',submit_date:'2024/10/15',file_url:''},
      {id:'QS-A82-0195',machine_code:'A82',brand:'JJ（確定）',board_name:'JJ+M2003A PCB',format:'新提案',amount:637.0,remarks:'PCBのみ',quote_no:'製見-2024-0121',submit_to:'井野さん',submit_date:'2024/09/25',file_url:''},
      {id:'QS-A82-0196',machine_code:'A82',brand:'JJ（確定）',board_name:'JJ+M2003A（中古基板使用）',format:'新提案',amount:612.0,remarks:'中古基板使用',quote_no:'製見-2024-0131',submit_to:'井野さん',submit_date:'2024/11/11',file_url:''},
      {id:'QS-A82-0197',machine_code:'A82',brand:'JJ（確定）',board_name:'オール新品',format:'新提案',amount:24979.0,remarks:'',quote_no:'製見-2024-0057',submit_to:'井野さん',submit_date:'2024/05/23',file_url:''},
      {id:'QS-A82-0198',machine_code:'A82',brand:'JJ（確定）',board_name:'オール新品',format:'新提案',amount:25477.0,remarks:'24年価格改定適用',quote_no:'製見-2024-0057-1',submit_to:'井野さん',submit_date:'2024/06/13',file_url:''},
      {id:'QS-A82-0199',machine_code:'A82',brand:'JJ（確定）',board_name:'オール新品（VDP_FUSED使用）',format:'新提案',amount:22643.0,remarks:'藤商事USED品使用 24年価格改定適用',quote_no:'製見-2024-0057-2',submit_to:'井野さん',submit_date:'2024/10/15',file_url:''},
      {id:'QS-A82-0200',machine_code:'A82',brand:'JJ（確定）',board_name:'オールリユース',format:'新提案',amount:1488.0,remarks:'ALL中古基板使用',quote_no:'製見-2024-125',submit_to:'井野さん',submit_date:'2024/10/15',file_url:''},
      {id:'QS-A82-0201',machine_code:'A82',brand:'JJ（確定）',board_name:'アフター',format:'新提案',amount:25470.0,remarks:'VDP新品使用 24年価格改定適用',quote_no:'製見-2024-124',submit_to:'井野さん',submit_date:'2024/10/15',file_url:''},
      {id:'QS-A82-0202',machine_code:'A82',brand:'JJ（確定）',board_name:'E基板リユース使用',format:'新提案',amount:14472.0,remarks:'藤商事USED品使用 24年価格改定適用',quote_no:'製見-2024-132',submit_to:'井野さん',submit_date:'2024/11/11',file_url:''},
      {id:'QS-A82-0203',machine_code:'A82',brand:'JJ（確定）',board_name:'D基板のみ新品（VDP_FUSED使用）',format:'新提案',amount:11647.0,remarks:'藤商事USED品使用 24年価格改定適用',quote_no:'製見-2024-133',submit_to:'井野さん',submit_date:'2024/11/11',file_url:''},
      {id:'QS-A77W-0204',machine_code:'A77W',brand:'JFJ（確定）',board_name:'新品',format:'新提案',amount:4446.0,remarks:'',quote_no:'製見2024-0073',submit_to:'井野さん',submit_date:'2024/08/29',file_url:''},
      {id:'QS-A77W-0205',machine_code:'A77W',brand:'JFJ（確定）',board_name:'リユース',format:'新提案',amount:562.0,remarks:'',quote_no:'製見2024-0074',submit_to:'井野さん',submit_date:'2024/05/29',file_url:''},
      {id:'QS-A77W-0206',machine_code:'A77W',brand:'JFJ（確定）',board_name:'JJ+M2003A PCB',format:'新提案',amount:637.0,remarks:'',quote_no:'製見2024-0121',submit_to:'井野さん',submit_date:'2024/09/25',file_url:''},
      {id:'QS-A77W-0207',machine_code:'A77W',brand:'JFJ（確定）',board_name:'オール新品(見本機のみ)',format:'新提案',amount:24979.0,remarks:'',quote_no:'',submit_to:'井野さん',submit_date:'',file_url:''},
      {id:'QS-A77W-0208',machine_code:'A77W',brand:'JFJ（確定）',board_name:'オール新品',format:'新提案',amount:25477.0,remarks:'',quote_no:'製見2024-0070-1',submit_to:'井野さん',submit_date:'2024/08/29',file_url:''},
      {id:'QS-A77W-0209',machine_code:'A77W',brand:'JFJ（確定）',board_name:'オールリユース',format:'新提案',amount:1488.0,remarks:'',quote_no:'製見2024-0071-1',submit_to:'井野さん',submit_date:'2024/08/29',file_url:''},
      {id:'QS-A77W-0210',machine_code:'A77W',brand:'JFJ（確定）',board_name:'アフター',format:'新提案',amount:25470.0,remarks:'',quote_no:'製見2024-0072',submit_to:'井野さん',submit_date:'2024/05/29',file_url:''},
      {id:'QS-A77W-0211',machine_code:'A77W',brand:'JFJ（確定）',board_name:'DEのみ新品',format:'新提案',amount:4312.0,remarks:'',quote_no:'製見2024-0077',submit_to:'井野さん',submit_date:'2024/05/29',file_url:''},
      {id:'QS-A77W-0212',machine_code:'A77W',brand:'JFJ（確定）',board_name:'オール新品（藤USED使用）',format:'新提案',amount:22643.0,remarks:'24年5月価格改定適用',quote_no:'製見2024-0113',submit_to:'井野さん',submit_date:'2024/09/13',file_url:''},
      {id:'QS-E48-0213',machine_code:'E48',brand:'RG（確定）',board_name:'新品（藤商事）',format:'新提案',amount:5354.0,remarks:'',quote_no:'製見2024-0083',submit_to:'大山さん',submit_date:'2024/06/05',file_url:''},
      {id:'QS-E48-0214',machine_code:'E48',brand:'RG（確定）',board_name:'新品（オレンジ）',format:'新提案',amount:5429.0,remarks:'',quote_no:'製見2024-0116',submit_to:'大山さん',submit_date:'2024/09/19',file_url:''},
      {id:'QS-E48-0215',machine_code:'E48',brand:'RG（確定）',board_name:'リユース',format:'新提案',amount:758.0,remarks:'E45（RG）→E48（RG）',quote_no:'製見2024-0117',submit_to:'大山さん',submit_date:'2024/09/19',file_url:''},
      {id:'QS-E48-0216',machine_code:'E48',brand:'RG（確定）',board_name:'オール新品',format:'新提案',amount:26129.0,remarks:'2024年価格改定',quote_no:'製見2024-0084',submit_to:'大山さん',submit_date:'2024/06/10',file_url:''},
      {id:'QS-E48-0217',machine_code:'E48',brand:'RG（確定）',board_name:'D/DEリユース（見本機）',format:'新提案',amount:10630.0,remarks:'2024年価格改定',quote_no:'製見2024-0108',submit_to:'大山さん',submit_date:'2024/07/24',file_url:''},
      {id:'QS-E48-0218',machine_code:'E48',brand:'RG（確定）',board_name:'オールリユース',format:'新提案',amount:1544.0,remarks:'2024年価格改定',quote_no:'製見2024-0119',submit_to:'大山さん',submit_date:'2024/09/18',file_url:''},
      {id:'QS-E48-0219',machine_code:'E48',brand:'RG（確定）',board_name:'D/DEリユース',format:'新提案',amount:10630.0,remarks:'2024年価格改定',quote_no:'製見2024-0120',submit_to:'大山さん',submit_date:'2024/09/18',file_url:''},
      {id:'QS-A79B-0220',machine_code:'A79B',brand:'RG（確定）',board_name:'RG+M2003A2',format:'新提案',amount:4609.0,remarks:'PCB含む',quote_no:'製見-2024-085',submit_to:'井野さん',submit_date:'2024/06/05',file_url:''},
      {id:'QS-A79B-0221',machine_code:'A79B',brand:'RG（確定）',board_name:'RG+M2003A2_リユース',format:'新提案',amount:612.0,remarks:'リユース／エコ',quote_no:'製見-2024-086',submit_to:'井野さん',submit_date:'2024/06/05',file_url:''},
      {id:'QS-A79B-0222',machine_code:'A79B',brand:'RG（確定）',board_name:'オール新品',format:'新提案',amount:25725.0,remarks:'',quote_no:'製見-2024-087',submit_to:'井野さん',submit_date:'2024/06/05',file_url:''},
      {id:'QS-A79B-0223',machine_code:'A79B',brand:'RG（確定）',board_name:'オール中古',format:'新提案',amount:1488.0,remarks:'',quote_no:'製見-2024-088',submit_to:'井野さん',submit_date:'2024/06/05',file_url:''},
      {id:'QS-A79B-0224',machine_code:'A79B',brand:'RG（確定）',board_name:'D/E中古使用（見本機）',format:'新提案',amount:4312.0,remarks:'',quote_no:'製見-2024-0108',submit_to:'井野さん',submit_date:'2024/07/29',file_url:''},
      {id:'QS-A79B-0225',machine_code:'A79B',brand:'RG（確定）',board_name:'アフター用',format:'新提案',amount:4306.0,remarks:'',quote_no:'製見-2024-0138',submit_to:'井野さん',submit_date:'2024/11/21',file_url:''},
      {id:'QS-D58-0226',machine_code:'D58',brand:'JFJ（確定）',board_name:'新品',format:'新',amount:3461.0,remarks:'',quote_no:'製見2024-0007',submit_to:'大山さん',submit_date:'2024/02/09',file_url:''},
      {id:'QS-D58-0227',machine_code:'D58',brand:'JFJ（確定）',board_name:'FJ+M2202A  PCB',format:'新',amount:685.0,remarks:'PCBのみ',quote_no:'製見2024-0008',submit_to:'大山さん',submit_date:'2024/02/09',file_url:''},
      {id:'QS-D58-0228',machine_code:'D58',brand:'JFJ（確定）',board_name:'JJ+M2202A（完）',format:'新提案',amount:3394.0,remarks:'',quote_no:'製見2024-0052',submit_to:'大山さん',submit_date:'2024/05/28',file_url:''},
      {id:'QS-D58-0229',machine_code:'D58',brand:'JFJ（確定）',board_name:'FJ+M2202A  PCB',format:'新提案',amount:637.0,remarks:'PCBのみ',quote_no:'製見2024-0065',submit_to:'大山さん',submit_date:'2024/05/28',file_url:''},
      {id:'QS-D58-0230',machine_code:'D58',brand:'JFJ（確定）',board_name:'長納期手配用（部品毎）',format:'',amount:null,remarks:'',quote_no:'24-MP002',submit_to:'井野さん',submit_date:'2024/02/07',file_url:''},
      {id:'QS-D58-0231',machine_code:'D58',brand:'JFJ（確定）',board_name:'オール新品_見本機',format:'旧',amount:25025.0,remarks:'見本機',quote_no:'製見2024-0015',submit_to:'大山さん',submit_date:'2024/03/07',file_url:''},
      {id:'QS-D58-0232',machine_code:'D58',brand:'JFJ（確定）',board_name:'オール新品',format:'新提案',amount:24979.0,remarks:'量産',quote_no:'製見2024-0059',submit_to:'大山さん',submit_date:'2024/05/28',file_url:''},
      {id:'QS-D58-0233',machine_code:'D58',brand:'JFJ（確定）',board_name:'オール新品_アフター',format:'新提案',amount:24973.0,remarks:'アフター',quote_no:'製見2024-0016',submit_to:'大山さん',submit_date:'2024/05/28',file_url:''},
      {id:'QS-D58-0234',machine_code:'D58',brand:'JFJ（確定）',board_name:'オール新品（D-Y品）※VDP_USED',format:'新提案',amount:23266.0,remarks:'量産（他社USED）',quote_no:'製見2024-0061',submit_to:'大山さん',submit_date:'2024/05/28',file_url:''},
      {id:'QS-D58-0235',machine_code:'D58',brand:'JFJ（確定）',board_name:'オール新品（D-X品）※VDP_USED',format:'新提案',amount:22146.0,remarks:'',quote_no:'製見2024-0062',submit_to:'大山さん',submit_date:'2024/05/28',file_url:''},
      {id:'QS-D58-0236',machine_code:'D58',brand:'JFJ（確定）',board_name:'オールリユース',format:'新提案',amount:1488.0,remarks:'ALLリユース',quote_no:'製見2024-0063',submit_to:'大山さん',submit_date:'2024/05/28',file_url:''},
      {id:'QS-D58-0237',machine_code:'D58',brand:'JFJ（確定）',board_name:'Dリユース',format:'新提案',amount:11986.0,remarks:'D基板リユース',quote_no:'製見2024-0064',submit_to:'大山さん',submit_date:'2024/05/28',file_url:''},
      {id:'QS-D58-0238',machine_code:'D58',brand:'JFJ（確定）',board_name:'D／DEリユース',format:'新提案',amount:9161.0,remarks:'D／DE基板リユース',quote_no:'製見2024-0094',submit_to:'大山さん',submit_date:'2024/06/17',file_url:''},
      {id:'QS-D58-0239',machine_code:'D58',brand:'JFJ（確定）',board_name:'オール新品_アフター（VDP_他社USED）',format:'新提案',amount:23259.0,remarks:'アフター',quote_no:'製見2024-0098',submit_to:'大山さん',submit_date:'2024/06/20',file_url:''},
      {id:'QS-PF45-0240',machine_code:'PF45',brand:'',board_name:'PF45枠払出制御基板（完）概算',format:'新',amount:4778.0,remarks:'',quote_no:'製見2023-0093',submit_to:'井野さん',submit_date:'2023/10/25',file_url:''},
      {id:'QS-PF45-0241',machine_code:'PF45',brand:'',board_name:'PF45枠払出制御基板（完）',format:'新',amount:4246.0,remarks:'SNBB028N市場品',quote_no:'製見2024-0017',submit_to:'井野さん',submit_date:'2024/03/12',file_url:''},
      {id:'QS-PF45-0242',machine_code:'PF45',brand:'',board_name:'PF45枠払出制御基板（完）',format:'新',amount:3845.0,remarks:'SNBB028N正規品',quote_no:'製見2024-0018',submit_to:'井野さん',submit_date:'2024/03/12',file_url:''},
      {id:'QS-PF45-0243',machine_code:'PF45',brand:'',board_name:'PF45枠払出制御基板（完）',format:'新提案',amount:3946.0,remarks:'SNBB028N市場品',quote_no:'製見2024-0017-1',submit_to:'井野さん',submit_date:'2024/04/02',file_url:''},
      {id:'QS-PF45-0244',machine_code:'PF45',brand:'',board_name:'PJ/PR45枠払出制御基板（完）',format:'新',amount:3920.0,remarks:'SNBB028N正規品',quote_no:'製見2024-0058',submit_to:'井野さん',submit_date:'2024/05/27',file_url:''},
      {id:'QS-PF45-0245',machine_code:'PF45',brand:'',board_name:'FJ+C2202B  PCB',format:'旧',amount:500.0,remarks:'',quote_no:'製見2024-0019',submit_to:'井野さん',submit_date:'2024/03/12',file_url:''},
      {id:'QS-PF45-0246',machine_code:'PF45',brand:'',board_name:'FJ+C2202B  PCB',format:'新提案',amount:496.0,remarks:'管理費削除　利益12％',quote_no:'製見2024-0019-1',submit_to:'井野さん',submit_date:'2024/04/02',file_url:''},
      {id:'QS-PF45-0247',machine_code:'PF45',brand:'',board_name:'C2202B PCBチェッカー',format:'',amount:220000.0,remarks:'',quote_no:'MP24-005',submit_to:'井野さん',submit_date:'2024/02/22',file_url:''},
      {id:'QS-PF45-0248',machine_code:'PF45',brand:'',board_name:'C2202B ICTピンボード',format:'',amount:456000.0,remarks:'',quote_no:'MP24-005',submit_to:'井野さん',submit_date:'2024/02/22',file_url:''},
      {id:'QS-PF45-0249',machine_code:'PF45',brand:'',board_name:'FJ+C2202B 仕掛基板',format:'',amount:3564.0,remarks:'',quote_no:'製見2024-0155',submit_to:'丸山さん',submit_date:'2025/01/15',file_url:''},
      {id:'QS-PF45-0250',machine_code:'PF45',brand:'',board_name:'JJ+C2202B 仕掛基板',format:'',amount:3639.0,remarks:'',quote_no:'製見2024-0155',submit_to:'丸山さん',submit_date:'2025/01/15',file_url:''},
      {id:'QS-PF45-0251',machine_code:'PF45',brand:'',board_name:'長納期手配用（部品毎）',format:'',amount:null,remarks:'',quote_no:'24-MP002',submit_to:'井野さん',submit_date:'2024/02/07',file_url:''},
      {id:'QS-A77B-0252',machine_code:'A77B',brand:'JFJ（確定）',board_name:'新品',format:'新提案',amount:4446.0,remarks:'',quote_no:'製見2024-0067',submit_to:'井野さん',submit_date:'2024/05/29',file_url:''},
      {id:'QS-A77B-0253',machine_code:'A77B',brand:'JFJ（確定）',board_name:'リユース',format:'新提案',amount:562.0,remarks:'',quote_no:'製見2024-0068',submit_to:'井野さん',submit_date:'2024/05/29',file_url:''},
      {id:'QS-A77B-0254',machine_code:'A77B',brand:'JFJ（確定）',board_name:'オール新品',format:'新提案',amount:25477.0,remarks:'',quote_no:'製見2024-0065',submit_to:'井野さん',submit_date:'2024/05/29',file_url:''},
      {id:'QS-A77B-0255',machine_code:'A77B',brand:'JFJ（確定）',board_name:'オールリユース',format:'新提案',amount:1488.0,remarks:'',quote_no:'製見2024-0066',submit_to:'井野さん',submit_date:'2024/05/29',file_url:''},
      {id:'QS-A77B-0256',machine_code:'A77B',brand:'JFJ（確定）',board_name:'アフター',format:'新提案',amount:25470.0,remarks:'',quote_no:'製見2024-0069',submit_to:'井野さん',submit_date:'2024/05/29',file_url:''},
      {id:'QS-A77B-0257',machine_code:'A77B',brand:'JFJ（確定）',board_name:'DEのみ新品',format:'新提案',amount:4312.0,remarks:'',quote_no:'製見2024-0095',submit_to:'井野さん',submit_date:'2024/06/17',file_url:''},
      {id:'QS-E47-0258',machine_code:'E47',brand:'JFJ',board_name:'新品',format:'新',amount:5502.0,remarks:'',quote_no:'製見2023-00122',submit_to:'大山さん',submit_date:'2023/12/20',file_url:''},
      {id:'QS-E47-0259',machine_code:'E47',brand:'JFJ',board_name:'新品',format:'新提案',amount:5429.0,remarks:'',quote_no:'製見2024-00054',submit_to:'大山さん',submit_date:'2024/05/21',file_url:''},
      {id:'QS-E47-0260',machine_code:'E47',brand:'JFJ',board_name:'リユース',format:'新提案',amount:833.0,remarks:'',quote_no:'製見2024-00022',submit_to:'大山さん',submit_date:'2024/05/21',file_url:''},
      {id:'QS-E47-0261',machine_code:'E47',brand:'JFJ',board_name:'オール新品',format:'旧',amount:25635.0,remarks:'',quote_no:'製見2023-00123',submit_to:'大山さん',submit_date:'2023/12/20',file_url:''},
      {id:'QS-E47-0262',machine_code:'E47',brand:'JFJ',board_name:'D基板リユース',format:'旧',amount:12756.0,remarks:'',quote_no:'製見2024-0002',submit_to:'大山さん',submit_date:'',file_url:''},
      {id:'QS-E47-0263',machine_code:'E47',brand:'JFJ',board_name:'D/DEリユース',format:'旧',amount:10720.0,remarks:'',quote_no:'製見2024-0005',submit_to:'大山さん',submit_date:'2024/02/08',file_url:''},
      {id:'QS-E47-0264',machine_code:'E47',brand:'JFJ',board_name:'オールリユース',format:'旧',amount:1414.0,remarks:'',quote_no:'製見2024-0006',submit_to:'大山さん',submit_date:'2024/02/08',file_url:''},
      {id:'QS-E47-0265',machine_code:'E47',brand:'JFJ',board_name:'オール新品',format:'新提案',amount:25635.0,remarks:'',quote_no:'製見2024-0023',submit_to:'大山さん',submit_date:'2024/05/21',file_url:''},
      {id:'QS-E47-0266',machine_code:'E47',brand:'JFJ',board_name:'D基板リユース',format:'新提案',amount:12357.0,remarks:'',quote_no:'製見2024-0024',submit_to:'大山さん',submit_date:'2024/05/21',file_url:''},
      {id:'QS-E47-0267',machine_code:'E47',brand:'JFJ',board_name:'D/DEリユース',format:'新提案',amount:10136.0,remarks:'',quote_no:'製見2024-0025',submit_to:'大山さん',submit_date:'2024/05/21',file_url:''},
      {id:'QS-E47-0268',machine_code:'E47',brand:'JFJ',board_name:'オールリユース',format:'新提案',amount:1544.0,remarks:'',quote_no:'製見2024-0026',submit_to:'大山さん',submit_date:'2024/05/21',file_url:''},
      {id:'QS-C45-0269',machine_code:'C45',brand:'',board_name:'FJ+M2003A5',format:'旧',amount:4726.0,remarks:'',quote_no:'製見-2023-0091',submit_to:'奥田さん',submit_date:'2023/10/19',file_url:''},
      {id:'QS-C45-0270',machine_code:'C45',brand:'',board_name:'FJ+M2003A5',format:'新提案',amount:3807.0,remarks:'PCB含まず',quote_no:'製見-2024-0048',submit_to:'井野さん',submit_date:'2024/05/14',file_url:''},
      {id:'QS-C45-0271',machine_code:'C45',brand:'',board_name:'オール新品',format:'旧',amount:25210.0,remarks:'',quote_no:'製見-2023-0090',submit_to:'奥田さん',submit_date:'2023/10/19',file_url:''},
      {id:'QS-C45-0272',machine_code:'C45',brand:'',board_name:'アフター',format:'旧',amount:25204.0,remarks:'',quote_no:'製見-2023-0125',submit_to:'奥田さん',submit_date:'2024/01/16',file_url:''},
      {id:'QS-C45-0273',machine_code:'C45',brand:'',board_name:'オール新品',format:'新提案',amount:24979.0,remarks:'',quote_no:'製見-2024-0046',submit_to:'井野さん',submit_date:'2024/05/14',file_url:''},
      {id:'QS-C45-0274',machine_code:'C45',brand:'',board_name:'アフター',format:'新提案',amount:24973.0,remarks:'',quote_no:'製見-2024-0047',submit_to:'井野さん',submit_date:'2024/05/14',file_url:''},
      {id:'QS-C45-0275',machine_code:'C45',brand:'',board_name:'オール中古',format:'新提案',amount:1488.0,remarks:'',quote_no:'製見-2024-0049',submit_to:'井野さん',submit_date:'2024/05/14',file_url:''},
      {id:'QS-C45-0276',machine_code:'C45',brand:'',board_name:'オール新品',format:'新提案',amount:25477.0,remarks:'',quote_no:'製見-2024-0046-1',submit_to:'井野さん',submit_date:'2024/09/03',file_url:''},
      {id:'QS-C45-0277',machine_code:'C45',brand:'',board_name:'オール新品',format:'新提案',amount:22643.0,remarks:'VDP藤商事USED品使用',quote_no:'製見-2024-0122',submit_to:'井野さん',submit_date:'2024/09/25',file_url:''},
      {id:'QS-A81-0278',machine_code:'A81',brand:'FJ（確定）',board_name:'FJ+M2003A5',format:'新',amount:4046.0,remarks:'PCB除く',quote_no:'製見-2023-0103',submit_to:'奥田さん',submit_date:'2023/12/20',file_url:''},
      {id:'QS-A81-0279',machine_code:'A81',brand:'FJ（確定）',board_name:'FJ+M2003A5 PCB',format:'新',amount:685.0,remarks:'PCBのみ',quote_no:'製見-2023-0104',submit_to:'奥田さん',submit_date:'2023/12/20',file_url:''},
      {id:'QS-A81-0280',machine_code:'A81',brand:'FJ（確定）',board_name:'FJ+M2003A5',format:'新提案',amount:3807.0,remarks:'PCB除く',quote_no:'製見-2024-0034',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-A81-0281',machine_code:'A81',brand:'FJ（確定）',board_name:'FJ+M2003A5 PCB',format:'新提案',amount:679.0,remarks:'PCBのみ',quote_no:'製見-2024-0033',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-A81-0282',machine_code:'A81',brand:'FJ（確定）',board_name:'オール新品',format:'旧',amount:25025.0,remarks:'',quote_no:'製見-2023-0105',submit_to:'奥田さん',submit_date:'2023/12/20',file_url:''},
      {id:'QS-A81-0283',machine_code:'A81',brand:'FJ（確定）',board_name:'オール新品',format:'新提案',amount:24979.0,remarks:'',quote_no:'製見-2024-0030',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-A81-0284',machine_code:'A81',brand:'FJ（確定）',board_name:'オール新品 アフター',format:'新提案',amount:24973.0,remarks:'',quote_no:'製見-2024-0029',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-A81-0285',machine_code:'A81',brand:'FJ（確定）',board_name:'オールリユース',format:'新提案',amount:1412.0,remarks:'',quote_no:'製見-2024-0031',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-A81-0286',machine_code:'A81',brand:'FJ（確定）',board_name:'D／DEリユース（Eのみ新品）',format:'新提案',amount:9086.0,remarks:'',quote_no:'製見-2024-0032',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-A78B-0287',machine_code:'A78B',brand:'JFJ（確定）',board_name:'JJ+M2003A12',format:'旧',amount:4670.0,remarks:'PCB含む、レーザー含む',quote_no:'製見-2023-0102',submit_to:'奥田さん',submit_date:'2023/11/14',file_url:''},
      {id:'QS-A78B-0288',machine_code:'A78B',brand:'JFJ（確定）',board_name:'JJ+M2003A12　中古品使用',format:'新',amount:624.0,remarks:'',quote_no:'製見-2024-0013',submit_to:'奥田さん',submit_date:'2024/02/20',file_url:''},
      {id:'QS-A78B-0289',machine_code:'A78B',brand:'JFJ（確定）',board_name:'オール新品（AX51501＿USED品）',format:'旧',amount:23855.0,remarks:'',quote_no:'製見-2023-0100',submit_to:'奥田さん',submit_date:'2023/11/07',file_url:''},
      {id:'QS-A78B-0290',machine_code:'A78B',brand:'JFJ（確定）',board_name:'オール新品（AX51501＿新品）',format:'旧',amount:25554.0,remarks:'使用しない（VDP新品在庫無いため）',quote_no:'製見-2023-0101',submit_to:'奥田さん',submit_date:'2023/11/09',file_url:''},
      {id:'QS-A78B-0291',machine_code:'A78B',brand:'JFJ（確定）',board_name:'オールリユース',format:'旧',amount:1362.0,remarks:'',quote_no:'製見-2024-0011',submit_to:'奥田さん',submit_date:'2024/02/20',file_url:''},
      {id:'QS-A78B-0292',machine_code:'A78B',brand:'JFJ（確定）',board_name:'アフター（新品）',format:'旧',amount:23849.0,remarks:'',quote_no:'製見-2024-0012',submit_to:'奥田さん',submit_date:'2024/02/20',file_url:''},
      {id:'QS-A78B-0293',machine_code:'A78B',brand:'JFJ（確定）',board_name:'アフター（リユース）',format:'旧',amount:1356.0,remarks:'',quote_no:'製見-2024-0021',submit_to:'奥田さん／井野さん',submit_date:'2024/03/21',file_url:''},
      {id:'QS-D57B-0294',machine_code:'D57B',brand:'FJ（確定）',board_name:'FJ+M2003A5',format:'新',amount:4046.0,remarks:'PCB除く',quote_no:'製見2023-0110',submit_to:'大山さん',submit_date:'2023/11/27',file_url:''},
      {id:'QS-D57B-0295',machine_code:'D57B',brand:'FJ（確定）',board_name:'FJ+M2003A5 PCB',format:'新',amount:685.0,remarks:'PCBのみ',quote_no:'製見2023-0111',submit_to:'大山さん',submit_date:'2023/11/27',file_url:''},
      {id:'QS-D57B-0296',machine_code:'D57B',brand:'FJ（確定）',board_name:'オール新品(AX51501USED)',format:'旧',amount:23855.0,remarks:'',quote_no:'製見2023-0085',submit_to:'大山さん',submit_date:'',file_url:''},
      {id:'QS-D57B-0297',machine_code:'D57B',brand:'FJ（確定）',board_name:'オールリユース（見本機組替）',format:'旧',amount:1362.0,remarks:'',quote_no:'製見2023-0109',submit_to:'大山さん',submit_date:'2023/11/27',file_url:''},
      {id:'QS-D57B-0298',machine_code:'D57B',brand:'FJ（確定）',board_name:'オール新品(AX51501USED)',format:'旧',amount:23854.0,remarks:'',quote_no:'製見2023-0108',submit_to:'大山さん',submit_date:'2023/11/27',file_url:''},
      {id:'QS-D57B-0299',machine_code:'D57B',brand:'FJ（確定）',board_name:'オール新品(AX51501新品)',format:'旧',amount:25552.0,remarks:'',quote_no:'製見2023-0115',submit_to:'大山さん',submit_date:'2023/11/27',file_url:''},
      {id:'QS-D57B-0300',machine_code:'D57B',brand:'FJ（確定）',board_name:'オール新品(AX51501USED)アフター',format:'旧',amount:23849.0,remarks:'',quote_no:'製見2023-0124',submit_to:'大山さん',submit_date:'2024/01/09',file_url:''},
      {id:'QS-D56-0301',machine_code:'D56',brand:'FJ（確定）',board_name:'FJ+M2003A5',format:'新',amount:4046.0,remarks:'',quote_no:'製見2023-0127',submit_to:'大山さん',submit_date:'2024/01/22',file_url:''},
      {id:'QS-D56-0302',machine_code:'D56',brand:'FJ（確定）',board_name:'FJ+M2003A5 PCB',format:'',amount:685.0,remarks:'PCBのみ',quote_no:'製見2023-0111',submit_to:'大山さん',submit_date:'D57流用',file_url:''},
      {id:'QS-D56-0303',machine_code:'D56',brand:'FJ（確定）',board_name:'オール新品',format:'旧',amount:25026.0,remarks:'OK',quote_no:'製見2023-0087',submit_to:'大山さん',submit_date:'',file_url:''},
      {id:'QS-D56-0304',machine_code:'D56',brand:'FJ（確定）',board_name:'ALL新品_アフター用',format:'旧',amount:25020.0,remarks:'',quote_no:'製見2023-0127',submit_to:'大山さん',submit_date:'2024/01/22',file_url:''},
      {id:'QS-D56-0305',machine_code:'D56',brand:'FJ（確定）',board_name:'見本機組替（オールリユース）',format:'旧',amount:1362.0,remarks:'',quote_no:'製見2024-0004',submit_to:'大山さん',submit_date:'2024/02/07',file_url:''},
      {id:'QS-D55L-0306',machine_code:'D55L',brand:'FJ（確定）',board_name:'FJ+M2003A3',format:'新',amount:4666.0,remarks:'レーザー含む',quote_no:'製見2023－0113',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-D55L-0307',machine_code:'D55L',brand:'FJ（確定）',board_name:'FJ+M2003A3　中古品使用',format:'新',amount:719.0,remarks:'レーザー含む',quote_no:'製見2023－0112-1',submit_to:'大山さん',submit_date:'2023/12/05',file_url:''},
      {id:'QS-D55L-0308',machine_code:'D55L',brand:'FJ（確定）',board_name:'ALL新品（見本機）',format:'旧旧',amount:24415.0,remarks:'22年価格改定',quote_no:'製見2023－020',submit_to:'大山さん',submit_date:'',file_url:''},
      {id:'QS-D55L-0309',machine_code:'D55L',brand:'FJ（確定）',board_name:'オールリユース',format:'旧',amount:1362.0,remarks:'',quote_no:'製見2023－0114-1',submit_to:'大山さん',submit_date:'2023/12/05',file_url:''},
      {id:'QS-D55L-0310',machine_code:'D55L',brand:'FJ（確定）',board_name:'ALL新品_リユース効果算出',format:'旧',amount:24494.0,remarks:'23年価格改定',quote_no:'製見2023－0117',submit_to:'大山さん',submit_date:'',file_url:''},
      {id:'QS-D55L-0311',machine_code:'D55L',brand:'FJ（確定）',board_name:'ALL新品_アフター用',format:'旧',amount:24489.0,remarks:'23年価格改定',quote_no:'製見2023－0119',submit_to:'大山さん',submit_date:'2023/12/20',file_url:''},
      {id:'QS-D55L-0312',machine_code:'D55L',brand:'FJ（確定）',board_name:'ALLリユース_アフター用',format:'旧',amount:1356.0,remarks:'23年価格改定',quote_no:'製見2023－0120',submit_to:'大山さん',submit_date:'2023/12/20',file_url:''},
      {id:'QS-A76BW-0313',machine_code:'A76BW',brand:'FJ（確定）',board_name:'FJ+M2003A3　中古品使用',format:'旧',amount:719.0,remarks:'レーザー含む',quote_no:'製見-2023-0062-1',submit_to:'奥田さん',submit_date:'2023/12/05',file_url:''},
      {id:'QS-A76BW-0314',machine_code:'A76BW',brand:'FJ（確定）',board_name:'FJ+M2003A3　中古品使用',format:'旧',amount:624.0,remarks:'レーザー含まず',quote_no:'製見-2023-0116',submit_to:'奥田さん',submit_date:'2023/12/05',file_url:''},
      {id:'QS-A76BW-0315',machine_code:'A76BW',brand:'FJ（確定）',board_name:'オールリユース（見本機組替）',format:'旧',amount:1362.0,remarks:'',quote_no:'製見-2023-0058',submit_to:'',submit_date:'2023/07/10',file_url:''},
      {id:'QS-A76BW-0316',machine_code:'A76BW',brand:'FJ（確定）',board_name:'ALL新品',format:'旧',amount:24060.0,remarks:'',quote_no:'製見-2023-0060',submit_to:'',submit_date:'2023/07/10',file_url:''},
      {id:'QS-A76BW-0317',machine_code:'A76BW',brand:'FJ（確定）',board_name:'オールリユース（見本機組替）',format:'旧',amount:1362.0,remarks:'A76→A76BW',quote_no:'製見-2023-0167',submit_to:'奥田さん',submit_date:'2023/12/05',file_url:''},
      {id:'QS-A76BW-0318',machine_code:'A76BW',brand:'FJ（確定）',board_name:'オールリユース（見本機組替）',format:'旧',amount:1362.0,remarks:'A76B→A76BW',quote_no:'製見-2023-0168',submit_to:'奥田さん',submit_date:'2023/12/05',file_url:''},
      {id:'QS-A76BW-0319',machine_code:'A76BW',brand:'FJ（確定）',board_name:'アフター（新品）',format:'旧',amount:24054.0,remarks:'',quote_no:'製見-2023-0121',submit_to:'奥田さん',submit_date:'2023/12/25',file_url:''},
      {id:'QS-A76BW-0320',machine_code:'A76BW',brand:'FJ（確定）',board_name:'アフター（リユース）',format:'旧',amount:1356.0,remarks:'',quote_no:'製見-2023-0123',submit_to:'奥田さん',submit_date:'2023/12/25',file_url:''},
      {id:'QS-A80-0321',machine_code:'A80',brand:'JFJ（確定）',board_name:'FJ+M2003A5',format:'旧',amount:4726.0,remarks:'',quote_no:'製見-2023-0032',submit_to:'奥田さん',submit_date:'2023/06/16',file_url:''},
      {id:'QS-A80-0322',machine_code:'A80',brand:'JFJ（確定）',board_name:'JJ+M2003A5',format:'新',amount:4048.0,remarks:'PCB除く',quote_no:'製見-2023-0094',submit_to:'奥田さん',submit_date:'2023/11/06',file_url:''},
      {id:'QS-A80-0323',machine_code:'A80',brand:'JFJ（確定）',board_name:'JJ+M2003A5 PCB',format:'新',amount:685.0,remarks:'PCBのみ',quote_no:'製見-2023-0097',submit_to:'奥田さん',submit_date:'2023/11/06',file_url:''},
      {id:'QS-A80-0324',machine_code:'A80',brand:'JFJ（確定）',board_name:'オール新品',format:'旧',amount:25210.0,remarks:'',quote_no:'製見-2023-0033',submit_to:'奥田さん',submit_date:'2023/06/16',file_url:''},
      {id:'QS-A80-0325',machine_code:'A80',brand:'JFJ（確定）',board_name:'オール新品（アフター）',format:'旧',amount:25204.0,remarks:'',quote_no:'製見-2023-0034',submit_to:'奥田さん',submit_date:'2023/06/16',file_url:''},
      {id:'QS-A80-0326',machine_code:'A80',brand:'JFJ（確定）',board_name:'オール新品（D-Y品）※VDP_USED',format:'旧',amount:23760.0,remarks:'',quote_no:'製見-2023-0095',submit_to:'奥田さん',submit_date:'2023/11/07',file_url:''},
      {id:'QS-A80-0327',machine_code:'A80',brand:'JFJ（確定）',board_name:'オールリユース（見本機組替）',format:'旧',amount:1362.0,remarks:'',quote_no:'製見-2023-0096',submit_to:'奥田さん',submit_date:'2023/11/07',file_url:''},
      {id:'QS-A80-0328',machine_code:'A80',brand:'JFJ（確定）',board_name:'Dのみリユース',format:'旧',amount:12600.0,remarks:'',quote_no:'製見-2023-0098',submit_to:'奥田さん',submit_date:'2023/11/07',file_url:''},
      {id:'QS-NA08-0329',machine_code:'NA08',brand:'RG（確定）',board_name:'NA08　ZEEG主',format:'旧',amount:477.0,remarks:'',quote_no:'製見2023-0088',submit_to:'井野さん　大山さん',submit_date:'2023/10/16',file_url:''},
      {id:'QS-NA08B-0330',machine_code:'NA08B',brand:'',board_name:'NA08B　ZEEG主',format:'旧',amount:477.0,remarks:'',quote_no:'製見2023-0118',submit_to:'大山さん',submit_date:'2023/12/12',file_url:''},
      {id:'QS-A79-0331',machine_code:'A79',brand:'RG',board_name:'RG+M2003A2',format:'旧旧',amount:4732.0,remarks:'',quote_no:'製見2023-0081',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-A79-0332',machine_code:'A79',brand:'RG',board_name:'RG+M2003A2',format:'旧',amount:4826.0,remarks:'',quote_no:'製見2022-0145-1',submit_to:'奥田さん',submit_date:'',file_url:''},
      {id:'QS-A79-0333',machine_code:'A79',brand:'RG',board_name:'RG+C1601E',format:'旧',amount:630.0,remarks:'リユース＆レーザー見積り',quote_no:'製見-2023-0086',submit_to:'山崎さん',submit_date:'2023/10/19',file_url:''},
      {id:'QS-A79-0334',machine_code:'A79',brand:'RG',board_name:'D/DEリユース',format:'旧',amount:10721.0,remarks:'',quote_no:'製見2023-0083',submit_to:'奥田さん',submit_date:'',file_url:''},
      {id:'QS-A79-0335',machine_code:'A79',brand:'RG',board_name:'ALLリユース',format:'旧',amount:1362.0,remarks:'',quote_no:'製見2023-0082',submit_to:'奥田さん',submit_date:'',file_url:''},
      {id:'QS-A79-0336',machine_code:'A79',brand:'RG',board_name:'ALL新品アフター',format:'旧',amount:25419.0,remarks:'',quote_no:'製見2023-0078',submit_to:'奥田さん',submit_date:'',file_url:''},
      {id:'QS-A79-0337',machine_code:'A79',brand:'RG',board_name:'オール新品',format:'旧',amount:25425.0,remarks:'',quote_no:'製見2023-0063',submit_to:'奥田さん',submit_date:'',file_url:''},
      {id:'QS-A79-0338',machine_code:'A79',brand:'RG',board_name:'Dリユースアフター',format:'旧',amount:12527.0,remarks:'',quote_no:'製見2023-0066',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-A79-0339',machine_code:'A79',brand:'RG',board_name:'Dリユースアフター',format:'新',amount:12507.0,remarks:'',quote_no:'製見2023-0066-1',submit_to:'奥田さん',submit_date:'2023/10/17',file_url:''},
      {id:'QS-A79-0340',machine_code:'A79',brand:'RG',board_name:'Dリユース',format:'旧',amount:12533.0,remarks:'',quote_no:'製見2023-0065',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-A79-0341',machine_code:'A79',brand:'RG',board_name:'Dリユース',format:'新',amount:12512.0,remarks:'',quote_no:'製見2023-0065-1',submit_to:'奥田さん',submit_date:'2023/08/28',file_url:''},
      {id:'QS-A79-0342',machine_code:'A79',brand:'RG',board_name:'D/Eリユース',format:'旧',amount:4548.0,remarks:'23年価格改定',quote_no:'製見2023-0083',submit_to:'奥田さん',submit_date:'',file_url:''},
      {id:'QS-A79e-0343',machine_code:'A79e',brand:'JFJ',board_name:'M2105C1',format:'旧',amount:4482.0,remarks:'FJ/JJ/RG',quote_no:'製見2023-0071',submit_to:'奥田さん',submit_date:'',file_url:''},
      {id:'QS-A79e-0344',machine_code:'A79e',brand:'JFJ',board_name:'ALL新品',format:'旧',amount:26173.0,remarks:'23年価格改定',quote_no:'製見2023-0075',submit_to:'奥田さん',submit_date:'',file_url:''},
      {id:'QS-A79e-0345',machine_code:'A79e',brand:'JFJ',board_name:'ALL新品アフター',format:'旧',amount:26167.0,remarks:'23年価格改定',quote_no:'製見2023-0077',submit_to:'奥田さん',submit_date:'',file_url:''},
      {id:'QS-A79e-0346',machine_code:'A79e',brand:'JFJ',board_name:'Dリユース',format:'旧',amount:13074.0,remarks:'23年価格改定',quote_no:'製見2023-0069',submit_to:'奥田さん',submit_date:'',file_url:''},
      {id:'QS-A79e-0347',machine_code:'A79e',brand:'JFJ',board_name:'Dリユースアフター',format:'旧',amount:13068.0,remarks:'23年価格改定',quote_no:'製見2023-0070',submit_to:'奥田さん',submit_date:'',file_url:''},
      {id:'QS-A79e-0348',machine_code:'A79e',brand:'JFJ',board_name:'ALLリユース',format:'旧',amount:1362.0,remarks:'23年価格改定',quote_no:'製見2023-0084',submit_to:'奥田さん',submit_date:'',file_url:''},
      {id:'QS-E46-0349',machine_code:'E46',brand:'JFJ',board_name:'新品',format:'旧',amount:4826.0,remarks:'',quote_no:'製見2022-0145-1',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-E46-0350',machine_code:'E46',brand:'JFJ',board_name:'D/DEリユース',format:'旧',amount:10721.0,remarks:'',quote_no:'製見2023-0080',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-E46-0351',machine_code:'E46',brand:'JFJ',board_name:'ALLリユース',format:'旧',amount:1416.0,remarks:'',quote_no:'製見2023-0079',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-E46-0352',machine_code:'E46',brand:'JFJ',board_name:'アフター(オール新品)',format:'旧',amount:25419.0,remarks:'',quote_no:'製見2023-0078',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-E46-0353',machine_code:'E46',brand:'JFJ',board_name:'オール新品',format:'旧',amount:25425.0,remarks:'',quote_no:'製見2023-0063',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-E46-0354',machine_code:'E46',brand:'JFJ',board_name:'アフター(Dリユース)',format:'旧',amount:12527.0,remarks:'',quote_no:'製見2023-0066',submit_to:'',submit_date:'',file_url:''},
      {id:'QS-E46-0355',machine_code:'E46',brand:'JFJ',board_name:'Dリユース',format:'旧',amount:12533.0,remarks:'',quote_no:'製見2023-0065',submit_to:'',submit_date:'',file_url:''},
  ];
}