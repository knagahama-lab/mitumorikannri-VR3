// ============================================================
// 49_cost_assessment.gs
// 見積書コスト査定ツール
//
// 基板1枚あたりの原価・売価に加えて、
//   ・通函（基板を入れる通い箱）1箱あたりの収容枚数
//   ・トラック1便あたりの通函積載台数
//   ・1便あたりの運送費
// から「基板1枚あたりの運送費」を算出し、1枚あたりの実質利益・粗利率を試算する。
// さらに年度別の予定販売枚数を入力すると、年度ごと・累計の利益を試算できる。
//
//   保存先：スプレッドシート「コスト査定」シート（1行＝1ケース）
//   運送費/枚・利益/枚・年度別利益などの計算はクライアント側
//   (cost_assessment_js.html の caCompute()) でその場で再計算する。
//   サーバー側は入力値（原価・通函仕様・年度別計画など）の保存・読込のみを担う。
// ============================================================

var COST_ASSESSMENT_SHEET = 'コスト査定';

function _initCostAssessmentSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(COST_ASSESSMENT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(COST_ASSESSMENT_SHEET);
    var h = ['ID', 'ケース名', '通函あたり枚数', 'トラック1便あたり通函台数', '1便あたり運送費',
      '原価/枚', '売価/枚', '年度別計画JSON', 'メモ', '作成日時', '更新日時'];
    var hr = sheet.getRange(1, 1, 1, h.length);
    hr.setValues([h]);
    hr.setBackground('#FFF3E0');
    hr.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(8, 300);
    sheet.setColumnWidth(9, 220);
  }
  return sheet;
}

/** 一覧取得 */
function apiCostAssessmentList() {
  try {
    var sheet = _initCostAssessmentSheet();
    var last = sheet.getLastRow();
    if (last <= 1) return { success: true, items: [] };
    var data = sheet.getRange(2, 1, last - 1, 11).getValues();
    var items = data.map(function (r) {
      var plan = [];
      try { plan = r[7] ? JSON.parse(r[7]) : []; } catch (e) {}
      return {
        id: String(r[0] || ''),
        caseName: String(r[1] || ''),
        boardPerTote: Number(r[2] || 0),
        totePerTruck: Number(r[3] || 0),
        freightPerTruck: Number(r[4] || 0),
        unitCost: Number(r[5] || 0),
        unitPrice: Number(r[6] || 0),
        yearlyPlan: plan,
        notes: String(r[8] || ''),
        createdAt: r[9] ? Utilities.formatDate(new Date(r[9]), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : '',
        updatedAt: r[10] ? Utilities.formatDate(new Date(r[10]), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : '',
      };
    }).filter(function (x) { return x.id; });
    items.sort(function (a, b) { return (b.updatedAt || '').localeCompare(a.updatedAt || ''); });
    return { success: true, items: items };
  } catch (e) { return { success: false, error: e.message }; }
}

/** 新規作成／上書き保存。p.id があれば更新、無ければ新規作成してIDを払い出す */
function apiCostAssessmentSave(p) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    p = p || {};
    var caseName = String(p.caseName || '').trim();
    if (!caseName) return { success: false, error: 'ケース名を入力してください' };

    var sheet = _initCostAssessmentSheet();
    var now = new Date();
    var id = String(p.id || '').trim();

    var yearlyPlan = (Array.isArray(p.yearlyPlan) ? p.yearlyPlan : [])
      .map(function (row) { return { year: Number(row.year) || '', qty: Number(row.qty) || 0 }; })
      .filter(function (row) { return row.year; });

    var createdAt = now;
    var last = sheet.getLastRow();
    var rowIdx = -1;
    if (id && last > 1) {
      var ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat().map(String);
      var idx = ids.indexOf(id);
      if (idx >= 0) {
        rowIdx = idx + 2;
        var existingCreated = sheet.getRange(rowIdx, 10).getValue();
        if (existingCreated) createdAt = existingCreated;
      }
    }
    if (!id) id = Utilities.getUuid();

    var row = [
      id, caseName,
      Number(p.boardPerTote) || 0,
      Number(p.totePerTruck) || 0,
      Number(p.freightPerTruck) || 0,
      Number(p.unitCost) || 0,
      Number(p.unitPrice) || 0,
      JSON.stringify(yearlyPlan),
      String(p.notes || ''),
      createdAt, now,
    ];

    if (rowIdx > 0) {
      sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    return { success: true, id: id };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function apiCostAssessmentDelete(p) {
  try {
    var id = String((p || {}).id || '').trim();
    if (!id) return { success: false, error: 'IDが必要です' };
    var sheet = _initCostAssessmentSheet();
    var last = sheet.getLastRow();
    if (last <= 1) return { success: false, error: 'データがありません' };
    var ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat().map(String);
    var idx = ids.indexOf(id);
    if (idx < 0) return { success: false, error: 'ケースが見つかりません' };
    sheet.deleteRow(idx + 2);
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}
