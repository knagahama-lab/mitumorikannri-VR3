// ============================================================
// 58_cpm_excel.gs
// 客先品番マスタを、オフラインで使える Excel（.xlsx）にして出力する
//
//   シート「構成表入力」：客先部品コードを入れると、関数で品名・単価が出る
//   シート「マスタ」    ：客先品番マスタ（部品コード1つにつき1行）
//   シート「使い方」
//   関数は古い Excel でも動くもの（IF / IFERROR / INDEX / MATCH / TODAY / SUM）だけを使う。
//   作り方：一時的なスプレッドシートを作る → xlsx に変換して取り出す → 一時ファイルはゴミ箱へ
// ============================================================

var CPX_INPUT_ROWS = 300;   // 入力できる行数
var CPX_TOP = 5;            // 入力の1行目（4行目が見出し）

/** マスタの行（部品コード順） */
function _cpxMasterRows() {
  var master = _cpMaster(), od = _cpOrders(), today = _cpToday();
  var toDate = function (s) { var m = String(s || '').match(/^(\d{4})\/(\d{2})\/(\d{2})$/); return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : ''; };
  var nz = function (v) { return v === null || v === undefined ? '' : v; };
  return Object.keys(master).sort().map(function (code) {
    var m = master[code], orders = od.byCode[code] || [], j = _cpJudge(m, orders, today), o = orders[0] || null;
    return [m.client, code, m.name, nz(j.cur), nz(m.newPrice), toDate(m.newDate), m.newQuoteNo || '', m.quoteNo || '', nz(m.quotePrice),
            o ? toDate(o.delivery || o.date) || (o.delivery || o.date || '') : toDate(m.lastDate), o ? (o.orderNo || '') : ''];
  });
}
var CPX_MASTER_HEAD = ['取引先', '客先部品コード', '客先品名', '現行単価', '新単価', '新単価適用日', '新単価の見積No', '見積No', '見積単価', '前回の発注日', '前回の発注書No'];

/** 「構成表入力」の1行分の関数（r は行番号） */
function _cpxRowFormulas(r) {
  var M = "'マスタ'!", m = '$M' + r;
  var get = function (col) { return '=IF(' + m + '="","",IF(INDEX(' + M + '$' + col + ':$' + col + ',' + m + ')="","",INDEX(' + M + '$' + col + ':$' + col + ',' + m + ')))'; };
  return [
    get('C'),                                                                                  // C 品名
    '=IF(' + m + '="","",IF(AND(ISNUMBER(F' + r + '),OR(G' + r + '="",G' + r + '<=TODAY())),F' + r + ',IF(ISNUMBER(E' + r + '),E' + r + ',IF(ISNUMBER(INDEX(' + M + '$I:$I,' + m + ')),INDEX(' + M + '$I:$I,' + m + '),""))))',   // D 次回の想定単価
    get('D'),                                                                                  // E 現行単価
    get('E'),                                                                                  // F 新単価
    get('F'),                                                                                  // G 新単価適用日
    null,                                                                                      // H 数量（入力）
    '=IF(AND(ISNUMBER(D' + r + '),ISNUMBER(H' + r + ')),D' + r + '*H' + r + ',"")',              // I 金額
    '=IF(' + m + '="","",IF(F' + r + '<>"",' + 'IF(INDEX(' + M + '$G:$G,' + m + ')="","",INDEX(' + M + '$G:$G,' + m + ')),IF(INDEX(' + M + '$H:$H,' + m + ')="","",INDEX(' + M + '$H:$H,' + m + '))))',   // J 見積No
    null,                                                                                      // K 構成表の品名（入力）
    '=IF($B' + r + '="","",IF(' + m + '="","未登録",IF(AND(ISNUMBER(F' + r + '),G' + r + '<>"",G' + r + '>TODAY()),"新単価予定","登録済")))',   // L 照合
    '=IF($B' + r + '="","",IFERROR(MATCH(UPPER(ASC(TRIM($B' + r + '&""))),' + M + '$B:$B,0),""))',   // M マスタの行（非表示）
  ];
}

/** API：Excel を作って base64 で返す。p.model があれば、その機種の弊社品番を構成表から入れておく */
function apiCpmExcel(p) {
  p = p || {};
  var tmp = null;
  try {
    var rows = _cpxMasterRows();
    var pre = [];
    if (p.model) {
      _csReadLines().forEach(function (r) {
        if (_cpStr(r[0]) !== String(p.model) || !_cpOurs(r[4], r[5])) return;
        var code = _cpmCode(r[6]); if (!code) return;
        pre.push({ code: code, qty: _cpNum(r[9]), name: _cpStr(r[7]) });
      });
    }
    var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmm');
    var fileName = '客先品番マスタ_構成表入力' + (p.model ? '_' + p.model : '') + '_' + stamp + '.xlsx';
    tmp = SpreadsheetApp.create('tmp_' + fileName);

    // ---- マスタ ----
    var ms = tmp.getSheets()[0].setName('マスタ');
    var nM = Math.max(rows.length, 1);
    ms.getRange(1, 1, 1, CPX_MASTER_HEAD.length).setValues([CPX_MASTER_HEAD]).setFontWeight('bold').setBackground('#dbeafe');
    ms.getRange(2, 2, nM, 1).setNumberFormat('@');
    ms.getRange(2, 7, nM, 2).setNumberFormat('@');
    ms.getRange(2, 11, nM, 1).setNumberFormat('@');
    if (rows.length) ms.getRange(2, 1, rows.length, CPX_MASTER_HEAD.length).setValues(rows);
    ms.getRange(2, 4, nM, 2).setNumberFormat('#,##0.00');
    ms.getRange(2, 9, nM, 1).setNumberFormat('#,##0.00');
    ms.getRange(2, 6, nM, 1).setNumberFormat('yyyy/mm/dd');
    ms.getRange(2, 10, nM, 1).setNumberFormat('yyyy/mm/dd');
    ms.setFrozenRows(1);
    [150, 130, 260, 80, 80, 95, 110, 110, 80, 95, 110].forEach(function (w, i) { ms.setColumnWidth(i + 1, w); });

    // ---- 構成表入力 ----
    var is = tmp.insertSheet('構成表入力', 0);
    var last = CPX_TOP + CPX_INPUT_ROWS - 1;
    is.getRange('A1').setValue('構成表 単価確認シート（客先部品コードを黄色の欄に入れてください）').setFontWeight('bold').setFontSize(13);
    is.getRange('A2:B2').setValues([['機種', p.model || '']]);
    is.getRange('D2:E2').setValues([['合計金額', '=SUM(I' + CPX_TOP + ':I' + last + ')']]);
    is.getRange('G2:H2').setValues([['マスタ出力日', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd')]]);
    is.getRange('A2').setFontWeight('bold'); is.getRange('D2').setFontWeight('bold'); is.getRange('G2').setFontWeight('bold');
    is.getRange('B2').setBackground('#fef9c3');
    is.getRange('E2').setNumberFormat('#,##0.00').setFontWeight('bold');
    var head = ['No', '客先部品コード', '品名', '次回の想定単価', '現行単価', '新単価', '新単価適用日', '数量', '金額', '見積No', '構成表の品名・メモ', '照合', 'マスタ行'];
    is.getRange(4, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff').setHorizontalAlignment('center');
    is.getRange(CPX_TOP, 2, CPX_INPUT_ROWS, 1).setNumberFormat('@').setBackground('#fef9c3');
    is.getRange(CPX_TOP, 8, CPX_INPUT_ROWS, 1).setBackground('#fef9c3');
    is.getRange(CPX_TOP, 11, CPX_INPUT_ROWS, 1).setBackground('#fffbeb');

    var vals = [], fx = [];
    for (var i = 0; i < CPX_INPUT_ROWS; i++) {
      var r = CPX_TOP + i, f = _cpxRowFormulas(r), x = pre[i];
      vals.push([i + 1, x ? x.code : '']);
      fx.push(f);
    }
    is.getRange(CPX_TOP, 1, CPX_INPUT_ROWS, 2).setValues(vals);
    // 関数（数量・メモの列は値として入れる）
    var colsF = [3, 4, 5, 6, 7, 9, 10, 12, 13];
    colsF.forEach(function (c) {
      is.getRange(CPX_TOP, c, CPX_INPUT_ROWS, 1).setFormulas(fx.map(function (f) { return [f[c - 3]]; }));
    });
    if (pre.length) {
      var n = Math.min(pre.length, CPX_INPUT_ROWS);
      is.getRange(CPX_TOP, 8, n, 1).setValues(pre.slice(0, n).map(function (x) { return [x.qty === null ? '' : x.qty]; }));
      is.getRange(CPX_TOP, 11, n, 1).setValues(pre.slice(0, n).map(function (x) { return [x.name]; }));
    }
    is.getRange(CPX_TOP, 4, CPX_INPUT_ROWS, 3).setNumberFormat('#,##0.00');
    is.getRange(CPX_TOP, 9, CPX_INPUT_ROWS, 1).setNumberFormat('#,##0.00');
    is.getRange(CPX_TOP, 7, CPX_INPUT_ROWS, 1).setNumberFormat('yyyy/mm/dd');
    is.getRange(CPX_TOP, 4, CPX_INPUT_ROWS, 1).setFontWeight('bold').setFontColor('#1d4ed8');
    is.getRange(4, 1, CPX_INPUT_ROWS + 1, head.length - 1).setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
    var lr = is.getRange(CPX_TOP, 12, CPX_INPUT_ROWS, 1);
    is.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('未登録').setBackground('#fee2e2').setFontColor('#b91c1c').setRanges([lr]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('新単価予定').setBackground('#fef3c7').setFontColor('#92400e').setRanges([lr]).build(),
    ]);
    is.setFrozenRows(4);
    [45, 130, 240, 105, 85, 85, 95, 60, 100, 110, 220, 85, 60].forEach(function (w, i) { is.setColumnWidth(i + 1, w); });
    is.hideColumns(13);

    // ---- 使い方 ----
    var hs = tmp.insertSheet('使い方');
    var help = [
      ['使い方'],
      ['1. 「構成表入力」シートの黄色の欄（客先部品コード）に、構成表の部品コードを入れます（貼り付けでもかまいません）。'],
      ['2. 品名・現行単価・新単価・適用日・見積No が「マスタ」シートから自動で表示されます。'],
      ['3. 「次回の想定単価」は、適用日を過ぎた新単価があれば新単価、なければ現行単価（最新の注文の単価）、どちらもなければ見積単価です。適用日は今日の日付と比べるので、日がたつと自動で切り替わります。'],
      ['4. 数量を入れると金額が出ます。合計は2行目の「合計金額」です。'],
      ['5. 照合が「未登録」の品番は、客先品番マスタにありません。見積・注文管理システムで登録してください。'],
      ['6. 部品コードは全角・小文字・前後の空白があっても照合します（ハイフン以外の記号は外して入れてください）。'],
      ['7. 「マスタ」シートは出力した時点の内容です。新しい単価にするには、システムからもう一度 Excel を出力してください。'],
      [''],
      ['マスタ件数', rows.length],
      ['出力日時', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')],
    ];
    help.forEach(function (h, i) { hs.getRange(i + 1, 1, 1, h.length).setValues([h]); });
    hs.getRange('A1').setFontWeight('bold').setFontSize(13);
    hs.setColumnWidth(1, 900);
    is.activate();

    SpreadsheetApp.flush();
    var url = 'https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx';
    var res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error('Excel に変換できませんでした（' + res.getResponseCode() + '）');
    return { success: true, fileName: fileName, base64: Utilities.base64Encode(res.getContent()), count: rows.length, prefilled: Math.min(pre.length, CPX_INPUT_ROWS) };
  } catch (e) {
    Logger.log('[apiCpmExcel] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  } finally {
    if (tmp) { try { DriveApp.getFileById(tmp.getId()).setTrashed(true); } catch (e2) {} }
  }
}
