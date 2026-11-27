// ============================================================
// 59_ship_check.gs
// 出荷明細（社内の出荷一覧 xlsx）の単価照合
//
//   出荷明細の「部品コード」は社内コード（例 1B184-HB239）なので、客先部品コードへ次の順で紐づける
//     ① 備考・受注No の「発注伝票番号-行番号」（例 206856-001）→ 注文書の明細 → 客先部品コード・注文単価
//     ② 客先品番マスタの「社内コード（出荷明細）」列（24列目）…①で紐づいたものは自動で覚える／手で登録もできる
//     ③ 客先品名が客先品番マスタの品名と同じ
//   照合する単価＝出荷日の時点の単価：適用日を過ぎた新単価 → 注文書の単価 → 最新の注文単価 → 見積単価
//   見積書・注文書のPDFリンクを各行に付けて、画面表示と Excel 出力をする
// ============================================================

var SC_COL_INNER = 24;   // 客先品番マスタ：社内コード（出荷明細）
var SC_HEAD_INNER = '社内コード（出荷明細）';

function _scInnerKey(s) { return String(s == null ? '' : s).normalize('NFKC').toUpperCase().replace(/\s+/g, ''); }
function _scDate(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd');
  return _cpDate(v);
}

/** 出荷明細の xlsx（base64）を読む：見出し行（部品コード・単価がある行）から下 */
function _scReadXlsx(base64, fileName) {
  var blob = Utilities.newBlob(Utilities.base64Decode(base64), MimeType.MICROSOFT_EXCEL, fileName || 'ship.xlsx');
  var conv = Drive.Files.create({ name: '[出荷明細照合] ' + (fileName || ''), mimeType: MimeType.GOOGLE_SHEETS }, blob);
  try {
    var v = SpreadsheetApp.openById(conv.id).getSheets()[0].getDataRange().getValues();
    var hi = -1;
    for (var i = 0; i < Math.min(v.length, 20); i++) {
      var t = v[i].map(function (x) { return String(x).trim(); });
      if (t.indexOf('部品コード') >= 0 && t.indexOf('単価') >= 0) { hi = i; break; }
    }
    if (hi < 0) throw new Error('見出し行（部品コード・単価）が見つかりません');
    var h = v[hi].map(function (x) { return String(x).trim(); });
    var c = function (name) { return h.indexOf(name); };
    var C = { date: c('出荷日'), client: c('客先'), model: c('機種'), code: c('部品コード'), name: c('品名'), cname: c('客先品名'),
              qty: c('出荷数'), price: c('単価'), amount: c('金額'), dest: c('納入先'), note: c('備考'), orderNo: c('受注No') };
    var g = function (r, k) { return C[k] >= 0 ? r[C[k]] : ''; };
    var rows = [];
    for (var j = hi + 1; j < v.length; j++) {
      var r = v[j];
      if (!String(g(r, 'code')).trim()) continue;
      rows.push({ date: _scDate(g(r, 'date')), client: _cpStr(g(r, 'client')), model: _cpStr(g(r, 'model')), inner: _cpStr(g(r, 'code')),
                  name: _cpStr(g(r, 'name')), cname: _cpStr(g(r, 'cname')), qty: _cpNum(g(r, 'qty')), price: _cpNum(g(r, 'price')),
                  amount: _cpNum(g(r, 'amount')), dest: _cpStr(g(r, 'dest')), note: _cpStr(g(r, 'note')), orderNo: _cpStr(g(r, 'orderNo')) });
    }
    return rows;
  } finally {
    try { Drive.Files.update({ trashed: true }, conv.id); } catch (e) { try { DriveApp.getFileById(conv.id).setTrashed(true); } catch (e2) {} }
  }
}

/** 照合に使う索引 */
function _scIndex() {
  var ss = getSpreadsheet();
  // 注文書の明細：伝票番号-行番号 → 明細
  var lines = _cpmOrderLines(), bySlip = {};
  var os = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  if (lines.length) {
    var slip = os.getRange(2, ORDER_COLS.ORDER_SLIP_NO, os.getLastRow() - 1, 1).getValues();
    lines.forEach(function (l) {
      var s = _cpStr(slip[l.row - 2][0]).replace(/\D/g, '');
      if (s && l.lineNo) bySlip[s + '-' + l.lineNo] = l;
    });
  }
  // 見積書PDF：見積番号 → URL
  var quoteUrl = {}, mg = getAllMgmtData();
  mg.forEach(function (r) {
    var no = _cpStr(r[MGMT_COLS.QUOTE_NO - 1]).toUpperCase(), u = _cpStr(r[MGMT_COLS.QUOTE_PDF_URL - 1]);
    if (no && u && !quoteUrl[no]) quoteUrl[no] = u;
  });
  // 客先品番マスタ：社内コード → 客先部品コード、品名 → 客先部品コード
  var sh = _cpmSheet(), last = sh.getLastRow(), byInner = {}, byName = {}, rowOf = {};
  if (last >= 2) {
    var w = Math.min(Math.max(sh.getLastColumn(), SC_COL_INNER), sh.getMaxColumns());
    sh.getRange(2, 1, last - 1, w).getValues().forEach(function (r, i) {
      var code = _cpmCode(r[CPM.CODE]); if (!code) return;
      if (!rowOf[code]) rowOf[code] = { row: i + 2, inner: _cpStr(r[SC_COL_INNER - 1]) };
      String(r[SC_COL_INNER - 1] || '').split(/[,、\/]/).forEach(function (x) { var k = _scInnerKey(x); if (k && !byInner[k]) byInner[k] = code; });
      var nk = _cpNameKey(r[CPM.NAME]); if (nk && !byName[nk]) byName[nk] = code;
    });
  }
  return { bySlip: bySlip, quoteUrl: quoteUrl, byInner: byInner, byName: byName, rowOf: rowOf, master: _cpMaster(), orders: _cpOrders() };
}

/** 1行の照合 */
function _scJudgeRow(s, ix) {
  var refs = [], re = /(\d{6})-(\d{1,3})/g, m, text = s.note + ' ' + s.orderNo;
  while ((m = re.exec(text))) refs.push(m[1] + '-' + Number(m[2]));
  var ol = null, via = '';
  for (var i = 0; i < refs.length && !ol; i++) if (ix.bySlip[refs[i]]) ol = ix.bySlip[refs[i]];
  var code = '';
  if (ol) { code = _cpmCode(ol.storedCode || ol.code); via = '伝票番号 ' + refs[0].replace(/-(\d+)$/, function (a, n) { return '-' + ('00' + n).slice(-3); }); }
  if (!code && ix.byInner[_scInnerKey(s.inner)]) { code = ix.byInner[_scInnerKey(s.inner)]; via = '社内コード（登録済み）'; }
  if (!code && s.cname && ix.byName[_cpNameKey(s.cname)]) { code = ix.byName[_cpNameKey(s.cname)]; via = '客先品名'; }
  var mm = code ? ix.master[code] || null : null;
  var date = s.date || _cpToday();
  // 照合する単価（出荷日の時点）
  var exp = null, src = '', qno = '';
  if (mm && mm.newPrice !== null && mm.newPrice !== undefined && (!mm.newDate || mm.newDate <= date)) { exp = mm.newPrice; src = '新単価' + (mm.newDate ? '（' + mm.newDate + '〜）' : ''); qno = mm.newQuoteNo || mm.quoteNo; }
  else if (ol && ol.price) { exp = ol.price; src = '注文書の単価'; qno = ol.linkedQuote || (mm && mm.quoteNo) || ''; }
  else if (code && (ix.orders.byCode[code] || [])[0]) { var lo = ix.orders.byCode[code][0]; exp = lo.price; src = '最新の注文（' + lo.orderNo + '）'; qno = (mm && mm.quoteNo) || ''; }
  else if (mm && mm.quotePrice) { exp = mm.quotePrice; src = '見積単価'; qno = mm.quoteNo; }
  if (!qno && mm) qno = mm.quoteNo || '';
  var judge = '', diff = null;
  if (!code) judge = '紐づけなし';
  else if (exp === null) judge = '単価なし';
  else if (!s.price) judge = '出荷単価0';
  else { diff = Math.round((s.price - exp) * 100) / 100; judge = Math.abs(diff) < 0.005 ? '一致' : '不一致'; }
  var orderUrl = ol ? ol.orderPdfUrl : (code && (ix.orders.byCode[code] || [])[0] ? ix.orders.byCode[code][0].pdfUrl : '');
  return {
    date: s.date, model: s.model, inner: s.inner, name: s.name, cname: s.cname, qty: s.qty, price: s.price, amount: s.amount, dest: s.dest, note: s.note,
    code: code, via: via, masterName: mm ? mm.name : '', expect: exp, expSrc: src, judge: judge, diff: diff,
    orderNo: ol ? ol.orderNo : '', slip: refs.length ? refs[0].replace(/-(\d+)$/,function (a, n) { return '-' + ('00' + n).slice(-3); }) : '', orderPrice: ol ? ol.price : null, orderUrl: orderUrl || '',
    quoteNo: qno || '', quoteUrl: qno ? (ix.quoteUrl[String(qno).toUpperCase()] || '') : '',
    learn: (ol && code && ix.rowOf[code] && ix.rowOf[code].inner.split(/[,、\/]/).map(_scInnerKey).indexOf(_scInnerKey(s.inner)) < 0) ? ix.rowOf[code].row : 0,
  };
}

/** 社内コードを客先品番マスタに覚えさせる（既存の値に追記） */
function _scLearn(sh, row, inner) {
  if (sh.getMaxColumns() < SC_COL_INNER) sh.insertColumnsAfter(sh.getMaxColumns(), SC_COL_INNER - sh.getMaxColumns());
  if (String(sh.getRange(1, SC_COL_INNER).getValue()) !== SC_HEAD_INNER) sh.getRange(1, SC_COL_INNER).setValue(SC_HEAD_INNER).setFontWeight('bold').setBackground('#DCFCE7');
  var cell = sh.getRange(row, SC_COL_INNER), cur = _cpStr(cell.getValue());
  var list = cur ? cur.split(/[,、\/]/).map(function (x) { return x.trim(); }).filter(String) : [];
  if (list.map(_scInnerKey).indexOf(_scInnerKey(inner)) >= 0) return false;
  list.push(inner);
  cell.setNumberFormat('@').setValue(list.join(','));
  return true;
}

/** API：出荷明細を照合する。p = { base64, fileName, excel:true } */
function apiShipCheck(p) {
  try {
    p = p || {};
    if (!p.base64) return { success: false, error: 'ファイルがありません' };
    var ship = _scReadXlsx(p.base64, p.fileName);
    if (!ship.length) return { success: false, error: '出荷明細の行がありません' };
    var ix = _scIndex(), sh = _cpmSheet(), learned = 0;
    var rows = ship.map(function (s) { return _scJudgeRow(s, ix); });
    rows.forEach(function (r) { if (r.learn && _scLearn(sh, r.learn, r.inner)) { learned++; ix.rowOf[r.code].inner += ',' + r.inner; } delete r.learn; });
    var sum = {};
    rows.forEach(function (r) { sum[r.judge] = (sum[r.judge] || 0) + 1; });
    var out = { success: true, rows: rows, summary: sum, learned: learned, period: ship[0].date + (ship[ship.length - 1].date !== ship[0].date ? '〜' + ship[ship.length - 1].date : '') };
    if (p.excel) { var x = _scExcel(rows, p.fileName); out.fileName = x.fileName; out.base64 = x.base64; }
    return out;
  } catch (e) {
    Logger.log('[apiShipCheck] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

/** API：社内コードと客先部品コードの紐づけを手で登録する。p = { inner, code } */
function apiShipMap(p) {
  try {
    var code = _cpmCode(p && p.code), inner = _cpStr(p && p.inner);
    if (!code || !inner) return { success: false, error: '社内コードと客先部品コードを入れてください' };
    var sh = _cpmSheet(), last = sh.getLastRow();
    var codes = last >= 2 ? sh.getRange(2, CPM.CODE + 1, last - 1, 1).getValues() : [];
    for (var i = 0; i < codes.length; i++) {
      if (_cpmCode(codes[i][0]) === code) { _scLearn(sh, i + 2, inner); return { success: true, row: i + 2 }; }
    }
    return { success: false, error: '客先品番マスタに ' + code + ' がありません。先に品番を登録してください' };
  } catch (e) { return { success: false, error: e.message }; }
}

/** 照合結果の Excel（見積書・注文書はリンク付き） */
function _scExcel(rows, srcName) {
  var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmm');
  var fileName = '出荷明細_単価照合_' + stamp + '.xlsx', tmp = SpreadsheetApp.create('tmp_' + fileName);
  try {
    var sh = tmp.getSheets()[0].setName('単価照合');
    var head = ['出荷日', '機種', '社内コード', '品名', '客先品名', '出荷数', '出荷単価', '金額', '客先部品コード', '紐づけ', 'マスタの品名',
                '照合単価', '照合単価の出どころ', '判定', '差額', '伝票番号-行', '注文単価', '見積No', '見積書', '注文書', '備考'];
    var q = function (s) { return String(s).replace(/"/g, '""'); };
    var link = function (u, t) { return u ? '=HYPERLINK("' + q(u) + '","' + q(t) + '")' : ''; };
    var nz = function (v) { return v === null || v === undefined ? '' : v; };
    var data = rows.map(function (r) {
      return [r.date, r.model, r.inner, r.name, r.cname, nz(r.qty), nz(r.price), nz(r.amount), r.code, r.via, r.masterName,
              nz(r.expect), r.expSrc, r.judge, nz(r.diff), r.slip, nz(r.orderPrice), r.quoteNo,
              link(r.quoteUrl, r.quoteNo ? '見積書 ' + r.quoteNo : ''), link(r.orderUrl, r.orderNo ? '注文書 ' + r.orderNo : '注文書'), r.note];
    });
    sh.getRange(1, 1).setValue('出荷明細 単価照合（' + (srcName || '') + '） 作成 ' + stamp).setFontWeight('bold');
    sh.getRange(3, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
    sh.getRange(4, 3, data.length, 1).setNumberFormat('@'); sh.getRange(4, 9, data.length, 1).setNumberFormat('@');
    sh.getRange(4, 1, data.length, head.length).setValues(data);
    [7, 8, 12, 15, 17].forEach(function (c) { sh.getRange(4, c, data.length, 1).setNumberFormat('#,##0.##'); });
    var jr = sh.getRange(4, 14, data.length, 1);
    var rule = function (t, bg, fg) { return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t).setBackground(bg).setFontColor(fg).setRanges([jr]).build(); };
    sh.setConditionalFormatRules([rule('一致', '#dcfce7', '#166534'), rule('不一致', '#fee2e2', '#b91c1c'), rule('紐づけなし', '#fef3c7', '#92400e'), rule('単価なし', '#fef3c7', '#92400e'), rule('出荷単価0', '#fef3c7', '#92400e')]);
    sh.getRange(3, 1, data.length + 1, head.length).setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
    sh.setFrozenRows(3);
    [80, 55, 120, 170, 260, 60, 75, 95, 100, 130, 220, 75, 150, 70, 70, 95, 70, 110, 130, 110, 200].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    SpreadsheetApp.flush();
    var res = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx',
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error('Excel に変換できませんでした（' + res.getResponseCode() + '）');
    return { fileName: fileName, base64: Utilities.base64Encode(res.getContent()) };
  } finally {
    try { DriveApp.getFileById(tmp.getId()).setTrashed(true); } catch (e) {}
  }
}
