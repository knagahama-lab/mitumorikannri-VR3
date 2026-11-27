// ============================================================
// 57_cpm_dedupe.gs
// 客先品番マスタの重複の整理と、重複させない仕組み
//
//   同じ「取引先＋客先部品コード」の行が複数あるとき、1行にまとめる。
//     ・単価と前回の注文（前回単価・数量・発注日・発注書No・管理ID）は、前回発注日が一番新しい行の値
//     ・品名・図番・機種・区分は一番新しい行。空欄はほかの行の値で補う
//     ・見積No・見積単価・弊社基板名・備考は、空欄ならほかの行の値で補う
//     ・確認済はどれかが「確認済」なら確認済、発注回数は一番大きい値、初回発注日は一番古い日
//     ・新単価（21〜23列目）は、新単価が入っている行のうち更新日時が新しい行の値
//   まとめる前に、シートのコピー（バックアップ）を作る。
//   注文書の取込（5分ごと）のついでに、1時間に1回だけ重複を確認して自動で整理する。
// ============================================================

var CPM_DEDUPE_PROP = 'CPM_DEDUPE_LAST';
var CPM_WIDTH = 23;   // 20列（34）＋新単価3列（56）

/** 追加する直前に、シートを読み直して同じキーがあるか確認（34 の cpmOnOrderSaved から呼ぶ） */
function _cpmExistsNow(sh, key) {
  var last = sh.getLastRow();
  if (last < 2) return false;
  return sh.getRange(2, 1, last - 1, 2).getValues().some(function (r) { return _cpmKey(r[0], r[1]) === key; });
}

function _cpmDedupeRows(rows) {
  var groups = {}, order = [];
  rows.forEach(function (r) {
    if (!String(r[CPM.CODE]).trim() && !String(r[CPM.CLIENT]).trim()) return;   // 空行
    var k = _cpmKey(r[CPM.CLIENT], r[CPM.CODE]);
    if (!groups[k]) { groups[k] = []; order.push(k); }
    groups[k].push(r);
  });
  var dup = 0, merged = order.map(function (k) {
    var g = groups[k];
    if (g.length === 1) return g[0];
    dup += g.length - 1;
    var d = function (r, c) { return _ofDate(r[c]) || ''; };
    var t = function (r) { return String(r[CPM.UPDATED_AT] || ''); };
    // 前回発注日が新しい順 → 更新日時が新しい順
    var s = g.slice().sort(function (a, b) { return d(b, CPM.LAST_DATE).localeCompare(d(a, CPM.LAST_DATE)) || t(b).localeCompare(t(a)); });
    var out = s[0].slice(); while (out.length < CPM_WIDTH) out.push('');
    var blank = function (v) { return v === '' || v === null || v === undefined; };
    for (var c = 0; c < CPM_WIDTH; c++) {
      if (!blank(out[c])) continue;
      for (var i = 1; i < s.length; i++) if (!blank(s[i][c])) { out[c] = s[i][c]; break; }
    }
    out[CPM.CODE] = _cpmCode(out[CPM.CODE]);
    out[CPM.CLIENT] = _cpmClientNorm(out[CPM.CLIENT]);
    if (g.some(function (r) { return String(r[CPM.CONFIRMED]) === 'TRUE'; })) out[CPM.CONFIRMED] = 'TRUE';
    out[CPM.COUNT] = Math.max.apply(null, g.map(function (r) { return Number(r[CPM.COUNT]) || 0; }));
    var firsts = g.map(function (r) { return d(r, CPM.FIRST_DATE); }).filter(String).sort();
    if (firsts.length) out[CPM.FIRST_DATE] = firsts[0];
    // 新単価は、入っている行のうち更新日時が新しいもの
    var np = g.filter(function (r) { return !blank(r[20]); }).sort(function (a, b) { return t(b).localeCompare(t(a)); })[0];
    if (np) { out[20] = np[20]; out[21] = np[21]; out[22] = np[22]; }
    var notes = [];
    g.forEach(function (r) { var n = String(r[CPM.NOTE] || '').trim(); if (n && notes.indexOf(n) < 0) notes.push(n); });
    out[CPM.NOTE] = notes.join(' / ');
    return out;
  });
  return { rows: merged, removed: dup };
}

/** 重複をまとめる。p.dryRun=true なら件数だけ返す */
function apiCpmDedupe(p) {
  p = p || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(p.auto ? 5000 : 30000)) return { success: false, error: 'ほかの処理が実行中です。少し待ってからもう一度お試しください' };
  try {
    var sh = _cpmSheet(), last = sh.getLastRow();
    if (last < 2) return { success: true, removed: 0 };
    var rows = sh.getRange(2, 1, last - 1, CPM_WIDTH).getValues();
    var r = _cpmDedupeRows(rows);
    if (!r.removed || p.dryRun) return { success: true, removed: r.removed, dryRun: !!p.dryRun };
    var name = '客先品番マスタ_重複整理前_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmm');
    var ss = getSpreadsheet();
    if (!ss.getSheetByName(name)) sh.copyTo(ss).setName(name);
    sh.getRange(2, 1, last - 1, CPM_WIDTH).clearContent();
    sh.getRange(2, 1, r.rows.length, CPM_WIDTH).setNumberFormat('@').setValues(r.rows.map(function (x) {
      return x.map(function (v) { return Object.prototype.toString.call(v) === '[object Date]' ? Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd') : v; });
    }));
    SpreadsheetApp.flush();
    Logger.log('[cpmDedupe] 重複 ' + r.removed + '行をまとめました（バックアップ: ' + name + '）');
    return { success: true, removed: r.removed, remaining: r.rows.length, backup: name };
  } catch (e) {
    Logger.log('[apiCpmDedupe] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** processDriveImports（5分ごと）から呼ぶ：1時間に1回、重複があれば整理 */
function cpmDedupeHourly() {
  try {
    var props = PropertiesService.getScriptProperties(), now = Date.now();
    if (now - Number(props.getProperty(CPM_DEDUPE_PROP) || 0) < 3600 * 1000) return;
    props.setProperty(CPM_DEDUPE_PROP, String(now));
    apiCpmDedupe({ auto: true });
  } catch (e) { Logger.log('[cpmDedupeHourly] ' + e.message); }
}
