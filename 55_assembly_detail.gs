// ============================================================
// 55_assembly_detail.gs
// 組立費 構成一覧（52）の編集機能
//   ① 手修正   … 見積ごとに 種類・使用基板・区分・世代・枠構成・提出日・金額・件名 を上書き（元の見積データは変えない）
//   ② 見積内訳 … 組立費見積書「見積詳細」と同じ項目（材料費／Assy費／組立／検査／梱包／輸送費／小計／管理費／利益／レーザーマーキング／値引き）
//   ③ 枠構成   … ハードウェア構成一覧表（機種別）を取り込み、機種ごとの枠（PFK25・SFK15 等）と基板構成を管理
//   シート: 組立費_手修正 / 組立費_見積内訳 / 枠構成_機種
// ============================================================

var AD_SHEET_OV = '組立費_手修正';
var AD_SHEET_DETAIL = '組立費_見積内訳';
var AD_SHEET_FRAME = '枠構成_機種';
var AD_OV_FIELDS = ['type', 'boards', 'variant', 'gen', 'frame', 'date', 'amount', 'subject'];
var AD_OV_HEADERS = ['見積キー', '種類', '使用基板', '区分', '世代', '枠構成', '提出日', '金額', '件名', '更新者', '更新日時'];
// 見積詳細の区分（小計に入る6つ → 管理費 → 利益 → レーザーマーキング → 値引き）
var AD_SUB = ['材料費', 'Assy費', '組立', '検査', '梱包', '輸送費'];
var AD_SECTIONS = AD_SUB.concat(['管理費', 'レーザーマーキング', '値引き']);
var AD_DETAIL_HEADERS = ['見積キー', '見積No'].concat(AD_SUB).concat(['小計', '管理費', '利益率%', '利益', 'レーザーマーキング', '値引き', '合計', '明細JSON', '更新者', '更新日時']);
var AD_FR_FIELDS = ['machine', 'title', 'kind', 'frame', 'M', 'D', 'vdp', 'DE', 'E', 'C', 'memo', 'version'];
var AD_FR_HEADERS = ['機種', 'タイトル', '種別', '枠構成', '主制御基板', '液晶制御基板', 'VDP', '液晶IF基板', '演出IF基板', '枠制御・メダル数制御基板', '備考', '一覧表の版', '手修正した項目', '更新日時'];

function _adSheet(name, headers) {
  var ss = getSpreadsheet(), sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (sh.getMaxColumns() < headers.length) sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
    sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#FFF3E0');
    sh.setFrozenRows(1);
  }
  return sh;
}
function _adRead(name, headers) {
  var sh = _adSheet(name, headers), last = sh.getLastRow();
  return last < 2 ? [] : sh.getRange(2, 1, last - 1, headers.length).getValues();
}
function _adUser() { try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; } }
function _adNum(v) {
  if (v === '' || v === null || v === undefined) return null;
  var s = String(v).normalize('NFKC').replace(/[,¥円%\s]/g, '').replace(/[−‐―ー]/g, '-');
  if (s === '') return null;
  var n = Number(s);
  if (isNaN(n)) throw new Error('金額は数値で入力してください（' + v + '）');
  return Math.round(n * 100) / 100;
}
function _adR2(n) { return Math.round((n + (n >= 0 ? 1e-9 : -1e-9)) * 100) / 100; }
/** キーの行へ書き込む（無ければ追加）。文字列固定の列なので文字で保存する */
function _adPut(name, headers, key, row) {
  var sh = _adSheet(name, headers), keys = _adRead(name, headers).map(function (r) { return String(r[0]); });
  var i = keys.indexOf(String(key)), at = i >= 0 ? i + 2 : sh.getLastRow() + 1;
  if (at > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
  sh.getRange(at, 1, 1, headers.length).setNumberFormat('@')
    .setValues([row.map(function (v) { return v === null || v === undefined ? '' : String(v); })]);
}
function _adLocked(fn) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); return fn(); }
  catch (e) { Logger.log('[55] ' + e.message + '\n' + e.stack); return { success: false, error: e.message }; }
  finally { try { lock.releaseLock(); } catch (e2) {} }
}

// ── ① 手修正 ──────────────────────────────────────────────
function _adOverrides() {
  var out = {};
  _adRead(AD_SHEET_OV, AD_OV_HEADERS).forEach(function (r) {
    if (!r[0]) return;
    var o = {};
    AD_OV_FIELDS.forEach(function (f, i) { var v = String(r[i + 1] === null ? '' : r[i + 1]).trim(); if (v !== '') o[f] = v; });
    out[String(r[0])] = o;
  });
  return out;
}
/** 手入力の基板名を配列へ（品番の形でなくても、そのままの名前で持つ） */
function _adSplitBoards(s) {
  var seen = {};
  return String(s || '').normalize('NFKC').toUpperCase().replace(/FJ\s*[+\-]\s*/g, '').split(/[\/,、+\n]+/)
    .map(function (x) { return x.trim(); })
    .filter(function (x) { if (!x || seen[x]) return false; seen[x] = 1; return true; });
}
/** p: { keys:[見積キー], field, value }  value 空 = 自動判定に戻す */
function apiAsmOverrideSave(p) {
  return _adLocked(function () {
    p = p || {};
    var keys = (p.keys || []).map(String).filter(String), fi = AD_OV_FIELDS.indexOf(p.field);
    if (!keys.length || fi < 0) return { success: false, error: '対象が特定できません' };
    var v = String(p.value === null || p.value === undefined ? '' : p.value).trim();
    if (p.field === 'amount' && v !== '') v = String(_adNum(v));
    if (p.field === 'date' && v !== '') { v = _aqDate(v); if (!v) return { success: false, error: '日付が正しくありません' }; }
    if (p.field === 'boards' && v !== '') v = _adSplitBoards(v).join(' / ');
    var cur = _adOverrides(), now = nowJST(), user = _adUser();
    keys.forEach(function (k) {
      var o = cur[k] || {}; o[p.field] = v;
      _adPut(AD_SHEET_OV, AD_OV_HEADERS, k, [k].concat(AD_OV_FIELDS.map(function (f) { return o[f] || ''; })).concat([user, now]));
    });
    return { success: true, value: v };
  });
}

// ── ② 見積内訳 ────────────────────────────────────────────
/** 明細行から 区分計・小計・利益・合計 を計算。d: { lines:[{sec,name,price,note}], rate, profit } */
function _adCalc(d) {
  var s = {}; AD_SECTIONS.forEach(function (k) { s[k] = 0; });
  (d.lines || []).forEach(function (l) { if (s[l.sec] !== undefined) s[l.sec] = _adR2(s[l.sec] + (Number(l.price) || 0)); });
  var has = function (x) { return x !== '' && x !== null && x !== undefined; };
  var sub = _adR2(AD_SUB.reduce(function (a, k) { return a + s[k]; }, 0));
  var rate = has(d.rate) ? Number(d.rate) : null;
  var profit = has(d.profit) ? Number(d.profit) : (rate === null ? 0 : _adR2(sub * rate / 100));
  return { sums: s, sub: sub, rate: rate, profit: profit, total: _adR2(sub + s['管理費'] + profit + s['レーザーマーキング'] + s['値引き']) };
}
function _adDetails() {
  var out = {}, j = AD_DETAIL_HEADERS.indexOf('明細JSON');
  _adRead(AD_SHEET_DETAIL, AD_DETAIL_HEADERS).forEach(function (r) {
    if (!r[0]) return;
    try { var d = JSON.parse(r[j] || '{}'); out[String(r[0])] = { lines: d.lines || [], rate: d.rate, profit: d.profit, calc: _adCalc(d) }; } catch (e) {}
  });
  return out;
}
/** p: { key, quoteNo, lines, rate, profit, remove } */
function apiAsmDetailSave(p) {
  return _adLocked(function () {
    p = p || {};
    if (!p.key) return { success: false, error: '見積が特定できません' };
    if (p.remove) {
      var keys = _adRead(AD_SHEET_DETAIL, AD_DETAIL_HEADERS).map(function (r) { return String(r[0]); }), i = keys.indexOf(String(p.key));
      if (i >= 0) _adSheet(AD_SHEET_DETAIL, AD_DETAIL_HEADERS).deleteRow(i + 2);
      return { success: true, detail: null };
    }
    var lines = (p.lines || []).map(function (l) {
      return { sec: String(l.sec || ''), name: String(l.name || '').trim(), price: _adNum(l.price) || 0, note: String(l.note || '').trim() };
    }).filter(function (l) { return AD_SECTIONS.indexOf(l.sec) >= 0 && (l.name || l.price || l.note); });
    var d = { lines: lines, rate: _adNum(p.rate), profit: _adNum(p.profit) };
    var c = _adCalc(d);
    _adPut(AD_SHEET_DETAIL, AD_DETAIL_HEADERS, p.key, [p.key, p.quoteNo || ''].concat(AD_SUB.map(function (k) { return c.sums[k]; }))
      .concat([c.sub, c.sums['管理費'], c.rate, c.profit, c.sums['レーザーマーキング'], c.sums['値引き'], c.total, JSON.stringify(d), _adUser(), nowJST()]));
    return { success: true, detail: { lines: d.lines, rate: d.rate, profit: d.profit, calc: c } };
  });
}

// ── ③ 枠構成（ハードウェア構成一覧表）─────────────────────
function _adMachineKey(s) { return String(s || '').normalize('NFKC').toUpperCase().replace(/\s/g, ''); }
function _adFrames() {
  return _adRead(AD_SHEET_FRAME, AD_FR_HEADERS).filter(function (r) { return r[0]; }).map(function (r) {
    var o = {}; AD_FR_FIELDS.forEach(function (f, i) { o[f] = String(r[i] === null ? '' : r[i]); });
    o.edited = String(r[12] || '').split(',').filter(String);
    return o;
  });
}
function _adFrameRow(o) { return AD_FR_FIELDS.map(function (f) { return o[f] || ''; }).concat([(o.edited || []).join(','), nowJST()]); }
/** 一覧表の取込 p: { rows:[{machine,...}], version }。手で直した項目と、より新しい版の値は上書きしない */
function apiFrameImport(p) {
  return _adLocked(function () {
    var rows = (p && p.rows) || [], ver = String((p && p.version) || '');
    if (!rows.length) return { success: false, error: '取り込む機種がありません' };
    var cur = _adFrames(), idx = {}, add = 0, upd = 0, skip = 0;
    cur.forEach(function (o, i) { idx[_adMachineKey(o.machine)] = i; });
    rows.forEach(function (r) {
      var k = _adMachineKey(r.machine); if (!k) return;
      if (idx[k] === undefined) {
        var n = { edited: [] };
        AD_FR_FIELDS.forEach(function (f) { n[f] = r[f] ? String(r[f]) : ''; });
        n.version = ver; idx[k] = cur.length; cur.push(n); add++; return;
      }
      var o = cur[idx[k]], ch = false;
      if (o.version && ver && o.version > ver) { skip++; return; }
      ['title', 'kind', 'frame', 'M', 'D', 'vdp', 'DE', 'E', 'C'].forEach(function (f) {
        if (!r[f] || o.edited.indexOf(f) >= 0 || o[f] === String(r[f])) return;
        o[f] = String(r[f]); ch = true;
      });
      if (ver) o.version = ver;
      if (ch) upd++;
    });
    var sh = _adSheet(AD_SHEET_FRAME, AD_FR_HEADERS);
    if (cur.length > sh.getMaxRows() - 1) sh.insertRowsAfter(sh.getMaxRows(), cur.length);
    sh.getRange(2, 1, cur.length, AD_FR_HEADERS.length).setNumberFormat('@').setValues(cur.map(_adFrameRow));
    return { success: true, added: add, updated: upd, skipped: skip, frames: _adFrames() };
  });
}
/** 1項目の編集・機種の追加・削除 p: { machine, field, value } / { add } / { machine, remove } */
function apiFrameSave(p) {
  return _adLocked(function () {
    p = p || {};
    var cur = _adFrames(), sh = _adSheet(AD_SHEET_FRAME, AD_FR_HEADERS);
    if (p.add) {
      var name = String(p.add).trim();
      if (!name) return { success: false, error: '機種コードを入力してください' };
      if (cur.some(function (o) { return _adMachineKey(o.machine) === _adMachineKey(name); })) return { success: false, error: name + ' は登録済みです' };
      _adPut(AD_SHEET_FRAME, AD_FR_HEADERS, name, _adFrameRow({ machine: name, edited: [] }));
      return { success: true, frames: _adFrames() };
    }
    var o = cur.filter(function (x) { return x.machine === p.machine; })[0];
    if (!o) return { success: false, error: '機種が見つかりません' };
    if (p.remove) {
      var i = _adRead(AD_SHEET_FRAME, AD_FR_HEADERS).map(function (r) { return String(r[0]); }).indexOf(String(p.machine));
      sh.deleteRow(i + 2);
      return { success: true, frames: _adFrames() };
    }
    if (AD_FR_FIELDS.indexOf(p.field) < 1) return { success: false, error: '編集できない項目です' };
    o[p.field] = String(p.value === null || p.value === undefined ? '' : p.value).trim();
    if (o.edited.indexOf(p.field) < 0) o.edited.push(p.field);
    _adPut(AD_SHEET_FRAME, AD_FR_HEADERS, o.machine, _adFrameRow(o));
    return { success: true };
  });
}

/** 文字列に含まれる機種コード（マスタにあるもの。長い名前を優先） */
function _adFindMachine(text, frames) {
  var t = ' ' + String(text || '').normalize('NFKC').toUpperCase() + ' ', hit = null;
  frames.forEach(function (o) {
    var k = _adMachineKey(o.machine); if (k.length < 2) return;
    var i = t.indexOf(k);
    while (i >= 0) {
      if (!/[A-Z0-9]/.test(t.charAt(i - 1)) && !/[A-Z0-9]/.test(t.charAt(i + k.length)) && (!hit || k.length > _adMachineKey(hit.machine).length)) hit = o;
      i = t.indexOf(k, i + 1);
    }
  });
  return hit;
}
/** 機種の基板（種類ごと） */
function _adFrameBoards(o) {
  var out = {};
  ['M', 'D', 'DE', 'E', 'C'].forEach(function (k) {
    var t = _aqTypeForBoard(k); if (!t) return;
    (out[t] = out[t] || []).push.apply(out[t], _aqBoards(o[k]));
  });
  return out;
}
/**
 * 見積の枠構成：機種名が取れればその枠。
 * 無ければ同じ基板構成を使う機種の枠（【E**】なら E で始まる機種だけ）を候補にし、1種類に決まる時だけ推定
 */
function _adGuessFrame(q, type, boards, frames) {
  var m = _adFindMachine([q.modelCode || '', q.subject || ''].join(' '), frames);
  if (m && m.frame) return { frame: m.frame, guess: false, cands: [m.frame] };
  if (!boards.length) return { frame: '', guess: false, cands: [] };
  var hm = _aqNorm(q.subject).match(/(?:^|[^A-Z0-9])([A-Z]{1,2})(?:\*+|_+)/), head = hm ? hm[1] : '', set = {};
  frames.forEach(function (o) {
    if (!o.frame || (head && _adMachineKey(o.machine).indexOf(head) !== 0)) return;
    var b = _adFrameBoards(o)[type] || [];
    if (b.length && boards.every(function (x) { return b.indexOf(x) >= 0; })) set[o.frame] = 1;
  });
  var c = Object.keys(set).sort();
  return { frame: c.length === 1 ? c[0] : '', guess: c.length === 1, cands: c };
}
