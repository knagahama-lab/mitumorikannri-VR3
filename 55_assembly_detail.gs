// ============================================================
// 55_assembly_detail.gs
// 組立費 構成一覧（52）の編集機能
//   ① 手修正   … 見積ごとに 種類・使用基板・区分・世代・枠構成・提出日・金額・件名 を上書き（元の見積データは変えない）
//   ② 見積内訳 … 組立費見積書「見積詳細」と同じ項目（材料費／Assy費／組立／検査／梱包／輸送費／小計／管理費／利益／レーザーマーキング／値引き）
//   ③ 枠構成   … ハードウェア構成一覧表（機種別）を取り込み、機種ごとの枠（PFK25・SFK15 等）と基板構成を管理
//   ④ 未来の見積 … これから必要な見積を登録（頭文字ごと）。該当する見積が登録されると「提出済み」になる
//   シート: 組立費_手修正 / 組立費_見積内訳 / 枠構成_機種 / 組立費_未来見積
// ============================================================

var AD_SHEET_OV = '組立費_手修正';
var AD_SHEET_DETAIL = '組立費_見積内訳';
var AD_SHEET_FRAME = '枠構成_機種';
var AD_OV_FIELDS = ['type', 'boards', 'variant', 'gen', 'frame', 'date', 'amount', 'subject', 'initial', 'life'];
var AD_OV_HEADERS = ['見積キー', '種類', '使用基板', '区分', '世代', '枠構成', '提出日', '金額', '件名', '更新者', '更新日時', '頭文字', '状態'];
var AD_OV_COL = { type: 1, boards: 2, variant: 3, gen: 4, frame: 5, date: 6, amount: 7, subject: 8, initial: 11, life: 12 };   // 頭文字・状態は後から追加した列（末尾）
var AD_LIVES = ['現役', '過去', '未来'];
var AD_SHEET_FUTURE = '組立費_未来見積';
var AD_FU_FIELDS = ['id', 'initial', 'type', 'frame', 'boards', 'variant', 'gen', 'subject', 'deadline', 'amount', 'machines', 'memo'];
var AD_FU_HEADERS = ['ID', '頭文字', '種類', '枠構成', '使用基板', '区分', '世代', '件名', '提出期限', '予定金額', '関連機種', 'メモ', '登録者', '登録日', '更新日時'];
// 見積詳細の区分（小計に入る6つ → 管理費 → 利益 → レーザーマーキング → 値引き）
var AD_SUB = ['材料費', 'Assy費', '組立', '検査', '梱包', '輸送費'];
var AD_SECTIONS = AD_SUB.concat(['管理費', 'レーザーマーキング', '値引き']);
var AD_DETAIL_HEADERS = ['見積キー', '見積No'].concat(AD_SUB).concat(['小計', '管理費', '利益率%', '利益', 'レーザーマーキング', '値引き', '合計', '明細JSON', '更新者', '更新日時']);
var AD_FR_FIELDS = ['machine', 'title', 'kind', 'frame', 'M', 'D', 'vdp', 'DE', 'E', 'C', 'memo', 'version'];
var AD_FR_HEADERS = ['機種', 'タイトル', '種別', '枠構成', '主制御基板', '液晶制御基板', 'VDP', '液晶IF基板', '演出IF基板', '枠制御・メダル数制御基板', '備考', '一覧表の版', '手修正した項目', '更新日時'];

// 保存はロックを待たない（スクリプトロックは注文書の自動取込などと共通で、待つと20秒で失敗するため）。
//   ・既にある行は、変わったセルだけを書き換える（行の並びは変えない。削除は行を空にする）
//   ・新しい行は末尾に追加する。同時に追加されて同じキーが2行できても、読む時にまとめる
var _AD_SS = null, _AD_SH = {};
function _adSheet(name, headers) {
  if (_AD_SH[name]) return _AD_SH[name];
  var ss = _AD_SS || (_AD_SS = getSpreadsheet()), sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (sh.getMaxColumns() < headers.length) sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
    sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#FFF3E0');
    sh.setFrozenRows(1);
  } else if (String(sh.getRange(1, headers.length, 1, 1).getValues()[0][0]) !== headers[headers.length - 1]) {
    // 後から増えた列の見出しを足す（既存の列の並びは変えない）
    sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#FFF3E0');
  }
  return (_AD_SH[name] = sh);
}
function _adRead(name, headers) {
  var sh = _adSheet(name, headers), last = sh.getLastRow();
  return last < 2 ? [] : sh.getRange(2, 1, last - 1, headers.length).getValues();
}
function _adStr(v) {
  if (v === null || v === undefined) return '';
  return Object.prototype.toString.call(v) === '[object Date]' ? Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd') : String(v).trim();
}
/** キー（A列）→ 行番号の一覧 */
function _adKeyRows(sh) {
  var last = sh.getLastRow(), out = {};
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) { var k = _adStr(r[0]); if (k) (out[k] = out[k] || []).push(i + 2); });
  return out;
}
function _adWriteRow(sh, at, row) { sh.getRange(at, 1, 1, row.length).setNumberFormat('@').setValues([row]); }
/** 末尾に追加。ロックがすぐ取れればまとめて、取れなければ1行ずつ（appendRow は同時でも重ならない） */
function _adAppend(sh, rows) {
  if (!rows.length) return;
  rows = rows.map(function (r) { return r.map(_adStr); });
  var lock = LockService.getScriptLock(), got = false;
  try { got = lock.tryLock(1500); } catch (e) {}
  if (got) {
    try {
      var at = sh.getLastRow() + 1;
      if (at + rows.length - 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), rows.length + 50);
      sh.getRange(at, 1, rows.length, rows[0].length).setNumberFormat('@').setValues(rows);
      SpreadsheetApp.flush();
    } finally { lock.releaseLock(); }
    return;
  }
  rows.forEach(function (r) {
    sh.appendRow([r[0]]);
    var f = _adKeyRows(sh)[r[0]];
    _adWriteRow(sh, f ? f[f.length - 1] : sh.getLastRow(), r);
  });
}
/** 行を空にする（行は詰めない） */
function _adClear(name, headers, key) {
  var sh = _adSheet(name, headers);
  (_adKeyRows(sh)[String(key)] || []).forEach(function (at) { sh.getRange(at, 1, 1, headers.length).clearContent(); });
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
/** キーの行へ書き込む（無ければ追加） */
function _adPut(name, headers, key, row) {
  var sh = _adSheet(name, headers), f = _adKeyRows(sh)[String(key)];
  row = row.map(_adStr);
  if (f) _adWriteRow(sh, f[f.length - 1], row); else _adAppend(sh, [row]);
}
function _adSafe(fn) {
  try { return fn(); }
  catch (e) { Logger.log('[55] ' + e.message + '\n' + e.stack); return { success: false, error: e.message }; }
}

// ── ① 手修正 ──────────────────────────────────────────────
function _adOverrides() {
  var out = {};
  _adRead(AD_SHEET_OV, AD_OV_HEADERS).forEach(function (r) {
    if (!r[0]) return;
    var k = _adStr(r[0]), o2 = out[k] || (out[k] = {});
    AD_OV_FIELDS.forEach(function (f) { var v = _adStr(r[AD_OV_COL[f]]); if (v !== '') o2[f] = v; });
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
function _adSplitInitials(s) {
  return String(s || '').normalize('NFKC').toUpperCase().split(/[^A-Z]+/).filter(function (x, i, a) { return x && a.indexOf(x) === i; });
}
/** p: { keys:[見積キー], field, value }  value 空 = 自動判定に戻す */
function apiAsmOverrideSave(p) {
  return _adSafe(function () {
    p = p || {};
    var keys = (p.keys || []).map(String).filter(String), fi = AD_OV_FIELDS.indexOf(p.field);
    if (!keys.length || fi < 0) return { success: false, error: '対象が特定できません' };
    var v = String(p.value === null || p.value === undefined ? '' : p.value).trim();
    if (p.field === 'amount' && v !== '') v = String(_adNum(v));
    if (p.field === 'date' && v !== '') { v = _aqDate(v); if (!v) return { success: false, error: '日付が正しくありません' }; }
    if (p.field === 'boards' && v !== '') v = _adSplitBoards(v).join(' / ');
    if (p.field === 'frame' && v !== '') v = _aqSplitFrame(v).join('/');
    if (p.field === 'gen' && v !== '') v = _aqGenJoin(_aqSplitGen(v));
    if (p.field === 'initial' && v !== '') { v = _adSplitInitials(v).join('/'); if (!v) return { success: false, error: '頭文字は英字で入力してください（例: A、NA、A/C）' }; }
    if (p.field === 'life' && v !== '' && AD_LIVES.indexOf(v) < 0) return { success: false, error: '状態は 現役・過去・未来 から選んでください' };
    // 変わったセル（その項目・更新者・更新日時）だけを書く
    var sh = _adSheet(AD_SHEET_OV, AD_OV_HEADERS), at = _adKeyRows(sh), now = nowJST(), user = _adUser(), rows = [], add = [];
    keys.filter(function (k, i) { return keys.indexOf(k) === i; }).forEach(function (k) {
      if (at[k]) { rows = rows.concat(at[k]); return; }
      var r = AD_OV_HEADERS.map(function () { return ''; });
      r[0] = k; r[AD_OV_COL[p.field]] = v; r[9] = user; r[10] = now;
      add.push(r);
    });
    if (rows.length) {
      [[AD_OV_COL[p.field], v], [9, user], [10, now]].forEach(function (x) {
        var c = String.fromCharCode(65 + x[0]);
        sh.getRangeList(rows.map(function (n) { return c + n; })).setNumberFormat('@').setValue(x[1]);
      });
    }
    _adAppend(sh, add);
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
    try { var d = JSON.parse(r[j] || '{}'); out[_adStr(r[0])] = { lines: d.lines || [], rate: d.rate, profit: d.profit, calc: _adCalc(d) }; } catch (e) {}
  });
  return out;
}
/** p: { key, quoteNo, lines, rate, profit, remove } */
function apiAsmDetailSave(p) {
  return _adSafe(function () {
    p = p || {};
    if (!p.key) return { success: false, error: '見積が特定できません' };
    if (p.remove) {
      _adClear(AD_SHEET_DETAIL, AD_DETAIL_HEADERS, p.key);
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
  var seen = {};
  return _adRead(AD_SHEET_FRAME, AD_FR_HEADERS).filter(function (r) { return r[0]; }).map(function (r) {
    var o = {}; AD_FR_FIELDS.forEach(function (f, i) { o[f] = _adStr(r[i]); });
    o.edited = _adStr(r[12]).split(',').filter(String);
    return o;
  }).filter(function (o) { var k = _adMachineKey(o.machine); if (seen[k]) return false; seen[k] = 1; return true; });
}
function _adFrameRow(o) { return AD_FR_FIELDS.map(function (f) { return o[f] || ''; }).concat([(o.edited || []).join(','), nowJST()]); }
/** 一覧表の取込 p: { rows:[{machine,...}], version }。手で直した項目と、より新しい版の値は上書きしない */
function apiFrameImport(p) {
  return _adSafe(function () {
    var rows = (p && p.rows) || [], ver = String((p && p.version) || '');
    if (!rows.length) return { success: false, error: '取り込む機種がありません' };
    var cur = _adFrames(), idx = {}, add = 0, upd = 0, skip = 0, news = [], changed = [];
    cur.forEach(function (o) { idx[_adMachineKey(o.machine)] = o; });
    rows.forEach(function (r) {
      var k = _adMachineKey(r.machine); if (!k) return;
      var o = idx[k];
      if (!o) {
        o = { edited: [] };
        AD_FR_FIELDS.forEach(function (f) { o[f] = r[f] ? String(r[f]) : ''; });
        o.version = ver; idx[k] = o; news.push(o); add++; return;
      }
      if (o.version && ver && o.version > ver) { skip++; return; }
      var ch = false, verCh = ver && o.version !== ver;
      ['title', 'kind', 'frame', 'M', 'D', 'vdp', 'DE', 'E', 'C'].forEach(function (f) {
        if (!r[f] || o.edited.indexOf(f) >= 0 || o[f] === String(r[f])) return;
        o[f] = String(r[f]); ch = true;
      });
      if (ver) o.version = ver;
      if (ch) upd++;
      if ((ch || verCh) && news.indexOf(o) < 0 && changed.indexOf(o) < 0) changed.push(o);
    });
    // 変わった機種の行だけ書き、新しい機種は末尾に追加
    var sh = _adSheet(AD_SHEET_FRAME, AD_FR_HEADERS), at = _adKeyRows(sh);
    changed.forEach(function (o) { var f = at[o.machine]; if (f) _adWriteRow(sh, f[0], _adFrameRow(o).map(_adStr)); });
    _adAppend(sh, news.map(_adFrameRow));
    return { success: true, added: add, updated: upd, skipped: skip, frames: _adFrames() };
  });
}
/** 1項目の編集・機種の追加・削除 p: { machine, field, value } / { add } / { machine, remove } */
function apiFrameSave(p) {
  return _adSafe(function () {
    p = p || {};
    var cur = _adFrames();
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
      _adClear(AD_SHEET_FRAME, AD_FR_HEADERS, o.machine);
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
  var t = ' ' + String(text || '').replace(/[\u2460-\u24FF]/g, ' ').normalize('NFKC').toUpperCase() + ' ', hit = null;
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
  var ini = q.initials || _aqSubjectInitials(q.subject).initials, set = {};
  frames.forEach(function (o) {
    if (!o.frame || (ini.length && ini.indexOf(_aqMachineInitial(o.machine)) < 0)) return;
    var b = _adFrameBoards(o)[type] || [];
    if (b.length && boards.every(function (x) { return b.indexOf(x) >= 0; })) set[o.frame] = 1;
  });
  var c = Object.keys(set).sort();
  return { frame: c.length === 1 ? c[0] : '', guess: c.length === 1, cands: c };
}

// ── ④ 未来の見積（これから必要な見積）────────────────────
function _adFutures() {
  return _adRead(AD_SHEET_FUTURE, AD_FU_HEADERS).filter(function (r) { return r[0]; }).map(function (r) {
    var o = {}; AD_FU_FIELDS.forEach(function (f, i) { o[f] = _adStr(r[i]); });
    o.user = _adStr(r[12]); o.createdAt = _adStr(r[13]);
    return o;
  });
}
/** 見積書の件名（頭文字入り）。例: 【D**】液晶演出制御基板(DE1802A/D2101A/E2503B)(完)D基板中古基板使用 */
function _adSubject(o) {
  var rank = function (b) { var i = ['DE', 'D', 'SNB', 'E', 'M', 'C', 'S'].indexOf(_aqBoardKind(b)); return i < 0 ? 9 : i; };
  var b = _adSplitBoards(o.boards).map(function (x, i) { return { x: x, i: i }; }).sort(function (p, q) { return rank(p.x) - rank(q.x) || p.i - q.i; }).map(function (p) { return p.x; });
  var name = String(o.type || '').replace(/（.*?）/g, '');
  if (name === '主基板') name = '主制御基板';
  var suf = { '全新品': '', 'D基板中古': 'D基板中古基板使用', 'オール中古': 'ALL中古基板使用', '中古基板使用': '中古基板使用', 'リユース（エコ）': 'リユース', '見本機': '見本機', 'アフター用': 'アフター用' }[o.variant];
  return '【' + _adSplitInitials(o.initial).map(function (x) { return x + '**'; }).join('/') + '】' + name + (b.length ? '(' + (b.length === 1 ? 'FJ+' : '') + b.join('/') + ')' : '') + '(完)' + (suf === undefined ? (o.variant || '') : suf);
}
/** 未来の見積を構成の形にする。登録日以降の見積で同じ構成が登録されたら fulfilledBy に見積No */
function _adFutureComps(groups, details) {
  return _adFutures().map(function (o) {
    var boards = _adSplitBoards(o.boards).sort(), bkey = boards.join('+') || '（基板不明）' + o.id, ini = _adSplitInitials(o.initial);
    var kind = _aqKind(o.variant), ch = _aqChannel(o.gen, kind), day = _aqDate(o.createdAt), hit = null;
    Object.keys(groups).forEach(function (k) {
      var g = groups[k];
      if (hit || !boards.length || g.type !== o.type || g.boardKey !== bkey || g.variant !== o.variant || !ini.every(function (x) { return g.initials.indexOf(x) >= 0; })) return;
      if (!_aqFrameOk(o.frame, g.frame) || !_aqGenOk(o.gen, g.gen)) return;
      hit = g.quotes.filter(function (q) { return q.date && day && q.date >= day && q.life !== '過去'; })[0] || null;
    });
    var l = { key: o.id, quoteNo: '', date: '', amount: o.amount ? Number(o.amount) : 0, pdfUrl: '', subject: o.subject || _adSubject(o), life: '未来' };
    return { key: 'F|' + o.id, future: true, futureId: o.id, fulfilledBy: hit ? hit.quoteNo : '', deadline: o.deadline, memo: o.memo, useFor: o.machines,
             type: o.type, boards: boards, boardKey: bkey, variant: o.variant, gen: o.gen, kind: kind, channel: ch.channel, needQuote: ch.needQuote,
             frame: o.frame, frameGuess: false, frameCands: [], initials: ini, machines: [], common: true, noInitial: false, life: '未来', lifeAuto: false,
             edited: {}, quotes: [], latest: l, detail: details[o.id] || null, count: 0 };
  });
}
/** p: { items:[{initial,type,frame,boards,variant,gen,deadline,amount,machines,memo}] } 追加 ／ { id, field, value } 編集 ／ { id, remove } 削除 */
function apiAsmFutureSave(p) {
  return _adSafe(function () {
    p = p || {};
    var clean = function (f, v) {
      v = String(v === null || v === undefined ? '' : v).trim();
      if (v === '') return '';
      if (f === 'initial') return _adSplitInitials(v).join('/');
      if (f === 'boards') return _adSplitBoards(v).join(' / ');
      if (f === 'frame') return _aqSplitFrame(v).join('/');
      if (f === 'gen') return _aqGenJoin(_aqSplitGen(v));
      if (f === 'amount') return String(_adNum(v));
      if (f === 'deadline') { var d = _aqDate(v); if (!d) throw new Error('提出期限の日付が正しくありません'); return d; }
      return v;
    };
    var now = nowJST(), cur = _adFutures();
    if (p.items) {
      if (!p.items.length) return { success: false, error: '追加する見積がありません' };
      var rows = p.items.map(function (it, i) {
        var o = {}; AD_FU_FIELDS.forEach(function (f) { o[f] = clean(f, it[f]); });
        if (!o.initial) throw new Error('頭文字を選んでください');
        if (!o.type) throw new Error('種類を選んでください');
        if (!o.variant) throw new Error('区分を選んでください');
        o.id = 'FQ-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + '-' + (i + 1) + Math.floor(Math.random() * 900 + 100);
        if (!o.subject) o.subject = _adSubject(o);
        var dup = cur.filter(function (x) { return ['initial', 'type', 'frame', 'boards', 'variant', 'gen'].every(function (f) { return x[f] === o[f]; }); })[0];
        return dup ? null : AD_FU_FIELDS.map(function (f) { return o[f]; }).concat([_adUser(), now, now]);
      }).filter(Boolean);
      if (!rows.length) return { success: false, error: '同じ内容の未来の見積が登録済みです' };
      _adAppend(_adSheet(AD_SHEET_FUTURE, AD_FU_HEADERS), rows);
      return { success: true, added: rows.length, skipped: p.items.length - rows.length };
    }
    var o = cur.filter(function (x) { return x.id === String(p.id || ''); })[0];
    if (!o) return { success: false, error: '未来の見積が見つかりません' };
    if (p.remove) {
      _adClear(AD_SHEET_FUTURE, AD_FU_HEADERS, o.id);
      return { success: true };
    }
    var f = p.field === 'date' ? 'deadline' : p.field;
    if (AD_FU_FIELDS.indexOf(f) < 1) return { success: false, error: '編集できない項目です' };
    var autoSubject = o.subject === _adSubject(o);
    o[f] = clean(f, p.value);
    if (autoSubject && f !== 'subject') o.subject = _adSubject(o);   // 件名を手で直していなければ、内容に合わせて作り直す
    _adPut(AD_SHEET_FUTURE, AD_FU_HEADERS, o.id, AD_FU_FIELDS.map(function (x) { return o[x]; }).concat([o.user, o.createdAt, now]));
    return { success: true, value: o[f], subject: o.subject };
  });
}
