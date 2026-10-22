// ============================================================
// 47_model_board_set.gs
// 機種ごとの基板構成（主基板／サブ基板／枠制御基板 …）と 価格の自動合計
//
//   ・機種マスタの機種に「役割 × 基板（基板価格表の基板名）× 数量」を登録すると、
//     基板価格表（46_board_price_table.gs）の金額から 売値合計・原価合計・利益・逆ざや を自動で合計する。
//   ・基板価格表の金額が改定されれば、機種の合計も自動で追随する（基板は 基板ID／基板名 で参照）。
//   ・基板価格表に未登録の基板名でも構成は保存でき、「未登録」として合計から外れる（後で同名の基板を登録すれば自動で紐づく）。
//   ・既存の「機種基板構成」シート（主基板/払出基板/D/E/DE）から 構成の候補を取り込める。
//
//   シート: 機種_基板セット
// ============================================================

var MBS_SHEET = '機種_基板セット';
var MBS_HEADERS = ['行ID', '機種コード', '役割', '基板ID', '基板名', '数量', '並び順', '更新日時', '更新者'];
var MBS_NCOLS = 9;
var MBS_ROLE_PRESETS = ['主基板', 'サブ基板', '枠制御基板', '払出基板', '液晶制御基板(D)', '演出基板(E)', '液晶IF基板(DE)'];
var MBS_LEGACY_SHEET = '機種基板構成';   // 30_price_master_import.gs の Excel 取込み表

function _mbsInitSheet() {
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName(MBS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MBS_SHEET);
    var hr = sh.getRange(1, 1, 1, MBS_NCOLS);
    hr.setValues([MBS_HEADERS]);
    hr.setBackground('#EDE7F6');
    hr.setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(2, 8, sh.getMaxRows() - 1, 1).setNumberFormat('@');
  }
  return sh;
}

// ------------------------------------------------------------
// 純粋関数（計算・解析）
// ------------------------------------------------------------
function _mbsBoardIndex(boards) {
  var byId = {}, byName = {};
  boards.forEach(function (b) { byId[b.id] = b; byName[_bpNormName(b.name)] = b; });
  return { byId: byId, byName: byName };
}

function _mbsResolve(line, idx) {
  if (line.boardId && idx.byId[line.boardId]) return idx.byId[line.boardId];
  return idx.byName[_bpNormName(line.boardName)] || null;
}

/** 1行（役割×基板×数量）の金額。基板が未登録・金額未入力なら complete=false（合計の「未確定」に数える） */
function _mbsLineCalc(line, board) {
  var q = Number(line.qty);
  if (!(q > 0)) q = 1;
  var out = {
    id: line.id || '', role: line.role, boardId: board ? board.id : '', boardName: board ? board.name : String(line.boardName || ''),
    qty: q, resolved: !!board, board: null, sale: null, cost: null, profit: null, inverted: false, complete: false,
  };
  if (!board) return out;
  var c = board.calc;
  out.board = { total: c.total, cost: c.cost, profit: c.profit, inverted: c.inverted };
  out.sale = c.total === null ? null : _bpRound(c.total * q, 2);
  out.cost = c.cost === null ? null : _bpRound(c.cost * q, 2);
  out.profit = c.profit === null ? null : _bpRound(c.profit * q, 2);
  out.inverted = out.profit !== null && out.profit < 0;
  out.complete = c.total !== null && c.profit !== null;
  return out;
}

/** 機種の合計。利益は「金額が揃っている基板」だけを合計し、揃っていない基板は incomplete として数える */
function _mbsTotals(calcs) {
  var sale = 0, saleN = 0, cost = 0, costN = 0, profit = 0, profitN = 0, rateBase = 0, costBase = 0, names = [];
  calcs.forEach(function (l) {
    if (l.sale !== null) { sale += l.sale; saleN++; }
    if (l.cost !== null) { cost += l.cost; costN++; }
    if (l.profit !== null) { profit += l.profit; profitN++; rateBase += l.sale; costBase += l.cost; }
    if (!l.complete) names.push(l.boardName || '(未選択)');
  });
  var p = profitN ? _bpRound(profit, 2) : null;
  return {
    sale: saleN ? _bpRound(sale, 2) : null,
    cost: costN ? _bpRound(cost, 2) : null,
    profit: p,
    rate: (p !== null && rateBase > 0) ? _bpRound(p / rateBase * 100, 1) : null,
    headroomRate: (p !== null && costBase > 0) ? _bpRound(p / costBase * 100, 1) : null,   // 原価合計がこの%上がると利益0
    inverted: p !== null && p < 0,
    incomplete: names.length, incompleteNames: names, lineCount: calcs.length,
  };
}

/** 機種コード文字列から「A84」「E64」のような機種記号を取り出す（"P:A84ｻﾌﾞ" → A84） */
function _mbsMachineTokens(s) {
  var t = String(s == null ? '' : s).normalize('NFKC').toUpperCase();
  var re = /(^|[^A-Z0-9])([A-Z]\d{2,3})(?!\d)/g, m, out = [];
  while ((m = re.exec(t)) !== null) { if (out.indexOf(m[2]) < 0) out.push(m[2]); }
  return out;
}

/** セルの基板コードを取り出す。"M2401A(M2105C3)" → [M2401A, M2105C3]、"E1901B/E2101B" → 2件、"なし"/"AG6R" → 0件 */
function _mbsParseBoardCell(s) {
  var t = String(s == null ? '' : s).normalize('NFKC').toUpperCase();
  return (t.match(/[A-Z]{1,4}\d{3,5}[A-Z0-9]*/g) || []).filter(function (x, i, a) { return a.indexOf(x) === i; });
}

// ------------------------------------------------------------
// 読み書き
// ------------------------------------------------------------
function _mbsReadAll() {
  var sh = _mbsInitSheet();
  var last = sh.getLastRow();
  var rows = last > 1 ? sh.getRange(2, 1, last - 1, MBS_NCOLS).getValues() : [];
  return rows.filter(function (r) { return r[0] !== '' && r[1] !== ''; }).map(function (r) {
    return { id: String(r[0]), modelCode: String(r[1]), role: String(r[2] || ''), boardId: String(r[3] || ''), boardName: String(r[4] || ''),
             qty: Number(r[5]) || 1, order: Number(r[6]) || 0, updatedAt: String(r[7] || ''), updater: String(r[8] || '') };
  });
}

function _mbsLinesFor(all, modelCode, idx) {
  var key = String(modelCode).trim();
  var lines = all.filter(function (l) { return l.modelCode === key; }).sort(function (a, b) { return a.order - b.order; });
  var calcs = lines.map(function (l) { return _mbsLineCalc(l, _mbsResolve(l, idx)); });
  return { lines: calcs, totals: _mbsTotals(calcs), updatedAt: lines.length ? lines[lines.length - 1].updatedAt : '', updater: lines.length ? lines[lines.length - 1].updater : '' };
}

/** 既存の「機種基板構成」（Excel取込み表）から この機種の構成候補を作る */
function _mbsCandidates(modelCode, idx) {
  var sh = getSpreadsheet().getSheetByName(MBS_LEGACY_SHEET);
  if (!sh || sh.getLastRow() <= 1) return [];
  var width = Math.max(sh.getLastColumn(), 2);
  var headers = sh.getRange(1, 1, 1, width).getValues()[0];
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();
  var want = _mbsMachineTokens(modelCode);
  if (!want.length) return [];
  var out = [];
  data.forEach(function (r) {
    var code = String(r[0] || '');
    var have = _mbsMachineTokens(code);
    if (!have.some(function (t) { return want.indexOf(t) >= 0; })) return;
    var lines = [], skipped = [];
    for (var c = 1; c < width; c++) {
      var raw = String(r[c] == null ? '' : r[c]).trim();
      if (!raw) continue;
      var tokens = _mbsParseBoardCell(raw);
      if (!tokens.length) { skipped.push({ role: String(headers[c] || ''), raw: raw }); continue; }
      var pick = tokens.filter(function (t) { return idx.byName[_bpNormName(t)]; })[0] || tokens[0];
      lines.push({ role: String(headers[c] || ''), boardName: pick, alts: tokens.filter(function (t) { return t !== pick; }), raw: raw, resolved: !!idx.byName[_bpNormName(pick)], qty: 1 });
    }
    if (lines.length) out.push({ label: code, lines: lines, skipped: skipped });
  });
  return out;
}

// ------------------------------------------------------------
// API
// ------------------------------------------------------------
function apiModelBoardSetGet(payload) {
  try {
    var modelCode = String((payload || {}).modelCode || '').trim();
    if (!modelCode) return { success: false, error: '機種コードが必要です' };
    var idx = _mbsBoardIndex(_bpBoardsLite());
    var res = _mbsLinesFor(_mbsReadAll(), modelCode, idx);
    return { success: true, modelCode: modelCode, lines: res.lines, totals: res.totals, updatedAt: res.updatedAt, updater: res.updater,
             candidates: res.lines.length ? [] : _mbsCandidates(modelCode, idx), rolePresets: MBS_ROLE_PRESETS };
  } catch (e) {
    Logger.log('[apiModelBoardSetGet] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

/** 機種別価格の一覧用：構成を持つ機種すべて＋機種マスタの機種一覧（構成の追加先の選択用） */
function apiModelBoardSetList() {
  try {
    var idx = _mbsBoardIndex(_bpBoardsLite());
    var all = _mbsReadAll();
    var names = {}, allModels = [];
    try {
      _getAllModelMasterRows().forEach(function (r) {
        var code = String(r[MODEL_MASTER_COLS.MODEL_CODE - 1] || '').trim();
        if (!code) return;
        names[code] = String(r[MODEL_MASTER_COLS.MODEL_NAME - 1] || '');
        allModels.push({ modelCode: code, modelName: names[code] });
      });
    } catch (e2) { Logger.log('[apiModelBoardSetList] 機種マスタ取得スキップ: ' + e2.message); }
    var codes = [];
    all.forEach(function (l) { if (codes.indexOf(l.modelCode) < 0) codes.push(l.modelCode); });
    var models = codes.map(function (c) {
      var r = _mbsLinesFor(all, c, idx);
      return { modelCode: c, modelName: names[c] || '', lines: r.lines, totals: r.totals, updatedAt: r.updatedAt, updater: r.updater };
    });
    models.sort(function (a, b) { return a.modelCode.localeCompare(b.modelCode, 'ja'); });
    return { success: true, models: models, allModels: allModels, rolePresets: MBS_ROLE_PRESETS };
  } catch (e) {
    Logger.log('[apiModelBoardSetList] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

/** 機種の基板構成を保存（その機種の構成を丸ごと置き換え。lines が空なら構成を削除） */
function apiModelBoardSetSave(payload) {
  var lock = LockService.getScriptLock();
  try {
    payload = payload || {};
    var modelCode = String(payload.modelCode || '').trim();
    if (!modelCode) return { success: false, error: '機種コードが必要です' };
    var raw = payload.lines || [];
    if (!Array.isArray(raw)) return { success: false, error: '構成の形式が不正です' };
    var idx = _mbsBoardIndex(_bpBoardsLite());
    var now = nowJST(), user = _bpUser();
    var newRows = [];
    for (var i = 0; i < raw.length; i++) {
      var l = raw[i] || {};
      var role = String(l.role || '').trim(), name = String(l.boardName || '').trim();
      if (!role && !name) continue;                                   // 空行は読み飛ばす
      if (!role) return { success: false, error: (i + 1) + '行目: 役割（主基板・サブ基板など）を入力してください' };
      if (!name) return { success: false, error: (i + 1) + '行目: 基板名を入力してください' };
      var qty = (l.qty === '' || l.qty == null) ? 1 : Number(l.qty);
      if (!(qty > 0) || qty > 999) return { success: false, error: (i + 1) + '行目: 数量は 1〜999 で入力してください' };
      var board = _mbsResolve({ boardId: String(l.boardId || ''), boardName: name }, idx);
      newRows.push([_bpId('MS'), modelCode, role, board ? board.id : '', board ? board.name : name, qty, newRows.length + 1, now, user]);
    }

    lock.waitLock(20000);
    var sh = _mbsInitSheet();
    var last = sh.getLastRow();
    var keep = last > 1 ? sh.getRange(2, 1, last - 1, MBS_NCOLS).getValues().filter(function (r) { return r[0] !== '' && String(r[1]) !== modelCode; }) : [];
    if (last > 1) sh.getRange(2, 1, last - 1, MBS_NCOLS).clearContent();
    var all = keep.concat(newRows);
    if (all.length) {
      sh.getRange(2, 8, all.length, 1).setNumberFormat('@');
      sh.getRange(2, 1, all.length, MBS_NCOLS).setValues(all);
    }
    var res = _mbsLinesFor(_mbsReadAll(), modelCode, idx);
    return { success: true, modelCode: modelCode, count: newRows.length, lines: res.lines, totals: res.totals };
  } catch (e) {
    Logger.log('[apiModelBoardSetSave] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}
