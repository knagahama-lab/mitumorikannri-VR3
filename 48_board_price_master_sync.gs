// ============================================================
// 48_board_price_master_sync.gs
// 基板価格表（46_board_price_table.gs）の金額 → 基板PCB価格マスタ（30_price_master_import.gs）へ反映
//
//   対応（基板価格表 → 基板PCB価格マスタ）
//     PCB原価                         → PCB原価
//     実装費 原価                     → 実装費
//     仕掛原価                        → 仕掛原価(PCB抜き)
//     PCB売価（PCB売値+利益+値引き）  → PCB公表単価
//     合計金額（売価）                → 仕掛公表単価
//     PCB売価 + 合計金額              → 公表単価合計
//   ・基板名とマスタの「基板コード」を照合（先頭の FJ+ は除き、全角半角・大小文字を区別しない）
//   ・基板価格表で空欄の項目はマスタを書き換えない
//   ・マスタに無い基板は最後の行に追加する
//   ・書き込みは apiMasterBoardPriceSaveRow と同じ方法（リンクを保ったまま、元の外部シートへも同期）
//   ・基板価格表の保存時・シート直接編集時に自動で反映。管理用に一括反映 apiBoardPriceSyncMasterAll も用意
// ============================================================

function _bpmCode(s) {
  return String(s == null ? '' : s).normalize('NFKC').toUpperCase().replace(/^\s*FJ\s*[+\-]?\s*/, '').replace(/[\s　]/g, '');
}

/** 基板価格表の値から、マスタへ書く { ヘッダー名: 値 } を作る（空欄は含めない） */
function _bpmUpdates(values, calc) {
  var u = {};
  var put = function (h, x) { if (x !== null && x !== undefined && x !== '') u[h] = x; };
  put('PCB原価', values.pcbPrice);
  put('実装費', values.mountCost);
  put('仕掛原価(PCB抜き)', values.matCost);
  put('PCB公表単価', calc.pcbTotal);
  put('仕掛公表単価', calc.total);
  if (calc.pcbTotal !== null || calc.total !== null) {
    put('公表単価合計', _bpRound((calc.pcbTotal || 0) + (calc.total || 0), 2));
  }
  return u;
}

/** マスタの 基板コード → 行番号[]（同じコードが複数行あれば全部） */
function _bpmIndex(sheet) {
  var idx = {};
  var last = sheet.getLastRow();
  if (last <= 1) return idx;
  var codeCol = MASTER_BOARD_PRICE_HEADERS.indexOf('基板コード') + 1;
  sheet.getRange(2, codeCol, last - 1, 1).getValues().forEach(function (r, i) {
    var k = _bpmCode(r[0]);
    if (k) (idx[k] = idx[k] || []).push(i + 2);
  });
  return idx;
}

/** 1基板をマスタへ反映。戻り値 { updated: 行数, added: true/false, skipped: true/false } */
function _bpSyncToMaster(name, values, calc, sheet, idx) {
  var code = _bpmCode(name);
  if (!code) return { skipped: true };
  var updates = _bpmUpdates(values, calc);
  if (!Object.keys(updates).length) return { skipped: true };
  sheet = sheet || _initMasterBoardPriceSheet();
  idx = idx || _bpmIndex(sheet);
  var rows = idx[code] || [];
  if (rows.length) {
    rows.forEach(function (row) { apiMasterBoardPriceSaveRow({ row: row, updates: updates }); });
    return { updated: rows.length };
  }
  // マスタに無い基板は最後の行に追加
  if (sheet.getLastRow() < 1) sheet.getRange(1, 1, 1, MASTER_BOARD_PRICE_HEADERS.length).setValues([MASTER_BOARD_PRICE_HEADERS]);
  var newRow = new Array(MASTER_BOARD_PRICE_HEADERS.length).fill('');
  newRow[MASTER_BOARD_PRICE_HEADERS.indexOf('基板コード')] = String(name).replace(/^\s*FJ\s*[+\-]?\s*/i, '').trim();
  Object.keys(updates).forEach(function (h) { var c = MASTER_BOARD_PRICE_HEADERS.indexOf(h); if (c >= 0) newRow[c] = updates[h]; });
  var at = sheet.getLastRow() + 1;
  sheet.getRange(at, 1, 1, newRow.length).setValues([newRow]);
  idx[code] = [at];
  return { added: true };
}

/** 基板価格表の1行（シートの配列）を反映（保存・直接編集の後に呼ぶ。失敗しても保存は成功扱い） */
function _bpSyncRowToMaster(name, values) {
  try {
    var r = _bpSyncToMaster(name, values, _bpCalc(values, _bpIncludePcb()));
    return r;
  } catch (e) {
    Logger.log('[BP→MASTER] ' + name + ': ' + e.message);
    return { error: e.message };
  }
}

/** 一括反映：基板価格表の全基板をマスタへ */
function apiBoardPriceSyncMasterAll() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var sheet = _initMasterBoardPriceSheet();
    var idx = _bpmIndex(sheet);
    var updated = 0, added = [], skipped = 0;
    _bpBoardsLite().forEach(function (b) {
      var r = _bpSyncToMaster(b.name, b.values, b.calc, sheet, idx);
      if (r.updated) updated++;
      else if (r.added) added.push(b.name);
      else skipped++;
    });
    return { success: true, updated: updated, added: added, skipped: skipped };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}
