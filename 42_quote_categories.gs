// ============================================================
// 42_quote_categories.gs
// 見積カテゴリ（仕掛基板 / PCB / 組立費 / ROM・RAM / 単品部品 …）を管理コンソールから追加・変更・削除
//
//   保存先：スクリプトプロパティ QUOTE_CATEGORY_LIST（全員共通）
//   1件：{ name, icon, palette, keywords[] }
//     palette  … 表示色（indigo / green / amber / purple / sky / rose / teal / slate）
//     keywords … 件名・基板名からカテゴリを自動推定する時の語句（未設定の見積に「推定?」で表示）
//   見積に付けたカテゴリは 見積台帳シートの「見積カテゴリ」列（LEDGER_COLS.CATEGORY）に保存される。
// ============================================================

var QCAT_PROP = 'QUOTE_CATEGORY_LIST';
var QCAT_DEFAULTS = [
  { name: '仕掛基板', icon: '🔩', palette: 'indigo', keywords: [] },
  { name: 'PCB',      icon: '🟩', palette: 'green',  keywords: [] },
  { name: '組立費',   icon: '🛠', palette: 'amber',  keywords: [] },
  { name: 'ROM・RAM', icon: '💾', palette: 'purple', keywords: [] },
  { name: '単品部品', icon: '📦', palette: 'sky',    keywords: [] },
];

function getQuoteCategoryList() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(QCAT_PROP);
    var list = raw ? JSON.parse(raw) : null;
    if (Array.isArray(list) && list.length) return list;
  } catch (e) { Logger.log('[getQuoteCategoryList] ' + e.message); }
  return JSON.parse(JSON.stringify(QCAT_DEFAULTS));
}
function getQuoteCategoryNames() {
  return getQuoteCategoryList().map(function (c) { return c.name; });
}

/** 一覧＋見積台帳での使用件数 */
function apiCategoryList() {
  try {
    var usage = {};
    try {
      getAllLedgerData().forEach(function (r) {
        var c = String(r[LEDGER_COLS.CATEGORY - 1] || '').trim();
        if (c) usage[c] = (usage[c] || 0) + 1;
      });
    } catch (e) {}
    return { success: true, items: getQuoteCategoryList(), usage: usage };
  } catch (e) { return { success: false, error: e.message }; }
}

/**
 * 保存：p.items = [{ name, orig, icon, palette, keywords }]
 *   orig     … 変更前の名前（名前を変えた時は台帳の該当カテゴリも付け替える）
 *   p.clearDeleted = true … 削除したカテゴリが付いている見積を「未分類」に戻す
 */
function apiCategorySave(p) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var before = getQuoteCategoryNames();
    var seen = {};
    var items = ((p && p.items) || []).map(function (c) {
      return { name: String(c.name || '').trim(), orig: String(c.orig || '').trim(), icon: String(c.icon || '').trim() || '🏷',
               palette: String(c.palette || 'slate'),
               keywords: (Array.isArray(c.keywords) ? c.keywords : String(c.keywords || '').split(/[,、\s]+/)).map(function (k) { return String(k).trim(); }).filter(Boolean) };
    });
    for (var i = 0; i < items.length; i++) {
      if (!items[i].name) return { success: false, error: (i + 1) + '行目のカテゴリ名が空です' };
      if (seen[items[i].name]) return { success: false, error: 'カテゴリ名「' + items[i].name + '」が重複しています' };
      if (items[i].name.length > 20) return { success: false, error: 'カテゴリ名は20文字以内にしてください：' + items[i].name };
      seen[items[i].name] = true;
    }
    if (!items.length) return { success: false, error: 'カテゴリを1件以上登録してください' };

    // 名前の変更（orig → name）と削除を見積台帳へ反映
    var renames = {};
    items.forEach(function (c) { if (c.orig && c.orig !== c.name) renames[c.orig] = c.name; });
    var keepNames = items.map(function (c) { return c.orig || c.name; });
    var deleted = before.filter(function (n) { return keepNames.indexOf(n) < 0; });
    var changedRows = 0;
    var sheet = getSpreadsheet().getSheetByName(CONFIG.SHEET_LEDGER);
    if (sheet && sheet.getLastRow() > 1 && (Object.keys(renames).length || (p.clearDeleted && deleted.length))) {
      var rng = sheet.getRange(2, LEDGER_COLS.CATEGORY, sheet.getLastRow() - 1, 1);
      var vals = rng.getValues();
      vals.forEach(function (r) {
        var v = String(r[0] || '').trim();
        if (renames[v]) { r[0] = renames[v]; changedRows++; }
        else if (p.clearDeleted && deleted.indexOf(v) >= 0) { r[0] = ''; changedRows++; }
      });
      if (changedRows) rng.setValues(vals);
    }

    var store = items.map(function (c) { return { name: c.name, icon: c.icon, palette: c.palette, keywords: c.keywords }; });
    PropertiesService.getScriptProperties().setProperty(QCAT_PROP, JSON.stringify(store));
    _applyLedgerCategoryValidation();
    return { success: true, items: store, changedRows: changedRows, deleted: deleted };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

/** 見積台帳シートの「見積カテゴリ」列の入力候補を現在のカテゴリに合わせる */
function _applyLedgerCategoryValidation() {
  try {
    var sheet = getSpreadsheet().getSheetByName(CONFIG.SHEET_LEDGER);
    if (!sheet) return;
    var rows = Math.max(sheet.getMaxRows() - 1, 1);
    sheet.getRange(2, LEDGER_COLS.CATEGORY, rows, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(getQuoteCategoryNames(), true).setAllowInvalid(true).build());
  } catch (e) { Logger.log('[_applyLedgerCategoryValidation] ' + e.message); }
}
