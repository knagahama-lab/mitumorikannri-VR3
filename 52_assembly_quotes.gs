// ============================================================
// 52_assembly_quotes.gs
// 組立費の見積を「構成（基板の組み合わせ）」ごとに一覧化し、進捗管理システムの機種予定と突き合わせる
//
//   ① 構成別 見積一覧：組立費の見積（カテゴリ「組立費」または件名に（完）・組立）を
//        種類（主制御／液晶演出制御／枠制御 など）＋使用基板の組み合わせ＋区分（新品／中古基板使用／アフター用／見本機）
//      でまとめ、最新の提出見積（見積No・提出日・金額・PDF）と過去の件数を出す
//   ② 機種別 予定：進捗管理システムのスプレッドシート（DB_機種・DB_登録機種。毎朝7時に更新）から
//        機種ごとの使用基板（M・D・DE・E・C）と日程を読み、必要な構成ごとに
//        「見積済み（どの見積がいくら）」か「要見積（いつまでに提出）」を判定する
//        提出期限 = 組立予定（量産組立 → 見本機組立 → 部品集約時期 の早いもの）− 準備日数（既定30日）
//   ③ 結果は見積管理のスプレッドシートにも書き出す（シート「組立費_構成別見積」「組立費_機種別予定」）
//      → 進捗管理側やスプレッドシートからも参照できる
// ============================================================

var AQ_PROGRESS_SS_DEFAULT = '18qSYWAIyskR9qfBPC8mGp85XcLRMe4i0WuXXERAJ4-c'; // 進捗管理システムのDBスプレッドシート
var AQ_PROP = 'ASSEMBLY_QUOTE_CFG';
var AQ_SHEET_COMP = '組立費_構成別見積';
var AQ_SHEET_PLAN = '組立費_機種別予定';
var AQ_SHEET_EXTRA = '組立費_追加費用';   // 構成ごとの 解体費・リユース組立費・検査費（手入力）
var AQ_EXTRA_FIELDS = [['dismantle', '解体費'], ['reuseAssy', 'リユース組立費'], ['inspect', '検査費']];

function _aqCfg() {
  var c = {};
  try { c = JSON.parse(PropertiesService.getScriptProperties().getProperty(AQ_PROP) || '{}'); } catch (e) {}
  return { leadDays: c.leadDays != null ? Number(c.leadDays) : 30, progressSsId: c.progressSsId || AQ_PROGRESS_SS_DEFAULT };
}

function _aqNorm(s) {
  return String(s == null ? '' : s).normalize('NFKC').toUpperCase().replace(/FJ\s*[+\-]\s*/g, '').replace(/[\s　]+/g, ' ');
}
/** 文字列から基板コードを抜き出す（M2401A, D2101A, DE2101A1, E2102B, C2501C, S2501D, SNB52163A, D1401C-R 等） */
function _aqBoards(s) {
  var t = _aqNorm(s), out = [], seen = {};
  var re = /(SNB\d{5}[A-Z]?|DE\d{4}[A-Z]\d?|[MDECS]\d{4}[A-Z]\d?)(-R)?/g, m;
  while ((m = re.exec(t))) {
    var prev = m.index > 0 ? t.charAt(m.index - 1) : '';
    if (/[A-Z0-9]/.test(prev)) continue; // 英数字の途中は除外
    var code = m[1]; // 末尾の -R（再生品の表記）は同じ基板として扱う
    if (!seen[code]) { seen[code] = 1; out.push(code); }
  }
  return out;
}
function _aqBoardKind(code) {
  if (/^SNB/.test(code)) return 'SNB';
  if (/^DE/.test(code)) return 'DE';
  return code.charAt(0);
}
// ── 組立の種類・区分（管理画面「⚙ 設定」で追加・変更できる。上から順にキーワードで判定）──
var AQ_DEFAULT_TYPES = [
  { name: '主基板',                     keywords: ['主制御', '主基板'],           boards: ['M'] },
  { name: '液晶演出制御基板',           keywords: ['液晶', '演出制御'],           boards: ['D', 'DE', 'E', 'SNB'] },
  { name: '枠制御基板（パチンコ）',     keywords: ['枠制御'],                     boards: ['C'] },
  { name: 'メダル数制御基板（スロット）', keywords: ['メダル'],                   boards: [] },
  { name: 'サブ制御基板',               keywords: ['サブ'],                       boards: ['S'] },
];
var AQ_DEFAULT_VARIANTS = [
  { name: '見本機',         keywords: ['見本機'] },
  { name: 'アフター用',     keywords: ['アフター'] },
  { name: 'リユース（エコ）', keywords: ['リユース', 'エコ', 'AG6R'] },
  { name: 'D基板中古品',    keywords: ['D基板中古', 'D基板 中古', 'D基板中古基板'] },
  { name: 'オール中古品',   keywords: ['ALL中古', 'オール中古', '全中古', 'ALL 中古'] },
  { name: '中古基板使用',   keywords: ['中古'] },
  { name: '新品',           keywords: [] },   // どれにも当たらない時
];
function _aqTypes() { var c = _aqCfgRaw(); return Array.isArray(c.types) && c.types.length ? c.types : AQ_DEFAULT_TYPES; }
function _aqVariants() { var c = _aqCfgRaw(); return Array.isArray(c.variants) && c.variants.length ? c.variants : AQ_DEFAULT_VARIANTS; }
function _aqCfgRaw() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(AQ_PROP) || '{}'); } catch (e) { return {}; } }
function _aqHasKw(s, kws) { return (kws || []).some(function (k) { k = _aqNorm(k).trim(); return k && s.indexOf(k) >= 0; }); }

/** 種類：件名のキーワード → 無ければ基板の頭文字から */
function _aqType(subject, boards) {
  var s = _aqNorm(subject), types = _aqTypes();
  for (var i = 0; i < types.length; i++) if (_aqHasKw(s, types[i].keywords)) return types[i].name;
  var k = boards.map(_aqBoardKind);
  for (var j = 0; j < types.length; j++) if ((types[j].boards || []).some(function (b) { return k.indexOf(b) >= 0; })) return types[j].name;
  return 'その他';
}
/** 区分：件名のキーワード（上から順）。どれにも当たらなければ キーワード無しの区分（新品） */
function _aqVariant(subject) {
  var s = _aqNorm(subject), vs = _aqVariants(), def = '新品';
  for (var i = 0; i < vs.length; i++) {
    if (!(vs[i].keywords || []).length) { def = vs[i].name; continue; }
    if (_aqHasKw(s, vs[i].keywords)) return vs[i].name;
  }
  return def;
}
/** 機種の使用基板（頭文字）→ 必要な組立の種類 */
function _aqTypeForBoard(kind) {
  var types = _aqTypes();
  for (var i = 0; i < types.length; i++) if ((types[i].boards || []).indexOf(kind) >= 0) return types[i].name;
  return '';
}

function _aqIsAssembly(q) {
  if (String(q.category || '') === '組立費') return true;
  var s = String(q.subject || '').normalize('NFKC');
  return /\(完\)|（完）|組立/.test(s);
}
function _aqDate(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd');
  var s = String(v).normalize('NFKC').trim();
  var m = s.match(/(\d{4})[\/\-年.](\d{1,2})(?:[\/\-月.](\d{1,2}))?/);
  if (!m) return '';
  return m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + (m[3] || '1')).slice(-2);
}
function _aqAddDays(ymd, n) {
  var m = String(ymd).match(/^(\d{4})\/(\d{2})\/(\d{2})$/); if (!m) return '';
  var d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]) + n);
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
}
function _aqDiff(a, b) {
  var p = function (s) { var m = s.match(/^(\d{4})\/(\d{2})\/(\d{2})$/); return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])); };
  return Math.round((p(b) - p(a)) / 86400000);
}

/** ① 構成別の見積一覧 */
function _aqCompositions() {
  var ql = _apiQuoteListGetAll();
  if (!ql.success) throw new Error(ql.error);
  var arch = (typeof ARCHIVE_STATUSES !== 'undefined') ? ARCHIVE_STATUSES : ['ボツ', '旧見積'];
  var groups = {}, extras = _aqExtras();
  ql.items.filter(_aqIsAssembly).forEach(function (q) {
    var boards = _aqBoards(q.subject + ' ' + (q.boardName || ''));
    var type = _aqType(q.subject, boards), variant = _aqVariant(q.subject);
    var bkey = boards.slice().sort().join('+') || '（基板不明）' + q.quoteNo;
    var key = type + '|' + bkey + '|' + variant;
    var g = groups[key] || (groups[key] = { key: key, type: type, boards: boards.slice().sort(), boardKey: bkey, variant: variant, quotes: [] });
    g.quotes.push({ quoteNo: q.quoteNo, date: _aqDate(q.issueDate), amount: q.quoteAmount || 0, pdfUrl: q.quotePdfUrl || '',
                    subject: q.subject || '', client: q.destCompany || '', status: q.status || '', modelCode: q.modelCode || '',
                    archived: arch.indexOf(q.status) >= 0 });
  });
  return Object.keys(groups).map(function (k) {
    var g = groups[k];
    g.quotes.sort(function (a, b) { return String(b.date).localeCompare(String(a.date)) || String(b.quoteNo).localeCompare(String(a.quoteNo)); });
    var live = g.quotes.filter(function (x) { return !x.archived; });
    g.latest = live[0] || g.quotes[0];
    var ex = extras[g.key] || {};
    g.extra = { dismantle: ex.dismantle == null ? null : ex.dismantle, reuseAssy: ex.reuseAssy == null ? null : ex.reuseAssy, inspect: ex.inspect == null ? null : ex.inspect, memo: ex.memo || '' };
    g.count = g.quotes.length;
    return g;
  }).sort(function (a, b) {
    var order = _aqTypes().map(function (t) { return t.name; }).concat(['その他']);
    return (order.indexOf(a.type) - order.indexOf(b.type)) || a.boardKey.localeCompare(b.boardKey) || a.variant.localeCompare(b.variant);
  });
}

function _aqExtraSheet() {
  var ss = getSpreadsheet(), sh = ss.getSheetByName(AQ_SHEET_EXTRA);
  if (!sh) {
    sh = ss.insertSheet(AQ_SHEET_EXTRA);
    sh.getRange(1, 1, 1, 8).setValues([['構成キー', '種類', '使用基板', '区分', '解体費', 'リユース組立費', '検査費', 'メモ']]).setFontWeight('bold').setBackground('#FFF3E0');
    sh.setFrozenRows(1);
  }
  return sh;
}
function _aqExtras() {
  var sh = _aqExtraSheet(), out = {};
  if (sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues().forEach(function (r, i) {
    if (!r[0]) return;
    out[String(r[0])] = { row: i + 2, dismantle: r[4] === '' ? null : Number(r[4]), reuseAssy: r[5] === '' ? null : Number(r[5]), inspect: r[6] === '' ? null : Number(r[6]), memo: String(r[7] || '') };
  });
  return out;
}
/** 追加費用の保存 p: { key, type, boards, variant, dismantle, reuseAssy, inspect, memo } */
function apiAssemblyExtraSave(p) {
  try {
    p = p || {};
    if (!p.key) return { success: false, error: '構成が特定できません' };
    var num = function (v) { if (v === '' || v == null) return ''; var n = Number(String(v).normalize('NFKC').replace(/[,¥￥円\s]/g, '')); if (isNaN(n)) throw new Error('金額は数値で入力してください'); return Math.round(n * 100) / 100; };
    var row = [p.key, p.type || '', p.boards || '', p.variant || '', num(p.dismantle), num(p.reuseAssy), num(p.inspect), p.memo || ''];
    var sh = _aqExtraSheet(), cur = _aqExtras()[p.key];
    if (cur) sh.getRange(cur.row, 1, 1, 8).setValues([row]); else sh.appendRow(row);
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

/** 進捗管理のDBシートを読む（見出し名で列を引く） */
function _aqReadSheet(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var h = v[0].map(function (x) { return String(x).trim(); });
  return v.slice(1).filter(function (r) { return String(r[0]).trim(); }).map(function (r) {
    var o = {}; h.forEach(function (k, i) { if (k) o[k] = r[i]; }); return o;
  });
}

/** ② 機種別の予定と見積状況 */
function _aqPlans(comps, cfg) {
  var ss = SpreadsheetApp.openById(cfg.progressSsId);
  var models = _aqReadSheet(ss, 'DB_機種');
  var reg = {}; _aqReadSheet(ss, 'DB_登録機種').forEach(function (r) { reg[String(r['機種']).trim()] = r; });
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');

  // 構成の索引：種類|基板の組み合わせ → 構成（区分ごと）
  var idx = {};
  comps.forEach(function (c) { var k = c.type + '|' + c.boardKey; (idx[k] = idx[k] || []).push(c); });
  // 基板1枚でも当たる構成（主制御・枠制御は1枚、液晶は全一致を優先）
  var findComps = function (type, boards) {
    var k = type + '|' + boards.slice().sort().join('+');
    if (idx[k]) return { exact: true, comps: idx[k] };
    var part = comps.filter(function (c) { return c.type === type && boards.length && boards.every(function (b) { return c.boards.indexOf(b) >= 0; }); });
    return { exact: false, comps: part };
  };

  var out = [];
  models.forEach(function (m) {
    var id = String(m['機種']).trim();
    var r = reg[id] || {};
    var col = function (k) { return _aqBoards(m[k] || ''); };
    var M = col('M'), D = col('D'), DE = col('DE'), E = col('E'), C = col('C');
    var needs = [], byType = {};
    [['M', M], ['D', D], ['DE', DE], ['E', E], ['C', C]].forEach(function (x) {
      if (!x[1].length) return;
      var t = _aqTypeForBoard(x[0]); if (!t) return;
      (byType[t] = byType[t] || []).push.apply(byType[t], x[1]);
    });
    Object.keys(byType).forEach(function (t) { needs.push({ type: t, boards: byType[t] }); });
    if (!needs.length) return;

    var dates = [
      ['量産組立', _aqDate(r['量産組立'])], ['見本機組立', _aqDate(r['見本機組立'])], ['部品集約時期', _aqDate(m['部品集約時期'])],
    ].filter(function (x) { return x[1]; }).sort(function (a, b) { return a[1].localeCompare(b[1]); });
    var base = dates[0] || null;
    var sales = _aqDate(m['販売時期']);
    // 終わった機種（販売時期・日程が全部90日以上前）は除外
    var last = [sales].concat(dates.map(function (x) { return x[1]; })).filter(String).sort().pop();
    if (last && _aqDiff(last, today) > 90) return;

    needs.forEach(function (n) {
      var f = findComps(n.type, n.boards);
      var quoted = f.comps.map(function (c) {
        return { variant: c.variant, quoteNo: c.latest.quoteNo, date: c.latest.date, amount: c.latest.amount, pdfUrl: c.latest.pdfUrl, boards: c.boards.join(' / ') };
      });
      var deadline = base ? _aqAddDays(base[1], -cfg.leadDays) : '';
      out.push({
        machine: id, customer: String(m['得意先'] || ''), title: String(m['タイトル'] || ''), sales: sales,
        type: n.type, boards: n.boards.join(' / '),
        status: quoted.length ? (f.exact ? '見積済み' : '一部一致') : '要見積',
        quotes: quoted,
        baseLabel: base ? base[0] : '', baseDate: base ? base[1] : '',
        deadline: deadline, leftDays: deadline ? _aqDiff(today, deadline) : null,
        progressStatus: String(r['ステータス'] || ''),
      });
    });
  });
  out.sort(function (a, b) {
    var sa = a.status === '要見積' ? 0 : 1, sb = b.status === '要見積' ? 0 : 1;
    return (sa - sb) || String(a.deadline || '9999').localeCompare(String(b.deadline || '9999')) || a.machine.localeCompare(b.machine);
  });
  return out;
}

function _aqWriteSheets(comps, plans) {
  var ss = getSpreadsheet(), now = nowJST();
  var write = function (name, headers, rows) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.clearContents();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#FFF3E0');
    if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sh.setFrozenRows(1);
    sh.getRange(1, headers.length + 2).setValue('最終更新: ' + now);
  };
  write(AQ_SHEET_COMP, ['種類', '使用基板', '区分', '最新見積No', '提出日', '金額', '解体費', 'リユース組立費', '検査費', '見積書PDF', '見積件数', '件名'],
    comps.map(function (c) { var l = c.latest || {}; var x = c.extra || {}; return [c.type, c.boards.join(' / '), c.variant, l.quoteNo || '', l.date || '', l.amount || '', x.dismantle == null ? '' : x.dismantle, x.reuseAssy == null ? '' : x.reuseAssy, x.inspect == null ? '' : x.inspect, l.pdfUrl || '', c.count, l.subject || '']; }));
  write(AQ_SHEET_PLAN, ['機種', '得意先', 'タイトル', '販売時期', '種類', '使用基板', '見積状況', '見積（区分:見積No 金額 提出日）', '基準日程', '基準日', '提出期限', '残り日数', '進捗ステータス'],
    plans.map(function (p) {
      return [p.machine, p.customer, p.title, p.sales, p.type, p.boards, p.status,
        p.quotes.map(function (q) { return q.variant + ':' + q.quoteNo + ' ¥' + Number(q.amount || 0).toLocaleString() + ' ' + q.date; }).join('\n'),
        p.baseLabel, p.baseDate, p.deadline, p.leftDays == null ? '' : p.leftDays, p.progressStatus];
    }));
}

/** API：構成別見積＋機種別予定（進捗管理が読めない時は①だけ返す） */
function apiAssemblyQuotes(p) {
  try {
    var cfg = _aqCfg();
    var comps = _aqCompositions();
    var plans = [], planError = '';
    try { plans = _aqPlans(comps, cfg); } catch (e) { planError = '進捗管理システムのスプレッドシートを読めませんでした：' + e.message; }
    try { _aqWriteSheets(comps, plans); } catch (e2) { Logger.log('[apiAssemblyQuotes write] ' + e2.message); }
    cfg.types = _aqTypes(); cfg.variants = _aqVariants();
    return { success: true, comps: comps, plans: plans, planError: planError, cfg: cfg, generatedAt: nowJST() };
  } catch (e) {
    Logger.log('[apiAssemblyQuotes] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

/** 設定 p: { leadDays, progressSsId } */
function apiAssemblyQuotesCfg(p) {
  try {
    p = p || {};
    var c = {}; try { c = JSON.parse(PropertiesService.getScriptProperties().getProperty(AQ_PROP) || '{}'); } catch (e) {}
    if (p.leadDays != null) c.leadDays = Math.max(0, Math.min(365, Number(p.leadDays) || 0));
    var clean = function (list, withBoards) {
      return (list || []).map(function (t) {
        var o = { name: String(t.name || '').trim(), keywords: (Array.isArray(t.keywords) ? t.keywords : String(t.keywords || '').split(/[,、\n]+/)).map(function (k) { return String(k).trim(); }).filter(Boolean) };
        if (withBoards) o.boards = (Array.isArray(t.boards) ? t.boards : String(t.boards || '').split(/[,、\s]+/)).map(function (k) { return String(k).trim().toUpperCase(); }).filter(Boolean);
        return o;
      }).filter(function (t) { return t.name; });
    };
    if (p.types) c.types = clean(p.types, true);
    if (p.variants) c.variants = clean(p.variants, false);
    if (p.resetLists) { delete c.types; delete c.variants; }
    if (p.progressSsId != null) {
      var m = String(p.progressSsId).match(/[-\w]{25,}/);
      c.progressSsId = m ? m[0] : '';
    }
    PropertiesService.getScriptProperties().setProperty(AQ_PROP, JSON.stringify(c));
    return apiAssemblyQuotes();
  } catch (e) { return { success: false, error: e.message }; }
}
