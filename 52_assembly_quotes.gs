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
/** 種類：件名から。無ければ基板から推定 */
function _aqType(subject, boards) {
  var s = _aqNorm(subject);
  if (/主制御/.test(s)) return '主制御基板';
  if (/液晶|演出制御/.test(s)) return '液晶演出制御基板';
  if (/枠制御|メダル/.test(s)) return '枠制御基板';
  if (/サブ/.test(s)) return 'サブ制御基板';
  var k = boards.map(_aqBoardKind);
  if (k.indexOf('M') >= 0) return '主制御基板';
  if (k.some(function (x) { return x === 'D' || x === 'DE' || x === 'E' || x === 'SNB'; })) return '液晶演出制御基板';
  if (k.indexOf('C') >= 0) return '枠制御基板';
  return 'その他';
}
function _aqVariant(subject) {
  var s = _aqNorm(subject);
  if (/見本機/.test(s)) return '見本機';
  if (/アフター/.test(s)) return 'アフター用';
  if (/中古|リユース|エコ/.test(s)) return '中古基板使用';
  return '新品';
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
  var groups = {};
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
    g.count = g.quotes.length;
    return g;
  }).sort(function (a, b) {
    var order = ['主制御基板', '液晶演出制御基板', '枠制御基板', 'サブ制御基板', 'その他'];
    return (order.indexOf(a.type) - order.indexOf(b.type)) || a.boardKey.localeCompare(b.boardKey) || a.variant.localeCompare(b.variant);
  });
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
    var needs = [];
    if (M.length) needs.push({ type: '主制御基板', boards: M });
    if (D.length || DE.length || E.length) needs.push({ type: '液晶演出制御基板', boards: D.concat(DE, E) });
    if (C.length) needs.push({ type: '枠制御基板', boards: C });
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
  write(AQ_SHEET_COMP, ['種類', '使用基板', '区分', '最新見積No', '提出日', '金額', '見積書PDF', '見積件数', '件名'],
    comps.map(function (c) { var l = c.latest || {}; return [c.type, c.boards.join(' / '), c.variant, l.quoteNo || '', l.date || '', l.amount || '', l.pdfUrl || '', c.count, l.subject || '']; }));
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
    if (p.progressSsId != null) {
      var m = String(p.progressSsId).match(/[-\w]{25,}/);
      c.progressSsId = m ? m[0] : '';
    }
    PropertiesService.getScriptProperties().setProperty(AQ_PROP, JSON.stringify(c));
    return apiAssemblyQuotes();
  } catch (e) { return { success: false, error: e.message }; }
}
