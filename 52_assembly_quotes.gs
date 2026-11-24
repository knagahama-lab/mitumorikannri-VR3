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
//   ③ 見積の要否は「基板世代（AG5／AG6R）×新品か中古再生か」で決まる商流で判定する
//        AG6R の中古再生 … サン電子商流（弊社に注文書が来る）→ 見積書の提出が必要
//        AG5  の中古再生 … イードリーム商流（弊社に注文書が来ない）→ 見積書の作成は不要
//        新品・見本機・アフター用 … 世代に関係なくサン電子商流（要見積）
//   ④ 結果は見積管理のスプレッドシートにも書き出す（シート「組立費_構成別見積」「組立費_機種別予定」）
//      → 進捗管理側やスプレッドシートからも参照できる
// ============================================================

var AQ_PROGRESS_SS_DEFAULT = '18qSYWAIyskR9qfBPC8mGp85XcLRMe4i0WuXXERAJ4-c'; // 進捗管理システムのDBスプレッドシート
var AQ_PROP = 'ASSEMBLY_QUOTE_CFG';
var AQ_SHEET_COMP = '組立費_構成別見積';
var AQ_SHEET_PLAN = '組立費_機種別予定';

// ── 基板世代と商流（見積の要否）。管理画面「⚙ 設定」で変更できる ──────────────
//   中古再生のときだけ、基板世代で商流が変わる：
//     AG6R → サン電子商流（弊社に注文書が来る）→ 見積「要」
//     AG5  → イードリーム商流（弊社に注文書が来ない）→ 見積「不要」
//   planReuse: 進捗管理の使用基板セルにこの世代名が入っていたら中古再生とみなす
//     （例：D列が「AG6R」＝新品の D／SNB 基板ではなく AG6R ユニットの再生を使う、の意味）
var AQ_DEFAULT_GENS = [
  { name: 'AG6R', keywords: ['AG6R'], planReuse: true,  channel: 'サン電子',     needQuote: '要' },
  { name: 'AG5',  keywords: ['AG5'],  planReuse: false, channel: 'イードリーム', needQuote: '不要' },
];
// 中古再生を示す言葉（「再販」と区別するため『再生』は2文字で持つ）
var AQ_DEFAULT_REUSE_KW = ['中古', '再生', 'リユース', 'エコ'];
var AQ_KIND_REUSE = '中古再生';
var AQ_KIND_NEW = '新品';
var AQ_GEN_UNKNOWN = '（世代不明）';

function _aqCfg() {
  var c = {};
  try { c = JSON.parse(PropertiesService.getScriptProperties().getProperty(AQ_PROP) || '{}'); } catch (e) {}
  return { leadDays: c.leadDays != null ? Number(c.leadDays) : 30, progressSsId: c.progressSsId || AQ_PROGRESS_SS_DEFAULT };
}

function _aqNorm(s) {
  return String(s == null ? '' : s).replace(/[\u2460-\u24FF]/g, ' ').normalize('NFKC').toUpperCase().replace(/FJ\s*[+\-]\s*/g, '').replace(/[\s　]+/g, ' ');
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
  { name: 'D基板中古',      keywords: ['D基板中古', 'D基板 中古', 'Dリユース', 'D-R', 'D中古'] },
  { name: 'オール中古',     keywords: ['ALL中古', 'オール中古', '全中古', 'ALL 中古'] },
  { name: 'リユース（エコ）', keywords: ['リユース', 'エコ', 'AG6R'] },
  { name: '中古基板使用',   keywords: ['中古'] },
  { name: '全新品',         keywords: [] },   // どれにも当たらない時
];
// 旧名 → 新名（設定画面で保存済みの一覧も読み替える）
var AQ_VARIANT_ALIAS = { '新品': '全新品', 'D基板中古品': 'D基板中古', 'オール中古品': 'オール中古' };
var AQ_MAIN_VARIANTS = ['全新品', 'D基板中古', 'オール中古'];   // 区分別価格の主要3区分
function _aqTypes() { var c = _aqCfgRaw(); return Array.isArray(c.types) && c.types.length ? c.types : AQ_DEFAULT_TYPES; }
function _aqVariants() {
  var c = _aqCfgRaw();
  if (!Array.isArray(c.variants) || !c.variants.length) return AQ_DEFAULT_VARIANTS;
  var list = c.variants.map(function (v) {
    var n = AQ_VARIANT_ALIAS[v.name] || v.name, def = AQ_DEFAULT_VARIANTS.filter(function (d) { return d.name === n; })[0];
    // 旧名で保存されていた区分には、新しいキーワード（Dリユース・D-R など）も足す
    var kw = (v.keywords || []).slice();
    if (AQ_VARIANT_ALIAS[v.name] && def) def.keywords.forEach(function (k) { if (kw.indexOf(k) < 0) kw.push(k); });
    return { name: n, keywords: kw };
  });
  // 「Dリユース」がエコ／リユースに取られないよう、D基板中古を先に判定する
  var di = -1, ri = -1;
  list.forEach(function (v, i) { if (v.name === 'D基板中古') di = i; if (v.name === 'リユース（エコ）') ri = i; });
  if (di > ri && ri >= 0) { var d = list.splice(di, 1)[0]; list.splice(ri, 0, d); }
  return list;
}
function _aqCfgRaw() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(AQ_PROP) || '{}'); } catch (e) { return {}; } }
function _aqHasKw(s, kws) { return (kws || []).some(function (k) { k = _aqNorm(k).trim(); return k && s.indexOf(k) >= 0; }); }
function _aqGens() { var c = _aqCfgRaw(); return Array.isArray(c.gens) && c.gens.length ? c.gens : AQ_DEFAULT_GENS; }
function _aqReuseKw() { var c = _aqCfgRaw(); return Array.isArray(c.reuseKeywords) && c.reuseKeywords.length ? c.reuseKeywords : AQ_DEFAULT_REUSE_KW; }

/** 基板世代（AG6R／AG5）を文字列から判定。上から順に判定し、見つからなければ '' */
function _aqGen(text) {
  var s = _aqNorm(text), gens = _aqGens();
  for (var i = 0; i < gens.length; i++) if (_aqHasKw(s, gens[i].keywords)) return gens[i].name;
  return '';
}
function _aqGenCfg(name) {
  var gens = _aqGens();
  for (var i = 0; i < gens.length; i++) if (gens[i].name === name) return gens[i];
  return null;
}
/** 新品か中古再生か */
function _aqKind(text) { return _aqHasKw(_aqNorm(text), _aqReuseKw()) ? AQ_KIND_REUSE : AQ_KIND_NEW; }
/** 商流と見積の要否。中古再生のときだけ世代で分かれる。世代が取れない中古再生は安全側で「要」 */
function _aqChannel(gen, kind) {
  if (kind !== AQ_KIND_REUSE) return { channel: 'サン電子', needQuote: '要' };
  var g = _aqGenCfg(gen);
  if (!g) return { channel: AQ_GEN_UNKNOWN, needQuote: '要' };
  return { channel: g.channel || 'サン電子', needQuote: String(g.needQuote || '要').indexOf('不要') >= 0 ? '不要' : '要' };
}

// ── 機種の頭文字（組立費の見積書は頭文字ごとに必要。件名に頭文字が無い見積は使えない）──
var AQ_DEFAULT_INITIALS = ['A', 'C', 'D', 'E', 'NA'];
function _aqInitials() { var c = _aqCfgRaw(); return Array.isArray(c.initials) && c.initials.length ? c.initials : AQ_DEFAULT_INITIALS; }
function _aqInitAlt() {
  return _aqInitials().slice().sort(function (a, b) { return b.length - a.length; })
    .map(function (x) { return x.replace(/[^A-Z0-9]/g, ''); }).filter(String).join('|');
}
/** 機種コードの頭文字（D61→D、NA09→NA）。登録されている頭文字でなければ '' */
function _aqMachineInitial(code) {
  var m = String(code || '').normalize('NFKC').toUpperCase().trim().match(/^([A-Z]+)\d/);
  return m && _aqInitials().indexOf(m[1]) >= 0 ? m[1] : '';
}
/**
 * 件名の頭文字：【A**】【E_】機種【C**】A__ ＝ 頭文字の機種すべてに使える共通見積（common）
 *               A85・D60B・NA07 のような機種コードだけ ＝ その機種専用の見積（machines）
 */
function _aqSubjectInitials(subject) {
  var s = String(subject || '').replace(/[\u2460-\u24FF]/g, ' ').normalize('NFKC').toUpperCase().replace(/[○◯●]/g, '*'), alt = _aqInitAlt(), out = [], machines = [], m;
  var add = function (x) { if (x && out.indexOf(x) < 0) out.push(x); };
  var re1 = new RegExp('(?:^|[^A-Z0-9])(' + alt + ')\\s*(?:\\*+|_+|-{2,})(?![A-Z0-9])', 'g');
  while ((m = re1.exec(s))) add(m[1]);
  var re2 = new RegExp('【\\s*(?:機種)?\\s*(' + alt + ')\\s*】|機種\\s*(' + alt + ')(?![A-Z0-9*_])', 'g');
  while ((m = re2.exec(s))) add(m[1] || m[2]);
  var common = out.length > 0;
  if (!common) {
    var re3 = new RegExp('(?:^|[^A-Z0-9])((' + alt + ')\\d{2,3}[A-Z]{0,2})(?![A-Z0-9])', 'g');
    while ((m = re3.exec(s))) { add(m[2]); if (machines.indexOf(m[1]) < 0) machines.push(m[1]); }
  }
  return { initials: out, machines: machines, common: common };
}

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
  var s = _aqNorm(subject), vs = _aqVariants(), def = '全新品';
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
  var groups = {}, ovs = _adOverrides(), details = _adDetails(), frames = _adFrames();
  ql.items.filter(function (q) { return _aqIsAssembly(q) || ovs[q.quoteNo || q.id]; }).forEach(function (q) {
    var qkey = String(q.quoteNo || q.id), ov = ovs[qkey] || {};
    var subject = ov.subject || q.subject || '';
    var boards = ov.boards ? _adSplitBoards(ov.boards) : _aqBoards(subject + ' ' + (q.boardName || ''));
    var type = ov.type || _aqType(subject, boards);
    var variant = ov.variant || _aqVariant(subject);
    variant = AQ_VARIANT_ALIAS[variant] || variant;
    var gen = ov.gen || _aqGen([subject, q.boardName || '', q.modelCode || ''].join(' '));
    // 区分を手で直した時は区分だけで 新品／中古再生 を決める
    var kind = _aqKind(ov.variant ? variant : [variant, subject].join(' '));
    var ch = _aqChannel(gen, kind);
    var si = _aqSubjectInitials(subject);
    var ini = ov.initial ? { initials: _adSplitInitials(ov.initial), machines: [], common: true } : si;
    var archived = arch.indexOf(q.status) >= 0;
    var life = ov.life || (archived ? '過去' : '現役');
    var fr = ov.frame ? { frame: ov.frame, guess: false, cands: [] } : _adGuessFrame({ subject: subject, modelCode: q.modelCode, initials: ini.initials }, type, boards, frames);
    var bkey = boards.slice().sort().join('+') || '（基板不明）' + qkey;
    // 頭文字（機種専用なら機種）が違えば別の見積。「未来」の見積は現役の見積の履歴に隠れないよう分ける
    var key = [type, bkey, variant, gen, fr.frame, ini.initials.slice().sort().join('/'), ini.machines.slice().sort().join('/'), life === '未来' ? '未来' : ''].join('|');
    var g = groups[key] || (groups[key] = { key: key, type: type, boards: boards.slice().sort(), boardKey: bkey, variant: variant,
                                            gen: gen, kind: kind, channel: ch.channel, needQuote: ch.needQuote,
                                            frame: fr.frame, frameGuess: fr.guess, frameCands: fr.cands, initials: ini.initials, machines: ini.machines, common: ini.common,
                                            edited: {}, quotes: [] });
    Object.keys(ov).forEach(function (f) { g.edited[f] = true; });
    g.quotes.push({ key: qkey, quoteNo: q.quoteNo, date: ov.date || _aqDate(q.issueDate), amount: ov.amount !== undefined ? Number(ov.amount) : (q.quoteAmount || 0), pdfUrl: q.quotePdfUrl || '',
                    subject: subject, client: q.destCompany || '', status: q.status || '', modelCode: q.modelCode || '',
                    archived: archived, life: life, lifeSet: !!ov.life, noInitial: !si.initials.length });
  });
  return Object.keys(groups).map(function (k) {
    var g = groups[k];
    g.quotes.sort(function (a, b) { return String(b.date).localeCompare(String(a.date)) || String(b.quoteNo).localeCompare(String(a.quoteNo)); });
    var live = g.quotes.filter(function (x) { return x.life !== '過去'; });
    g.latest = live[0] || g.quotes[0];
    g.life = g.latest.life; g.lifeAuto = !g.latest.lifeSet; g.noInitial = g.latest.noInitial;
    g.detail = details[g.latest.key] || null;
    g.count = g.quotes.length;
    return g;
  }).concat(_adFutureComps(groups, details)).sort(function (a, b) {
    var order = _aqTypes().map(function (t) { return t.name; }).concat(['その他']);
    return (order.indexOf(a.type) - order.indexOf(b.type)) || String(a.frame || '~').localeCompare(String(b.frame || '~')) || a.boardKey.localeCompare(b.boardKey)
      || a.initials.join('/').localeCompare(b.initials.join('/')) || String(a.gen || '').localeCompare(String(b.gen || '')) || a.variant.localeCompare(b.variant);
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
function _aqPlans(allComps, cfg, stat) {
  // 過去の見積は使わない。まだ提出していない「未来」の登録は「作成予定」の判定にだけ使う
  var comps = allComps.filter(function (c) { return !c.future && c.life !== '過去'; });
  var futures = allComps.filter(function (c) { return c.future && !c.fulfilledBy; });
  stat = stat || {};
  var ss = SpreadsheetApp.openById(cfg.progressSsId);
  var models = _aqReadSheet(ss, 'DB_機種');
  stat.models = models.length; stat.noBoards = 0; stat.ended = 0; stat.noSheet = !ss.getSheetByName('DB_機種');
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

  var out = [], frames = _adFrames();
  models.forEach(function (m) {
    var id = String(m['機種']).trim();
    var r = reg[id] || {};
    var fm = frames.filter(function (o) { return _adMachineKey(o.machine) === _adMachineKey(id); })[0] || {};
    // 使用基板：生産計画の欄 → 空なら進捗管理のHW構成一覧表の欄 → 空なら枠構成マスタ
    var HW = { M: 'HW:主制御基板', D: 'HW:液晶制御基板', DE: 'HW:液晶IF基板', E: 'HW:演出IF基板' };
    var raw = function (k) { return String(m[k] || (HW[k] && m[HW[k]]) || fm[k] || ''); };
    var col = function (k) { return _aqBoards(raw(k)); };
    var needs = [], byType = {}, rawByType = {};
    [['M', col('M')], ['D', col('D')], ['DE', col('DE')], ['E', col('E')], ['C', col('C')]].forEach(function (x) {
      var t = _aqTypeForBoard(x[0]); if (!t) return;
      var cell = raw(x[0]);
      // 基板品番が入っていなくても、セルに世代名（D列の「AG6R」など）があれば構成として数える
      if (!x[1].length && !_aqGen(cell)) return;
      byType[t] = byType[t] || [];
      if (x[1].length) byType[t].push.apply(byType[t], x[1]);
      rawByType[t] = (rawByType[t] || '') + ' ' + cell;
    });
    Object.keys(byType).forEach(function (t) { needs.push({ type: t, boards: byType[t], raw: rawByType[t] || '' }); });
    if (!needs.length) { stat.noBoards++; return; }

    var dates = [
      ['量産組立', _aqDate(r['量産組立'])], ['見本機組立', _aqDate(r['見本機組立'])], ['部品集約時期', _aqDate(m['部品集約時期'])],
    ].filter(function (x) { return x[1]; }).sort(function (a, b) { return a[1].localeCompare(b[1]); });
    var base = dates[0] || null;
    var sales = _aqDate(m['販売時期']);
    // 終わった機種（販売時期・日程が全部90日以上前）は除外
    var last = [sales].concat(dates.map(function (x) { return x[1]; })).filter(String).sort().pop();
    if (last && _aqDiff(last, today) > 90) { stat.ended++; return; }

    var mName = [id, String(m['タイトル'] || '')].join(' ');
    var frame = fm.frame || String(m['長納期:枠'] || '');
    // 頭文字ごとの見積が要るのは藤商事の機種（A・C・D・E・NA など、設定にある頭文字）
    var ini = !m['得意先'] || String(m['得意先']).indexOf('藤商事') >= 0 ? _aqMachineInitial(id) : '';
    needs.forEach(function (n) {
      // 世代は 構成セル（D列の「AG6R」など）→ 機種名 の順で判定
      var genCell = _aqGen(n.raw), gen = genCell || _aqGen(mName) || _aqGen(m['HW:VDP'] || fm.vdp || ''), genGuess = false, gc = _aqGenCfg(gen);
      // 必要区分：中古再生の言葉があるか、構成セルの世代名が「構成表の表記で中古再生」設定のとき
      var kind = (_aqKind(n.raw + ' ' + mName) === AQ_KIND_REUSE || (gc && gc.planReuse && genCell)) ? AQ_KIND_REUSE : AQ_KIND_NEW;
      var f = findComps(n.type, n.boards);
      // 構成表・機種名に世代の表記が無いときは、基板が完全一致する見積の世代を借りる（推定）
      if (!gen && f.exact) {
        var withGen = f.comps.filter(function (c) { return c.gen; });
        if (withGen.length) { gen = withGen[0].gen; genGuess = true; }
      }
      var ch = _aqChannel(gen, kind);
      // 必要区分（と世代）に合う見積だけを「この構成の見積」とみなす。世代が空の見積は世代不明として許容
      var quoted = f.comps.map(function (c) {
        // 違う点。1つでもあれば、その見積は流用できない
        var diff = [];
        if (ini) {
          if (!c.initials.length) diff.push('頭文字なし');
          else if (c.initials.indexOf(ini) < 0) diff.push('頭文字');
          else if (!c.common && c.machines.length && c.machines.indexOf(_adMachineKey(id)) < 0) diff.push('他機種専用');
        }
        if ((c.kind || AQ_KIND_NEW) !== kind) diff.push('区分');
        if (gen && c.gen && c.gen !== gen) diff.push('世代');
        if (frame && c.frame && c.frame !== frame) diff.push('枠');
        return { variant: c.variant, gen: c.gen || '', frame: c.frame || '', kind: c.kind || AQ_KIND_NEW, fit: !diff.length, diff: diff,
                 initials: c.initials, life: c.life,
                 quoteNo: c.latest.quoteNo, date: c.latest.date, amount: c.latest.amount, pdfUrl: c.latest.pdfUrl, boards: c.boards.join(' / ') };
      });
      // 中古再生は使用基板が入れ替わる（AG6Rユニットが新品の D／SNB 基板の代わりになる）ので、
      // 基板の完全一致を待たず「世代と区分が合う見積」があれば見積済みとみなす
      var fitCount = quoted.filter(function (x) { return x.fit && (f.exact || (gen && x.gen === gen)); }).length;
      var bset = n.boards.slice().sort().join('+');
      var planned = futures.filter(function (c) {
        return c.type === n.type && c.boardKey === bset && c.kind === kind && (!ini || c.initials.indexOf(ini) >= 0)
          && (!gen || !c.gen || c.gen === gen) && (!frame || !c.frame || c.frame === frame);
      });
      var only = function (list) { return quoted.some(function (x) { return x.diff.length && x.diff.every(function (d) { return list.indexOf(d) >= 0; }); }); };
      var status = ch.needQuote === '不要' ? '見積不要'
        : fitCount ? '見積済み'
        : planned.length ? '作成予定'
        : !quoted.length ? '要見積'
        : !f.exact ? '一部一致'
        : only(['頭文字', '頭文字なし', '他機種専用']) ? '頭文字違い'
        : only(['枠']) ? '枠違い'
        : '区分違い';   // 構成は一致するが、必要な区分（新品／中古再生）の見積が無い＝単価が違うので流用できない
      var deadline = base ? _aqAddDays(base[1], -cfg.leadDays) : '';
      out.push({
        initial: ini, boardKey: bset, planned: planned.map(function (c) { return c.futureId; }),
        frame: frame, machine: id, customer: String(m['得意先'] || ''), title: String(m['タイトル'] || ''), sales: sales,
        type: n.type, boards: n.boards.join(' / '),
        gen: gen, genGuess: genGuess, kind: kind, channel: ch.channel, needQuote: ch.needQuote,
        status: status,
        quotes: quoted,
        baseLabel: base ? base[0] : '', baseDate: base ? base[1] : '',
        deadline: deadline, leftDays: deadline ? _aqDiff(today, deadline) : null,
        progressStatus: String(r['ステータス'] || ''),
      });
    });
  });
  // 対応が要るものを上に、提出期限の近い順
  var rank = function (s) { var i = ['要見積', '頭文字違い', '枠違い', '区分違い', '一部一致', '作成予定', '見積済み'].indexOf(s); return i < 0 ? 9 : i; };
  out.sort(function (a, b) {
    return (rank(a.status) - rank(b.status)) || String(a.deadline || '9999').localeCompare(String(b.deadline || '9999')) || a.machine.localeCompare(b.machine);
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
  write(AQ_SHEET_COMP, ['状態', '頭文字', '種類', '枠構成', '使用基板', '区分', '世代', '新品/中古', '商流', '最新見積No', '提出日', '金額'].concat(AD_SUB).concat(['小計', '管理費', '利益', 'レーザーマーキング', '値引き', '内訳合計', '見積書PDF', '見積件数', '件名']),
    comps.map(function (c) {
      var l = c.latest || {}, x = c.detail && c.detail.calc, v = function (k) { return x ? x.sums[k] : ''; };
      return [c.life || '', c.initials.join('/') + (c.machines.length ? '（' + c.machines.join('/') + '専用）' : ''), c.type, (c.frame || '') + (c.frameGuess ? '（推定）' : ''), c.boards.join(' / '), c.variant, c.gen || '', c.kind || '', c.channel || '', l.quoteNo || '', l.date || '', l.amount || '']
        .concat(AD_SUB.map(v)).concat([x ? x.sub : '', v('管理費'), x ? x.profit : '', v('レーザーマーキング'), v('値引き'), x ? x.total : '', l.pdfUrl || '', c.count, l.subject || '']);
    }));
  write(AQ_SHEET_PLAN, ['機種', '頭文字', '得意先', 'タイトル', '販売時期', '枠構成', '種類', '使用基板', '世代', '必要区分', '商流', '見積要否', '見積状況', '見積（区分:見積No 金額 提出日）', '基準日程', '基準日', '提出期限', '残り日数', '進捗ステータス'],
    plans.map(function (p) {
      return [p.machine, p.initial || '', p.customer, p.title, p.sales, p.frame || '', p.type, p.boards, (p.gen || '') + (p.genGuess ? '（推定）' : ''), p.kind || '', p.channel || '', p.needQuote || '', p.status,
        p.quotes.map(function (q) { return (q.fit ? '' : '△') + q.variant + ':' + q.quoteNo + ' ¥' + Number(q.amount || 0).toLocaleString() + ' ' + q.date; }).join('\n'),
        p.baseLabel, p.baseDate, p.deadline, p.leftDays == null ? '' : p.leftDays, p.progressStatus];
    }));
}

/** API：構成別見積＋機種別予定（進捗管理が読めない時は①だけ返す） */
function apiAssemblyQuotes(p) {
  try {
    var cfg = _aqCfg();
    var comps = _aqCompositions();
    var plans = [], planError = '', stat = {};
    try { plans = _aqPlans(comps, cfg, stat); } catch (e) { planError = '進捗管理システムのスプレッドシートを読めませんでした：' + e.message; }
    if (!planError && !plans.length) {
      planError = stat.noSheet ? '進捗管理のスプレッドシートに「DB_機種」シートがありません。進捗管理システムで「スプレッドシートへ反映」を実行してください。'
        : !stat.models ? '進捗管理の「DB_機種」に機種がありません。'
        : '対象の機種がありません（DB_機種 ' + stat.models + '機種のうち、使用基板が不明 ' + stat.noBoards + '、販売・日程が90日以上前 ' + stat.ended + '）。枠構成タブでハードウェア構成一覧表を取り込むと、使用基板を補えます。';
    }
    try { _aqWriteSheets(comps, plans); } catch (e2) { Logger.log('[apiAssemblyQuotes write] ' + e2.message); }
    cfg.initials = _aqInitials(); cfg.types = _aqTypes(); cfg.variants = _aqVariants(); cfg.gens = _aqGens(); cfg.reuseKeywords = _aqReuseKw();
    return { success: true, comps: comps, plans: plans, planError: planError, cfg: cfg, frames: _adFrames(), mainVariants: AQ_MAIN_VARIANTS, sections: AD_SECTIONS, subSections: AD_SUB, generatedAt: nowJST() };
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
    if (p.gens) {
      c.gens = (p.gens || []).map(function (g) {
        return { name: String(g.name || '').trim(),
                 keywords: (Array.isArray(g.keywords) ? g.keywords : String(g.keywords || '').split(/[,、\n]+/)).map(function (k) { return String(k).trim(); }).filter(Boolean),
                 planReuse: !!g.planReuse,
                 channel: String(g.channel || '').trim() || 'サン電子',
                 needQuote: String(g.needQuote || '要').indexOf('不要') >= 0 ? '不要' : '要' };
      }).filter(function (g) { return g.name; });
    }
    if (p.reuseKeywords != null) {
      c.reuseKeywords = (Array.isArray(p.reuseKeywords) ? p.reuseKeywords : String(p.reuseKeywords).split(/[,、\s\n]+/))
        .map(function (k) { return String(k).trim(); }).filter(Boolean);
    }
    if (p.initials != null) {
      c.initials = (Array.isArray(p.initials) ? p.initials : String(p.initials).split(/[,、\s\/]+/))
        .map(function (k) { return String(k).normalize('NFKC').toUpperCase().replace(/[^A-Z]/g, ''); }).filter(function (k, i, a) { return k && a.indexOf(k) === i; });
    }
    if (p.resetLists) { delete c.initials; delete c.types; delete c.variants; delete c.gens; delete c.reuseKeywords; }
    if (p.progressSsId != null) {
      var m = String(p.progressSsId).match(/[-\w]{25,}/);
      c.progressSsId = m ? m[0] : '';
    }
    PropertiesService.getScriptProperties().setProperty(AQ_PROP, JSON.stringify(c));
    return apiAssemblyQuotes();
  } catch (e) { return { success: false, error: e.message }; }
}
