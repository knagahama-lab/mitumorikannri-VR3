// ============================================================
// 54_composition_sheet.gs
// 客先支給の「構成表」（機種ごとの部品構成・発注書番号つき）を取り込み、
// 品名で注文書・見積書と突き合わせて「注文が来ているか／見積が出ているか」を管理する
//
//   構成表の品名は客先の注文書の品名と一致しているので、品名をキーに
//   VR3 の注文書シート・見積書シートを引き当てられる。
//
//   ① 取込：Google ドライブの「構成表インポート」フォルダに置いた xlsx（またはファイル単体指定）を
//        Google スプレッドシートに変換して読み、シート「構成表_明細」「構成表_機種」に展開する
//        同じ機種を取り込み直すと、その機種の行を新しい構成表で置き換える（修正版の上書き）
//   ② 突合：品名を正規化して注文書・見積書を引き当て、構成表の発注書番号と VR3 の注文を
//        両方並べて差異を出す
//          注文あり         … 構成表に発注書番号があり、VR3 にも同じ番号の注文がある
//          注文あり(番号違い) … 品名では注文があるが、発注書番号が一致しない
//          構成表のみ       … 構成表に発注書番号があるが VR3 に注文が無い（取込漏れの疑い）
//          VR3のみ          … 構成表は空欄だが VR3 に注文がある（構成表が古い疑い）
//          未発注           … どちらにも無い
//   ③ 共有：別スプレッドシート「構成表 管理表」へ片方向で転記する（閲覧・共有専用。書き戻しはしない）
//
//   構成表のレイアウト（E63 の実物で確認）
//     1行目A列=機種、2行目A列=版（日付）
//     見出し行（B列が「発注先」の行）… A:発注先コード B:発注先 C:部品コード
//                                      D〜H:品名（列が右にずれるほど下の階層） I:部品名
//                                      J:納入先 K:単位数 L:発注書番号 M:発注書数量 N:発注単価 O:金額 P:備考
//     1つの品名に複数の発注書がある場合は、品名を空欄にした行が続く（備考に 全新品／全中古／Dリユース 等の区分）
// ============================================================

var CS_SHEET_LINE  = '構成表_明細';
var CS_SHEET_MODEL = '構成表_機種';
var CS_PROP = 'COMPOSITION_CFG';   // { importFolderId, doneFolderId, shareSsId }
var CS_SHARE_NAME = '構成表 管理表（見積管理VR3 連携）';

var CS_LINE_HEADERS = ['機種', '版', '行No', '階層', '発注先コード', '発注先', '部品コード', '品名', '納入先', '単位数',
                       '発注書番号', '発注書数量', '発注単価', '金額', '区分',
                       '注文書状況', 'VR3注文番号', 'VR3発注伝票番号', 'VR3数量', 'VR3単価', '単価差',
                       '見積状況', '見積No', '見積単価', '取込日時', '元ファイル'];
var CS_MODEL_HEADERS = ['機種', '版', '品目数', '明細行数', '発注書番号あり', '未発注', '注文あり', '構成表のみ', 'VR3のみ',
                        '番号違い', '単価差あり', '見積なし', '総額', '取込日時', '元ファイル', '注記'];

var CS_ST_OK     = '注文あり';
var CS_ST_NO_NUM = '注文あり(番号違い)';
var CS_ST_COMP   = '構成表のみ';
var CS_ST_VR3    = 'VR3のみ';
var CS_ST_NONE   = '未発注';

// ------------------------------------------------------------
// 設定
// ------------------------------------------------------------
function _csCfg() {
  var c = {};
  try { c = JSON.parse(PropertiesService.getScriptProperties().getProperty(CS_PROP) || '{}'); } catch (e) {}
  return { importFolderId: c.importFolderId || '', doneFolderId: c.doneFolderId || '', shareSsId: c.shareSsId || '',
           notes: c.notes || {} };
}
function _csSaveCfg(c) { PropertiesService.getScriptProperties().setProperty(CS_PROP, JSON.stringify(c)); }
/** URL でも ID でも受け取れるようにする */
function _csId(s) { var m = String(s == null ? '' : s).match(/[-\w]{25,}/); return m ? m[0] : ''; }

/** 品名・番号の正規化（全角→半角、空白除去、大文字化） */
function _csNorm(s) {
  return String(s == null ? '' : s).normalize('NFKC').replace(/[\s　]+/g, '').toUpperCase();
}
function _csNum(v) {
  if (v === '' || v == null) return null;
  if (typeof v === 'number') return v;
  var n = Number(String(v).normalize('NFKC').replace(/[,¥￥円\s]/g, ''));
  return isNaN(n) ? null : n;
}
function _csText(v) { return String(v == null ? '' : v).trim(); }
function _csDate(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd');
  return _csText(v);
}

// ------------------------------------------------------------
// 構成表の読み取り
// ------------------------------------------------------------
/**
 * 変換済みスプレッドシートの1シートを構成表として読む。
 * 構成表でなければ null を返す（見出し行が見つからない）。
 */
function _csParseSheet(sh, fallbackModel) {
  var last = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (last < 2 || lastCol < 5) return null;
  var v = sh.getRange(1, 1, last, lastCol).getValues();

  // 見出し行＝「発注先」が入っている行（上から30行まで探す）
  var hr = -1, supplierCol = -1;
  for (var r = 0; r < Math.min(v.length, 30) && hr < 0; r++) {
    for (var c = 0; c < lastCol; c++) {
      if (_csNorm(v[r][c]) === '発注先') { hr = r; supplierCol = c; break; }
    }
  }
  if (hr < 0) return null;

  // 列を見出し名で引く
  var ix = { supplier: supplierCol };
  for (var c2 = 0; c2 < lastCol; c2++) {
    var h = _csNorm(v[hr][c2]);
    if (/^部品.?コード$/.test(h)) { if (ix.partCode == null) ix.partCode = c2; }
    else if (h === '納入先') ix.dest = c2;
    else if (h === '単位数') ix.unitQty = c2;
    else if (h === '発注書番号') ix.orderNo = c2;
    else if (h === '発注書数量') ix.orderQty = c2;
    else if (h === '発注単価') ix.price = c2;
    else if (h === '金額') ix.amount = c2;
    else if (h === '備考') ix.note = c2;
  }
  if (ix.partCode == null || ix.dest == null) return null;
  // 品名は「部品コード」の右から「納入先」の左までの帯。右にずれるほど下の階層
  var nameFrom = ix.partCode + 1, nameTo = ix.dest - 1;
  if (nameTo < nameFrom) return null;

  var model = _csText(v[0][0]) || _csText(fallbackModel).replace(/\.(xlsx|xlsm|xls)$/i, '');
  var ver = _csText(v.length > 1 ? v[1][0] : '') || _csText(sh.getName());

  var lines = [], prev = null, seq = 0, total = null;
  for (var r2 = hr + 1; r2 < v.length; r2++) {
    var row = v[r2];
    // 「総額」行に当たったら終わり（金額欄の合計を拾っておく）
    var isTotal = false;
    for (var c3 = 0; c3 < lastCol; c3++) if (_csNorm(row[c3]) === '総額') { isTotal = true; break; }
    if (isTotal) { total = _csNum(row[ix.amount]); break; }

    // 品名（帯の中の最初の非空セル）と階層
    var itemName = '', level = 0;
    for (var c4 = nameFrom; c4 <= nameTo; c4++) {
      if (_csText(row[c4])) { itemName = _csText(row[c4]); level = c4 - nameFrom + 1; break; }
    }
    var partCode = _csText(row[ix.partCode]);
    var orderNo = _csText(row[ix.orderNo]);
    var supplier = _csText(row[ix.supplier]);

    // 品名も部品コードも発注書番号も無い行は飛ばす（金額の式だけが入った空行）
    if (!itemName && !partCode && !orderNo) continue;

    if (!itemName && !partCode && orderNo && prev) {
      // 同じ品名の2本目以降の発注書
      lines.push({
        seq: ++seq, level: prev.level, supplierCode: prev.supplierCode, supplier: prev.supplier,
        partCode: prev.partCode, itemName: prev.itemName, dest: prev.dest, unitQty: prev.unitQty,
        orderNo: orderNo, orderQty: _csNum(row[ix.orderQty]), price: _csNum(row[ix.price]),
        amount: _csNum(row[ix.amount]), note: _csText(row[ix.note]), cont: true,
      });
      continue;
    }
    if (!itemName) continue;   // 部品コードだけの行は親が無いので飛ばす

    prev = {
      seq: ++seq, level: level, supplierCode: _csText(row[0]), supplier: supplier,
      partCode: partCode, itemName: itemName, dest: _csText(row[ix.dest]), unitQty: _csNum(row[ix.unitQty]),
      orderNo: orderNo, orderQty: _csNum(row[ix.orderQty]), price: _csNum(row[ix.price]),
      amount: _csNum(row[ix.amount]), note: _csText(row[ix.note]), cont: false,
    };
    lines.push(prev);
  }
  if (!lines.length) return null;
  // 客先から支給されたままの原本には 発注書番号・数量・単価・金額 の列が無いことがある（弊社で書き加えて使う）
  var cols = { orderNo: ix.orderNo != null, orderQty: ix.orderQty != null, price: ix.price != null,
               amount: ix.amount != null, note: ix.note != null };
  return { model: model, ver: ver, lines: lines, total: total, sheetName: sh.getName(), cols: cols };
}

/** xlsx を Google スプレッドシートに変換して構成表として読む（変換物は読み終えたらゴミ箱へ） */
function _csReadFile(fileId) {
  var f = DriveApp.getFileById(fileId);
  var mime = f.getMimeType();
  var out = [];
  if (mime === MimeType.GOOGLE_SHEETS) {
    SpreadsheetApp.openById(fileId).getSheets().forEach(function (sh) {
      var p = _csParseSheet(sh, f.getName()); if (p) out.push(p);
    });
    out.forEach(function (p) { p.fileName = f.getName(); p.fileUrl = f.getUrl(); });
    return out;
  }
  var conv = Drive.Files.create({ name: '[構成表取込] ' + f.getName(), mimeType: MimeType.GOOGLE_SHEETS }, f.getBlob());
  try {
    SpreadsheetApp.openById(conv.id).getSheets().forEach(function (sh) {
      var p = _csParseSheet(sh, f.getName()); if (p) out.push(p);
    });
    out.forEach(function (p) { p.fileName = f.getName(); p.fileUrl = f.getUrl(); });
    return out;
  } finally {
    try { Drive.Files.update({ trashed: true }, conv.id); }
    catch (e) { try { DriveApp.getFileById(conv.id).setTrashed(true); } catch (e2) {} }
  }
}

// ------------------------------------------------------------
// 注文書・見積書の索引（品名で引く）
// ------------------------------------------------------------
function _csOrderIndex() {
  var map = {};
  var sh = getSpreadsheet().getSheetByName(CONFIG.SHEET_ORDERS);
  if (!sh || sh.getLastRow() < 2) return map;
  var need = ORDER_COLS.FOLDER_URL;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), need)).getValues();
  v.forEach(function (r) {
    var key = _csNorm(r[ORDER_COLS.ITEM_NAME - 1]);
    if (!key) return;
    (map[key] = map[key] || []).push({
      orderNo: _csText(r[ORDER_COLS.ORDER_NO - 1]),
      slipNo: _csText(r[ORDER_COLS.ORDER_SLIP_NO - 1]),
      model: _csText(r[ORDER_COLS.MODEL_CODE - 1]),
      qty: _csNum(r[ORDER_COLS.QTY - 1]),
      price: _csNum(r[ORDER_COLS.UNIT_PRICE - 1]),
      date: _csDate(r[ORDER_COLS.ORDER_DATE - 1]),
    });
  });
  return map;
}

function _csQuoteIndex() {
  // 見積書シート: 管理ID,見積番号,発行日,送り先会社名,送り先担当者名,行No,品名,仕様,数量,単位,単価,…
  var map = {};
  var sh = getSpreadsheet().getSheetByName(CONFIG.SHEET_QUOTES);
  if (!sh || sh.getLastRow() < 2) return map;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), 12)).getValues();
  v.forEach(function (r) {
    var key = _csNorm(r[6]);   // 7列目 = 品名
    if (!key) return;
    (map[key] = map[key] || []).push({ quoteNo: _csText(r[1]), date: _csDate(r[2]), price: _csNum(r[10]) });
  });
  return map;
}

/** 1行の突合結果 */
function _csMatch(line, oidx, qidx) {
  var key = _csNorm(line.itemName);
  var orders = oidx[key] || [], quotes = qidx[key] || [];
  var num = _csNorm(line.orderNo);
  var hit = num ? orders.filter(function (o) { return _csNorm(o.slipNo) === num || _csNorm(o.orderNo) === num; }) : [];
  var use = hit[0] || orders[0] || null;

  var status = num
    ? (hit.length ? CS_ST_OK : (orders.length ? CS_ST_NO_NUM : CS_ST_COMP))
    : (orders.length ? CS_ST_VR3 : CS_ST_NONE);

  // 単価差は番号まで一致した注文だけで見る（別ロットの単価と比べても意味が無い）
  var diff = (hit.length && line.price != null && use && use.price != null) ? Math.round((line.price - use.price) * 100) / 100 : null;
  var q = quotes.sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); })[0] || null;

  return {
    status: status,
    orderNo: use ? use.orderNo : '', slipNo: use ? use.slipNo : '',
    qty: use ? use.qty : null, price: use ? use.price : null, diff: diff,
    quoteStatus: q ? '見積あり' : '見積なし', quoteNo: q ? q.quoteNo : '', quotePrice: q ? q.price : null,
  };
}

// ------------------------------------------------------------
// シート
// ------------------------------------------------------------
function _csLineSheet() {
  var ss = getSpreadsheet(), sh = ss.getSheetByName(CS_SHEET_LINE);
  if (!sh) {
    sh = ss.insertSheet(CS_SHEET_LINE);
    sh.getRange(1, 1, 1, CS_LINE_HEADERS.length).setValues([CS_LINE_HEADERS])
      .setBackground('#E8F5E9').setFontWeight('bold').setFontSize(10);
    sh.setFrozenRows(1);
    // 部品コード・発注書番号の先頭0や日付化を防ぐ
    sh.getRange(1, 7, sh.getMaxRows(), 2).setNumberFormat('@');
    sh.getRange(1, 11, sh.getMaxRows(), 1).setNumberFormat('@');
  }
  return sh;
}
function _csModelSheet() {
  var ss = getSpreadsheet(), sh = ss.getSheetByName(CS_SHEET_MODEL);
  if (!sh) {
    sh = ss.insertSheet(CS_SHEET_MODEL);
    sh.getRange(1, 1, 1, CS_MODEL_HEADERS.length).setValues([CS_MODEL_HEADERS])
      .setBackground('#E3F2FD').setFontWeight('bold').setFontSize(10);
    sh.setFrozenRows(1);
  }
  return sh;
}

function _csReadLines() {
  var sh = _csLineSheet();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, CS_LINE_HEADERS.length).getValues()
    .filter(function (r) { return _csText(r[0]); });
}

/** 取り込んだ構成表を「構成表_明細」に書く（同じ機種の既存行は消して置き換える） */
function _csWriteModel(parsed, oidx, qidx) {
  var sh = _csLineSheet(), now = nowJST();
  var model = parsed.model;
  var keep = _csReadLines().filter(function (r) { return _csText(r[0]) !== model; });

  var rows = parsed.lines.map(function (l) {
    var m = _csMatch(l, oidx, qidx);
    return [model, parsed.ver, l.seq, l.level, l.supplierCode, l.supplier, l.partCode, l.itemName, l.dest,
            l.unitQty == null ? '' : l.unitQty,
            l.orderNo, l.orderQty == null ? '' : l.orderQty, l.price == null ? '' : l.price, l.amount == null ? '' : l.amount, l.note,
            m.status, m.orderNo, m.slipNo, m.qty == null ? '' : m.qty, m.price == null ? '' : m.price, m.diff == null ? '' : m.diff,
            m.quoteStatus, m.quoteNo, m.quotePrice == null ? '' : m.quotePrice,
            now, parsed.fileName || ''];
  });

  var all = keep.concat(rows);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, CS_LINE_HEADERS.length).clearContent();
  if (all.length) sh.getRange(2, 1, all.length, CS_LINE_HEADERS.length).setValues(all);
  return rows.length;
}

/** 「構成表_機種」を明細から作り直す */
function _csRebuildModelSheet() {
  var rows = _csReadLines(), by = {};
  rows.forEach(function (r) {
    var k = _csText(r[0]);
    var g = by[k] || (by[k] = { model: k, ver: _csText(r[1]), items: {}, lines: 0, withNo: 0, none: 0,
                                ok: 0, comp: 0, vr3: 0, noNum: 0, diff: 0, noQuote: 0, total: 0,
                                at: _csText(r[24]), file: _csText(r[25]) });
    g.lines++;
    if (_csText(r[7])) g.items[_csNorm(r[7])] = 1;
    if (_csText(r[10])) g.withNo++;
    var st = _csText(r[15]);
    if (st === CS_ST_OK) g.ok++;
    else if (st === CS_ST_COMP) g.comp++;
    else if (st === CS_ST_VR3) g.vr3++;
    else if (st === CS_ST_NO_NUM) g.noNum++;
    else if (st === CS_ST_NONE) g.none++;
    if (_csNum(r[20])) g.diff++;   // 単価差 0（＝一致を確認できた）は「単価差あり」に数えない
    if (_csText(r[21]) === '見積なし') g.noQuote++;
    var amt = _csNum(r[13]); if (amt) g.total += amt;
  });

  var notes = _csCfg().notes || {};
  var out = Object.keys(by).sort().map(function (k) {
    var g = by[k];
    return [g.model, g.ver, Object.keys(g.items).length, g.lines, g.withNo, g.none,
            g.ok, g.comp, g.vr3, g.noNum, g.diff, g.noQuote, g.total, g.at, g.file, notes[g.model] || ''];
  });
  var sh = _csModelSheet();
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, CS_MODEL_HEADERS.length).clearContent();
  if (out.length) sh.getRange(2, 1, out.length, CS_MODEL_HEADERS.length).setValues(out);
  return out.length;
}

/** 既存の明細を注文書・見積書と再突合する（構成表を読み直さずに状況だけ更新） */
function _csRematchAll() {
  var sh = _csLineSheet();
  if (sh.getLastRow() < 2) return 0;
  var oidx = _csOrderIndex(), qidx = _csQuoteIndex(), now = nowJST();
  var n = sh.getLastRow() - 1;
  var v = sh.getRange(2, 1, n, CS_LINE_HEADERS.length).getValues();
  var upd = v.map(function (r) {
    if (!_csText(r[0])) return r.slice(15, 25);
    var m = _csMatch({ itemName: _csText(r[7]), orderNo: _csText(r[10]), price: _csNum(r[12]) }, oidx, qidx);
    return [m.status, m.orderNo, m.slipNo, m.qty == null ? '' : m.qty, m.price == null ? '' : m.price,
            m.diff == null ? '' : m.diff, m.quoteStatus, m.quoteNo, m.quotePrice == null ? '' : m.quotePrice, now];
  });
  sh.getRange(2, 16, n, 10).setValues(upd);
  _csRebuildModelSheet();
  return n;
}

// ------------------------------------------------------------
// 取込
// ------------------------------------------------------------
/** ファイル1つを取り込む。戻り値は取り込んだ機種の一覧 */
function _csImportFile(fileId) {
  var parsedList = _csReadFile(fileId);
  if (!parsedList.length) throw new Error('構成表として読める表が見つかりませんでした（見出しに「発注先」「部品コード」「納入先」がある表を探します）');
  var oidx = _csOrderIndex(), qidx = _csQuoteIndex(), done = [];
  var cfg = _csCfg(), notes = cfg.notes || {};
  parsedList.forEach(function (p) {
    var n = _csWriteModel(p, oidx, qidx);
    // 発注書番号の列が無い構成表は「注文が来ているか」を構成表側から判定できないので、はっきり残す
    var miss = [];
    if (p.cols && !p.cols.orderNo) miss.push('発注書番号');
    if (p.cols && !p.cols.price) miss.push('発注単価');
    if (p.cols && !p.cols.amount) miss.push('金額');
    notes[p.model] = miss.length
      ? 'この構成表には ' + miss.join('・') + ' の列がありません（客先から支給されたままの原本の可能性）。発注書番号を書き加えた方の構成表を取り込むと、注文の有無を突き合わせられます。'
      : '';
    done.push({ model: p.model, ver: p.ver, lines: n, sheetName: p.sheetName, note: notes[p.model] });
  });
  cfg.notes = notes; _csSaveCfg(cfg);
  _csRebuildModelSheet();
  return done;
}

/** p: { fileId | url } */
function apiCompImportFile(p) {
  try {
    p = p || {};
    var id = _csId(p.fileId || p.url || '');
    if (!id) return { success: false, error: 'ファイルのURLまたはIDを指定してください' };
    var done = _csImportFile(id);
    var warn = done.filter(function (d) { return d.note; }).map(function (d) { return d.model + '：' + d.note; }).join('\n');
    return { success: true, imported: done,
             message: done.map(function (d) { return d.model + '（' + d.lines + '行）'; }).join('、') + ' を取り込みました'
               + (warn ? '\n⚠ ' + warn : '') };
  } catch (e) { return { success: false, error: _csErr(e) }; }
}

/**
 * 「構成表インポート」フォルダを見て、入っている xlsx を全部取り込む。
 * 済んだファイルは処理済みフォルダ（未設定ならインポートフォルダ内の「処理済み」）へ移す。
 * processDriveImports（5分ごと）から呼ばれる。
 */
function importCompositionSheetsFromFolder() {
  var cfg = _csCfg();
  if (!cfg.importFolderId) return 0;
  var folder;
  try { folder = DriveApp.getFolderById(cfg.importFolderId); }
  catch (e) { Logger.log('[COMPOSITION] インポートフォルダが開けません: ' + e.message); return 0; }

  var done;
  if (cfg.doneFolderId) { try { done = DriveApp.getFolderById(cfg.doneFolderId); } catch (e2) { done = null; } }
  if (!done) {
    var it = folder.getFoldersByName('処理済み');
    done = it.hasNext() ? it.next() : folder.createFolder('処理済み');
  }

  var files = folder.getFiles(), count = 0;
  while (files.hasNext()) {
    var f = files.next();
    if (!/\.(xlsx|xlsm|xls)$/i.test(f.getName()) && f.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;
    if (/^~\$/.test(f.getName())) continue;   // Excel の一時ファイル
    try {
      var r = _csImportFile(f.getId());
      f.moveTo(done);
      count += r.length;
      Logger.log('[COMPOSITION] 取込: ' + f.getName() + ' → ' + r.map(function (x) { return x.model; }).join('、'));
    } catch (e3) {
      Logger.log('[COMPOSITION] 取込失敗: ' + f.getName() + ' / ' + e3.message);
    }
  }
  return count;
}

/**
 * パソコンから直接アップロードして取り込む（ドライブに置かなくてよい）
 * p: { fileName, mimeType, base64Data }
 * 取り込んだ元ファイルは「処理済み」フォルダに残す（未設定ならマイドライブ直下）
 */
function apiCompUpload(p) {
  var tmp = null;
  try {
    p = p || {};
    if (!p.base64Data || !p.fileName) return { success: false, error: 'ファイルを読み取れませんでした' };
    if (!/[.](xlsx|xlsm|xls)$/i.test(p.fileName)) return { success: false, error: 'Excelファイル（.xlsx）を選んでください：' + p.fileName };
    var mime = p.mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    var blob = Utilities.newBlob(Utilities.base64Decode(p.base64Data), mime, p.fileName);

    var cfg = _csCfg(), dest = null;
    if (cfg.doneFolderId) { try { dest = DriveApp.getFolderById(cfg.doneFolderId); } catch (e) {} }
    if (!dest && cfg.importFolderId) {
      try {
        var f = DriveApp.getFolderById(cfg.importFolderId);
        var it = f.getFoldersByName('処理済み');
        dest = it.hasNext() ? it.next() : f.createFolder('処理済み');
      } catch (e2) {}
    }
    tmp = (dest || DriveApp.getRootFolder()).createFile(blob);

    var done = _csImportFile(tmp.getId());
    var warn = done.filter(function (d) { return d.note; }).map(function (d) { return d.model + '：' + d.note; }).join('\n');
    return { success: true, imported: done, fileUrl: tmp.getUrl(),
             message: done.map(function (d) { return d.model + '（' + d.lines + '行）'; }).join('、') + ' を取り込みました'
               + (warn ? '\n⚠ ' + warn : '') };
  } catch (e) {
    // 読めなかったファイルをドライブに残さない
    if (tmp) { try { tmp.setTrashed(true); } catch (e3) {} }
    return { success: false, error: _csErr(e) };
  }
}

function apiCompImportFolder() {
  try {
    var cfg = _csCfg();
    if (!cfg.importFolderId) return { success: false, error: '「構成表インポート」フォルダが未設定です。⚙ 設定 でフォルダのURLを入れてください' };
    var n = importCompositionSheetsFromFolder();
    return { success: true, message: n ? n + '件の構成表を取り込みました' : 'フォルダに未処理の構成表はありませんでした' };
  } catch (e) { return { success: false, error: _csErr(e) }; }
}

// ------------------------------------------------------------
// 共有スプレッドシートへの転記（片方向）
// ------------------------------------------------------------
function _csShareSs() {
  var cfg = _csCfg();
  if (cfg.shareSsId) { try { return SpreadsheetApp.openById(cfg.shareSsId); } catch (e) {} }
  var ss = SpreadsheetApp.create(CS_SHARE_NAME);
  cfg.shareSsId = ss.getId(); _csSaveCfg(cfg);
  return ss;
}

/** 共有スプレッドシートへ「明細」「機種サマリ」を書き出す */
function apiCompExport() {
  try {
    var ss = _csShareSs(), now = nowJST();
    var lines = _csReadLines();
    var models = (function () {
      var sh = _csModelSheet();
      return sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, CS_MODEL_HEADERS.length).getValues();
    })();

    var write = function (name, headers, rows, color) {
      var sh = ss.getSheetByName(name) || ss.insertSheet(name);
      sh.clear();
      sh.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground(color).setFontWeight('bold').setFontSize(10);
      if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
      sh.setFrozenRows(1);
      var f = sh.getFilter(); if (f) f.remove();
      if (rows.length) sh.getRange(1, 1, rows.length + 1, headers.length).createFilter();
      sh.getRange(1, headers.length + 2).setValue('最終更新: ' + now + '（見積管理VR3からの転記。こちらで編集しても戻りません）');
      return sh;
    };

    var msh = write('機種サマリ', CS_MODEL_HEADERS, models, '#E3F2FD');
    var lsh = write('明細', CS_LINE_HEADERS, lines, '#E8F5E9');

    // 注文書状況に色を付ける（明細の16列目）
    if (lines.length) {
      var rng = lsh.getRange(2, 16, lines.length, 1);
      var rules = [
        { t: CS_ST_NONE,   bg: '#FEE2E2', fg: '#B91C1C' },
        { t: CS_ST_COMP,   bg: '#FFEDD5', fg: '#9A3412' },
        { t: CS_ST_VR3,    bg: '#FEF3C7', fg: '#92400E' },
        { t: CS_ST_NO_NUM, bg: '#FEF3C7', fg: '#92400E' },
        { t: CS_ST_OK,     bg: '#DCFCE7', fg: '#166534' },
      ].map(function (x) {
        return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(x.t)
          .setBackground(x.bg).setFontColor(x.fg).setRanges([rng]).build();
      });
      lsh.setConditionalFormatRules(rules);
    }

    // 既定の空シートが残っていたら消す
    ss.getSheets().forEach(function (sh) {
      if (['シート1', 'Sheet1'].indexOf(sh.getName()) >= 0 && ss.getSheets().length > 1) {
        try { ss.deleteSheet(sh); } catch (e) {}
      }
    });

    return { success: true, url: ss.getUrl(), ssId: ss.getId(),
             message: '明細' + lines.length + '行・機種' + models.length + '件を転記しました' };
  } catch (e) { return { success: false, error: _csErr(e) }; }
}

// ------------------------------------------------------------
// 画面用 API
// ------------------------------------------------------------
function apiCompList(p) {
  var step = '開始';
  try {
    p = p || {};
    step = '設定の読み込み';
    var cfg = _csCfg();
    step = 'スプレッドシートを開く';
    var ss = getSpreadsheet();
    step = 'シート「' + CS_SHEET_LINE + '」の準備';
    var lsh = _csLineSheet();
    step = 'シート「' + CS_SHEET_MODEL + '」の準備';
    var msh = _csModelSheet();
    if (p.rematch) { step = '注文書・見積書との再突合'; _csRematchAll(); }
    step = '明細の読み込み';
    var lineRows = _csReadLines();
    step = '機種サマリの読み込み';
    var modelRows = msh.getLastRow() < 2 ? [] : msh.getRange(2, 1, msh.getLastRow() - 1, CS_MODEL_HEADERS.length).getValues();

    step = '画面へ渡す形に整える';
    // google.script.run は Date をそのまま返せないことがあるので文字列にしておく
    var cell = function (x) {
      if (x instanceof Date) return Utilities.formatDate(x, 'Asia/Tokyo', 'yyyy/MM/dd');
      return (x === null || x === undefined) ? '' : x;
    };
    var obj = function (headers, r) { var o = {}; headers.forEach(function (h, i) { o[h] = cell(r[i]); }); return o; };
    return {
      success: true,
      models: modelRows.map(function (r) { return obj(CS_MODEL_HEADERS, r); }),
      lines: lineRows.map(function (r) { return obj(CS_LINE_HEADERS, r); }),
      cfg: { importFolderId: cfg.importFolderId, doneFolderId: cfg.doneFolderId, shareSsId: cfg.shareSsId,
             shareUrl: cfg.shareSsId ? 'https://docs.google.com/spreadsheets/d/' + cfg.shareSsId + '/edit' : '' },
      lineHeaders: CS_LINE_HEADERS, modelHeaders: CS_MODEL_HEADERS,
      generatedAt: nowJST(),
    };
  } catch (e) {
    var msg = _csErr(e);
    Logger.log('[apiCompList] ' + step + ': ' + msg);
    return { success: false, error: '［' + step + '］で失敗しました: ' + msg, step: step };
  }
}

/** 例外から必ず読めるメッセージを作る（message が無い例外でも空にしない） */
function _csErr(e) {
  if (!e) return '原因不明のエラー';
  var m = (typeof e === 'string') ? e : (e.message || e.toString() || '');
  return String(m).trim() || '原因不明のエラー';
}

/** p: { importFolderId, doneFolderId, shareSsId } */
function apiCompCfgSave(p) {
  try {
    p = p || {};
    var c = _csCfg();
    if (p.importFolderId != null) c.importFolderId = _csId(p.importFolderId);
    if (p.doneFolderId != null) c.doneFolderId = _csId(p.doneFolderId);
    if (p.shareSsId != null) c.shareSsId = _csId(p.shareSsId);
    _csSaveCfg(c);
    return apiCompList();
  } catch (e) { return { success: false, error: _csErr(e) }; }
}

/** 機種1件分の明細を消す */
function apiCompDeleteModel(p) {
  try {
    p = p || {};
    var model = _csText(p.model);
    if (!model) return { success: false, error: '機種を指定してください' };
    var sh = _csLineSheet();
    var keep = _csReadLines().filter(function (r) { return _csText(r[0]) !== model; });
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, CS_LINE_HEADERS.length).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, CS_LINE_HEADERS.length).setValues(keep);
    var cfg = _csCfg(); if (cfg.notes) { delete cfg.notes[model]; _csSaveCfg(cfg); }
    _csRebuildModelSheet();
    return { success: true, message: model + ' の明細を削除しました' };
  } catch (e) { return { success: false, error: _csErr(e) }; }
}
