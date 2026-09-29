// ============================================================
// 33_quote_entry.gs
// 見積入力（基幹システム風の見積書作成・修正・改訂・印刷）
//
//   ・見積の全項目（ヘッダ＋明細）は「見積入力」シートに JSON で保持（見積No.×改訂で1行）
//   ・保存時に 管理シート（見積行）と 見積書シート（明細行）へ同期するので、
//     見積書一覧・見積→注文カレンダー・AI紐づけ・単価検索にそのまま載る
//   ・消費税は税率×税区分ごとに1回端数処理（切り捨て）＝インボイス方式
// ============================================================

var QE_SHEET   = '見積入力';
var QE_HEADERS = ['管理ID','見積No','改訂','見積日','得意先','件名','税込合計','承認','更新者','更新日時','最新','データ(JSON)'];
var QE_C = { ID:0, NO:1, REV:2, DATE:3, CLIENT:4, SUBJECT:5, TOTAL:6, APPROVED:7, UPDATED_BY:8, UPDATED_AT:9, LATEST:10, JSON:11 };
var QE_COMPANY_PROP = 'QE_COMPANY_INFO';

function _qeSheet() {
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName(QE_SHEET);
  if (sh) return sh;
  sh = ss.insertSheet(QE_SHEET);
  sh.getRange(1, 1, 1, QE_HEADERS.length).setValues([QE_HEADERS])
    .setBackground('#DBEAFE').setFontWeight('bold').setFontSize(10);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, sh.getMaxRows(), QE_HEADERS.length).setNumberFormat('@');
  return sh;
}

function _qeRows() {
  var sh = _qeSheet();
  var last = sh.getLastRow();
  if (last <= 1) return [];
  return sh.getRange(2, 1, last - 1, QE_HEADERS.length).getValues();
}

// ── 金額計算（フロントと同じロジック。保存時はサーバー側の結果を正とする） ──
function _qeCalc(doc) {
  var groups = {};
  var subtotal = 0, taxTotal = 0, total = 0;
  (doc.lines || []).forEach(function(l) {
    if (l.kind === '摘要') { l.amount = 0; l.tax = 0; return; }
    var qty = Number(l.qty) || 0;
    if (!qty && Number(l.perCase) && Number(l.cases)) qty = Number(l.perCase) * Number(l.cases);
    l.qty = qty;
    var amt = Math.round(qty * (Number(l.price) || 0));
    if (l.kind === '値引') amt = -Math.abs(amt);
    l.amount = amt;
    var rate = Number(l.taxRate) || 0;
    var type = l.taxType || '外税';
    if (type === '非課税') rate = 0;
    l.tax = type === '内税' ? Math.floor(amt * rate / (100 + rate)) : Math.floor(amt * rate / 100);
    var key = type + '|' + rate;
    groups[key] = groups[key] || { type: type, rate: rate, amount: 0 };
    groups[key].amount += amt;
  });
  var breakdown = [];
  Object.keys(groups).forEach(function(k) {
    var g = groups[k];
    var tax = g.type === '内税' ? Math.floor(g.amount * g.rate / (100 + g.rate)) : Math.floor(g.amount * g.rate / 100);
    var excl = g.type === '内税' ? g.amount - tax : g.amount;
    subtotal += excl; taxTotal += tax; total += excl + tax;
    breakdown.push({ type: g.type, rate: g.rate, base: excl, tax: tax });
  });
  doc.subtotal = subtotal; doc.tax = taxTotal; doc.total = total; doc.breakdown = breakdown;
  return doc;
}

function _qeIsBlankLine(l) {
  return !String(l.code || '').trim() && !String(l.name || '').trim() && !String(l.spec || '').trim() && !Number(l.qty) && !Number(l.price);
}

// ── 採番：数字のみの見積No.の最大値＋1（無ければ 100001） ──
function _qeNextNo() {
  var max = 0;
  var bump = function(v) { var s = String(v || '').trim(); if (/^\d{1,10}$/.test(s)) max = Math.max(max, Number(s)); };
  _qeRows().forEach(function(r) { bump(r[QE_C.NO]); });
  getAllMgmtData().forEach(function(r) { bump(r[MGMT_COLS.QUOTE_NO - 1]); });
  return String(max ? max + 1 : 100001);
}

function apiQeNextNo() {
  try { return { success: true, quoteNo: _qeNextNo() }; }
  catch (e) { return { success: false, error: e.message }; }
}

// ── 一覧 ──
function apiQeList(p) {
  try {
    var kw = normalizeText((p && p.keyword) || '');
    var all = !!(p && p.allRevisions);
    var items = _qeRows().filter(function(r) { return all || String(r[QE_C.LATEST]) !== 'FALSE'; }).map(function(r) {
      return {
        mgmtId: String(r[QE_C.ID]), quoteNo: String(r[QE_C.NO]), rev: Number(r[QE_C.REV]) || 1,
        quoteDate: _ofDate(r[QE_C.DATE]), client: String(r[QE_C.CLIENT]), subject: String(r[QE_C.SUBJECT]),
        total: Number(r[QE_C.TOTAL]) || 0, approved: String(r[QE_C.APPROVED]) === 'TRUE',
        updatedBy: String(r[QE_C.UPDATED_BY]), updatedAt: String(r[QE_C.UPDATED_AT]), latest: String(r[QE_C.LATEST]) !== 'FALSE',
      };
    }).filter(function(it) {
      return !kw || normalizeText([it.quoteNo, it.client, it.subject].join(' ')).indexOf(kw) >= 0;
    });
    items.sort(function(a, b) { return String(b.updatedAt).localeCompare(String(a.updatedAt)); });
    return { success: true, items: items.slice(0, 300) };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ── 取得（見積入力データが無い既存見積は 管理シート＋見積書シート から組み立てる） ──
function apiQeGet(p) {
  try {
    p = p || {};
    var rows = _qeRows();
    var hit = null;
    rows.forEach(function(r) {
      var matchNo = p.quoteNo && String(r[QE_C.NO]) === String(p.quoteNo) && (p.rev ? Number(r[QE_C.REV]) === Number(p.rev) : String(r[QE_C.LATEST]) !== 'FALSE');
      var matchId = !p.quoteNo && p.mgmtId && String(r[QE_C.ID]) === String(p.mgmtId) && String(r[QE_C.LATEST]) !== 'FALSE';
      if (matchNo || matchId) hit = r;
    });
    if (hit) {
      var doc = JSON.parse(String(hit[QE_C.JSON] || '{}'));
      doc.updatedBy = String(hit[QE_C.UPDATED_BY]); doc.updatedAt = String(hit[QE_C.UPDATED_AT]);
      doc.revisions = rows.filter(function(r) { return String(r[QE_C.NO]) === String(doc.quoteNo); })
        .map(function(r) { return { rev: Number(r[QE_C.REV]) || 1, updatedAt: String(r[QE_C.UPDATED_AT]), total: Number(r[QE_C.TOTAL]) || 0 }; })
        .sort(function(a, b) { return a.rev - b.rev; });
      return { success: true, doc: doc, source: 'entry' };
    }
    return _qeImportExisting(p);
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function _qeImportExisting(p) {
  var mg = getAllMgmtData().map(_rowToObject).filter(function(o) {
    return (p.mgmtId && o.id === String(p.mgmtId)) || (p.quoteNo && o.quoteNo === String(p.quoteNo) && !(o.orderNo && !o.quotePdfUrl));
  })[0];
  if (!mg) return { success: false, error: '見積が見つかりません' };
  var lines = [];
  var qs = getSpreadsheet().getSheetByName(CONFIG.SHEET_QUOTES);
  var destPerson = '';
  if (qs && qs.getLastRow() > 1) {
    qs.getRange(2, 1, qs.getLastRow() - 1, QUOTE_COLS.FOLDER_URL).getValues().forEach(function(r) {
      if (String(r[QUOTE_COLS.MGMT_ID - 1]) !== mg.id) return;
      destPerson = destPerson || String(r[QUOTE_COLS.DEST_PERSON - 1] || '');
      var remarks = String(r[QUOTE_COLS.REMARKS - 1] || '');
      var code = (remarks.match(/品番:([^\s／]+)/) || [])[1] || '';
      lines.push({
        kind: '通常', code: code, name: String(r[QUOTE_COLS.ITEM_NAME - 1] || ''), spec: String(r[QUOTE_COLS.SPEC - 1] || ''),
        qty: Number(r[QUOTE_COLS.QTY - 1]) || 0, unit: String(r[QUOTE_COLS.UNIT - 1] || ''),
        price: Number(r[QUOTE_COLS.UNIT_PRICE - 1]) || 0, taxRate: 10, taxType: '外税',
        note: remarks.replace(/品番:[^\s／]+\s*／?\s*/, ''),
      });
    });
  }
  var doc = _qeCalc({
    mgmtId: mg.id, quoteNo: mg.quoteNo, rev: 1, quoteDate: mg.quoteDate, dept: '', assignee: mg.assignee,
    currency: '円', approved: false, orderType: mg.orderType, modelCode: mg.modelCode, boardName: '',
    clientCode: '', client: mg.client, clientPerson: destPerson, billCode: '', billTo: mg.client, contact: '',
    subject: mg.subject, deliveryPlace: '貴社ご指定場所', deliveryTerm: '別途ご相談', payment: '', validity: '発行日より1ヶ月',
    note: mg.memo, lines: lines, pdfUrl: mg.quotePdfUrl,
  });
  return { success: true, doc: doc, source: 'imported' };
}

// ── 保存（p.doc / p.asRevision） ──
function apiQeSave(p) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var doc = (p && p.doc) || {};
    doc.lines = (doc.lines || []).filter(function(l) { return !_qeIsBlankLine(l); });
    if (!String(doc.client || '').trim()) return { success: false, error: '得意先を入力してください' };
    if (!doc.lines.length) return { success: false, error: '明細を1行以上入力してください' };
    _qeCalc(doc);
    doc.quoteDate = _ofDate(doc.quoteDate) || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
    doc.quoteNo = String(doc.quoteNo || '').trim() || _qeNextNo();

    var sh   = _qeSheet();
    var rows = _qeRows();
    var user = _ofUser();
    var now  = nowJST();
    var sameNo = [];
    rows.forEach(function(r, i) { if (String(r[QE_C.NO]) === doc.quoteNo) sameNo.push({ r: r, row: i + 2 }); });

    // 別の見積（管理IDが違う）で同じ見積No.が既にあれば拒否
    if (!p.asRevision && sameNo.length && doc.mgmtId && sameNo.some(function(x) { return String(x.r[QE_C.ID]) !== String(doc.mgmtId); })) {
      return { success: false, error: '見積No. ' + doc.quoteNo + ' は別の見積で使用済みです' };
    }
    if (!doc.mgmtId && sameNo.length) return { success: false, error: '見積No. ' + doc.quoteNo + ' は使用済みです（「開く」から修正してください）' };

    if (p.asRevision) {
      if (!sameNo.length) return { success: false, error: '改訂版は保存済みの見積からのみ作成できます' };
      doc.rev = sameNo.reduce(function(m, x) { return Math.max(m, Number(x.r[QE_C.REV]) || 1); }, 1) + 1;
      doc.approved = false;
      sameNo.forEach(function(x) { sh.getRange(x.row, QE_C.LATEST + 1).setValue('FALSE'); });
    }
    doc.rev = Number(doc.rev) || 1;
    doc.mgmtId = doc.mgmtId || _qeFindMgmtIdByQuoteNo(doc.quoteNo) || generateMgmtId();
    doc.approvedBy = doc.approved ? (doc.approvedBy || user) : '';

    var rowVals = [doc.mgmtId, doc.quoteNo, String(doc.rev), doc.quoteDate, doc.client, doc.subject || '', String(doc.total),
                   doc.approved ? 'TRUE' : 'FALSE', user, now, 'TRUE', ''];
    var toStore = JSON.parse(JSON.stringify(doc));
    delete toStore.revisions; delete toStore.updatedAt; delete toStore.updatedBy;
    rowVals[QE_C.JSON] = JSON.stringify(toStore);

    var target = sameNo.filter(function(x) { return (Number(x.r[QE_C.REV]) || 1) === doc.rev; })[0];
    if (target && !p.asRevision) sh.getRange(target.row, 1, 1, QE_HEADERS.length).setValues([rowVals]);
    else sh.appendRow(rowVals);

    _qeSyncMgmt(doc, now);
    _qeSyncQuoteLines(doc);
    return { success: true, doc: apiQeGet({ quoteNo: doc.quoteNo }).doc };
  } catch (e) {
    Logger.log('[apiQeSave] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function _qeFindMgmtIdByQuoteNo(quoteNo) {
  var hit = getAllMgmtData().map(_rowToObject).filter(function(o) {
    return o.quoteNo === quoteNo && !(o.orderNo && !o.quotePdfUrl);
  })[0];
  return hit ? hit.id : '';
}

/** 管理シートの見積行を作成／更新（ステータスは既存値を保持） */
function _qeSyncMgmt(doc, now) {
  var sheet = getSpreadsheet().getSheetByName(CONFIG.SHEET_MANAGEMENT);
  var last  = sheet.getLastRow();
  var ids   = last > 1 ? sheet.getRange(2, MGMT_COLS.ID, last - 1, 1).getValues().map(function(r) { return String(r[0]); }) : [];
  var idx   = ids.indexOf(String(doc.mgmtId));
  var width = Math.max(sheet.getLastColumn(), MGMT_COLS.BOARD_NAME);
  var row   = idx >= 0 ? sheet.getRange(idx + 2, 1, 1, width).getValues()[0] : new Array(width).fill('');
  var set = function(col, v) { row[col - 1] = v; };
  set(MGMT_COLS.ID, doc.mgmtId);
  set(MGMT_COLS.QUOTE_NO, doc.quoteNo);
  set(MGMT_COLS.SUBJECT, doc.subject || '');
  set(MGMT_COLS.CLIENT, doc.client);
  if (idx < 0) set(MGMT_COLS.STATUS, CONFIG.STATUS.PLANNED);
  set(MGMT_COLS.QUOTE_DATE, doc.quoteDate);
  set(MGMT_COLS.QUOTE_AMOUNT, doc.subtotal);
  set(MGMT_COLS.TAX, doc.tax);
  set(MGMT_COLS.TOTAL, doc.total);
  if (doc.orderType) set(MGMT_COLS.ORDER_TYPE, doc.orderType);
  if (doc.modelCode) set(MGMT_COLS.MODEL_CODE, doc.modelCode);
  if (doc.assignee)  set(MGMT_COLS.ASSIGNEE, doc.assignee);
  set(MGMT_COLS.MEMO, doc.note || '');
  if (idx < 0) set(MGMT_COLS.CREATED_AT, now);
  set(MGMT_COLS.UPDATED_AT, now);
  set(MGMT_COLS.IS_LATEST, 'TRUE');
  set(MGMT_COLS.REVISION_NO, String(doc.rev));
  if (doc.boardName) set(MGMT_COLS.BOARD_NAME, doc.boardName);
  if (idx >= 0) sheet.getRange(idx + 2, 1, 1, width).setValues([row]);
  else sheet.appendRow(row);
  if (doc.modelCode) { try { _ensureModelCode(doc.modelCode, ''); } catch (e) {} }
}

/** 見積書シートの明細を差し替え（検索・単価照合・AI紐づけ用） */
function _qeSyncQuoteLines(doc) {
  var sheet = getSpreadsheet().getSheetByName(CONFIG.SHEET_QUOTES);
  if (!sheet) return;
  var last = sheet.getLastRow();
  if (last > 1) {
    var ids = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (var i = ids.length - 1; i >= 0; i--) {
      if (String(ids[i][0]) === String(doc.mgmtId)) sheet.deleteRow(i + 2);
    }
  }
  var rows = doc.lines.filter(function(l) { return l.kind !== '摘要'; }).map(function(l, i) {
    var remarks = [l.code ? '品番:' + l.code : '', l.kind !== '通常' ? l.kind : '', l.note || ''].filter(String).join(' ／ ');
    var r = new Array(QUOTE_COLS.FOLDER_URL).fill('');
    r[QUOTE_COLS.MGMT_ID - 1]      = doc.mgmtId;
    r[QUOTE_COLS.QUOTE_NO - 1]     = doc.quoteNo;
    r[QUOTE_COLS.ISSUE_DATE - 1]   = doc.quoteDate;
    r[QUOTE_COLS.DEST_COMPANY - 1] = doc.client;
    r[QUOTE_COLS.DEST_PERSON - 1]  = doc.clientPerson || '';
    r[QUOTE_COLS.LINE_NO - 1]      = i + 1;
    r[QUOTE_COLS.ITEM_NAME - 1]    = l.name || l.code || '';
    r[QUOTE_COLS.SPEC - 1]         = l.spec || '';
    r[QUOTE_COLS.QTY - 1]          = l.qty;
    r[QUOTE_COLS.UNIT - 1]         = l.unit || '';
    r[QUOTE_COLS.UNIT_PRICE - 1]   = Number(l.price) || 0;
    r[QUOTE_COLS.AMOUNT - 1]       = l.amount;
    r[QUOTE_COLS.REMARKS - 1]      = remarks;
    r[QUOTE_COLS.PDF_URL - 1]      = doc.pdfUrl || '';
    return r;
  });
  if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

// ── 入力補助マスタ（得意先・品目・部門・担当者・自社情報） ──
function apiQeMasters() {
  try {
    var clients = {}, items = {}, depts = {}, assignees = {};
    try { getClientMasterList().forEach(function(c) { if (!c.isFallback) clients[c.name] = { code: '', name: c.name }; }); } catch (e) {}
    getAllMgmtData().forEach(function(r) {
      var c = String(r[MGMT_COLS.CLIENT - 1] || '').trim(); if (c && !clients[c]) clients[c] = { code: '', name: c };
      var a = String(r[MGMT_COLS.ASSIGNEE - 1] || '').trim(); if (a) assignees[a] = true;
    });
    // 見積書シートの過去明細（後勝ち＝新しい単価）
    var qs = getSpreadsheet().getSheetByName(CONFIG.SHEET_QUOTES);
    if (qs && qs.getLastRow() > 1) {
      qs.getRange(2, 1, qs.getLastRow() - 1, QUOTE_COLS.REMARKS).getValues().forEach(function(r) {
        var name = String(r[QUOTE_COLS.ITEM_NAME - 1] || '').trim();
        if (!name) return;
        var code = (String(r[QUOTE_COLS.REMARKS - 1] || '').match(/品番:([^\s／]+)/) || [])[1] || '';
        items[code || name] = { code: code, name: name, spec: String(r[QUOTE_COLS.SPEC - 1] || ''), unit: String(r[QUOTE_COLS.UNIT - 1] || ''),
                                price: Number(r[QUOTE_COLS.UNIT_PRICE - 1]) || 0, date: _toDateStr(r[QUOTE_COLS.ISSUE_DATE - 1]) };
      });
    }
    // 見積入力の得意先コード・部門
    _qeRows().forEach(function(r) {
      try {
        var d = JSON.parse(String(r[QE_C.JSON] || '{}'));
        if (d.client) clients[d.client] = { code: d.clientCode || (clients[d.client] || {}).code || '', name: d.client };
        if (d.dept) depts[d.dept] = true;
        if (d.assignee) assignees[d.assignee] = true;
      } catch (e) {}
    });
    var company = {};
    try { company = JSON.parse(PropertiesService.getScriptProperties().getProperty(QE_COMPANY_PROP) || '{}'); } catch (e) {}
    return {
      success: true,
      clients: Object.keys(clients).map(function(k) { return clients[k]; }),
      items: Object.keys(items).map(function(k) { return items[k]; }).slice(-2000),
      depts: Object.keys(depts), assignees: Object.keys(assignees),
      company: company, me: _ofUser(),
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function apiQeSaveCompany(p) {
  try {
    var c = (p && p.company) || {};
    var keep = { name: c.name || '', zip: c.zip || '', address: c.address || '', tel: c.tel || '', fax: c.fax || '',
                 regNo: c.regNo || '', bank: c.bank || '' };
    PropertiesService.getScriptProperties().setProperty(QE_COMPANY_PROP, JSON.stringify(keep));
    return { success: true, company: keep };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
