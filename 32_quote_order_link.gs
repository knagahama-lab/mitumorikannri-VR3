// ============================================================
// 32_quote_order_link.gs
// 見積書 ⇔ 注文書 の対応を PDF を開かずに一目で分かるようにする
//
//   apiQuoteOrderLinks()          … 見積ごとの受注状況（注文No.・受領日・経過日数・候補）
//   apiQuoteOrderCalendar({year,month}) … 見積提出日／注文受領日をペアで返す（カレンダー用）
//
// 紐づけの判定元（どれか1つでも一致すれば「受注済み」とみなす）
//   A. 注文側の管理行に見積番号が入っている（MGMT_COLS.QUOTE_NO）
//   B. 見積側の管理行に注文番号が入っている（MGMT_COLS.ORDER_NO）
//   C. 注文書シートの「見積番号(紐づけ)」列（ORDER_COLS.LINKED_QUOTE）
// 未紐づけの見積には、機種コード＋基板名/件名が一致する注文を「候補」として返す（確定はしない）。
// 見積の提出日は 見積台帳の「メール送信日」→ 発行日 → 管理シートの見積日 の順で採用。
// ============================================================

function _qolDays(from, to) {
  var a = new Date(String(from).replace(/-/g, '/'));
  var b = new Date(String(to).replace(/-/g, '/'));
  if (isNaN(a.getTime()) || isNaN(b.getTime())) return null;
  return Math.round((b.getTime() - a.getTime()) / 86400000);
}

function _qolNorm(s) {
  return normalizeText(s).replace(/[\s\-_・（）()【】\[\]]/g, '');
}

/** 見積・注文・紐づけ情報をまとめて構築する */
function _qolBuild() {
  var mgmt = getAllMgmtData().map(_rowToObject).filter(function(o) {
    return o.isLatest === '' || String(o.isLatest).toUpperCase() === 'TRUE';
  });

  // 注文（注文番号を持つ行）
  var orders = {}; // orderNo -> order（注文書PDFを持つ行＝注文本体を優先）
  mgmt.slice().sort(function(a, b) { return (b.orderPdfUrl ? 1 : 0) - (a.orderPdfUrl ? 1 : 0); }).forEach(function(o) {
    var no = String(o.orderNo || '').trim();
    if (!no || orders[no]) return;
    orders[no] = {
      mgmtId: o.id, orderNo: no, orderDate: o.orderDate, client: o.client, subject: o.subject,
      modelCode: o.modelCode, orderAmount: o.orderAmount, orderPdfUrl: o.orderPdfUrl,
      orderType: o.orderType, deliveryDate: o.deliveryDate, quoteNos: {},
    };
  });

  // 見積（見積番号を持つ行）
  var quotes = {}; // quoteNo -> quote
  mgmt.forEach(function(o) {
    var no = String(o.quoteNo || '').trim();
    if (!no) return;
    // 注文側の行（注文番号あり＆見積PDFなし）は見積本体ではなく紐づけ情報として扱う
    var isOrderRow = o.orderNo && !o.quotePdfUrl && o.orderPdfUrl;
    if (isOrderRow) {
      orders[o.orderNo] && (orders[o.orderNo].quoteNos[no] = 'A');
      return;
    }
    if (!quotes[no]) {
      quotes[no] = {
        mgmtId: o.id, quoteNo: no, subject: o.subject, client: o.client, modelCode: o.modelCode,
        boardName: '', quoteDate: o.quoteDate, submitDate: '', quoteAmount: o.quoteAmount,
        quotePdfUrl: o.quotePdfUrl, status: o.status, orderNos: {},
      };
    }
    if (o.orderNo) {
      quotes[no].orderNos[o.orderNo] = 'B';
      if (orders[o.orderNo]) orders[o.orderNo].quoteNos[no] = 'B';
    }
  });
  // 注文行が見積番号を持つケース（A）を見積側へ反映
  Object.keys(orders).forEach(function(ono) {
    Object.keys(orders[ono].quoteNos).forEach(function(qno) {
      if (quotes[qno]) quotes[qno].orderNos[ono] = quotes[qno].orderNos[ono] || 'A';
    });
  });

  var ss = getSpreadsheet();
  // C. 注文書シートの紐づけ列
  var os = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  if (os && os.getLastRow() > 1) {
    os.getRange(2, 1, os.getLastRow() - 1, ORDER_COLS.LINKED_QUOTE).getValues().forEach(function(r) {
      var ono = String(r[ORDER_COLS.ORDER_NO - 1] || '').trim();
      var qno = String(r[ORDER_COLS.LINKED_QUOTE - 1] || '').trim();
      if (!ono || !qno) return;
      if (orders[ono]) orders[ono].quoteNos[qno] = orders[ono].quoteNos[qno] || 'C';
      if (quotes[qno]) quotes[qno].orderNos[ono] = quotes[qno].orderNos[ono] || 'C';
    });
  }

  // 見積台帳：提出日（メール送信日）・基板名・台帳のみの見積
  try {
    getAllLedgerData().forEach(function(r) {
      var qno    = String(r[LEDGER_COLS.QUOTE_NO - 1] || '').trim();
      var status = String(r[LEDGER_COLS.STATUS - 1] || '').trim();
      if (!qno || status === 'キャンセル' || status === '__MACHINE_FOLDER__') return;
      var sent  = _toDateStr(r[LEDGER_COLS.SENT_DATE - 1]);
      var issue = _toDateStr(r[LEDGER_COLS.ISSUE_DATE - 1]);
      if (!quotes[qno]) {
        quotes[qno] = {
          mgmtId: '', ledgerId: String(r[LEDGER_COLS.LEDGER_ID - 1] || ''), quoteNo: qno,
          subject: String(r[LEDGER_COLS.SUBJECT - 1] || ''), client: String(r[LEDGER_COLS.DEST - 1] || ''),
          modelCode: String(r[LEDGER_COLS.MACHINE_CODE - 1] || ''), boardName: '',
          quoteDate: issue, submitDate: '', quoteAmount: _toNum(r[LEDGER_COLS.AMOUNT - 1]),
          quotePdfUrl: String(r[LEDGER_COLS.SAVE_URL - 1] || ''), status: status, orderNos: {},
        };
        Object.keys(orders).forEach(function(ono) {
          if (orders[ono].quoteNos[qno]) quotes[qno].orderNos[ono] = orders[ono].quoteNos[qno];
        });
      }
      var q = quotes[qno];
      if (sent && !q.submitDate) q.submitDate = sent;
      if (!q.quoteDate && issue) q.quoteDate = issue;
      if (!q.boardName) q.boardName = String(r[LEDGER_COLS.BOARD_NAME - 1] || '');
      if (!q.modelCode) q.modelCode = String(r[LEDGER_COLS.MACHINE_CODE - 1] || '');
    });
  } catch (e) { Logger.log('[_qolBuild ledger] ' + e.message); }

  // D. 客先部品コード：見積に登録した部品コードが、見積提出日以降の注文明細に出てきたら受注とみなす
  try {
    if (typeof _cpmRows === 'function' && typeof _cpmOrderLines === 'function') {
      var codesByQuote = {};
      _cpmRows().forEach(function(r) {
        var qno = String(r[CPM.QUOTE_NO] || '').trim();
        if (qno && quotes[qno]) (codesByQuote[qno] = codesByQuote[qno] || []).push(r[CPM.CLIENT] + '|' + _cpmCode(r[CPM.CODE]));
      });
      if (Object.keys(codesByQuote).length) {
        var linesByKey = {};
        _cpmOrderLines().forEach(function(l) {
          if (l.code && l.orderNo) (linesByKey[l.clientKey + '|' + l.code] = linesByKey[l.clientKey + '|' + l.code] || []).push(l);
        });
        Object.keys(codesByQuote).forEach(function(qno) {
          var q = quotes[qno];
          var base = String(q.submitDate || q.quoteDate || '');
          codesByQuote[qno].forEach(function(key) {
            (linesByKey[key] || []).forEach(function(l) {
              if (base && l.orderDate && String(l.orderDate) < base) return; // 見積より前の注文は対象外
              q.orderNos[l.orderNo] = q.orderNos[l.orderNo] || 'D';
              if (orders[l.orderNo]) orders[l.orderNo].quoteNos[qno] = orders[l.orderNo].quoteNos[qno] || 'D';
            });
          });
        });
      }
    }
  } catch (e) { Logger.log('[_qolBuild partcode] ' + e.message); }

  return { quotes: quotes, orders: orders };
}

/** 未紐づけ見積に対する注文候補（機種コード一致 かつ 基板名 or 件名の一部一致、見積日以降の注文） */
function _qolCandidates(q, orders, linkedOrderNos) {
  var codes = String(q.modelCode || '').split(/[,、\/\s]+/).map(_qolNorm).filter(Boolean);
  if (!codes.length) return [];
  var board = _qolNorm(q.boardName);
  var subj  = _qolNorm(q.subject).substring(0, 12);
  var base  = q.submitDate || q.quoteDate;
  var out = [];
  Object.keys(orders).forEach(function(ono) {
    if (linkedOrderNos[ono]) return;
    var o = orders[ono];
    var oc = String(o.modelCode || '').split(/[,、\/\s]+/).map(_qolNorm).filter(Boolean);
    if (!oc.some(function(c) { return codes.indexOf(c) >= 0; })) return;
    if (base && o.orderDate && _qolDays(base, o.orderDate) < 0) return;
    var os = _qolNorm(o.subject);
    if ((board && os.indexOf(board) >= 0) || (subj && os.indexOf(subj) >= 0) || (!board && !subj)) {
      out.push({ orderNo: ono, orderDate: o.orderDate, mgmtId: o.mgmtId, subject: o.subject });
    }
  });
  return out.slice(0, 5);
}

function _qolSummary(q, orders, linkedOrderNos, today) {
  var list = Object.keys(q.orderNos).map(function(ono) {
    var o = orders[ono] || { orderNo: ono };
    return {
      orderNo: ono, mgmtId: o.mgmtId || '', orderDate: o.orderDate || '', orderPdfUrl: o.orderPdfUrl || '',
      orderAmount: o.orderAmount || 0, orderType: o.orderType || '', via: q.orderNos[ono],
    };
  }).sort(function(a, b) { return String(a.orderDate).localeCompare(String(b.orderDate)); });
  var submit = q.submitDate || q.quoteDate || '';
  var first  = list.length ? list[0].orderDate : '';
  var cancelled = q.status === 'キャンセル' || q.status === '失注';
  return {
    quoteMgmtId: q.mgmtId, ledgerId: q.ledgerId || '', quoteNo: q.quoteNo, subject: q.subject, client: q.client,
    modelCode: q.modelCode, boardName: q.boardName, quoteAmount: q.quoteAmount, quotePdfUrl: q.quotePdfUrl,
    submitDate: submit, submitFromLedger: !!q.submitDate, status: q.status,
    orders: list,
    state: list.length ? 'ordered' : cancelled ? 'lost' : 'waiting',
    leadDays: (submit && first) ? _qolDays(submit, first) : null,             // 提出→受注の日数
    waitingDays: (!list.length && submit) ? _qolDays(submit, today) : null,   // 注文待ち経過日数
    candidates: (!list.length && !cancelled) ? _qolCandidates(q, orders, linkedOrderNos) : [],
  };
}

function apiQuoteOrderLinks() {
  try {
    var b = _qolBuild();
    var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
    var linked = {};
    Object.keys(b.quotes).forEach(function(qno) { Object.keys(b.quotes[qno].orderNos).forEach(function(o) { linked[o] = true; }); });

    var byQuote = {};
    Object.keys(b.quotes).forEach(function(qno) { byQuote[qno] = _qolSummary(b.quotes[qno], b.orders, linked, today); });

    // 注文側から見た「元の見積」
    var byOrder = {};
    Object.keys(b.orders).forEach(function(ono) {
      var o = b.orders[ono];
      byOrder[ono] = {
        mgmtId: o.mgmtId, orderDate: o.orderDate, client: o.client, subject: o.subject,
        modelCode: o.modelCode, orderPdfUrl: o.orderPdfUrl, quotes: [],
      };
      byOrder[ono].quotes = Object.keys(o.quoteNos).map(function(qno) {
        var q = b.quotes[qno] || {};
        var submit = q.submitDate || q.quoteDate || '';
        return {
          quoteNo: qno, quoteMgmtId: q.mgmtId || '', submitDate: submit, quotePdfUrl: q.quotePdfUrl || '',
          leadDays: (submit && b.orders[ono].orderDate) ? _qolDays(submit, b.orders[ono].orderDate) : null,
        };
      });
    });
    return { success: true, today: today, byQuote: byQuote, byOrder: byOrder };
  } catch (e) {
    Logger.log('[apiQuoteOrderLinks] ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

/**
 * 指定月に「見積提出」または「注文受領」があるペアを返す。
 * pairs: [{ quoteNo, submitDate, orderNo, orderDate, leadDays, state, ... }]
 */
function apiQuoteOrderCalendar(p) {
  try {
    var year  = Number(p && p.year)  || new Date().getFullYear();
    var month = Number(p && p.month) || (new Date().getMonth() + 1);
    var ym    = year + '/' + ('0' + month).slice(-2);
    var links = apiQuoteOrderLinks();
    if (!links.success) return links;
    var pairs = [];
    Object.keys(links.byQuote).forEach(function(qno) {
      var s = links.byQuote[qno];
      var base = {
        quoteNo: qno, quoteMgmtId: s.quoteMgmtId, ledgerId: s.ledgerId, subject: s.subject, client: s.client,
        modelCode: s.modelCode, submitDate: s.submitDate, quotePdfUrl: s.quotePdfUrl, state: s.state,
        waitingDays: s.waitingDays, candidates: s.candidates.length,
      };
      if (!s.orders.length) {
        if (String(s.submitDate).indexOf(ym) === 0) pairs.push(Object.assign({}, base, { orderNo: '', orderDate: '' }));
        return;
      }
      s.orders.forEach(function(o) {
        if (String(s.submitDate).indexOf(ym) !== 0 && String(o.orderDate).indexOf(ym) !== 0) return;
        pairs.push(Object.assign({}, base, {
          orderNo: o.orderNo, orderMgmtId: o.mgmtId, orderDate: o.orderDate, orderPdfUrl: o.orderPdfUrl,
          leadDays: (s.submitDate && o.orderDate) ? _qolDays(s.submitDate, o.orderDate) : null,
        }));
      });
    });
    // 見積に紐づかない注文（見積不明）も受領日に表示する
    Object.keys(links.byOrder).forEach(function(ono) {
      var o = links.byOrder[ono];
      if (o.quotes.length || String(o.orderDate).indexOf(ym) !== 0) return;
      pairs.push({
        quoteNo: '', submitDate: '', subject: o.subject, client: o.client, modelCode: o.modelCode,
        state: 'noquote', orderNo: ono, orderMgmtId: o.mgmtId, orderDate: o.orderDate, orderPdfUrl: o.orderPdfUrl,
      });
    });
    pairs.sort(function(a, b) {
      return String(a.submitDate || a.orderDate).localeCompare(String(b.submitDate || b.orderDate));
    });
    return { success: true, year: year, month: month, today: links.today, pairs: pairs };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
