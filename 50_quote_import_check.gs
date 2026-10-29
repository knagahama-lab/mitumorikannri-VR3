// ============================================================
// 50_quote_import_check.gs
// 「スプレッドシートには転記されているのに、見積書一覧に出ない」の確認用
//   最近登録された見積の行（管理シート）を新しい順に返し、一覧に出ない／見えにくい理由を付ける
//     ・見積番号が空                       → 一覧では「（番号なし）」で表示
//     ・同じ見積番号の新しい行がある       → 一覧には新しい方だけ表示
//     ・ステータスがボツ／旧見積           → アーカイブ扱い（「アーカイブを表示」で見える）
//     ・見積台帳の同じ番号がボツ／旧見積／キャンセル → 一覧ではそのステータスに置き換わる
//     ・取引先が空・自社                   → 取引先タブ（藤商事など）では出ない
// ============================================================

function apiQuoteImportCheck(p) {
  try {
    var days = Math.max(1, Math.min(90, Number((p && p.days) || 14)));
    var since = new Date(Date.now() - days * 86400000);
    var rows = getAllMgmtData();

    // 見積番号ごとの最新行（下の行ほど新しい）
    var latestByNo = {};
    rows.forEach(function (r, i) {
      var no = String(r[MGMT_COLS.QUOTE_NO - 1] || '').trim();
      if (no) latestByNo[no] = i;
    });
    // 見積台帳のステータス
    var ledgerSt = {};
    try {
      getAllLedgerData().forEach(function (lr) {
        var no = String(lr[LEDGER_COLS.QUOTE_NO - 1] || '').trim();
        if (no) ledgerSt[no] = String(lr[LEDGER_COLS.STATUS - 1] || '').trim();
      });
    } catch (e) {}
    var archived = (typeof ARCHIVE_STATUSES !== 'undefined') ? ARCHIVE_STATUSES : ['ボツ', '旧見積'];

    var out = [];
    rows.forEach(function (r, i) {
      var pdf = String(r[MGMT_COLS.QUOTE_PDF_URL - 1] || '').trim();
      var no = String(r[MGMT_COLS.QUOTE_NO - 1] || '').trim();
      if (!pdf && !no) return;                                   // 見積の行ではない
      if (!no && String(r[MGMT_COLS.ORDER_NO - 1] || '').trim()) return; // 注文だけの行
      var cv = r[MGMT_COLS.CREATED_AT - 1];
      var cd = (cv instanceof Date) ? cv : (cv ? new Date(String(cv).replace(/-/g, '/')) : null);
      if (!cd || isNaN(cd.getTime()) || cd < since) return;

      var status = String(r[MGMT_COLS.STATUS - 1] || '');
      var client = String(r[MGMT_COLS.CLIENT - 1] || '');
      var reasons = [];
      if (!no) reasons.push('見積番号が空（OCRで読めていません）');
      if (no && latestByNo[no] !== i) reasons.push('同じ見積番号の新しい行があるため、一覧には新しい方を表示');
      if (archived.indexOf(status) >= 0) reasons.push('ステータスが「' + status + '」（アーカイブ扱い）');
      var ls = no ? ledgerSt[no] : '';
      if (ls && (archived.indexOf(ls) >= 0 || ls === 'キャンセル')) reasons.push('見積台帳のステータスが「' + ls + '」');
      if (!client || (typeof _cpmIsSelf === 'function' && _cpmIsSelf(client))) reasons.push('取引先が空または自社（取引先タブでは出ません）');
      out.push({
        id: String(r[MGMT_COLS.ID - 1] || ''), quoteNo: no, subject: String(r[MGMT_COLS.SUBJECT - 1] || ''),
        client: client, status: status, ledgerStatus: ls || '',
        quoteDate: _toDateStr(r[MGMT_COLS.QUOTE_DATE - 1]), createdAt: Utilities.formatDate(cd, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
        pdfUrl: pdf, reasons: reasons,
      });
    });
    out.reverse(); // 新しい順
    return { success: true, days: days, rows: out };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
