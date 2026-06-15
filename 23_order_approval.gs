// ============================================================
// 23_order_approval.gs
// 営業承認ワークフロー
//  - 注文書単価確認後、資材購買課へ承認メールを送信
//  - 承認後にステータスを「受注済み」に更新
// ============================================================
//
// 【修正履歴】
//   - CONFIG.SHEETS.MGMT  → CONFIG.SHEET_MANAGEMENT  (01 config and setup.gs に合わせる)
//   - CONFIG.SHEETS.QUOTE → CONFIG.SHEET_QUOTES       (01 config and setup.gs に合わせる)
//   - ステータス更新ロジックの二重ループを削除・整理
// ============================================================

// ── メイン: 承認メール送信 ──────────────────────────────────
function _apiApproveOrderAndNotify(payload) {
  try {
    var mgmtId   = String(payload.mgmtId  || '');
    var toEmail  = String(payload.to      || '').trim();
    var ccEmail  = String(payload.cc      || '').trim();
    var subject  = String(payload.subject || '').trim();
    var memo     = String(payload.memo    || '').trim();
    var updateSt = payload.updateStatus !== false; // デフォルトtrue

    if (!mgmtId)  return { success: false, error: '管理IDが必要です' };
    if (!toEmail) return { success: false, error: '送信先メールアドレスを入力してください' };
    if (!subject) return { success: false, error: '件名を入力してください' };

    var approver  = Session.getActiveUser().getEmail() || '';
    var systemUrl = ScriptApp.getService().getUrl()    || '';

    // ── 注文書データをシートから取得 ──
    var ss    = getSpreadsheet();

    // ★ 修正: CONFIG.SHEETS.MGMT → CONFIG.SHEET_MANAGEMENT
    var sheet = ss.getSheetByName(CONFIG.SHEET_MANAGEMENT);
    if (!sheet) return { success: false, error: '管理シートが見つかりません' };

    var allData  = sheet.getDataRange().getValues();
    var orderRow = null;
    var orderRowIndex = -1; // ★ 追加: ステータス更新用に行番号を記憶
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][MGMT_COLS.ID - 1]) === mgmtId) {
        orderRow      = allData[i];
        orderRowIndex = i + 1; // スプレッドシートは1始まり、ヘッダー分+1
        break;
      }
    }
    if (!orderRow) return { success: false, error: '注文書が見つかりません: ' + mgmtId };

    var order = {
      id:           mgmtId,
      orderNo:      String(orderRow[MGMT_COLS.ORDER_NO      - 1] || ''),
      slipNo:       String(orderRow[MGMT_COLS.ORDER_SLIP_NO - 1] || ''),
      subject:      String(orderRow[MGMT_COLS.SUBJECT       - 1] || ''),
      client:       String(orderRow[MGMT_COLS.CLIENT        - 1] || ''),
      modelCode:    String(orderRow[MGMT_COLS.MODEL_CODE    - 1] || ''),
      orderDate:    _toDateStr(orderRow[MGMT_COLS.ORDER_DATE    - 1]),
      deliveryDate: _toDateStr(orderRow[MGMT_COLS.DELIVERY_DATE - 1]),
      orderAmount:  _toNum(orderRow[MGMT_COLS.ORDER_AMOUNT  - 1]),
      orderType:    String(orderRow[MGMT_COLS.ORDER_TYPE    - 1] || ''),
      orderPdfUrl:  String(orderRow[MGMT_COLS.ORDER_PDF_URL - 1] || ''),
      quotePdfUrl:  String(orderRow[MGMT_COLS.QUOTE_PDF_URL - 1] || ''),
      quoteNo:      String(orderRow[MGMT_COLS.QUOTE_NO      - 1] || ''),
    };

    // ── 紐づき見積書データ取得 ──
    var quote = null;
    if (order.quoteNo) {
      // ★ 修正: CONFIG.SHEETS.QUOTE → CONFIG.SHEET_QUOTES
      var qSheet = ss.getSheetByName(CONFIG.SHEET_QUOTES);
      if (qSheet) {
        var qData = qSheet.getDataRange().getValues();
        for (var j = 1; j < qData.length; j++) {
          var rowQNo = String(qData[j][QUOTE_COLS.QUOTE_NO - 1] || '').trim();
          if (rowQNo === order.quoteNo) {
            quote = {
              quoteNo:     rowQNo,
              issueDate:   _toDateStr(qData[j][QUOTE_COLS.ISSUE_DATE   - 1]),
              destCompany: String(qData[j][QUOTE_COLS.DEST_COMPANY - 1] || ''),
              pdfUrl:      String(qData[j][QUOTE_COLS.PDF_URL      - 1] || '') || order.quotePdfUrl,
            };
            break;
          }
        }
      }
    }

    // ── メール本文生成 ──
    var htmlBody  = _buildApprovalHtmlEmail(order, quote, memo, approver, systemUrl);
    var plainBody = _buildApprovalPlainEmail(order, quote, memo, approver, systemUrl);

    // ── メール送信 ──
    var opts = { name: '見積・注文管理システム（営業）', htmlBody: htmlBody };
    if (ccEmail) opts.cc = ccEmail;
    GmailApp.sendEmail(toEmail, subject, plainBody, opts);
    Logger.log('[APPROVAL] 承認メール送信: ' + toEmail + ' | ' + mgmtId);

    // ── ステータス更新（受注済み）──
    // ★ 修正: 二重ループを削除し、取得済みの orderRowIndex を使用
    if (updateSt && orderRowIndex > 0) {
      sheet.getRange(orderRowIndex, MGMT_COLS.STATUS).setValue('受注済み');
      Logger.log('[APPROVAL] ステータス更新: 行' + orderRowIndex + ' → 受注済み');
    }

    return { success: true, sentTo: toEmail, mgmtId: mgmtId };

  } catch(e) {
    Logger.log('[APPROVAL] エラー: ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}

// ── 資材購買メールアドレスを返す ────────────────────────────
function _apiGetApprovalSettings() {
  try {
    var props = PropertiesService.getScriptProperties();
    return {
      success:          true,
      procurementEmail: props.getProperty('PROCUREMENT_EMAIL') || '',
      salesEmails:      props.getProperty('SALES_EMAILS')      || '',
      approvalButtonEnabled: props.getProperty('APPROVAL_BUTTON_ENABLED') !== 'false',
    };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ── 承認ボタンの表示/非表示などの設定を保存 ────────────────────
function _apiSaveApprovalSettings(p) {
  try {
    var props = PropertiesService.getScriptProperties();
    if (p.approvalButtonEnabled !== undefined) {
      props.setProperty('APPROVAL_BUTTON_ENABLED', p.approvalButtonEnabled ? 'true' : 'false');
    }
    if (p.procurementEmail !== undefined) {
      props.setProperty('PROCUREMENT_EMAIL', p.procurementEmail || '');
    }
    if (p.salesEmails !== undefined) {
      props.setProperty('SALES_EMAILS', p.salesEmails || '');
    }
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// HTML メール本文
// ============================================================
function _buildApprovalHtmlEmail(order, quote, memo, approver, systemUrl) {
  function row(label, val, accent) {
    return '<tr>'
      + '<th style="text-align:left;padding:8px 14px;background:' + (accent ? '#f0fdf4' : '#f8fafc') + ';border:1px solid #e2e8f0;white-space:nowrap;width:130px;font-weight:600;color:#475569;font-size:13px;">'
      + label + '</th>'
      + '<td style="padding:8px 14px;border:1px solid #e2e8f0;color:#1e293b;font-size:13px;">' + (val || '—') + '</td>'
      + '</tr>';
  }

  var orderRows = ''
    + row('発注書番号',  order.orderNo)
    + row('発注伝票No',  order.slipNo)
    + row('件名',        order.subject)
    + row('顧客名',      order.client)
    + row('機種',        order.modelCode)
    + row('種別',        order.orderType)
    + row('発注日',      order.orderDate)
    + row('納期',        order.deliveryDate)
    + row('注文金額',    order.orderAmount ? '&yen;' + Number(order.orderAmount).toLocaleString() : '—', true);

  var orderPdfBtn = order.orderPdfUrl
    ? '<div style="margin-top:12px;"><a href="' + order.orderPdfUrl + '" target="_blank" '
      + 'style="display:inline-block;background:#0f766e;color:#fff;text-decoration:none;padding:9px 20px;border-radius:6px;font-weight:700;font-size:13px;">&#128196; 注文書PDF を開く</a></div>'
    : '';

  var quoteSection = '';
  if (quote) {
    var quoteRows = ''
      + row('見積書番号', quote.quoteNo)
      + row('発行日',     quote.issueDate)
      + row('宛先',       quote.destCompany);
    var quotePdfBtn = quote.pdfUrl
      ? '<div style="margin-top:12px;"><a href="' + quote.pdfUrl + '" target="_blank" '
        + 'style="display:inline-block;background:#2563eb;color:#fff;text-decoration:none;padding:9px 20px;border-radius:6px;font-weight:700;font-size:13px;">&#128196; 見積書PDF を開く</a></div>'
      : '';
    quoteSection = '<div style="margin-top:24px;">'
      + '<h3 style="font-size:14px;font-weight:700;color:#2563eb;margin:0 0 10px;border-bottom:2px solid #dbeafe;padding-bottom:6px;">&#128203; 紐づき見積書</h3>'
      + '<table style="border-collapse:collapse;width:100%;">' + quoteRows + '</table>'
      + quotePdfBtn
      + '</div>';
  }

  var memoSection = memo
    ? '<div style="margin-top:20px;background:#fefce8;border:1px solid #fef08a;border-radius:8px;padding:14px;">'
      + '<strong style="color:#854d0e;font-size:13px;">&#128221; 備考</strong>'
      + '<p style="margin:6px 0 0;color:#713f12;font-size:13px;">' + memo.replace(/\n/g, '<br>') + '</p>'
      + '</div>'
    : '';

  var systemBtn = systemUrl
    ? '<a href="' + systemUrl + '" target="_blank" style="display:inline-block;background:#6366f1;color:#fff;text-decoration:none;padding:5px 14px;border-radius:5px;font-size:11px;font-weight:600;margin-left:10px;">&#128279; システムを開く</a>'
    : '';

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
    + '<body style="font-family:\'Helvetica Neue\',Arial,\'Hiragino Kaku Gothic ProN\',sans-serif;font-size:14px;color:#334155;margin:0;padding:0;background:#f1f5f9;">'
    + '<div style="max-width:660px;margin:24px auto;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.12);">'

    // ヘッダー
    + '<div style="background:linear-gradient(135deg,#16a34a 0%,#15803d 100%);color:#fff;padding:22px 28px;">'
    + '<div style="font-size:20px;font-weight:800;letter-spacing:.02em;">&#9989; 注文書 承認・資材手配依頼</div>'
    + '<div style="font-size:12px;margin-top:5px;opacity:.85;">単価確認完了 ― 販売管理登録・手配をお願いします</div>'
    + '</div>'

    // 本文
    + '<div style="background:#fff;padding:26px 28px;">'
    + '<p style="margin:0 0 20px;line-height:1.8;color:#475569;">資材購買担当者様<br><br>'
    + '下記注文書の単価確認が完了しました。<br>'
    + '販売管理への登録および手配をお願いいたします。</p>'

    + '<h3 style="font-size:14px;font-weight:700;color:#0f766e;margin:0 0 10px;border-bottom:2px solid #ccfbf1;padding-bottom:6px;">&#128203; 注文書情報</h3>'
    + '<table style="border-collapse:collapse;width:100%;">' + orderRows + '</table>'
    + orderPdfBtn
    + quoteSection
    + memoSection
    + '</div>'

    // フッター
    + '<div style="background:#f8fafc;padding:14px 28px;border-top:1px solid #e2e8f0;display:flex;align-items:center;flex-wrap:wrap;gap:4px;">'
    + '<span style="font-size:11px;color:#94a3b8;">送信者: ' + approver + '</span>'
    + systemBtn
    + '</div>'
    + '</div></body></html>';
}

// ============================================================
// プレーンテキスト本文（フォールバック）
// ============================================================
function _buildApprovalPlainEmail(order, quote, memo, approver, systemUrl) {
  var lines = [
    '資材購買担当者様',
    '',
    '下記注文書の単価確認が完了しました。',
    '販売管理への登録および手配をお願いいたします。',
    '',
    '--- 注文書情報 ---',
    '発注書番号 : ' + (order.orderNo      || '—'),
    '発注伝票No : ' + (order.slipNo        || '—'),
    '件名       : ' + (order.subject       || '—'),
    '顧客名     : ' + (order.client        || '—'),
    '機種       : ' + (order.modelCode     || '—'),
    '種別       : ' + (order.orderType     || '—'),
    '発注日     : ' + (order.orderDate     || '—'),
    '納期       : ' + (order.deliveryDate  || '—'),
    '注文金額   : ' + (order.orderAmount   ? '¥' + Number(order.orderAmount).toLocaleString() : '—'),
  ];
  if (order.orderPdfUrl) lines.push('注文書PDF  : ' + order.orderPdfUrl);

  if (quote) {
    lines.push('', '--- 紐づき見積書 ---');
    lines.push('見積書番号 : ' + (quote.quoteNo     || '—'));
    lines.push('発行日     : ' + (quote.issueDate   || '—'));
    lines.push('宛先       : ' + (quote.destCompany || '—'));
    if (quote.pdfUrl) lines.push('見積書PDF  : ' + quote.pdfUrl);
  }

  if (memo) { lines.push('', '--- 備考 ---', memo); }

  lines.push('', '---');
  lines.push('送信者 : ' + approver);
  if (systemUrl) lines.push('システム : ' + systemUrl);
  return lines.join('\n');
}