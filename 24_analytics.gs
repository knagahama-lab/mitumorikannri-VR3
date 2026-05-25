// ============================================================
// 24_analytics.gs
// 分析・レポート機能
//  - 月別・年別の見積/注文集計
//  - ステータス・顧客・機種 ランキング
// ============================================================

function apiGetAnalysisReport(payload) {
  try {
    payload = payload || {};
    var year = parseInt(payload.year || new Date().getFullYear(), 10);

    var mgmtData = getAllMgmtData();

    // 月別集計マップ初期化
    var monthlyQuotes = {};
    var monthlyOrders = {};
    for (var m = 1; m <= 12; m++) {
      var key = year + '-' + (m < 10 ? '0' : '') + m;
      monthlyQuotes[key] = { count: 0, amount: 0 };
      monthlyOrders[key] = { count: 0, amount: 0 };
    }

    var totalQuotes   = 0, totalOrders   = 0;
    var totalQuoteAmt = 0, totalOrderAmt = 0;
    var statusCount   = {};
    var clientCount   = {};
    var modelCount    = {};
    var yearsSet      = {};

    mgmtData.forEach(function(r) {
      var qDate  = _toDateStr(r[MGMT_COLS.QUOTE_DATE   - 1]);
      var oDate  = _toDateStr(r[MGMT_COLS.ORDER_DATE   - 1]);
      var qNo    = String(r[MGMT_COLS.QUOTE_NO          - 1] || '').trim();
      var oNo    = String(r[MGMT_COLS.ORDER_NO          - 1] || '').trim();
      var qAmt   = _toNum(r[MGMT_COLS.QUOTE_AMOUNT      - 1]);
      var oAmt   = _toNum(r[MGMT_COLS.ORDER_AMOUNT      - 1]);
      var status = String(r[MGMT_COLS.STATUS            - 1] || '').trim();
      var client = String(r[MGMT_COLS.CLIENT            - 1] || '').trim();
      var mc     = String(r[MGMT_COLS.MODEL_CODE        - 1] || '').trim();

      // 利用可能な年を収集
      if (qDate) yearsSet[qDate.substring(0, 4)] = true;
      if (oDate) yearsSet[oDate.substring(0, 4)] = true;

      // ステータス別（全年）
      if (status) {
        statusCount[status] = (statusCount[status] || 0) + 1;
      }

      // 見積集計（対象年）
      if (qNo && qDate && parseInt(qDate.substring(0, 4), 10) === year) {
        var qKey = qDate.substring(0, 7);
        totalQuotes++;
        totalQuoteAmt += qAmt;
        if (monthlyQuotes[qKey]) {
          monthlyQuotes[qKey].count++;
          monthlyQuotes[qKey].amount += qAmt;
        }
      }

      // 注文集計（対象年）
      if (oNo && oDate && parseInt(oDate.substring(0, 4), 10) === year) {
        var oKey = oDate.substring(0, 7);
        totalOrders++;
        totalOrderAmt += oAmt;
        if (monthlyOrders[oKey]) {
          monthlyOrders[oKey].count++;
          monthlyOrders[oKey].amount += oAmt;
        }
      }

      // 顧客別・機種別（対象年）
      var refDate = qDate || oDate;
      if (refDate && parseInt(refDate.substring(0, 4), 10) === year) {
        if (client) {
          clientCount[client] = (clientCount[client] || 0) + 1;
        }
        if (mc) {
          if (!modelCount[mc]) modelCount[mc] = { quotes: 0, orders: 0, amount: 0 };
          if (qNo) modelCount[mc].quotes++;
          if (oNo) { modelCount[mc].orders++; modelCount[mc].amount += oAmt; }
        }
      }
    });

    // 月別配列
    var months = [];
    for (var m2 = 1; m2 <= 12; m2++) {
      var key2 = year + '-' + (m2 < 10 ? '0' : '') + m2;
      months.push({
        month:    m2,
        label:    m2 + '月',
        quotes:   monthlyQuotes[key2].count,
        quoteAmt: monthlyQuotes[key2].amount,
        orders:   monthlyOrders[key2].count,
        orderAmt: monthlyOrders[key2].amount,
      });
    }

    // 顧客ランキング TOP10
    var clientRank = Object.keys(clientCount).map(function(c) {
      return { name: c, count: clientCount[c] };
    }).sort(function(a, b) { return b.count - a.count; }).slice(0, 10);

    // 機種ランキング TOP10（見積件数順）
    var modelRank = Object.keys(modelCount).map(function(mc) {
      return { code: mc, quotes: modelCount[mc].quotes, orders: modelCount[mc].orders, amount: modelCount[mc].amount };
    }).sort(function(a, b) { return b.quotes - a.quotes; }).slice(0, 10);

    // ステータス配列（件数降順）
    var statusRank = Object.keys(statusCount).map(function(s) {
      return { status: s, count: statusCount[s] };
    }).sort(function(a, b) { return b.count - a.count; });

    // 利用可能な年リスト（降順）
    var availableYears = Object.keys(yearsSet).map(Number).sort(function(a, b) { return b - a; });

    return JSON.parse(JSON.stringify({
      success:        true,
      year:           year,
      totalQuotes:    totalQuotes,
      totalOrders:    totalOrders,
      totalQuoteAmt:  totalQuoteAmt,
      totalOrderAmt:  totalOrderAmt,
      months:         months,
      statusRank:     statusRank,
      clientRank:     clientRank,
      modelRank:      modelRank,
      availableYears: availableYears,
    }));
  } catch(e) {
    Logger.log('[apiGetAnalysisReport ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}
