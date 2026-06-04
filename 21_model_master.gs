// ============================================================
// 21_model_master.gs
// 機種マスタ管理 + 全データ双方向連携
//
// 【機能】
//   ② 機種マスタ CRUD（機種別 見積書・注文書 管理・閲覧）
//   ③ 機種マスタ ↔ 基板マスタ(BOM) ↔ 見積書一覧 ↔ 見積台帳 双方向同期
//
// 【シート】
//   主SS: '機種マスタ' シート（本ファイルで管理）
//   BOM SS: '機種マスタ' シート（08_parts_management.gs の BOM_SS_ID）
// ============================================================

var MODEL_MASTER_SHEET = '機種マスタ';
var MODEL_MASTER_COLS  = {
  ID          : 1,  // 機種ID (MM-YYYYMMDD-XXXX)
  MODEL_CODE  : 2,  // 機種コード（全シート共通リンクキー）
  MODEL_NAME  : 3,  // 機種名
  DESCRIPTION : 4,  // 説明・備考
  BOARD_NAMES : 5,  // 基板名一覧（BOM SSから自動取得）
  CREATED_AT  : 6,  // 登録日時
  UPDATED_AT  : 7,  // 更新日時
  MODEL_TYPE  : 8,  // 型式
  RELEASE_DATE: 9,  // 発売日
  VISIBLE     : 10, // 表示/非表示（TRUE=表示 / FALSE=非表示）
};
var MODEL_MASTER_COL_COUNT = 10;

// ============================================================
// 初期化
// ============================================================

function initModelMasterSheet() {
  try {
    var ss      = getSpreadsheet();
    var headers = ['機種ID','機種コード','機種名','説明・備考','基板名一覧(自動)','登録日時','更新日時','型式','発売日','表示'];
    var sheet   = _createOrSetupSheet(ss, MODEL_MASTER_SHEET, headers, '#E3F2FD');
    // 基板名列は自動列なので薄いグレー
    if (sheet.getLastRow() <= 1) {
      sheet.getRange(2, MODEL_MASTER_COLS.BOARD_NAMES, 1000, 1).setBackground('#f5f5f5');
    }
    Logger.log('[initModelMasterSheet] 機種マスタシート初期化完了');
    return { success: true };
  } catch(e) {
    Logger.log('[initModelMasterSheet ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// 内部ユーティリティ
// ============================================================

function _getModelMasterSheet() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(MODEL_MASTER_SHEET);
  if (!sheet) {
    initModelMasterSheet();
    sheet = ss.getSheetByName(MODEL_MASTER_SHEET);
  }
  return sheet;
}

function _getAllModelMasterRows() {
  var sheet = _getModelMasterSheet();
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var colCount = sheet.getLastColumn() >= MODEL_MASTER_COL_COUNT ? MODEL_MASTER_COL_COUNT : 7;
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, colCount).getValues()
    .filter(function(r) { return String(r[MODEL_MASTER_COLS.MODEL_CODE - 1]).trim() !== ''; });
}

function _findModelRowNum(modelCode) {
  var sheet = _getModelMasterSheet();
  if (!sheet || sheet.getLastRow() <= 1) return -1;
  var codes = sheet.getRange(2, MODEL_MASTER_COLS.MODEL_CODE, sheet.getLastRow() - 1, 1)
                   .getValues().flat();
  var idx = codes.findIndex(function(c) {
    return String(c).trim() === String(modelCode).trim();
  });
  return idx >= 0 ? idx + 2 : -1; // 1-based
}

function _generateModelId() {
  return 'MM-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd') + '-' +
    (Math.floor(Math.random() * 9000) + 1000);
}

// ============================================================
// API: 機種マスタ一覧（統計付き）
// ============================================================

function apiModelMasterList() {
  try {
    // ★ 複数機種コード対応の統計集計ヘルパー
    // 管理シートのモデルコード（単一）を、機種マスタの複数コードエントリにマッピング
    var mgmtData = getAllMgmtData();
    // まず管理シートの単一コード別に統計を集計
    var singleStatsMap = {};
    mgmtData.forEach(function(r) {
      var mc = String(r[MGMT_COLS.MODEL_CODE - 1] || '').trim();
      if (!mc) return;
      if (!singleStatsMap[mc]) {
        singleStatsMap[mc] = { quoteCount: 0, orderCount: 0, totalAmount: 0, latestDate: '' };
      }
      if (String(r[MGMT_COLS.QUOTE_NO - 1] || '').trim()) singleStatsMap[mc].quoteCount++;
      if (String(r[MGMT_COLS.ORDER_NO  - 1] || '').trim()) {
        singleStatsMap[mc].orderCount++;
        singleStatsMap[mc].totalAmount += _toNum(r[MGMT_COLS.ORDER_AMOUNT - 1]);
      }
      var d = _toDateStr(r[MGMT_COLS.ORDER_DATE - 1] || r[MGMT_COLS.QUOTE_DATE - 1]);
      if (d > singleStatsMap[mc].latestDate) singleStatsMap[mc].latestDate = d;
    });
    // 台帳の単一コード別集計も追加
    var ledgerData = getAllLedgerData();
    ledgerData.forEach(function(r) {
      var mc = String(r[LEDGER_COLS.MACHINE_CODE - 1] || '').trim();
      if (mc && !singleStatsMap[mc]) {
        singleStatsMap[mc] = { quoteCount: 0, orderCount: 0, totalAmount: 0, latestDate: '' };
      }
    });

    // 機種マスタ行の複数コード（カンマ区切り）に対して統計を合算する関数
    function _mergeStats(modelCodeField) {
      var codes = String(modelCodeField || '').split(/[,、\s]+/).map(function(c){ return c.trim(); }).filter(Boolean);
      var merged = { quoteCount: 0, orderCount: 0, totalAmount: 0, latestDate: '' };
      codes.forEach(function(c) {
        var s = singleStatsMap[c];
        if (!s) return;
        merged.quoteCount  += s.quoteCount;
        merged.orderCount  += s.orderCount;
        merged.totalAmount += s.totalAmount;
        if (s.latestDate > merged.latestDate) merged.latestDate = s.latestDate;
      });
      // 単一コードとしてもチェック（完全一致フォールバック）
      var exact = singleStatsMap[String(modelCodeField || '').trim()];
      if (exact && merged.quoteCount === 0 && merged.orderCount === 0) return exact;
      return merged;
    }

    // ※ 自動追加は廃止 — 削除した機種コードが復活するバグを防ぐため
    // （新規登録時は onDocRegistered() が _ensureModelCode() を呼ぶ）

    // 最新状態で再取得してレスポンス構築
    var rows  = _getAllModelMasterRows();
    var items = rows.map(function(r) {
      var mc      = String(r[MODEL_MASTER_COLS.MODEL_CODE  - 1] || '').trim();
      var stats   = _mergeStats(mc);
      var visible = r[MODEL_MASTER_COLS.VISIBLE - 1];
      // VISIBLE が明示的に FALSE のものは非表示（空欄・TRUE は表示）
      var isVisible = (visible === false || String(visible).toUpperCase() === 'FALSE') ? false : true;
      return {
        id          : String(r[MODEL_MASTER_COLS.ID          - 1] || ''),
        modelCode   : mc,
        modelName   : String(r[MODEL_MASTER_COLS.MODEL_NAME  - 1] || ''),
        description : String(r[MODEL_MASTER_COLS.DESCRIPTION - 1] || ''),
        boardNames  : String(r[MODEL_MASTER_COLS.BOARD_NAMES - 1] || ''),
        createdAt   : _toDateStr(r[MODEL_MASTER_COLS.CREATED_AT - 1]),
        updatedAt   : _toDateStr(r[MODEL_MASTER_COLS.UPDATED_AT - 1]),
        modelType   : String(r[MODEL_MASTER_COLS.MODEL_TYPE   - 1] || ''),
        releaseDate : _toDateStr(r[MODEL_MASTER_COLS.RELEASE_DATE - 1]),
        visible     : isVisible,
        quoteCount  : stats.quoteCount,
        orderCount  : stats.orderCount,
        totalAmount : stats.totalAmount,
        latestDate  : stats.latestDate,
      };
    });

    // ★ VISIBLE=false のエントリを一覧から除外（非表示設定）
    var visibleItems = items.filter(function(item) { return item.visible !== false; });

    // 最終活動日 → 更新日時の順でソート
    visibleItems.sort(function(a, b) {
      var da = a.latestDate || a.updatedAt || '';
      var db = b.latestDate || b.updatedAt || '';
      return db.localeCompare(da);
    });

    return JSON.parse(JSON.stringify({ success: true, items: visibleItems }));
  } catch(e) {
    Logger.log('[apiModelMasterList ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// API: 機種詳細取得（関連見積・注文・基板・台帳付き）
// ============================================================

function apiModelMasterGet(payload) {
  try {
    var modelCode  = String((payload || {}).modelCode  || '').trim();
    var filterCode = String((payload || {}).filterCode || '').trim(); // ★ 検索で使った特定コード
    if (!modelCode) return { success: false, error: '機種コードが必要です' };

    // ★ filterCode が指定された場合はそのコードのみで絞り込む
    // （例: modelCode="E64,E65,E66", filterCode="E66" → E66だけ表示）
    var modelCodes = filterCode
      ? [filterCode]
      : modelCode.split(/[,、\s]+/).map(function(c){ return c.trim(); }).filter(Boolean);

    function _matchMC(val) {
      var valCodes = String(val || '').split(/[,、\s]+/).map(function(c){ return c.trim(); }).filter(Boolean);
      return valCodes.some(function(vc){ return modelCodes.indexOf(vc) >= 0; });
    }

    // ── 機種マスタ情報（先頭コードで取得）──
    // filterCodeがある場合でもマスタ情報はエントリ全体（modelCode）の先頭コードで引く
    var primaryCode = modelCode.split(/[,、\s]+/).map(function(c){ return c.trim(); }).filter(Boolean)[0];
    var masterInfo  = { modelCode: modelCode, modelName: '', description: '', boardNames: '' };
    var rowNum      = _findModelRowNum(primaryCode);
    if (rowNum > 0) {
      var sheet = _getModelMasterSheet();
      var row   = sheet.getRange(rowNum, 1, 1, 7).getValues()[0];
      masterInfo = {
        id          : String(row[MODEL_MASTER_COLS.ID          - 1] || ''),
        modelCode   : modelCode,
        modelName   : String(row[MODEL_MASTER_COLS.MODEL_NAME  - 1] || ''),
        description : String(row[MODEL_MASTER_COLS.DESCRIPTION - 1] || ''),
        boardNames  : String(row[MODEL_MASTER_COLS.BOARD_NAMES - 1] || ''),
        createdAt   : _toDateStr(row[MODEL_MASTER_COLS.CREATED_AT - 1]),
        updatedAt   : _toDateStr(row[MODEL_MASTER_COLS.UPDATED_AT - 1]),
      };
    }

    // ── 関連見積書・注文書（管理シートから・複数コード対応）──
    var mgmtData      = getAllMgmtData();
    var relatedQuotes = [];
    var relatedOrders = [];
    var seenQuotes    = {};
    var seenOrders    = {};

    mgmtData.forEach(function(r) {
      var mc = String(r[MGMT_COLS.MODEL_CODE - 1] || '').trim();
      if (!_matchMC(mc)) return;

      var qNo = String(r[MGMT_COLS.QUOTE_NO - 1] || '').trim();
      var oNo = String(r[MGMT_COLS.ORDER_NO  - 1] || '').trim();

      if (qNo && !seenQuotes[qNo]) {
        seenQuotes[qNo] = true;
        relatedQuotes.push({
          mgmtId      : String(r[MGMT_COLS.ID           - 1] || ''),
          quoteNo     : qNo,
          client      : String(r[MGMT_COLS.CLIENT        - 1] || ''),
          quoteDate   : _toDateStr(r[MGMT_COLS.QUOTE_DATE - 1]),
          amount      : _toNum(r[MGMT_COLS.QUOTE_AMOUNT  - 1]),
          status      : String(r[MGMT_COLS.STATUS        - 1] || ''),
          pdfUrl      : String(r[MGMT_COLS.QUOTE_PDF_URL - 1] || ''),
          subject     : String(r[MGMT_COLS.SUBJECT       - 1] || ''),
          linked      : _isLinkedVal(r[MGMT_COLS.LINKED  - 1]),
          matchedCode : mc,  // ★ どのコードにマッチしたか
        });
      }
      if (oNo && !seenOrders[oNo]) {
        seenOrders[oNo] = true;
        relatedOrders.push({
          mgmtId       : String(r[MGMT_COLS.ID              - 1] || ''),
          orderNo      : oNo,
          client       : String(r[MGMT_COLS.CLIENT           - 1] || ''),
          orderDate    : _toDateStr(r[MGMT_COLS.ORDER_DATE   - 1]),
          amount       : _toNum(r[MGMT_COLS.ORDER_AMOUNT     - 1]),
          status       : String(r[MGMT_COLS.STATUS           - 1] || ''),
          pdfUrl       : String(r[MGMT_COLS.ORDER_PDF_URL    - 1] || ''),
          deliveryDate : _toDateStr(r[MGMT_COLS.DELIVERY_DATE - 1]),
          orderType    : String(r[MGMT_COLS.ORDER_TYPE       - 1] || ''),
          orderSlipNo  : String(r[MGMT_COLS.ORDER_SLIP_NO    - 1] || ''),
          linked       : _isLinkedVal(r[MGMT_COLS.LINKED     - 1]),
          matchedCode  : mc,  // ★ どのコードにマッチしたか
        });
      }
    });

    relatedQuotes.sort(function(a,b){ return (b.quoteDate||'').localeCompare(a.quoteDate||''); });
    relatedOrders.sort(function(a,b){ return (b.orderDate||'').localeCompare(a.orderDate||''); });

    // ── 見積台帳（複数機種コード対応）──
    var ledgerData    = getAllLedgerData();
    var relatedLedger = ledgerData
      .filter(function(r){ return _matchMC(r[LEDGER_COLS.MACHINE_CODE - 1]); })
      .map(function(r){ return _ledgerRowToObject(r); });

    // ★② 台帳と見積書を統合（mergedQuotes）quoteNoで重複除去・PDFを補完
    var mergedMap = {};
    relatedLedger.forEach(function(l) {
      var key = l.quoteNo || ('LED_' + l.ledgerId);
      mergedMap[key] = { quoteNo: l.quoteNo || '', subject: l.subject || '', client: l.dest || '', issueDate: l.issueDate || '', amount: l.amount || 0, status: l.status || '', pdfUrl: l.saveUrl || '', linked: false, source: 'ledger', compositionType: l.compositionType || '' };
    });
    relatedQuotes.forEach(function(q) {
      var key = q.quoteNo;
      if (mergedMap[key]) {
        if (!mergedMap[key].pdfUrl && q.pdfUrl) mergedMap[key].pdfUrl = q.pdfUrl;
        mergedMap[key].linked = q.linked;
        mergedMap[key].source = 'both';
      } else {
        mergedMap[key] = { quoteNo: q.quoteNo, subject: q.subject || '', client: q.client || '', issueDate: q.quoteDate || '', amount: q.amount || 0, status: q.status || '', pdfUrl: q.pdfUrl || '', linked: q.linked, source: 'mgmt', compositionType: '' };
      }
    });
    var mergedQuotes = Object.keys(mergedMap).map(function(k){ return mergedMap[k]; });
    mergedQuotes.sort(function(a,b){ return (b.issueDate||'').localeCompare(a.issueDate||''); });

    // ── BOM SS から基板一覧を取得 ──
    var boards = _getBoardsForModel(primaryCode);

    if (boards.length > 0 && rowNum > 0) {
      var boardNamesStr = boards.map(function(b){ return b.boardName; }).join('、');
      if (boardNamesStr !== masterInfo.boardNames) {
        var sh = _getModelMasterSheet();
        sh.getRange(rowNum, MODEL_MASTER_COLS.BOARD_NAMES).setValue(boardNamesStr);
        sh.getRange(rowNum, MODEL_MASTER_COLS.UPDATED_AT).setValue(nowJST());
        masterInfo.boardNames = boardNamesStr;
      }
    }

    return JSON.parse(JSON.stringify({
      success       : true,
      masterInfo    : masterInfo,
      relatedQuotes : relatedQuotes,
      relatedOrders : relatedOrders,
      relatedLedger : relatedLedger,
      mergedQuotes  : mergedQuotes,  // ★ 台帳＋見積書の統合リスト
      boards        : boards,
    }));
  } catch(e) {
    Logger.log('[apiModelMasterGet ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// API: 機種マスタ保存（新規/更新 + BOM SS 同期）
// ============================================================

function apiModelMasterSave(payload) {
  try {
    payload = payload || {};
    var modelCode   = String(payload.modelCode   || '').trim();
    var modelName   = String(payload.modelName   || '').trim();
    var description = String(payload.description || '').trim();

    if (!modelCode) return { success: false, error: '機種コードは必須です' };

    var sheet  = _getModelMasterSheet();
    var rowNum = _findModelRowNum(modelCode);
    var now    = nowJST();

    if (rowNum > 0) {
      // 更新
      sheet.getRange(rowNum, MODEL_MASTER_COLS.MODEL_NAME  ).setValue(modelName);
      sheet.getRange(rowNum, MODEL_MASTER_COLS.DESCRIPTION ).setValue(description);
      sheet.getRange(rowNum, MODEL_MASTER_COLS.UPDATED_AT  ).setValue(now);
      if (payload.modelType   !== undefined) sheet.getRange(rowNum, MODEL_MASTER_COLS.MODEL_TYPE   ).setValue(String(payload.modelType   || ''));
      if (payload.releaseDate !== undefined) sheet.getRange(rowNum, MODEL_MASTER_COLS.RELEASE_DATE ).setValue(String(payload.releaseDate || ''));
      if (payload.visible     !== undefined) sheet.getRange(rowNum, MODEL_MASTER_COLS.VISIBLE      ).setValue(payload.visible === false || payload.visible === 'false' ? false : true);
      // 機種名変更を見積台帳（BOARD_NAME列）へカスケード反映
      try { if (modelName) syncModelNameToMgmt(modelCode, modelName); } catch(e2) {}
    } else {
      // 新規（10列対応）
      sheet.appendRow([
        _generateModelId(), modelCode, modelName, description, '', now, now,
        String(payload.modelType   || ''),
        String(payload.releaseDate || ''),
        true,  // VISIBLE: デフォルト表示
      ]);
      // 削除済みリストから除外（再登録を許可）
      _removeFromDeletedModelCodes(modelCode);
    }

    // BOM SS と同期（失敗しても本体は成功扱い）
    try {
      _syncToBomModelMaster(modelCode, modelName, description);
    } catch(e2) {
      Logger.log('[apiModelMasterSave] BOM同期失敗（本体は成功）: ' + e2.message);
    }

    return { success: true, modelCode: modelCode };
  } catch(e) {
    Logger.log('[apiModelMasterSave ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// API: 機種マスタ 表示/非表示切り替え
// ============================================================

function apiModelMasterSetVisible(payload) {
  try {
    var modelCode = String((payload || {}).modelCode || '').trim();
    var visible   = payload.visible !== false && payload.visible !== 'false'; // true/false
    if (!modelCode) return { success: false, error: '機種コードが必要です' };
    var rowNum = _findModelRowNum(modelCode);
    if (rowNum < 0) return { success: false, error: '機種コードが見つかりません' };
    var sheet = _getModelMasterSheet();
    sheet.getRange(rowNum, MODEL_MASTER_COLS.VISIBLE   ).setValue(visible);
    sheet.getRange(rowNum, MODEL_MASTER_COLS.UPDATED_AT).setValue(nowJST());
    return { success: true, modelCode: modelCode, visible: visible };
  } catch(e) {
    Logger.log('[apiModelMasterSetVisible ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// API: 機種マスタ 管理表（全件・非表示含む）取得
// ============================================================

function apiModelMasterGetAll() {
  try {
    var rows = _getAllModelMasterRows();
    var items = rows.map(function(r) {
      var visible = r[MODEL_MASTER_COLS.VISIBLE - 1];
      var isVisible = !(visible === false || String(visible).toUpperCase() === 'FALSE');
      return {
        id          : String(r[MODEL_MASTER_COLS.ID          - 1] || ''),
        modelCode   : String(r[MODEL_MASTER_COLS.MODEL_CODE  - 1] || '').trim(),
        modelName   : String(r[MODEL_MASTER_COLS.MODEL_NAME  - 1] || ''),
        description : String(r[MODEL_MASTER_COLS.DESCRIPTION - 1] || ''),
        boardNames  : String(r[MODEL_MASTER_COLS.BOARD_NAMES - 1] || ''),
        createdAt   : _toDateStr(r[MODEL_MASTER_COLS.CREATED_AT - 1]),
        updatedAt   : _toDateStr(r[MODEL_MASTER_COLS.UPDATED_AT - 1]),
        modelType   : String(r[MODEL_MASTER_COLS.MODEL_TYPE   - 1] || ''),
        releaseDate : _toDateStr(r[MODEL_MASTER_COLS.RELEASE_DATE - 1]),
        visible     : isVisible,
      };
    });
    return JSON.parse(JSON.stringify({ success: true, items: items }));
  } catch(e) {
    Logger.log('[apiModelMasterGetAll ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// API: 機種マスタ削除
// ============================================================

function apiModelMasterDelete(payload) {
  try {
    var modelCode = String((payload || {}).modelCode || '').trim();
    if (!modelCode) return { success: false, error: '機種コードが必要です' };
    var rowNum = _findModelRowNum(modelCode);
    if (rowNum < 0) return { success: false, error: '機種コードが見つかりません: ' + modelCode };
    _getModelMasterSheet().deleteRow(rowNum);
    // 削除済みコードをブロックリストに追加して自動再登録を防ぐ
    _addToDeletedModelCodes(modelCode);
    return { success: true };
  } catch(e) {
    Logger.log('[apiModelMasterDelete ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// 削除済み機種コード ブロックリスト管理
// ============================================================

function _addToDeletedModelCodes(modelCode) {
  try {
    var props = PropertiesService.getScriptProperties();
    var raw   = props.getProperty('DELETED_MODEL_CODES') || '[]';
    var list  = JSON.parse(raw);
    if (list.indexOf(modelCode) < 0) list.push(modelCode);
    props.setProperty('DELETED_MODEL_CODES', JSON.stringify(list));
  } catch(e) { Logger.log('[_addToDeletedModelCodes ERROR] ' + e.message); }
}

function _removeFromDeletedModelCodes(modelCode) {
  try {
    var props = PropertiesService.getScriptProperties();
    var raw   = props.getProperty('DELETED_MODEL_CODES') || '[]';
    var list  = JSON.parse(raw).filter(function(c) { return c !== modelCode; });
    props.setProperty('DELETED_MODEL_CODES', JSON.stringify(list));
  } catch(e) { Logger.log('[_removeFromDeletedModelCodes ERROR] ' + e.message); }
}

function _isDeletedModelCode(modelCode) {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('DELETED_MODEL_CODES') || '[]';
    return JSON.parse(raw).indexOf(modelCode) >= 0;
  } catch(e) { return false; }
}

// ============================================================
// API: 全機種マスタ再構築
// （管理シート・見積台帳から機種コードを収集して機種マスタを同期）
// ============================================================

function apiRebuildModelMaster() {
  try {
    var added = 0;

    // 管理シートから収集
    var seenCodes = {};
    getAllMgmtData().forEach(function(r) {
      var mc = String(r[MGMT_COLS.MODEL_CODE - 1] || '').trim();
      if (mc && !seenCodes[mc]) { seenCodes[mc] = true; }
    });
    // 見積台帳から収集
    getAllLedgerData().forEach(function(r) {
      var mc = String(r[LEDGER_COLS.MACHINE_CODE - 1] || '').trim();
      if (mc && !seenCodes[mc]) { seenCodes[mc] = true; }
    });

    // upsert
    Object.keys(seenCodes).forEach(function(mc) {
      var rn = _findModelRowNum(mc);
      if (rn < 0) { _ensureModelCode(mc, ''); added++; }
    });

    // 基板名を一括更新
    _refreshAllBoardNames();

    Logger.log('[apiRebuildModelMaster] 完了: ' + Object.keys(seenCodes).length + '機種 / 新規追加: ' + added);
    return { success: true, total: Object.keys(seenCodes).length, added: added };
  } catch(e) {
    Logger.log('[apiRebuildModelMaster ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// 機種マスタの全行の基板名列を BOM から再取得して更新
function _refreshAllBoardNames() {
  try {
    var sheet = _getModelMasterSheet();
    var rows  = _getAllModelMasterRows();
    rows.forEach(function(r, i) {
      var mc     = String(r[MODEL_MASTER_COLS.MODEL_CODE - 1] || '').trim();
      if (!mc) return;
      var boards = _getBoardsForModel(mc);
      if (boards.length > 0) {
        var names  = boards.map(function(b) { return b.boardName; }).join('、');
        var rowNum = i + 2;
        sheet.getRange(rowNum, MODEL_MASTER_COLS.BOARD_NAMES).setValue(names);
        sheet.getRange(rowNum, MODEL_MASTER_COLS.UPDATED_AT ).setValue(nowJST());
      }
    });
  } catch(e) {
    Logger.log('[_refreshAllBoardNames ERROR] ' + e.message);
  }
}

// ============================================================
// 内部: 機種コードを機種マスタに自動登録（upsert）
// 見積書・注文書の登録時にフックして呼ぶ
// ============================================================

function _ensureModelCode(modelCode, modelName) {
  try {
    modelCode = String(modelCode || '').trim();
    if (!modelCode) return;
    if (_isDeletedModelCode(modelCode)) return; // 削除済みはスキップ（復活させない）
    if (_findModelRowNum(modelCode) >= 0) return; // 既存はスキップ

    var sheet = _getModelMasterSheet();
    var now   = nowJST();
    sheet.appendRow([
      _generateModelId(),
      modelCode,
      modelName || '',  // ★ 機種名は明示的に設定された場合のみ。空欄のままにする（コードや基板名で埋めない）
      '', '', now, now,
    ]);
    Logger.log('[_ensureModelCode] 機種マスタに追加: ' + modelCode);
  } catch(e) {
    Logger.log('[_ensureModelCode ERROR] ' + e.message);
  }
}

// ============================================================
// 内部: BOM SS の機種マスタと同期（双方向: 主SS→BOM SS）
// ============================================================

function _syncToBomModelMaster(modelCode, modelName, description) {
  var bomSs     = SpreadsheetApp.openById(BOM_SS_ID);
  var bomSheet  = bomSs.getSheetByName(BOM_SHEET.PRODUCTS);
  if (!bomSheet) throw new Error('BOM SS に機種マスタシートが見つかりません');

  var last = bomSheet.getLastRow();
  if (last <= 1) {
    // 空シート → 追加
    var newId = 'BPROD-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd') + '-' +
                (Math.floor(Math.random() * 9000) + 1000);
    bomSheet.appendRow([newId, modelName || modelCode, modelCode, description || '']);
    return;
  }

  // BOM 機種マスタの列: [機種ID, 機種名, 機種コード, 説明]
  var codeCol = 3; // 機種コードは3列目
  var codes   = bomSheet.getRange(2, codeCol, last - 1, 1).getValues().flat();
  var idx     = codes.findIndex(function(c) {
    return String(c).trim() === String(modelCode).trim();
  });

  if (idx >= 0) {
    var rowNum = idx + 2;
    bomSheet.getRange(rowNum, 2).setValue(modelName   || modelCode); // 機種名
    bomSheet.getRange(rowNum, 4).setValue(description || '');        // 説明
  } else {
    var newId2 = 'BPROD-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd') + '-' +
                 (Math.floor(Math.random() * 9000) + 1000);
    bomSheet.appendRow([newId2, modelName || modelCode, modelCode, description || '']);
  }
  Logger.log('[_syncToBomModelMaster] BOM同期完了: ' + modelCode);
}

// ============================================================
// 内部: BOM SS から機種コードに紐づく基板一覧を取得
// BOM_SHEET.PRODUCTS: [機種ID, 機種名, 機種コード, 説明]
// BOM_SHEET.BOARDS:   [基板ID, 機種ID, 基板名, コード, 説明, バージョン]
// ============================================================

function _getBoardsForModel(modelCode) {
  try {
    var bomSs     = SpreadsheetApp.openById(BOM_SS_ID);
    var prodSheet = bomSs.getSheetByName(BOM_SHEET.PRODUCTS);
    var brdSheet  = bomSs.getSheetByName(BOM_SHEET.BOARDS);
    if (!prodSheet || !brdSheet) return [];

    // 機種コード → 機種IDを取得
    var prodLast = prodSheet.getLastRow();
    if (prodLast <= 1) return [];
    var prodData  = prodSheet.getRange(2, 1, prodLast - 1, 4).getValues();
    var productId = null;
    prodData.forEach(function(r) {
      if (String(r[2] || '').trim() === String(modelCode).trim()) {
        productId = String(r[0] || '').trim();
      }
    });
    if (!productId) return [];

    // 機種IDで基板マスタを絞り込む
    var brdLast = brdSheet.getLastRow();
    if (brdLast <= 1) return [];
    var brdData = brdSheet.getRange(2, 1, brdLast - 1, 6).getValues();
    return brdData
      .filter(function(r) { return String(r[1] || '').trim() === productId; })
      .map(function(r) {
        return {
          boardId  : String(r[0] || ''),
          productId: String(r[1] || ''),
          boardName: String(r[2] || ''),
          code     : String(r[3] || ''),
          desc     : String(r[4] || ''),
          version  : String(r[5] || ''),
        };
      });
  } catch(e) {
    Logger.log('[_getBoardsForModel ERROR] ' + e.message);
    return [];
  }
}

// ============================================================
// ③ データ連携: 見積書・注文書登録後フック
// 02_ocr_and_processing.gs や 05_matching_engine.gs から呼び出す
// ============================================================

/**
 * 見積書または注文書の登録・更新後に呼ぶ。
 * modelCode が設定されていれば機種マスタに自動登録。
 * @param {string} modelCode
 * @param {string} [modelName] - 任意で渡すと機種名として登録
 */
function onDocRegistered(modelCode, modelName) {
  if (!modelCode || String(modelCode).trim() === '') return;
  _ensureModelCode(String(modelCode).trim(), modelName || '');
}

// ============================================================
// ③ データ連携: 見積台帳の機種コードを機種マスタに同期
// ============================================================

function syncLedgerModelCodes() {
  try {
    var ledgerData = getAllLedgerData();
    ledgerData.forEach(function(r) {
      var mc = String(r[LEDGER_COLS.MACHINE_CODE - 1] || '').trim();
      var bn = String(r[LEDGER_COLS.BOARD_NAME   - 1] || '').trim();
      if (mc) _ensureModelCode(mc, '');  // ★ 基板名を機種名として使わない
    });
    Logger.log('[syncLedgerModelCodes] 見積台帳→機種マスタ同期完了');
    return { success: true };
  } catch(e) {
    Logger.log('[syncLedgerModelCodes ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// ③ データ連携: 機種マスタの機種名変更を管理シートへ反映
// ============================================================

function syncModelNameToMgmt(modelCode, newModelName) {
  try {
    // 管理シートは MODEL_CODE を参照するだけで名前は持っていないので
    // 見積台帳の boardName 列を更新する
    var ss         = getSpreadsheet();
    var ledgerSheet = ss.getSheetByName(CONFIG.SHEET_LEDGER);
    if (!ledgerSheet || ledgerSheet.getLastRow() <= 1) return { success: true };

    var last    = ledgerSheet.getLastRow();
    var mcVals  = ledgerSheet.getRange(2, LEDGER_COLS.MACHINE_CODE, last - 1, 1).getValues().flat();
    var updated = 0;
    mcVals.forEach(function(mc, i) {
      if (String(mc || '').trim() === String(modelCode).trim()) {
        ledgerSheet.getRange(i + 2, LEDGER_COLS.BOARD_NAME).setValue(newModelName);
        updated++;
      }
    });
    Logger.log('[syncModelNameToMgmt] 更新行数: ' + updated);
    return { success: true, updated: updated };
  } catch(e) {
    Logger.log('[syncModelNameToMgmt ERROR] ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// ③ データ連携: BOM 基板マスタ変更を機種マスタに反映
// 基板マスタの CRUD 後（08_parts_management.gs）からフック
// ============================================================

function onBoardChanged(productId) {
  try {
    if (!productId) return;
    // productId → modelCode を逆引き
    var bomSs     = SpreadsheetApp.openById(BOM_SS_ID);
    var prodSheet = bomSs.getSheetByName(BOM_SHEET.PRODUCTS);
    if (!prodSheet || prodSheet.getLastRow() <= 1) return;
    var prodData  = prodSheet.getRange(2, 1, prodSheet.getLastRow() - 1, 3).getValues();
    var modelCode = null;
    prodData.forEach(function(r) {
      if (String(r[0] || '').trim() === String(productId).trim()) {
        modelCode = String(r[2] || '').trim(); // 機種コード（3列目）
      }
    });
    if (!modelCode) return;

    // 基板名を機種マスタに更新
    var boards  = _getBoardsForModel(modelCode);
    var rowNum  = _findModelRowNum(modelCode);
    if (rowNum > 0 && boards.length > 0) {
      var names = boards.map(function(b) { return b.boardName; }).join('、');
      var sheet = _getModelMasterSheet();
      sheet.getRange(rowNum, MODEL_MASTER_COLS.BOARD_NAMES).setValue(names);
      sheet.getRange(rowNum, MODEL_MASTER_COLS.UPDATED_AT ).setValue(nowJST());
      Logger.log('[onBoardChanged] 機種マスタ基板名更新: ' + modelCode + ' → ' + names);
    }
  } catch(e) {
    Logger.log('[onBoardChanged ERROR] ' + e.message);
  }
}

// ============================================================
// トリガー登録
// ============================================================

function setupModelMasterTriggers() {
  var existing = ScriptApp.getProjectTriggers().map(function(t){ return t.getHandlerFunction(); });
  if (existing.indexOf('dailySyncModelMaster') < 0) {
    ScriptApp.newTrigger('dailySyncModelMaster').timeBased().atHour(8).everyDays(1).create();
    Logger.log('[setupModelMasterTriggers] dailySyncModelMaster トリガー登録（毎日8時）');
  }
}

// 毎朝8時: 機種マスタを自動同期
function dailySyncModelMaster() {
  Logger.log('[dailySyncModelMaster] 開始');
  syncLedgerModelCodes();
  _refreshAllBoardNames();
  Logger.log('[dailySyncModelMaster] 完了');
}

// ============================================================
// テスト
// ============================================================

function testModelMasterList() {
  var res = apiModelMasterList();
  Logger.log('機種数: ' + (res.items || []).length);
  (res.items || []).slice(0, 3).forEach(function(item) {
    Logger.log(item.modelCode + ' / ' + item.modelName + ' / 見積' + item.quoteCount + '件 注文' + item.orderCount + '件');
  });
}

function testModelMasterGet() {
  var rows = _getAllModelMasterRows();
  if (rows.length === 0) { Logger.log('機種マスタにデータがありません'); return; }
  var mc  = String(rows[0][MODEL_MASTER_COLS.MODEL_CODE - 1]).trim();
  var res = apiModelMasterGet({ modelCode: mc });
  Logger.log('機種コード: ' + mc);
  Logger.log('関連見積: ' + res.relatedQuotes.length + '件');
  Logger.log('関連注文: ' + res.relatedOrders.length + '件');
  Logger.log('基板: ' + res.boards.length + '件');
}
