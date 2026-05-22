// ============================================================
// 22_custom_doc.gs
// カスタム書類タイプ管理（汎用書類テンプレート）
// 家計簿・指示書・請求書・在庫管理 など任意の書類を管理
// ============================================================

// ── 内部ユーティリティ ──────────────────────────────────────

function _getCustomDocTypes() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('CUSTOM_DOC_TYPES');
    return raw ? JSON.parse(raw) : [];
  } catch(e) { return []; }
}

function _saveCustomDocTypes(types) {
  PropertiesService.getScriptProperties().setProperty('CUSTOM_DOC_TYPES', JSON.stringify(types));
}

function _getCustomDocSheet(type) {
  if (!type.spreadsheetId) return null;
  var ss = SpreadsheetApp.openById(type.spreadsheetId);
  var sheetName = type.sheetName || 'データ';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var sysHeaders = ['ID', 'ステータス', '登録日時', '更新日時'];
    var fieldHeaders = (type.fields || []).map(function(f){ return f.label; });
    var allHeaders = sysHeaders.concat(fieldHeaders);
    sheet.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders])
      .setBackground('#E8F0FE').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, allHeaders.length);
  }
  return sheet;
}

// ── 書類タイプ設定 CRUD ──────────────────────────────────────

function apiCustomDocTypeList() {
  try {
    return { success: true, types: _getCustomDocTypes() };
  } catch(e) { return { success: false, error: e.message }; }
}

function apiCustomDocTypeSave(payload) {
  try {
    var types = _getCustomDocTypes();
    var type  = payload.type;
    if (!type || !type.name) return { success: false, error: '名前は必須です' };
    if (!type.id) {
      type.id = 'cd_' +
        Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + '_' +
        (Math.floor(Math.random() * 9000) + 1000);
    }
    var idx = types.findIndex(function(t){ return t.id === type.id; });
    if (idx >= 0) types[idx] = type;
    else          types.push(type);
    _saveCustomDocTypes(types);
    return { success: true, id: type.id };
  } catch(e) { return { success: false, error: e.message }; }
}

function apiCustomDocTypeDelete(payload) {
  try {
    var types = _getCustomDocTypes().filter(function(t){ return t.id !== payload.id; });
    _saveCustomDocTypes(types);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── データ CRUD ───────────────────────────────────────────────

function apiCustomDocGetAll(payload) {
  try {
    var types = _getCustomDocTypes();
    var type  = types.find(function(t){ return t.id === payload.typeId; });
    if (!type) return { success: false, error: '書類タイプが見つかりません' };
    if (!type.spreadsheetId) return {
      success: true, items: [], fields: type.fields || [], statusOptions: (type.statusOptions||'').split(',').filter(Boolean)
    };

    var sheet = _getCustomDocSheet(type);
    if (!sheet || sheet.getLastRow() <= 1) return {
      success: true, items: [], fields: type.fields || [], statusOptions: (type.statusOptions||'').split(',').filter(Boolean)
    };

    var sysKeys    = ['_id', '_status', '_createdAt', '_updatedAt'];
    var fieldKeys  = (type.fields || []).map(function(f){ return f.key; });
    var allKeys    = sysKeys.concat(fieldKeys);
    var totalCols  = allKeys.length;

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, totalCols).getValues();
    var items = data
      .filter(function(row){ return String(row[0]).trim() !== ''; })
      .map(function(row){
        var obj = {};
        allKeys.forEach(function(key, i){
          var v = row[i];
          obj[key] = (v instanceof Date) ? _toDateStr(v) : String(v === null || v === undefined ? '' : v);
        });
        return obj;
      });

    return {
      success: true,
      items: items,
      fields: type.fields || [],
      statusOptions: (type.statusOptions||'').split(',').filter(Boolean)
    };
  } catch(e) { return { success: false, error: e.message }; }
}

function apiCustomDocSave(payload) {
  try {
    var types = _getCustomDocTypes();
    var type  = types.find(function(t){ return t.id === payload.typeId; });
    if (!type)               return { success: false, error: '書類タイプが見つかりません' };
    if (!type.spreadsheetId) return { success: false, error: 'スプレッドシートIDが設定されていません' };

    var sheet     = _getCustomDocSheet(type);
    var fieldKeys = (type.fields || []).map(function(f){ return f.key; });
    var entry     = payload.entry || {};
    var now       = nowJST();

    if (entry._id) {
      // 更新
      var lastRow = sheet.getLastRow();
      if (lastRow <= 1) return { success: false, error: 'データが見つかりません' };
      var ids    = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String);
      var rowIdx = ids.indexOf(String(entry._id));
      if (rowIdx < 0) return { success: false, error: 'IDが見つかりません: ' + entry._id };
      var rowNum  = rowIdx + 2;
      var created = sheet.getRange(rowNum, 3).getValue();
      var rowData = [entry._id, entry._status || '', created, now];
      fieldKeys.forEach(function(k){ rowData.push(entry[k] !== undefined ? entry[k] : ''); });
      sheet.getRange(rowNum, 1, 1, rowData.length).setValues([rowData]);
      return { success: true, id: entry._id };
    } else {
      // 新規
      var newId = 'CD-' +
        Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd') + '-' +
        (Math.floor(Math.random() * 9000) + 1000);
      var defaultStatus = (type.statusOptions||'').split(',').filter(Boolean)[0] || '';
      var rowData = [newId, entry._status || defaultStatus, now, now];
      fieldKeys.forEach(function(k){ rowData.push(entry[k] !== undefined ? entry[k] : ''); });
      sheet.appendRow(rowData);
      return { success: true, id: newId };
    }
  } catch(e) { return { success: false, error: e.message }; }
}

function apiCustomDocDelete(payload) {
  try {
    var types = _getCustomDocTypes();
    var type  = types.find(function(t){ return t.id === payload.typeId; });
    if (!type || !type.spreadsheetId) return { success: false, error: '設定が見つかりません' };

    var sheet   = _getCustomDocSheet(type);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, error: 'データが見つかりません' };

    var ids    = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String);
    var rowIdx = ids.indexOf(String(payload.entryId));
    if (rowIdx < 0) return { success: false, error: 'IDが見つかりません' };
    sheet.deleteRow(rowIdx + 2);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── Gemini OCR（カスタム書類用） ─────────────────────────────

function apiCustomDocOcr(payload) {
  try {
    var types = _getCustomDocTypes();
    var type  = types.find(function(t){ return t.id === payload.typeId; });
    if (!type) return { success: false, error: '書類タイプが見つかりません' };

    var apiKey = getGeminiApiKey();
    if (!apiKey) return { success: false, error: 'Gemini APIキーが設定されていません' };

    var ocrFields = (type.fields || []).filter(function(f){ return f.ocr; });
    if (ocrFields.length === 0) return { success: false, error: 'OCR対象フィールドが設定されていません' };

    var jsonTemplate = '{' + ocrFields.map(function(f){
      return '"' + f.key + '": "' + f.label + 'の値"';
    }).join(', ') + '}';

    var prompt = type.ocrPrompt
      ? type.ocrPrompt + '\n\n必ず以下のJSON形式のみで返答してください（コードブロック・説明文なし）:\n' + jsonTemplate
      : '画像から次の項目を抽出してください: ' +
        ocrFields.map(function(f){ return f.label + '（' + f.key + '）'; }).join('、') +
        '\n\n必ず以下のJSON形式のみで返答してください（コードブロック・説明文なし）:\n' + jsonTemplate;

    var url  = CONFIG.GEMINI_API_ENDPOINT + CONFIG.GEMINI_PRIMARY_MODEL + ':generateContent?key=' + apiKey;
    var body = {
      contents: [{ parts: [
        { text: prompt },
        { inline_data: { mime_type: payload.mimeType || 'image/jpeg', data: payload.base64 } }
      ]}],
      generationConfig: { temperature: 0.1 }
    };

    var res  = fetchWithRetry(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(body), muteHttpExceptions: true
    });
    var data = JSON.parse(res.getContentText());
    var text = '';
    try { text = data.candidates[0].content.parts[0].text; } catch(e2) {}

    var jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) return { success: false, error: 'OCR結果をJSONとして解析できませんでした: ' + text.substring(0, 200) };

    var extracted = JSON.parse(jsonMatch[0]);
    return { success: true, extracted: extracted };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── プリセットテンプレート ────────────────────────────────────

function apiCustomDocPresets() {
  return { success: true, presets: [
    {
      name: '家計簿',
      icon: '💰',
      color: '#16a34a',
      sheetName: '家計簿データ',
      statusOptions: '未確認,確認済み',
      ocrPrompt: 'レシートや領収書から日付・カテゴリ（食費/交通費/光熱費/娯楽/医療/日用品/その他）・金額・支払先を抽出してください。',
      fields: [
        { key:'date',      label:'日付',     type:'date',     required:true,  showInTable:true,  ocr:true  },
        { key:'category',  label:'カテゴリ',  type:'select',   required:true,  showInTable:true,  ocr:true,  options:'食費,交通費,光熱費,娯楽,医療,日用品,その他' },
        { key:'amount',    label:'金額（円）', type:'number',   required:true,  showInTable:true,  ocr:true  },
        { key:'payee',     label:'支払先',    type:'text',     required:false, showInTable:true,  ocr:true  },
        { key:'payMethod', label:'支払方法',  type:'select',   required:false, showInTable:false, ocr:false, options:'現金,クレジットカード,電子マネー,振込,その他' },
        { key:'memo',      label:'メモ',      type:'textarea', required:false, showInTable:false, ocr:false }
      ]
    },
    {
      name: '指示書',
      icon: '📋',
      color: '#2563eb',
      sheetName: '指示書データ',
      statusOptions: '未対応,対応中,完了,保留',
      ocrPrompt: '指示書や作業指示書から指示日・件名・担当者・内容・期限を抽出してください。',
      fields: [
        { key:'issueDate', label:'指示日',   type:'date',     required:true,  showInTable:true,  ocr:true  },
        { key:'title',     label:'件名',     type:'text',     required:true,  showInTable:true,  ocr:true  },
        { key:'assignee',  label:'担当者',   type:'text',     required:false, showInTable:true,  ocr:true  },
        { key:'dueDate',   label:'期限',     type:'date',     required:false, showInTable:true,  ocr:true  },
        { key:'priority',  label:'優先度',   type:'select',   required:false, showInTable:true,  ocr:false, options:'高,中,低' },
        { key:'content',   label:'指示内容', type:'textarea', required:false, showInTable:false, ocr:true  },
        { key:'memo',      label:'備考',     type:'textarea', required:false, showInTable:false, ocr:false }
      ]
    },
    {
      name: '請求書',
      icon: '🧾',
      color: '#d97706',
      sheetName: '請求書データ',
      statusOptions: '未払い,支払済み,確認中,キャンセル',
      ocrPrompt: '請求書から請求日・請求書番号・発行元・金額・支払期限を抽出してください。',
      fields: [
        { key:'invoiceDate', label:'請求日',    type:'date',     required:true,  showInTable:true,  ocr:true  },
        { key:'invoiceNo',   label:'請求書番号', type:'text',     required:false, showInTable:true,  ocr:true  },
        { key:'issuer',      label:'発行元',    type:'text',     required:true,  showInTable:true,  ocr:true  },
        { key:'amount',      label:'金額（円）', type:'number',   required:true,  showInTable:true,  ocr:true  },
        { key:'dueDate',     label:'支払期限',  type:'date',     required:false, showInTable:true,  ocr:true  },
        { key:'description', label:'内容',      type:'textarea', required:false, showInTable:false, ocr:true  },
        { key:'memo',        label:'備考',      type:'textarea', required:false, showInTable:false, ocr:false }
      ]
    },
    {
      name: '在庫管理',
      icon: '📦',
      color: '#7c3aed',
      sheetName: '在庫データ',
      statusOptions: '在庫あり,在庫少,在庫切れ,発注中',
      ocrPrompt: '納品書や在庫リストから品名・品番・数量・単位・単価を抽出してください。',
      fields: [
        { key:'itemName',  label:'品名',      type:'text',     required:true,  showInTable:true,  ocr:true  },
        { key:'itemCode',  label:'品番',      type:'text',     required:false, showInTable:true,  ocr:true  },
        { key:'qty',       label:'数量',      type:'number',   required:true,  showInTable:true,  ocr:true  },
        { key:'unit',      label:'単位',      type:'text',     required:false, showInTable:true,  ocr:false },
        { key:'unitPrice', label:'単価（円）', type:'number',   required:false, showInTable:true,  ocr:true  },
        { key:'location',  label:'保管場所',  type:'text',     required:false, showInTable:true,  ocr:false },
        { key:'memo',      label:'備考',      type:'textarea', required:false, showInTable:false, ocr:false }
      ]
    }
  ]};
}
