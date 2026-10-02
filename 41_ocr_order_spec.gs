// ============================================================
// 41_ocr_order_spec.gs
// 注文書OCRの「出力形式とルール」を1か所にまとめたもの
//
//   注文書のOCRは 13_ocr_extended.gs の
//     ・_buildTextStructurePrompt（PDFのテキストをGeminiで構造化）
//     ・_buildOcrPrompt（PDFを画像としてGeminiで読取）
//   の2経路があり（02_ocr_and_processing.gs にも同名関数があるが 13 の定義が優先される）、
//   どちらからも本ファイルの _ocrOrderSpec() を使う。注文書の読み取り方を変える時はここだけ直せばよい。
//
//   対応している書式の特徴
//     ・藤商事：部品コード（10桁）が品名の上段にある
//     ・コナミ：1つのPDFにページごとに別の発注番号の発注書がある／明細番号 00010…／品目欄が英数字コード
// ============================================================

function _ocrSelfName() {
  return (PropertiesService.getScriptProperties().getProperty('SELF_COMPANY_NAMES') || 'サン電子').split(',')[0];
}

/** 注文書の出力JSON形式（1件分） */
function _ocrOrderSchema() {
  return [
    '{',
    '  "actionType": "new / revision / cancellation",',
    '  "reason": "差し替えやキャンセルの理由（新規なら空文字）",',
    '  "documentNo": "発注書番号・注文番号",',
    '  "documentDate": "発注日 YYYY/MM/DD",',
    '  "clientName": "発注元の会社名（この発注書を発行した会社＝顧客）",',
    '  "issuerName": "発行元の会社名（社名ロゴ・住所・印がある側）",',
    '  "subject": "件名",',
    '  "modelCode": "機種コード（なければ空文字）",',
    '  "orderSlipNo": "発注伝票番号（なければ空文字）",',
    '  "linkedQuoteNo": "対応する見積番号（記載があれば。なければ空文字）",',
    '  "orderType": "試作 または 量産（「試作購買」「試作」の記載は試作。不明なら空文字）",',
    '  "subtotal": 小計(数値), "tax": 消費税(数値), "totalAmount": 税込合計(数値),',
    '  "lineItems": [',
    '    {"lineNo":"明細の行番号・項目番号（00010 など）","partCode":"客先の部品コード・品番・品目コード","itemName":"品名","drawingNo":"図番・型式（なければ空文字）","spec":"仕様","firstDelivery":"初回納入日・納期 YYYY/MM/DD","deliveryDest":"納入先","qty":数量,"unit":"単位","unitPrice":単価,"amount":金額,"quoteRef":"この明細が参照する見積番号（なければ空文字）","remarks":"備考"}',
    '  ],',
    '  "additionalOrders": []',
    '}',
  ].join('\n');
}

/** 注文書の読み取りルール */
function _ocrOrderRules() {
  var self = _ocrSelfName();
  return [
    '## 注文書の読み取りルール',
    '- clientName／issuerName は発注書を発行した会社（顧客）。宛先「殿」「御中」側の「' + self + '」は受注者なので入れない',
    '- partCode：客先の部品コード。品名の上段や左にある数字（例: 2601156062）や、品目欄の英数字コード（例: KAM1121A, SNB52163A）。itemName には含めない。品目欄にコードしか無い場合は itemName にも同じコードを入れる',
    '- 明細の行番号・項目番号（00010, 00020 など）は partCode ではなく lineNo',
    '- 書類全体に納期・納入期日が1つだけ書かれている場合も、各明細の firstDelivery に入れる',
    '- 明細ごとに見積番号の参照がある場合は quoteRef に。全明細で同じなら linkedQuoteNo にも入れる',
    '- 1つのPDFに発注番号の異なる発注書が複数ある場合（ページごとに別の発注書など）は合算しない。1件目をトップレベル、2件目以降を additionalOrders に上と同じ形式（lineItems を含む）で1件ずつ入れる。1件だけなら additionalOrders は空配列',
    '- 「差し替え」「訂正」「版数更新」→ revision、「中止」「取消」「キャンセル」→ cancellation、それ以外 → new',
  ].join('\n');
}

/** 注文書の完全な指示（ヘッダ＋形式＋共通ルール＋注文書ルール） */
function _ocrOrderSpec(intro, commonRules) {
  return [intro, '', _ocrOrderSchema(), '', commonRules || '', '', _ocrOrderRules()].join('\n');
}
