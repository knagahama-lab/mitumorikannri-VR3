"""
agents/procurement_agent/prompts.py

procurement_agent が使用するシステムプロンプトと
差し替え依頼メール・ERP 登録確認などのテンプレート集。
"""

# ============================================================
# システムプロンプト
# ============================================================
SYSTEM_PROMPT = """\
あなたは「資材購買アシスタント Agent（procurement_agent）」です。
営業アシスタントから受け取った情報をもとに、以下を担当します。

## フロー A: 単価 OK → 販売管理登録
  1. 注文書データを ERP（販売管理システム）に登録する
  2. 登録完了を関係者に通知する

## フロー B: 単価 NG・見積書提出済み → 注文書差し替え依頼
  1. 見積書 PDF URL と単価相違内容を確認する
  2. お客様資材担当者へ「注文書差し替え依頼メール」を作成・送信する
  3. 社内担当者へ差し替え依頼状況を共有する

## 行動原則
- ツールを使用する前に、実行内容を簡潔に説明してください
- 実際に送信・登録を行う前に、内容をユーザーに確認してください
- メール文面は丁寧かつ簡潔な日本語ビジネス文体で作成してください
- 金額・番号などの重要情報は必ず原文を引用してください
"""


# ============================================================
# 注文書差し替え依頼メール
# ============================================================
def replacement_request_email(
    order_doc,
    quote_record,
    price_mismatches: list[dict],
    quote_url: str,
) -> str:
    """
    お客様の資材担当者へ送る注文書差し替え依頼メールの本文。

    Args:
        order_doc: 差し替え対象の注文書
        quote_record: 正しい単価が含まれる見積書
        price_mismatches: [{"itemName": str, "orderPrice": float, "quotePrice": float}, ...]
        quote_url: 見積書 PDF のダウンロード URL
    """
    mismatch_table = "\n".join(
        f"  ・{m['itemName']}: "
        f"ご発注単価 ¥{m['orderPrice']:,.0f} → "
        f"弊社見積単価 ¥{m['quotePrice']:,.0f}"
        for m in price_mismatches
    )

    return f"""\
{order_doc.client_name} 御中
資材購買ご担当者様

いつも大変お世話になっております。

このたびご発行いただきました発注書について、
単価に相違がございましたため、大変恐れ入りますが
**注文書の再発行（差し替え）**をお願いしたく、ご連絡申し上げます。

【対象発注書】
　発注書番号: {order_doc.document_no}
　件名: {order_doc.subject}
　発注日: {order_doc.document_date}

【単価相違内容】
{mismatch_table}

弊社お見積書（{quote_record.quote_no}）の単価が正しい金額となります。
下記 URL よりご確認いただけますでしょうか。

　見積書 PDF: {quote_url}

誠にお手数をおかけしますが、
上記見積書の単価にてご発注書を再発行いただけますようお願い申し上げます。

ご不明な点がございましたら、担当者までお気軽にご連絡ください。

よろしくお願いいたします。
"""


# ============================================================
# 販売管理登録完了通知
# ============================================================
def erp_registered_notification(
    order_doc,
    quote_record,
    erp_order_id: str,
) -> str:
    """ERP 登録完了後の社内通知メッセージ"""
    return f"""\
【販売管理登録完了】

以下の注文書が販売管理システムへ登録されました。

【注文書情報】
　発注書番号 : {order_doc.document_no}
　発注元     : {order_doc.client_name}
　件名       : {order_doc.subject}
　発注日     : {order_doc.document_date}
　合計金額   : ¥{order_doc.total_amount:,.0f}

【紐づく見積書】
　見積番号   : {quote_record.quote_no if quote_record else "—"}

【ERP 受注 ID】
　{erp_order_id}

以上、ご確認よろしくお願いいたします。
"""


# ============================================================
# 差し替え完了後の社内共有メッセージ
# ============================================================
def replacement_sent_notification(order_doc, sent_to: str) -> str:
    """差し替え依頼メール送信完了後の社内共有メッセージ"""
    return f"""\
【注文書差し替え依頼 送信済み】

以下の案件について、お客様へ注文書差し替えの依頼メールを送信しました。

　発注書番号 : {order_doc.document_no}
　発注元     : {order_doc.client_name}
　件名       : {order_doc.subject}
　送信先     : {sent_to}

お客様からの再発行をお待ちください。
再発行後、再度単価チェックを実施して販売管理へ登録します。
"""
