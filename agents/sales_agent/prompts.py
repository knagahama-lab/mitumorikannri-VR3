"""
agents/sales_agent/prompts.py

sales_agent が使用するシステムプロンプトと
各ステップで使う動的プロンプトのテンプレート集。
"""

# ============================================================
# システムプロンプト
# ============================================================
SYSTEM_PROMPT = """\
あなたは「営業アシスタント Agent（sales_agent）」です。
注文書の受領から単価チェック・社内通知までを担当します。

## あなたの責務
1. 受領した注文書 PDF を OCR 解析し、単価・品名・数量を抽出する
2. 管理システムから該当する見積書を検索し、単価を照合する
3. 単価が正しい場合 → 営業担当者へ押印・回覧の指示を出す
4. 単価が間違っている場合 → 見積書の提出状況を確認し、
   - 提出済みなら見積書 URL を資材購買（procurement_agent）に連絡する
   - 未提出なら営業担当者へ「見積書提出タスク」の通知とメール下書きを作成する

## 行動原則
- ツールを使用する前に、実行内容をユーザーに簡潔に説明してください
- 判断に迷う場合はユーザーに確認してください
- 金額・番号などの重要情報は必ず原文を引用してください
- メール文面は丁寧かつ簡潔な日本語ビジネス文体で作成してください
"""

# ============================================================
# 単価チェック確認プロンプト
# ============================================================
def price_check_prompt(order_doc, check_result) -> str:
    """単価チェック結果をもとに次のアクションを判断させるプロンプト"""
    mismatch_detail = ""
    if not check_result.is_correct:
        rows = "\n".join(
            f"  ・{m['itemName']}: 注文書 {m['orderPrice']:,.0f}円 ／ 見積書 {m['quotePrice']:,.0f}円"
            f"（差異 {m['diffRate']:.1f}%）"
            for m in check_result.mismatches
        )
        mismatch_detail = f"\n【単価相違明細】\n{rows}"

    return f"""\
注文書の単価チェックが完了しました。

【注文書情報】
- 発注書番号: {order_doc.document_no}
- 発注元: {order_doc.client_name}
- 件名: {order_doc.subject}
- 合計金額: ¥{order_doc.total_amount:,.0f}
- 照合見積番号: {check_result.matched_quote_no or "（未特定）"}

【チェック結果】
{check_result.summary()}{mismatch_detail}

次のアクションを決定し、適切なツールを実行してください。
"""


# ============================================================
# 営業担当者への通知メール（見積書未提出）
# ============================================================
def quote_not_submitted_email(order_doc, recipient_name: str = "営業担当者") -> str:
    """
    見積書が未提出の場合に営業担当者へ送る通知メールの本文テンプレート。
    """
    return f"""\
{recipient_name} 様

お疲れ様です。

受領した注文書の単価チェックにおいて、以下の件について
**見積書がまだ提出されていない**ことが確認されました。

【注文書情報】
　発注書番号: {order_doc.document_no}
　発注元: {order_doc.client_name}
　件名: {order_doc.subject}
　発注日: {order_doc.document_date}

お客様の購買担当者へ**早急に見積書をご提出**いただけますでしょうか。
見積書の提出後、本件の注文書処理を再開いたします。

何かご不明な点がございましたら、担当者までご連絡ください。

よろしくお願いいたします。
"""


# ============================================================
# 資材購買への見積 URL 連絡メッセージ
# ============================================================
def quote_url_notification(order_doc, quote_record, quote_url: str) -> str:
    """
    資材購買（procurement_agent）へ見積 URL を連絡するメッセージテンプレート。
    """
    return f"""\
【見積書URL連絡 / 注文書差し替え依頼案件】

注文書の単価相違が確認されました。
以下の見積書 URL をご確認の上、お客様へ注文書の差し替えをご依頼ください。

【注文書情報】
　発注書番号: {order_doc.document_no}
　発注元: {order_doc.client_name}
　件名: {order_doc.subject}
　発注日: {order_doc.document_date}
　合計金額: ¥{order_doc.total_amount:,.0f}

【紐づく見積書】
　見積番号: {quote_record.quote_no}
　件名: {quote_record.subject}
　金額: ¥{quote_record.total_amount:,.0f}
　見積書 PDF URL: {quote_url}

上記の見積書の単価が正しい金額となります。
注文書の差し替えをお客様にご依頼をお願いします。
"""


# ============================================================
# 押印・回覧依頼メッセージ（単価正常時）
# ============================================================
def approval_circulation_message(order_doc, quote_record) -> str:
    """営業担当者への押印・回覧依頼メッセージ"""
    return f"""\
【注文書 単価確認完了 — 押印・回覧依頼】

単価チェックの結果、以下の注文書の単価に問題がないことを確認しました。
営業担当者による**押印・回覧**をお願いします。

【注文書情報】
　発注書番号: {order_doc.document_no}
　発注元: {order_doc.client_name}
　件名: {order_doc.subject}
　発注日: {order_doc.document_date}
　合計金額: ¥{order_doc.total_amount:,.0f}

【照合見積書】
　見積番号: {quote_record.quote_no if quote_record else '—'}
　金額: ¥{quote_record.total_amount:,.0f if quote_record else 0}

確認後、資材購買担当者へ販売管理登録の指示をお願いします。
"""
