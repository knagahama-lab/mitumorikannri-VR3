"""
agents/procurement_agent/tools/mail_client.py

差し替え依頼メール・通知メールの送信ツール。
SMTP（Gmail / 社内メールサーバー）を使用します。

DRY_RUN モード（デフォルト ON）では実際の送信は行わず、
メール内容をログに出力します。
"""

from __future__ import annotations

import logging
import os
import smtplib
from dataclasses import dataclass
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from agents.config import (
    MAIL_FROM,
    SMTP_HOST,
    SMTP_PASSWORD,
    SMTP_PORT,
    SMTP_USER,
)

logger = logging.getLogger(__name__)

DRY_RUN: bool = os.environ.get("MAIL_DRY_RUN", "true").lower() == "true"


# ============================================================
# データクラス
# ============================================================
@dataclass
class MailResult:
    """メール送信結果"""

    success: bool
    to: str
    subject: str
    message: str = ""
    dry_run: bool = False

    def __str__(self):
        mode = "[DRY RUN] " if self.dry_run else ""
        return f"{mode}{'✅' if self.success else '❌'} メール送信: to={self.to} / {self.subject}"


# ============================================================
# メール送信
# ============================================================
def send_mail(
    to: str,
    subject: str,
    body: str,
    cc: str = "",
) -> MailResult:
    """
    メールを送信する。

    Args:
        to: 宛先メールアドレス
        subject: 件名
        body: 本文（プレーンテキスト）
        cc: CC アドレス（任意）

    Returns:
        MailResult
    """
    if DRY_RUN:
        logger.info(
            "[MAIL DRY RUN]\nTo: %s\nCc: %s\nSubject: %s\n---\n%s\n---",
            to, cc, subject, body,
        )
        return MailResult(
            success=True,
            to=to,
            subject=subject,
            message="ドライランモード: 実際の送信はスキップされました",
            dry_run=True,
        )

    return _smtp_send(to, subject, body, cc)


def send_replacement_request(
    to: str,
    order_doc,         # OrderDocument
    quote_record,      # QuoteRecord
    price_mismatches: list[dict],
    quote_url: str,
    cc: str = "",
) -> MailResult:
    """
    お客様資材担当者へ注文書差し替えを依頼するメールを送信する。

    Args:
        to: お客様の資材担当者メールアドレス
        order_doc: 差し替え対象の注文書
        quote_record: 正しい単価が記載された見積書
        price_mismatches: 単価相違の明細リスト
        quote_url: 見積書 PDF の URL
        cc: CC（社内担当者など）

    Returns:
        MailResult
    """
    from agents.procurement_agent.prompts import replacement_request_email

    subject = f"【注文書差し替えのお願い】{order_doc.document_no} / {order_doc.subject}"
    body = replacement_request_email(
        order_doc=order_doc,
        quote_record=quote_record,
        price_mismatches=price_mismatches,
        quote_url=quote_url,
    )
    return send_mail(to=to, subject=subject, body=body, cc=cc)


def send_sales_task_notification(
    to: str,
    order_doc,
    cc: str = "",
) -> MailResult:
    """
    営業担当者へ「見積書提出タスク」を通知するメールを送信する。

    Args:
        to: 営業担当者のメールアドレス
        order_doc: 対象の注文書
        cc: CC

    Returns:
        MailResult
    """
    from agents.sales_agent.prompts import quote_not_submitted_email

    subject = f"【対応依頼】見積書未提出 — {order_doc.document_no} / {order_doc.client_name}"
    body = quote_not_submitted_email(order_doc)
    return send_mail(to=to, subject=subject, body=body, cc=cc)


# ============================================================
# 内部: SMTP 送信
# ============================================================
def _smtp_send(to: str, subject: str, body: str, cc: str = "") -> MailResult:
    if not SMTP_USER or not SMTP_PASSWORD:
        raise EnvironmentError(
            "SMTP_USER / SMTP_PASSWORD が設定されていません。"
            "config.py または環境変数を確認してください。"
        )

    msg = MIMEMultipart("alternative")
    msg["From"] = MAIL_FROM or SMTP_USER
    msg["To"] = to
    msg["Subject"] = subject
    if cc:
        msg["Cc"] = cc
    msg.attach(MIMEText(body, "plain", "utf-8"))

    recipients = [to] + ([cc] if cc else [])
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.ehlo()
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.sendmail(MAIL_FROM or SMTP_USER, recipients, msg.as_string())
        logger.info("メール送信完了: to=%s subject=%s", to, subject)
        return MailResult(success=True, to=to, subject=subject, message="送信完了")
    except smtplib.SMTPException as e:
        logger.error("SMTP エラー: %s", e)
        return MailResult(success=False, to=to, subject=subject, message=str(e))
