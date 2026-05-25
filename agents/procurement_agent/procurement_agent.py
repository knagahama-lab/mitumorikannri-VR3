"""
agents/procurement_agent/procurement_agent.py

資材購買アシスタント Agent のメインロジック。

フロー担当範囲（sales_agent の結果を受け取る）:
  A. 単価 OK  → ERP 販売管理登録 → 登録完了通知
  B. 単価 NG・見積書提出済み
       → 見積書を確認
       → お客様へ注文書差し替え依頼メール作成・送信
       → 社内共有
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from enum import Enum
from typing import Optional

import anthropic

from agents import config
from agents.procurement_agent import prompts
from agents.procurement_agent.tools.erp_client import (
    ErpRegistrationResult,
    register_sales_order,
)
from agents.procurement_agent.tools.mail_client import (
    MailResult,
    send_mail,
    send_replacement_request,
)
from agents.sales_agent.sales_agent import NextAction, SalesAgentResult

logger = logging.getLogger(__name__)


# ============================================================
# フロー結果の型定義
# ============================================================
class ProcurementOutcome(str, Enum):
    ERP_REGISTERED     = "erp_registered"     # 販売管理登録完了
    REPLACEMENT_SENT   = "replacement_sent"   # 差し替え依頼送信完了
    REPLACEMENT_DRAFT  = "replacement_draft"  # 差し替え依頼下書き作成（送信前確認）
    ERROR              = "error"


@dataclass
class ProcurementAgentResult:
    """procurement_agent の処理結果"""
    outcome: ProcurementOutcome
    erp_result: Optional[ErpRegistrationResult] = None
    mail_result: Optional[MailResult] = None
    notification_message: str = ""  # 社内共有メッセージ
    agent_log: list[str] = field(default_factory=list)

    def summary(self) -> str:
        return (
            f"[ProcurementAgent] outcome={self.outcome.value} | "
            f"erp={self.erp_result} | mail={self.mail_result}"
        )


# ============================================================
# ツール定義（Claude tool_use スキーマ）
# ============================================================
TOOLS: list[dict] = [
    {
        "name": "register_to_erp",
        "description": (
            "注文書データを ERP（販売管理システム）に登録します。"
            "単価が正しいと確認された注文書のみ実行してください。"
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "approved_by": {
                    "type": "string",
                    "description": "押印・回覧した営業担当者名（わかれば）",
                    "default": "",
                },
            },
            "required": [],
        },
    },
    {
        "name": "create_replacement_request_mail",
        "description": (
            "単価相違があった場合にお客様へ送る注文書差し替え依頼メールの下書きを作成します。"
            "実際の送信前にユーザー確認を求めます。"
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "to": {
                    "type": "string",
                    "description": "送信先（お客様の資材担当者メールアドレス）",
                },
                "cc": {
                    "type": "string",
                    "description": "CC（社内担当者など、任意）",
                    "default": "",
                },
            },
            "required": ["to"],
        },
    },
    {
        "name": "send_replacement_request_mail",
        "description": (
            "作成済みの差し替え依頼メールを実際に送信します。"
            "必ず create_replacement_request_mail を先に実行してください。"
        ),
        "input_schema": {
            "type": "object",
            "properties": {},
            "required": [],
        },
    },
    {
        "name": "send_internal_notification",
        "description": "社内担当者へ処理完了・状況共有メッセージを送信します。",
        "input_schema": {
            "type": "object",
            "properties": {
                "to": {
                    "type": "string",
                    "description": "社内通知先メールアドレス",
                },
                "message": {
                    "type": "string",
                    "description": "通知本文",
                },
                "subject": {
                    "type": "string",
                    "description": "件名",
                },
            },
            "required": ["to", "message", "subject"],
        },
    },
    {
        "name": "report_completion",
        "description": "全処理が完了したことを記録して Agent を終了します。",
        "input_schema": {
            "type": "object",
            "properties": {
                "outcome": {
                    "type": "string",
                    "enum": [o.value for o in ProcurementOutcome],
                    "description": "処理結果の種別",
                },
                "summary": {
                    "type": "string",
                    "description": "完了サマリー（社内共有メッセージ）",
                },
            },
            "required": ["outcome", "summary"],
        },
    },
]


# ============================================================
# Agent 本体
# ============================================================
class ProcurementAgent:
    """
    資材購買アシスタント Agent。
    SalesAgentResult を受け取り、ERP 登録または差し替え依頼を実行する。

    使用例:
        agent = ProcurementAgent()
        result = agent.run(sales_result)
    """

    MAX_ITERATIONS = 10

    def __init__(self):
        self._client = anthropic.Anthropic(api_key=config.ANTHROPIC_API_KEY)
        self._sales_result: Optional[SalesAgentResult] = None
        self._mail_draft: Optional[dict] = None  # 下書き保存用
        self._erp_result: Optional[ErpRegistrationResult] = None
        self._mail_result: Optional[MailResult] = None
        self._final_result: Optional[dict] = None
        self._log: list[str] = []

    # ── public ──────────────────────────────────────────────
    def run(self, sales_result: SalesAgentResult) -> ProcurementAgentResult:
        """
        エントリーポイント。sales_agent の結果を受け取って処理を実行する。

        Args:
            sales_result: SalesAgent.run() の返却値

        Returns:
            ProcurementAgentResult
        """
        self._sales_result = sales_result
        self._log.append(
            f"[開始] next_action={sales_result.next_action.value} | "
            f"order={sales_result.order_doc.document_no}"
        )

        # フロー B: 単価 NG・未提出 は procurement_agent の担当外
        # (sales_agent が営業に通知済み)
        if sales_result.next_action == NextAction.NOTIFY_SALES:
            return ProcurementAgentResult(
                outcome=ProcurementOutcome.ERROR,
                notification_message=(
                    "見積書が未提出のため procurement_agent の処理対象外です。"
                    "営業担当者が見積書を提出後、再度フローを実行してください。"
                ),
                agent_log=self._log,
            )

        # ── Agent に処理させるためのプロンプト組み立て ──
        context = self._build_context_message()
        messages = [{"role": "user", "content": context}]

        for iteration in range(self.MAX_ITERATIONS):
            response = self._client.messages.create(
                model=config.CLAUDE_MODEL,
                max_tokens=4096,
                system=prompts.SYSTEM_PROMPT,
                tools=TOOLS,
                messages=messages,
            )
            self._log.append(
                f"[iter {iteration}] stop_reason={response.stop_reason}, "
                f"blocks={len(response.content)}"
            )

            tool_results = []
            for block in response.content:
                if block.type == "tool_use":
                    result = self._dispatch_tool(block.name, block.input)
                    tool_results.append(
                        {
                            "type": "tool_result",
                            "tool_use_id": block.id,
                            "content": json.dumps(result, ensure_ascii=False),
                        }
                    )
                    self._log.append(f"  tool={block.name} → {str(result)[:120]}")

                elif block.type == "text" and block.text:
                    self._log.append(f"  text={block.text[:80]}")

            if self._final_result:
                return self._build_result()

            if response.stop_reason == "end_turn" and not tool_results:
                logger.warning("Agent が report_completion を呼ばずに終了しました")
                break

            messages.append({"role": "assistant", "content": response.content})
            if tool_results:
                messages.append({"role": "user", "content": tool_results})

        raise RuntimeError(
            f"ProcurementAgent: {self.MAX_ITERATIONS} イテレーション内に完了しませんでした\n"
            + "\n".join(self._log[-5:])
        )

    # ── コンテキストメッセージ ───────────────────────────────
    def _build_context_message(self) -> str:
        sr = self._sales_result
        od = sr.order_doc
        qr = sr.quote_record

        base = (
            f"営業アシスタントから以下の情報を受け取りました。\n\n"
            f"【注文書情報】\n"
            f"　発注書番号: {od.document_no}\n"
            f"　発注元: {od.client_name}\n"
            f"　件名: {od.subject}\n"
            f"　発注日: {od.document_date}\n"
            f"　合計金額: ¥{od.total_amount:,.0f}\n\n"
        )

        if qr:
            base += (
                f"【紐づく見積書】\n"
                f"　見積番号: {qr.quote_no}\n"
                f"　ステータス: {qr.status}\n"
                f"　見積 PDF URL: {sr.quote_url or qr.quote_pdf_url}\n\n"
            )

        if sr.next_action == NextAction.APPROVAL_CIRCULATION:
            base += (
                "【次アクション】単価確認済み — 販売管理（ERP）への登録をお願いします。\n"
                "register_to_erp ツールを実行し、完了後に社内通知を送ってください。\n"
            )
        elif sr.next_action == NextAction.NOTIFY_PROCUREMENT:
            mismatches = sr.price_check.mismatches if sr.price_check else []
            mismatch_str = "\n".join(
                f"  ・{m['itemName']}: 注文書 ¥{m['orderPrice']:,.0f} → 見積書 ¥{m['quotePrice']:,.0f}"
                for m in mismatches
            )
            base += (
                f"【次アクション】単価相違あり — お客様へ差し替え依頼が必要です。\n"
                f"【単価相違明細】\n{mismatch_str}\n\n"
                f"create_replacement_request_mail でメール下書きを作成し、\n"
                f"内容を確認後 send_replacement_request_mail で送信してください。\n"
            )

        return base

    # ── tool dispatch ────────────────────────────────────────
    def _dispatch_tool(self, name: str, inputs: dict) -> dict:
        try:
            if name == "register_to_erp":
                return self._tool_register_erp(**inputs)
            if name == "create_replacement_request_mail":
                return self._tool_create_mail(**inputs)
            if name == "send_replacement_request_mail":
                return self._tool_send_mail()
            if name == "send_internal_notification":
                return self._tool_send_notification(**inputs)
            if name == "report_completion":
                return self._tool_report_completion(**inputs)
            return {"error": f"未知のツール: {name}"}
        except Exception as e:
            logger.exception("ツール %s で例外発生", name)
            return {"error": str(e)}

    # ── ツール実装 ───────────────────────────────────────────
    def _tool_register_erp(self, approved_by: str = "") -> dict:
        sr = self._sales_result
        self._erp_result = register_sales_order(
            order_doc=sr.order_doc,
            quote_record=sr.quote_record,
            approved_by=approved_by,
        )
        return {
            "success": self._erp_result.success,
            "erpOrderId": self._erp_result.erp_order_id,
            "message": self._erp_result.message,
            "dryRun": self._erp_result.dry_run,
        }

    def _tool_create_mail(self, to: str, cc: str = "") -> dict:
        sr = self._sales_result
        mismatches = sr.price_check.mismatches if sr.price_check else []
        from agents.procurement_agent.prompts import replacement_request_email
        body = replacement_request_email(
            order_doc=sr.order_doc,
            quote_record=sr.quote_record,
            price_mismatches=mismatches,
            quote_url=sr.quote_url,
        )
        subject = (
            f"【注文書差し替えのお願い】"
            f"{sr.order_doc.document_no} / {sr.order_doc.subject}"
        )
        self._mail_draft = {"to": to, "cc": cc, "subject": subject, "body": body}
        return {
            "success": True,
            "to": to,
            "subject": subject,
            "bodyPreview": body[:200] + "...",
            "message": "下書きを作成しました。send_replacement_request_mail で送信できます。",
        }

    def _tool_send_mail(self) -> dict:
        if not self._mail_draft:
            return {"error": "先に create_replacement_request_mail を実行してください"}
        d = self._mail_draft
        self._mail_result = send_mail(
            to=d["to"],
            subject=d["subject"],
            body=d["body"],
            cc=d.get("cc", ""),
        )
        return {
            "success": self._mail_result.success,
            "to": self._mail_result.to,
            "subject": self._mail_result.subject,
            "dryRun": self._mail_result.dry_run,
            "message": self._mail_result.message,
        }

    def _tool_send_notification(self, to: str, message: str, subject: str) -> dict:
        result = send_mail(to=to, subject=subject, body=message)
        return {
            "success": result.success,
            "dryRun": result.dry_run,
            "message": result.message,
        }

    def _tool_report_completion(self, outcome: str, summary: str) -> dict:
        self._final_result = {"outcome": outcome, "summary": summary}
        return {"success": True, "recorded": outcome}

    # ── 結果オブジェクト組み立て ─────────────────────────────
    def _build_result(self) -> ProcurementAgentResult:
        assert self._final_result is not None
        return ProcurementAgentResult(
            outcome=ProcurementOutcome(self._final_result["outcome"]),
            erp_result=self._erp_result,
            mail_result=self._mail_result,
            notification_message=self._final_result.get("summary", ""),
            agent_log=self._log,
        )
