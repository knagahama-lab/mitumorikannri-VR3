"""
agents/sales_agent/sales_agent.py

営業アシスタント Agent のメインロジック。

フロー担当範囲:
  [注文書受領]
      → OCR で単価を抽出
      → 管理システムから見積書を検索・単価照合
      → [単価 OK] → 押印・回覧依頼メッセージを生成 → FlowResult 返却
      → [単価 NG]
            → 見積書が提出済み → 見積 URL を取得 → FlowResult 返却（procurement_agent へ渡す）
            → 見積書が未提出  → 通知メール下書き → FlowResult 返却（営業へ通知）
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Optional

import anthropic

from agents import config
from agents.sales_agent import prompts
from agents.sales_agent.tools.ocr_parser import OrderDocument, parse_order_pdf_via_gas
from agents.sales_agent.tools.system_api import (
    PriceCheckResult,
    QuoteRecord,
    check_order_prices,
    find_linked_quote,
    get_quote_detail,
    get_quote_url,
    is_quote_submitted,
)

logger = logging.getLogger(__name__)


# ============================================================
# フロー結果の型定義
# ============================================================
class NextAction(str, Enum):
    """sales_agent が判定した次アクション"""
    APPROVAL_CIRCULATION = "approval_circulation"   # 単価OK → 押印・回覧
    NOTIFY_PROCUREMENT   = "notify_procurement"     # 単価NG・提出済 → 購買へURL連絡
    NOTIFY_SALES         = "notify_sales"           # 単価NG・未提出 → 営業へタスク通知


@dataclass
class SalesAgentResult:
    """sales_agent の処理結果（run_flow.py へ返す）"""
    next_action: NextAction
    order_doc: OrderDocument
    quote_record: Optional[QuoteRecord] = None
    price_check: Optional[PriceCheckResult] = None
    quote_url: str = ""
    message: str = ""               # 生成したメッセージ本文（メールや通知）
    subject: str = ""               # メール件名
    agent_log: list[str] = field(default_factory=list)  # Agent の思考ログ

    def summary(self) -> str:
        return (
            f"[SalesAgent] next={self.next_action.value} | "
            f"order={self.order_doc.document_no} | "
            f"quote={self.quote_record.quote_no if self.quote_record else '—'}"
        )


# ============================================================
# ツール定義（Claude tool_use スキーマ）
# ============================================================
TOOLS: list[dict] = [
    {
        "name": "ocr_order_pdf",
        "description": "注文書 PDF を OCR 解析して発注情報（単価・品名・数量など）を抽出します。",
        "input_schema": {
            "type": "object",
            "properties": {
                "pdf_path": {
                    "type": "string",
                    "description": "解析対象の PDF ファイルの絶対パスまたは相対パス",
                },
                "order_type": {
                    "type": "string",
                    "description": "注文種別: '試作' / '量産' / '' (不明)",
                    "default": "",
                },
            },
            "required": ["pdf_path"],
        },
    },
    {
        "name": "find_quote_for_order",
        "description": "注文書に対応する見積書を管理システムから検索します。",
        "input_schema": {
            "type": "object",
            "properties": {
                "client_name": {
                    "type": "string",
                    "description": "発注元の会社名（部分一致で検索）",
                },
                "model_code": {
                    "type": "string",
                    "description": "機種コード（なければ空文字）",
                    "default": "",
                },
                "linked_quote_no": {
                    "type": "string",
                    "description": "注文書に記載されている見積番号（あれば）",
                    "default": "",
                },
            },
            "required": ["client_name"],
        },
    },
    {
        "name": "check_unit_prices",
        "description": "注文書の単価と見積書の単価を照合します。",
        "input_schema": {
            "type": "object",
            "properties": {
                "mgmt_id": {
                    "type": "string",
                    "description": "照合する見積書の管理 ID",
                },
            },
            "required": ["mgmt_id"],
        },
    },
    {
        "name": "get_quote_pdf_url",
        "description": "指定した管理 ID の見積書 PDF URL を取得します。",
        "input_schema": {
            "type": "object",
            "properties": {
                "mgmt_id": {
                    "type": "string",
                    "description": "見積書の管理 ID",
                },
            },
            "required": ["mgmt_id"],
        },
    },
    {
        "name": "report_result",
        "description": (
            "フロー判定結果を確定して終了します。"
            "次アクション・メッセージ本文・メール件名を指定してください。"
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "next_action": {
                    "type": "string",
                    "enum": [a.value for a in NextAction],
                    "description": "次のアクション種別",
                },
                "message": {
                    "type": "string",
                    "description": "通知・メール本文（Markdown 可）",
                },
                "subject": {
                    "type": "string",
                    "description": "メール件名（通知メッセージのタイトル）",
                    "default": "",
                },
            },
            "required": ["next_action", "message"],
        },
    },
]


# ============================================================
# Agent 本体
# ============================================================
class SalesAgent:
    """
    注文書の単価チェックを担当する営業アシスタント Agent。

    使用例:
        agent = SalesAgent()
        result = agent.run("path/to/order.pdf", order_type="試作")
    """

    MAX_ITERATIONS = 10  # 無限ループ防止

    def __init__(self):
        self._client = anthropic.Anthropic(api_key=config.ANTHROPIC_API_KEY)
        self._order_doc: Optional[OrderDocument] = None
        self._quote_record: Optional[QuoteRecord] = None
        self._price_check: Optional[PriceCheckResult] = None
        self._quote_url: str = ""
        self._final_result: Optional[dict] = None
        self._log: list[str] = []

    # ── public ──────────────────────────────────────────────
    def run(self, pdf_path: str, order_type: str = "") -> SalesAgentResult:
        """
        エントリーポイント。PDF パスを受け取ってフローを実行し結果を返す。

        Args:
            pdf_path: 注文書 PDF のパス
            order_type: 注文種別（"試作" / "量産" / ""）

        Returns:
            SalesAgentResult
        """
        self._log.append(f"[開始] pdf={pdf_path}, order_type={order_type}")

        messages = [
            {
                "role": "user",
                "content": (
                    f"以下の注文書 PDF を処理してください。\n"
                    f"PDFパス: {pdf_path}\n"
                    f"注文種別: {order_type or '不明'}\n\n"
                    "まず OCR で注文書を解析し、次に管理システムで見積書を検索して"
                    "単価を照合してください。"
                ),
            }
        ]

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

            # ── tool_use ブロックを処理 ──
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

            # ── 終了条件チェック ──
            if self._final_result:
                return self._build_result()

            if response.stop_reason == "end_turn" and not tool_results:
                # ツールを呼ばずに終了 → report_result が呼ばれなかった
                logger.warning("Agent が report_result を呼ばずに終了しました")
                break

            # ── 次ターンへメッセージを追加 ──
            messages.append({"role": "assistant", "content": response.content})
            if tool_results:
                messages.append({"role": "user", "content": tool_results})

        # MAX_ITERATIONS 超過 or 異常終了
        raise RuntimeError(
            f"SalesAgent: {self.MAX_ITERATIONS} イテレーション内に完了しませんでした\n"
            + "\n".join(self._log[-5:])
        )

    # ── tool dispatch ────────────────────────────────────────
    def _dispatch_tool(self, name: str, inputs: dict) -> dict:
        try:
            if name == "ocr_order_pdf":
                return self._tool_ocr_order_pdf(**inputs)
            if name == "find_quote_for_order":
                return self._tool_find_quote(**inputs)
            if name == "check_unit_prices":
                return self._tool_check_prices(**inputs)
            if name == "get_quote_pdf_url":
                return self._tool_get_quote_url(**inputs)
            if name == "report_result":
                return self._tool_report_result(**inputs)
            return {"error": f"未知のツール: {name}"}
        except Exception as e:
            logger.exception("ツール %s で例外発生", name)
            return {"error": str(e)}

    # ── ツール実装 ───────────────────────────────────────────
    def _tool_ocr_order_pdf(self, pdf_path: str, order_type: str = "") -> dict:
        self._order_doc = parse_order_pdf_via_gas(pdf_path, order_type)
        return {
            "success": True,
            "documentNo": self._order_doc.document_no,
            "documentDate": self._order_doc.document_date,
            "clientName": self._order_doc.client_name,
            "subject": self._order_doc.subject,
            "modelCode": self._order_doc.model_code,
            "linkedQuoteNo": self._order_doc.linked_quote_no,
            "totalAmount": self._order_doc.total_amount,
            "lineCount": len(self._order_doc.line_items),
            "ocrQuality": self._order_doc.ocr_quality,
            "lineItems": [li.to_dict() for li in self._order_doc.line_items],
        }

    def _tool_find_quote(
        self,
        client_name: str,
        model_code: str = "",
        linked_quote_no: str = "",
    ) -> dict:
        if not self._order_doc:
            return {"error": "先に ocr_order_pdf を実行してください"}

        # linked_quote_no が渡された場合は order_doc に設定してから検索
        if linked_quote_no:
            self._order_doc.linked_quote_no = linked_quote_no

        self._quote_record = find_linked_quote(self._order_doc)
        if not self._quote_record:
            return {
                "success": False,
                "message": "対応する見積書が見つかりませんでした",
                "isSubmitted": False,
            }

        submitted = is_quote_submitted(self._quote_record)
        return {
            "success": True,
            "mgmtId": self._quote_record.mgmt_id,
            "quoteNo": self._quote_record.quote_no,
            "subject": self._quote_record.subject,
            "clientName": self._quote_record.client_name,
            "issueDate": self._quote_record.issue_date,
            "totalAmount": self._quote_record.total_amount,
            "status": self._quote_record.status,
            "isSubmitted": submitted,
            "quotePdfUrl": self._quote_record.quote_pdf_url,
        }

    def _tool_check_prices(self, mgmt_id: str) -> dict:
        if not self._order_doc:
            return {"error": "先に ocr_order_pdf を実行してください"}

        detail = get_quote_detail(mgmt_id)
        self._price_check = check_order_prices(self._order_doc, detail)
        return {
            "isCorrect": self._price_check.is_correct,
            "mismatches": self._price_check.mismatches,
            "matchedQuoteNo": self._price_check.matched_quote_no,
            "message": self._price_check.message,
        }

    def _tool_get_quote_url(self, mgmt_id: str) -> dict:
        url = get_quote_url(mgmt_id)
        self._quote_url = url
        return {"success": True, "quotePdfUrl": url}

    def _tool_report_result(
        self,
        next_action: str,
        message: str,
        subject: str = "",
    ) -> dict:
        self._final_result = {
            "next_action": next_action,
            "message": message,
            "subject": subject,
        }
        return {"success": True, "recorded": next_action}

    # ── 結果オブジェクト組み立て ─────────────────────────────
    def _build_result(self) -> SalesAgentResult:
        assert self._final_result is not None
        assert self._order_doc is not None
        return SalesAgentResult(
            next_action=NextAction(self._final_result["next_action"]),
            order_doc=self._order_doc,
            quote_record=self._quote_record,
            price_check=self._price_check,
            quote_url=self._quote_url,
            message=self._final_result.get("message", ""),
            subject=self._final_result.get("subject", ""),
            agent_log=self._log,
        )
