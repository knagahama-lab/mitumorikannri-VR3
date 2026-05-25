"""
agents/procurement_agent/tools/erp_client.py

販売管理システム（ERP）への登録ツール。
注文書データを ERP API 経由で販売管理に投入します。

実際の ERP API に合わせて _post_to_erp() の中身を書き換えてください。
現状は「ドライラン」モード（DRY_RUN=True）がデフォルトです。
"""

from __future__ import annotations

import logging
import os
from dataclasses import dataclass
from datetime import datetime

import requests

from agents.config import ERP_API_KEY, ERP_BASE_URL, ERP_TIMEOUT

logger = logging.getLogger(__name__)

# DRY_RUN=True のときは ERP に実際のリクエストを送らず、ログだけ出す
DRY_RUN: bool = os.environ.get("ERP_DRY_RUN", "true").lower() == "true"


# ============================================================
# データクラス
# ============================================================
@dataclass
class ErpRegistrationResult:
    """ERP 登録結果"""

    success: bool
    erp_order_id: str = ""       # ERP 側で発番された受注 ID
    message: str = ""
    dry_run: bool = False

    def __str__(self):
        mode = "[DRY RUN] " if self.dry_run else ""
        return (
            f"{mode}{'✅' if self.success else '❌'} ERP登録: "
            f"erp_order_id={self.erp_order_id or '—'} / {self.message}"
        )


# ============================================================
# ERP 登録
# ============================================================
def register_sales_order(
    order_doc,         # OrderDocument
    quote_record=None, # QuoteRecord | None
    approved_by: str = "",
) -> ErpRegistrationResult:
    """
    注文書データを販売管理（ERP）に登録する。

    Args:
        order_doc: OCR 解析済みの OrderDocument
        quote_record: 紐づく見積書レコード（あれば）
        approved_by: 押印・回覧した営業担当者名

    Returns:
        ErpRegistrationResult
    """
    payload = _build_erp_payload(order_doc, quote_record, approved_by)

    if DRY_RUN:
        logger.info("[ERP DRY RUN] 以下のデータを登録予定:\n%s", payload)
        return ErpRegistrationResult(
            success=True,
            erp_order_id=f"DRY-{int(datetime.now().timestamp())}",
            message="ドライランモード: ERP への実際の登録はスキップされました",
            dry_run=True,
        )

    return _post_to_erp(payload)


def _build_erp_payload(order_doc, quote_record, approved_by: str) -> dict:
    """ERP API 用のリクエストボディを組み立てる"""
    return {
        "order_no": order_doc.document_no,
        "order_date": order_doc.document_date,
        "client_name": order_doc.client_name,
        "subject": order_doc.subject,
        "model_code": order_doc.model_code,
        "order_slip_no": order_doc.order_slip_no,
        "order_type": order_doc.order_type,
        "subtotal": order_doc.subtotal,
        "tax": order_doc.tax,
        "total_amount": order_doc.total_amount,
        "linked_quote_no": (quote_record.quote_no if quote_record else ""),
        "approved_by": approved_by,
        "registered_at": datetime.now().isoformat(),
        "line_items": [
            {
                "item_name": li.item_name,
                "spec": li.spec,
                "qty": li.qty,
                "unit": li.unit,
                "unit_price": li.unit_price,
                "amount": li.amount,
                "first_delivery": li.first_delivery,
                "delivery_dest": li.delivery_dest,
                "remarks": li.remarks,
            }
            for li in order_doc.line_items
        ],
    }


def _post_to_erp(payload: dict) -> ErpRegistrationResult:
    """ERP API への実際の POST リクエスト（本番実装）"""
    if not ERP_BASE_URL or "example.com" in ERP_BASE_URL:
        raise EnvironmentError(
            "ERP_BASE_URL が設定されていません。config.py または環境変数を確認してください。"
        )

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {ERP_API_KEY}",
    }
    try:
        resp = requests.post(
            f"{ERP_BASE_URL}/sales_orders",
            json=payload,
            headers=headers,
            timeout=ERP_TIMEOUT,
        )
        resp.raise_for_status()
        data = resp.json()
        return ErpRegistrationResult(
            success=True,
            erp_order_id=str(data.get("id") or data.get("order_id") or ""),
            message=data.get("message", "登録完了"),
        )
    except requests.HTTPError as e:
        logger.error("ERP API HTTPエラー: %s — %s", e, e.response.text[:200])
        return ErpRegistrationResult(
            success=False,
            message=f"HTTP {e.response.status_code}: {e.response.text[:100]}",
        )
    except requests.RequestException as e:
        logger.error("ERP API 通信エラー: %s", e)
        return ErpRegistrationResult(success=False, message=str(e))
