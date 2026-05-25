"""
agents/sales_agent/tools/system_api.py

見積・注文管理システム（GAS Web App）との通信ツール。
- 見積書の検索・URL 取得
- 注文書の照合（単価チェック）
- 管理シートへのデータ参照
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Optional

import requests

from agents.config import GAS_WEB_APP_URL, GAS_REQUEST_TIMEOUT, PRICE_TOLERANCE_RATE

logger = logging.getLogger(__name__)


# ============================================================
# データクラス
# ============================================================
@dataclass
class QuoteRecord:
    """管理システムから取得した見積書レコード"""

    mgmt_id: str
    quote_no: str
    subject: str
    client_name: str
    issue_date: str
    total_amount: float
    quote_pdf_url: str
    status: str
    linked: bool
    order_no: str = ""
    order_pdf_url: str = ""
    model_code: str = ""


@dataclass
class PriceCheckResult:
    """単価チェック結果"""

    is_correct: bool
    mismatches: list[dict] = field(default_factory=list)  # 不一致明細のリスト
    matched_quote_no: str = ""
    message: str = ""

    def summary(self) -> str:
        if self.is_correct:
            return f"✅ 単価正常（見積書: {self.matched_quote_no}）"
        lines = "\n".join(
            f"  ・{m['itemName']}: 注文書 {m['orderPrice']:,.0f}円 ≠ 見積書 {m['quotePrice']:,.0f}円"
            for m in self.mismatches
        )
        return f"❌ 単価相違 {len(self.mismatches)}件:\n{lines}"


# ============================================================
# 管理システム API クライアント
# ============================================================
def _call_gas(action: str, payload: dict | None = None) -> dict:
    """GAS Web App の handleApiRequest を呼び出す汎用関数"""
    body = {"action": action, "payload": payload or {}}
    try:
        resp = requests.post(
            GAS_WEB_APP_URL,
            json=body,
            timeout=GAS_REQUEST_TIMEOUT,
        )
        resp.raise_for_status()
        data = resp.json()
    except requests.RequestException as e:
        raise RuntimeError(f"GAS API 通信エラー [{action}]: {e}") from e

    if not data.get("success"):
        raise RuntimeError(f"GAS API エラー [{action}]: {data.get('error', '不明')}")
    return data


def get_quote_by_client(
    client_name: str,
    model_code: str = "",
    limit: int = 5,
) -> list[QuoteRecord]:
    """
    顧客名・機種コードで見積書を検索して返す。

    Args:
        client_name: 顧客名（部分一致）
        model_code: 機種コード（空文字なら不問）
        limit: 最大件数

    Returns:
        マッチした見積書レコードのリスト（新しい順）
    """
    data = _call_gas("quoteListGetAll", {})
    items: list[dict] = data.get("items", [])

    results = []
    client_kw = client_name.strip().lower()
    model_kw = model_code.strip().lower()

    for item in items:
        # 顧客名フィルター
        dest = (item.get("destCompany") or item.get("client") or "").lower()
        if client_kw and client_kw not in dest:
            continue
        # 機種コードフィルター
        if model_kw and model_kw not in (item.get("modelCode") or "").lower():
            continue
        results.append(
            QuoteRecord(
                mgmt_id=item.get("id", ""),
                quote_no=item.get("quoteNo", ""),
                subject=item.get("subject", ""),
                client_name=item.get("destCompany") or item.get("client", ""),
                issue_date=item.get("issueDate", ""),
                total_amount=float(item.get("quoteAmount") or 0),
                quote_pdf_url=item.get("quotePdfUrl", ""),
                status=item.get("status", ""),
                linked=bool(item.get("linked")),
                order_no=item.get("orderNo", ""),
                order_pdf_url=item.get("orderPdfUrl", ""),
                model_code=item.get("modelCode", ""),
            )
        )

    # 新しい順にソート、件数制限
    results.sort(key=lambda r: r.issue_date, reverse=True)
    return results[:limit]


def get_quote_detail(mgmt_id: str) -> dict:
    """
    見積書の詳細（明細行含む）を取得する。

    Returns:
        GAS getQuoteDetail API の生レスポンス dict
    """
    return _call_gas("getQuoteDetail", {"mgmtId": mgmt_id})


def get_quote_url(mgmt_id: str) -> str:
    """
    指定した管理 ID の見積書 PDF URL を返す。
    URL が空の場合は空文字を返す。
    """
    detail = get_quote_detail(mgmt_id)
    mgmt = detail.get("mgmt") or {}
    return mgmt.get("quotePdfUrl", "")


def check_order_prices(
    order_doc,  # OrderDocument
    quote_detail: dict,
) -> PriceCheckResult:
    """
    注文書の単価 vs 見積書の単価を照合する。

    照合ルール:
      - 見積書の lineItems と注文書の lineItems を品名で突き合わせる
      - 単価の差が PRICE_TOLERANCE_RATE を超えたらアウト
      - 見積書に存在しない品名はスキップ（新規品目扱い）

    Args:
        order_doc: OCR 結果の OrderDocument
        quote_detail: get_quote_detail() の返却値

    Returns:
        PriceCheckResult
    """
    # 見積書明細を品名で辞書化
    quote_lines: list[dict] = quote_detail.get("quoteLines") or []
    quote_price_map: dict[str, float] = {}
    for line in quote_lines:
        name = (line.get("itemName") or "").strip()
        price = float(line.get("unitPrice") or 0)
        if name and price > 0:
            quote_price_map[name] = price

    mismatches = []
    for item in order_doc.line_items:
        name = item.item_name.strip()
        order_price = item.unit_price
        if not name or order_price <= 0:
            continue
        if name not in quote_price_map:
            logger.debug("照合スキップ（見積書に存在しない品目）: %s", name)
            continue
        quote_price = quote_price_map[name]
        diff_rate = abs(order_price - quote_price) / quote_price if quote_price else 1.0
        if diff_rate > PRICE_TOLERANCE_RATE:
            mismatches.append(
                {
                    "itemName": name,
                    "orderPrice": order_price,
                    "quotePrice": quote_price,
                    "diffRate": round(diff_rate * 100, 2),
                }
            )

    is_correct = len(mismatches) == 0
    return PriceCheckResult(
        is_correct=is_correct,
        mismatches=mismatches,
        matched_quote_no=quote_detail.get("mgmt", {}).get("quoteNo", ""),
        message="単価一致" if is_correct else f"{len(mismatches)}件の単価相違あり",
    )


def find_linked_quote(order_doc) -> Optional[QuoteRecord]:
    """
    注文書に紐づく見積書を管理システムから探す。

    優先順位:
      1. 注文書に linkedQuoteNo が記載されている
      2. 顧客名 + 機種コードで直近の見積書を検索

    Returns:
        QuoteRecord or None（見つからない場合）
    """
    # ① 注文書の記載番号で直接検索
    if order_doc.linked_quote_no:
        data = _call_gas("quoteListGetAll", {})
        for item in data.get("items", []):
            if item.get("quoteNo", "") == order_doc.linked_quote_no:
                return QuoteRecord(
                    mgmt_id=item["id"],
                    quote_no=item["quoteNo"],
                    subject=item.get("subject", ""),
                    client_name=item.get("destCompany") or item.get("client", ""),
                    issue_date=item.get("issueDate", ""),
                    total_amount=float(item.get("quoteAmount") or 0),
                    quote_pdf_url=item.get("quotePdfUrl", ""),
                    status=item.get("status", ""),
                    linked=bool(item.get("linked")),
                    order_no=item.get("orderNo", ""),
                    order_pdf_url=item.get("orderPdfUrl", ""),
                    model_code=item.get("modelCode", ""),
                )

    # ② 顧客名 + 機種コードで検索
    candidates = get_quote_by_client(
        client_name=order_doc.client_name,
        model_code=order_doc.model_code,
        limit=1,
    )
    return candidates[0] if candidates else None


def is_quote_submitted(quote_record: QuoteRecord) -> bool:
    """
    見積書が顧客へ提出済みかどうか。
    ステータスが「送信済み」「受領」「受注済み」「納品済み」なら提出済みとみなす。
    """
    submitted_statuses = {"送信済み", "受領", "受注済み", "納品済み", "受領（差し替え）"}
    return quote_record.status in submitted_statuses
