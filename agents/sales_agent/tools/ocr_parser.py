"""
agents/sales_agent/tools/ocr_parser.py

注文書 PDF から単価・品名・数量・発注番号などを抽出するツール。
Gemini Vision API を使用した OCR 処理と、
GAS 側（15_ocr_review_ui.gs）の ocrPreview API の両方をサポートします。
"""

import base64
import json
import re
from pathlib import Path
from typing import Optional

import requests

from agents.config import (
    GAS_WEB_APP_URL,
    GAS_REQUEST_TIMEOUT,
    GEMINI_API_KEY,
    GEMINI_MODEL,
)


# ============================================================
# データクラス
# ============================================================
class LineItem:
    """注文書の明細行 1 件分"""

    def __init__(self, data: dict):
        self.item_name: str = data.get("itemName", "")
        self.spec: str = data.get("spec", "")
        self.qty: float = float(data.get("qty") or 0)
        self.unit: str = data.get("unit", "")
        self.unit_price: float = float(data.get("unitPrice") or 0)
        self.amount: float = float(data.get("amount") or 0)
        self.first_delivery: str = data.get("firstDelivery", "")
        self.delivery_dest: str = data.get("deliveryDest", "")
        self.remarks: str = data.get("remarks", "")

    def to_dict(self) -> dict:
        return {
            "itemName": self.item_name,
            "spec": self.spec,
            "qty": self.qty,
            "unit": self.unit,
            "unitPrice": self.unit_price,
            "amount": self.amount,
            "firstDelivery": self.first_delivery,
            "deliveryDest": self.delivery_dest,
            "remarks": self.remarks,
        }

    def __repr__(self):
        return (
            f"LineItem({self.item_name!r}, qty={self.qty}, "
            f"unitPrice={self.unit_price:,.0f})"
        )


class OrderDocument:
    """注文書 OCR 結果全体"""

    def __init__(self, data: dict):
        self.document_no: str = data.get("documentNo", "")
        self.document_date: str = data.get("documentDate", "")
        self.client_name: str = data.get("clientName", "")
        self.subject: str = data.get("subject", "")
        self.model_code: str = data.get("modelCode", "")
        self.order_slip_no: str = data.get("orderSlipNo", "")
        self.linked_quote_no: str = data.get("linkedQuoteNo", "")
        self.order_type: str = data.get("orderType", "")
        self.subtotal: float = float(data.get("subtotal") or 0)
        self.tax: float = float(data.get("tax") or 0)
        self.total_amount: float = float(data.get("totalAmount") or 0)
        self.action_type: str = data.get("actionType", "new")
        self.reason: str = data.get("reason", "")
        self.line_items: list[LineItem] = [
            LineItem(item) for item in (data.get("lineItems") or [])
        ]
        # OCR メタ情報
        self.ocr_quality: int = int(data.get("quality") or 0)
        self.pdf_url: str = data.get("pdfUrl", "")
        self.session_id: str = data.get("sessionId", "")
        self.raw: dict = data  # 元データ保持

    @property
    def is_valid(self) -> bool:
        return bool(self.document_no and self.line_items)

    def __repr__(self):
        return (
            f"OrderDocument({self.document_no!r}, "
            f"lines={len(self.line_items)}, total={self.total_amount:,.0f})"
        )


# ============================================================
# GAS Web App 経由での OCR（推奨）
# ============================================================
def parse_order_pdf_via_gas(
    pdf_path: str | Path,
    order_type: str = "",
) -> OrderDocument:
    """
    PDF を GAS の ocrPreview API に送信し、注文書データを返す。

    GAS 側の 15_ocr_review_ui.gs::apiOcrPreview() を呼び出し、
    Gemini OCR の結果をそのまま受け取る。

    Args:
        pdf_path: ローカルの PDF ファイルパス
        order_type: 注文種別（"試作" / "量産" / ""）

    Returns:
        OrderDocument: OCR 結果

    Raises:
        RuntimeError: OCR 失敗 or 通信エラー
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF が見つかりません: {pdf_path}")

    # base64 エンコード
    b64 = base64.b64encode(pdf_path.read_bytes()).decode("utf-8")

    payload = {
        "action": "ocrPreview",
        "payload": {
            "base64Data": b64,
            "fileName": pdf_path.name,
            "docType": "order",
            "orderType": order_type,
        },
    }

    resp = requests.post(
        GAS_WEB_APP_URL,
        json=payload,
        timeout=GAS_REQUEST_TIMEOUT,
    )
    resp.raise_for_status()
    data = resp.json()

    if not data.get("success"):
        raise RuntimeError(f"GAS OCR 失敗: {data.get('error', '不明なエラー')}")

    ocr_result: dict = data.get("ocrResult", {})
    ocr_result["quality"] = data.get("quality", 0)
    ocr_result["pdfUrl"] = data.get("pdfUrl", "")
    ocr_result["sessionId"] = data.get("sessionId", "")

    return OrderDocument(ocr_result)


# ============================================================
# Gemini API 直接呼び出しでの OCR（GAS 未経由）
# ============================================================
def parse_order_pdf_direct(
    pdf_path: str | Path,
) -> OrderDocument:
    """
    Gemini API を直接呼び出して注文書 PDF を OCR する。
    GAS を経由しないため、Python 単独で動作させたい場合に使用。

    Args:
        pdf_path: ローカルの PDF ファイルパス

    Returns:
        OrderDocument

    Raises:
        RuntimeError: OCR 失敗 or API エラー
    """
    if not GEMINI_API_KEY:
        raise EnvironmentError("GEMINI_API_KEY が設定されていません")

    pdf_path = Path(pdf_path)
    b64 = base64.b64encode(pdf_path.read_bytes()).decode("utf-8")

    prompt = _build_order_ocr_prompt()
    url = (
        f"https://generativelanguage.googleapis.com/v1beta/models/"
        f"{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
    )
    body = {
        "contents": [
            {
                "parts": [
                    {"text": prompt},
                    {
                        "inline_data": {
                            "mime_type": "application/pdf",
                            "data": b64,
                        }
                    },
                ]
            }
        ],
        "generationConfig": {"temperature": 0},
    }

    resp = requests.post(url, json=body, timeout=GAS_REQUEST_TIMEOUT)
    resp.raise_for_status()

    raw_text: str = (
        resp.json()
        .get("candidates", [{}])[0]
        .get("content", {})
        .get("parts", [{}])[0]
        .get("text", "")
    )

    ocr_data = _extract_json(raw_text)
    return OrderDocument(ocr_data)


# ============================================================
# ユーティリティ
# ============================================================
def _build_order_ocr_prompt() -> str:
    return (
        "あなたはOCR専門家です。添付PDF（発注書・注文書）を解析し、"
        "以下のJSON形式のみで返してください。説明文不要。\n"
        '{\n'
        ' "documentNo": "発注書番号",\n'
        ' "documentDate": "発注日(YYYY/MM/DD)",\n'
        ' "clientName": "発注元企業名",\n'
        ' "subject": "件名",\n'
        ' "modelCode": "機種コード（なければ空文字）",\n'
        ' "orderSlipNo": "伝票番号（なければ空文字）",\n'
        ' "linkedQuoteNo": "紐づく見積番号（なければ空文字）",\n'
        ' "orderType": "試作 または 量産（不明なら空文字）",\n'
        ' "subtotal": 小計(数値),\n'
        ' "tax": 消費税(数値),\n'
        ' "totalAmount": 合計(数値),\n'
        ' "lineItems": [\n'
        '   {"itemName":"品名","spec":"仕様","firstDelivery":"初回納入日(YYYY/MM/DD)",'
        '"deliveryDest":"納入先","qty":数量,"unit":"単位","unitPrice":単価,"amount":金額,'
        '"remarks":"備考"}\n'
        ' ]\n'
        '}\n'
        "ルール: 有効なJSONのみ。金額は数値。不明は空文字か0。合計行はlineItemsに含めない。"
    )


def _extract_json(text: str) -> dict:
    """LLM の応答テキストから JSON ブロックを抽出してパース"""
    # ```json ... ``` ブロックがある場合
    match = re.search(r"```(?:json)?\s*([\s\S]+?)\s*```", text)
    if match:
        text = match.group(1)
    # 先頭・末尾の空白を除去して直接パース
    text = text.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError as e:
        raise RuntimeError(f"OCR レスポンスの JSON パース失敗: {e}\n---\n{text[:300]}")
