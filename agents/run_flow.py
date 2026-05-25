"""
agents/run_flow.py

注文書処理フロー全体のエントリーポイント。

使用例:
    # ① コマンドラインから直接実行
    python -m agents.run_flow path/to/order.pdf --order-type 試作

    # ② Python から呼び出す
    from agents.run_flow import run
    run("path/to/order.pdf", order_type="量産")

フロー:
    [注文書受領]
        ↓
    SalesAgent: OCR → 見積書検索 → 単価照合
        ↓
    ┌─ 単価OK  → ProcurementAgent: ERP登録 → 登録完了通知
    │
    ├─ 単価NG・提出済み → ProcurementAgent: 差し替え依頼メール作成・送信
    │
    └─ 単価NG・未提出  → 営業へ通知メール送信（procurement_agent は待機）
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

# ── ロギング設定（run_flow.py から起動する場合） ──
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


def run(
    pdf_path: str,
    order_type: str = "",
    customer_mail: str = "",
    internal_notify_mail: str = "",
    verbose: bool = False,
) -> dict:
    """
    注文書処理フローのメイン関数。

    Args:
        pdf_path: 注文書 PDF のパス（ローカルまたはアクセス可能なパス）
        order_type: 注文種別 "試作" / "量産" / ""（不明）
        customer_mail: お客様資材担当者のメールアドレス（差し替え依頼送信先）
        internal_notify_mail: 社内通知先メールアドレス
        verbose: True にすると Agent の思考ログを表示

    Returns:
        {
            "status": "ok" | "error",
            "next_action": str,
            "sales_summary": str,
            "procurement_summary": str | None,
            "details": dict,
        }
    """
    # ── 設定バリデーション ──
    from agents import config
    missing = config.validate()
    if missing:
        logger.error("必須環境変数が未設定: %s", missing)
        return {
            "status": "error",
            "message": f"環境変数が未設定です: {missing}",
        }

    # ── Step 1: SalesAgent 実行 ──
    logger.info("=" * 60)
    logger.info("STEP 1: SalesAgent 実行開始")
    logger.info("  PDF: %s", pdf_path)
    logger.info("  種別: %s", order_type or "不明")
    logger.info("=" * 60)

    from agents.sales_agent.sales_agent import SalesAgent, NextAction
    sales_agent = SalesAgent()

    try:
        sales_result = sales_agent.run(pdf_path, order_type=order_type)
    except Exception as e:
        logger.exception("SalesAgent 実行エラー")
        return {"status": "error", "message": f"SalesAgent エラー: {e}"}

    if verbose:
        _print_log("SalesAgent ログ", sales_result.agent_log)

    logger.info("SalesAgent 完了: %s", sales_result.summary())

    # ── フロー C: 単価NG・未提出 → 営業通知 ──
    if sales_result.next_action == NextAction.NOTIFY_SALES:
        logger.info("→ 単価NG・見積未提出: 営業担当者へ通知")
        _send_sales_notification(sales_result, internal_notify_mail)
        return {
            "status": "ok",
            "next_action": sales_result.next_action.value,
            "sales_summary": sales_result.summary(),
            "procurement_summary": None,
            "details": {
                "order_no": sales_result.order_doc.document_no,
                "message": sales_result.message,
                "note": "見積書が未提出のため、営業担当者へタスク通知を送信しました。",
            },
        }

    # ── Step 2: ProcurementAgent 実行 ──
    logger.info("=" * 60)
    logger.info("STEP 2: ProcurementAgent 実行開始")
    logger.info("  next_action: %s", sales_result.next_action.value)
    logger.info("=" * 60)

    from agents.procurement_agent.procurement_agent import ProcurementAgent
    proc_agent = ProcurementAgent()

    # customer_mail が指定されている場合、差し替え依頼の宛先として注入
    # （実際の運用では GAS 管理システムや顧客マスタから取得する）
    if customer_mail and sales_result.next_action == NextAction.NOTIFY_PROCUREMENT:
        # procurement_agent のコンテキストに宛先を伝える方法として
        # sales_result の message に埋め込むか、別途設定する
        sales_result.order_doc.raw["_customer_mail"] = customer_mail

    try:
        proc_result = proc_agent.run(sales_result)
    except Exception as e:
        logger.exception("ProcurementAgent 実行エラー")
        return {
            "status": "error",
            "message": f"ProcurementAgent エラー: {e}",
            "sales_summary": sales_result.summary(),
        }

    if verbose:
        _print_log("ProcurementAgent ログ", proc_result.agent_log)

    logger.info("ProcurementAgent 完了: %s", proc_result.summary())

    # ── 社内共有通知 ──
    if internal_notify_mail and proc_result.notification_message:
        _send_internal_notification(
            to=internal_notify_mail,
            subject=f"【処理完了】注文書 {sales_result.order_doc.document_no}",
            message=proc_result.notification_message,
        )

    # ── 結果を返す ──
    return {
        "status": "ok",
        "next_action": sales_result.next_action.value,
        "sales_summary": sales_result.summary(),
        "procurement_summary": proc_result.summary(),
        "details": {
            "order_no": sales_result.order_doc.document_no,
            "order_client": sales_result.order_doc.client_name,
            "quote_no": sales_result.quote_record.quote_no if sales_result.quote_record else None,
            "erp_order_id": proc_result.erp_result.erp_order_id if proc_result.erp_result else None,
            "mail_sent_to": proc_result.mail_result.to if proc_result.mail_result else None,
            "outcome": proc_result.outcome.value,
        },
    }


# ============================================================
# 内部ヘルパー
# ============================================================
def _send_sales_notification(sales_result, notify_to: str):
    """営業担当者へのタスク通知メール送信"""
    from agents import config
    from agents.procurement_agent.tools.mail_client import send_mail

    to = notify_to or config.SALES_NOTIFY_EMAIL
    if not to:
        logger.warning("営業通知先メールアドレスが未設定のためスキップ")
        return

    send_mail(
        to=to,
        subject=(
            f"【対応依頼】見積書未提出 — "
            f"{sales_result.order_doc.document_no} / "
            f"{sales_result.order_doc.client_name}"
        ),
        body=sales_result.message,
    )


def _send_internal_notification(to: str, subject: str, message: str):
    """社内担当者への完了通知"""
    from agents.procurement_agent.tools.mail_client import send_mail
    send_mail(to=to, subject=subject, body=message)


def _print_log(title: str, logs: list[str]):
    print(f"\n{'─' * 50}")
    print(f"  {title}")
    print('─' * 50)
    for line in logs:
        print(f"  {line}")
    print('─' * 50 + "\n")


# ============================================================
# CLI エントリーポイント
# ============================================================
def _cli():
    parser = argparse.ArgumentParser(
        description="注文書処理フロー — 単価チェック → ERP登録 or 差し替え依頼",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  # 試作注文書を処理（ドライランモード）
  python -m agents.run_flow orders/order_001.pdf --order-type 試作 --verbose

  # 実際に差し替えメールを送信する場合
  MAIL_DRY_RUN=false python -m agents.run_flow orders/order_001.pdf \\
    --customer-mail customer@example.com \\
    --notify-mail internal@example.com
        """,
    )
    parser.add_argument("pdf_path", help="注文書 PDF のパス")
    parser.add_argument(
        "--order-type", default="", choices=["試作", "量産", ""],
        help="注文種別（デフォルト: 自動判定）",
    )
    parser.add_argument(
        "--customer-mail", default="",
        help="差し替え依頼の送信先（お客様資材担当者）",
    )
    parser.add_argument(
        "--notify-mail", default="",
        help="社内通知先メールアドレス",
    )
    parser.add_argument(
        "--verbose", action="store_true",
        help="Agent の思考ログを表示する",
    )
    args = parser.parse_args()

    # PDF ファイルの存在確認
    if not Path(args.pdf_path).exists():
        print(f"エラー: PDF ファイルが見つかりません — {args.pdf_path}", file=sys.stderr)
        sys.exit(1)

    result = run(
        pdf_path=args.pdf_path,
        order_type=args.order_type,
        customer_mail=args.customer_mail,
        internal_notify_mail=args.notify_mail,
        verbose=args.verbose,
    )

    print("\n" + "=" * 60)
    print("  処理完了")
    print("=" * 60)

    import json
    print(json.dumps(result, ensure_ascii=False, indent=2))

    sys.exit(0 if result.get("status") == "ok" else 1)


if __name__ == "__main__":
    _cli()
