"""
agents/config.py
各種システムURL・APIキー・共通設定

環境変数から読み込み、未設定の場合は .env ファイルを参照します。
本番運用では必ず環境変数で設定してください。
"""

import os
from pathlib import Path

# ── .env の自動読み込み（python-dotenv がある場合のみ）──
try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent.parent / ".env")
except ImportError:
    pass  # dotenv 未インストール時はスキップ

# ============================================================
# Anthropic API
# ============================================================
ANTHROPIC_API_KEY: str = os.environ.get("ANTHROPIC_API_KEY", "")
CLAUDE_MODEL: str = os.environ.get("CLAUDE_MODEL", "claude-opus-4-5")

# ============================================================
# 見積・注文管理システム（Google Apps Script Web App）
# ============================================================
GAS_WEB_APP_URL: str = os.environ.get(
    "GAS_WEB_APP_URL",
    "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec",
)
# GAS へのリクエスト共通ヘッダー
GAS_REQUEST_TIMEOUT: int = int(os.environ.get("GAS_REQUEST_TIMEOUT", "30"))

# ============================================================
# メール設定（差し替え依頼・通知メール送信）
# ============================================================
SMTP_HOST: str = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT: int = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER: str = os.environ.get("SMTP_USER", "")
SMTP_PASSWORD: str = os.environ.get("SMTP_PASSWORD", "")  # Gmail: アプリパスワード
MAIL_FROM: str = os.environ.get("MAIL_FROM", SMTP_USER)

# 社内通知先（営業宛て タスク通知）
SALES_NOTIFY_EMAIL: str = os.environ.get("SALES_NOTIFY_EMAIL", "sales@example.com")
# 資材購買担当者
PROCUREMENT_EMAIL: str = os.environ.get("PROCUREMENT_EMAIL", "procurement@example.com")

# ============================================================
# 販売管理システム（ERP）
# ============================================================
ERP_BASE_URL: str = os.environ.get("ERP_BASE_URL", "https://erp.example.com/api/v1")
ERP_API_KEY: str = os.environ.get("ERP_API_KEY", "")
ERP_TIMEOUT: int = int(os.environ.get("ERP_TIMEOUT", "20"))

# ============================================================
# OCR 設定
# ============================================================
# Gemini API（GAS 側と同じキーを使う場合）
GEMINI_API_KEY: str = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL: str = os.environ.get("GEMINI_MODEL", "gemini-1.5-flash-latest")

# 単価照合の許容誤差（例: 0.01 = ±1%）
PRICE_TOLERANCE_RATE: float = float(os.environ.get("PRICE_TOLERANCE_RATE", "0.01"))

# ============================================================
# バリデーション（起動時チェック用）
# ============================================================
REQUIRED_VARS = [
    ("ANTHROPIC_API_KEY", ANTHROPIC_API_KEY),
    ("GAS_WEB_APP_URL",   GAS_WEB_APP_URL),
]

def validate() -> list[str]:
    """未設定の必須環境変数を返す（空リストなら OK）"""
    missing = [name for name, val in REQUIRED_VARS if not val or "YOUR_" in val]
    return missing
