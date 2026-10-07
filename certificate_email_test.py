"""Send an explicitly requested Gmail SMTP test message.

This command is intentionally separate from certificate creation and never
attaches or sends certificates.
"""

from __future__ import annotations

import argparse
import os
import smtplib
import sys
from email.message import EmailMessage
from pathlib import Path


ENV_FILE = Path(".env")
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 465


def load_dotenv(path: Path) -> None:
    """Load only missing environment variables from a small local .env file."""
    if not path.is_file():
        return

    for line_number, raw_line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            raise ValueError(f".env 第 {line_number} 行格式錯誤，請使用 KEY=VALUE。")
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if not key:
            raise ValueError(f".env 第 {line_number} 行缺少設定名稱。")
        os.environ.setdefault(key, value)


def settings() -> tuple[str, str]:
    load_dotenv(ENV_FILE)
    username = os.environ.get("GMAIL_SMTP_USER", "").strip()
    app_password = os.environ.get("GMAIL_SMTP_APP_PASSWORD", "").strip()
    missing = []
    if not username:
        missing.append("GMAIL_SMTP_USER")
    if not app_password:
        missing.append("GMAIL_SMTP_APP_PASSWORD")
    if missing:
        raise ValueError(".env 缺少 " + "、".join(missing) + "；請填寫後再試。")
    return username, app_password


def send_test_email(recipient: str, subject: str) -> None:
    username, app_password = settings()
    message = EmailMessage()
    message["From"] = username
    message["To"] = recipient
    message["Subject"] = subject
    message.set_content(
        "這是一封由研習證書製作工具手動寄出的 Gmail SMTP 測試信。\n\n"
        "此測試不會附加或寄送任何證書。"
    )
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30) as server:
        server.login(username, app_password)
        server.send_message(message)


def send_certificate_email(recipient: str, name: str, course: str, attachment: Path) -> None:
    """Send one PDF certificate using the local Gmail settings.

    This function deliberately returns no credentials and leaves SMTP errors for
    the caller to classify without printing protocol details.
    """
    username, app_password = settings()
    message = EmailMessage()
    message["From"] = username
    message["To"] = recipient
    message["Subject"] = f"{course}｜研習證書"
    message.set_content(
        f"{name} 您好：\n\n"
        f"感謝您完成「{course}」。您的研習證書 PDF 已附在本信中，請妥善保存。\n\n"
        "此信由研習證書製作工具寄送。"
    )
    message.add_attachment(attachment.read_bytes(), maintype="application", subtype="pdf", filename=attachment.name)
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30) as server:
        server.login(username, app_password)
        server.send_message(message)


def main() -> int:
    parser = argparse.ArgumentParser(description="以 .env 中的 Gmail SMTP 設定手動寄送一封測試信。")
    parser.add_argument("--to", required=True, help="測試信收件者 email")
    parser.add_argument("--subject", default="研習證書工具 Gmail SMTP 測試", help="測試信主旨")
    args = parser.parse_args()

    try:
        send_test_email(args.to, args.subject)
    except (OSError, ValueError, smtplib.SMTPException) as error:
        # Never include configuration values or SMTP protocol details in output.
        print(f"測試信未寄出：{type(error).__name__}。請確認 .env 設定、網路與 Gmail 應用程式密碼。", file=sys.stderr)
        return 1

    print(f"測試信已寄出至：{args.to}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
