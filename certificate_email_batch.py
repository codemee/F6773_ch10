"""Create certificates, convert them to PDF, and explicitly send them by email.

Every batch has a manifest.  A delivered email is recorded before the next
entry is processed, so resuming the same batch never sends it again.
"""

from __future__ import annotations

import argparse
import json
import smtplib
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

from certificate_email_test import send_certificate_email
from certificate_maker import Attendee, fill_template, read_attendees, safe_filename
from certificate_to_pdf import convert_one


MANIFEST_NAME = "寄送狀態.json"
REPORT_NAME = "處理與寄送結果.txt"


def now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def save_manifest(path: Path, data: dict[str, Any]) -> None:
    data["updated_at"] = now()
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temporary.replace(path)


def new_batch_folder(root: Path) -> Path:
    root.mkdir(parents=True, exist_ok=True)
    stem = f"寄送批次_{datetime.now():%Y%m%d_%H%M%S}"
    folder = root / stem
    suffix = 2
    while folder.exists():
        folder = root / f"{stem}_{suffix}"
        suffix += 1
    folder.mkdir()
    (folder / "Word").mkdir()
    (folder / "PDF").mkdir()
    return folder


def attendee_entry(attendee: Attendee) -> dict[str, Any]:
    base_name = f"{safe_filename(attendee.student_id)}_{safe_filename(attendee.name)}_研習證書"
    return {
        "key": f"{attendee.row_number}:{attendee.student_id}",
        "row_number": attendee.row_number,
        "student_id": attendee.student_id,
        "name": attendee.name,
        "email": attendee.email,
        "course": attendee.course,
        "completion_date": attendee.completion_date,
        "word_file": f"Word/{base_name}.docx",
        "pdf_file": f"PDF/{base_name}.pdf",
        "word_status": "pending",
        "pdf_status": "pending",
        "email_status": "not_sent",
        "failure_reason": "",
        "email_attempts": 0,
    }


def create_manifest(batch: Path, roster: Path, template: Path) -> dict[str, Any]:
    attendees, issues = read_attendees(roster)
    return {
        "version": 1,
        "created_at": now(),
        "updated_at": now(),
        "roster": str(roster.resolve()),
        "template": str(template.resolve()),
        "validation_issues": issues,
        "entries": [attendee_entry(attendee) for attendee in attendees],
    }


def load_manifest(batch: Path) -> tuple[Path, dict[str, Any]]:
    path = batch / MANIFEST_NAME
    if not path.is_file():
        raise ValueError(f"找不到批次狀態檔：{path}")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise ValueError("無法讀取批次狀態檔。") from error
    if data.get("version") != 1 or not isinstance(data.get("entries"), list):
        raise ValueError("批次狀態檔格式不支援。")
    return path, data


def failure(error: BaseException) -> str:
    """Keep a useful class-level failure reason without leaking SMTP details."""
    if isinstance(error, smtplib.SMTPAuthenticationError):
        return "SMTPAuthenticationError：Gmail 拒絕登入，請檢查應用程式密碼。"
    if isinstance(error, smtplib.SMTPRecipientsRefused):
        return "SMTPRecipientsRefused：收件者地址被 SMTP 伺服器拒絕。"
    if isinstance(error, smtplib.SMTPException):
        return f"{type(error).__name__}：SMTP 寄送失敗。"
    if isinstance(error, OSError):
        return f"{type(error).__name__}：無法連線至 Gmail SMTP。"
    if isinstance(error, ValueError):
        return "設定錯誤：請檢查 .env 中的 Gmail 設定。"
    return f"{type(error).__name__}：處理失敗。"


def process_entry(entry: dict[str, Any], batch: Path, template: Path, timeout: int, send: bool) -> None:
    if entry["email_status"] == "sent":
        return
    word = batch / entry["word_file"]
    pdf = batch / entry["pdf_file"]
    attendee = Attendee(entry["row_number"], entry["student_id"], entry["name"], entry["email"], entry["course"], entry["completion_date"])

    if entry["word_status"] != "success" or not word.is_file():
        try:
            fill_template(template, word, attendee)
            entry["word_status"] = "success"
            entry["pdf_status"] = "pending"
            entry["failure_reason"] = ""
        except (OSError, ValueError) as error:
            entry["word_status"] = "failed"
            entry["pdf_status"] = "not_created"
            entry["email_status"] = "not_sent"
            entry["failure_reason"] = str(error)
            return

    if entry["pdf_status"] != "success" or not pdf.is_file():
        result = convert_one(word, pdf.parent, timeout)
        if result["status"] != "success":
            entry["pdf_status"] = "failed"
            entry["email_status"] = "not_sent"
            entry["failure_reason"] = result["error"]
            return
        entry["pdf_status"] = "success"
        entry["failure_reason"] = ""

    if not send:
        entry["email_status"] = "dry_run_ready"
        return
    try:
        entry["email_attempts"] += 1
        send_certificate_email(entry["email"], entry["name"], entry["course"], pdf)
        entry["email_status"] = "sent"
        entry["failure_reason"] = ""
    except (OSError, ValueError, smtplib.SMTPException) as error:
        entry["email_status"] = "failed"
        entry["failure_reason"] = failure(error)


def write_report(path: Path, data: dict[str, Any], dry_run: bool) -> None:
    entries = data["entries"]
    sent = sum(entry["email_status"] == "sent" for entry in entries)
    ready = sum(entry["email_status"] == "dry_run_ready" for entry in entries)
    failed = [entry for entry in entries if entry["word_status"] == "failed" or entry["pdf_status"] == "failed" or entry["email_status"] == "failed"]
    lines = ["證書製作、PDF 轉檔與寄送結果", "=" * 30, f"模式：{'Dry run（未實際寄信）' if dry_run else '實際寄送'}", f"成功寄送：{sent} 封", f"待實際寄送：{ready} 封", f"失敗：{len(failed)} 項", ""]
    if data["validation_issues"]:
        lines.extend(["名單資料問題（未建立或寄送）："] + [f"- {issue}" for issue in data["validation_issues"]] + [""])
    lines.append("每位參加者：")
    for entry in entries:
        status = f"Word={entry['word_status']}；PDF={entry['pdf_status']}；寄送={entry['email_status']}"
        reason = f"；原因：{entry['failure_reason']}" if entry["failure_reason"] else ""
        lines.append(
            f"- 第 {entry['row_number']} 列｜{entry['student_id']}／{entry['name']}｜收件者：{entry['email']}"
            f"｜附件：{entry['pdf_file']}｜{status}{reason}"
        )
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="建立 Word/PDF 證書並選擇性以 Gmail 寄送 PDF。預設 dry run，不實際寄信。")
    parser.add_argument("--roster", type=Path, default=Path("參加者名單.xlsx"), help="新批次使用的 Excel 名單")
    parser.add_argument("--template", type=Path, default=Path("研習證書範本.docx"), help="新批次使用的 Word 範本")
    parser.add_argument("--output-root", type=Path, default=Path("證書寄送批次"), help="新批次的上層資料夾")
    parser.add_argument("--batch-dir", type=Path, help="重跑既有批次；會略過已成功寄送的郵件")
    parser.add_argument("--send", action="store_true", help="實際透過 .env 的 Gmail SMTP 寄送；未指定時只做 dry run")
    parser.add_argument("--timeout", type=int, default=120, help="每份 PDF 使用 Microsoft Word 的最長轉檔秒數")
    args = parser.parse_args()
    if args.timeout <= 0:
        parser.error("--timeout 必須大於 0。")

    if args.batch_dir:
        batch = args.batch_dir
        try:
            manifest_path, data = load_manifest(batch)
        except ValueError as error:
            parser.error(str(error))
        template = Path(data["template"])
    else:
        if not args.roster.is_file() or not args.template.is_file():
            missing = [str(path) for path in (args.roster, args.template) if not path.is_file()]
            parser.error("找不到檔案：" + "、".join(missing))
        batch = new_batch_folder(args.output_root)
        manifest_path = batch / MANIFEST_NAME
        data = create_manifest(batch, args.roster, args.template)
        save_manifest(manifest_path, data)
        template = args.template

    if not template.is_file():
        parser.error(f"找不到範本：{template}")

    for entry in data["entries"]:
        process_entry(entry, batch, template, args.timeout, args.send)
        save_manifest(manifest_path, data)
    report = batch / REPORT_NAME
    write_report(report, data, dry_run=not args.send)
    save_manifest(manifest_path, data)

    sent = sum(entry["email_status"] == "sent" for entry in data["entries"])
    pending = sum(entry["email_status"] in {"not_sent", "dry_run_ready", "failed"} for entry in data["entries"])
    failed = sum(entry["word_status"] == "failed" or entry["pdf_status"] == "failed" or entry["email_status"] == "failed" for entry in data["entries"])
    print(f"批次資料夾：{batch}")
    print(f"成功寄送：{sent} 封；尚待處理或重試：{pending} 項。")
    print(f"處理報告：{report}")
    # A successful dry run intentionally leaves emails pending, which is not an error.
    if not args.send:
        return 0 if not failed and not data["validation_issues"] else 1
    return 0 if not pending and not data["validation_issues"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
