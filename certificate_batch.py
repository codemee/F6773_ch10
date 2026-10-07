"""Run certificate creation and PDF export as one isolated batch."""

from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

from certificate_maker import Attendee, fill_template, read_attendees, safe_filename
from certificate_to_pdf import convert_one


def new_batch_folder(root: Path) -> Path:
    root.mkdir(parents=True, exist_ok=True)
    stem = f"批次_{datetime.now():%Y%m%d_%H%M%S}"
    folder = root / stem
    suffix = 2
    while folder.exists():
        folder = root / f"{stem}_{suffix}"
        suffix += 1
    folder.mkdir()
    return folder


def result_line(row: int | str, student_id: str, name: str, word: str, pdf: str, detail: str = "") -> str:
    identity = " / ".join(value for value in (student_id, name) if value) or "資料檢核"
    return f"- 第 {row} 列｜{identity}｜Word：{word}｜PDF：{pdf}" + (f"｜{detail}" if detail else "")


def process_attendee(attendee: Attendee, template: Path, word_dir: Path, pdf_dir: Path, timeout: int) -> str:
    filename = f"{safe_filename(attendee.student_id)}_{safe_filename(attendee.name)}_研習證書.docx"
    word_file = word_dir / filename
    try:
        fill_template(template, word_file, attendee)
    except (OSError, ValueError) as error:
        return result_line(attendee.row_number, attendee.student_id, attendee.name, "失敗", "未轉檔", str(error))

    pdf_result = convert_one(word_file, pdf_dir, timeout)
    if pdf_result["status"] == "success":
        return result_line(attendee.row_number, attendee.student_id, attendee.name, f"成功（Word/{filename}）", f"成功（PDF/{pdf_result['output']}）")
    return result_line(attendee.row_number, attendee.student_id, attendee.name, f"成功（Word/{filename}）", "失敗，已保留 Word", pdf_result["error"])


def main() -> int:
    parser = argparse.ArgumentParser(description="從 Excel 名單與 Word 範本建立獨立批次的 Word/PDF 證書。")
    parser.add_argument("--roster", type=Path, default=Path("參加者名單.xlsx"), help="Excel 參加者名單")
    parser.add_argument("--template", type=Path, default=Path("研習證書範本.docx"), help="Word 證書範本")
    parser.add_argument("--output-root", type=Path, default=Path("證書批次輸出"), help="所有批次的上層資料夾")
    parser.add_argument("--timeout", type=int, default=120, help="每份 PDF 使用 Microsoft Word 的最長轉檔秒數")
    args = parser.parse_args()
    if args.timeout <= 0:
        parser.error("--timeout 必須大於 0。")

    if not args.roster.is_file() or not args.template.is_file():
        missing = [str(path) for path in (args.roster, args.template) if not path.is_file()]
        parser.error("找不到檔案：" + "、".join(missing))

    batch = new_batch_folder(args.output_root)
    word_dir = batch / "Word"
    pdf_dir = batch / "PDF"
    word_dir.mkdir()
    pdf_dir.mkdir()
    attendees, issues = read_attendees(args.roster)
    results = ["完整證書流程處理結果", "=" * 24, f"批次資料夾：{batch.name}", f"Excel：{args.roster.name}", f"範本：{args.template.name}", ""]
    results.extend(result_line("-", "", "", "未製作", "未轉檔", issue) for issue in issues)
    results.extend(process_attendee(attendee, args.template, word_dir, pdf_dir, args.timeout) for attendee in attendees)
    report = batch / "處理結果.txt"
    report.write_text("\n".join(results) + "\n", encoding="utf-8")

    succeeded = sum("｜Word：成功" in line and "｜PDF：成功" in line for line in results)
    failed = len(issues) + len(attendees) - succeeded
    print(f"批次完成：{batch}")
    print(f"Word 與 PDF 都成功：{succeeded} 份；需要處理：{failed} 項。")
    print(f"處理報告：{report}")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
