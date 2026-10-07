"""Convert an existing folder of Word certificates to PDF via Microsoft Word."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import tempfile
from pathlib import Path


def write_report(path: Path, results: list[dict[str, str]]) -> None:
    successful = [item for item in results if item["status"] == "success"]
    failed = [item for item in results if item["status"] != "success"]
    lines = ["Word 證書轉 PDF 結果", "=" * 20, f"成功轉換：{len(successful)} 份"]
    lines.extend(f"- {item['source']} -> {item['output']}" for item in successful)
    lines.extend(["", f"失敗：{len(failed)} 項"])
    lines.extend(f"- {item['source']}：{item['error']}" for item in failed) if failed else lines.append("- 無")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def word_worker(sources: list[Path], output_folder: Path, state_path: Path) -> None:
    import ctypes
    from ctypes import wintypes

    import win32api
    import win32com.client
    import win32process

    state = {"results": [], "stage": "啟動 Microsoft Word", "word_pid": None}
    log = output_folder / "轉檔紀錄.txt"

    def save(stage: str) -> None:
        state["stage"] = stage
        temporary = state_path.with_suffix(".tmp")
        temporary.write_text(json.dumps(state, ensure_ascii=False), encoding="utf-8")
        temporary.replace(state_path)
        with log.open("a", encoding="utf-8") as stream:
            stream.write(stage + "\n")

    word = None
    save("啟動 Microsoft Word")
    try:
        previous_pids = set(win32process.EnumProcesses())
        word = win32com.client.DispatchEx("Word.Application")
        # Word.Application has no Hwnd property. Identify only a newly
        # created WINWORD process; never terminate an existing user session.
        query_image = ctypes.WinDLL("kernel32", use_last_error=True).QueryFullProcessImageNameW
        query_image.argtypes = [wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
        query_image.restype = wintypes.BOOL
        candidates = []
        for pid in set(win32process.EnumProcesses()) - previous_pids:
            try:
                handle = win32api.OpenProcess(0x0400, False, pid)
                try:
                    image_path = ctypes.create_unicode_buffer(32768)
                    size = wintypes.DWORD(len(image_path))
                    if query_image(int(handle), 0, image_path, ctypes.byref(size)) and image_path.value.lower().endswith("\\winword.exe"):
                        candidates.append((pid, win32process.GetProcessTimes(handle)["CreationTime"].isoformat()))
                finally:
                    handle.Close()
            except OSError:
                continue
        if len(candidates) == 1:
            state["word_pid"], state["word_created"] = candidates[0]
        word.Visible = False
        word.DisplayAlerts = 0
        save("Word 已啟動")
        for source in sources:
            document = None
            pdf_path = output_folder / f"{source.stem}.pdf"
            result = {"source": source.name, "output": pdf_path.name, "status": "failed", "error": ""}
            try:
                save(f"開啟：{source.name}")
                document = word.Documents.Open(
                    str(source), ConfirmConversions=False, ReadOnly=True,
                    AddToRecentFiles=False, Visible=False,
                )
                save(f"匯出：{source.name}")
                document.ExportAsFixedFormat(str(pdf_path), 17)
                if not pdf_path.is_file() or pdf_path.stat().st_size == 0:
                    raise RuntimeError("Word 未產生有效的 PDF 檔案。")
                result["status"] = "success"
            except Exception as error:
                result["error"] = str(error)
            finally:
                state["results"].append(result)
                save(f"{result['status']}：{source.name}")
                if document is not None:
                    save(f"關閉文件：{source.name}")
                    document.Close(False)
    except Exception as error:
        done = {item["source"] for item in state["results"]}
        for source in sources:
            if source.name not in done:
                state["results"].append({"source": source.name, "output": f"{source.stem}.pdf", "status": "failed", "error": str(error)})
        save(f"Word 錯誤：{error}")
    finally:
        if word is not None:
            save("關閉 Microsoft Word")
            word.Quit()
        save("轉檔程序結束")


def stop_worker_word(state: dict) -> None:
    """Stop only the isolated Word process created by this conversion."""
    if not state.get("word_pid"):
        return
    import win32api
    import win32process

    try:
        handle = win32api.OpenProcess(0x0400 | 0x0001, False, state["word_pid"])
        try:
            created = win32process.GetProcessTimes(handle)["CreationTime"].isoformat()
            if created == state.get("word_created"):
                win32api.TerminateProcess(handle, 1)
        finally:
            handle.Close()
    except OSError:
        pass


def convert_files(sources: list[Path], output_folder: Path, timeout: int = 120) -> tuple[list[dict[str, str]], str]:
    """Export selected documents in an isolated Word worker with a time limit."""
    if timeout <= 0:
        raise ValueError("timeout 必須大於 0。")
    sources = [source.resolve() for source in sources if not source.name.startswith("~$")]
    if not sources:
        return [], ""
    output_folder = output_folder.resolve()
    output_folder.mkdir(parents=True, exist_ok=True)
    worker_error = ""
    state = {}
    if sys.platform != "win32":
        worker_error = "Microsoft Word 轉檔僅支援 Windows。"
    else:
        with tempfile.TemporaryDirectory(prefix=".word-pdf-", dir=output_folder) as temporary:
            state_path = Path(temporary) / "state.json"
            request_path = Path(temporary) / "sources.json"
            request_path.write_text(json.dumps([str(source) for source in sources]), encoding="utf-8")
            command = [sys.executable, str(Path(__file__).resolve()), "--_worker",
                       str(request_path), str(output_folder), str(state_path)]
            timed_out = False
            try:
                completed = subprocess.run(command, capture_output=True, text=True,
                                           encoding="utf-8", errors="replace", timeout=timeout)
                if completed.returncode:
                    worker_error = completed.stderr.strip() or "Word 轉檔程序異常結束。"
            except subprocess.TimeoutExpired:
                timed_out = True
            except OSError as error:
                worker_error = str(error)
            if state_path.exists():
                state = json.loads(state_path.read_text(encoding="utf-8"))
            if timed_out:
                stage = state.get("stage", "啟動 Microsoft Word")
                worker_error = f"Microsoft Word 在 {timeout} 秒內未完成。停在：{stage}。"
            if worker_error:
                stop_worker_word(state)
    results = state.get("results", [])
    done = {item["source"] for item in results}
    results.extend({"source": source.name, "output": f"{source.stem}.pdf",
                    "status": "failed", "error": worker_error or "Word 未回傳文件轉檔結果。"}
                   for source in sources if source.name not in done)
    return results, worker_error


def convert_one(source: Path, output_folder: Path, timeout: int = 120) -> dict[str, str]:
    """Convert only this document, leaving other batch documents untouched."""
    results, _ = convert_files([source], output_folder, timeout)
    if not results:
        return {"source": source.name, "output": f"{source.stem}.pdf", "status": "failed",
                "error": "Word 暫存檔不進行轉檔。"}
    return results[0]


def main() -> int:
    parser = argparse.ArgumentParser(description="使用 Microsoft Word 將既有 DOCX 證書轉成 PDF。")
    parser.add_argument("--input", type=Path, default=Path("輸出證書"), help="含有 DOCX 證書的資料夾")
    parser.add_argument("--output", type=Path, default=Path("輸出 PDF 證書"), help="PDF 輸出資料夾")
    parser.add_argument("--timeout", type=int, default=120, help="等待 Word 完成整批轉檔的秒數（預設 120）")
    args = parser.parse_args()
    if args.timeout <= 0:
        parser.error("--timeout 必須大於 0。")
    if not args.input.is_dir():
        print(f"找不到輸入資料夾：{args.input}", file=sys.stderr)
        return 2
    if sys.platform != "win32":
        print("Microsoft Word 轉檔僅支援 Windows。", file=sys.stderr)
        return 2
    args.output.mkdir(parents=True, exist_ok=True)
    results, worker_error = convert_files(sorted(args.input.glob("*.docx")), args.output, args.timeout)
    if worker_error:
        print(worker_error, file=sys.stderr)
    if not results:
        results = [{"source": "(沒有 DOCX 檔案)", "output": "", "status": "failed", "error": "輸入資料夾內沒有 DOCX 證書。"}]
    report = args.output / "轉檔報告.txt"
    write_report(report, results)
    successful = sum(item["status"] == "success" for item in results)
    failed = len(results) - successful
    print(f"成功轉換 {successful} 份，失敗 {failed} 項。報告：{report}")
    return 0 if failed == 0 and not worker_error else 1


if __name__ == "__main__":
    if len(sys.argv) == 5 and sys.argv[1] == "--_worker":
        word_worker([Path(source) for source in json.loads(Path(sys.argv[2]).read_text(encoding="utf-8"))],
                    Path(sys.argv[3]), Path(sys.argv[4]))
    else:
        raise SystemExit(main())
