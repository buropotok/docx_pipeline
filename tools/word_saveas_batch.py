from __future__ import annotations

import argparse
import json
import os
import sys
import time
import traceback
from pathlib import Path


def log_event(log_path: str | None, message: str) -> None:
    if not log_path:
        return
    path = Path(log_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as f:
        f.write(message.rstrip() + "\n")


def start_word():
    import win32com.client  # type: ignore
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    return app


def close_doc(doc, log_path: str | None) -> None:
    if doc is None:
        return
    try:
        log_event(log_path, "[saveas-batch] closing document")
        doc.Close(False)
    except Exception as exc:
        log_event(log_path, f"[saveas-batch] close doc warning: {exc}")


def quit_word(app, log_path: str | None) -> None:
    if app is None:
        return
    try:
        log_event(log_path, "[saveas-batch] quitting Word")
        app.Quit()
    except Exception as exc:
        log_event(log_path, f"[saveas-batch] quit warning: {exc}")


def process_job(app, job: dict) -> dict:
    uid = str(job["uid"])
    in_docx = os.path.abspath(str(job["in_docx"]))
    out_docx = os.path.abspath(str(job["out_docx"]))
    log_path = str(job.get("log_path") or "")

    os.makedirs(os.path.dirname(out_docx) or ".", exist_ok=True)

    started = time.time()
    doc = None
    try:
        log_event(log_path, f"[saveas-batch] uid={uid}")
        log_event(log_path, f"[saveas-batch] input: {in_docx}")
        log_event(log_path, f"[saveas-batch] output: {out_docx}")
        try:
            log_event(log_path, f"[saveas-batch] word_version={app.Version}")
        except Exception:
            log_event(log_path, "[saveas-batch] word_version=unavailable")

        log_event(log_path, f"[saveas-batch] opening document path={in_docx} ReadOnly=True")
        doc = app.Documents.Open(in_docx, ReadOnly=True)

        log_event(log_path, f"[saveas-batch] saveas path={out_docx} file_format=16")
        doc.SaveAs2(out_docx, FileFormat=16)

        if not os.path.exists(out_docx):
            raise FileNotFoundError(f"SaveAs2 finished but output file not found: {out_docx}")

        duration_ms = int((time.time() - started) * 1000)
        log_event(log_path, f"[saveas-batch] done duration_ms={duration_ms}")
        return {
            "uid": uid,
            "status": "ok",
            "in_docx": in_docx,
            "out_docx": out_docx,
            "duration_ms": duration_ms,
        }
    except Exception as exc:
        duration_ms = int((time.time() - started) * 1000)
        hresult = getattr(exc, "hresult", None)
        tb = traceback.format_exc()
        log_event(log_path, f"[saveas-batch] ERROR type={type(exc).__name__} hresult={hresult} message={exc}")
        log_event(log_path, tb)
        return {
            "uid": uid,
            "status": "failed",
            "in_docx": in_docx,
            "out_docx": out_docx,
            "duration_ms": duration_ms,
            "error_type": type(exc).__name__,
            "error_message": str(exc),
            "traceback": tb,
        }
    finally:
        close_doc(doc, log_path)


def main() -> int:
    parser = argparse.ArgumentParser(description="Batch materialize DOC/DOCX via one Word.Application")
    parser.add_argument("--jobs", required=True, help="Path to JSON file with jobs")
    parser.add_argument("--results", required=True, help="Path to JSON file for results")
    parser.add_argument("--restart-every", type=int, default=200, help="Restart Word every N documents")
    args = parser.parse_args()

    jobs_path = Path(args.jobs)
    results_path = Path(args.results)
    restart_every = max(1, int(args.restart_every))

    with jobs_path.open("r", encoding="utf-8") as f:
        jobs = json.load(f)

    if not isinstance(jobs, list):
        print("[saveas-batch] jobs JSON must be a list", file=sys.stderr)
        return 2

    try:
        import win32com.client  # noqa: F401
    except Exception as exc:
        print(f"[saveas-batch] pywin32 import failed: {exc}", file=sys.stderr)
        return 2

    results: list[dict] = []
    app = None
    processed_since_restart = 0

    try:
        app = start_word()

        for idx, job in enumerate(jobs, start=1):
            if processed_since_restart >= restart_every:
                quit_word(app, None)
                app = start_word()
                processed_since_restart = 0

            result = process_job(app, job)
            results.append(result)
            processed_since_restart += 1

            # Conservative recovery: restart Word after any failed file.
            if result["status"] != "ok":
                quit_word(app, str(job.get("log_path") or ""))
                app = start_word()
                processed_since_restart = 0

            if idx % 50 == 0:
                print(f"[saveas-batch] processed {idx}/{len(jobs)}")

    finally:
        quit_word(app, None)

    results_path.parent.mkdir(parents=True, exist_ok=True)
    with results_path.open("w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    ok_count = sum(1 for r in results if r["status"] == "ok")
    fail_count = len(results) - ok_count
    print(f"[saveas-batch] done ok={ok_count} failed={fail_count}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
