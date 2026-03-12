from __future__ import annotations

import argparse
import ctypes
import json
import os
import sqlite3
import subprocess
import sys
import threading
import time
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[4]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from docx_pipeline.config.settings import get_settings


def ts() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")


def log_line(log_path: Path, message: str) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("a", encoding="utf-8") as f:
        f.write(f"{ts()} | {message.rstrip()}\n")


def get_connection() -> sqlite3.Connection:
    settings = get_settings()
    conn = sqlite3.connect(settings.db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def fetch_documents_by_ids(document_ids: list[int]) -> list[dict[str, Any]]:
    if not document_ids:
        return []
    placeholders = ",".join("?" for _ in document_ids)
    order_map = {doc_id: idx for idx, doc_id in enumerate(document_ids)}
    conn = get_connection()
    try:
        rows = conn.execute(
            f"""
            SELECT id, uid, source_abs_path, artifacts_abs_path
            FROM document
            WHERE id IN ({placeholders})
            """,
            tuple(document_ids),
        ).fetchall()
        docs = [dict(r) for r in rows]
        docs.sort(key=lambda d: order_map[int(d["id"])])
        return docs
    finally:
        conn.close()


def get_word_pid(app) -> int | None:
    try:
        hwnd = int(app.Hwnd)
    except Exception:
        return None
    pid = ctypes.c_ulong()
    ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return int(pid.value) if pid.value else None


def start_word(log_path: Path):
    import win32com.client  # type: ignore

    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    pid = get_word_pid(app)
    log_line(log_path, f"[saveas_batch] Word.Application started pid={pid}")
    return app, pid


def kill_word_process(pid: int | None, log_path: Path) -> None:
    if not pid:
        log_line(log_path, "[saveas_batch] kill requested but pid is unknown")
        return
    try:
        completed = subprocess.run(
            ["taskkill", "/PID", str(pid), "/F", "/T"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        log_line(log_path, "[stdout]")
        if completed.stdout:
            for line in completed.stdout.splitlines():
                log_line(log_path, line)
        log_line(log_path, "[stderr]")
        if completed.stderr:
            for line in completed.stderr.splitlines():
                log_line(log_path, line)
        log_line(log_path, f"[saveas_batch] taskkill exit_code={completed.returncode}")
    except Exception as exc:
        log_line(log_path, f"[saveas_batch] taskkill failed: {type(exc).__name__}: {exc}")


def close_doc(doc, log_path: Path) -> None:
    if doc is None:
        return
    try:
        doc.Close(False)
    except Exception as exc:
        log_line(log_path, f"[saveas_batch] doc close warning: {type(exc).__name__}: {exc}")


def quit_word(app, log_path: Path) -> None:
    if app is None:
        return
    try:
        app.Quit()
        log_line(log_path, "[saveas_batch] Word.Application quit")
    except Exception as exc:
        log_line(log_path, f"[saveas_batch] word quit warning: {type(exc).__name__}: {exc}")


def _saveas_operation(app, in_docx: str, out_docx: str, result_box: dict[str, Any]) -> None:
    doc = None
    try:
        doc = app.Documents.Open(in_docx, ReadOnly=True)
        doc.SaveAs2(out_docx, FileFormat=16)
        result_box["ok"] = True
    except Exception as exc:
        result_box["exc"] = exc
    finally:
        close_doc(doc, Path(result_box["log_path"]))


def saveas_once(app, pid: int | None, uid: str, in_docx: str, out_docx: str, log_path: Path, timeout_sec: int) -> dict[str, Any]:
    started = time.time()
    log_line(log_path, f"========== {uid} ==========")
    log_line(log_path, f"[saveas_batch] input={in_docx}")
    log_line(log_path, f"[saveas_batch] output={out_docx}")

    result_box: dict[str, Any] = {"log_path": str(log_path)}
    thread = threading.Thread(
        target=_saveas_operation,
        args=(app, in_docx, out_docx, result_box),
        daemon=True,
    )
    thread.start()
    thread.join(timeout=timeout_sec)

    duration_ms = int((time.time() - started) * 1000)

    if thread.is_alive():
        log_line(log_path, "[stdout]")
        log_line(log_path, f"[saveas_batch] timeout after {timeout_sec}s")
        log_line(log_path, "[exit_code]")
        log_line(log_path, "1")
        log_line(log_path, "[duration_ms]")
        log_line(log_path, str(duration_ms))
        kill_word_process(pid, log_path)
        return {"status": "retry_needed", "reason": f"timeout>{timeout_sec}s", "duration_ms": duration_ms}

    exc = result_box.get("exc")
    if exc is not None:
        log_line(log_path, "[stdout]")
        log_line(log_path, f"[saveas_batch] exception: {type(exc).__name__}: {exc}")
        log_line(log_path, "[exit_code]")
        log_line(log_path, "1")
        log_line(log_path, "[duration_ms]")
        log_line(log_path, str(duration_ms))
        kill_word_process(pid, log_path)
        return {"status": "retry_needed", "reason": f"{type(exc).__name__}: {exc}", "duration_ms": duration_ms}

    if not os.path.exists(out_docx):
        log_line(log_path, "[stdout]")
        log_line(log_path, "[saveas_batch] output file was not created")
        log_line(log_path, "[exit_code]")
        log_line(log_path, "1")
        log_line(log_path, "[duration_ms]")
        log_line(log_path, str(duration_ms))
        kill_word_process(pid, log_path)
        return {"status": "retry_needed", "reason": "missing_output", "duration_ms": duration_ms}

    log_line(log_path, "[stdout]")
    log_line(log_path, "[saveas_batch] ok")
    log_line(log_path, "[exit_code]")
    log_line(log_path, "0")
    log_line(log_path, "[duration_ms]")
    log_line(log_path, str(duration_ms))
    return {"status": "ok", "reason": None, "duration_ms": duration_ms}


def process_documents(
    documents: list[dict[str, Any]],
    log_path: Path,
    timeout_sec: int,
    retries: int,
    retry_sleep_sec: int,
    restart_every: int,
) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    app = None
    pid = None
    docs_since_restart = 0

    try:
        for doc in documents:
            uid = str(doc["uid"])
            in_docx = str(Path(doc["source_abs_path"]))
            out_docx = str(Path(doc["artifacts_abs_path"]) / "materialized.docx")

            attempt_no = 0
            while True:
                if app is None:
                    app, pid = start_word(log_path)
                    docs_since_restart = 0

                attempt_no += 1
                log_line(log_path, f"[saveas_batch] attempt={attempt_no}/{retries}")
                result = saveas_once(
                    app=app,
                    pid=pid,
                    uid=uid,
                    in_docx=in_docx,
                    out_docx=out_docx,
                    log_path=log_path,
                    timeout_sec=timeout_sec,
                )

                if result["status"] == "ok":
                    docs_since_restart += 1
                    results.append(
                        {
                            "document_id": int(doc["id"]),
                            "uid": uid,
                            "status": "ok",
                            "duration_ms": result["duration_ms"],
                        }
                    )
                    break

                app = None
                pid = None

                if attempt_no >= retries:
                    results.append(
                        {
                            "document_id": int(doc["id"]),
                            "uid": uid,
                            "status": "failed_saveas",
                            "duration_ms": result["duration_ms"],
                            "error_message": result["reason"],
                        }
                    )
                    break

                time.sleep(retry_sleep_sec)

            if app is not None and docs_since_restart >= restart_every:
                quit_word(app, log_path)
                app = None
                pid = None
                docs_since_restart = 0

    finally:
        if app is not None:
            quit_word(app, log_path)

    return results


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--ids", required=True, help="Comma-separated document ids")
    parser.add_argument("--log", required=True, help="Path to global run log")
    parser.add_argument("--results", required=True, help="Path to results JSON")
    parser.add_argument("--timeout-sec", type=int, default=5)
    parser.add_argument("--retries", type=int, default=3)
    parser.add_argument("--retry-sleep-sec", type=int, default=1)
    parser.add_argument("--restart-every", type=int, default=100)
    args = parser.parse_args()

    document_ids = [int(x) for x in args.ids.split(",") if x.strip()]
    log_path = Path(args.log)
    results_path = Path(args.results)

    docs = fetch_documents_by_ids(document_ids)
    log_line(log_path, f"[saveas_batch] starting batch size={len(docs)} ids={document_ids}")

    results = process_documents(
        documents=docs,
        log_path=log_path,
        timeout_sec=args.timeout_sec,
        retries=args.retries,
        retry_sleep_sec=args.retry_sleep_sec,
        restart_every=args.restart_every,
    )

    results_path.parent.mkdir(parents=True, exist_ok=True)
    with results_path.open("w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    ok_count = sum(1 for r in results if r["status"] == "ok")
    fail_count = len(results) - ok_count
    log_line(log_path, f"[saveas_batch] finished ok={ok_count} failed={fail_count}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
