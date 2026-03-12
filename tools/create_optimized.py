from __future__ import annotations

import argparse
import json
import multiprocessing as mp
import sqlite3
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from queue import Empty
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

VALIDATION_TABLE = "reconstruction_validation_run"
INPUT_RAW_NAME = "raw.json"
INPUT_OPTIMIZED_NAME = "raw_optimized.json"
INPUT_DONOR_NAME = "materialized_with_ids.docx"
OUTPUT_DOCX_NAME = "optimized.docx"

RECONSTRUCTION_RUN_ID = f"reconstruct_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

from docx_pipeline.config.settings import get_settings


def ts() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")


def log_line(log_path: Path, message: str) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("a", encoding="utf-8") as f:
        f.write(f"{ts()} | {message.rstrip()}\n")


def log_section(log_path: Path, title: str) -> None:
    log_line(log_path, f"========== {title} ==========")


def get_connection() -> sqlite3.Connection:
    settings = get_settings()
    settings.db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(settings.db_path, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.execute("PRAGMA busy_timeout = 30000;")
    conn.execute("PRAGMA journal_mode = WAL;")
    conn.execute("PRAGMA synchronous = NORMAL;")
    return conn


def fetch_documents(limit: int | None) -> list[dict[str, Any]]:
    conn = get_connection()
    try:
        sql = """
            SELECT
                id,
                uid,
                mail_uid,
                attachment_index,
                source_filename,
                source_abs_path,
                artifacts_abs_path,
                processing_status
            FROM document
            WHERE processing_status IN (
                'discovered',
                'materialized',
                'ids_injected',
                'parsed_raw',
                'failed',
                'failed_saveas',
                'failed_attrs',
                'failed_parser',
                'failed_optimizer',
                'optimized',
                'failed_reconstructor',
                'materializing',
                'reconstructing',
                'injecting_ids',
                'parsing_raw',
                'optimizing'
            )
            ORDER BY id
        """
        if limit is not None:
            sql += f" LIMIT {int(limit)}"
        rows = conn.execute(sql).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def update_status(conn: sqlite3.Connection, document_id: int, status: str) -> None:
    conn.execute(
        """
        UPDATE document
        SET processing_status = ?, update_date = datetime('now')
        WHERE id = ?
        """,
        (status, document_id),
    )
    conn.commit()


def ensure_validation_table(conn: sqlite3.Connection) -> None:
    conn.execute(
        f"""
        CREATE TABLE IF NOT EXISTS {VALIDATION_TABLE} (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id TEXT NOT NULL,
            created_at TEXT NOT NULL DEFAULT (datetime('now')),
            document_id INTEGER NOT NULL,
            uid TEXT NOT NULL,

            source_status TEXT,
            final_status TEXT,

            raw_paragraph_count INTEGER,
            optimized_paragraph_count INTEGER,
            raw_table_count INTEGER,
            optimized_table_count INTEGER,

            raw_text_chars INTEGER,
            optimized_text_chars INTEGER,

            raw_tab_runs INTEGER,
            optimized_tab_runs INTEGER,

            raw_pformat_tabs INTEGER,
            optimized_pformat_tabs INTEGER,

            raw_shape_runs INTEGER,
            optimized_shape_runs INTEGER,

            raw_picture_runs INTEGER,
            optimized_picture_runs INTEGER,

            raw_unknown_runs INTEGER,
            optimized_unknown_runs INTEGER,

            paragraphs_changed INTEGER,
            paragraphs_with_new_tabs INTEGER,
            consecutive_tab_run_issues INTEGER,
            invalid_tab_pos_count INTEGER,
            suspicious_tab_pos_count INTEGER,

            validation_ok INTEGER NOT NULL DEFAULT 1,
            validation_errors_json TEXT,
            summary_json TEXT
        )
        """
    )
    conn.commit()


def load_json_file(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def iter_paragraphs(content: list[dict[str, Any]]):
    for item in content:
        if not isinstance(item, dict):
            continue
        item_type = item.get("type")
        if item_type == "paragraph":
            yield item
        elif item_type == "table":
            for row in item.get("rows", []) or []:
                for cell in row.get("cells", []) or []:
                    for child in cell.get("content", []) or []:
                        if isinstance(child, dict) and child.get("type") == "paragraph":
                            yield child


def iter_tables(content: list[dict[str, Any]]):
    for item in content:
        if isinstance(item, dict) and item.get("type") == "table":
            yield item


def collect_run_stats(content: list[dict[str, Any]]) -> dict[str, int]:
    text_chars = 0
    tab_runs = 0
    shape_runs = 0
    picture_runs = 0
    unknown_runs = 0

    for p in iter_paragraphs(content):
        for run in p.get("runs", []) or []:
            rtype = run.get("type")
            if rtype == "text":
                text_chars += len(run.get("text") or "")
            elif rtype == "tab":
                tab_runs += 1
            elif rtype == "shape":
                shape_runs += 1
            elif rtype == "picture":
                picture_runs += 1
            elif rtype in {"break", "cr", "sym"}:
                pass
            else:
                unknown_runs += 1

    return {
        "text_chars": text_chars,
        "tab_runs": tab_runs,
        "shape_runs": shape_runs,
        "picture_runs": picture_runs,
        "unknown_runs": unknown_runs,
    }


def count_pformat_tabs(content: list[dict[str, Any]]) -> int:
    total = 0
    for p in iter_paragraphs(content):
        total += len((p.get("p_format") or {}).get("tabs") or [])
    return total


def paragraph_signature(paragraph: dict[str, Any]) -> str:
    payload = {
        "runs": paragraph.get("runs", []),
        "p_format": paragraph.get("p_format", {}),
    }
    return json.dumps(payload, ensure_ascii=False, sort_keys=True)


def consecutive_tab_issues(content: list[dict[str, Any]]) -> int:
    issues = 0
    for p in iter_paragraphs(content):
        prev_tab = False
        for run in p.get("runs", []) or []:
            is_tab = run.get("type") == "tab"
            if prev_tab and is_tab:
                issues += 1
            prev_tab = is_tab
    return issues


def usable_page_width(doc: dict[str, Any]) -> int | None:
    page_setup = ((doc.get("document_info") or {}).get("page_setup") or {})
    width = page_setup.get("pageWidthTwip")
    left = page_setup.get("marginLeftTwip", 0)
    right = page_setup.get("marginRightTwip", 0)
    if isinstance(width, int):
        return width - int(left or 0) - int(right or 0)
    return None


def tab_position_issues(content: list[dict[str, Any]], usable_width_twip: int | None) -> tuple[int, int]:
    invalid = 0
    suspicious = 0
    limit = (usable_width_twip + 300) if isinstance(usable_width_twip, int) else None

    for p in iter_paragraphs(content):
        tabs = (p.get("p_format") or {}).get("tabs") or []
        prev = None
        for t in tabs:
            pos = t.get("posTwip")
            if not isinstance(pos, int):
                invalid += 1
                continue
            if pos <= 0:
                invalid += 1
            if limit is not None and pos > limit:
                suspicious += 1
            if prev is not None and pos < prev:
                suspicious += 1
            prev = pos

    return invalid, suspicious


def collect_validation_stats(raw_doc: dict[str, Any], optimized_doc: dict[str, Any]) -> dict[str, Any]:
    raw_content = raw_doc.get("content", []) or []
    opt_content = optimized_doc.get("content", []) or []

    raw_paragraphs = list(iter_paragraphs(raw_content))
    opt_paragraphs = list(iter_paragraphs(opt_content))
    raw_para_by_id = {p.get("id"): p for p in raw_paragraphs if p.get("id")}
    opt_para_by_id = {p.get("id"): p for p in opt_paragraphs if p.get("id")}

    raw_tables = list(iter_tables(raw_content))
    opt_tables = list(iter_tables(opt_content))

    raw_run_stats = collect_run_stats(raw_content)
    opt_run_stats = collect_run_stats(opt_content)

    changed = 0
    paragraphs_with_new_tabs = 0
    for pid, raw_p in raw_para_by_id.items():
        opt_p = opt_para_by_id.get(pid)
        if opt_p is None:
            continue
        if paragraph_signature(raw_p) != paragraph_signature(opt_p):
            changed += 1

        raw_tabs = sum(1 for r in (raw_p.get("runs") or []) if r.get("type") == "tab")
        opt_tabs = sum(1 for r in (opt_p.get("runs") or []) if r.get("type") == "tab")
        if opt_tabs > raw_tabs:
            paragraphs_with_new_tabs += 1

    consecutive_issues = consecutive_tab_issues(opt_content)
    invalid_tab_pos_count, suspicious_tab_pos_count = tab_position_issues(
        opt_content,
        usable_page_width(optimized_doc),
    )

    errors: list[str] = []
    if len(raw_paragraphs) != len(opt_paragraphs):
        errors.append(f"paragraph_count_changed: raw={len(raw_paragraphs)} optimized={len(opt_paragraphs)}")
    if len(raw_tables) != len(opt_tables):
        errors.append(f"table_count_changed: raw={len(raw_tables)} optimized={len(opt_tables)}")
    if opt_run_stats["text_chars"] <= 0:
        errors.append("optimized_text_chars_is_zero")
    if opt_run_stats["unknown_runs"] > raw_run_stats["unknown_runs"]:
        errors.append(
            f"unknown_runs_increased: raw={raw_run_stats['unknown_runs']} optimized={opt_run_stats['unknown_runs']}"
        )
    if consecutive_issues > 0:
        errors.append(f"consecutive_tab_runs={consecutive_issues}")
    if invalid_tab_pos_count > 0:
        errors.append(f"invalid_tab_pos_count={invalid_tab_pos_count}")

    summary = {
        "raw_paragraph_count": len(raw_paragraphs),
        "optimized_paragraph_count": len(opt_paragraphs),
        "raw_table_count": len(raw_tables),
        "optimized_table_count": len(opt_tables),
        "raw_text_chars": raw_run_stats["text_chars"],
        "optimized_text_chars": opt_run_stats["text_chars"],
        "raw_tab_runs": raw_run_stats["tab_runs"],
        "optimized_tab_runs": opt_run_stats["tab_runs"],
        "raw_pformat_tabs": count_pformat_tabs(raw_content),
        "optimized_pformat_tabs": count_pformat_tabs(opt_content),
        "raw_shape_runs": raw_run_stats["shape_runs"],
        "optimized_shape_runs": opt_run_stats["shape_runs"],
        "raw_picture_runs": raw_run_stats["picture_runs"],
        "optimized_picture_runs": opt_run_stats["picture_runs"],
        "raw_unknown_runs": raw_run_stats["unknown_runs"],
        "optimized_unknown_runs": opt_run_stats["unknown_runs"],
        "paragraphs_changed": changed,
        "paragraphs_with_new_tabs": paragraphs_with_new_tabs,
        "consecutive_tab_run_issues": consecutive_issues,
        "invalid_tab_pos_count": invalid_tab_pos_count,
        "suspicious_tab_pos_count": suspicious_tab_pos_count,
        "validation_ok": 1 if not errors else 0,
        "validation_errors_json": json.dumps(errors, ensure_ascii=False),
    }
    summary["summary_json"] = json.dumps(summary, ensure_ascii=False, sort_keys=True)
    return summary


def _word_pid_from_hwnd(hwnd: int) -> int | None:
    try:
        import win32process  # type: ignore

        if not hwnd:
            return None

        _, pid = win32process.GetWindowThreadProcessId(int(hwnd))
        return int(pid) if pid else None
    except Exception:
        return None


def saveas_worker(task_q: mp.Queue, result_q: mp.Queue) -> None:
    app = None
    word_pid: int | None = None
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore

        pythoncom.CoInitialize()
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        try:
            hwnd = int(app.Hwnd)
        except Exception:
            hwnd = 0
        word_pid = _word_pid_from_hwnd(hwnd)

        result_q.put({
            "kind": "worker_started",
            "word_pid": word_pid,
            "stdout_lines": [f"opening WordCom pid={word_pid}"],
        })

        while True:
            job = task_q.get()
            if job is None:
                break

            uid = str(job["uid"])
            source_path = str(job["source_path"])
            out_path = str(job["materialized_path"])
            doc = None
            started = time.time()
            stdout_lines = [
                f"input={source_path}",
                f"output={out_path}",
            ]
            status = "ok"
            exit_code = 0
            error_message = None

            try:
                Path(out_path).parent.mkdir(parents=True, exist_ok=True)
                doc = app.Documents.Open(source_path, ReadOnly=True)
                doc.SaveAs2(out_path, FileFormat=16)
                if not Path(out_path).exists():
                    raise FileNotFoundError(f"Output file not found after SaveAs2: {out_path}")
                stdout_lines.append("ok")
            except Exception as exc:
                status = "error"
                exit_code = 1
                error_message = f"{type(exc).__name__}: {exc}"
                stdout_lines.append(error_message)
            finally:
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass

            duration_ms = int((time.time() - started) * 1000)
            result_q.put({
                "kind": "document_result",
                "uid": uid,
                "status": status,
                "exit_code": exit_code,
                "stdout_lines": stdout_lines,
                "duration_ms": duration_ms,
                "error_message": error_message,
                "word_pid": word_pid,
            })

    except Exception as exc:
        result_q.put({
            "kind": "worker_fatal",
            "status": "error",
            "exit_code": 1,
            "stdout_lines": [f"{type(exc).__name__}: {exc}"],
            "duration_ms": 0,
            "error_message": f"{type(exc).__name__}: {exc}",
            "word_pid": word_pid,
        })
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        try:
            import pythoncom  # type: ignore
            pythoncom.CoUninitialize()
        except Exception:
            pass


class SaveAsBatchSession:
    def __init__(self, log_path: Path, timeout_sec: int):
        self.log_path = log_path
        self.timeout_sec = timeout_sec
        self.proc: mp.Process | None = None
        self.task_q: mp.Queue | None = None
        self.result_q: mp.Queue | None = None
        self.word_pid: int | None = None

    def start(self) -> None:
        self.task_q = mp.Queue()
        self.result_q = mp.Queue()
        self.proc = mp.Process(target=saveas_worker, args=(self.task_q, self.result_q), daemon=True)
        self.proc.start()
        started = time.time()
        while True:
            remaining = max(0.1, self.timeout_sec - (time.time() - started))
            try:
                msg = self.result_q.get(timeout=remaining)
            except Empty:
                self.terminate()
                raise RuntimeError("saveas worker start timeout")

            if msg.get("kind") == "worker_started":
                self.word_pid = msg.get("word_pid")
                log_line(self.log_path, "[run_saveas_batch]")
                for line in msg.get("stdout_lines", []):
                    log_line(self.log_path, line)
                return

            if msg.get("kind") == "worker_fatal":
                self.terminate()
                raise RuntimeError(msg.get("error_message") or "saveas worker fatal error")

    def terminate(self) -> None:
        if self.proc is not None and self.proc.is_alive():
            self.proc.terminate()
            self.proc.join(timeout=2)

        if self.proc is not None and self.proc.is_alive():
            try:
                self.proc.kill()
            except Exception:
                pass

        try:
            import subprocess

            if self.word_pid:
                subprocess.run(
                    ["taskkill", "/PID", str(self.word_pid), "/F", "/T"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    check=False,
                )
            else:
                subprocess.run(
                    ["taskkill", "/IM", "WINWORD.EXE", "/F", "/T"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    check=False,
                )
        except Exception:
            pass

        self.proc = None
        self.task_q = None
        self.result_q = None
        self.word_pid = None

    def shutdown(self) -> None:
        try:
            if self.task_q is not None:
                self.task_q.put(None)
        except Exception:
            pass
        if self.proc is not None:
            self.proc.join(timeout=5)
        self.terminate()

    def save_one(self, uid: str, source_path: Path, materialized_path: Path) -> dict[str, Any]:
        if self.task_q is None or self.result_q is None:
            raise RuntimeError("saveas session is not started")

        self.task_q.put({
            "uid": uid,
            "source_path": str(source_path),
            "materialized_path": str(materialized_path),
        })

        try:
            msg = self.result_q.get(timeout=self.timeout_sec)
        except Empty:
            self.terminate()
            return {
                "uid": uid,
                "status": "timeout",
                "exit_code": 1,
                "stdout_lines": [f"timeout after {self.timeout_sec}s"],
                "duration_ms": self.timeout_sec * 1000,
                "error_message": f"timeout>{self.timeout_sec}s",
            }

        if msg.get("kind") == "document_result":
            return msg

        if msg.get("kind") == "worker_fatal":
            self.terminate()
            return {
                "uid": uid,
                "status": "error",
                "exit_code": 1,
                "stdout_lines": msg.get("stdout_lines", []),
                "duration_ms": msg.get("duration_ms", 0),
                "error_message": msg.get("error_message"),
            }

        self.terminate()
        return {
            "uid": uid,
            "status": "error",
            "exit_code": 1,
            "stdout_lines": [f"unexpected worker message: {msg!r}"],
            "duration_ms": 0,
            "error_message": "unexpected_worker_message",
        }


def run_subprocess_with_retry(
    cmd: list[str],
    log_path: Path,
    timeout_sec: int,
    retries: int,
    retry_sleep_sec: int,
) -> tuple[bool, str | None]:
    last_error: str | None = None

    for attempt in range(1, retries + 1):
        started = time.time()
        log_line(log_path, f"[attempt] {attempt}/{retries}")
        log_line(log_path, "[command]")
        log_line(log_path, " ".join(cmd))

        try:
            completed = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout_sec,
            )
            duration_ms = int((time.time() - started) * 1000)

            log_line(log_path, "[stdout]")
            if completed.stdout:
                for line in completed.stdout.splitlines():
                    log_line(log_path, line)

            log_line(log_path, "[stderr]")
            if completed.stderr:
                for line in completed.stderr.splitlines():
                    log_line(log_path, line)

            log_line(log_path, "[exit_code]")
            log_line(log_path, str(completed.returncode))
            log_line(log_path, "[duration_ms]")
            log_line(log_path, str(duration_ms))

            if completed.returncode == 0:
                return True, None

            last_error = f"exit_code={completed.returncode}"

        except subprocess.TimeoutExpired:
            duration_ms = int((time.time() - started) * 1000)
            log_line(log_path, "[stdout]")
            log_line(log_path, f"timeout after {timeout_sec}s")
            log_line(log_path, "[exit_code]")
            log_line(log_path, "1")
            log_line(log_path, "[duration_ms]")
            log_line(log_path, str(duration_ms))
            last_error = f"timeout>{timeout_sec}s"

        except Exception as exc:
            duration_ms = int((time.time() - started) * 1000)
            log_line(log_path, "[stdout]")
            log_line(log_path, f"{type(exc).__name__}: {exc}")
            log_line(log_path, "[exit_code]")
            log_line(log_path, "1")
            log_line(log_path, "[duration_ms]")
            log_line(log_path, str(duration_ms))
            last_error = f"{type(exc).__name__}: {exc}"

        if attempt < retries:
            time.sleep(retry_sleep_sec)

    return False, last_error


def ensure_valid_json_file(path: Path) -> None:
    if not path.exists():
        raise FileNotFoundError(f"JSON file not found: {path}")
    if path.stat().st_size == 0:
        raise ValueError(f"JSON file is empty: {path}")
    with path.open("r", encoding="utf-8") as f:
        json.load(f)


def process_saveas_batches(
    docs: list[dict[str, Any]],
    log_path: Path,
    batch_size: int,
    timeout_sec: int,
    retries: int,
    retry_sleep_sec: int,
    restart_every: int,
) -> set[int]:
    conn = get_connection()
    ok_ids: set[int] = set()

    try:
        batches = [docs[i:i + batch_size] for i in range(0, len(docs), batch_size)]

        for batch_idx, batch in enumerate(batches, start=1):
            log_section(log_path, f"SAVEAS BATCH {batch_idx}/{len(batches)} size={len(batch)}")
            session: SaveAsBatchSession | None = None
            docs_since_restart = 0

            try:
                for doc in batch:
                    doc_id = int(doc["id"])
                    uid = str(doc["uid"])
                    current_status = str(doc["processing_status"])
                    source_path = Path(doc["source_abs_path"])
                    artifact_dir = Path(doc["artifacts_abs_path"])
                    artifact_dir.mkdir(parents=True, exist_ok=True)
                    materialized_path = artifact_dir / "materialized.docx"

                    log_section(log_path, uid)
                    log_line(log_path, f"[current_status] {current_status}")
                    update_status(conn, doc_id, "materializing")

                    for attempt in range(1, retries + 1):
                        if session is None:
                            session = SaveAsBatchSession(log_path=log_path, timeout_sec=timeout_sec)
                            session.start()
                            docs_since_restart = 0

                        log_line(log_path, f"[run_saveas_batch] attempt={attempt}/{retries}")
                        log_line(log_path, f"[run_saveas_batch] word_pid={session.word_pid}")
                        result = session.save_one(uid=uid, source_path=source_path, materialized_path=materialized_path)

                        log_line(log_path, "[stdout]")
                        for line in result.get("stdout_lines", []):
                            log_line(log_path, line)
                        log_line(log_path, "[exit_code]")
                        log_line(log_path, str(result.get("exit_code", 1)))
                        log_line(log_path, "[duration_ms]")
                        log_line(log_path, str(result.get("duration_ms", 0)))

                        if result.get("status") == "ok" and materialized_path.exists():
                            update_status(conn, doc_id, "materialized")
                            ok_ids.add(doc_id)
                            docs_since_restart += 1
                            log_line(log_path, "[final_status] materialized")
                            break

                        if session is not None:
                            session.terminate()
                            session = None
                        if attempt < retries:
                            time.sleep(retry_sleep_sec)
                    else:
                        update_status(conn, doc_id, "failed_saveas")
                        log_line(log_path, "[final_status] failed_saveas")

                    if session is not None and docs_since_restart >= restart_every:
                        session.shutdown()
                        session = None
                        docs_since_restart = 0
            finally:
                if session is not None:
                    session.shutdown()
    finally:
        conn.close()

    return ok_ids


def process_attrs(
    docs: list[dict[str, Any]],
    log_path: Path,
    attrs_timeout_sec: int,
    retries: int,
    retry_sleep_sec: int,
) -> None:
    python_exe = sys.executable
    add_attrs_script = PROJECT_ROOT / "src" / "docx_pipeline" / "pipeline" / "saveas" / "add_custom_attrs.py"

    conn = get_connection()
    try:
        for doc in docs:
            doc_id = int(doc["id"])
            uid = str(doc["uid"])
            current_status = str(doc["processing_status"])
            artifact_dir = Path(doc["artifacts_abs_path"])
            materialized_path = artifact_dir / "materialized.docx"
            with_ids_path = artifact_dir / "materialized_with_ids.docx"

            log_section(log_path, uid)
            log_line(log_path, f"[current_status] {current_status}")

            update_status(conn, doc_id, "injecting_ids")
            ok, error = run_subprocess_with_retry(
                cmd=[python_exe, str(add_attrs_script), "--in", str(materialized_path), "--out", str(with_ids_path)],
                log_path=log_path,
                timeout_sec=attrs_timeout_sec,
                retries=retries,
                retry_sleep_sec=retry_sleep_sec,
            )
            if not ok:
                update_status(conn, doc_id, "failed_attrs")
                log_line(log_path, f"[final_status] failed_attrs error={error}")
                continue

            if not with_ids_path.exists():
                update_status(conn, doc_id, "failed_attrs")
                log_line(log_path, f"[final_status] failed_attrs error=FileNotFoundError: {with_ids_path} not found")
                continue

            update_status(conn, doc_id, "ids_injected")
            log_line(log_path, "[final_status] ids_injected")
    finally:
        conn.close()


def process_parser(
    docs: list[dict[str, Any]],
    log_path: Path,
    parser_timeout_sec: int,
    retries: int,
    retry_sleep_sec: int,
) -> None:
    python_exe = sys.executable
    parser_script = PROJECT_ROOT / "src" / "docx_pipeline" / "pipeline" / "parse" / "parser.py"

    conn = get_connection()
    try:
        for doc in docs:
            doc_id = int(doc["id"])
            uid = str(doc["uid"])
            current_status = str(doc["processing_status"])
            artifact_dir = Path(doc["artifacts_abs_path"])
            with_ids_path = artifact_dir / "materialized_with_ids.docx"
            raw_json_path = artifact_dir / "raw.json"

            log_section(log_path, uid)
            log_line(log_path, f"[current_status] {current_status}")

            update_status(conn, doc_id, "parsing_raw")
            ok, error = run_subprocess_with_retry(
                cmd=[python_exe, str(parser_script), "--in", str(with_ids_path), "--out", str(raw_json_path)],
                log_path=log_path,
                timeout_sec=parser_timeout_sec,
                retries=retries,
                retry_sleep_sec=retry_sleep_sec,
            )
            if not ok:
                update_status(conn, doc_id, "failed_parser")
                log_line(log_path, f"[final_status] failed_parser error={error}")
                continue

            try:
                ensure_valid_json_file(raw_json_path)
            except Exception as exc:
                update_status(conn, doc_id, "failed_parser")
                log_line(log_path, f"[final_status] failed_parser error={type(exc).__name__}: {exc}")
                continue

            update_status(conn, doc_id, "parsed_raw")
            log_line(log_path, "[final_status] parsed_raw")
    finally:
        conn.close()


def process_optimizer(
    docs: list[dict[str, Any]],
    log_path: Path,
    optimizer_timeout_sec: int,
    retries: int,
    retry_sleep_sec: int,
) -> None:
    settings = get_settings()
    python_exe = sys.executable
    optimizer_script = PROJECT_ROOT / "tools" / "optimization" / "optimizer_layout_v2.py"

    conn = get_connection()
    try:
        for doc in docs:
            doc_id = int(doc["id"])
            uid = str(doc["uid"])
            current_status = str(doc["processing_status"])
            artifact_dir = Path(doc["artifacts_abs_path"])
            raw_json_path = artifact_dir / "raw.json"
            raw_optimized_path = artifact_dir / "raw.json"

            log_section(log_path, uid)
            log_line(log_path, f"[current_status] {current_status}")

            update_status(conn, doc_id, "optimizing")
            ok, error = run_subprocess_with_retry(
                cmd=[
                    python_exe,
                    str(optimizer_script),
                    "--db",
                    str(settings.db_path),
                    "--uid",
                    uid,
                    "--status",
                    "optimizing",
                    "--output-status",
                    "optimized",
                ],
                log_path=log_path,
                timeout_sec=optimizer_timeout_sec,
                retries=retries,
                retry_sleep_sec=retry_sleep_sec,
            )
            if not ok:
                update_status(conn, doc_id, "failed_optimizer")
                log_line(log_path, f"[final_status] failed_optimizer error={error}")
                continue

            try:
                ensure_valid_json_file(raw_json_path)
                ensure_valid_json_file(raw_optimized_path)
            except Exception as exc:
                update_status(conn, doc_id, "failed_optimizer")
                log_line(log_path, f"[final_status] failed_optimizer error={type(exc).__name__}: {exc}")
                continue

            update_status(conn, doc_id, "optimized")
            log_line(log_path, "[final_status] optimized")
    finally:
        conn.close()


def insert_validation_stats(
    conn: sqlite3.Connection,
    *,
    doc_id: int,
    uid: str,
    source_status: str,
    final_status: str,
    stats: dict[str, Any],
) -> None:
    conn.execute(
        f"""
        INSERT INTO {VALIDATION_TABLE} (
            run_id,
            document_id,
            uid,
            source_status,
            final_status,
            raw_paragraph_count,
            optimized_paragraph_count,
            raw_table_count,
            optimized_table_count,
            raw_text_chars,
            optimized_text_chars,
            raw_tab_runs,
            optimized_tab_runs,
            raw_pformat_tabs,
            optimized_pformat_tabs,
            raw_shape_runs,
            optimized_shape_runs,
            raw_picture_runs,
            optimized_picture_runs,
            raw_unknown_runs,
            optimized_unknown_runs,
            paragraphs_changed,
            paragraphs_with_new_tabs,
            consecutive_tab_run_issues,
            invalid_tab_pos_count,
            suspicious_tab_pos_count,
            validation_ok,
            validation_errors_json,
            summary_json
        ) VALUES (
            ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
        )
        """,
        (
            RECONSTRUCTION_RUN_ID,
            doc_id,
            uid,
            source_status,
            final_status,
            stats["raw_paragraph_count"],
            stats["optimized_paragraph_count"],
            stats["raw_table_count"],
            stats["optimized_table_count"],
            stats["raw_text_chars"],
            stats["optimized_text_chars"],
            stats["raw_tab_runs"],
            stats["optimized_tab_runs"],
            stats["raw_pformat_tabs"],
            stats["optimized_pformat_tabs"],
            stats["raw_shape_runs"],
            stats["optimized_shape_runs"],
            stats["raw_picture_runs"],
            stats["optimized_picture_runs"],
            stats["raw_unknown_runs"],
            stats["optimized_unknown_runs"],
            stats["paragraphs_changed"],
            stats["paragraphs_with_new_tabs"],
            stats["consecutive_tab_run_issues"],
            stats["invalid_tab_pos_count"],
            stats["suspicious_tab_pos_count"],
            stats["validation_ok"],
            stats["validation_errors_json"],
            stats["summary_json"],
        ),
    )
    conn.commit()


def process_reconstruction(
    docs: list[dict[str, Any]],
    log_path: Path,
    reconstructor_timeout_sec: int,
    retries: int,
    retry_sleep_sec: int,
) -> None:
    conn = get_connection()
    try:
        ensure_validation_table(conn)
        python_exe = sys.executable
        reconstructor_script = PROJECT_ROOT / "src" / "reconstructor.py"

        for doc in docs:
            doc_id = int(doc["id"])
            uid = str(doc["uid"])
            current_status = str(doc["processing_status"])
            artifact_dir = Path(doc["artifacts_abs_path"])
            raw_json_path = artifact_dir / INPUT_RAW_NAME
            raw_optimized_path = artifact_dir / INPUT_OPTIMIZED_NAME
            donor_docx_path = artifact_dir / INPUT_DONOR_NAME
            output_docx_path = artifact_dir / OUTPUT_DOCX_NAME

            log_section(log_path, uid)
            log_line(log_path, f"[current_status] {current_status}")

            try:
                raw_doc = load_json_file(raw_json_path)
                optimized_doc = load_json_file(raw_optimized_path)
                validation_stats = collect_validation_stats(raw_doc, optimized_doc)
                log_line(log_path, f"[validation_ok] {validation_stats['validation_ok']}")
                log_line(log_path, f"[validation_errors] {validation_stats['validation_errors_json']}")
            except Exception as exc:
                update_status(conn, doc_id, "failed_reconstructor")
                log_line(log_path, f"[final_status] failed_reconstructor error={type(exc).__name__}: {exc}")
                continue

            update_status(conn, doc_id, "reconstructing")
            ok, error = run_subprocess_with_retry(
                cmd=[
                    python_exe,
                    str(reconstructor_script),
                    "--in-json",
                    str(raw_optimized_path),
                    "--out-docx",
                    str(output_docx_path),
                    "--donor-docx",
                    str(donor_docx_path),
                ],
                log_path=log_path,
                timeout_sec=reconstructor_timeout_sec,
                retries=retries,
                retry_sleep_sec=retry_sleep_sec,
            )
            if not ok or not output_docx_path.exists():
                update_status(conn, doc_id, "failed_reconstructor")
                insert_validation_stats(conn, doc_id=doc_id, uid=uid, source_status=current_status, final_status="failed_reconstructor", stats=validation_stats)
                log_line(log_path, f"[final_status] failed_reconstructor error={error or 'output_docx_missing'}")
                continue

            update_status(conn, doc_id, "reconstructed")
            insert_validation_stats(conn, doc_id=doc_id, uid=uid, source_status=current_status, final_status="reconstructed", stats=validation_stats)
            log_line(log_path, "[final_status] reconstructed")
    finally:
        conn.close()


def split_docs_by_stage(
    docs: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    saveas_statuses = {
        "discovered",
        "failed",
        "failed_saveas",
        "materializing",
    }
    attrs_statuses = {
        "materialized",
        "failed_attrs",
        "injecting_ids",
    }
    parser_statuses = {
        "ids_injected",
        "failed_parser",
        "parsing_raw",
    }
    optimizer_statuses = {
        "parsed_raw",
        "failed_optimizer",
        "optimizing",
    }
    reconstruction_statuses = {
        "optimized",
        "failed_reconstructor",
        "reconstructing",
    }

    saveas_docs: list[dict[str, Any]] = []
    attrs_docs: list[dict[str, Any]] = []
    parser_docs: list[dict[str, Any]] = []
    optimizer_docs: list[dict[str, Any]] = []
    reconstruction_docs: list[dict[str, Any]] = []

    for doc in docs:
        status = str(doc["processing_status"])
        if status in saveas_statuses:
            saveas_docs.append(doc)
        elif status in attrs_statuses:
            attrs_docs.append(doc)
        elif status in parser_statuses:
            parser_docs.append(doc)
        elif status in optimizer_statuses:
            optimizer_docs.append(doc)
        elif status in reconstruction_statuses:
            reconstruction_docs.append(doc)

    return saveas_docs, attrs_docs, parser_docs, optimizer_docs, reconstruction_docs


def refresh_doc_statuses(docs: list[dict[str, Any]]) -> list[dict[str, Any]]:
    if not docs:
        return []

    ids = [int(doc["id"]) for doc in docs]
    placeholders = ",".join("?" for _ in ids)

    conn = get_connection()
    try:
        rows = conn.execute(
            f"""
            SELECT
                id,
                uid,
                mail_uid,
                attachment_index,
                source_filename,
                source_abs_path,
                artifacts_abs_path,
                processing_status
            FROM document
            WHERE id IN ({placeholders})
            ORDER BY id
            """,
            ids,
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def main() -> None:
    mp.freeze_support()

    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--batch-size", type=int, default=100)
    parser.add_argument("--saveas-timeout-sec", type=int, default=5)
    parser.add_argument("--attrs-timeout-sec", type=int, default=5)
    parser.add_argument("--parser-timeout-sec", type=int, default=30)
    parser.add_argument("--optimizer-timeout-sec", type=int, default=60)
    parser.add_argument("--reconstructor-timeout-sec", type=int, default=180)
    parser.add_argument("--retries", type=int, default=3)
    parser.add_argument("--retry-sleep-sec", type=int, default=1)
    parser.add_argument("--restart-every", type=int, default=100)
    args = parser.parse_args()

    run_started = time.time()
    log_dir = PROJECT_ROOT / "tools" / "log"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"create_optimized_{time.strftime('%Y%m%d_%H%M%S')}.log"

    docs = fetch_documents(args.limit)

    log_section(log_path, "RUN START")
    log_line(log_path, f"[start_time] {ts()}")
    log_line(log_path, f"[documents_requested] {args.limit if args.limit is not None else 'ALL'}")
    log_line(log_path, f"[documents_selected] {len(docs)}")
    log_line(log_path, f"[batch_size] {args.batch_size}")
    log_line(log_path, f"[project_root] {PROJECT_ROOT}")

    if not docs:
        log_line(log_path, "[run_status] nothing_to_do")
        print(f"Run log: {log_path}")
        return

    saveas_docs, attrs_docs, parser_docs, optimizer_docs, reconstruction_docs = split_docs_by_stage(docs)

    log_line(log_path, f"[saveas_docs] {len(saveas_docs)}")
    log_line(log_path, f"[attrs_docs] {len(attrs_docs)}")
    log_line(log_path, f"[parser_docs] {len(parser_docs)}")
    log_line(log_path, f"[optimizer_docs] {len(optimizer_docs)}")
    log_line(log_path, f"[reconstruction_docs] {len(reconstruction_docs)}")

    if saveas_docs:
        ok_ids = process_saveas_batches(
            docs=saveas_docs,
            log_path=log_path,
            batch_size=args.batch_size,
            timeout_sec=args.saveas_timeout_sec,
            retries=args.retries,
            retry_sleep_sec=args.retry_sleep_sec,
            restart_every=args.restart_every,
        )
        saveas_ok_docs = [doc for doc in saveas_docs if int(doc["id"]) in ok_ids]
        if saveas_ok_docs:
            attrs_docs = attrs_docs + refresh_doc_statuses(saveas_ok_docs)

    if attrs_docs:
        process_attrs(
            docs=attrs_docs,
            log_path=log_path,
            attrs_timeout_sec=args.attrs_timeout_sec,
            retries=args.retries,
            retry_sleep_sec=args.retry_sleep_sec,
        )
        parser_docs = parser_docs + refresh_doc_statuses(attrs_docs)

    parser_docs = [doc for doc in parser_docs if str(doc["processing_status"]) == "ids_injected"]
    if parser_docs:
        process_parser(
            docs=parser_docs,
            log_path=log_path,
            parser_timeout_sec=args.parser_timeout_sec,
            retries=args.retries,
            retry_sleep_sec=args.retry_sleep_sec,
        )
        optimizer_docs = optimizer_docs + refresh_doc_statuses(parser_docs)

    optimizer_docs = [doc for doc in optimizer_docs if str(doc["processing_status"]) == "parsed_raw"]
    if optimizer_docs:
        process_optimizer(
            docs=optimizer_docs,
            log_path=log_path,
            optimizer_timeout_sec=args.optimizer_timeout_sec,
            retries=args.retries,
            retry_sleep_sec=args.retry_sleep_sec,
        )
        reconstruction_docs = reconstruction_docs + refresh_doc_statuses(optimizer_docs)

    reconstruction_docs = [doc for doc in reconstruction_docs if str(doc["processing_status"]) == "optimized"]
    if reconstruction_docs:
        process_reconstruction(
            docs=reconstruction_docs,
            log_path=log_path,
            reconstructor_timeout_sec=args.reconstructor_timeout_sec,
            retries=args.retries,
            retry_sleep_sec=args.retry_sleep_sec,
        )

    total_duration_ms = int((time.time() - run_started) * 1000)
    log_section(log_path, "RUN END")
    log_line(log_path, f"[reconstruction_run_id] {RECONSTRUCTION_RUN_ID}")
    log_line(log_path, f"[duration_ms] {total_duration_ms}")
    print(f"Run log: {log_path}")


if __name__ == "__main__":
    main()