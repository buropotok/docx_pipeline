
from __future__ import annotations

import argparse
import json
import multiprocessing as mp
import sqlite3
import subprocess
import sys
import time
from pathlib import Path
from queue import Empty
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[1]
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
                'failed',
                'failed_saveas',
                'failed_attrs',
                'failed_parser'
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

        # fallback: если Word остался висеть
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
                    source_path = Path(doc["source_abs_path"])
                    artifact_dir = Path(doc["artifacts_abs_path"])
                    artifact_dir.mkdir(parents=True, exist_ok=True)
                    materialized_path = artifact_dir / "materialized.docx"

                    log_section(log_path, uid)
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

                        # Any error/timeout: restart word and retry current doc.
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


def process_attrs_and_parser(
    docs: list[dict[str, Any]],
    log_path: Path,
    attrs_timeout_sec: int,
    parser_timeout_sec: int,
    retries: int,
    retry_sleep_sec: int,
) -> None:
    settings = get_settings()
    python_exe = sys.executable
    add_attrs_script = settings.project_root / "src" / "docx_pipeline" / "pipeline" / "saveas" / "add_custom_attrs.py"
    parser_script = settings.project_root / "src" / "docx_pipeline" / "pipeline" / "parse" / "parser.py"

    conn = get_connection()
    try:
        for doc in docs:
            doc_id = int(doc["id"])
            uid = str(doc["uid"])
            artifact_dir = Path(doc["artifacts_abs_path"])
            materialized_path = artifact_dir / "materialized.docx"
            with_ids_path = artifact_dir / "materialized_with_ids.docx"
            raw_json_path = artifact_dir / "raw.json"

            log_section(log_path, uid)

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

            update_status(conn, doc_id, "ids_injected")

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


def main() -> None:
    mp.freeze_support()

    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--batch-size", type=int, default=100)
    parser.add_argument("--saveas-timeout-sec", type=int, default=5)
    parser.add_argument("--attrs-timeout-sec", type=int, default=5)
    parser.add_argument("--parser-timeout-sec", type=int, default=30)
    parser.add_argument("--retries", type=int, default=3)
    parser.add_argument("--retry-sleep-sec", type=int, default=1)
    parser.add_argument("--restart-every", type=int, default=100)
    args = parser.parse_args()

    settings = get_settings()
    run_started = time.time()
    log_dir = settings.project_root / "tools" / "log"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"create_raw_{time.strftime('%Y%m%d_%H%M%S')}.log"

    docs = fetch_documents(args.limit)

    log_section(log_path, "RUN START")
    log_line(log_path, f"[start_time] {ts()}")
    log_line(log_path, f"[documents_requested] {args.limit if args.limit is not None else 'ALL'}")
    log_line(log_path, f"[documents_selected] {len(docs)}")
    log_line(log_path, f"[batch_size] {args.batch_size}")

    if not docs:
        log_line(log_path, "[run_status] nothing_to_do")
        print(f"Run log: {log_path}")
        return

    ok_ids = process_saveas_batches(
        docs=docs,
        log_path=log_path,
        batch_size=args.batch_size,
        timeout_sec=args.saveas_timeout_sec,
        retries=args.retries,
        retry_sleep_sec=args.retry_sleep_sec,
        restart_every=args.restart_every,
    )

    ok_docs = [doc for doc in docs if int(doc["id"]) in ok_ids]

    process_attrs_and_parser(
        docs=ok_docs,
        log_path=log_path,
        attrs_timeout_sec=args.attrs_timeout_sec,
        parser_timeout_sec=args.parser_timeout_sec,
        retries=args.retries,
        retry_sleep_sec=args.retry_sleep_sec,
    )

    total_duration_ms = int((time.time() - run_started) * 1000)
    log_section(log_path, "RUN END")
    log_line(log_path, f"[duration_ms] {total_duration_ms}")
    print(f"Run log: {log_path}")


if __name__ == "__main__":
    main()
