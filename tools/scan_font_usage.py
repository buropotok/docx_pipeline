#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import json
import os
import sqlite3
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, Iterable, Iterator, List, Optional, Set, Tuple


DEFAULT_RAW_FILENAME = "raw.json"


@dataclass(frozen=True)
class FontUsageSignature:
    font_title: str
    normalized_title: str
    size_pt: Optional[float]
    bold: Optional[int]
    italic: Optional[int]
    underline: Optional[int]
    all_caps: Optional[int]
    small_caps: Optional[int]
    strike: Optional[int]
    double_strike: Optional[int]
    outline: Optional[int]
    shadow: Optional[int]
    emboss: Optional[int]
    imprint: Optional[int]
    rtl: Optional[int]
    lang: Optional[str]


@dataclass
class UsageAggregate:
    runs_count: int = 0
    chars_count: int = 0
    paragraph_ids: Optional[Set[str]] = None
    table_ids: Optional[Set[str]] = None
    first_seen_path: Optional[str] = None

    def __post_init__(self) -> None:
        if self.paragraph_ids is None:
            self.paragraph_ids = set()
        if self.table_ids is None:
            self.table_ids = set()


class ScanError(Exception):
    pass


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Scan raw.json files and populate font/font_usage/document_font_usage tables."
    )
    parser.add_argument("--db", required=True, help="Path to SQLite database.")
    parser.add_argument(
        "--artifacts-root",
        default=None,
        help="Artifacts root. Optional if document.artifacts_abs_path is populated.",
    )
    parser.add_argument(
        "--raw-filename",
        default=DEFAULT_RAW_FILENAME,
        help="Artifact filename to read. Default: raw.json",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Limit number of documents to scan.",
    )
    parser.add_argument(
        "--uid",
        action="append",
        default=[],
        help="Scan only specific document uid. Can be repeated.",
    )
    parser.add_argument(
        "--status",
        action="append",
        default=["parsed_raw"],
        help="Filter by document.processing_status. Default: parsed_raw. Can be repeated.",
    )
    parser.add_argument(
        "--where",
        default=None,
        help="Additional SQL WHERE fragment for document selection.",
    )
    parser.add_argument(
        "--commit-every",
        type=int,
        default=100,
        help="Commit after every N scanned documents. Default: 100",
    )
    parser.add_argument(
        "--reset-document-data",
        action="store_true",
        help="Delete document_font_usage and document_font_profile for each scanned document before reinsert.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Parse and aggregate, but do not write to the DB.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Verbose logging.",
    )
    return parser.parse_args()


def now_utc() -> str:
    return datetime.utcnow().isoformat(timespec="seconds") + "Z"


def normalize_font_title(title: Optional[str]) -> str:
    value = (title or "").strip()
    if not value:
        return "unknown"
    return " ".join(value.split()).lower()


def normalize_lang(lang: Optional[str]) -> Optional[str]:
    if lang is None:
        return None
    lang = lang.strip()
    return lang or None


def to_sql_bool(value: Any) -> Optional[int]:
    if value is None:
        return None
    return 1 if bool(value) else 0


def normalize_underline(value: Any) -> Optional[int]:
    if value is None:
        return None
    if value is False:
        return 0
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in {"", "none", "false"}:
            return 0
        return 1
    return 1 if bool(value) else 0


def half_points_to_pt(value: Any) -> Optional[float]:
    if value is None:
        return None
    try:
        return float(value) / 2.0
    except (TypeError, ValueError):
        return None


def pick_font_name(rfonts: Optional[Dict[str, Any]]) -> Optional[str]:
    if not isinstance(rfonts, dict):
        return None
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        value = rfonts.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()
    return None


class StyleResolver:
    def __init__(self, data: Dict[str, Any]) -> None:
        self.doc_defaults = data.get("doc_defaults", {}) or {}
        self.styles = data.get("styles", {}) or {}
        self._p_style_cache: Dict[str, Dict[str, Any]] = {}
        self._r_style_cache: Dict[str, Dict[str, Any]] = {}

    def resolve_paragraph_r_format(self, p_style_id: Optional[str]) -> Dict[str, Any]:
        if not p_style_id:
            return dict((self.doc_defaults.get("r_format") or {}))
        if p_style_id in self._p_style_cache:
            return dict(self._p_style_cache[p_style_id])

        merged = dict((self.doc_defaults.get("r_format") or {}))
        chain = self._style_chain(p_style_id)
        for style in chain:
            if style.get("type") == "paragraph":
                merged = deep_merge_r_format(merged, style.get("r_format") or {})
        self._p_style_cache[p_style_id] = merged
        return dict(merged)

    def resolve_run_style_r_format(self, r_style_id: Optional[str], char_style_id: Optional[str]) -> Dict[str, Any]:
        style_id = r_style_id or char_style_id
        if not style_id:
            return {}
        if style_id in self._r_style_cache:
            return dict(self._r_style_cache[style_id])

        merged: Dict[str, Any] = {}
        chain = self._style_chain(style_id)
        for style in chain:
            if style.get("type") in {"character", "paragraph"}:
                merged = deep_merge_r_format(merged, style.get("r_format") or {})
        self._r_style_cache[style_id] = merged
        return dict(merged)

    def _style_chain(self, style_id: str) -> List[Dict[str, Any]]:
        chain: List[Dict[str, Any]] = []
        seen: Set[str] = set()
        current = style_id
        while current and current not in seen and current in self.styles:
            seen.add(current)
            style = self.styles[current] or {}
            chain.append(style)
            current = style.get("based_on")
        chain.reverse()
        return chain


def deep_merge_r_format(base: Dict[str, Any], override: Dict[str, Any]) -> Dict[str, Any]:
    result = dict(base or {})
    for key, value in (override or {}).items():
        if key == "rFonts" and isinstance(value, dict):
            fonts = dict(result.get("rFonts") or {})
            fonts.update(value)
            result["rFonts"] = fonts
        else:
            result[key] = value
    return result


def build_signature(effective_r_format: Dict[str, Any]) -> FontUsageSignature:
    font_title = pick_font_name(effective_r_format.get("rFonts")) or "Unknown"
    normalized_title = normalize_font_title(font_title)
    return FontUsageSignature(
        font_title=font_title,
        normalized_title=normalized_title,
        size_pt=half_points_to_pt(effective_r_format.get("font_size_half_points")),
        bold=to_sql_bool(effective_r_format.get("bold")),
        italic=to_sql_bool(effective_r_format.get("italic")),
        underline=normalize_underline(effective_r_format.get("underline")),
        all_caps=to_sql_bool(effective_r_format.get("caps")),
        small_caps=to_sql_bool(effective_r_format.get("small_caps")),
        strike=to_sql_bool(effective_r_format.get("strike")),
        double_strike=to_sql_bool(effective_r_format.get("double_strike")),
        outline=to_sql_bool(effective_r_format.get("outline")),
        shadow=to_sql_bool(effective_r_format.get("shadow")),
        emboss=to_sql_bool(effective_r_format.get("emboss")),
        imprint=to_sql_bool(effective_r_format.get("imprint")),
        rtl=to_sql_bool(effective_r_format.get("rtl")),
        lang=normalize_lang(effective_r_format.get("lang")),
    )


def iter_paragraphs(content: List[Dict[str, Any]], table_stack: Optional[List[str]] = None) -> Iterator[Tuple[Dict[str, Any], List[str]]]:
    table_stack = list(table_stack or [])
    for item in content or []:
        item_type = item.get("type")
        if item_type == "paragraph":
            yield item, table_stack
        elif item_type == "table":
            table_id = item.get("id") or "<table>"
            next_stack = table_stack + [table_id]
            for row in item.get("rows", []) or []:
                for cell in row.get("cells", []) or []:
                    yield from iter_paragraphs(cell.get("content", []) or [], next_stack)


def aggregate_document(data: Dict[str, Any]) -> Tuple[Dict[FontUsageSignature, UsageAggregate], Counter, Counter]:
    resolver = StyleResolver(data)
    usage_map: Dict[FontUsageSignature, UsageAggregate] = {}
    font_counter: Counter = Counter()
    font_usage_counter: Counter = Counter()

    for paragraph, table_stack in iter_paragraphs(data.get("content", []) or []):
        p_id = paragraph.get("id") or "<paragraph>"
        p_style_id = paragraph.get("p_style_id") or paragraph.get("style_id")
        paragraph_base = resolver.resolve_paragraph_r_format(p_style_id)

        for run in paragraph.get("runs", []) or []:
            run_type = run.get("type")
            if run_type not in {"text", "sym"}:
                continue

            run_style = resolver.resolve_run_style_r_format(run.get("r_style_id"), run.get("char_style_id"))
            effective = deep_merge_r_format(paragraph_base, run_style)
            effective = deep_merge_r_format(effective, run.get("r_format") or {})
            signature = build_signature(effective)

            agg = usage_map.setdefault(signature, UsageAggregate())
            agg.runs_count += 1
            text = run.get("text") or ""
            agg.chars_count += len(text)
            agg.paragraph_ids.add(p_id)
            agg.table_ids.update(table_stack)
            if agg.first_seen_path is None:
                agg.first_seen_path = run.get("id") or f"{p_id}.<run>"

            font_counter[signature.normalized_title] += 1
            font_usage_counter[signature] += 1

    return usage_map, font_counter, font_usage_counter


def connect_db(db_path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def create_scan_run(conn: sqlite3.Connection, dry_run: bool) -> Optional[int]:
    if dry_run:
        return None
    cur = conn.execute(
        "INSERT INTO font_scan_run(status, started_at) VALUES (?, ?)",
        ("running", now_utc()),
    )
    return int(cur.lastrowid)


def finish_scan_run(
    conn: sqlite3.Connection,
    scan_run_id: Optional[int],
    *,
    status: str,
    scanned_documents: int,
    inserted_fonts: int,
    inserted_font_usages: int,
    inserted_links: int,
    error_message: Optional[str] = None,
    dry_run: bool = False,
) -> None:
    if dry_run or scan_run_id is None:
        return
    conn.execute(
        """
        UPDATE font_scan_run
           SET finished_at = ?,
               status = ?,
               scanned_documents = ?,
               inserted_fonts = ?,
               inserted_font_usages = ?,
               inserted_links = ?,
               error_message = ?
         WHERE id = ?
        """,
        (
            now_utc(),
            status,
            scanned_documents,
            inserted_fonts,
            inserted_font_usages,
            inserted_links,
            error_message,
            scan_run_id,
        ),
    )
    conn.commit()


def get_documents(conn: sqlite3.Connection, args: argparse.Namespace) -> List[sqlite3.Row]:
    sql = [
        "SELECT id, uid, artifacts_abs_path, source_abs_path, processing_status",
        "FROM document",
        "WHERE 1=1",
    ]
    params: List[Any] = []

    if args.uid:
        placeholders = ",".join("?" for _ in args.uid)
        sql.append(f"AND uid IN ({placeholders})")
        params.extend(args.uid)

    statuses = args.status or ["parsed_raw"]
    if statuses:
        placeholders = ",".join("?" for _ in statuses)
        sql.append(f"AND processing_status IN ({placeholders})")
        params.extend(statuses)

    if args.where:
        sql.append(f"AND ({args.where})")

    sql.append("ORDER BY id")
    if args.limit:
        sql.append("LIMIT ?")
        params.append(args.limit)

    return list(conn.execute("\n".join(sql), params))


def resolve_raw_path(row: sqlite3.Row, args: argparse.Namespace) -> str:
    artifacts_abs_path = row["artifacts_abs_path"]
    uid = row["uid"]

    if artifacts_abs_path:
        return os.path.join(artifacts_abs_path, args.raw_filename)
    if args.artifacts_root:
        return os.path.join(args.artifacts_root, uid, args.raw_filename)
    raise ScanError(f"Cannot resolve artifacts path for uid={uid}")


def ensure_font(conn: sqlite3.Connection, signature: FontUsageSignature, counters: Dict[str, int], dry_run: bool) -> int:
    row = conn.execute(
        "SELECT id FROM font WHERE normalized_title = ?",
        (signature.normalized_title,),
    ).fetchone()
    if row:
        return int(row["id"])
    if dry_run:
        counters["inserted_fonts"] += 1
        return -1
    cur = conn.execute(
        "INSERT INTO font(title, normalized_title, created_at, updated_at) VALUES (?, ?, ?, ?)",
        (signature.font_title, signature.normalized_title, now_utc(), now_utc()),
    )
    counters["inserted_fonts"] += 1
    return int(cur.lastrowid)


def ensure_font_usage(
    conn: sqlite3.Connection,
    font_id: int,
    signature: FontUsageSignature,
    counters: Dict[str, int],
    dry_run: bool,
) -> int:
    query = """
        SELECT id
          FROM font_usage
         WHERE font_id IS ?
           AND size_pt IS ?
           AND bold IS ?
           AND italic IS ?
           AND underline IS ?
           AND all_caps IS ?
           AND small_caps IS ?
           AND strike IS ?
           AND double_strike IS ?
           AND outline IS ?
           AND shadow IS ?
           AND emboss IS ?
           AND imprint IS ?
           AND rtl IS ?
           AND lang IS ?
         LIMIT 1
    """
    params = (
        font_id,
        signature.size_pt,
        signature.bold,
        signature.italic,
        signature.underline,
        signature.all_caps,
        signature.small_caps,
        signature.strike,
        signature.double_strike,
        signature.outline,
        signature.shadow,
        signature.emboss,
        signature.imprint,
        signature.rtl,
        signature.lang,
    )
    row = conn.execute(query, params).fetchone()
    if row:
        return int(row["id"])
    if dry_run:
        counters["inserted_font_usages"] += 1
        return -1
    cur = conn.execute(
        """
        INSERT INTO font_usage(
            font_id, size_pt, bold, italic, underline, all_caps, small_caps,
            strike, double_strike, outline, shadow, emboss, imprint, rtl, lang,
            created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        params + (now_utc(), now_utc()),
    )
    counters["inserted_font_usages"] += 1
    return int(cur.lastrowid)


def upsert_document_font_usage(
    conn: sqlite3.Connection,
    document_id: int,
    font_usage_id: int,
    aggregate: UsageAggregate,
    scan_run_id: Optional[int],
    counters: Dict[str, int],
    dry_run: bool,
) -> None:
    if dry_run:
        counters["inserted_links"] += 1
        return

    existing = conn.execute(
        "SELECT id FROM document_font_usage WHERE document_id = ? AND font_usage_id = ?",
        (document_id, font_usage_id),
    ).fetchone()

    payload = (
        aggregate.runs_count,
        aggregate.chars_count,
        len(aggregate.paragraph_ids),
        len(aggregate.table_ids),
        aggregate.first_seen_path,
        now_utc(),
        scan_run_id,
        document_id,
        font_usage_id,
    )

    if existing:
        conn.execute(
            """
            UPDATE document_font_usage
               SET runs_count = ?,
                   chars_count = ?,
                   paragraphs_count = ?,
                   tables_count = ?,
                   first_seen_path = ?,
                   updated_at = ?,
                   scan_run_id = ?
             WHERE document_id = ? AND font_usage_id = ?
            """,
            payload,
        )
    else:
        conn.execute(
            """
            INSERT INTO document_font_usage(
                runs_count, chars_count, paragraphs_count, tables_count,
                first_seen_path, updated_at, scan_run_id, document_id, font_usage_id,
                created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            payload + (now_utc(),),
        )
    counters["inserted_links"] += 1


def upsert_document_font_profile(
    conn: sqlite3.Connection,
    document_id: int,
    aggregate_by_font_usage_id: Dict[int, UsageAggregate],
    font_ids_by_usage_id: Dict[int, int],
    dry_run: bool,
) -> None:
    if dry_run:
        return

    total_runs = sum(item.runs_count for item in aggregate_by_font_usage_id.values())
    total_chars = sum(item.chars_count for item in aggregate_by_font_usage_id.values())
    unique_font_usages_count = len(aggregate_by_font_usage_id)
    unique_fonts_count = len(set(font_ids_by_usage_id.values()))

    primary_font_usage_id: Optional[int] = None
    if aggregate_by_font_usage_id:
        primary_font_usage_id = max(
            aggregate_by_font_usage_id.items(),
            key=lambda pair: (pair[1].chars_count, pair[1].runs_count, -pair[0]),
        )[0]

    existing = conn.execute(
        "SELECT document_id FROM document_font_profile WHERE document_id = ?",
        (document_id,),
    ).fetchone()

    if existing:
        conn.execute(
            """
            UPDATE document_font_profile
               SET primary_font_usage_id = ?,
                   total_runs_count = ?,
                   total_chars_count = ?,
                   unique_font_usages_count = ?,
                   unique_fonts_count = ?,
                   updated_at = ?
             WHERE document_id = ?
            """,
            (
                primary_font_usage_id,
                total_runs,
                total_chars,
                unique_font_usages_count,
                unique_fonts_count,
                now_utc(),
                document_id,
            ),
        )
    else:
        conn.execute(
            """
            INSERT INTO document_font_profile(
                document_id, primary_font_usage_id, total_runs_count,
                total_chars_count, unique_font_usages_count, unique_fonts_count,
                created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                document_id,
                primary_font_usage_id,
                total_runs,
                total_chars,
                unique_font_usages_count,
                unique_fonts_count,
                now_utc(),
                now_utc(),
            ),
        )


def reset_document_usage(conn: sqlite3.Connection, document_id: int, dry_run: bool) -> None:
    if dry_run:
        return
    conn.execute("DELETE FROM document_font_usage WHERE document_id = ?", (document_id,))
    conn.execute("DELETE FROM document_font_profile WHERE document_id = ?", (document_id,))


def scan_document(
    conn: sqlite3.Connection,
    row: sqlite3.Row,
    args: argparse.Namespace,
    counters: Dict[str, int],
    scan_run_id: Optional[int],
) -> None:
    raw_path = resolve_raw_path(row, args)
    if not os.path.exists(raw_path):
        raise ScanError(f"raw.json not found: {raw_path}")

    with open(raw_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    usage_map, _, _ = aggregate_document(data)
    document_id = int(row["id"])

    if args.reset_document_data:
        reset_document_usage(conn, document_id, args.dry_run)

    aggregate_by_font_usage_id: Dict[int, UsageAggregate] = {}
    font_ids_by_usage_id: Dict[int, int] = {}

    for signature, aggregate in usage_map.items():
        font_id = ensure_font(conn, signature, counters, args.dry_run)
        font_usage_id = ensure_font_usage(conn, font_id, signature, counters, args.dry_run)
        upsert_document_font_usage(
            conn,
            document_id=document_id,
            font_usage_id=font_usage_id,
            aggregate=aggregate,
            scan_run_id=scan_run_id,
            counters=counters,
            dry_run=args.dry_run,
        )
        aggregate_by_font_usage_id[font_usage_id] = aggregate
        font_ids_by_usage_id[font_usage_id] = font_id

    upsert_document_font_profile(
        conn,
        document_id=document_id,
        aggregate_by_font_usage_id=aggregate_by_font_usage_id,
        font_ids_by_usage_id=font_ids_by_usage_id,
        dry_run=args.dry_run,
    )


def main() -> int:
    args = parse_args()
    conn = connect_db(args.db)
    counters: Dict[str, int] = defaultdict(int)
    scan_run_id: Optional[int] = None
    scanned_documents = 0

    try:
        documents = get_documents(conn, args)
        if not documents:
            print("[scan_font_usage] no documents matched", file=sys.stderr)
            return 0

        scan_run_id = create_scan_run(conn, args.dry_run)

        for index, row in enumerate(documents, start=1):
            uid = row["uid"]
            try:
                scan_document(conn, row, args, counters, scan_run_id)
                scanned_documents += 1
                if args.verbose:
                    print(f"[scan_font_usage] ok uid={uid}")
            except Exception as exc:
                print(f"[scan_font_usage] ERROR uid={uid}: {exc}", file=sys.stderr)
                raise

            if not args.dry_run and (index % max(1, args.commit_every) == 0):
                conn.commit()

        if not args.dry_run:
            conn.commit()

        finish_scan_run(
            conn,
            scan_run_id,
            status="done",
            scanned_documents=scanned_documents,
            inserted_fonts=counters["inserted_fonts"],
            inserted_font_usages=counters["inserted_font_usages"],
            inserted_links=counters["inserted_links"],
            dry_run=args.dry_run,
        )

        print(
            "[scan_font_usage] done "
            f"scanned_documents={scanned_documents} "
            f"inserted_fonts={counters['inserted_fonts']} "
            f"inserted_font_usages={counters['inserted_font_usages']} "
            f"inserted_links={counters['inserted_links']}"
        )
        return 0

    except Exception as exc:
        if not args.dry_run:
            conn.rollback()
        finish_scan_run(
            conn,
            scan_run_id,
            status="error",
            scanned_documents=scanned_documents,
            inserted_fonts=counters["inserted_fonts"],
            inserted_font_usages=counters["inserted_font_usages"],
            inserted_links=counters["inserted_links"],
            error_message=str(exc),
            dry_run=args.dry_run,
        )
        print(f"[scan_font_usage] failed: {exc}", file=sys.stderr)
        return 1
    finally:
        conn.close()


if __name__ == "__main__":
    raise SystemExit(main())
