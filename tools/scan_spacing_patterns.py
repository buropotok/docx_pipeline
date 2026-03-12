#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import csv
import json
import os
import re
import sqlite3
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, asdict
from typing import Any, Dict, Iterator, List, Optional, Tuple


WHITESPACE_GROUP_RE = re.compile(r"(\t+| {2,})")
DEFAULT_RAW_FILENAME = "raw.json"


@dataclass
class PatternRow:
    document_id: int
    uid: str
    paragraph_id: str
    paragraph_path: str
    in_table: int
    table_path: str
    run_id: str
    run_index: int
    group_kind: str
    group_text: str
    group_len: int
    text_left: str
    text_right: str
    paragraph_alignment: Optional[str]
    preserve: Optional[int]
    font_name: Optional[str]
    font_size_pt: Optional[float]


class ScanError(Exception):
    pass


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Scan raw.json files and export spacing/tab patterns to CSV/JSON."
    )
    parser.add_argument("--db", required=True, help="Path to SQLite database.")
    parser.add_argument("--out-csv", required=True, help="CSV output path.")
    parser.add_argument("--out-json", default=None, help="Optional JSON summary output path.")
    parser.add_argument("--artifacts-root", default=None, help="Fallback artifacts root.")
    parser.add_argument("--raw-filename", default=DEFAULT_RAW_FILENAME)
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--uid", action="append", default=[])
    parser.add_argument("--status", action="append", default=[])
    parser.add_argument("--min-spaces", type=int, default=3, help="Minimum spaces to record. Default: 3")
    parser.add_argument("--verbose", action="store_true")
    return parser.parse_args()


def connect_db(db_path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def get_documents(conn: sqlite3.Connection, args: argparse.Namespace) -> List[sqlite3.Row]:
    sql = [
        "SELECT id, uid, artifacts_abs_path, processing_status",
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

    sql.append("ORDER BY id")
    if args.limit:
        sql.append("LIMIT ?")
        params.append(args.limit)

    return list(conn.execute("\n".join(sql), params))


def resolve_raw_path(row: sqlite3.Row, args: argparse.Namespace) -> str:
    artifacts_abs_path = row["artifacts_abs_path"]
    if artifacts_abs_path:
        return os.path.join(artifacts_abs_path, args.raw_filename)
    if args.artifacts_root:
        return os.path.join(args.artifacts_root, row["uid"], args.raw_filename)
    raise ScanError(f"Cannot resolve artifacts path for uid={row['uid']}")


def pick_font_name(rfonts: Optional[Dict[str, Any]]) -> Optional[str]:
    if not isinstance(rfonts, dict):
        return None
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        value = rfonts.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()
    return None


def half_points_to_pt(value: Any) -> Optional[float]:
    if value is None:
        return None
    try:
        return float(value) / 2.0
    except (TypeError, ValueError):
        return None


def iter_paragraphs(content: List[Dict[str, Any]], path: Optional[List[str]] = None) -> Iterator[Tuple[Dict[str, Any], List[str], List[str]]]:
    path = list(path or [])
    for item in content or []:
        item_type = item.get("type")
        if item_type == "paragraph":
            yield item, path, []
        elif item_type == "table":
            table_id = item.get("id") or "<table>"
            table_path = path + [table_id]
            for row in item.get("rows", []) or []:
                row_id = row.get("id") or "<row>"
                for cell in row.get("cells", []) or []:
                    cell_id = cell.get("id") or "<cell>"
                    yield from iter_paragraphs(cell.get("content", []) or [], table_path + [row_id, cell_id])


def scan_paragraph(
    *,
    document_id: int,
    uid: str,
    paragraph: Dict[str, Any],
    path: List[str],
    min_spaces: int,
) -> List[PatternRow]:
    rows: List[PatternRow] = []
    runs = paragraph.get("runs", []) or []
    paragraph_id = paragraph.get("id") or "<paragraph>"
    paragraph_path = "/".join(path + [paragraph_id]) if path else paragraph_id
    in_table = 1 if path else 0
    table_path = "/".join(path) if path else ""
    paragraph_alignment = (paragraph.get("p_format") or {}).get("alignment")

    for run_index, run in enumerate(runs):
        run_type = run.get("type")
        if run_type == "tab":
            rows.append(
                PatternRow(
                    document_id=document_id,
                    uid=uid,
                    paragraph_id=paragraph_id,
                    paragraph_path=paragraph_path,
                    in_table=in_table,
                    table_path=table_path,
                    run_id=run.get("id") or f"{paragraph_id}.tab_{run_index}",
                    run_index=run_index,
                    group_kind="tab_run",
                    group_text="\\t",
                    group_len=1,
                    text_left=extract_neighbor_text(runs, run_index, direction=-1),
                    text_right=extract_neighbor_text(runs, run_index, direction=1),
                    paragraph_alignment=paragraph_alignment,
                    preserve=int(bool((run.get("meta") or {}).get("preserve"))) if run.get("meta") else None,
                    font_name=None,
                    font_size_pt=None,
                )
            )
            continue

        if run_type != "text":
            continue

        text = run.get("text") or ""
        for match in WHITESPACE_GROUP_RE.finditer(text):
            token = match.group(0)
            if token.startswith(" ") and len(token) < min_spaces:
                continue
            kind = "tabs" if token.startswith("\t") else "spaces"
            left = text[: match.start()][-20:]
            right = text[match.end() :][:20]
            r_format = run.get("r_format") or {}
            rows.append(
                PatternRow(
                    document_id=document_id,
                    uid=uid,
                    paragraph_id=paragraph_id,
                    paragraph_path=paragraph_path,
                    in_table=in_table,
                    table_path=table_path,
                    run_id=run.get("id") or f"{paragraph_id}.run_{run_index}",
                    run_index=run_index,
                    group_kind=kind,
                    group_text=token.replace("\t", "\\t"),
                    group_len=len(token),
                    text_left=left,
                    text_right=right,
                    paragraph_alignment=paragraph_alignment,
                    preserve=int(bool((run.get("meta") or {}).get("preserve"))) if run.get("meta") else None,
                    font_name=pick_font_name(r_format.get("rFonts")),
                    font_size_pt=half_points_to_pt(r_format.get("font_size_half_points")),
                )
            )
    return rows


def extract_neighbor_text(runs: List[Dict[str, Any]], idx: int, direction: int) -> str:
    j = idx + direction
    while 0 <= j < len(runs):
        run = runs[j]
        if run.get("type") == "text":
            text = run.get("text") or ""
            if direction < 0:
                return text[-20:]
            return text[:20]
        j += direction
    return ""


def ensure_parent_dir(path: str) -> None:
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)


def main() -> int:
    args = parse_args()
    conn = connect_db(args.db)

    try:
        docs = get_documents(conn, args)
        ensure_parent_dir(args.out_csv)
        if args.out_json:
            ensure_parent_dir(args.out_json)

        all_rows: List[PatternRow] = []
        summary = {
            "documents_scanned": 0,
            "patterns_total": 0,
            "by_kind": Counter(),
            "by_length": Counter(),
            "in_table": Counter(),
            "alignment": Counter(),
        }

        for row in docs:
            raw_path = resolve_raw_path(row, args)
            if not os.path.exists(raw_path):
                raise ScanError(f"raw.json not found: {raw_path}")

            with open(raw_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            doc_rows: List[PatternRow] = []
            for paragraph, path, _ in iter_paragraphs(data.get("content", []) or []):
                doc_rows.extend(
                    scan_paragraph(
                        document_id=int(row["id"]),
                        uid=row["uid"],
                        paragraph=paragraph,
                        path=path,
                        min_spaces=args.min_spaces,
                    )
                )

            all_rows.extend(doc_rows)
            summary["documents_scanned"] += 1
            for item in doc_rows:
                summary["patterns_total"] += 1
                summary["by_kind"][item.group_kind] += 1
                summary["by_length"][str(item.group_len)] += 1
                summary["in_table"][str(item.in_table)] += 1
                summary["alignment"][item.paragraph_alignment or "<none>"] += 1

            if args.verbose:
                print(f"[scan_spacing_patterns] ok uid={row['uid']} patterns={len(doc_rows)}")

        with open(args.out_csv, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=list(asdict(all_rows[0]).keys()) if all_rows else list(PatternRow.__annotations__.keys()))
            writer.writeheader()
            for item in all_rows:
                writer.writerow(asdict(item))

        if args.out_json:
            serializable_summary = {
                "documents_scanned": summary["documents_scanned"],
                "patterns_total": summary["patterns_total"],
                "by_kind": dict(summary["by_kind"]),
                "by_length": dict(summary["by_length"]),
                "in_table": dict(summary["in_table"]),
                "alignment": dict(summary["alignment"]),
            }
            with open(args.out_json, "w", encoding="utf-8") as f:
                json.dump(serializable_summary, f, ensure_ascii=False, indent=2)

        print(
            f"[scan_spacing_patterns] done documents_scanned={summary['documents_scanned']} patterns_total={summary['patterns_total']}"
        )
        return 0

    except Exception as exc:
        print(f"[scan_spacing_patterns] failed: {exc}", file=sys.stderr)
        return 1
    finally:
        conn.close()


if __name__ == "__main__":
    raise SystemExit(main())
