#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import copy
import json
import logging
import os
import re
import sqlite3
import sys
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, Iterable, List, Optional, Tuple

DEFAULT_STATUS = "parsed_raw"
DEFAULT_OUTPUT_STATUS = "optimized"
DEFAULT_INPUT_NAME = "raw.json"
DEFAULT_OUTPUT_NAME = "raw_optimized.json"
DEFAULT_MIN_SPACE_GROUP = 4
DEFAULT_FONT = "Times New Roman"
DEFAULT_SIZE_HALF_POINTS = 24
DEFAULT_CHAR_WIDTH = 120.0
DEFAULT_SPACE_WIDTH = 60.0
TAB_EXISTS_EPSILON = 5

SPACE_GROUP_RE = re.compile(r" {4,}")
PROSE_PUNCT_RE = re.compile(r"[\.,;:!?]")
MULTISPACE_RE = re.compile(r" {2,}")


@dataclass
class ParagraphContext:
    in_table: bool = False
    table_path: Optional[str] = None


@dataclass
class ParagraphStats:
    spaces_found: int = 0
    spaces_converted: int = 0
    trailing_trimmed: int = 0
    table_candidates_skipped: int = 0
    preserve_skipped: int = 0
    prose_skipped: int = 0
    paragraphs_changed: int = 0

    def merge(self, other: "ParagraphStats") -> None:
        self.spaces_found += other.spaces_found
        self.spaces_converted += other.spaces_converted
        self.trailing_trimmed += other.trailing_trimmed
        self.table_candidates_skipped += other.table_candidates_skipped
        self.preserve_skipped += other.preserve_skipped
        self.prose_skipped += other.prose_skipped
        self.paragraphs_changed += other.paragraphs_changed


class FontMetrics:
    def __init__(self) -> None:
        self.metrics: Dict[str, Dict[float, Dict[str, float]]] = {}
        self.aliases = {
            "unknown": "times new roman",
            "times": "times new roman",
        }

    @staticmethod
    def normalize_font_name(name: Optional[str]) -> str:
        if not name:
            return ""
        return " ".join(str(name).strip().lower().split())

    def load_from_db(self, conn: sqlite3.Connection) -> None:
        rows = conn.execute(
            """
            SELECT
                f.normalized_title,
                fm.base_size_pt,
                fm.char_text,
                fm.char_code,
                fm.width_units
            FROM font_metric fm
            JOIN font f ON f.id = fm.font_id
            """
        ).fetchall()
        for row in rows:
            font_name = self.normalize_font_name(row[0])
            base_size = float(row[1])
            char_text = row[2]
            char_code = row[3]
            width = float(row[4])
            if char_text is None or char_text == "":
                try:
                    char_text = chr(int(char_code))
                except Exception:
                    continue
            self.metrics.setdefault(font_name, {}).setdefault(base_size, {})[char_text] = width

    def resolve_font_key(self, font_name: Optional[str]) -> str:
        normalized = self.normalize_font_name(font_name)
        if normalized in self.metrics:
            return normalized
        alias = self.aliases.get(normalized)
        if alias and alias in self.metrics:
            return alias
        default_key = self.normalize_font_name(DEFAULT_FONT)
        if default_key in self.metrics:
            return default_key
        return normalized

    def get_char_width(self, char: str, font_name: Optional[str], font_size_half_points: Optional[int]) -> float:
        size_half = font_size_half_points or DEFAULT_SIZE_HALF_POINTS
        size_pt = size_half / 2.0
        font_key = self.resolve_font_key(font_name)
        base_map = self.metrics.get(font_key)
        if not base_map:
            return DEFAULT_SPACE_WIDTH if char == " " else DEFAULT_CHAR_WIDTH
        if 12.0 in base_map:
            base_size = 12.0
        else:
            base_size = sorted(base_map.keys())[0]
        char_map = base_map[base_size]
        if char in char_map:
            base_width = char_map[char]
        elif char == " ":
            base_width = char_map.get(" ", DEFAULT_SPACE_WIDTH)
        else:
            base_width = DEFAULT_CHAR_WIDTH
        return base_width * (size_pt / base_size)

    def text_width(self, text: str, font_name: Optional[str], font_size_half_points: Optional[int]) -> float:
        return sum(self.get_char_width(ch, font_name, font_size_half_points) for ch in text)


def str2bool(value: str) -> bool:
    if isinstance(value, bool):
        return value
    normalized = str(value).strip().lower()
    if normalized in {"1", "true", "yes", "y", "on"}:
        return True
    if normalized in {"0", "false", "no", "n", "off"}:
        return False
    raise argparse.ArgumentTypeError(f"Invalid boolean value: {value}")


def setup_logger() -> Tuple[logging.Logger, str]:
    tools_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    log_dir = os.path.join(tools_dir, "log")
    os.makedirs(log_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = os.path.join(log_dir, f"optimize_spaces_{timestamp}.log")

    logger = logging.getLogger(f"optimize_spaces_{timestamp}")
    logger.setLevel(logging.INFO)
    logger.propagate = False

    formatter = logging.Formatter("%(asctime)s | %(message)s", datefmt="%Y-%m-%d %H:%M:%S")

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    stream_handler = logging.StreamHandler(sys.stderr)
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)
    return logger, log_path


def connect_db(db_path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def fetch_documents(
    conn: sqlite3.Connection,
    status: str,
    limit: Optional[int],
    uid: Optional[str],
) -> List[sqlite3.Row]:
    params: List[Any] = []
    sql = """
        SELECT id, uid, artifacts_abs_path, processing_status
        FROM document
        WHERE artifacts_abs_path IS NOT NULL
    """
    if uid:
        sql += " AND uid = ?"
        params.append(uid)
    else:
        sql += " AND processing_status = ?"
        params.append(status)
    sql += " ORDER BY id"
    if limit is not None:
        sql += " LIMIT ?"
        params.append(limit)
    return list(conn.execute(sql, params).fetchall())


def merge_format(base: Dict[str, Any], extra: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    result = dict(base)
    if extra:
        result.update(extra)
    return result


def resolve_style_format(
    styles: Dict[str, Any],
    style_id: Optional[str],
    key: str,
    cache: Dict[Tuple[Optional[str], str], Dict[str, Any]],
) -> Dict[str, Any]:
    cache_key = (style_id, key)
    if cache_key in cache:
        return cache[cache_key]
    if not style_id or style_id not in styles:
        cache[cache_key] = {}
        return {}
    style = styles[style_id]
    parent_id = style.get("based_on")
    result = resolve_style_format(styles, parent_id, key, cache)
    result = merge_format(result, style.get(key))
    cache[cache_key] = result
    return result


def get_effective_paragraph_indent(
    para: Dict[str, Any],
    styles: Dict[str, Any],
    doc_defaults: Dict[str, Any],
    style_cache: Dict[Tuple[Optional[str], str], Dict[str, Any]],
) -> int:
    p_format = dict(doc_defaults.get("p_format") or {})
    p_style_id = para.get("p_style_id") or para.get("style_id")
    p_format = merge_format(p_format, resolve_style_format(styles, p_style_id, "p_format", style_cache))
    p_format = merge_format(p_format, para.get("p_format"))
    return int(p_format.get("indent_start_twip") or 0)


def get_effective_run_format(
    run: Dict[str, Any],
    para: Dict[str, Any],
    styles: Dict[str, Any],
    doc_defaults: Dict[str, Any],
    style_cache: Dict[Tuple[Optional[str], str], Dict[str, Any]],
) -> Dict[str, Any]:
    result = dict(doc_defaults.get("r_format") or {})
    p_style_id = para.get("p_style_id") or para.get("style_id")
    result = merge_format(result, resolve_style_format(styles, p_style_id, "r_format", style_cache))
    char_style_id = run.get("char_style_id") or run.get("r_style_id")
    result = merge_format(result, resolve_style_format(styles, char_style_id, "r_format", style_cache))
    result = merge_format(result, run.get("r_format"))
    return result


def extract_font_and_size(effective_r_format: Dict[str, Any]) -> Tuple[str, int]:
    r_fonts = effective_r_format.get("rFonts") or {}
    font_name = (
        r_fonts.get("ascii")
        or r_fonts.get("hAnsi")
        or r_fonts.get("cs")
        or r_fonts.get("eastAsia")
        or DEFAULT_FONT
    )
    size_half = int(effective_r_format.get("font_size_half_points") or DEFAULT_SIZE_HALF_POINTS)
    return font_name, size_half


def is_prose_like(para: Dict[str, Any]) -> bool:
    """
    Conservative prose detector.

    Important:
    - alignment='justify' by itself is NOT enough to treat a paragraph as prose.
      In forms, many label/value lines are justified too.
    - We only skip as prose when there are strong textual signals of running text.
    """
    text_parts: List[str] = []
    for run in para.get("runs", []):
        if run.get("type") == "text":
            text_parts.append(run.get("text", ""))
    full_text = "".join(text_parts)

    visible_text = full_text.strip()
    if not visible_text:
        return False

    words = [w for w in re.split(r"\s+", visible_text) if w]
    has_punct = PROSE_PUNCT_RE.search(visible_text) is not None
    has_multi_space = MULTISPACE_RE.search(full_text) is not None
    text_len = len(visible_text)

    # Short form-like lines are not prose even if justified.
    if text_len <= 80 and len(words) <= 8:
        return False

    # Long sentences/paragraphs with punctuation are prose.
    if text_len >= 120:
        return True
    if text_len >= 80 and len(words) >= 10 and has_punct:
        return True
    if len(words) >= 14 and has_punct:
        return True

    # Multiple explicit spacing groups usually indicate layout, not prose.
    if has_multi_space:
        return False

    return False


def clone_run(run: Dict[str, Any], text: str, suffix: str) -> Dict[str, Any]:
    cloned = copy.deepcopy(run)
    cloned["text"] = text
    if run.get("id"):
        cloned["id"] = f"{run['id']}{suffix}"
    return cloned


def build_tab_run(source_run: Dict[str, Any], leading: bool) -> Dict[str, Any]:
    tab_run: Dict[str, Any] = {"type": "tab"}
    if source_run.get("parent_id"):
        tab_run["parent_id"] = source_run["parent_id"]
    if source_run.get("id"):
        tab_run["id"] = f"{source_run['id']}.tab"
    if leading:
        tab_run["meta"] = {"leading": True}
    return tab_run


def ensure_tab_stop(p_format: Dict[str, Any], target_pos: float) -> bool:
    tabs = p_format.setdefault("tabs", [])
    rounded_target = int(round(target_pos))
    exists = any(abs(int(tab.get("posTwip", 0)) - rounded_target) < TAB_EXISTS_EPSILON for tab in tabs)
    if exists:
        return False
    tabs.append({"posTwip": rounded_target, "val": "left"})
    tabs.sort(key=lambda item: int(item.get("posTwip", 0)))
    return True


def trim_trailing_spaces(runs: List[Dict[str, Any]]) -> int:
    trimmed = 0
    for idx in range(len(runs) - 1, -1, -1):
        run = runs[idx]
        run_type = run.get("type")
        if run_type != "text":
            if run_type in {"tab", "break"}:
                break
            continue
        text = run.get("text", "")
        if text == "":
            continue
        new_text = text.rstrip(" ")
        removed = len(text) - len(new_text)
        if removed > 0:
            run["text"] = new_text
            trimmed += removed
        if new_text:
            break
    return trimmed


def is_visible_text_run(run: Optional[Dict[str, Any]]) -> bool:
    if not run or run.get("type") != "text":
        return False
    return (run.get("text") or "").strip(" ") != ""


def find_prev_visible_text_run(runs: List[Dict[str, Any]], start_index: int) -> Optional[Dict[str, Any]]:
    for idx in range(start_index - 1, -1, -1):
        run = runs[idx]
        if run.get("type") == "text" and is_visible_text_run(run):
            return run
        if run.get("type") in {"tab", "break", "picture"}:
            break
    return None


def find_next_visible_text_run(runs: List[Dict[str, Any]], start_index: int) -> Optional[Dict[str, Any]]:
    for idx in range(start_index + 1, len(runs)):
        run = runs[idx]
        if run.get("type") == "text" and is_visible_text_run(run):
            return run
        if run.get("type") in {"tab", "break", "picture"}:
            break
    return None


def optimize_paragraph_spaces(
    para: Dict[str, Any],
    ctx: ParagraphContext,
    font_metrics: FontMetrics,
    styles: Dict[str, Any],
    doc_defaults: Dict[str, Any],
    style_cache: Dict[Tuple[Optional[str], str], Dict[str, Any]],
    default_tab_stop: int,
    min_space_group: int,
) -> ParagraphStats:
    stats = ParagraphStats()
    runs = para.get("runs", [])
    if not runs:
        return stats

    stats.trailing_trimmed = trim_trailing_spaces(runs)
    current_pos = float(get_effective_paragraph_indent(para, styles, doc_defaults, style_cache))
    indent_pos = current_pos
    p_format = copy.deepcopy(para.get("p_format") or {})
    prose_like = is_prose_like(para)
    changed = stats.trailing_trimmed > 0

    new_runs: List[Dict[str, Any]] = []
    for run_index, run in enumerate(runs):
        run_type = run.get("type")
        if run_type == "tab":
            new_runs.append(run)
            current_pos += default_tab_stop
            continue
        if run_type != "text":
            new_runs.append(run)
            continue

        text = run.get("text", "")
        if not text:
            new_runs.append(run)
            continue

        effective_r = get_effective_run_format(run, para, styles, doc_defaults, style_cache)
        font_name, size_half = extract_font_and_size(effective_r)

        # Separate case: a standalone run consisting only of spaces between/after visible runs.
        if text and text.strip(" ") == "" and len(text) >= min_space_group:
            stats.spaces_found += 1
            prev_visible = find_prev_visible_text_run(runs, run_index)
            next_visible = find_next_visible_text_run(runs, run_index)
            leading = prev_visible is None and current_pos <= indent_pos + 1

            if ctx.in_table:
                stats.table_candidates_skipped += 1
                current_pos += font_metrics.text_width(text, font_name, size_half)
                new_runs.append(run)
                continue
            if prose_like and prev_visible is not None and next_visible is not None:
                stats.prose_skipped += 1
                current_pos += font_metrics.text_width(text, font_name, size_half)
                new_runs.append(run)
                continue

            if leading or prev_visible is not None or next_visible is not None:
                target_pos = current_pos + font_metrics.text_width(text, font_name, size_half)
                ensure_tab_stop(p_format, target_pos)
                new_runs.append(build_tab_run(run, leading=leading))
                current_pos = target_pos
                stats.spaces_converted += 1
                changed = True
                continue

            current_pos += font_metrics.text_width(text, font_name, size_half)
            new_runs.append(run)
            continue

        matches = [m for m in SPACE_GROUP_RE.finditer(text) if len(m.group(0)) >= min_space_group]
        if not matches:
            current_pos += font_metrics.text_width(text, font_name, size_half)
            new_runs.append(run)
            continue

        stats.spaces_found += len(matches)
        if ctx.in_table:
            stats.table_candidates_skipped += len(matches)
            current_pos += font_metrics.text_width(text, font_name, size_half)
            new_runs.append(run)
            continue
        if prose_like:
            stats.prose_skipped += len(matches)
            current_pos += font_metrics.text_width(text, font_name, size_half)
            new_runs.append(run)
            continue

        cursor = 0
        part_counter = 1
        run_changed = False
        for match in matches:
            start, end = match.span()
            before = text[cursor:start]
            spaces = text[start:end]
            after_exists = end < len(text) and text[end:].strip(" ") != ""
            leading = (cursor == 0 and start == 0 and current_pos <= indent_pos + 1)

            if before:
                new_runs.append(clone_run(run, before, f".s{part_counter}"))
                current_pos += font_metrics.text_width(before, font_name, size_half)
                part_counter += 1

            should_convert = leading or after_exists

            if should_convert:
                target_pos = current_pos + font_metrics.text_width(spaces, font_name, size_half)
                ensure_tab_stop(p_format, target_pos)
                new_runs.append(build_tab_run(run, leading=leading))
                current_pos = target_pos
                stats.spaces_converted += 1
                changed = True
                run_changed = True
            else:
                new_runs.append(clone_run(run, spaces, f".s{part_counter}"))
                current_pos += font_metrics.text_width(spaces, font_name, size_half)
                part_counter += 1

            cursor = end

        tail = text[cursor:]
        if tail or not run_changed:
            tail_suffix = f".s{part_counter}"
            if not run_changed and cursor == 0:
                new_runs.append(run)
            elif tail:
                new_runs.append(clone_run(run, tail, tail_suffix))
            if tail:
                current_pos += font_metrics.text_width(tail, font_name, size_half)

    if changed:
        para["runs"] = new_runs
        if p_format:
            para["p_format"] = p_format
        stats.paragraphs_changed += 1
    return stats



def walk_content_items(content: Iterable[Dict[str, Any]], ctx: ParagraphContext) -> Iterable[Tuple[Dict[str, Any], ParagraphContext]]:
    for item in content:
        item_type = item.get("type")
        if item_type == "paragraph":
            yield item, ctx
        elif item_type == "table":
            table_ctx = ParagraphContext(in_table=True, table_path=item.get("id") or ctx.table_path)
            for row in item.get("rows", []):
                for cell in row.get("cells", []):
                    cell_content = cell.get("content", [])
                    yield from walk_content_items(cell_content, table_ctx)


def process_document(
    data: Dict[str, Any],
    font_metrics: FontMetrics,
    min_space_group: int,
) -> ParagraphStats:
    styles = data.get("styles") or {}
    doc_defaults = data.get("doc_defaults") or {}
    style_cache: Dict[Tuple[Optional[str], str], Dict[str, Any]] = {}
    default_tab_stop = int((data.get("document_info") or {}).get("settings", {}).get("defaultTabStopTwip") or 708)

    total = ParagraphStats()
    for para, ctx in walk_content_items(data.get("content", []), ParagraphContext()):
        para_stats = optimize_paragraph_spaces(
            para=para,
            ctx=ctx,
            font_metrics=font_metrics,
            styles=styles,
            doc_defaults=doc_defaults,
            style_cache=style_cache,
            default_tab_stop=default_tab_stop,
            min_space_group=min_space_group,
        )
        total.merge(para_stats)
    return total


def write_json(path: str, data: Dict[str, Any]) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def update_document_status(conn: sqlite3.Connection, document_id: int, output_status: str) -> None:
    conn.execute(
        """
        UPDATE document
        SET processing_status = ?, update_date = datetime('now')
        WHERE id = ?
        """,
        (output_status, document_id),
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="Optimize leading/middle spaces into tab stops and write raw_optimized.json")
    parser.add_argument("--db", required=True, help="Path to SQLite database")
    parser.add_argument("--limit", type=int, default=None, help="How many documents to process")
    parser.add_argument("--status", default=DEFAULT_STATUS, help="Input processing_status filter")
    parser.add_argument("--output-status", default=DEFAULT_OUTPUT_STATUS, help="Status to write after success")
    parser.add_argument("--dry-run", action="store_true", help="Run without writing files or DB updates")
    parser.add_argument("--uid", default=None, help="Process one specific uid")
    parser.add_argument("--overwrite", type=str2bool, default=True, help="Overwrite raw_optimized.json if it exists (default: true)")
    parser.add_argument("--min-space-group", type=int, default=DEFAULT_MIN_SPACE_GROUP, help="Minimum consecutive spaces to consider for conversion")
    parser.add_argument("--input-name", default=DEFAULT_INPUT_NAME, help="Input JSON filename inside artifacts folder")
    parser.add_argument("--output-name", default=DEFAULT_OUTPUT_NAME, help="Output JSON filename inside artifacts folder")
    args = parser.parse_args()

    logger, log_path = setup_logger()
    conn = connect_db(args.db)
    font_metrics = FontMetrics()
    try:
        font_metrics.load_from_db(conn)
    except Exception as exc:
        logger.error(f"error | bootstrap | failed to load font_metric: {exc}")
        return 1

    documents = fetch_documents(conn, args.status, args.limit, args.uid)
    logger.info(f"optimized | bootstrap | selected_documents={len(documents)} dry_run={args.dry_run} log_path={log_path}")

    processed = 0
    try:
        for doc in documents:
            doc_id = int(doc["id"])
            uid = str(doc["uid"])
            artifacts_path = doc["artifacts_abs_path"]
            if not artifacts_path:
                logger.error(f"error | {uid} | artifacts_abs_path is empty")
                continue

            input_path = os.path.join(artifacts_path, args.input_name)
            output_path = os.path.join(artifacts_path, args.output_name)
            print(uid)

            try:
                if not os.path.exists(input_path):
                    raise FileNotFoundError(f"Input file not found: {input_path}")
                if os.path.exists(output_path) and not args.overwrite:
                    logger.info(f"optimized | {uid} | skipped existing output overwrite=false")
                    continue

                with open(input_path, "r", encoding="utf-8") as f:
                    data = json.load(f)

                optimized = copy.deepcopy(data)
                stats = process_document(optimized, font_metrics, args.min_space_group)

                if not args.dry_run:
                    write_json(output_path, optimized)
                    update_document_status(conn, doc_id, args.output_status)
                    conn.commit()

                logger.info(
                    "optimized | %s | found=%s converted=%s trimmed=%s skipped_table=%s skipped_preserve=%s skipped_prose=%s paragraphs_changed=%s output=%s",
                    uid,
                    stats.spaces_found,
                    stats.spaces_converted,
                    stats.trailing_trimmed,
                    stats.table_candidates_skipped,
                    stats.preserve_skipped,
                    stats.prose_skipped,
                    stats.paragraphs_changed,
                    output_path if not args.dry_run else "dry-run",
                )
                processed += 1
            except Exception as exc:
                if not args.dry_run:
                    conn.rollback()
                logger.error(f"error | {uid} | {exc}")
        logger.info(f"optimized | summary | processed={processed} selected={len(documents)} dry_run={args.dry_run}")
    finally:
        conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
