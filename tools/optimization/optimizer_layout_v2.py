from __future__ import annotations

import argparse
import copy
import datetime as dt
import json
import math
import os
import sqlite3
from dataclasses import dataclass
from typing import Any, Dict, Iterable, List, Optional, Tuple

EMU_PER_TWIP = 635
DEFAULT_TAB_WIDTH_TWIPS = 720
DEFAULT_PARAGRAPH_HEIGHT_RESERVE_TWIPS = 400
FONT_ALIASES = {"unknown": "times new roman"}
OBJECT_RUN_TYPES = {"shape", "picture"}
TABLE_ROW_HEIGHT_MULTIPLIER = 1.5
OUTER_PARAGRAPH_HEIGHT_MULTIPLIER = 1.0
FLOATING_TABLE_PARAGRAPH_BUFFER = 2
INLINE_SPACER_TEXT = " "


class RunLogger:
    def __init__(self, base_dir: str):
        ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        log_dir = os.path.join(base_dir, "tools", "log")
        os.makedirs(log_dir, exist_ok=True)
        self.path = os.path.join(log_dir, f"optimizer_layout_{ts}.log")

    def write(self, status: str, uid: str, message: str) -> None:
        now = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(self.path, "a", encoding="utf-8") as f:
            f.write(f"{now} | {status} | {uid} | {message}\n")


class FontMetricResolver:
    def __init__(self, conn: sqlite3.Connection):
        self.font_by_id: Dict[int, str] = {}
        self.metrics: Dict[Tuple[str, float], Dict[int, float]] = {}
        self.base_sizes: Dict[str, float] = {}
        self._load(conn)

    def _load(self, conn: sqlite3.Connection) -> None:
        cur = conn.cursor()
        cur.execute("SELECT id, normalized_title FROM font")
        for font_id, normalized_title in cur.fetchall():
            self.font_by_id[int(font_id)] = str(normalized_title).strip().lower()

        cur.execute("SELECT font_id, base_size_pt, char_code, width_units FROM font_metric")
        for font_id, base_size_pt, char_code, width_units in cur.fetchall():
            font_name = self.font_by_id.get(int(font_id))
            if not font_name:
                continue
            key = (font_name, float(base_size_pt))
            self.metrics.setdefault(key, {})[int(char_code)] = float(width_units)
            self.base_sizes[font_name] = float(base_size_pt)

    def _normalize_font_name(self, font_name: Optional[str]) -> str:
        name = (font_name or "").strip().lower()
        if not name:
            name = "unknown"
        return FONT_ALIASES.get(name, name)

    def resolve_font_name(self, run, doc_defaults, style_r_format=None) -> str:
        r_format = run.get("r_format") or {}
        fonts = r_format.get("rFonts") or {}

        for key in ("ascii", "hAnsi", "eastAsia", "cs"):
            val = fonts.get(key)
            if isinstance(val, str) and val.strip():
                return self._normalize_font_name(val)

        if style_r_format:
            fonts = style_r_format.get("rFonts") or {}
            for key in ("ascii", "hAnsi", "eastAsia", "cs"):
                val = fonts.get(key)
                if isinstance(val, str) and val.strip():
                    return self._normalize_font_name(val)

        defaults = (((doc_defaults or {}).get("r_format") or {}).get("rFonts") or {})
        for key in ("ascii", "hAnsi", "eastAsia", "cs"):
            val = defaults.get(key)
            if isinstance(val, str) and val.strip():
                return self._normalize_font_name(val)

        return "times new roman"

    def resolve_size_pt(self, run, doc_defaults, style_r_format=None):
        r_format = run.get("r_format") or {}
        sz = r_format.get("font_size_half_points")
        if isinstance(sz, int):
            return sz / 2.0

        if style_r_format:
            sz = style_r_format.get("font_size_half_points")
            if isinstance(sz, int):
                return sz / 2.0

        defaults = ((doc_defaults or {}).get("r_format") or {})
        sz = defaults.get("font_size_half_points")
        if isinstance(sz, int):
            return sz / 2.0

        return 12.0

    def char_width(self, ch: str, font_name: str, size_pt: float) -> float:
        font_name = self._normalize_font_name(font_name)
        base_size = self.base_sizes.get(font_name)
        if base_size is None:
            font_name = "times new roman"
            base_size = self.base_sizes.get(font_name, 12.0)
        metric_map = self.metrics.get((font_name, base_size), {})
        width = metric_map.get(ord(ch))
        if width is None:
            if ch == " ":
                width = metric_map.get(ord(" "), 180.0)
            else:
                width = metric_map.get(ord("n")) or metric_map.get(ord("a")) or 360.0
        return width * (size_pt / float(base_size))

    def text_width_twips(self, text: str, run: Dict[str, Any], doc_defaults: Dict[str, Any], style_r_format: Optional[Dict[str, Any]] = None) -> float:
        if not text:
            return 0.0
        font_name = self.resolve_font_name(run, doc_defaults, style_r_format)
        size_pt = self.resolve_size_pt(run, doc_defaults, style_r_format)
        return sum(self.char_width(ch, font_name, size_pt) for ch in text)


@dataclass
class Unit:
    kind: str
    source_run: Dict[str, Any]
    text: str = ""
    width_twips: float = 0.0
    tab_val: str = "left"
    tab_pos_twip: Optional[int] = None
    spacing_twip: Optional[int] = None


def emu_to_twips(v: Optional[int]) -> Optional[int]:
    if v is None:
        return None
    return int(round(v / EMU_PER_TWIP))


def parse_size_type_twips(size_obj: Optional[Dict[str, Any]]) -> Optional[int]:
    if not isinstance(size_obj, dict):
        return None
    if size_obj.get("type") == "dxa" and isinstance(size_obj.get("w"), int):
        return int(size_obj["w"])
    return None


def object_width_twips(run: Dict[str, Any]) -> float:
    if run.get("type") == "picture":
        ext = run.get("extent") or {}
        if isinstance(ext, dict):
            return float(emu_to_twips(ext.get("cx")) or 0)
    if run.get("type") == "shape":
        shp = run.get("shape") or {}
        ext = shp.get("extent") or {}
        if isinstance(ext, dict):
            return float(emu_to_twips(ext.get("cx")) or 0)
    return 0.0


def shape_is_floating_unsafe(run: Dict[str, Any]) -> bool:
    if run.get("type") != "shape":
        return False

    shape = run.get("shape") or {}
    positioning = shape.get("positioning")

    if not isinstance(positioning, dict):
        return False

    layout = str(
        positioning.get("layout")
        or positioning.get("type")
        or positioning.get("anchor")
        or ""
    ).strip().lower()

    wrap = str(
        positioning.get("wrap")
        or shape.get("wrap")
        or ""
    ).strip().lower()

    # Явно безопасные случаи
    if layout == "inline":
        return False

    if wrap in {"behindtext", "behind", "infrontoftext", "front", "none"}:
        return False

    # Если positioning есть, но это не inline-safe случай — считаем shape опасной
    return True


def floating_shape_affected_paragraphs(run: Dict[str, Any]) -> int:
    shape = run.get("shape") or {}
    extent = shape.get("extent") or {}
    cy_emu = extent.get("cy")

    if not isinstance(cy_emu, int):
        return 1

    shape_height_twips = emu_to_twips(cy_emu)
    if not isinstance(shape_height_twips, int) or shape_height_twips <= 0:
        return 1

    return 2 + int(math.ceil(shape_height_twips / DEFAULT_PARAGRAPH_HEIGHT_RESERVE_TWIPS))


def paragraph_floating_shape_skip_span(paragraph: Dict[str, Any]) -> int:
    spans = [floating_shape_affected_paragraphs(run) for run in (paragraph.get("runs") or []) if isinstance(run, dict) and shape_is_floating_unsafe(run)]
    return max(spans) if spans else 0


def split_text_chunks(text: str) -> List[Tuple[str, str]]:
    if not text:
        return []
    out: List[Tuple[str, str]] = []
    buf: List[str] = []
    mode: Optional[str] = None
    for ch in text:
        cur = "space" if ch == " " else "text"
        if mode is None:
            mode = cur
            buf.append(ch)
            continue
        if cur == mode:
            buf.append(ch)
        else:
            out.append((mode, "".join(buf)))
            buf = [ch]
            mode = cur
    if buf:
        out.append((mode or "text", "".join(buf)))
    return out

def paragraph_existing_tabs(paragraph: Dict[str, Any]) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    for tab in ((paragraph.get("p_format") or {}).get("tabs") or []):
        if not isinstance(tab, dict):
            continue
        pos = tab.get("posTwip")
        val = tab.get("val")
        if isinstance(pos, int) and isinstance(val, str):
            out.append({"posTwip": int(pos), "val": val, "leader": tab.get("leader")})
    out.sort(key=lambda x: x["posTwip"])
    return out


def next_tab_stop(current_x: float, existing_tabs: List[Dict[str, Any]], default_tab_width: int) -> Tuple[int, str]:
    for tab in existing_tabs:
        pos = tab["posTwip"]
        if pos > current_x + 0.5:
            return int(pos), str(tab.get("val") or "left")
    if default_tab_width <= 0:
        default_tab_width = DEFAULT_TAB_WIDTH_TWIPS
    return int(math.floor(current_x / default_tab_width) + 1) * default_tab_width, "left"


def build_units(paragraph: Dict[str, Any], doc_defaults: Dict[str, Any], resolver: FontMetricResolver, existing_tabs: List[Dict[str, Any]], default_tab_width: int, style_r_format: Optional[Dict[str, Any]] = None) -> List[Unit]:
    units: List[Unit] = []
    current_x = 0.0
    for run in paragraph.get("runs") or []:
        run_type = run.get("type")
        if run_type == "text":
            text = run.get("text") or ""
            spacing_twip = ((run.get("r_format") or {}).get("spacing_twip"))
            for kind, chunk in split_text_chunks(text):
                chunk_width = resolver.text_width_twips(chunk, run, doc_defaults, style_r_format)
                if isinstance(spacing_twip, int) and chunk:
                    chunk_width += int(spacing_twip)
                units.append(Unit(kind="space" if kind == "space" else "text", source_run=run, text=chunk, width_twips=chunk_width, spacing_twip=spacing_twip if isinstance(spacing_twip, int) else None))
                current_x += units[-1].width_twips
        elif run_type == "tab":
            pos, val = next_tab_stop(current_x, existing_tabs, default_tab_width)
            units.append(Unit(kind="tab", source_run=run, width_twips=max(0.0, float(pos) - current_x), tab_val=val, tab_pos_twip=pos))
            current_x = float(pos)
        elif run_type in OBJECT_RUN_TYPES:
            width = object_width_twips(run)
            units.append(Unit(kind="object", source_run=run, width_twips=width))
            current_x += width
        else:
            units.append(Unit(kind="other", source_run=run, width_twips=0.0))
    return units


def prose_like(paragraph: Dict[str, Any]) -> bool:
    full_text = "".join((run.get("text") or "") for run in (paragraph.get("runs") or []) if run.get("type") == "text")

    # Правило 1: для prose-like учитываем только видимые символы.
    visible_text = "".join(ch for ch in full_text if not ch.isspace())
    if len(visible_text) >= 120:
        return True

    # Правило 2:
    # если есть хотя бы одна большая последовательность пробелов (> 10),
    # абзац должен оптимизироваться, т.е. prose-like = False.
    # Правило 3:
    # если есть более 3 последовательностей пробелов длиной > 4,
    # абзац тоже должен оптимизироваться.
    space_runs: List[int] = []
    current_space_run = 0
    for ch in full_text:
        if ch == " ":
            current_space_run += 1
        else:
            if current_space_run > 0:
                space_runs.append(current_space_run)
                current_space_run = 0
    if current_space_run > 0:
        space_runs.append(current_space_run)

    if any(run_len > 10 for run_len in space_runs):
        return False

    if sum(1 for run_len in space_runs if run_len > 4) > 3:
        return False

    raw_tokens = [w for w in visible_text.split() if w]
    meaningful_tokens: List[str] = []

    for token in raw_tokens:
        # Любой токен с буквами считаем содержательным словом
        if any(ch.isalpha() for ch in token):
            meaningful_tokens.append(token)
            continue

        # Чистую пунктуацию и одиночные цифры не считаем "словами" для prose эвристики
        stripped = token.strip(".,;:!?()[]{}\"'`-–—_/\\")
        if len(stripped) >= 2:
            meaningful_tokens.append(token)

    if len(meaningful_tokens) >= 12:
        return True

    return (
        len(meaningful_tokens) >= 8
        and "   " not in full_text
        and "\t" not in full_text
    )


def unit_is_anchor(unit: Unit) -> bool:
    return unit.kind in {"text", "object"}


def unit_is_gap_member(unit: Unit) -> bool:
    return unit.kind in {"space", "tab"}


def anchor_width(unit: Optional[Unit]) -> float:
    return 0.0 if unit is None else unit.width_twips if unit.kind in {"text", "object"} else 0.0


def anchor_start_from_tabstop(tab_pos: int, tab_val: str, next_anchor_width: float) -> float:
    tab_val = (tab_val or "left").lower()
    if tab_val == "right":
        return float(tab_pos) - next_anchor_width
    if tab_val == "center":
        return float(tab_pos) - (next_anchor_width / 2.0)
    return float(tab_pos)


def rightmost_tab_in_gap(gap_units: List[Unit]) -> Optional[Unit]:
    tabs = [u for u in gap_units if u.kind == "tab" and isinstance(u.tab_pos_twip, int)]
    if not tabs:
        return None
    return max(tabs, key=lambda u: int(u.tab_pos_twip or 0))


def suffix_width_after_unit(gap_units: List[Unit], unit: Unit) -> float:
    passed = False
    total = 0.0
    for u in gap_units:
        if not passed:
            if u is unit:
                passed = True
            continue
        total += u.width_twips
    return total


def gap_total_width(gap_units: List[Unit]) -> float:
    return sum(u.width_twips for u in gap_units)

def snap_twip(value: int, step: int = 30) -> int:
    return int(round(value / step) * step)

def is_plain_interword_gap(gap_units: List[Unit], left_anchor: Optional[Unit], right_anchor: Optional[Unit]) -> bool:
    if left_anchor is None or right_anchor is None:
        return False
    if left_anchor.kind != "text" or right_anchor.kind != "text":
        return False
    return len(gap_units) == 1 and gap_units[0].kind == "space" and gap_units[0].text == " "


def dedupe_and_sort_tabs(existing: List[Dict[str, Any]], added: Iterable[Tuple[int, str]]) -> List[Dict[str, Any]]:
    merged: List[Dict[str, Any]] = []
    seen: List[Tuple[int, str]] = []

    def push(pos: int, val: str = "left", leader: Optional[str] = None) -> None:
        for cur_pos, cur_val in seen:
            if abs(cur_pos - pos) <= 5 and cur_val == val:
                return
        seen.append((pos, val))
        rec = {"posTwip": int(pos), "val": val}
        if leader:
            rec["leader"] = leader
        merged.append(rec)

    for tab in existing:
        push(int(tab["posTwip"]), str(tab.get("val") or "left"), tab.get("leader"))
    for pos, val in added:
        push(int(pos), str(val or "left"))
    merged.sort(key=lambda x: x["posTwip"])
    return merged


def unit_is_layout_carrier(unit: Unit) -> bool:
    return unit.kind in {"tab", "inline_spacer"}


def group_units_for_runs(units: List[Unit]) -> List[List[Unit]]:
    groups: List[List[Unit]] = []

    for unit in units:
        if unit_is_layout_carrier(unit):
            groups.append([unit])
            continue

        if not groups:
            groups.append([unit])
            continue

        last_group = groups[-1]
        last_unit = last_group[-1]

        if unit_is_layout_carrier(last_unit):
            groups.append([unit])
            continue

        if unit.source_run is last_unit.source_run:
            last_group.append(unit)
        else:
            groups.append([unit])

    return groups


def materialize_run_from_group(group: List[Unit], parent_id: str, counter: int) -> Dict[str, Any]:
    first_unit = group[0]
    run = copy.deepcopy(first_unit.source_run)

    if first_unit.kind == "tab":
        run["type"] = "tab"
        run.pop("text", None)
        meta = dict(run.get("meta") or {})
        meta.pop("preserve", None)
        if meta:
            run["meta"] = meta
        else:
            run.pop("meta", None)
    elif first_unit.kind == "inline_spacer":
        run["type"] = "text"
        run["text"] = first_unit.text or INLINE_SPACER_TEXT
        meta = dict(run.get("meta") or {})
        meta["preserve"] = True
        run["meta"] = meta
        r_format = dict(run.get("r_format") or {})
        if isinstance(first_unit.spacing_twip, int):
            r_format["spacing_twip"] = int(first_unit.spacing_twip)
        else:
            r_format.pop("spacing_twip", None)
        if r_format:
            run["r_format"] = r_format
        else:
            run.pop("r_format", None)
    elif all(unit.kind in {"text", "space"} for unit in group):
        run["type"] = "text"
        run["text"] = "".join(unit.text for unit in group)
        if any(unit.kind == "space" and unit.text for unit in group):
            meta = dict(run.get("meta") or {})
            meta["preserve"] = True
            run["meta"] = meta
    else:
        # Non-text groups are expected to remain singleton and preserve the original run payload.
        pass

    run["id"] = f"{parent_id}.run_{counter}"
    run["parent_id"] = parent_id
    return run

def rebuild_runs_from_units(units: List[Unit], parent_id: str) -> List[Dict[str, Any]]:
    runs: List[Dict[str, Any]] = []
    groups = group_units_for_runs(units)
    for counter, group in enumerate(groups, start=1):
        if not group:
            continue
        runs.append(materialize_run_from_group(group, parent_id, counter))
    return runs


def measure_space_width_twips(run: Dict[str, Any], doc_defaults: Dict[str, Any], resolver: FontMetricResolver, style_r_format: Optional[Dict[str, Any]] = None) -> float:
    return resolver.text_width_twips(INLINE_SPACER_TEXT, run, doc_defaults, style_r_format)


def make_inline_spacer_unit(source_run: Dict[str, Any], target_width_twips: float, doc_defaults: Dict[str, Any], resolver: FontMetricResolver, style_r_format: Optional[Dict[str, Any]] = None) -> Unit:
    base_space_width = measure_space_width_twips(source_run, doc_defaults, resolver, style_r_format)
    if target_width_twips <= base_space_width + 1.0:   # небольшая корректировка для округления
        return Unit(kind="space", source_run=source_run, text=INLINE_SPACER_TEXT, width_twips=base_space_width)
    return Unit(
        kind="inline_spacer",
        source_run=source_run,
        text=INLINE_SPACER_TEXT,
        width_twips=float(target_width_twips),
        spacing_twip=max(0, int(round(target_width_twips - base_space_width))),
    )


def first_anchor_info(units: List[Unit], starts: List[float]) -> Optional[Dict[str, Any]]:
    for idx, unit in enumerate(units):
        if unit_is_anchor(unit):
            start_x = starts[idx]
            return {
                "index": idx,
                "unit": unit,
                "start_x": start_x,
                "end_x": start_x + unit.width_twips,
            }
    return None


def leading_space_width_before_index(units: List[Unit], end_index: int) -> float:
    total = 0.0
    for idx in range(min(end_index, len(units))):
        unit = units[idx]
        if unit.kind != "space":
            break
        total += unit.width_twips
    return total


def leading_space_end_index(units: List[Unit], end_index: int) -> int:
    idx = 0
    limit = min(end_index, len(units))
    while idx < limit and units[idx].kind == "space":
        idx += 1
    return idx



def paragraph_indent_markers(p_format: Dict[str, Any]) -> Tuple[int, int, int, int, int]:
    left_base = int(p_format.get("indent_start_twip") or 0)
    first_line_delta = int(p_format.get("indent_first_line_twip") or 0)
    hanging_delta = int(p_format.get("indent_hanging_twip") or 0)

    if hanging_delta != 0:
        first_line_marker = left_base - hanging_delta
    else:
        first_line_marker = left_base + first_line_delta

    other_lines_marker = left_base
    return left_base, first_line_delta, hanging_delta, first_line_marker, other_lines_marker


def apply_first_line_marker_to_pformat(p_format: Dict[str, Any], new_first_line_marker: int, keep_other_lines_marker: int) -> Dict[str, Any]:
    updated = dict(p_format or {})
    left_base = int(keep_other_lines_marker)
    updated["indent_start_twip"] = left_base
    updated.pop("indent_first_line_twip", None)

    if new_first_line_marker < left_base:
        hanging_delta = left_base - int(new_first_line_marker)
        if hanging_delta != 0:
            updated["indent_hanging_twip"] = hanging_delta
        else:
            updated.pop("indent_hanging_twip", None)
    else:
        updated.pop("indent_hanging_twip", None)
        first_line_delta = int(new_first_line_marker) - left_base
        if first_line_delta != 0:
            updated["indent_first_line_twip"] = first_line_delta

    return updated


def paragraph_effective_origins(p_format: Dict[str, Any]) -> Tuple[int, int, int, int, int]:
    indent_start = int(p_format.get("indent_start_twip") or 0)
    indent_first_line = int(p_format.get("indent_first_line_twip") or 0)
    indent_hanging = int(p_format.get("indent_hanging_twip") or 0)

    if indent_hanging != 0:
        first_line_origin = indent_start - indent_hanging
        other_lines_origin = indent_start
    else:
        first_line_origin = indent_start + indent_first_line
        other_lines_origin = indent_start

    return indent_start, indent_first_line, indent_hanging, first_line_origin, other_lines_origin


def apply_first_line_origin_to_pformat(p_format: Dict[str, Any], new_first_line_origin: int, keep_other_lines_origin: int) -> Dict[str, Any]:
    updated = dict(p_format or {})
    indent_hanging = int(updated.get("indent_hanging_twip") or 0)

    if indent_hanging != 0:
        updated["indent_start_twip"] = int(keep_other_lines_origin)
        if new_first_line_origin <= keep_other_lines_origin:
            new_hanging = int(keep_other_lines_origin) - int(new_first_line_origin)
            if new_hanging != 0:
                updated["indent_hanging_twip"] = int(new_hanging)
            else:
                updated.pop("indent_hanging_twip", None)
            updated.pop("indent_first_line_twip", None)
        else:
            updated.pop("indent_hanging_twip", None)
            new_first_line = int(new_first_line_origin) - int(keep_other_lines_origin)
            if new_first_line != 0:
                updated["indent_first_line_twip"] = int(new_first_line)
            else:
                updated.pop("indent_first_line_twip", None)
        return updated

    indent_start = int(updated.get("indent_start_twip") or 0)
    new_first_line = int(new_first_line_origin) - indent_start
    if new_first_line != 0:
        updated["indent_first_line_twip"] = int(new_first_line)
    else:
        updated.pop("indent_first_line_twip", None)
    return updated


def build_anchor_positions(units: List[Unit], starts: List[float]) -> Dict[int, Dict[str, float]]:
    out: Dict[int, Dict[str, float]] = {}
    for idx, unit in enumerate(units):
        if unit_is_anchor(unit):
            start_x = starts[idx]
            out[idx] = {
                "start": start_x,
                "end": start_x + unit.width_twips,
                "width": unit.width_twips,
            }
    return out


def zero_stats() -> Dict[str, int]:
    return {"paragraphs": 0, "gaps_found": 0, "gaps_converted": 0, "trailing_removed": 0, "skipped_table": 0, "skipped_prose": 0, "skipped_uncertain": 0}


def add_stats(dst: Dict[str, int], src: Dict[str, int]) -> None:
    for key, value in src.items():
        dst[key] = dst.get(key, 0) + int(value)


def optimize_paragraph(paragraph: Dict[str, Any], doc_defaults: Dict[str, Any], resolver: FontMetricResolver, context: Dict[str, Any], min_space_group: int, default_tab_width: int, optimize_in_tables: bool, styles: Dict[str, Any]) -> Dict[str, int]:
    stats = zero_stats()
    stats["paragraphs"] = 1
    if context.get("in_table") and not optimize_in_tables:
        stats["skipped_table"] = 1
        return stats
    if prose_like(paragraph):
        stats["skipped_prose"] = 1
        return stats
    # Получаем стиль абзаца по p_style_id
    p_style_id = paragraph.get("p_style_id")
    paragraph_style = styles.get(p_style_id) if p_style_id else None
    paragraph_style_r = (paragraph_style.get("r_format") if paragraph_style else {}) or {}

    existing_tabs_full = paragraph_existing_tabs(paragraph)
    units = build_units(paragraph, doc_defaults, resolver, existing_tabs_full, default_tab_width, paragraph_style_r)
    if not units:
        return stats

    p_format = dict(paragraph.get("p_format") or {})
    _, _, _, first_line_marker_original, other_lines_marker_original = paragraph_indent_markers(p_format)
    first_line_marker_effective = first_line_marker_original

    starts: List[float] = []
    current_x = float(first_line_marker_original)
    for unit in units:
        starts.append(current_x)
        current_x += unit.width_twips

    anchor_map = build_anchor_positions(units, starts)
    first_anchor = first_anchor_info(units, starts)
    first_anchor_idx: Optional[int] = None
    leading_spaces_width = 0.0
    leading_spaces_end = 0
    special_inline_gap_after_first_anchor = False
    forced_leading_tab_pos: Optional[int] = None
    force_inline_spacer_mode = bool(context.get("force_inline_spacer_mode"))
    drop_leading_spaces = False

    # (остальная логика без изменений)
    if first_line_marker_original < 0 and first_anchor is not None:
        first_anchor_idx = int(first_anchor["index"])
        first_anchor_start_x = float(anchor_map[first_anchor_idx]["start"])
        first_anchor_end_x = float(anchor_map[first_anchor_idx]["end"])
        leading_spaces_width = leading_space_width_before_index(units, first_anchor_idx)
        leading_spaces_end = leading_space_end_index(units, first_anchor_idx)
        has_leading_spaces = leading_spaces_width > 0.5

        if first_anchor_start_x < 0.0 < first_anchor_end_x:
            # Case A: first anchor crosses zero.
            if has_leading_spaces:
                drop_leading_spaces = True
                first_line_marker_effective = int(round(first_anchor_start_x))
        elif first_anchor_start_x < 0.0 and first_anchor_end_x <= 0.0:
            # Case B: first anchor is fully in the negative zone.
            if has_leading_spaces:
                drop_leading_spaces = True
                first_line_marker_effective = int(round(first_anchor_start_x))
            special_inline_gap_after_first_anchor = True
        elif first_anchor_start_x >= 0.0:
            # Case C: first anchor is already in the positive zone.
            if has_leading_spaces:
                drop_leading_spaces = True
            first_line_marker_effective = 10
            forced_leading_tab_pos = max(0, int(round(first_anchor_start_x)))

    out_units: List[Unit] = []
    added_tabs: List[Tuple[int, str]] = []
    if forced_leading_tab_pos is not None:
        out_units.append(Unit(kind="tab", source_run=units[0].source_run, width_twips=0.0, tab_val="left", tab_pos_twip=forced_leading_tab_pos))
        added_tabs.append((forced_leading_tab_pos, "left"))

    i = 0

    while i < len(units):
        unit = units[i]
        if drop_leading_spaces and i < leading_spaces_end and unit.kind == "space":
            i += 1
            continue

        if not unit_is_gap_member(unit):
            out_units.append(unit)
            i += 1
            continue

        gap_start = i
        while i < len(units) and unit_is_gap_member(units[i]):
            i += 1
        gap_end = i
        gap_units = units[gap_start:gap_end]
        stats["gaps_found"] += 1

        left_anchor: Optional[Unit] = None
        left_anchor_idx: Optional[int] = None
        li = gap_start - 1
        while li >= 0:
            if unit_is_anchor(units[li]):
                left_anchor = units[li]
                left_anchor_idx = li
                break
            if units[li].kind == "other":
                break
            li -= 1

        right_anchor: Optional[Unit] = None
        right_anchor_idx: Optional[int] = None
        ri = gap_end
        while ri < len(units):
            if unit_is_anchor(units[ri]):
                right_anchor = units[ri]
                right_anchor_idx = ri
                break
            if units[ri].kind == "other":
                break
            ri += 1

        if right_anchor is None or right_anchor_idx is None:
            # Special case: trailing tabs at paragraph end are layout carriers.
            # Preserve one trailing tab (without tab stop) and convert preceding
            # spaces to inline_spacer instead of deleting the whole trailing gap.
            trailing_tabs = [u for u in gap_units if u.kind == "tab"]
            if trailing_tabs:
                first_tab_idx = next((idx for idx, u in enumerate(gap_units) if u.kind == "tab"), None)
                if first_tab_idx is not None:
                    space_prefix = gap_units[:first_tab_idx]
                    space_prefix_width = sum(u.width_twips for u in space_prefix if u.kind == "space")
                    space_source = next((u.source_run for u in space_prefix if u.kind == "space"), None)
                    tab_source = trailing_tabs[0].source_run

                    if space_prefix_width > 0.5 and space_source is not None:
                        out_units.append(
                            make_inline_spacer_unit(
                                space_source,
                                space_prefix_width,
                                doc_defaults,
                                resolver,
                                paragraph_style_r
                            )
                        )
                    out_units.append(Unit(kind="tab", source_run=tab_source, width_twips=0.0, tab_val="left", tab_pos_twip=None))
                    stats["gaps_converted"] += 1
                    continue
            stats["trailing_removed"] += 1
            continue

        if is_plain_interword_gap(gap_units, left_anchor, right_anchor):
            out_units.extend(gap_units)
            continue

        total_space_chars = sum(len(u.text) for u in gap_units if u.kind == "space")
        has_tab = any(u.kind == "tab" for u in gap_units)
        if not has_tab and total_space_chars < min_space_group:
            out_units.extend(gap_units)
            continue

        right_anchor_start_x = float(anchor_map[right_anchor_idx]["start"])
        right_anchor_end_x = float(anchor_map[right_anchor_idx]["end"])
        right_anchor_width = float(anchor_map[right_anchor_idx]["width"])

        if (
            special_inline_gap_after_first_anchor
            and first_anchor_idx is not None
            and left_anchor_idx == first_anchor_idx
            and gap_units
        ):
            left_anchor_end_x = float(anchor_map[left_anchor_idx]["end"])
            required_gap_width = max(0.0, right_anchor_start_x - left_anchor_end_x)
            out_units.append(make_inline_spacer_unit(gap_units[0].source_run, required_gap_width, doc_defaults, resolver, paragraph_style_r))
            special_inline_gap_after_first_anchor = False
            stats["gaps_converted"] += 1
            continue

        if force_inline_spacer_mode:
            if left_anchor_idx is not None:
                left_anchor_end_x = float(anchor_map[left_anchor_idx]["end"])
                required_gap_width = max(0.0, right_anchor_start_x - left_anchor_end_x)
            else:
                required_gap_width = max(0.0, gap_total_width(gap_units))

            out_units.append(
                make_inline_spacer_unit(
                    gap_units[0].source_run,
                    required_gap_width,
                    doc_defaults,
                    resolver,
                    paragraph_style_r
                )
            )
            stats["gaps_converted"] += 1
            continue

        dominating_tab = rightmost_tab_in_gap(gap_units)
        if dominating_tab is not None and dominating_tab.tab_pos_twip is not None:
            final_tab_val = dominating_tab.tab_val or "left"
        else:
            final_tab_val = "left"

        if final_tab_val == "right":
            final_tab_pos = int(round(right_anchor_end_x))
        elif final_tab_val == "center":
            final_tab_pos = int(round(right_anchor_start_x + (right_anchor_width / 2.0)))
        else:
            final_tab_pos = int(round(right_anchor_start_x))

        if final_tab_pos < 0:
            stats["skipped_uncertain"] += 1
            out_units.extend(gap_units)
            continue

        out_units.append(Unit(kind="tab", source_run=gap_units[0].source_run, width_twips=0.0, tab_val=final_tab_val, tab_pos_twip=final_tab_pos))
        added_tabs.append((final_tab_pos, final_tab_val))
        stats["gaps_converted"] += 1

    paragraph["runs"] = rebuild_runs_from_units(out_units, paragraph["id"])
    if first_line_marker_effective != first_line_marker_original or added_tabs:
        p_format = dict(paragraph.get("p_format") or {})
        if first_line_marker_effective != first_line_marker_original:
            p_format = apply_first_line_marker_to_pformat(p_format, int(first_line_marker_effective), int(other_lines_marker_original))
        if added_tabs:
            p_format["tabs"] = dedupe_and_sort_tabs(existing_tabs_full, added_tabs)
        paragraph["p_format"] = p_format
    return stats

def table_cell_width_twips(table: Dict[str, Any], row: Dict[str, Any], cell_index: int) -> Optional[int]:
    # (без изменений)
    cells = row.get("cells") or []
    if cell_index >= len(cells):
        return None
    cell = cells[cell_index]
    tc_pr = cell.get("tcPr") or {}
    tc_w = parse_size_type_twips(tc_pr.get("tcW"))
    if tc_w is not None:
        return tc_w
    grid = table.get("tbl_grid") or []
    if isinstance(grid, list) and cell_index < len(grid) and isinstance(grid[cell_index], int):
        return int(grid[cell_index])
    tbl_pr = table.get("tblPr") or {}
    return parse_size_type_twips(tbl_pr.get("tblW"))


def table_is_floating(table: Dict[str, Any]) -> bool:
    tbl_pr = table.get("tblPr") or {}
    return isinstance(tbl_pr.get("tblpPr"), dict)


def paragraph_font_height_twips(paragraph: Dict[str, Any], doc_defaults: Dict[str, Any], resolver: FontMetricResolver) -> int:
    max_size_pt = 0.0
    for run in paragraph.get("runs") or []:
        if not isinstance(run, dict):
            continue
        max_size_pt = max(max_size_pt, resolver.resolve_size_pt(run, doc_defaults))

    if max_size_pt <= 0.0:
        defaults = (doc_defaults or {}).get("r_format") or {}
        sz = defaults.get("font_size_half_points")
        if isinstance(sz, int):
            max_size_pt = sz / 2.0
        else:
            max_size_pt = 12.0

    return max(1, int(round(max_size_pt * 20.0)))


def estimate_paragraph_height_twips(
    paragraph: Dict[str, Any],
    doc_defaults: Dict[str, Any],
    resolver: FontMetricResolver,
    multiplier: float,
) -> int:
    base = paragraph_font_height_twips(paragraph, doc_defaults, resolver)
    return max(1, int(math.ceil(base * float(multiplier))))


def estimate_cell_content_height_twips(
    cell: Dict[str, Any],
    doc_defaults: Dict[str, Any],
    resolver: FontMetricResolver,
) -> int:
    total = 0
    for item in cell.get("content") or []:
        if not isinstance(item, dict):
            continue
        if item.get("type") == "paragraph":
            total += estimate_paragraph_height_twips(item, doc_defaults, resolver, TABLE_ROW_HEIGHT_MULTIPLIER)
        elif item.get("type") == "table":
            total += estimate_table_height_twips(item, doc_defaults, resolver)
    return total


def estimate_table_row_height_twips(
    row: Dict[str, Any],
    doc_defaults: Dict[str, Any],
    resolver: FontMetricResolver,
) -> int:
    tr_pr = row.get("trPr") or {}
    tr_height = tr_pr.get("trHeight") or {}
    explicit_val = tr_height.get("val") if isinstance(tr_height, dict) else None
    h_rule = str(tr_height.get("hRule") or "").strip().lower() if isinstance(tr_height, dict) else ""

    content_height = 0
    for cell in row.get("cells") or []:
        if isinstance(cell, dict):
            content_height = max(content_height, estimate_cell_content_height_twips(cell, doc_defaults, resolver))

    if isinstance(explicit_val, int):
        if h_rule == "exact":
            return max(1, int(explicit_val))
        return max(int(explicit_val), content_height, 1)

    return max(content_height, DEFAULT_PARAGRAPH_HEIGHT_RESERVE_TWIPS)


def estimate_table_height_twips(table: Dict[str, Any], doc_defaults: Dict[str, Any], resolver: FontMetricResolver) -> int:
    total = 0
    for row in table.get("rows") or []:
        if isinstance(row, dict):
            total += estimate_table_row_height_twips(row, doc_defaults, resolver)
    return max(total, DEFAULT_PARAGRAPH_HEIGHT_RESERVE_TWIPS)


def process_shape_content(shape_run: Dict[str, Any], doc_defaults: Dict[str, Any], resolver: FontMetricResolver, context: Dict[str, Any], min_space_group: int, default_tab_width: int, optimize_in_tables: bool, styles: Dict[str, Any]) -> Dict[str, int]:
    stats = zero_stats()
    shape = shape_run.get("shape") or {}
    content = shape.get("content")
    if not isinstance(content, list):
        return stats
    shape_ctx = dict(context)
    shape_ctx["in_shape"] = True
    ext = (shape.get("extent") or {}).get("cx")
    if isinstance(ext, int):
        shape_ctx["shape_width_twips"] = emu_to_twips(ext)
    for item in content:
        if isinstance(item, dict) and item.get("type") == "paragraph":
            add_stats(stats, optimize_paragraph(item, doc_defaults, resolver, shape_ctx, min_space_group, default_tab_width, optimize_in_tables, styles))
            for run in item.get("runs") or []:
                if isinstance(run, dict) and run.get("type") == "shape":
                    add_stats(stats, process_shape_content(run, doc_defaults, resolver, shape_ctx, min_space_group, default_tab_width, optimize_in_tables, styles))
    return stats


def process_content(content: List[Dict[str, Any]], doc_defaults: Dict[str, Any], resolver: FontMetricResolver, context: Dict[str, Any], min_space_group: int, default_tab_width: int, optimize_in_tables: bool, styles: Dict[str, Any]) -> Dict[str, int]:
    stats = zero_stats()
    skip_remaining_paragraphs = 0
    inline_spacer_remaining_paragraphs = 0

    for item in content:
        if not isinstance(item, dict):
            continue
        if item.get("type") == "paragraph":
            local_skip_span = paragraph_floating_shape_skip_span(item)
            apply_inline_spacer_mode = inline_spacer_remaining_paragraphs > 0

            if skip_remaining_paragraphs > 0:
                stats["skipped_uncertain"] += 1
                skip_remaining_paragraphs -= 1
                if apply_inline_spacer_mode:
                    inline_spacer_remaining_paragraphs -= 1
                continue

            if local_skip_span > 0:
                stats["skipped_uncertain"] += 1
                skip_remaining_paragraphs = max(skip_remaining_paragraphs, local_skip_span - 1)
                if apply_inline_spacer_mode:
                    inline_spacer_remaining_paragraphs -= 1
                continue

            paragraph_context = dict(context)
            if apply_inline_spacer_mode:
                paragraph_context["force_inline_spacer_mode"] = True

            add_stats(
                stats,
                optimize_paragraph(
                    item,
                    doc_defaults,
                    resolver,
                    paragraph_context,
                    min_space_group,
                    default_tab_width,
                    optimize_in_tables,
                    styles,
                ),
            )
            if apply_inline_spacer_mode:
                inline_spacer_remaining_paragraphs -= 1

            for run in item.get("runs") or []:
                if isinstance(run, dict) and run.get("type") == "shape":
                    add_stats(stats, process_shape_content(run, doc_defaults, resolver, context, min_space_group, default_tab_width, optimize_in_tables, styles))
        elif item.get("type") == "table":
            for row in item.get("rows") or []:
                for idx, cell in enumerate(row.get("cells") or []):
                    if not isinstance(cell, dict):
                        continue
                    cell_ctx = dict(context)
                    cell_ctx["in_table"] = True
                    cell_ctx.pop("force_inline_spacer_mode", None)
                    cell_ctx["cell_width_twips"] = table_cell_width_twips(item, row, idx)
                    add_stats(stats, process_content(cell.get("content") or [], doc_defaults, resolver, cell_ctx, min_space_group, default_tab_width, optimize_in_tables, styles))

            if table_is_floating(item):
                table_height_twips = estimate_table_height_twips(item, doc_defaults, resolver)
                outer_paragraph_height_twips = max(
                    1,
                    int(
                        math.ceil(
                            paragraph_font_height_twips(
                                {"runs": [], "p_format": {}},
                                doc_defaults,
                                resolver,
                            ) * OUTER_PARAGRAPH_HEIGHT_MULTIPLIER
                        )
                    ),
                )
                affected_count = int(math.ceil(table_height_twips / float(outer_paragraph_height_twips))) + FLOATING_TABLE_PARAGRAPH_BUFFER
                inline_spacer_remaining_paragraphs = max(inline_spacer_remaining_paragraphs, affected_count)
    return stats


def fetch_documents(conn: sqlite3.Connection, status: str, uid: Optional[str], limit: Optional[int]) -> List[sqlite3.Row]:
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    if uid:
        cur.execute("SELECT id, uid, artifacts_abs_path, processing_status FROM document WHERE uid = ?", (uid,))
        rows = cur.fetchall()
        if status:
            rows = [r for r in rows if r["processing_status"] == status]
        return rows
    sql = "SELECT id, uid, artifacts_abs_path, processing_status FROM document WHERE processing_status = ? ORDER BY id"
    params: List[Any] = [status]
    if isinstance(limit, int) and limit > 0:
        sql += " LIMIT ?"
        params.append(limit)
    cur.execute(sql, params)
    return cur.fetchall()

def read_json(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def write_json(path: str, payload: Dict[str, Any]) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def update_status(conn: sqlite3.Connection, document_id: int, status: str) -> None:
    cur = conn.cursor()
    cur.execute("UPDATE document SET processing_status = ? WHERE id = ?", (status, document_id))
    conn.commit()


def optimize_document(row: sqlite3.Row, args: argparse.Namespace, conn: sqlite3.Connection, resolver: FontMetricResolver, logger: RunLogger) -> None:
    uid = row["uid"]
    artifacts_dir = row["artifacts_abs_path"]
    input_path = os.path.join(artifacts_dir, args.input_name)
    output_path = os.path.join(artifacts_dir, args.output_name)
    try:
        if not os.path.exists(input_path):
            logger.write("error", uid, f"missing_input={input_path}")
            print(uid)
            return
        if os.path.exists(output_path) and not args.overwrite:
            logger.write("optimized", uid, f"skipped_existing_output={output_path}")
            print(uid)
            return

        raw = read_json(input_path)
        doc_defaults = raw.get("doc_defaults") or {}
        default_tab_width = int((((raw.get("document_info") or {}).get("settings") or {}).get("defaultTabStopTwip")) or args.default_tab_width_twips or DEFAULT_TAB_WIDTH_TWIPS)
        styles = raw.get("styles") or {}
        stats = process_content(raw.get("content") or [], doc_defaults, resolver, {"in_table": False, "in_shape": False}, args.min_space_group, default_tab_width, args.optimize_in_tables, styles)

        if not args.dry_run:
            write_json(output_path, raw)
            # update_status(conn, int(row["id"]), args.output_status)

        logger.write("optimized", uid, f"paragraphs={stats['paragraphs']} gaps={stats['gaps_found']} converted={stats['gaps_converted']} trimmed={stats['trailing_removed']} skipped_table={stats['skipped_table']} skipped_prose={stats['skipped_prose']} skipped_uncertain={stats['skipped_uncertain']} dry_run={str(bool(args.dry_run)).lower()}")
        print(uid)
    except Exception as e:
        logger.write("error", uid, f"{type(e).__name__}: {e}")
        print(uid)


def str2bool(v: str) -> bool:
    val = str(v).strip().lower()
    if val in {"1", "true", "yes", "y", "on"}:
        return True
    if val in {"0", "false", "no", "n", "off"}:
        return False
    raise argparse.ArgumentTypeError(f"Invalid boolean value: {v}")


def detect_project_root(script_path: str) -> str:
    return os.path.abspath(os.path.join(os.path.dirname(script_path), "..", ".."))


def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Anchor/gap layout optimizer for RAW JSON")
    p.add_argument("--db", required=True, help="Path to SQLite database")
    p.add_argument("--limit", type=int, default=None, help="How many documents to process")
    p.add_argument("--status", default="parsed_raw", help="Input processing_status")
    p.add_argument("--output-status", default="optimized", help="Output processing_status")
    p.add_argument("--dry-run", action="store_true", help="Run without writing output or DB status")
    p.add_argument("--uid", default=None, help="Process one UID")
    p.add_argument("--overwrite", type=str2bool, default=True, help="Overwrite output file if it exists")
    p.add_argument("--input-name", default="raw.json", help="Input JSON file inside artifacts dir")
    p.add_argument("--output-name", default="raw_optimized.json", help="Output JSON file inside artifacts dir")
    p.add_argument("--default-tab-width-twips", type=int, default=DEFAULT_TAB_WIDTH_TWIPS, help="Fallback default tab width in twips")
    p.add_argument("--min-space-group", type=int, default=4, help="Minimum spaces in positioning gap if no tabs")
    p.add_argument("--optimize-in-tables", type=str2bool, default=True, help="Optimize paragraphs inside table cells")
    return p


def main() -> None:
    args = build_arg_parser().parse_args()
    project_root = detect_project_root(os.path.abspath(__file__))
    logger = RunLogger(project_root)
    conn = sqlite3.connect(args.db)
    conn.row_factory = sqlite3.Row
    try:
        resolver = FontMetricResolver(conn)
        rows = fetch_documents(conn, args.status, args.uid, args.limit)
        for row in rows:
            optimize_document(row, args, conn, resolver, logger)
    finally:
        conn.close()


if __name__ == "__main__":
    main()
