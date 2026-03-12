
# optimizer_layout.py
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
FONT_ALIASES = {"unknown": "times new roman"}
OBJECT_RUN_TYPES = {"shape", "picture"}


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

    def resolve_font_name(self, run: Dict[str, Any], doc_defaults: Dict[str, Any]) -> str:
        r_format = run.get("r_format") or {}
        fonts = r_format.get("rFonts") or {}

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

    def resolve_size_pt(self, run: Dict[str, Any], doc_defaults: Dict[str, Any]) -> float:
        r_format = run.get("r_format") or {}
        sz = r_format.get("font_size_half_points")
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

    def text_width_twips(self, text: str, run: Dict[str, Any], doc_defaults: Dict[str, Any]) -> float:
        if not text:
            return 0.0

        font_name = self.resolve_font_name(run, doc_defaults)
        size_pt = self.resolve_size_pt(run, doc_defaults)

        return sum(self.char_width(ch, font_name, size_pt) for ch in text)


@dataclass
class Token:
    kind: str
    text: str
    source_run: Dict[str, Any]
    width_twips: float
    object_type: Optional[str] = None


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


def object_width_twips(run: Dict[str, Any]) -> Optional[int]:
    if run.get("type") == "picture":
        ext = run.get("extent") or {}
        if isinstance(ext, dict):
            return emu_to_twips(ext.get("cx"))

    if run.get("type") == "shape":
        shp = run.get("shape") or {}
        ext = shp.get("extent") or {}
        if isinstance(ext, dict):
            return emu_to_twips(ext.get("cx"))

    return None


def detect_project_root(script_path: str) -> str:
    return os.path.abspath(os.path.join(os.path.dirname(script_path), "..", ".."))


def str2bool(v: str) -> bool:
    val = str(v).strip().lower()
    if val in {"1", "true", "yes", "y", "on"}:
        return True
    if val in {"0", "false", "no", "n", "off"}:
        return False
    raise argparse.ArgumentTypeError(f"Invalid boolean value: {v}")


def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Layout optimizer for RAW JSON corpus")

    p.add_argument("--db", required=True)
    p.add_argument("--limit", type=int, default=None)
    p.add_argument("--status", default="parsed_raw")
    p.add_argument("--output-status", default="optimized")
    p.add_argument("--dry-run", action="store_true")
    p.add_argument("--uid", default=None)
    p.add_argument("--overwrite", type=str2bool, default=True)

    p.add_argument("--input-name", default="raw.json")
    p.add_argument("--output-name", default="raw_optimized.json")

    p.add_argument("--default-tab-width-twips", type=int, default=DEFAULT_TAB_WIDTH_TWIPS)
    p.add_argument("--min-space-group", type=int, default=4)
    p.add_argument("--optimize-in-tables", type=str2bool, default=True)

    return p


def main():
    args = build_arg_parser().parse_args()

    project_root = detect_project_root(os.path.abspath(__file__))
    logger = RunLogger(project_root)

    conn = sqlite3.connect(args.db)
    conn.row_factory = sqlite3.Row

    print("optimizer_layout started")
    print("DB:", args.db)

    conn.close()


if __name__ == "__main__":
    main()
