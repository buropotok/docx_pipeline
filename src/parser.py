# UltimateParserV43.py
# DOCX -> RAW JSON (schema v2.9) with Normal base style and character styles
# Parser Version: v43
# Schema Version: 2.9
# Rules Version: 0.3

# Deterministic, visually-lossless for "forms" subset (no tables/images/fields/hyperlinks).
#
# Based on UltimateParserV42, with major improvements:
# 1) Normal base style with full default attributes (Times New Roman, 12pt, etc.)
# 2) Paragraph styles stored as diff from Normal
# 3) Character styles introduced; runs use char_style_id instead of diff
# 4) Numbering pStyle mapping preserved

from __future__ import annotations

import argparse
import json
import os
import zipfile
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
BASE_DIR = "."

# Default Normal style values (based on Word's Normal + observed donor)
NORMAL_P_FORMAT: Dict[str, Any] = {
    "alignment": "LEFT",
    "indentStartTwip": 0,
    "indentEndTwip": 0,
    "indentFirstLineTwip": 0,
    "indentHangingTwip": 0,
    "spaceBeforeTwip": 0,
    "spaceAfterTwip": 0,
    "lineTwip": 240,
    "lineRule": "AUTO",
    "keepNext": False,
    "keepLines": False,
    "pageBreakBefore": False,
    "widowControl": True,
    "snapToGrid": False,
    "contextualSpacing": False,
    # textAlignment not set (AUTO by absence)
    # tabs empty list, numbering absent
}
NORMAL_R_FORMAT: Dict[str, Any] = {
    "rFonts": {
        "ascii": "Times New Roman",
        "hAnsi": "Times New Roman",
        "eastAsia": "SimSun;宋体",
        "cs": "Times New Roman"
    },
    "font_size_half_points": 24,
    "bold": False,
    "italic": False,
    "underline": None,          # absent means none
    "color": "auto",
    "vertical_align": "baseline",
    "all_caps": False,
    "charSpacingTwip": 0,
    "positionHalfPoints": 0,
    "lang": {
        "eastAsia": "zh-CN"
    }
}


def qn(tag: str) -> str:
    pfx, local = tag.split(":")
    if pfx != "w":
        raise ValueError(f"Unsupported prefix: {pfx}")
    return f"{{{W_NS}}}{local}"


def _int_attr(el: etree._Element, attr_local: str) -> Optional[int]:
    v = el.get(f"{{{W_NS}}}{attr_local}")
    if v is None:
        return None
    try:
        return int(v)
    except Exception:
        return None


def _str_attr(el: etree._Element, attr_local: str) -> Optional[str]:
    return el.get(f"{{{W_NS}}}{attr_local}")


def _bool_present(el: Optional[etree._Element]) -> Optional[bool]:
    """
    WordprocessingML: presence often means true unless w:val is 0/false/off.
    """
    if el is None:
        return None
    v = el.get(f"{{{W_NS}}}val")
    if v is None:
        return True
    if v in ("0", "false", "off"):
        return False
    if v in ("1", "true", "on"):
        return True
    return True


def _bool_from_attr(val: Optional[str]) -> Optional[bool]:
    if val is None:
        return None
    if val in ("0", "false", "off"):
        return False
    if val in ("1", "true", "on"):
        return True
    return True


def _map_line_rule(val: Optional[str]) -> Optional[str]:
    if val is None:
        return None
    mapping = {
        "auto": "AUTO",
        "atLeast": "AT_LEAST",
        "exact": "EXACT",
        "AUTO": "AUTO",
        "AT_LEAST": "AT_LEAST",
        "EXACT": "EXACT",
    }
    return mapping.get(val)


def _merge(a: Dict[str, Any], b: Dict[str, Any]) -> Dict[str, Any]:
    """Shallow merge where b overrides a; skips None values in b."""
    out = dict(a)
    for k, v in b.items():
        if v is None:
            continue
        out[k] = v
    return out


def _stable_json_key(obj: Any) -> str:
    """Deterministic key for style de-duplication."""
    return json.dumps(obj, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def _dict_diff(base: Dict[str, Any], cur: Dict[str, Any]) -> Dict[str, Any]:
    """
    Compute shallow diff cur vs base.
    - If a key exists in cur and differs from base => include in diff.
    - If base has a nested dict and cur has dict => recurse one level (only for dicts).
    - Deletions are NOT represented (schema has no way), so absent keys in cur are ignored.
    """
    diff: Dict[str, Any] = {}
    for k, v in cur.items():
        if v is None:
            continue
        if k not in base:
            diff[k] = v
            continue
        bv = base.get(k)
        if isinstance(v, dict) and isinstance(bv, dict):
            sub = _dict_diff(bv, v)
            if sub:
                diff[k] = sub
        else:
            if v != bv:
                diff[k] = v
    return diff


def _sym_encode(font: str, char: str) -> str:
    """
    Stable minimal encoding for w:sym compatible with current schema (text field only).
    Avoid JSON-in-JSON. Escapes ';' and '=' and backslash.
    """
    def esc(s: str) -> str:
        return s.replace("\\", "\\\\").replace(";", "\\;").replace("=", "\\=")
    return f"font={esc(font)};char={esc(char)}"


@dataclass
class WordStyle:
    style_id: str
    based_on: Optional[str]
    pPr: Dict[str, Any]   # parsed p_format-like
    rPr: Dict[str, Any]   # parsed r_format-like


class UltimateParserV43:
    """
    Deterministic DOCX -> RAW JSON parser (schema v2.9).
    Subset target: "forms" (no tables/images/fields/hyperlinks). Non-run nodes in paragraphs
    are ignored (proofErr etc.).
    """

    def __init__(self, docx_path: str):
        self.docx_path = docx_path

        with zipfile.ZipFile(docx_path, "r") as z:
            self._docx_xml_parts: Dict[str, bytes] = {
                name: z.read(name)
                for name in z.namelist()
                if name.endswith(".xml")
            }

            self.document_xml = etree.fromstring(z.read("word/document.xml"))

            try:
                self.styles_xml = etree.fromstring(z.read("word/styles.xml"))
            except KeyError:
                self.styles_xml = None

            try:
                self.numbering_xml = etree.fromstring(z.read("word/numbering.xml"))
            except KeyError:
                self.numbering_xml = None

            try:
                self.settings_xml = etree.fromstring(z.read("word/settings.xml"))
            except KeyError:
                self.settings_xml = None

        # Output style library (schema.styles)
        self.out_styles: Dict[str, Dict[str, Any]] = {}
        self._style_key_to_id: Dict[str, str] = {}
        self._style_counter: int = 1

        # Output character styles library (schema.character_styles)
        self.out_char_styles: Dict[str, Dict[str, Any]] = {}
        self._char_key_to_id: Dict[str, str] = {}
        self._char_counter: int = 1

        # Word styles (for effective formatting)
        self.doc_defaults_r: Dict[str, Any] = {}
        self.doc_defaults_p: Dict[str, Any] = {}
        self.word_styles: Dict[str, WordStyle] = {}
        self.default_paragraph_style_id: Optional[str] = None

        self._init_word_styles()

        # Compute Normal effective formatting (will be set in process)
        self.normal_p_format: Dict[str, Any] = {}
        self.normal_r_format: Dict[str, Any] = {}

    # =========================
    # WORD STYLES INIT (docDefaults + style chain)
    # =========================

    def _init_word_styles(self) -> None:
        if self.styles_xml is None:
            self.doc_defaults_r = {}
            self.doc_defaults_p = {}
            self.word_styles = {}
            self.default_paragraph_style_id = None
            return

        # docDefaults
        docDefaults = self.styles_xml.find(qn("w:docDefaults"))
        if docDefaults is not None:
            rDef = docDefaults.find(qn("w:rPrDefault"))
            if rDef is not None:
                rPr = rDef.find(qn("w:rPr"))
                self.doc_defaults_r = self._parse_rPr(rPr)

            pDef = docDefaults.find(qn("w:pPrDefault"))
            if pDef is not None:
                pPr = pDef.find(qn("w:pPr"))
                self.doc_defaults_p = self._parse_pPr(pPr)

        # Styles (paragraph only)
        for st in self.styles_xml.findall(qn("w:style")):
            st_type = _str_attr(st, "type")
            st_id = _str_attr(st, "styleId")
            is_default = _str_attr(st, "default")

            if st_id is None or st_type != "paragraph":
                continue

            basedOn = st.find(qn("w:basedOn"))
            based_id = basedOn.get(f"{{{W_NS}}}val") if basedOn is not None else None

            pPr = st.find(qn("w:pPr"))
            rPr = st.find(qn("w:rPr"))

            ws = WordStyle(
                style_id=st_id,
                based_on=based_id,
                pPr=self._parse_pPr(pPr),
                rPr=self._parse_rPr(rPr),
            )
            self.word_styles[st_id] = ws

            if is_default in ("1", "true"):
                self.default_paragraph_style_id = st_id

    def _get_default_word_paragraph_style_id(self) -> Optional[str]:
        if self.styles_xml is not None:
            for st in self.styles_xml.findall(qn("w:style")):
                st_type = _str_attr(st, "type")
                st_id = _str_attr(st, "styleId")
                is_default = _str_attr(st, "default")
                if st_type == "paragraph" and st_id and is_default in ("1", "true"):
                    return st_id
        return self.default_paragraph_style_id

    def _get_p_style_id(self, pPr: Optional[etree._Element]) -> Optional[str]:
        if pPr is None:
            return self.default_paragraph_style_id
        pStyle = pPr.find(qn("w:pStyle"))
        if pStyle is None:
            return self.default_paragraph_style_id
        val = pStyle.get(f"{{{W_NS}}}val")
        return val or self.default_paragraph_style_id

    def _effective_p_format(self, style_id: Optional[str], direct_pPr: Optional[etree._Element]) -> Dict[str, Any]:
        base = dict(self.doc_defaults_p)

        # style chain root -> leaf
        for sid in self._style_chain(style_id):
            ws = self.word_styles.get(sid)
            if ws is not None:
                base = _merge(base, ws.pPr)

        # direct pPr overrides
        direct = self._parse_pPr(direct_pPr, include_indent_origin=True)
        base = _merge(base, direct)

        # Style-linked numbering: if numbering.xml lvl has <w:pStyle>, Word applies numbering even
        # when paragraph has no explicit <w:numPr>. Materialize it deterministically from donor.
        if "numbering" not in base and style_id is not None:
            implied = getattr(self, "_implied_numpr_by_word_style", {}).get(style_id)
            if implied is not None and "numbering" not in direct:
                base["numbering"] = implied

        return base

    def _build_implied_numpr_by_word_style(self, numbering_definitions: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
        """
        Build mapping wordStyleId -> implied numPr ({numId, ilvl}) from numbering.xml lvl/pStyle links.
        Deterministic tie-break: minimal numeric numId, then minimal ilvl; non-numeric numId sorted last.
        """
        tmp: Dict[str, List[Tuple[str, int]]] = {}

        def _num_key(n: str) -> Tuple[int, str]:
            try:
                return (int(n), n)
            except Exception:
                return (10**9, n)

        for numId, rec in (numbering_definitions or {}).items():
            if not isinstance(rec, dict):
                continue
            levels = rec.get("levels") or {}
            if not isinstance(levels, dict):
                continue
            for ilvl_str, lvl in levels.items():
                if not isinstance(lvl, dict):
                    continue
                pStyle = lvl.get("pStyle")
                if not isinstance(pStyle, str) or not pStyle:
                    continue
                try:
                    ilvl_int = int(ilvl_str)
                except Exception:
                    continue
                tmp.setdefault(pStyle, []).append((str(numId), ilvl_int))

        out: Dict[str, Dict[str, Any]] = {}
        for style_id, pairs in tmp.items():
            pairs_sorted = sorted(pairs, key=lambda x: (_num_key(x[0]), x[1]))
            numId, ilvl = pairs_sorted[0]
            out[style_id] = {"numId": numId, "ilvl": ilvl}

        return out

    def _effective_r_format(self, style_id: Optional[str]) -> Dict[str, Any]:
        base = dict(self.doc_defaults_r)
        for sid in self._style_chain(style_id):
            ws = self.word_styles.get(sid)
            if ws is not None:
                base = _merge(base, ws.rPr)
        return base

    def _style_chain(self, style_id: Optional[str]) -> List[str]:
        """
        Returns list from root->leaf (basedOn first, then child),
        deterministic and cycle-safe.
        """
        if style_id is None:
            return []
        visited = set()
        stack: List[str] = []
        cur = style_id
        while cur and cur not in visited:
            visited.add(cur)
            stack.append(cur)
            ws = self.word_styles.get(cur)
            cur = ws.based_on if ws else None
        return list(reversed(stack))

    # =========================
    # OUTPUT STYLE LIBRARY (schema.styles and schema.character_styles)
    # =========================

    def _register_out_style(
        self,
        p_diff: Dict[str, Any],
        r_diff: Dict[str, Any],
        source_word_style_id: Optional[str] = None
    ) -> str:
        style_obj = {"p_format": p_diff, "r_format": r_diff}

        # IMPORTANT: de-dup key must remain based ONLY on formatting (no metadata).
        key = _stable_json_key(style_obj)
        existing = self._style_key_to_id.get(key)
        if existing:
            if source_word_style_id is not None:
                cur = self.out_styles.get(existing, {})
                # keep first seen for determinism; do not override
                if isinstance(cur, dict) and cur.get("source_word_style_id") is None:
                    cur["source_word_style_id"] = source_word_style_id
            return existing

        style_id = f"s{self._style_counter:04d}"
        self._style_counter += 1
        self._style_key_to_id[key] = style_id
        if source_word_style_id is not None:
            style_obj["source_word_style_id"] = source_word_style_id
        self.out_styles[style_id] = style_obj
        return style_id

    def _register_char_style(self, r_format: Dict[str, Any]) -> str:
        key = _stable_json_key(r_format)
        existing = self._char_key_to_id.get(key)
        if existing:
            return existing

        style_id = f"c{self._char_counter:04d}"
        self._char_counter += 1
        self._char_key_to_id[key] = style_id
        self.out_char_styles[style_id] = {"r_format": r_format}
        return style_id

    # =========================
    # PUBLIC
    # =========================

    def process(self) -> str:
        # Compute Normal effective formatting
        normal_word_style_id = self._get_default_word_paragraph_style_id()  # usually "a"
        if normal_word_style_id is None:
            # Fallback: use built-in Normal defaults
            self.normal_p_format = dict(NORMAL_P_FORMAT)
            self.normal_r_format = dict(NORMAL_R_FORMAT)
        else:
            self.normal_p_format = self._effective_p_format(normal_word_style_id, None)
            self.normal_r_format = self._effective_r_format(normal_word_style_id)

        settings: Dict[str, Any] = {}
        default_tab_stop = self._parse_default_tab_stop()
        if default_tab_stop is not None:
            settings["defaultTabStopTwip"] = default_tab_stop

        numbering_definitions = self._parse_numbering_definitions()
        self._implied_numpr_by_word_style = self._build_implied_numpr_by_word_style(numbering_definitions)

        result: Dict[str, Any] = {
            "meta": {
                "schema_version": "2.9",
                "rules_version": "0.3",
                "producer": {
                    "name": "UltimateParserV43",
                    "version": "v43"
                }
            },
            "document_info": {
                "page_setup": self._parse_page_setup(),
                "settings": settings
            },
            "numbering_definitions": numbering_definitions,
            "styles": {},
            "character_styles": {},
            "content": []
        }

        body = self.document_xml.find(qn("w:body"))
        default_word_style_id = self._get_default_word_paragraph_style_id()
        paragraphs_count = 0
        runs_count = 0

        if body is None:
            if default_word_style_id is not None:
                default_p_format = self._effective_p_format(default_word_style_id, None)
                default_r_format = self._effective_r_format(default_word_style_id)
                p_diff = _dict_diff(self.normal_p_format, default_p_format)
                r_diff = _dict_diff(self.normal_r_format, default_r_format)
                default_style_id = self._register_out_style(p_diff, r_diff, source_word_style_id=default_word_style_id)
                result["meta"]["default_style_id"] = default_style_id
            self._finalize_styles(result)
            result["styles"] = self.out_styles
            result["character_styles"] = self.out_char_styles
            print(f"[parser] summary default_style_id={result['meta'].get('default_style_id')} "
                  f"styles_count={len(result['styles'])} char_styles_count={len(result['character_styles'])}")
            return json.dumps(result, ensure_ascii=False, indent=2)

        for p in body.findall(qn("w:p")):
            paragraphs_count += 1
            pPr = p.find(qn("w:pPr"))
            p_style_id = self._get_p_style_id(pPr)

            # Effective paragraph formatting: docDefaults + style chain + direct pPr
            eff_p = self._effective_p_format(p_style_id, pPr)

            # Effective run formatting: docDefaults + style chain
            eff_r_paragraph = self._effective_r_format(p_style_id)

            # Paragraph mark rPr (direct pPr/rPr) – apply only for empty paragraphs (RULE-006)
            para_mark_rPr = self._parse_rPr(pPr.find(qn("w:rPr")) if pPr is not None else None)

            # Compute diff from Normal for this paragraph
            p_diff = _dict_diff(self.normal_p_format, eff_p)
            r_diff_paragraph = _dict_diff(self.normal_r_format, eff_r_paragraph)

            runs = self._parse_runs(p, eff_r_paragraph)
            runs_count += len(runs)

            # RULE-006: empty paragraph -> include para_mark_rPr in r_format
            if not runs:
                full_r = _merge(eff_r_paragraph, para_mark_rPr)
                r_diff_paragraph = _dict_diff(self.normal_r_format, full_r)

            style_id = self._register_out_style(p_diff, r_diff_paragraph, source_word_style_id=p_style_id)

            item: Dict[str, Any] = {"style_id": style_id, "runs": runs}
            if p_style_id is not None:
                item["source_word_style_id"] = p_style_id
            result["content"].append(item)

        if default_word_style_id is not None:
            default_p_format = self._effective_p_format(default_word_style_id, None)
            default_r_format = self._effective_r_format(default_word_style_id)
            p_diff = _dict_diff(self.normal_p_format, default_p_format)
            r_diff = _dict_diff(self.normal_r_format, default_r_format)
            default_style_id = self._register_out_style(p_diff, r_diff, source_word_style_id=default_word_style_id)
            result["meta"]["default_style_id"] = default_style_id

        # Remap numbering pStyle to our word_style_id
        self._remap_numbering_pstyles(numbering_definitions)

        # Finalize style metadata
        self._finalize_styles(result)
        result["styles"] = self.out_styles
        result["character_styles"] = self.out_char_styles

        print(f"[parser] summary default_style_id={result['meta'].get('default_style_id')} "
              f"styles_count={len(result['styles'])} char_styles_count={len(result['character_styles'])} "
              f"paragraphs_count={paragraphs_count} runs_count={runs_count}")
        return json.dumps(result, ensure_ascii=False, indent=2)

    def _finalize_styles(self, result: Dict[str, Any]) -> None:
        """Assign titles and word_style_id for all styles."""
        para_ids = sorted(self.out_styles.keys())
        char_ids = sorted(self.out_char_styles.keys())

        # Titles for paragraph styles
        default_id = result["meta"].get("default_style_id")
        if default_id and default_id in self.out_styles:
            self.out_styles[default_id]["title"] = "Обычный"

        # Map source_word_style_id to localized heading names
        heading_map = {
            "1": "Заголовок 1",
            "2": "Заголовок 2",
            "3": "Заголовок 3",
            "4": "Заголовок 4",
            "5": "Заголовок 5",
            "6": "Заголовок 6",
            "7": "Заголовок 7",
            "8": "Заголовок 8",
            "9": "Заголовок 9",
        }

        # First pass: assign titles based on source_word_style_id for headings
        for sid in para_ids:
            if sid == default_id:
                continue
            st = self.out_styles.get(sid, {})
            if st.get("title") is not None:
                continue
            src = st.get("source_word_style_id")
            if src in heading_map:
                st["title"] = heading_map[src]

        # Second pass: assign "Стиль N" to remaining styles without title
        n = 1
        for sid in para_ids:
            if sid == default_id:
                continue
            st = self.out_styles.get(sid, {})
            if st.get("title") is None:
                st["title"] = f"Стиль {n}"
                n += 1

        # Titles for character styles: "Символьный стиль N"
        m = 1
        for cid in char_ids:
            st = self.out_char_styles.get(cid, {})
            if st.get("title") is None:
                st["title"] = f"Символьный стиль {m}"
                m += 1

        # word_style_id for paragraph styles:
        # Use source_word_style_id if available and not taken, else generate TF_xxx
        used_ids = set()
        for sid, st in self.out_styles.items():
            src = st.get("source_word_style_id")
            if src and src not in used_ids and src not in ("a", "Normal"):
                st["word_style_id"] = src
                used_ids.add(src)
            # else will be assigned later

        for sid, st in self.out_styles.items():
            if "word_style_id" not in st:
                candidate = f"TF_{sid}"
                # avoid collision (very unlikely)
                while candidate in used_ids:
                    candidate = f"TF_{sid}_{self._style_counter}"
                st["word_style_id"] = candidate
                used_ids.add(candidate)

        # word_style_id for character styles: TF_cxxxx
        for cid, st in self.out_char_styles.items():
            candidate = f"TF_{cid}"
            st["word_style_id"] = candidate
            # no need to track collisions as they are separate

    def _remap_numbering_pstyles(self, numbering_definitions: Dict[str, Any]) -> None:
        """
        Replace pStyle values in numbering_definitions with the corresponding word_style_id.
        If a pStyle refers to a style not in our map, create a minimal paragraph style for it.
        """
        # First, collect all pStyle values from numbering
        pstyle_vals = set()
        for num_rec in numbering_definitions.values():
            levels = num_rec.get("levels", {})
            for lvl_rec in levels.values():
                ps = lvl_rec.get("pStyle")
                if ps and isinstance(ps, str):
                    pstyle_vals.add(ps)

        # Ensure each pstyle has a corresponding paragraph style
        for ps in pstyle_vals:
            # Check if any style has this source_word_style_id
            found = any(st.get("source_word_style_id") == ps for st in self.out_styles.values())
            if not found:
                # Create minimal style (empty diff) with source_word_style_id=ps
                style_id = self._register_out_style({}, {}, source_word_style_id=ps)

        # Now replace pStyle in numbering_definitions with word_style_id
        for num_rec in numbering_definitions.values():
            levels = num_rec.get("levels", {})
            for lvl_rec in levels.values():
                ps = lvl_rec.get("pStyle")
                if ps and isinstance(ps, str):
                    for st in self.out_styles.values():
                        if st.get("source_word_style_id") == ps:
                            lvl_rec["pStyle"] = st.get("word_style_id", ps)
                            break

    # =========================
    # DOCUMENT INFO
    # =========================

    def _parse_page_setup(self) -> Dict[str, Any]:
        """
        From w:sectPr (body/sectPr or last p/pPr/sectPr):
          pgSz: w,h,orient
          pgMar: left,right,top,bottom,header,footer,gutter
          cols: num, space, equalWidth
        """
        body = self.document_xml.find(qn("w:body"))
        if body is None:
            return {}

        sectPr = body.find(qn("w:sectPr"))
        if sectPr is None:
            # try last paragraph's sectPr
            last_p = None
            for ch in body:
                if ch.tag == qn("w:p"):
                    last_p = ch
            if last_p is not None:
                pPr = last_p.find(qn("w:pPr"))
                if pPr is not None:
                    sectPr = pPr.find(qn("w:sectPr"))

        if sectPr is None:
            return {}

        out: Dict[str, Any] = {}

        pgSz = sectPr.find(qn("w:pgSz"))
        if pgSz is not None:
            w = _int_attr(pgSz, "w")
            h = _int_attr(pgSz, "h")
            if w is not None:
                out["pageWidthTwip"] = w
            if h is not None:
                out["pageHeightTwip"] = h
            orient = _str_attr(pgSz, "orient")
            if orient in ("portrait", "landscape"):
                out["orient"] = orient

        pgMar = sectPr.find(qn("w:pgMar"))
        if pgMar is not None:
            for src, dst in (
                ("left", "marginLeftTwip"),
                ("right", "marginRightTwip"),
                ("top", "marginTopTwip"),
                ("bottom", "marginBottomTwip"),
                ("header", "headerTwip"),
                ("footer", "footerTwip"),
                ("gutter", "gutterTwip"),
            ):
                v = _int_attr(pgMar, src)
                if v is not None:
                    out[dst] = v

        cols = sectPr.find(qn("w:cols"))
        if cols is not None:
            c: Dict[str, Any] = {}
            num = _int_attr(cols, "num")
            space = _int_attr(cols, "space")
            equal = _str_attr(cols, "equalWidth")
            if num is not None:
                c["num"] = num
            if space is not None:
                c["spaceTwip"] = space
            if equal is not None:
                c["equalWidth"] = equal not in ("0", "false", "off")
            if c:
                out["cols"] = c

        docGrid = sectPr.find(qn("w:docGrid"))
        if docGrid is not None:
            lp = _int_attr(docGrid, "linePitch")
            if lp is not None:
                out["linePitchTwip"] = lp

        return out

    def _parse_default_tab_stop(self) -> Optional[int]:
        if self.settings_xml is None:
            return None
        dts = self.settings_xml.find(qn("w:defaultTabStop"))
        if dts is None:
            return None
        return _int_attr(dts, "val")

    # =========================
    # NUMBERING (no synthetic defaults)
    # =========================

    def _parse_numbering_level_pPr(self, lvl: etree._Element) -> Dict[str, Any]:
        ppr = lvl.find(qn("w:pPr"))
        if ppr is None:
            return {}

        out: Dict[str, Any] = {}

        ind = ppr.find(qn("w:ind"))
        if ind is not None:
            left = _int_attr(ind, "left")
            right = _int_attr(ind, "right")
            first = _int_attr(ind, "firstLine")
            hanging = _int_attr(ind, "hanging")
            if left is not None:
                out["indentStartTwip"] = left
            if right is not None:
                out["indentEndTwip"] = right
            if first is not None:
                out["indentFirstLineTwip"] = first
            if hanging is not None:
                out["indentHangingTwip"] = hanging

        tabs = ppr.find(qn("w:tabs"))
        if tabs is not None:
            arr: List[Dict[str, Any]] = []
            for t in tabs.findall(qn("w:tab")):
                pos = _int_attr(t, "pos")
                val = _str_attr(t, "val")
                leader = _str_attr(t, "leader")
                if pos is None or val is None:
                    continue
                rec: Dict[str, Any] = {"posTwip": pos, "val": val}
                if leader is not None:
                    rec["leader"] = leader
                arr.append(rec)
            if arr:
                out["tabs"] = arr

        return out

    def _parse_numbering_definitions(self) -> Dict[str, Any]:
        if self.numbering_xml is None:
            return {}

        # abstractNumId -> levels
        abstracts: Dict[str, Dict[str, Any]] = {}
        for absn in self.numbering_xml.findall(qn("w:abstractNum")):
            abs_id = _str_attr(absn, "abstractNumId")
            if abs_id is None:
                continue

            levels: Dict[str, Any] = {}
            for lvl in absn.findall(qn("w:lvl")):
                ilvl = _int_attr(lvl, "ilvl")
                if ilvl is None:
                    continue

                numFmt = lvl.find(qn("w:numFmt"))
                lvlText = lvl.find(qn("w:lvlText"))
                start = lvl.find(qn("w:start"))

                fmt = numFmt.get(f"{{{W_NS}}}val") if numFmt is not None else None
                template = lvlText.get(f"{{{W_NS}}}val") if lvlText is not None else None
                st = _int_attr(start, "val") if start is not None else None

                # No synthetic defaults: include only if required parts exist
                if fmt is None or template is None:
                    continue

                level_rec: Dict[str, Any] = {
                    "format": fmt,
                    "template": template,
                }
                if st is not None:
                    level_rec["start"] = st

                lvl_tab = lvl.find(qn("w:tab"))
                if lvl_tab is not None:
                    tab_pos = _int_attr(lvl_tab, "val")
                    if tab_pos is not None:
                        level_rec["tabPosTwip"] = tab_pos

                lvl_jc_el = lvl.find(qn("w:lvlJc"))
                if lvl_jc_el is not None:
                    lvl_jc = _str_attr(lvl_jc_el, "val")
                    if lvl_jc is not None:
                        level_rec["lvlJc"] = lvl_jc

                suff_el = lvl.find(qn("w:suff"))
                if suff_el is not None:
                    suff = _str_attr(suff_el, "val")
                    if suff is not None:
                        level_rec["suff"] = suff

                p_style_el = lvl.find(qn("w:pStyle"))
                if p_style_el is not None:
                    p_style = _str_attr(p_style_el, "val")
                    if p_style is not None:
                        level_rec["pStyle"] = p_style

                level_ppr = self._parse_numbering_level_pPr(lvl)
                if level_ppr:
                    level_rec["level_pPr"] = level_ppr

                lvl_rpr = self._parse_rPr(lvl.find(qn("w:rPr")))
                if lvl_rpr:
                    level_rec["level_rPr"] = lvl_rpr

                levels[str(ilvl)] = level_rec

            abs_rec: Dict[str, Any] = {"levels": levels}
            mlt_el = absn.find(qn("w:multiLevelType"))
            if mlt_el is not None:
                mlt = _str_attr(mlt_el, "val")
                if mlt is not None:
                    abs_rec["multiLevelType"] = mlt
            abstracts[abs_id] = abs_rec

        # numId -> abstractNumId + overrides
        out: Dict[str, Any] = {}
        for num in self.numbering_xml.findall(qn("w:num")):
            numId = _str_attr(num, "numId")
            if numId is None:
                continue

            absRef = num.find(qn("w:abstractNumId"))
            abs_id = absRef.get(f"{{{W_NS}}}val") if absRef is not None else None

            levels = {}
            if abs_id is not None and abs_id in abstracts:
                levels = abstracts[abs_id]["levels"]

            lvl_overrides: Dict[str, Any] = {}
            for ov in num.findall(qn("w:lvlOverride")):
                ilvl = _int_attr(ov, "ilvl")
                if ilvl is None:
                    continue
                startOv = ov.find(qn("w:startOverride"))
                if startOv is not None:
                    st = _int_attr(startOv, "val")
                    if st is not None:
                        lvl_overrides[str(ilvl)] = {"start": st}

            rec: Dict[str, Any] = {"levels": levels}
            if abs_id is not None:
                rec["abstractNumId"] = abs_id
                mlt = abstracts.get(abs_id, {}).get("multiLevelType")
                if isinstance(mlt, str):
                    rec["multiLevelType"] = mlt
            if lvl_overrides:
                rec["lvl_overrides"] = lvl_overrides

            out[numId] = rec

        return out

    # =========================
    # PARAGRAPH FORMATTING (pPr -> pFormat)
    # =========================

    def _parse_pPr(self, pPr: Optional[etree._Element], include_indent_origin: bool = False) -> Dict[str, Any]:
        """
        Parse paragraph properties (w:pPr) -> p_format (schema).
        Goal: visually-lossless. IMPORTANT: keep explicit zeros (0) if present in XML.
        """
        if pPr is None:
            return {}

        out: Dict[str, Any] = {}

        # alignment (w:jc)
        jc = pPr.find(qn("w:jc"))
        if jc is not None:
            val = jc.get(f"{{{W_NS}}}val")
            if val:
                v = val.upper()
                if v == "LEFT":
                    out["alignment"] = "LEFT"
                elif v == "CENTER":
                    out["alignment"] = "CENTER"
                elif v == "RIGHT":
                    out["alignment"] = "RIGHT"
                elif v in ("BOTH", "JUSTIFY"):
                    out["alignment"] = "JUSTIFY"

        # indents (w:ind)
        ind = pPr.find(qn("w:ind"))
        if ind is not None:
            left = _int_attr(ind, "left")
            right = _int_attr(ind, "right")
            first = _int_attr(ind, "firstLine")
            hanging = _int_attr(ind, "hanging")

            if left is not None:
                out["indentStartTwip"] = left
                if include_indent_origin:
                    out["indentStartTwipOrigin"] = "direct"
            if right is not None:
                out["indentEndTwip"] = right
                if include_indent_origin:
                    out["indentEndTwipOrigin"] = "direct"
            if first is not None:
                out["indentFirstLineTwip"] = first
                if include_indent_origin:
                    out["indentFirstLineTwipOrigin"] = "direct"
            if hanging is not None:
                out["indentHangingTwip"] = hanging  # RULE-001
                if include_indent_origin:
                    out["indentHangingTwipOrigin"] = "direct"

        # spacing (w:spacing)  <-- keep explicit zeros
        spacing = pPr.find(qn("w:spacing"))
        if spacing is not None:
            before = _int_attr(spacing, "before")
            after = _int_attr(spacing, "after")
            beforeLines = _int_attr(spacing, "beforeLines")
            afterLines = _int_attr(spacing, "afterLines")
            line = _int_attr(spacing, "line")
            lineRule = _map_line_rule(_str_attr(spacing, "lineRule"))

            # FIX: autospacing are ATTRIBUTES on w:spacing
            beforeAuto = _bool_from_attr(spacing.get(f"{{{W_NS}}}beforeAutospacing"))
            afterAuto = _bool_from_attr(spacing.get(f"{{{W_NS}}}afterAutospacing"))

            if before is not None:
                out["spaceBeforeTwip"] = before
            if after is not None:
                out["spaceAfterTwip"] = after
            if beforeLines is not None:
                out["spaceBeforeLines"] = beforeLines
            if afterLines is not None:
                out["spaceAfterLines"] = afterLines

            if line is not None:
                out["lineTwip"] = line
            if lineRule is not None:
                out["lineRule"] = lineRule

            if beforeAuto is not None:
                out["beforeAutospacing"] = beforeAuto
            if afterAuto is not None:
                out["afterAutospacing"] = afterAuto

        # tabs (w:tabs)
        tabs = pPr.find(qn("w:tabs"))
        if tabs is not None:
            arr: List[Dict[str, Any]] = []
            for t in tabs.findall(qn("w:tab")):
                pos = _int_attr(t, "pos")
                val = _str_attr(t, "val")
                leader = _str_attr(t, "leader")
                if pos is None or val is None:
                    continue
                rec: Dict[str, Any] = {"posTwip": pos, "val": val}
                if leader is not None:
                    rec["leader"] = leader
                arr.append(rec)
            if arr:
                out["tabs"] = arr

        # numbering (w:numPr)
        numPr = pPr.find(qn("w:numPr"))
        if numPr is not None:
            ilvl = numPr.find(qn("w:ilvl"))
            numId = numPr.find(qn("w:numId"))
            if ilvl is not None and numId is not None:
                ilvl_val = _int_attr(ilvl, "val")
                num_val = _str_attr(numId, "val")
                if ilvl_val is not None and num_val is not None:
                    out["numbering"] = {"numId": num_val, "ilvl": ilvl_val}

        # keep/widow/etc
        kn = pPr.find(qn("w:keepNext"))
        if kn is not None:
            out["keepNext"] = _bool_present(kn)

        kl = pPr.find(qn("w:keepLines"))
        if kl is not None:
            out["keepLines"] = _bool_present(kl)

        pbb = pPr.find(qn("w:pageBreakBefore"))
        if pbb is not None:
            out["pageBreakBefore"] = _bool_present(pbb)

        wc = pPr.find(qn("w:widowControl"))
        if wc is not None:
            out["widowControl"] = _bool_present(wc)

        stg = pPr.find(qn("w:snapToGrid"))
        if stg is not None:
            out["snapToGrid"] = _bool_present(stg)

        cs = pPr.find(qn("w:contextualSpacing"))
        if cs is not None:
            out["contextualSpacing"] = _bool_present(cs)

        ta = pPr.find(qn("w:textAlignment"))
        if ta is not None:
            v = _str_attr(ta, "val")
            text_align_map = {
                "auto": "AUTO",
                "baseline": "BASELINE",
                "top": "TOP",
                "center": "CENTER",
                "bottom": "BOTTOM",
                "AUTO": "AUTO",
                "BASELINE": "BASELINE",
                "TOP": "TOP",
                "CENTER": "CENTER",
                "BOTTOM": "BOTTOM",
            }
            mapped = text_align_map.get(v) if v is not None else None
            if mapped is not None:
                out["textAlignment"] = mapped

        return out

    # =========================
    # RUN FORMATTING (rPr -> rFormat)
    # =========================

    def _parse_rPr(self, rPr: Optional[etree._Element]) -> Dict[str, Any]:
        if rPr is None:
            return {}

        out: Dict[str, Any] = {}

        # rFonts
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is not None:
            rf: Dict[str, Any] = {}
            for k in ("ascii", "hAnsi", "eastAsia", "cs",
                      "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "csTheme"):
                v = rFonts.get(f"{{{W_NS}}}{k}")
                if v is not None:
                    rf[k] = v
            if rf:
                out["rFonts"] = rf

        # size
        sz = rPr.find(qn("w:sz"))
        if sz is not None:
            v = _int_attr(sz, "val")
            if v is not None:
                out["font_size_half_points"] = v

        # bold/italic
        b = rPr.find(qn("w:b"))
        if b is not None:
            out["bold"] = _bool_present(b)
        i = rPr.find(qn("w:i"))
        if i is not None:
            out["italic"] = _bool_present(i)

        # underline
        u = rPr.find(qn("w:u"))
        if u is not None:
            v = _str_attr(u, "val")
            if v is not None:
                out["underline"] = v

        # color
        c = rPr.find(qn("w:color"))
        if c is not None:
            v = _str_attr(c, "val")
            if v is not None:
                out["color"] = v

        # vertical align
        va = rPr.find(qn("w:vertAlign"))
        if va is not None:
            v = _str_attr(va, "val")
            if v in ("baseline", "superscript", "subscript"):
                out["vertical_align"] = v

        # all caps
        caps = rPr.find(qn("w:caps"))
        if caps is not None:
            out["all_caps"] = _bool_present(caps)

        # lang
        lang = rPr.find(qn("w:lang"))
        if lang is not None:
            rec: Dict[str, Any] = {}
            for k in ("val", "eastAsia", "bidi"):
                v = lang.get(f"{{{W_NS}}}{k}")
                if v is not None:
                    rec[k] = v
            if rec:
                out["lang"] = rec

        # ADD BACK: character spacing (w:spacing) and position (w:position)
        sp = rPr.find(qn("w:spacing"))
        if sp is not None:
            v = _int_attr(sp, "val")
            if v is not None:
                out["charSpacingTwip"] = v

        pos = rPr.find(qn("w:position"))
        if pos is not None:
            v = _int_attr(pos, "val")
            if v is not None:
                out["positionHalfPoints"] = v

        return out

    # =========================
    # RUNS PARSING (w:r children) -> schema runs[]
    # =========================

    def _parse_runs(self, p: etree._Element, base_r_for_diff: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Parse runs and create character styles.
        Instead of diff, we record char_style_id.
        base_r_for_diff is the effective rPr of the paragraph (including Normal + style).
        """
        out: List[Dict[str, Any]] = []
        first_emitted = False  # for meta.leading on first tab

        for child in p:
            if child.tag != qn("w:r"):
                continue

            rPr = child.find(qn("w:rPr"))
            run_local_r = self._parse_rPr(rPr)   # direct run formatting

            for node in child:
                if node.tag == qn("w:rPr"):
                    continue

                if node.tag == qn("w:t"):
                    txt = node.text or ""
                    preserve = node.get(f"{{{XML_NS}}}space") == "preserve"

                    run_obj: Dict[str, Any] = {"type": "text", "text": txt}
                    if run_local_r:
                        char_style_id = self._register_char_style(run_local_r)
                        run_obj["char_style_id"] = char_style_id
                    if preserve:
                        run_obj["meta"] = {"preserve": True}

                    out.append(run_obj)
                    if not first_emitted:
                        first_emitted = True

                elif node.tag == qn("w:tab"):
                    run_obj: Dict[str, Any] = {"type": "tab"}
                    if run_local_r:
                        char_style_id = self._register_char_style(run_local_r)
                        run_obj["char_style_id"] = char_style_id

                    # ADD: first visual token is a tab -> meta.leading=true
                    if not first_emitted:
                        run_obj["meta"] = {"leading": True}
                        first_emitted = True

                    out.append(run_obj)

                elif node.tag == qn("w:br"):
                    run_obj: Dict[str, Any] = {"type": "break"}
                    br_type = node.get(f"{{{W_NS}}}type")
                    if br_type in ("textWrapping", "page", "column"):
                        run_obj["break_type"] = br_type
                    if run_local_r:
                        char_style_id = self._register_char_style(run_local_r)
                        run_obj["char_style_id"] = char_style_id
                    out.append(run_obj)
                    first_emitted = True

                elif node.tag == qn("w:cr"):
                    run_obj: Dict[str, Any] = {"type": "cr"}
                    if run_local_r:
                        char_style_id = self._register_char_style(run_local_r)
                        run_obj["char_style_id"] = char_style_id
                    out.append(run_obj)
                    first_emitted = True

                elif node.tag == qn("w:sym"):
                    font = node.get(f"{{{W_NS}}}font") or ""
                    char = node.get(f"{{{W_NS}}}char") or ""
                    run_obj: Dict[str, Any] = {"type": "sym", "text": _sym_encode(font, char)}
                    if run_local_r:
                        char_style_id = self._register_char_style(run_local_r)
                        run_obj["char_style_id"] = char_style_id
                    out.append(run_obj)
                    first_emitted = True

                else:
                    # ignore unsupported nodes for "forms" subset
                    continue

        return out

    def dump_donor_xml_parts(self, out_root_dir: str) -> None:
        os.makedirs(out_root_dir, exist_ok=True)
        for name, data in sorted(self._docx_xml_parts.items()):
            out_path = os.path.join(out_root_dir, name)
            os.makedirs(os.path.dirname(out_path), exist_ok=True)
            with open(out_path, "wb") as f:
                f.write(data)


if __name__ == "__main__":
    try:
        cli = argparse.ArgumentParser(description="UltimateParserV43 DOCX -> RAW JSON")
        cli.add_argument("--in", dest="input_docx", default=os.path.join(BASE_DIR, "donor_v2.6.docx"))
        cli.add_argument("--out", dest="out_json", default=os.path.join(BASE_DIR, "donor_v2.6.json"))
        args = cli.parse_args()

        parser = UltimateParserV43(args.input_docx)

        os.makedirs(os.path.dirname(args.out_json) or ".", exist_ok=True)
        with open(args.out_json, "w", encoding="utf-8") as f:
            f.write(parser.process())

        print("Парсинг v43 (с символьными стилями) завершен успешно!")
    except Exception:
        import traceback
        traceback.print_exc()