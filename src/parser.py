# UltimateParserV43.py
# Schema v2.14 parser: imports styles.xml into JSON styles library; uses p_style_id/r_style_id + inline r_format; numbering + pictures preserved; legacy synthetic styles removed.
# Parser Version: v43
# Schema Version: 2.15
# Rules Version: 0.3

# Deterministic, visually-lossless for "forms" subset (no tables/images/fields/hyperlinks).

from __future__ import annotations

import argparse
import json
import os
import zipfile
from typing import Any, Dict, List, Optional, Tuple

from lxml import etree
import sys
sys.path.insert(0, os.path.dirname(__file__))
from parser_picture import parse_picture_node
from parser_table import parse_table_node

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
MY_NS = "https://translatefactory/schema/custom-id"
BASE_DIR = "."

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
    # raw.schema.json v2.15 enum: "auto" | "atLeast" | "exact"
    if val in ("auto", "atLeast", "exact"):
        return val
    # tolerate legacy variants if they appear (should not be emitted further)
    legacy = {
        "AUTO": "auto",
        "AT_LEAST": "atLeast",
        "EXACT": "exact",
    }
    return legacy.get(val)


def _require_my_id(el: etree._Element, what: str) -> str:
    v = el.get(f"{{{MY_NS}}}id")
    if v is None or not str(v).strip():
        raise ValueError(f"Missing required my:id for {what}")
    return str(v).strip()


def _sym_encode(font: str, char: str) -> str:
    """
    Stable minimal encoding for w:sym compatible with current schema (text field only).
    Avoid JSON-in-JSON. Escapes ';' and '=' and backslash.
    """
    def esc(s: str) -> str:
        return s.replace("\\", "\\\\").replace(";", "\\;").replace("=", "\\=")
    return f"font={esc(font)};char={esc(char)}"


class UltimateParserV43:
    """
    Deterministic DOCX -> RAW JSON parser (schema v2.12).
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

            # Загрузка отношений документа (document.xml.rels)
            self.relationships = {}
            try:
                rels_data = z.read("word/_rels/document.xml.rels")
                rels_xml = etree.fromstring(rels_data)
                # Пространство имён для отношений
                rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
                for rel in rels_xml.findall(f"{{{rels_ns}}}Relationship"):
                    r_id = rel.get("Id")
                    target = rel.get("Target")
                    if r_id and target:
                        self.relationships[r_id] = target
            except KeyError:
                # Если файла отношений нет, оставляем пустой словарь
                pass

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

        # Счётчик для run_id
        self._paragraph_counter = 1  # больше не используется для корневых абзацев, но оставим для совместимости
        self._table_counter = 1
        self._row_counter = 1
        self.default_paragraph_style_id: Optional[str] = None

        self._init_word_styles()

    # =========================
    # WORD STYLES INIT
    # =========================

    def _init_word_styles(self) -> None:
        if self.styles_xml is None:
            self.default_paragraph_style_id = None
            return

        for st in self.styles_xml.findall(qn("w:style")):
            st_type = _str_attr(st, "type")
            st_id = _str_attr(st, "styleId")
            is_default = _str_attr(st, "default")

            if st_id is not None and st_type == "paragraph" and is_default in ("1", "true"):
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

    # =========================
    # PUBLIC
    # =========================

    def process(self) -> str:
        settings: Dict[str, Any] = {}
        default_tab_stop = self._parse_default_tab_stop()
        if default_tab_stop is not None:
            settings["defaultTabStopTwip"] = default_tab_stop

        numbering_definitions = self._parse_numbering_definitions()
        self._implied_numpr_by_word_style = self._build_implied_numpr_by_word_style(numbering_definitions)

        default_style_id = self._get_default_word_paragraph_style_id()

        result: Dict[str, Any] = {
            "meta": {
                "schema_version": "2.15",
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
            "doc_defaults": self._parse_doc_defaults_v212(),
            "latent_styles": self._parse_latent_styles_v212(),
            "styles": self._parse_styles_v212(),
            "content": []
        }

        if isinstance(default_style_id, str) and default_style_id:
            result["meta"]["default_style_id"] = default_style_id

        # Обходим дочерние элементы body: только корневые w:p и w:tbl.
        # Контракт v2.15: единственный источник id для корневых элементов — атрибут my:id.
        body = self.document_xml.find(qn("w:body"))
        if body is not None:
            for child in body:
                if child.tag == qn("w:p"):
                    elem_id = _require_my_id(child, what="root paragraph (w:body > w:p)")
                    result["content"].append(
                        self._parse_paragraph_element(child, parent_id=None, index=None, root_id=elem_id)
                    )
                elif child.tag == qn("w:tbl"):
                    table_id = _require_my_id(child, what="root table (w:body > w:tbl)")
                    result["content"].append(parse_table_node(self, child, table_id))
                # остальные элементы игнорируем – passthrough (sectPr, bookmarks, etc.)

        if os.getenv("DOCX_PIPELINE_VALIDATE_RAW") == "1":
            self._validate_raw_v212(result)
        return json.dumps(result, ensure_ascii=False, indent=2)

    def _get_inline_p_style_id(self, pPr: Optional[etree._Element]) -> str:
        if pPr is not None:
            pStyle = pPr.find(qn("w:pStyle"))
            if pStyle is not None:
                val = pStyle.get(f"{{{W_NS}}}val")
                if val:
                    return val
        default_style = self._get_default_word_paragraph_style_id()
        if default_style:
            return default_style
        # Если ни явный стиль, ни дефолтный не найдены, возвращаем "Normal"
        # Это соответствует поведению Word и не ломает визуальную точность.
        # При желании можно заменить на None и доработать схему.
        return "Normal"

    def _to_schema_p_format(self, old: Dict[str, Any]) -> Dict[str, Any]:
        out: Dict[str, Any] = {}
        alignment_map = {
            "LEFT": "left",
            "CENTER": "center",
            "RIGHT": "right",
            "JUSTIFY": "justify",
            "DISTRIBUTE": "distribute",
            "left": "left",
            "center": "center",
            "right": "right",
            "justify": "justify",
            "distribute": "distribute",
        }
        text_alignment_map = {
            "AUTO": "auto",
            "BASELINE": "baseline",
            "TOP": "top",
            "CENTER": "center",
            "BOTTOM": "bottom",
            "auto": "auto",
            "baseline": "baseline",
            "top": "top",
            "center": "center",
            "bottom": "bottom",
        }
        line_rule_map = {
            "AUTO": "auto",
            "AT_LEAST": "atLeast",
            "EXACT": "exact",
            "auto": "auto",
            "atLeast": "atLeast",
            "exact": "exact",
        }
        key_map = {
            # Существующие поля
            "lineTwip": "line_spacing_twip",
            "spaceBeforeTwip": "space_before_twip",
            "spaceAfterTwip": "space_after_twip",
            "indentStartTwip": "indent_start_twip",
            "indentEndTwip": "indent_end_twip",
            "indentFirstLineTwip": "indent_first_line_twip",
            "indentHangingTwip": "indent_hanging_twip",
            "keepNext": "keep_next",
            "keepLines": "keep_lines",
            "pageBreakBefore": "page_break_before",
            "widowControl": "widow_control",
            "contextualSpacing": "contextual_spacing",
            "snapToGrid": "snap_to_grid",
            "beforeAutospacing": "before_autospacing",
            "afterAutospacing": "after_autospacing",
            "spaceBeforeLines": "space_before_lines",
            "spaceAfterLines": "space_after_lines",
        }
        for k, v in old.items():
            if v is None:
                continue
            if k == "alignment":
                mapped = alignment_map.get(v)
                if mapped is not None:
                    out["alignment"] = mapped
            elif k == "textAlignment":
                mapped = text_alignment_map.get(v)
                if mapped is not None:
                    out["text_alignment"] = mapped
            elif k == "lineRule":
                mapped = line_rule_map.get(v)
                if mapped is not None:
                    out["line_rule"] = mapped
            elif k == "numbering" and isinstance(v, dict):
                num_id = v.get("numId")
                ilvl = v.get("ilvl")
                if num_id is not None and ilvl is not None:
                    out["list_info"] = {"numId": str(num_id), "ilvl": str(ilvl)}
            elif k in key_map:
                out[key_map[k]] = v
            elif k == "tabs":
                out["tabs"] = v
        return out

    def _to_schema_r_format(self, old: Dict[str, Any]) -> Dict[str, Any]:
        out: Dict[str, Any] = {}
        vert_map = {
            "baseline": "baseline",
            "superscript": "superscript",
            "subscript": "subscript",
        }
        key_map = {
            "all_caps": "caps",
            "vertical_align": "vert_align",
            "charSpacingTwip": "spacing_twip",
            "positionHalfPoints": "position_half_points",
            "font_size_half_points": "font_size_half_points",
            "rFonts": "rFonts",
            "lang": "lang",
            "bold": "bold",
            "italic": "italic",
            "underline": "underline",
            "color": "color",
        }
        for k, v in old.items():
            if v is None:
                continue
            if k == "vertical_align":
                mapped = vert_map.get(v)
                if mapped is not None:
                    out["vert_align"] = mapped
            elif k in key_map:
                out[key_map[k]] = v
        return out

    def _parse_doc_defaults_v212(self) -> Dict[str, Any]:
        if self.styles_xml is None:
            return {"p_format": {}, "r_format": {}}

        p_old: Dict[str, Any] = {}
        r_old: Dict[str, Any] = {}
        docDefaults = self.styles_xml.find(qn("w:docDefaults"))
        if docDefaults is not None:
            pDef = docDefaults.find(qn("w:pPrDefault"))
            if pDef is not None:
                p_old = self._parse_pPr(pDef.find(qn("w:pPr")))
            rDef = docDefaults.find(qn("w:rPrDefault"))
            if rDef is not None:
                r_old = self._parse_rPr(rDef.find(qn("w:rPr")))

        return {
            "p_format": self._to_schema_p_format(p_old),
            "r_format": self._to_schema_r_format(r_old),
        }

    def _parse_latent_styles_v212(self) -> Dict[str, Any]:
        if self.styles_xml is None:
            return {}
        latent = self.styles_xml.find(qn("w:latentStyles"))
        if latent is None:
            return {}

        out: Dict[str, Any] = {}
        bool_attr_map = {
            "defLockedState": "defaultLockedState",
            "defSemiHidden": "defaultSemiHiddenState",
            "defUnhideWhenUsed": "defaultUnhideWhenUsedState",
            "defQFormat": "defaultQFormatState",
        }
        for src, dst in bool_attr_map.items():
            bv = _bool_from_attr(latent.get(f"{{{W_NS}}}{src}"))
            if bv is not None:
                out[dst] = bv

        ui = _int_attr(latent, "defUIPriority")
        if ui is not None:
            out["defaultUiPriority"] = ui

        exceptions: Dict[str, Any] = {}
        for exc in latent.findall(qn("w:lsdException")):
            name = _str_attr(exc, "name")
            if not name:
                continue
            rec: Dict[str, Any] = {}
            for src, dst in (
                ("locked", "locked"),
                ("semiHidden", "semiHidden"),
                ("unhideWhenUsed", "unhideWhenUsed"),
                ("qFormat", "qFormat"),
            ):
                bv = _bool_from_attr(exc.get(f"{{{W_NS}}}{src}"))
                if bv is not None:
                    rec[dst] = bv
            pr = _int_attr(exc, "uiPriority")
            if pr is not None:
                rec["uiPriority"] = pr
            if rec:
                exceptions[name] = rec
        if exceptions:
            out["exceptions"] = exceptions
        return out

    def _parse_styles_v212(self) -> Dict[str, Any]:
        out: Dict[str, Any] = {}
        if self.styles_xml is None:
            return out

        styles_by_id: Dict[str, etree._Element] = {}
        for st in self.styles_xml.findall(qn("w:style")):
            st_id = _str_attr(st, "styleId")
            st_type = _str_attr(st, "type")
            if st_id and st_type in {"paragraph", "character", "table", "numbering"}:
                styles_by_id[st_id] = st

        for style_id in sorted(styles_by_id.keys()):
            st = styles_by_id[style_id]
            st_type = _str_attr(st, "type")
            if st_type is None:
                continue
            rec: Dict[str, Any] = {"type": st_type}

            nm = st.find(qn("w:name"))
            if nm is not None:
                name = _str_attr(nm, "val")
                if name is not None:
                    rec["name"] = name

            for tag, key in (("w:basedOn", "based_on"), ("w:next", "next"), ("w:link", "link")):
                el = st.find(qn(tag))
                if el is not None:
                    v = _str_attr(el, "val")
                    if v is not None:
                        rec[key] = v

            is_default = _bool_from_attr(st.get(f"{{{W_NS}}}default"))
            if is_default is not None:
                rec["is_default"] = is_default
            custom = _bool_from_attr(st.get(f"{{{W_NS}}}customStyle"))
            if custom is not None:
                rec["custom"] = custom

            ui_pr = st.find(qn("w:uiPriority"))
            if ui_pr is not None:
                ui_val = _int_attr(ui_pr, "val")
                if ui_val is not None:
                    rec["ui_priority"] = ui_val

            for tag, key in (
                ("w:qFormat", "q_format"),
                ("w:semiHidden", "semi_hidden"),
                ("w:unhideWhenUsed", "unhide_when_used"),
                ("w:locked", "locked"),
            ):
                el = st.find(qn(tag))
                if el is not None:
                    bv = _bool_present(el)
                    if bv is not None:
                        rec[key] = bv

            if st_type == "paragraph":
                p_format = self._to_schema_p_format(self._parse_pPr(st.find(qn("w:pPr"))))
                if p_format:
                    rec["p_format"] = p_format

            if st_type in {"paragraph", "character"}:
                r_format = self._to_schema_r_format(self._parse_rPr(st.find(qn("w:rPr"))))
                if r_format:
                    rec["r_format"] = r_format

            out[style_id] = rec

        return out

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
    # PARAGRAPH PARSING (выделено в отдельный метод)
    # =========================

    def _parse_paragraph_element(self, p: etree._Element, parent_id: Optional[str], index: Optional[int], root_id: Optional[str] = None) -> Dict[str, Any]:
        """
        Парсит один элемент <w:p> и возвращает словарь абзаца.
        Используется как для абзацев в body, так и для абзацев внутри таблиц.

        Аргументы:
            p – элемент <w:p>
            parent_id – идентификатор родителя (для вложенных абзацев)
            index – порядковый номер абзаца внутри родителя (начиная с 1) – только для вложенных
            root_id – готовый идентификатор для корневого абзаца (когда parent_id is None)
        """
        pPr = p.find(qn("w:pPr"))
        p_style_id = self._get_inline_p_style_id(pPr)

        if parent_id is not None:
            # вложенный абзац (в ячейке таблицы)
            p_id = f"{parent_id}.p_{index}"
        elif root_id is not None:
            # корневой абзац с явным id из SDT
            p_id = root_id
        else:
            # старый режим (для обратной совместимости, но в новом коде не должен использоваться)
            p_id = f"p_{self._paragraph_counter}"
            self._paragraph_counter += 1

        runs = self._parse_runs(p, p_id)  # передаём id абзаца как родителя для run'ов
        # Важно: для корневых абзацев self._paragraph_counter НЕ увеличивается

        item: Dict[str, Any] = {
            "type": "paragraph",
            "id": p_id,
            "p_style_id": p_style_id,
            "runs": runs,
        }

        p_inline = self._to_schema_p_format(self._parse_pPr(pPr, include_indent_origin=False))
        if p_inline:
            item["p_format"] = p_inline

        return item

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
            # Schema v2.12 expects rFormat.lang to be a string.
            v = lang.get(f"{{{W_NS}}}val") or lang.get(f"{{{W_NS}}}eastAsia") or lang.get(f"{{{W_NS}}}bidi")
            if v:
                out["lang"] = v

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

    def _parse_runs(self, p: etree._Element, parent_id: str) -> List[Dict[str, Any]]:
        """
        Parse runs preserving token order and inline formatting.
        parent_id – идентификатор родительского абзаца (например, "p_1")
        """
        out: List[Dict[str, Any]] = []
        first_emitted = False  # for meta.leading on first tab
        run_counter = 1  # локальный счётчик run'ов для этого абзаца

        for child in p:
            if run_counter > 1000:  # защита от бесконечного цикла (на всякий случай)
                break
            if child.tag != qn("w:r"):
                continue

            rPr = child.find(qn("w:rPr"))
            run_local_r = self._to_schema_r_format(self._parse_rPr(rPr))
            r_style_id = None
            if rPr is not None:
                r_style = rPr.find(qn("w:rStyle"))
                if r_style is not None:
                    r_style_id = _str_attr(r_style, "val")

            def _attach_style_fields(run_obj: Dict[str, Any]) -> None:
                if r_style_id:
                    run_obj["r_style_id"] = r_style_id
                if run_local_r:
                    run_obj["r_format"] = run_local_r

            for node in child:
                if node.tag == qn("w:rPr"):
                    continue

                if node.tag == qn("w:t"):
                    txt = node.text or ""
                    preserve = node.get(f"{{{XML_NS}}}space") == "preserve"

                    run_obj: Dict[str, Any] = {"type": "text", "text": txt}
                    run_obj["id"] = f"{parent_id}.run_{run_counter}"
                    run_obj["parent_id"] = parent_id
                    _attach_style_fields(run_obj)
                    if preserve:
                        run_obj["meta"] = {"preserve": True}

                    out.append(run_obj)
                    if not first_emitted:
                        first_emitted = True

                elif node.tag == qn("w:tab"):
                    run_obj: Dict[str, Any] = {"type": "tab"}
                    run_obj["id"] = f"{parent_id}.run_{run_counter}"
                    run_obj["parent_id"] = parent_id
                    _attach_style_fields(run_obj)

                    # ADD: first visual token is a tab -> meta.leading=true
                    if not first_emitted:
                        run_obj["meta"] = {"leading": True}
                        first_emitted = True

                    out.append(run_obj)

                elif node.tag == qn("w:br"):
                    run_obj: Dict[str, Any] = {"type": "break"}
                    run_obj["id"] = f"{parent_id}.run_{run_counter}"
                    run_obj["parent_id"] = parent_id
                    br_type = node.get(f"{{{W_NS}}}type")
                    if br_type in ("textWrapping", "page", "column"):
                        run_obj["break_type"] = br_type
                    _attach_style_fields(run_obj)
                    out.append(run_obj)
                    first_emitted = True

                elif node.tag == qn("w:cr"):
                    run_obj: Dict[str, Any] = {"type": "cr"}
                    run_obj["id"] = f"{parent_id}.run_{run_counter}"
                    run_obj["parent_id"] = parent_id
                    _attach_style_fields(run_obj)
                    out.append(run_obj)
                    first_emitted = True

                elif node.tag == qn("w:softHyphen"):
                    run_obj: Dict[str, Any] = {"type": "softHyphen"}
                    run_obj["id"] = f"{parent_id}.run_{run_counter}"
                    run_obj["parent_id"] = parent_id
                    _attach_style_fields(run_obj)
                    out.append(run_obj)
                    first_emitted = True

                elif node.tag == qn("w:noBreakHyphen"):
                    run_obj: Dict[str, Any] = {"type": "noBreakHyphen"}
                    run_obj["id"] = f"{parent_id}.run_{run_counter}"
                    run_obj["parent_id"] = parent_id
                    _attach_style_fields(run_obj)
                    out.append(run_obj)
                    first_emitted = True

                elif node.tag == qn("w:sym"):
                    font = node.get(f"{{{W_NS}}}font") or ""
                    char = node.get(f"{{{W_NS}}}char") or ""
                    run_obj: Dict[str, Any] = {
                        "type": "sym",
                        "text": _sym_encode(font, char),
                        "id": f"{parent_id}.run_{run_counter}",
                        "parent_id": parent_id
                    }
                    _attach_style_fields(run_obj)
                    out.append(run_obj)
                    first_emitted = True

                elif node.tag == qn("w:drawing") or node.tag == qn("w:pict"):
                    # Обработка изображения
                    run_id = f"{parent_id}.run_{run_counter}"
                    # run_counter увеличится в конце цикла
                    pic_data = parse_picture_node(node, run_id, self.relationships)
                    if pic_data:
                        _attach_style_fields(pic_data)
                        out.append(pic_data)
                        first_emitted = True
                else:
                    # ignore unsupported nodes for "forms" subset
                    continue
                run_counter += 1  # увеличиваем локальный счётчик после каждого обработанного run'а

        return out

    def _validate_raw_v212(self, result: Dict[str, Any]) -> None:
        required = [
            ("meta",),
            ("meta", "schema_version"),
            ("document_info",),
            ("numbering_definitions",),
            ("doc_defaults",),
            ("doc_defaults", "p_format"),
            ("doc_defaults", "r_format"),
            ("latent_styles",),
            ("styles",),
            ("content",),
        ]
        for path in required:
            cur: Any = result
            for part in path:
                if not isinstance(cur, dict) or part not in cur:
                    raise ValueError(f"RAW validation failed: missing key {'/'.join(path)}")
                cur = cur[part]

        forbidden = {"character_styles", "style_id", "char_style_id", "source_word_style_id"}

        def _scan(obj: Any, path: str) -> None:
            if isinstance(obj, dict):
                for k, v in obj.items():
                    if k in forbidden:
                        raise ValueError(f"RAW validation failed: forbidden field '{k}' at {path or '/'}")
                    _scan(v, f"{path}/{k}")
            elif isinstance(obj, list):
                for i, it in enumerate(obj):
                    _scan(it, f"{path}[{i}]")

        _scan(result, "")

        styles = result.get("styles", {})
        if not isinstance(styles, dict):
            raise ValueError("RAW validation failed: styles must be an object")

        for i, item in enumerate(result.get("content", [])):
            if not isinstance(item, dict):
                raise ValueError(f"RAW validation failed: content[{i}] must be an object")

            item_type = item.get("type")
            if item_type == "paragraph":
                if "p_style_id" not in item:
                    raise ValueError(f"RAW validation failed: content[{i}] missing p_style_id")
                if "runs" not in item or not isinstance(item["runs"], list):
                    raise ValueError(f"RAW validation failed: content[{i}] missing runs list")
                for j, run in enumerate(item["runs"]):
                    if not isinstance(run, dict):
                        raise ValueError(f"RAW validation failed: content[{i}].runs[{j}] must be object")
                    r_style_id = run.get("r_style_id")
                    if isinstance(r_style_id, str) and r_style_id not in styles:
                        print(f"[parser] warn: r_style_id '{r_style_id}' missing in styles at content[{i}].runs[{j}]")
                    if run.get("type") == "picture":
                        file_val = run.get("file")
                        if not isinstance(file_val, str) or not file_val.startswith("media/"):
                            raise ValueError(
                                f"RAW validation failed: picture file must start with media/ at content[{i}].runs[{j}]")
                        ext = run.get("extent")
                        if ext is not None:
                            if not isinstance(ext, dict):
                                raise ValueError(
                                    f"RAW validation failed: picture extent must be object at content[{i}].runs[{j}]")
                            for key in ("cx", "cy"):
                                if key in ext and not isinstance(ext[key], int):
                                    raise ValueError(
                                        f"RAW validation failed: picture extent.{key} must be integer at content[{i}].runs[{j}]")
            elif item_type == "table":
                if "id" not in item:
                    raise ValueError(f"RAW validation failed: table at content[{i}] missing id")
                if "tblPr" in item and not isinstance(item["tblPr"], dict):
                    raise ValueError(f"RAW validation failed: table tblPr must be object at content[{i}]")
                if "tbl_grid" in item and not isinstance(item["tbl_grid"], list):
                    raise ValueError(f"RAW validation failed: table tbl_grid must be array at content[{i}]")
                if "rows" not in item or not isinstance(item["rows"], list):
                    raise ValueError(f"RAW validation failed: table at content[{i}] missing rows list")
                for j, row in enumerate(item["rows"]):
                    if not isinstance(row, dict):
                        raise ValueError(f"RAW validation failed: table[{i}].rows[{j}] must be object")
                    if "id" not in row:
                        raise ValueError(f"RAW validation failed: table[{i}].rows[{j}] missing id")
                    if "cells" not in row or not isinstance(row["cells"], list):
                        raise ValueError(f"RAW validation failed: table[{i}].rows[{j}] missing cells list")
                    for k, cell in enumerate(row["cells"]):
                        if not isinstance(cell, dict):
                            raise ValueError(f"RAW validation failed: table[{i}].rows[{j}].cells[{k}] must be object")
                        if "content" not in cell or not isinstance(cell["content"], list):
                            raise ValueError(
                                f"RAW validation failed: table[{i}].rows[{j}].cells[{k}] missing content list")
            elif item_type == "shape":
                if "shape" not in item:
                    raise ValueError(f"RAW validation failed: shape at content[{i}] missing shape property")
            else:
                raise ValueError(f"RAW validation failed: unknown content type '{item_type}' at content[{i}]")

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
