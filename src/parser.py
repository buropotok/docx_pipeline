# UltimateParserV41.py
# DOCX -> RAW JSON (schema v2.7.x / v2.8-ish) using lxml (NO python-docx)
# Parser Version: v41
# Schema Version: 2.8.1
# Rules Version: 0.2

# Deterministic, visually-lossless for "forms" subset (no tables/images/fields/hyperlinks).
#
# Based on your UltimateParserV41, with careful incremental improvements (NO hardcode):
# 1) FIX: spacing autospacing flags parsed correctly from w:spacing attributes (not as child elements).
# 2) ADD: spacing beforeLines/afterLines -> spaceBeforeLines/spaceAfterLines (if present).
# 3) ADD BACK: rPr char spacing + position (charSpacingTwip, positionHalfPoints) that existed earlier.
# 4) ADD: meta.leading=true for the first tab run in a paragraph (schema supports meta.leading).
#
# Rules enforced:
# - RULE-001: hanging -> indentHangingTwip
# - RULE-003: preserve spaces (no normalization; store xml:space preserve marker)
# - RULE-004: do not merge runs during parsing (do not concatenate across w:r or across w:t)
# - RULE-006: if paragraph has runs -> base_r = effective style r_format
#            if paragraph has no runs -> base_r includes paragraph mark rPr merged in

from __future__ import annotations

import json
import os
import zipfile
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
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
    # If some unexpected value appears, treat as True for presence semantics
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


class UltimateParserV41:
    """
    Deterministic DOCX -> RAW JSON parser (schema v2.7-ish).
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

            self.settings_xml_raw: Optional[bytes] = None
            try:
                self.settings_xml_raw = z.read("word/settings.xml")
                self.settings_xml = etree.fromstring(self.settings_xml_raw)
            except KeyError:
                self.settings_xml = None

        # Output style library (schema.styles)
        self.out_styles: Dict[str, Dict[str, Any]] = {}
        self._style_key_to_id: Dict[str, str] = {}
        self._style_counter: int = 1

        # Word styles (for effective formatting)
        self.doc_defaults_r: Dict[str, Any] = {}
        self.doc_defaults_p: Dict[str, Any] = {}
        self.word_styles: Dict[str, WordStyle] = {}
        self.default_paragraph_style_id: Optional[str] = None

        self._init_word_styles()

    # =========================
    # PUBLIC
    # =========================

    def process(self) -> str:
        settings: Dict[str, Any] = {}
        default_tab_stop = self._parse_default_tab_stop()
        if default_tab_stop is not None:
            settings["defaultTabStopTwip"] = default_tab_stop
        if self.settings_xml_raw is not None:
            settings["raw_settings_xml"] = self.settings_xml_raw.decode("utf-8")

        result: Dict[str, Any] = {
            "meta": {
                "schema_version": "2.8.1",
                "rules_version": "0.2",
                "producer": {
                    "name": "UltimateParserV41",
                    "version": "v41"
                }
            },
            "document_info": {
                "page_setup": self._parse_page_setup(),
                "settings": settings
            },
            "numbering_definitions": self._parse_numbering_definitions(),
            "styles": {},
            "content": []
        }

        body = self.document_xml.find(qn("w:body"))
        if body is None:
            result["styles"] = self.out_styles
            return json.dumps(result, ensure_ascii=False, indent=2)

        for p in body.findall(qn("w:p")):
            pPr = p.find(qn("w:pPr"))
            p_style_id = self._get_p_style_id(pPr)

            # Effective paragraph formatting: docDefaults + style chain + direct pPr
            base_p_format = self._effective_p_format(p_style_id, pPr)

            # Effective run formatting: docDefaults + style chain
            style_r_format = self._effective_r_format(p_style_id)

            # Paragraph mark rPr (direct pPr/rPr) – apply only for empty paragraphs (RULE-006)
            para_mark_rPr = self._parse_rPr(pPr.find(qn("w:rPr")) if pPr is not None else None)

            runs = self._parse_runs(p, base_r_for_diff=style_r_format)

            # RULE-006:
            if runs:
                base_r_for_style = style_r_format
            else:
                base_r_for_style = _merge(style_r_format, para_mark_rPr)

            style_id = self._register_out_style(base_p_format, base_r_for_style)

            result["content"].append({
                "style_id": style_id,
                "runs": runs
            })

        result["styles"] = self.out_styles
        return json.dumps(result, ensure_ascii=False, indent=2)

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
        base = _merge(base, self._parse_pPr(direct_pPr))

        return base

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
    # OUTPUT STYLE LIBRARY (schema.styles)
    # =========================

    def _register_out_style(self, p_format: Dict[str, Any], r_format: Dict[str, Any]) -> str:
        style_obj = {
            "p_format": p_format or {},
            "r_format": r_format or {},
        }
        key = _stable_json_key(style_obj)
        existing = self._style_key_to_id.get(key)
        if existing:
            return existing

        style_id = f"s{self._style_counter:04d}"
        self._style_counter += 1
        self._style_key_to_id[key] = style_id
        self.out_styles[style_id] = style_obj
        return style_id

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
                levels[str(ilvl)] = level_rec

            abstracts[abs_id] = {"levels": levels}

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
            if lvl_overrides:
                rec["lvl_overrides"] = lvl_overrides

            out[numId] = rec

        return out

    # =========================
    # PARAGRAPH FORMATTING (pPr -> pFormat)
    # =========================

    def _parse_pPr(self, pPr: Optional[etree._Element]) -> Dict[str, Any]:
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
            if right is not None:
                out["indentEndTwip"] = right
            if first is not None:
                out["indentFirstLineTwip"] = first
            if hanging is not None:
                out["indentHangingTwip"] = hanging  # RULE-001

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
        RULE-004: do not merge runs during parsing:
          - never combine content across different w:r
          - within a single w:r, do NOT merge different w:t nodes
        """
        out: List[Dict[str, Any]] = []
        first_emitted = False  # for meta.leading on first tab

        for child in p:
            if child.tag != qn("w:r"):
                continue

            rPr = child.find(qn("w:rPr"))
            r_cur = self._parse_rPr(rPr)
            r_diff = _dict_diff(base_r_for_diff, r_cur) if r_cur else {}

            for node in child:
                if node.tag == qn("w:rPr"):
                    continue

                if node.tag == qn("w:t"):
                    txt = node.text or ""
                    preserve = node.get(f"{{{XML_NS}}}space") == "preserve"

                    run_obj: Dict[str, Any] = {"type": "text", "text": txt}
                    if r_diff:
                        run_obj["diff"] = r_diff
                    if preserve:
                        run_obj["meta"] = {"preserve": True}

                    out.append(run_obj)
                    first_emitted = True

                elif node.tag == qn("w:tab"):
                    run_obj: Dict[str, Any] = {"type": "tab"}
                    if r_diff:
                        run_obj["diff"] = r_diff

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
                    if r_diff:
                        run_obj["diff"] = r_diff
                    out.append(run_obj)
                    first_emitted = True

                elif node.tag == qn("w:cr"):
                    run_obj: Dict[str, Any] = {"type": "cr"}
                    if r_diff:
                        run_obj["diff"] = r_diff
                    out.append(run_obj)
                    first_emitted = True

                elif node.tag == qn("w:sym"):
                    font = node.get(f"{{{W_NS}}}font") or ""
                    char = node.get(f"{{{W_NS}}}char") or ""
                    run_obj: Dict[str, Any] = {"type": "sym", "text": _sym_encode(font, char)}
                    if r_diff:
                        run_obj["diff"] = r_diff
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
        input_name = "donor_v2.6.docx"
        input_docx = os.path.join(BASE_DIR, input_name)

        parser = UltimateParserV41(input_docx)

        out_json = os.path.join(BASE_DIR, os.path.splitext(input_name)[0] + ".json")
        with open(out_json, "w", encoding="utf-8") as f:
            f.write(parser.process())

        raw_donor_dir = os.path.join(os.path.dirname(out_json), "raw", "donor")
        parser.dump_donor_xml_parts(raw_donor_dir)

        print("Парсинг v4.1+ (lxml) завершен успешно!")
    except Exception:
        import traceback
        traceback.print_exc()
