# UltimateReconstructorV10.py
# Reconstructor Version: v10
# Schema Version: 2.8.2
# Rules Version: 0.2

# RAW JSON -> DOCX reconstructor using lxml (NO python-docx)
# Target: visually deterministic for "forms" subset (no tables/images/fields/hyperlinks)

from __future__ import annotations

import argparse
import json
import os
import zipfile
from typing import Any, Dict, Optional, List, Tuple

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

NSMAP_W = {"w": W_NS, "r": R_NS}
BASE_DIR = "."


def qn_w(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def qn_r(local: str) -> str:
    return f"{{{R_NS}}}{local}"


def _w_el(tag_local: str, attrib: Optional[Dict[str, str]] = None, nsmap=None) -> etree._Element:
    return etree.Element(qn_w(tag_local), attrib=attrib or {}, nsmap=nsmap)


def _w_sub(parent: etree._Element, tag_local: str, attrib: Optional[Dict[str, str]] = None) -> etree._Element:
    el = etree.SubElement(parent, qn_w(tag_local), attrib=attrib or {})
    return el


def _set_w_attr(el: etree._Element, local: str, val: Any) -> None:
    el.set(f"{{{W_NS}}}{local}", str(val))


def _set_w_attr_int(el: etree._Element, local: str, val: Any) -> None:
    iv = _safe_int(val)
    if iv is None:
        return
    _set_w_attr(el, local, iv)


def _needs_xml_preserve(text: str) -> bool:
    """
    Deterministic rule to keep visual spaces:
    - leading/trailing space
    - two consecutive spaces anywhere
    """
    if not text:
        return False
    if text[0] == " " or text[-1] == " ":
        return True
    if "  " in text:
        return True
    return False


def _safe_int(v: Any) -> Optional[int]:
    try:
        if v is None:
            return None
        return int(v)
    except Exception:
        return None


def _pick_docdefaults_from_styles(styles: Dict[str, Any], content: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Choose the most frequent style_id from content and use its r_format as docDefaults.
    Deterministic tie-break: smallest style_id.
    """
    counts: Dict[str, int] = {}
    for p in content:
        sid = p.get("style_id")
        if not sid:
            continue
        counts[sid] = counts.get(sid, 0) + 1

    if not counts:
        return {}

    best_count = max(counts.values())
    best_ids = sorted([sid for sid, c in counts.items() if c == best_count])
    best = best_ids[0]
    st = styles.get(best, {})
    rfmt = st.get("r_format", {})
    return rfmt if isinstance(rfmt, dict) else {}


class UltimateReconstructorV10:
    def __init__(self, raw_json_path: str):
        self.raw_json_path = raw_json_path
        with open(raw_json_path, "r", encoding="utf-8") as f:
            self.data = json.load(f)

        # validate minimal keys (soft)
        for k in ("meta", "document_info", "numbering_definitions", "styles", "content"):
            if k not in self.data:
                raise ValueError(f"RAW JSON missing required key: {k}")

        self.meta: Dict[str, Any] = self.data["meta"] or {}
        self.document_info: Dict[str, Any] = self.data["document_info"] or {}
        self.numbering_definitions: Dict[str, Any] = self.data["numbering_definitions"] or {}
        self.styles: Dict[str, Any] = self.data["styles"] or {}
        self.content: List[Dict[str, Any]] = self.data["content"] or []

        self.default_style_id: Optional[str] = None
        default_style_id = self.meta.get("default_style_id")
        if isinstance(default_style_id, str) and default_style_id in self.styles:
            self.default_style_id = default_style_id

        unsupported_run_types = []
        for p_item in self.content:
            for run in (p_item.get("runs") or []):
                rtype = run.get("type")
                if rtype in ("softHyphen", "noBreakHyphen"):
                    unsupported_run_types.append(rtype)

        print(
            f"[recon] summary input_schema={self.meta.get('schema_version')} input_rules={self.meta.get('rules_version')} "
            f"default_style_id={self.default_style_id} paragraphs_count={len(self.content)} "
            f"unsupported_runs_count={len(unsupported_run_types)}"
        )

    # =========================
    # PUBLIC
    # =========================

    def build_docx(self, out_docx_path: str) -> None:
        package_files: Dict[str, bytes] = {}

        # XML parts
        document_xml = self._build_document_xml()
        styles_xml = self._build_styles_xml()
        numbering_xml = self._build_numbering_xml()

        settings_xml = self._build_settings_xml()

        # Relationships & content types
        rels_root = self._build_root_rels()
        doc_rels = self._build_document_rels(has_numbering=bool(self.numbering_definitions), has_styles=True, has_settings=True)
        content_types = self._build_content_types(has_numbering=bool(self.numbering_definitions), has_styles=True, has_settings=True)

        package_files["word/document.xml"] = self._serialize_xml(document_xml, standalone=True)
        package_files["word/styles.xml"] = self._serialize_xml(styles_xml, standalone=True)
        package_files["word/settings.xml"] = self._serialize_xml(settings_xml, standalone=True)
        if self.numbering_definitions:
            package_files["word/numbering.xml"] = self._serialize_xml(numbering_xml, standalone=True)

        package_files["_rels/.rels"] = self._serialize_xml(rels_root, standalone=True)
        package_files["word/_rels/document.xml.rels"] = self._serialize_xml(doc_rels, standalone=True)
        package_files["[Content_Types].xml"] = self._serialize_xml(content_types, standalone=True)

        raw_reconstructed_dir = os.path.join(os.path.dirname(self.raw_json_path), "raw", "reconstructed")
        self._dump_reconstructed_parts(package_files, raw_reconstructed_dir)

        # Write docx (zip)
        os.makedirs(os.path.dirname(out_docx_path), exist_ok=True)
        with zipfile.ZipFile(out_docx_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
            # deterministic ordering
            for name in sorted(package_files.keys()):
                z.writestr(name, package_files[name])

    # =========================
    # BUILD: document.xml
    # =========================

    def _build_document_xml(self) -> etree._Element:
        doc = _w_el("document", nsmap={
            "w": W_NS,
            "r": R_NS
        })
        body = _w_sub(doc, "body")

        for p_item in self.content:
            p = _w_sub(body, "p")
            style_id = p_item.get("style_id")
            runs = p_item.get("runs", [])
            if style_id is None:
                style_id = ""

            st = self.styles.get(style_id, {"p_format": {}, "r_format": {}})
            p_format = st.get("p_format", {}) or {}
            base_r = st.get("r_format", {}) or {}

            # paragraph properties
            pPr = self._build_pPr(p_format)
            if pPr is not None:
                p.append(pPr)

            # runs
            if not runs:
                # empty paragraph: keep it empty (format already in pPr)
                # note: Word uses paragraph mark's rPr inside pPr/rPr, but we store base_r in style.
                # For visual determinism, add paragraph mark rPr to pPr if style has run-format.
                if base_r:
                    if pPr is None:
                        pPr = _w_sub(p, "pPr")
                    rPr_mark = self._build_rPr(base_r)
                    if rPr_mark is not None:
                        # paragraph mark formatting is inside pPr/rPr
                        # avoid duplicating if already exists
                        _w_sub(pPr, "rPr")
                        pPr.find(qn_w("rPr")).extend(list(rPr_mark))
                continue

            for run in runs:
                rtype = run.get("type")
                if rtype not in ("text", "tab", "break", "sym", "cr"):
                    # unsupported run types do not contribute in v10 (RULE-RUN-UNSUPPORTED)
                    continue

                r = _w_sub(p, "r")
                diff = run.get("diff", {}) or {}
                effective_r = self._merge_formats(base_r, diff)

                rPr = self._build_rPr(effective_r)
                if rPr is not None:
                    r.append(rPr)

                if rtype == "text":
                    txt = run.get("text", "")
                    t = _w_sub(r, "t")
                    if _needs_xml_preserve(txt):
                        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    t.text = txt

                elif rtype == "tab":
                    _w_sub(r, "tab")

                elif rtype == "break":
                    br = _w_sub(r, "br")
                    # optional break_type from RAW schema (textWrapping/page/column)
                    bt = run.get("break_type")
                    if bt in ("textWrapping", "page", "column"):
                        _set_w_attr(br, "type", bt)

                elif rtype == "cr":
                    _w_sub(r, "cr")

                elif rtype == "sym":
                    # In your schema: sym.text is a string; we treat it as literal char if possible.
                    # If it's JSON like {"font":"Wingdings","char":"F0A7"}, you can adjust later.
                    sym_text = run.get("text", "")
                    # Best-effort: output as plain text
                    t = _w_sub(r, "t")
                    if _needs_xml_preserve(sym_text):
                        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    t.text = sym_text

                else:
                    # unreachable due to early type guard
                    continue

        # sectPr from document_info.page_setup (must exist per your requirement)
        sectPr = self._build_sectPr(self.document_info.get("page_setup", {}) or {})
        if sectPr is not None:
            body.append(sectPr)

        return doc

    # =========================
    # BUILD: pPr
    # =========================

    def _build_pPr(self, p_format: Dict[str, Any]) -> Optional[etree._Element]:
        if not p_format:
            return None

        pPr = _w_el("pPr")

        # alignment
        align = p_format.get("alignment")
        if align in ("LEFT", "CENTER", "RIGHT", "JUSTIFY"):
            jc = _w_sub(pPr, "jc")
            _set_w_attr(jc, "val", align.lower() if align != "JUSTIFY" else "both")

        # indents
        ind_keys = ("indentStartTwip", "indentEndTwip", "indentFirstLineTwip", "indentHangingTwip")
        if any(k in p_format for k in ind_keys):
            ind = _w_sub(pPr, "ind")
            if "indentStartTwip" in p_format:
                _set_w_attr_int(ind, "left", p_format.get("indentStartTwip"))
            if "indentEndTwip" in p_format:
                _set_w_attr_int(ind, "right", p_format.get("indentEndTwip"))
            if "indentFirstLineTwip" in p_format:
                _set_w_attr_int(ind, "firstLine", p_format.get("indentFirstLineTwip"))
            if "indentHangingTwip" in p_format:
                _set_w_attr_int(ind, "hanging", p_format.get("indentHangingTwip"))

        # spacing
        if any(k in p_format for k in (
            "spaceBeforeTwip",
            "spaceAfterTwip",
            "spaceBeforeLines",
            "spaceAfterLines",
            "beforeAutospacing",
            "afterAutospacing",
            "lineTwip",
            "lineRule",
        )):
            sp = _w_sub(pPr, "spacing")
            if "spaceBeforeTwip" in p_format:
                _set_w_attr_int(sp, "before", p_format.get("spaceBeforeTwip"))
            if "spaceAfterTwip" in p_format:
                _set_w_attr_int(sp, "after", p_format.get("spaceAfterTwip"))
            if "spaceBeforeLines" in p_format:
                _set_w_attr_int(sp, "beforeLines", p_format.get("spaceBeforeLines"))
            if "spaceAfterLines" in p_format:
                _set_w_attr_int(sp, "afterLines", p_format.get("spaceAfterLines"))
            if "beforeAutospacing" in p_format:
                _set_w_attr(sp, "beforeAutospacing", "1" if p_format.get("beforeAutospacing") else "0")
            if "afterAutospacing" in p_format:
                _set_w_attr(sp, "afterAutospacing", "1" if p_format.get("afterAutospacing") else "0")
            if "lineTwip" in p_format:
                _set_w_attr_int(sp, "line", p_format.get("lineTwip"))
            lr = p_format.get("lineRule")
            if lr == "AUTO":
                _set_w_attr(sp, "lineRule", "auto")
            elif lr == "AT_LEAST":
                _set_w_attr(sp, "lineRule", "atLeast")
            elif lr == "EXACT":
                _set_w_attr(sp, "lineRule", "exact")

        # tabs (tab stops)
        tabs = p_format.get("tabs")
        if isinstance(tabs, list) and tabs:
            tabs_el = _w_sub(pPr, "tabs")
            for t in tabs:
                if not isinstance(t, dict):
                    continue
                pos = _safe_int(t.get("posTwip"))
                val = t.get("val")
                if pos is None or val is None:
                    continue
                tab_el = _w_sub(tabs_el, "tab")
                _set_w_attr(tab_el, "pos", pos)
                _set_w_attr(tab_el, "val", val)
                leader = t.get("leader")
                if leader is not None:
                    _set_w_attr(tab_el, "leader", leader)

        # numbering
        num = p_format.get("numbering")
        if isinstance(num, dict):
            numId = num.get("numId")
            ilvl = num.get("ilvl")
            if numId is not None and ilvl is not None:
                ilvl_int = _safe_int(ilvl)
                if ilvl_int is not None:
                    numPr = _w_sub(pPr, "numPr")
                    ilvl_el = _w_sub(numPr, "ilvl")
                    _set_w_attr(ilvl_el, "val", ilvl_int)
                    numId_el = _w_sub(numPr, "numId")
                    _set_w_attr(numId_el, "val", str(numId))

        # boolean paragraph flags (emit explicit false as val="0")
        for k, tag in (
            ("keepNext", "keepNext"),
            ("keepLines", "keepLines"),
            ("pageBreakBefore", "pageBreakBefore"),
            ("widowControl", "widowControl"),
            ("snapToGrid", "snapToGrid"),
            ("contextualSpacing", "contextualSpacing"),
        ):
            if k not in p_format:
                continue
            el = _w_sub(pPr, tag)
            if p_format.get(k) is False:
                _set_w_attr(el, "val", "0")

        text_alignment = p_format.get("textAlignment")
        if text_alignment in ("AUTO", "BASELINE", "TOP", "CENTER", "BOTTOM"):
            ta = _w_sub(pPr, "textAlignment")
            _set_w_attr(ta, "val", text_alignment.lower())

        return pPr

    # =========================
    # BUILD: rPr
    # =========================

    def _build_rPr(self, r_format: Dict[str, Any]) -> Optional[etree._Element]:
        if not r_format:
            return None

        rPr = _w_el("rPr")

        # rFonts
        rFonts = r_format.get("rFonts")
        if isinstance(rFonts, dict) and rFonts:
            rf = _w_sub(rPr, "rFonts")
            for k, v in rFonts.items():
                if v is None:
                    continue
                # store as w:ascii etc.
                _set_w_attr(rf, k, v)

        # size
        if "font_size_half_points" in r_format:
            szv = _safe_int(r_format.get("font_size_half_points"))
            if szv is not None:
                sz = _w_sub(rPr, "sz")
                _set_w_attr(sz, "val", szv)

        # bold/italic
        if "bold" in r_format:
            b = _w_sub(rPr, "b")
            if r_format.get("bold") is False:
                _set_w_attr(b, "val", "0")
        if "italic" in r_format:
            i = _w_sub(rPr, "i")
            if r_format.get("italic") is False:
                _set_w_attr(i, "val", "0")

        # underline
        if "underline" in r_format:
            u = _w_sub(rPr, "u")
            _set_w_attr(u, "val", r_format.get("underline"))

        # color
        if "color" in r_format:
            c = _w_sub(rPr, "color")
            _set_w_attr(c, "val", r_format.get("color"))

        # vertical_align
        if "vertical_align" in r_format:
            va = r_format.get("vertical_align")
            if va in ("baseline", "superscript", "subscript"):
                v = _w_sub(rPr, "vertAlign")
                _set_w_attr(v, "val", va)

        # all_caps
        if "all_caps" in r_format:
            if r_format.get("all_caps") is True:
                _w_sub(rPr, "caps")
            else:
                caps = _w_sub(rPr, "caps")
                _set_w_attr(caps, "val", "0")

        # lang
        lang = r_format.get("lang")
        if isinstance(lang, dict) and lang:
            le = _w_sub(rPr, "lang")
            for k, v in lang.items():
                if v is None:
                    continue
                _set_w_attr(le, k, v)

        # (optional extended) charSpacingTwip -> w:spacing
        if "charSpacingTwip" in r_format:
            spv = _safe_int(r_format.get("charSpacingTwip"))
            if spv is not None:
                sp = _w_sub(rPr, "spacing")
                _set_w_attr(sp, "val", spv)

        # (optional extended) positionHalfPoints -> w:position
        if "positionHalfPoints" in r_format:
            posv = _safe_int(r_format.get("positionHalfPoints"))
            if posv is not None:
                pos = _w_sub(rPr, "position")
                _set_w_attr(pos, "val", posv)

        return rPr

    # =========================
    # BUILD: sectPr
    # =========================

    def _build_sectPr(self, page_setup: Dict[str, Any]) -> Optional[etree._Element]:
        # Always create sectPr if any page setup present
        if not isinstance(page_setup, dict) or not page_setup:
            # still create minimal sectPr? You said "Page setup обязательно".
            sectPr = _w_el("sectPr")
            return sectPr

        sectPr = _w_el("sectPr")

        # pgSz
        if any(k in page_setup for k in ("pageWidthTwip", "pageHeightTwip", "orient")):
            pgSz = _w_sub(sectPr, "pgSz")
            if "pageWidthTwip" in page_setup:
                _set_w_attr_int(pgSz, "w", page_setup.get("pageWidthTwip"))
            if "pageHeightTwip" in page_setup:
                _set_w_attr_int(pgSz, "h", page_setup.get("pageHeightTwip"))
            orient = page_setup.get("orient")
            if orient in ("portrait", "landscape"):
                _set_w_attr(pgSz, "orient", orient)

        # pgMar
        if any(k in page_setup for k in ("marginLeftTwip", "marginRightTwip", "marginTopTwip", "marginBottomTwip", "headerTwip", "footerTwip", "gutterTwip")):
            pgMar = _w_sub(sectPr, "pgMar")
            if "marginTopTwip" in page_setup:
                _set_w_attr_int(pgMar, "top", page_setup.get("marginTopTwip"))
            if "marginRightTwip" in page_setup:
                _set_w_attr_int(pgMar, "right", page_setup.get("marginRightTwip"))
            if "marginBottomTwip" in page_setup:
                _set_w_attr_int(pgMar, "bottom", page_setup.get("marginBottomTwip"))
            if "marginLeftTwip" in page_setup:
                _set_w_attr_int(pgMar, "left", page_setup.get("marginLeftTwip"))
            if "headerTwip" in page_setup:
                _set_w_attr_int(pgMar, "header", page_setup.get("headerTwip"))
            if "footerTwip" in page_setup:
                _set_w_attr_int(pgMar, "footer", page_setup.get("footerTwip"))
            if "gutterTwip" in page_setup:
                _set_w_attr_int(pgMar, "gutter", page_setup.get("gutterTwip"))

        # cols
        cols = page_setup.get("cols")
        if isinstance(cols, dict) and cols:
            cols_el = _w_sub(sectPr, "cols")
            if "num" in cols:
                _set_w_attr_int(cols_el, "num", cols.get("num"))
            if "spaceTwip" in cols:
                _set_w_attr_int(cols_el, "space", cols.get("spaceTwip"))
            if "equalWidth" in cols:
                _set_w_attr(cols_el, "equalWidth", "1" if cols.get("equalWidth") else "0")

        # docGrid is optional; not in schema now. Skip.

        return sectPr

    # =========================
    # BUILD: styles.xml
    # =========================

    def _build_styles_xml(self) -> etree._Element:
        styles = etree.Element(qn_w("styles"), nsmap={"w": W_NS})

        # docDefaults from meta.default_style_id r_format when valid, otherwise fallback to most common style's r_format (deterministic)
        use_default_style_id = self.default_style_id is not None and self.default_style_id in self.styles
        if use_default_style_id:
            dd_style = self.styles.get(self.default_style_id, {})
            dd_r = dd_style.get("r_format", {}) if isinstance(dd_style, dict) else {}
        else:
            dd_r = _pick_docdefaults_from_styles(self.styles, self.content)

        docDefaults = _w_sub(styles, "docDefaults")
        rPrDefault = _w_sub(docDefaults, "rPrDefault")
        rPr = _w_sub(rPrDefault, "rPr")

        # Apply rFonts + sz + lang into docDefaults if present
        if isinstance(dd_r, dict):
            if "rFonts" in dd_r:
                rf = _w_sub(rPr, "rFonts")
                for k, v in (dd_r.get("rFonts") or {}).items():
                    if v is None:
                        continue
                    _set_w_attr(rf, k, v)
            if "font_size_half_points" in dd_r:
                sz = _w_sub(rPr, "sz")
                _set_w_attr_int(sz, "val", dd_r.get("font_size_half_points"))
            if "lang" in dd_r:
                lang = _w_sub(rPr, "lang")
                for k, v in (dd_r.get("lang") or {}).items():
                    if v is None:
                        continue
                    _set_w_attr(lang, k, v)

        # Minimal default paragraph style ("Normal")
        st = _w_sub(styles, "style", attrib={
            f"{{{W_NS}}}type": "paragraph",
            f"{{{W_NS}}}default": "1",
            f"{{{W_NS}}}styleId": "Normal"
        })
        name = _w_sub(st, "name")
        _set_w_attr(name, "val", "Normal")

        if use_default_style_id:
            normal_style = self.styles.get(self.default_style_id, {})
            if isinstance(normal_style, dict):
                normal_p = normal_style.get("p_format", {}) or {}
                normal_r = normal_style.get("r_format", {}) or {}
                if isinstance(normal_p, dict) and normal_p:
                    pPr = self._build_pPr(normal_p)
                    if pPr is not None:
                        st.append(pPr)
                if isinstance(normal_r, dict) and normal_r:
                    rPr = self._build_rPr(normal_r)
                    if rPr is not None:
                        st.append(rPr)

        return styles

    # =========================
    # BUILD: numbering.xml
    # =========================

    def _build_numbering_xml(self) -> etree._Element:
        numbering = etree.Element(qn_w("numbering"), nsmap={"w": W_NS})

        # Collect abstractNum definitions by abstractNumId from JSON records
        # JSON shape: numId -> {abstractNumId?, levels{ilvl: {...}}, lvl_overrides?}
        abstract_map: Dict[str, Dict[str, Any]] = {}
        for numId, rec in self.numbering_definitions.items():
            abs_id = rec.get("abstractNumId")
            levels = rec.get("levels", {}) or {}
            if abs_id is None:
                raise ValueError(f"Contract violation: missing abstractNumId for numId={numId} in numbering. Run effective materializer or preserve numbering mappings.")
            if abs_id not in abstract_map:
                abstract_map[abs_id] = {"levels": levels}
            else:
                # keep first; deterministic
                pass

        # Write abstractNum in sorted order for determinism
        for abs_id in sorted(abstract_map.keys(), key=lambda x: str(x)):
            abs_el = _w_sub(numbering, "abstractNum", attrib={f"{{{W_NS}}}abstractNumId": str(abs_id)})

            levels = abstract_map[abs_id].get("levels", {}) or {}
            # levels keys are strings of ints
            for ilvl_str in sorted(levels.keys(), key=lambda x: int(x) if str(x).isdigit() else str(x)):
                lvl_rec = levels[ilvl_str] or {}
                fmt = lvl_rec.get("format")
                if not isinstance(fmt, str):
                    continue

                template = lvl_rec.get("template")
                if not isinstance(template, str):
                    continue

                lvl_el = _w_sub(abs_el, "lvl", attrib={f"{{{W_NS}}}ilvl": str(ilvl_str)})
                if "start" in lvl_rec:
                    sv = _safe_int(lvl_rec.get("start"))
                    if sv is not None:
                        start = _w_sub(lvl_el, "start")
                        _set_w_attr(start, "val", sv)

                lvl_el = _w_sub(abs_el, "lvl", attrib={f"{{{W_NS}}}ilvl": str(ilvl_str)})
                if "start" in lvl_rec:
                    sv = _safe_int(lvl_rec.get("start"))
                    if sv is not None:
                        start = _w_sub(lvl_el, "start")
                        _set_w_attr(start, "val", sv)

                fmt = lvl_rec.get("format")
                if not isinstance(fmt, str):
                    continue
                numFmt = _w_sub(lvl_el, "numFmt")
                _set_w_attr(numFmt, "val", fmt)

                template = lvl_rec.get("template")
                if not isinstance(template, str):
                    continue
                lvlText = _w_sub(lvl_el, "lvlText")
                _set_w_attr(lvlText, "val", template)

                # minimal; can be extended later (lvlJc/pPr/rPr)

        # Write nums
        for numId in sorted(self.numbering_definitions.keys(), key=lambda x: int(x) if str(x).isdigit() else str(x)):
            rec = self.numbering_definitions[numId] or {}
            num_el = _w_sub(numbering, "num", attrib={f"{{{W_NS}}}numId": str(numId)})

            abs_id = rec.get("abstractNumId")
            if abs_id is None:
                raise ValueError(f"Contract violation: missing abstractNumId for numId={numId} in numbering. Run effective materializer or preserve numbering mappings.")

            abs_ref = _w_sub(num_el, "abstractNumId")
            _set_w_attr(abs_ref, "val", str(abs_id))

            # lvl overrides
            ovs = rec.get("lvl_overrides", {}) or {}
            if isinstance(ovs, dict) and ovs:
                for ilvl_str in sorted(ovs.keys(), key=lambda x: int(x) if str(x).isdigit() else str(x)):
                    ov = ovs[ilvl_str] or {}
                    if "start" in ov:
                        ov_start = _safe_int(ov.get("start"))
                        if ov_start is not None:
                            lvlOv = _w_sub(num_el, "lvlOverride", attrib={f"{{{W_NS}}}ilvl": str(ilvl_str)})
                            st = _w_sub(lvlOv, "startOverride")
                            _set_w_attr(st, "val", ov_start)

        return numbering

    # =========================
    # BUILD: settings.xml
    # =========================

    def _build_settings_xml(self) -> etree._Element:
        settings = etree.Element(qn_w("settings"), nsmap={"w": W_NS})

        dts = None
        doc_settings = (self.document_info.get("settings") or {})
        if isinstance(doc_settings, dict):
            dts = doc_settings.get("defaultTabStopTwip")

        if dts is not None:
            dts_int = _safe_int(dts)
            if dts_int is not None:
                el = _w_sub(settings, "defaultTabStop")
                _set_w_attr(el, "val", dts_int)

        return settings

    def _dump_reconstructed_parts(self, package_files: Dict[str, bytes], out_root_dir: str) -> None:
        os.makedirs(out_root_dir, exist_ok=True)
        for name, data in sorted(package_files.items()):
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            out_path = os.path.join(out_root_dir, name)
            os.makedirs(os.path.dirname(out_path), exist_ok=True)
            with open(out_path, "wb") as f:
                f.write(data)


    # =========================
    # RELATIONSHIPS & CONTENT TYPES
    # =========================

    def _build_root_rels(self) -> etree._Element:
        rels = etree.Element(f"{{{PR_NS}}}Relationships", nsmap={None: PR_NS})

        rel = etree.SubElement(rels, f"{{{PR_NS}}}Relationship")
        rel.set("Id", "rId1")
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
        rel.set("Target", "word/document.xml")

        return rels

    def _build_document_rels(self, has_numbering: bool, has_styles: bool, has_settings: bool) -> etree._Element:
        rels = etree.Element(f"{{{PR_NS}}}Relationships", nsmap={None: PR_NS})

        rid = 1
        if has_styles:
            rel = etree.SubElement(rels, f"{{{PR_NS}}}Relationship")
            rel.set("Id", f"rId{rid}")
            rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
            rel.set("Target", "styles.xml")
            rid += 1

        if has_numbering:
            rel = etree.SubElement(rels, f"{{{PR_NS}}}Relationship")
            rel.set("Id", f"rId{rid}")
            rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering")
            rel.set("Target", "numbering.xml")
            rid += 1

        if has_settings:
            rel = etree.SubElement(rels, f"{{{PR_NS}}}Relationship")
            rel.set("Id", f"rId{rid}")
            rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings")
            rel.set("Target", "settings.xml")
            rid += 1

        return rels

    def _build_content_types(self, has_numbering: bool, has_styles: bool, has_settings: bool) -> etree._Element:
        types = etree.Element(f"{{{CT_NS}}}Types", nsmap={None: CT_NS})

        # Defaults
        d1 = etree.SubElement(types, f"{{{CT_NS}}}Default")
        d1.set("Extension", "rels")
        d1.set("ContentType", "application/vnd.openxmlformats-package.relationships+xml")

        d2 = etree.SubElement(types, f"{{{CT_NS}}}Default")
        d2.set("Extension", "xml")
        d2.set("ContentType", "application/xml")

        # Overrides
        o1 = etree.SubElement(types, f"{{{CT_NS}}}Override")
        o1.set("PartName", "/word/document.xml")
        o1.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")

        if has_styles:
            o = etree.SubElement(types, f"{{{CT_NS}}}Override")
            o.set("PartName", "/word/styles.xml")
            o.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")

        if has_numbering:
            o = etree.SubElement(types, f"{{{CT_NS}}}Override")
            o.set("PartName", "/word/numbering.xml")
            o.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")

        if has_settings:
            o = etree.SubElement(types, f"{{{CT_NS}}}Override")
            o.set("PartName", "/word/settings.xml")
            o.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml")

        return types

    # =========================
    # HELPERS
    # =========================

    def _serialize_xml(self, root: etree._Element, standalone: bool = True) -> bytes:
        return etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=standalone,
            pretty_print=False
        )

    def _merge_formats(self, base: Dict[str, Any], diff: Dict[str, Any]) -> Dict[str, Any]:
        out = dict(base or {})
        for k, v in (diff or {}).items():
            if v is None:
                continue
            out[k] = v
        return out


if __name__ == "__main__":
    try:
        cli = argparse.ArgumentParser(description="UltimateReconstructorV10 RAW JSON -> DOCX")
        cli.add_argument("--in-json", dest="input_json", default=os.path.join(BASE_DIR, "donor_v2.6.json"))
        cli.add_argument("--out-docx", dest="output_docx", default=os.path.join(BASE_DIR, "donor_v2.6_reconstructed.docx"))
        args = cli.parse_args()

        recon = UltimateReconstructorV10(args.input_json)
        recon.build_docx(args.output_docx)
        print("Реконструкция v4.1_plus завершена успешно!")
    except Exception:
        import traceback
        traceback.print_exc()
