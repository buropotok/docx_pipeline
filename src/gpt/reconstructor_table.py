# reconstructor_table.py
# DOCX Pipeline v2.15 (contract-compliant)
# Table reconstruction helpers (rows/cells/nested paragraphs) for reconstructor.py
#
# Principles:
# - Tables/rows are modified "surgically" on donor XML.
# - Table rows are rebuilt as an ordered list from JSON `rows`.
# - Cells are matched by index (no cell ids in XML).
# - Paragraphs inside cells are treated as managed elements with generated ids:
#     {row_id}.cell_{cell_index}.p_{p_index}
# - Paragraph runs are rebuilt by reconstructor.py; this module only positions/clones
#   and applies tcPr/trPr/tblPr patches (partial, only provided fields).

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple, Callable
from copy import deepcopy
from lxml import etree
import re

_NEW_ID_RE = re.compile(r"\.\d+$")
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MY_NS = "https://translatefactory/schema/custom-id"

def is_new_id(elem_id: str) -> bool:
    return bool(_NEW_ID_RE.search(str(elem_id)))

def qn_w(local: str) -> str:
    return f"{{{W_NS}}}{local}"

def qn_my(local: str) -> str:
    return f"{{{MY_NS}}}{local}"

def _safe_int(v: Any) -> Optional[int]:
    try:
        if v is None:
            return None
        return int(v)
    except Exception:
        return None

def _get_or_create(parent: etree._Element, tag_local: str) -> etree._Element:
    el = parent.find(qn_w(tag_local))
    if el is None:
        el = etree.SubElement(parent, qn_w(tag_local))
    return el

def _remove_children_by_tag(parent: etree._Element, tag_local: str) -> None:
    for ch in list(parent):
        if ch.tag == qn_w(tag_local):
            parent.remove(ch)

# -----------------------
# Patch helpers (minimal, partial overwrite only)
# -----------------------

def patch_tbl_grid(tbl: etree._Element, grid: Optional[List[int]]) -> None:
    if grid is None:
        return
    _remove_children_by_tag(tbl, "tblGrid")
    tblGrid = etree.SubElement(tbl, qn_w("tblGrid"))
    for w in grid:
        col = etree.SubElement(tblGrid, qn_w("gridCol"))
        iw = _safe_int(w)
        if iw is None:
            raise ValueError("tbl_grid must contain integers (twips)")
        col.set(qn_w("w"), str(iw))

def _patch_size_type(parent: etree._Element, tag_local: str, value: Optional[Dict[str, Any]]) -> None:
    if value is None:
        return
    _remove_children_by_tag(parent, tag_local)
    el = etree.SubElement(parent, qn_w(tag_local))
    w = _safe_int(value.get("w"))
    typ = value.get("type")
    if w is not None:
        el.set(qn_w("w"), str(w))
    if typ is not None:
        el.set(qn_w("type"), str(typ))

def _patch_val_attr(parent: etree._Element, tag_local: str, attr_local: str, value: Optional[Any]) -> None:
    if value is None:
        return
    _remove_children_by_tag(parent, tag_local)
    el = etree.SubElement(parent, qn_w(tag_local))
    el.set(qn_w(attr_local), str(value))

def patch_tblPr(tbl: etree._Element, tblPr_patch: Optional[Dict[str, Any]]) -> None:
    if not isinstance(tblPr_patch, dict) or not tblPr_patch:
        return
    tblPr = _get_or_create(tbl, "tblPr")

    # Common keys produced by parser_table.py
    if "tblStyle" in tblPr_patch:
        _patch_val_attr(tblPr, "tblStyle", "val", tblPr_patch["tblStyle"])
    if "tblOverlap" in tblPr_patch:
        _patch_val_attr(tblPr, "tblOverlap", "val", tblPr_patch["tblOverlap"])
    if "jc" in tblPr_patch:
        _patch_val_attr(tblPr, "jc", "val", tblPr_patch["jc"])

    if "tblW" in tblPr_patch:
        _patch_size_type(tblPr, "tblW", tblPr_patch["tblW"])
    if "tblInd" in tblPr_patch:
        _patch_size_type(tblPr, "tblInd", tblPr_patch["tblInd"])
    if "tblCellSpacing" in tblPr_patch:
        _patch_size_type(tblPr, "tblCellSpacing", tblPr_patch["tblCellSpacing"])

    # Borders
    if "tblBorders" in tblPr_patch:
        _remove_children_by_tag(tblPr, "tblBorders")
        borders_el = etree.SubElement(tblPr, qn_w("tblBorders"))
        borders = tblPr_patch["tblBorders"] or {}
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            bd = borders.get(side)
            if not isinstance(bd, dict):
                continue
            side_el = etree.SubElement(borders_el, qn_w(side))
            for a in ("val", "sz", "color", "space", "themeColor", "themeTint", "themeShade"):
                if a in bd and bd[a] is not None:
                    side_el.set(qn_w(a), str(bd[a]))

    # Cell margins
    if "tblCellMar" in tblPr_patch:
        _remove_children_by_tag(tblPr, "tblCellMar")
        mar_el = etree.SubElement(tblPr, qn_w("tblCellMar"))
        mar = tblPr_patch["tblCellMar"] or {}
        for side in ("top", "left", "bottom", "right"):
            sd = mar.get(side)
            if not isinstance(sd, dict):
                continue
            side_el = etree.SubElement(mar_el, qn_w(side))
            if "w" in sd and sd["w"] is not None:
                side_el.set(qn_w("w"), str(sd["w"]))
            if "type" in sd and sd["type"] is not None:
                side_el.set(qn_w("type"), str(sd["type"]))

    if "tblLayout" in tblPr_patch:
        # {"type": "..."}
        lay = tblPr_patch["tblLayout"]
        if isinstance(lay, dict) and "type" in lay:
            _remove_children_by_tag(tblPr, "tblLayout")
            el = etree.SubElement(tblPr, qn_w("tblLayout"))
            el.set(qn_w("type"), str(lay["type"]))

    if "tblLook" in tblPr_patch:
        look = tblPr_patch["tblLook"]
        if isinstance(look, dict):
            el = tblPr.find(qn_w("tblLook"))
            if el is None:
                # создавать можно только если val явно задан,
                # иначе это "переизобретение" и риск invalid DOCX
                if look.get("val") is None:
                    raise ValueError("tblLook missing w:val: can't create tblLook without donor val")
                el = etree.SubElement(tblPr, qn_w("tblLook"))

            # val: если пришёл — обновляем, если нет — сохраняем донорский.
            if look.get("val") is not None:
                el.set(qn_w("val"), str(look["val"]))
            elif el.get(qn_w("val")) is None:
                # донорского нет и в JSON нет => получится битый tblLook
                raise ValueError("tblLook has no w:val in donor and JSON")

            # флаги — точечно, только если пришли
            for flag, attr in (
                    ("firstRow", "firstRow"),
                    ("lastRow", "lastRow"),
                    ("firstColumn", "firstColumn"),
                    ("lastColumn", "lastColumn"),
                    ("bandRow", "bandRow"),
                    ("bandCol", "bandCol"),
                    ("noHBand", "noHBand"),
                    ("noVBand", "noVBand"),
            ):
                if flag in look and look[flag] is not None:
                    el.set(qn_w(attr), "1" if bool(look[flag]) else "0")
            # raw val (optional)
            if "val" in look and look["val"] is not None:
                el.set(qn_w("val"), str(look["val"]))

def patch_trPr(tr: etree._Element, trPr_patch: Optional[Dict[str, Any]]) -> None:
    if not isinstance(trPr_patch, dict) or not trPr_patch:
        return
    trPr = _get_or_create(tr, "trPr")
    # Minimal: handle "cantSplit" and "tblHeader" if present
    if "cantSplit" in trPr_patch:
        _remove_children_by_tag(trPr, "cantSplit")
        if bool(trPr_patch["cantSplit"]):
            etree.SubElement(trPr, qn_w("cantSplit"))
    if "tblHeader" in trPr_patch:
        _remove_children_by_tag(trPr, "tblHeader")
        if bool(trPr_patch["tblHeader"]):
            etree.SubElement(trPr, qn_w("tblHeader"))

def patch_tcPr(tc: etree._Element, tcPr_patch: Optional[Dict[str, Any]]) -> None:
    if not isinstance(tcPr_patch, dict) or not tcPr_patch:
        return
    tcPr = _get_or_create(tc, "tcPr")
    # Minimal: width + gridSpan + vMerge + shading + borders + margins
    if "tcW" in tcPr_patch:
        _patch_size_type(tcPr, "tcW", tcPr_patch["tcW"])
    if "gridSpan" in tcPr_patch:
        _patch_val_attr(tcPr, "gridSpan", "val", tcPr_patch["gridSpan"])
    if "vMerge" in tcPr_patch:
        vm = tcPr_patch["vMerge"]
        _remove_children_by_tag(tcPr, "vMerge")
        el = etree.SubElement(tcPr, qn_w("vMerge"))
        if vm is not None:
            el.set(qn_w("val"), str(vm))
    if "shd" in tcPr_patch:
        shd = tcPr_patch["shd"]
        _remove_children_by_tag(tcPr, "shd")
        if isinstance(shd, dict):
            el = etree.SubElement(tcPr, qn_w("shd"))
            for a in ("val", "color", "fill", "themeColor", "themeTint", "themeShade"):
                if a in shd and shd[a] is not None:
                    el.set(qn_w(a), str(shd[a]))
    if "tcBorders" in tcPr_patch:
        _remove_children_by_tag(tcPr, "tcBorders")
        borders_el = etree.SubElement(tcPr, qn_w("tcBorders"))
        borders = tcPr_patch["tcBorders"] or {}
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            bd = borders.get(side)
            if not isinstance(bd, dict):
                continue
            side_el = etree.SubElement(borders_el, qn_w(side))
            for a in ("val", "sz", "color", "space", "themeColor", "themeTint", "themeShade"):
                if a in bd and bd[a] is not None:
                    side_el.set(qn_w(a), str(bd[a]))
    if "tcMar" in tcPr_patch:
        _remove_children_by_tag(tcPr, "tcMar")
        mar_el = etree.SubElement(tcPr, qn_w("tcMar"))
        mar = tcPr_patch["tcMar"] or {}
        for side in ("top", "left", "bottom", "right"):
            sd = mar.get(side)
            if not isinstance(sd, dict):
                continue
            side_el = etree.SubElement(mar_el, qn_w(side))
            if "w" in sd and sd["w"] is not None:
                side_el.set(qn_w("w"), str(sd["w"]))
            if "type" in sd and sd["type"] is not None:
                side_el.set(qn_w("type"), str(sd["type"]))

# -----------------------
# Cell paragraph positioning (contract rules)
# -----------------------

@dataclass
class CellParaIndex:
    paras_in_order: List[etree._Element]
    id_by_el: Dict[int, str]  # id(elem) -> para_id
    el_by_id: Dict[str, etree._Element]  # para_id -> element

def index_cell_paragraphs(tc: etree._Element, row_id: str, cell_index_1based: int) -> CellParaIndex:
    paras = [ch for ch in list(tc) if ch.tag == qn_w("p")]
    cell_id = f"{row_id}.cell_{cell_index_1based}"
    id_by_el: Dict[int, str] = {}
    el_by_id: Dict[str, etree._Element] = {}
    for i, p in enumerate(paras, start=1):
        pid = f"{cell_id}.p_{i}"
        id_by_el[id(p)] = pid
        el_by_id[pid] = p
    return CellParaIndex(paras_in_order=paras, id_by_el=id_by_el, el_by_id=el_by_id)

def replace_cell_paragraphs(tc: etree._Element, new_paras: List[etree._Element]) -> None:
    # Preserve tcPr if present, drop/replace only paragraphs
    children = list(tc)
    tcPr = tc.find(qn_w("tcPr"))
    # Remove all paragraph children
    for ch in children:
        if ch.tag == qn_w("p"):
            tc.remove(ch)

    # Insert new paragraphs after tcPr if tcPr exists, else at end
    insert_pos = 0
    if tcPr is not None:
        cur = list(tc)
        for i, ch in enumerate(cur):
            if ch is tcPr:
                insert_pos = i + 1
                break
    for p in new_paras:
        tc.insert(insert_pos, p)
        insert_pos += 1

def apply_cell_paragraph_ops(
    tc: etree._Element,
    row_id: str,
    cell_index_1based: int,
    json_paragraphs: List[Dict[str, Any]],
    *,
    clone_paragraph: Callable[[etree._Element], etree._Element],
) -> Tuple[List[Tuple[str, etree._Element, Dict[str, Any]]], List[etree._Element]]:
    """
    Applies deletion / insertion / movement of paragraphs inside a cell.

    Returns:
      - planned_modifications: list of tuples (para_id, para_el, para_json) for paragraphs that should be modified (pPr+runs) by caller
      - final_paragraph_elements: ordered list of paragraph elements to set into cell
    """
    idx = index_cell_paragraphs(tc, row_id, cell_index_1based)
    working: List[etree._Element] = list(idx.paras_in_order)

    # For cloning sources we keep original mapping even if element is removed from working
    source_by_id = dict(idx.el_by_id)

    # Explicit deletions
    explicitly_deleted_ids = set()
    for pj in json_paragraphs:
        if pj.get("deleted") is True:
            pid = pj.get("id")
            if not pid:
                raise ValueError("Cell paragraph with deleted:true missing id")
            explicitly_deleted_ids.add(pid)

    if explicitly_deleted_ids:
        working = [p for p in working if idx.id_by_el.get(id(p)) not in explicitly_deleted_ids]

    def find_index_by_pid(pid: str) -> int:
        for i, p in enumerate(working):
            if idx.id_by_el.get(id(p)) == pid:
                return i
        return -1

    planned: List[Tuple[str, etree._Element, Dict[str, Any]]] = []
    keep_ids: List[str] = []

    for pj in json_paragraphs:
        if pj.get("deleted") is True:
            continue
        pid = pj.get("id")
        if not pid:
            raise ValueError("Cell paragraph missing id")

        is_new = is_new_id(str(pid))
        if is_new:
            derive_from = pj.get("derive_from")
            if not derive_from:
                raise ValueError(f"New cell paragraph '{pid}' missing derive_from (fatal).")
            src = source_by_id.get(str(derive_from))
            if src is None:
                raise ValueError(f"derive_from paragraph '{derive_from}' not found in cell for new paragraph '{pid}'.")
            p_el = clone_paragraph(src)
            idx.id_by_el[id(p_el)] = str(pid)
        else:
            p_el = idx.el_by_id.get(str(pid))
            if p_el is None:
                raise ValueError(f"Cell paragraph '{pid}' not found in donor (fatal).")

        anchor = pj.get("anchor")
        position = pj.get("position")
        if (anchor is None) ^ (position is None):
            raise ValueError(f"Cell paragraph '{pid}' has only one of anchor/position (fatal).")

        if anchor is not None:
            anchor = str(anchor)
            a_idx = find_index_by_pid(anchor)
            if a_idx < 0:
                raise ValueError(f"Cell paragraph '{pid}' anchor '{anchor}' not found (fatal).")
            if p_el in working:
                working.remove(p_el)
                a_idx = find_index_by_pid(anchor)
                if a_idx < 0:
                    raise ValueError(f"Cell paragraph '{pid}' anchor '{anchor}' disappeared after removal (fatal).")
            ins = a_idx if position == "before" else a_idx + 1
            working.insert(ins, p_el)
        else:
            if p_el not in working:
                working.append(p_el)

        planned.append((str(pid), p_el, pj))
        keep_ids.append(str(pid))

    keep_set = set(keep_ids)
    final = [p for p in working if idx.id_by_el.get(id(p)) in keep_set]

    return planned, final

# -----------------------
# Row processing
# -----------------------

def iter_row_cells(tr: etree._Element) -> List[etree._Element]:
    return [ch for ch in list(tr) if ch.tag == qn_w("tc")]

def set_table_rows(tbl: etree._Element, rows: List[etree._Element]) -> None:
    for ch in list(tbl):
        if ch.tag == qn_w("tr"):
            tbl.remove(ch)
    for r in rows:
        tbl.append(r)

def clone_row(tr: etree._Element) -> etree._Element:
    return deepcopy(tr)