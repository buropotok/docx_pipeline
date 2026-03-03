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
#
# IMPORTANT: All patches are applied without deleting existing elements.
#            We only update attributes of existing elements, or create new ones
#            only if the corresponding element doesn't exist at all.

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

def _update_attributes(el: etree._Element, attrs: Dict[str, Any], ns=W_NS) -> None:
    """Update attributes of an element from a dict, converting values to string."""
    for key, value in attrs.items():
        if value is not None:
            el.set(f"{{{ns}}}{key}", str(value))

def _update_subelement_attributes(parent: etree._Element, tag_local: str, attrs: Dict[str, Any]) -> None:
    """Find or create subelement, then update its attributes."""
    el = parent.find(qn_w(tag_local))
    if el is None:
        el = etree.SubElement(parent, qn_w(tag_local))
    _update_attributes(el, attrs)

def _update_border_side(borders_el: etree._Element, side: str, border_dict: Optional[Dict[str, Any]]) -> None:
    """Update a single border side element with given attributes, or leave untouched if None."""
    if border_dict is None:
        return
    side_el = borders_el.find(qn_w(side))
    if side_el is None:
        side_el = etree.SubElement(borders_el, qn_w(side))
    _update_attributes(side_el, border_dict)

def _update_margin_side(mar_el: etree._Element, side: str, margin_dict: Optional[Dict[str, Any]]) -> None:
    """Update a single margin side element with given attributes (w, type)."""
    if margin_dict is None:
        return
    side_el = mar_el.find(qn_w(side))
    if side_el is None:
        side_el = etree.SubElement(mar_el, qn_w(side))
    if "w" in margin_dict and margin_dict["w"] is not None:
        side_el.set(qn_w("w"), str(margin_dict["w"]))
    if "type" in margin_dict and margin_dict["type"] is not None:
        side_el.set(qn_w("type"), str(margin_dict["type"]))

def _update_size_type(parent: etree._Element, tag_local: str, value: Optional[Dict[str, Any]]) -> None:
    """Update or create a sizeType element (like w:tblW) with w and type attributes."""
    if value is None:
        return
    el = parent.find(qn_w(tag_local))
    if el is None:
        el = etree.SubElement(parent, qn_w(tag_local))
    if "w" in value and value["w"] is not None:
        el.set(qn_w("w"), str(value["w"]))
    if "type" in value and value["type"] is not None:
        el.set(qn_w("type"), str(value["type"]))

def _update_val_attr(parent: etree._Element, tag_local: str, attr_local: str, value: Optional[Any]) -> None:
    """Update or create a simple element with a val attribute."""
    if value is None:
        return
    el = parent.find(qn_w(tag_local))
    if el is None:
        el = etree.SubElement(parent, qn_w(tag_local))
    el.set(qn_w(attr_local), str(value))

# -----------------------
# Patch helpers (minimal, partial overwrite only)
# -----------------------

def patch_tbl_grid(tbl: etree._Element, grid: Optional[List[int]]) -> None:
    """
    Replace tblGrid completely if grid is provided.
    This is a structural change, so it's allowed to replace.
    """
    if grid is None:
        return
    # Remove old grid
    for ch in list(tbl):
        if ch.tag == qn_w("tblGrid"):
            tbl.remove(ch)
    # Create new
    tblGrid = etree.SubElement(tbl, qn_w("tblGrid"))
    for w in grid:
        col = etree.SubElement(tblGrid, qn_w("gridCol"))
        iw = _safe_int(w)
        if iw is None:
            raise ValueError("tbl_grid must contain integers (twips)")
        col.set(qn_w("w"), str(iw))

def patch_tblPr(tbl: etree._Element, tblPr_patch: Optional[Dict[str, Any]]) -> None:
    """Surgically update table properties without removing whole elements."""
    if not isinstance(tblPr_patch, dict) or not tblPr_patch:
        return
    tblPr = _get_or_create(tbl, "tblPr")

    # Simple val attributes
    if "tblStyle" in tblPr_patch:
        _update_val_attr(tblPr, "tblStyle", "val", tblPr_patch["tblStyle"])
    if "tblOverlap" in tblPr_patch:
        _update_val_attr(tblPr, "tblOverlap", "val", tblPr_patch["tblOverlap"])
    if "jc" in tblPr_patch:
        _update_val_attr(tblPr, "jc", "val", tblPr_patch["jc"])

    # SizeType elements
    if "tblW" in tblPr_patch:
        _update_size_type(tblPr, "tblW", tblPr_patch["tblW"])
    if "tblInd" in tblPr_patch:
        _update_size_type(tblPr, "tblInd", tblPr_patch["tblInd"])
    if "tblCellSpacing" in tblPr_patch:
        _update_size_type(tblPr, "tblCellSpacing", tblPr_patch["tblCellSpacing"])

    # Borders: update or create individual sides
    if "tblBorders" in tblPr_patch:
        borders_data = tblPr_patch["tblBorders"]
        if borders_data is None:
            borders_data = {}
        borders_el = tblPr.find(qn_w("tblBorders"))
        if borders_el is None:
            borders_el = etree.SubElement(tblPr, qn_w("tblBorders"))
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            side_data = borders_data.get(side)
            _update_border_side(borders_el, side, side_data)

    # Cell margins
    if "tblCellMar" in tblPr_patch:
        mar_data = tblPr_patch["tblCellMar"]
        if mar_data is None:
            mar_data = {}
        mar_el = tblPr.find(qn_w("tblCellMar"))
        if mar_el is None:
            mar_el = etree.SubElement(tblPr, qn_w("tblCellMar"))
        for side in ("top", "left", "bottom", "right"):
            side_data = mar_data.get(side)
            _update_margin_side(mar_el, side, side_data)

    # tblLayout: simple type attribute
    if "tblLayout" in tblPr_patch:
        lay = tblPr_patch["tblLayout"]
        if isinstance(lay, dict) and "type" in lay:
            _update_val_attr(tblPr, "tblLayout", "type", lay["type"])

    # tblLook: update flags, preserve val
    if "tblLook" in tblPr_patch:
        look_data = tblPr_patch["tblLook"]
        if not isinstance(look_data, dict):
            return
        tblLook_el = tblPr.find(qn_w("tblLook"))
        if tblLook_el is None:
            # In a valid donor, it should exist, but if not, we create only if val is provided
            if look_data.get("val") is None:
                raise ValueError("tblLook missing w:val: can't create tblLook without donor val")
            tblLook_el = etree.SubElement(tblPr, qn_w("tblLook"))
        # Update flags if present
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
            if flag in look_data and look_data[flag] is not None:
                tblLook_el.set(qn_w(attr), "1" if bool(look_data[flag]) else "0")
        # Update val only if provided
        if "val" in look_data and look_data["val"] is not None:
            tblLook_el.set(qn_w("val"), str(look_data["val"]))

def patch_trPr(tr: etree._Element, trPr_patch: Optional[Dict[str, Any]]) -> None:
    """Surgically update row properties."""
    if not isinstance(trPr_patch, dict) or not trPr_patch:
        return
    trPr = _get_or_create(tr, "trPr")

    # cantSplit and tblHeader are presence elements (no attributes)
    if "cantSplit" in trPr_patch:
        # If true, ensure element exists; if false, remove it
        existing = trPr.find(qn_w("cantSplit"))
        if bool(trPr_patch["cantSplit"]):
            if existing is None:
                etree.SubElement(trPr, qn_w("cantSplit"))
        else:
            if existing is not None:
                trPr.remove(existing)

    if "tblHeader" in trPr_patch:
        existing = trPr.find(qn_w("tblHeader"))
        if bool(trPr_patch["tblHeader"]):
            if existing is None:
                etree.SubElement(trPr, qn_w("tblHeader"))
        else:
            if existing is not None:
                trPr.remove(existing)

    # trHeight: update or create
    if "trHeight" in trPr_patch:
        height = trPr_patch["trHeight"]
        if isinstance(height, dict):
            height_el = trPr.find(qn_w("trHeight"))
            if height_el is None:
                height_el = etree.SubElement(trPr, qn_w("trHeight"))
            if "val" in height and height["val"] is not None:
                height_el.set(qn_w("val"), str(height["val"]))
            if "hRule" in height and height["hRule"] is not None:
                height_el.set(qn_w("hRule"), str(height["hRule"]))

    # jc (alignment)
    if "jc" in trPr_patch:
        _update_val_attr(trPr, "jc", "val", trPr_patch["jc"])

def patch_tcPr(tc: etree._Element, tcPr_patch: Optional[Dict[str, Any]]) -> None:
    """Surgically update cell properties."""
    if not isinstance(tcPr_patch, dict) or not tcPr_patch:
        return
    tcPr = _get_or_create(tc, "tcPr")

    # tcW
    if "tcW" in tcPr_patch:
        _update_size_type(tcPr, "tcW", tcPr_patch["tcW"])

    # gridSpan
    if "gridSpan" in tcPr_patch:
        _update_val_attr(tcPr, "gridSpan", "val", tcPr_patch["gridSpan"])

    # vMerge
    if "vMerge" in tcPr_patch:
        vm_val = tcPr_patch["vMerge"]
        existing = tcPr.find(qn_w("vMerge"))
        if vm_val is not None:
            if existing is None:
                existing = etree.SubElement(tcPr, qn_w("vMerge"))
            existing.set(qn_w("val"), str(vm_val))
        else:
            # If vm_val is None, we might want to remove it? Typically vMerge with no val means "restart".
            # But we'll keep existing behavior: set val if provided.
            pass

    # shd
    if "shd" in tcPr_patch:
        shd_data = tcPr_patch["shd"]
        if isinstance(shd_data, dict):
            shd_el = tcPr.find(qn_w("shd"))
            if shd_el is None:
                shd_el = etree.SubElement(tcPr, qn_w("shd"))
            _update_attributes(shd_el, shd_data)

    # tcBorders
    if "tcBorders" in tcPr_patch:
        borders_data = tcPr_patch["tcBorders"]
        if borders_data is None:
            borders_data = {}
        borders_el = tcPr.find(qn_w("tcBorders"))
        if borders_el is None:
            borders_el = etree.SubElement(tcPr, qn_w("tcBorders"))
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            side_data = borders_data.get(side)
            _update_border_side(borders_el, side, side_data)

    # tcMar
    if "tcMar" in tcPr_patch:
        mar_data = tcPr_patch["tcMar"]
        if mar_data is None:
            mar_data = {}
        mar_el = tcPr.find(qn_w("tcMar"))
        if mar_el is None:
            mar_el = etree.SubElement(tcPr, qn_w("tcMar"))
        for side in ("top", "left", "bottom", "right"):
            side_data = mar_data.get(side)
            _update_margin_side(mar_el, side, side_data)

    # Other simple properties
    if "vAlign" in tcPr_patch:
        _update_val_attr(tcPr, "vAlign", "val", tcPr_patch["vAlign"])
    if "textDirection" in tcPr_patch:
        _update_val_attr(tcPr, "textDirection", "val", tcPr_patch["textDirection"])
    if "tcFitText" in tcPr_patch:
        existing = tcPr.find(qn_w("tcFitText"))
        if bool(tcPr_patch["tcFitText"]):
            if existing is None:
                etree.SubElement(tcPr, qn_w("tcFitText"))
        else:
            if existing is not None:
                tcPr.remove(existing)
    if "noWrap" in tcPr_patch:
        existing = tcPr.find(qn_w("noWrap"))
        if bool(tcPr_patch["noWrap"]):
            if existing is None:
                etree.SubElement(tcPr, qn_w("noWrap"))
        else:
            if existing is not None:
                tcPr.remove(existing)

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