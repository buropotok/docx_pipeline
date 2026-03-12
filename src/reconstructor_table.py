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
# - All patches are applied surgically - we only update existing elements,
#   never delete and recreate unless absolutely necessary.

from __future__ import annotations

# print("=" * 50)
# print("reconstructor_table.py ЗАГРУЖЕН!!!")
# print("=" * 50)

from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple, Callable
from copy import deepcopy
from lxml import etree
import re

_NEW_ID_RE = re.compile(r"\.\d+$")
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MY_NS = "https://translatefactory/schema/custom-id"


def is_new_id(elem_id: str) -> bool:
    """Check if element id indicates a newly created element (contains .digits at end)."""
    return bool(_NEW_ID_RE.search(str(elem_id)))


def qn_w(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def qn_my(local: str) -> str:
    return f"{{{MY_NS}}}{local}"


def _safe_int(v: Any) -> Optional[int]:
    """Safely convert value to int, return None if conversion fails."""
    try:
        if v is None:
            return None
        return int(v)
    except Exception:
        return None


def _update_attributes(el: etree._Element, attrs: Dict[str, Any], ns=W_NS) -> None:
    """
    Update attributes of an element from a dict, converting values to string.
    Only updates attributes that are present in the dict.
    """
    for key, value in attrs.items():
        if value is not None:
            el.set(f"{{{ns}}}{key}", str(value))


def _get_or_create(parent: etree._Element, tag_local: str) -> etree._Element:
    """Find or create a child element with the given tag."""
    el = parent.find(qn_w(tag_local))
    if el is None:
        el = etree.SubElement(parent, qn_w(tag_local))
    return el


def _update_val_attr(
    parent: etree._Element, tag_local: str, attr_local: str, value: Optional[Any]
) -> None:
    """
    Update or create a simple element with a val attribute.
    Example: <w:tblStyle w:val="..." />
    """
    if value is None:
        return
    el = parent.find(qn_w(tag_local))
    if el is None:
        el = etree.SubElement(parent, qn_w(tag_local))
    el.set(qn_w(attr_local), str(value))


def _update_size_type(
    parent: etree._Element, tag_local: str, value: Optional[Dict[str, Any]]
) -> None:
    """
    Update or create a sizeType element (like w:tblW, w:tcW) with w and type attributes.
    Example: <w:tblW w:w="5000" w:type="dxa" />
    """
    if value is None:
        return
    el = parent.find(qn_w(tag_local))
    if el is None:
        el = etree.SubElement(parent, qn_w(tag_local))
    if "w" in value and value["w"] is not None:
        el.set(qn_w("w"), str(value["w"]))
    if "type" in value and value["type"] is not None:
        el.set(qn_w("type"), str(value["type"]))


def _update_border_side(
    borders_el: etree._Element, side: str, border_dict: Optional[Dict[str, Any]]
) -> None:
    """
    Update a single border side element with given attributes.
    If the side element doesn't exist, it will be created.
    """
    if border_dict is None:
        return
    side_el = borders_el.find(qn_w(side))
    if side_el is None:
        side_el = etree.SubElement(borders_el, qn_w(side))
    _update_attributes(side_el, border_dict)


def _update_margin_side(
    mar_el: etree._Element, side: str, margin_dict: Optional[Dict[str, Any]]
) -> None:
    """
    Update a single margin side element with given attributes (w, type).
    Example: <w:top w:w="120" w:type="dxa" />
    """
    if margin_dict is None:
        return
    side_el = mar_el.find(qn_w(side))
    if side_el is None:
        side_el = etree.SubElement(mar_el, qn_w(side))
    if "w" in margin_dict and margin_dict["w"] is not None:
        side_el.set(qn_w("w"), str(margin_dict["w"]))
    if "type" in margin_dict and margin_dict["type"] is not None:
        side_el.set(qn_w("type"), str(margin_dict["type"]))


def _update_bool_presence(
    parent: etree._Element, tag_local: str, value: Optional[bool]
) -> None:
    """
    Update a boolean presence element (like w:cantSplit, w:tblHeader).
    If value is True, ensure element exists.
    If value is False, ensure element does NOT exist.
    If value is None, do nothing.
    """
    if value is None:
        return
    existing = parent.find(qn_w(tag_local))
    if value:
        if existing is None:
            etree.SubElement(parent, qn_w(tag_local))
    else:
        if existing is not None:
            parent.remove(existing)


# ----------------------------------------------------------------------
# Public patch functions (surgical updates)
# ----------------------------------------------------------------------


def patch_tbl_grid(tbl: etree._Element, grid: Optional[List[int]]) -> None:
    """
    Replace tblGrid completely if grid is provided.
    This is a structural change, so it's allowed to replace.
    """
    if grid is None:
        return

    # Remove old tblGrid
    for ch in list(tbl):
        if ch.tag == qn_w("tblGrid"):
            tbl.remove(ch)

    # Create new tblGrid
    tblGrid = etree.SubElement(tbl, qn_w("tblGrid"))
    for w in grid:
        col = etree.SubElement(tblGrid, qn_w("gridCol"))
        iw = _safe_int(w)
        if iw is None:
            raise ValueError("tbl_grid must contain integers (twips)")
        col.set(qn_w("w"), str(iw))


def patch_tblPr(tbl: etree._Element, tblPr_patch: Optional[Dict[str, Any]]) -> None:
    """
    Surgically update table properties without removing whole elements.
    Only updates fields that are present in the patch.
    """
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

    # Borders - update individual sides
    if "tblBorders" in tblPr_patch:
        borders_data = tblPr_patch["tblBorders"] or {}
        borders_el = tblPr.find(qn_w("tblBorders"))
        if borders_el is None:
            borders_el = etree.SubElement(tblPr, qn_w("tblBorders"))
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            side_data = borders_data.get(side)
            _update_border_side(borders_el, side, side_data)

    # Cell margins - update individual sides
    if "tblCellMar" in tblPr_patch:
        mar_data = tblPr_patch["tblCellMar"] or {}
        mar_el = tblPr.find(qn_w("tblCellMar"))
        if mar_el is None:
            mar_el = etree.SubElement(tblPr, qn_w("tblCellMar"))
        for side in ("top", "left", "bottom", "right"):
            side_data = mar_data.get(side)
            _update_margin_side(mar_el, side, side_data)

    # tblLayout - simple type attribute
    if "tblLayout" in tblPr_patch:
        lay = tblPr_patch["tblLayout"]
        if isinstance(lay, dict) and "type" in lay:
            _update_val_attr(tblPr, "tblLayout", "type", lay["type"])
        elif isinstance(lay, str):
            _update_val_attr(tblPr, "tblLayout", "type", lay)

    # tblLook - update flags, preserve val if not overridden
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

        # Update val only if explicitly provided
        if "val" in look_data and look_data["val"] is not None:
            tblLook_el.set(qn_w("val"), str(look_data["val"]))
        # Otherwise preserve existing val (do nothing)


def patch_trPr(tr: etree._Element, trPr_patch: Optional[Dict[str, Any]]) -> None:
    """
    Surgically update row properties.
    """
    if not isinstance(trPr_patch, dict) or not trPr_patch:
        return

    trPr = _get_or_create(tr, "trPr")

    # Boolean presence elements
    if "cantSplit" in trPr_patch:
        _update_bool_presence(trPr, "cantSplit", trPr_patch["cantSplit"])
    if "tblHeader" in trPr_patch:
        _update_bool_presence(trPr, "tblHeader", trPr_patch["tblHeader"])

    # trHeight
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
    """
    Surgically update cell properties.
    """
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
        # If vm_val is None, we keep existing (do nothing)

    # shd (shading)
    if "shd" in tcPr_patch:
        shd_data = tcPr_patch["shd"]
        if isinstance(shd_data, dict):
            shd_el = tcPr.find(qn_w("shd"))
            if shd_el is None:
                shd_el = etree.SubElement(tcPr, qn_w("shd"))
            _update_attributes(shd_el, shd_data)

    # tcBorders
    if "tcBorders" in tcPr_patch:
        borders_data = tcPr_patch["tcBorders"] or {}
        borders_el = tcPr.find(qn_w("tcBorders"))
        if borders_el is None:
            borders_el = etree.SubElement(tcPr, qn_w("tcBorders"))
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            side_data = borders_data.get(side)
            _update_border_side(borders_el, side, side_data)

    # tcMar (cell margins)
    if "tcMar" in tcPr_patch:
        mar_data = tcPr_patch["tcMar"] or {}
        mar_el = tcPr.find(qn_w("tcMar"))
        if mar_el is None:
            mar_el = etree.SubElement(tcPr, qn_w("tcMar"))
        for side in ("top", "left", "bottom", "right"):
            side_data = mar_data.get(side)
            _update_margin_side(mar_el, side, side_data)

    # vAlign
    if "vAlign" in tcPr_patch:
        _update_val_attr(tcPr, "vAlign", "val", tcPr_patch["vAlign"])

    # textDirection
    if "textDirection" in tcPr_patch:
        _update_val_attr(tcPr, "textDirection", "val", tcPr_patch["textDirection"])

    # tcFitText (boolean presence)
    if "tcFitText" in tcPr_patch:
        _update_bool_presence(tcPr, "tcFitText", tcPr_patch["tcFitText"])

    # noWrap (boolean presence)
    if "noWrap" in tcPr_patch:
        _update_bool_presence(tcPr, "noWrap", tcPr_patch["noWrap"])


# ----------------------------------------------------------------------
# Cell paragraph positioning (contract rules)
# ----------------------------------------------------------------------


@dataclass
class CellParaIndex:
    """Index of paragraphs inside a cell."""
    paras_in_order: List[etree._Element]
    id_by_el: Dict[int, str]  # id(elem) -> para_id
    el_by_id: Dict[str, etree._Element]  # para_id -> element


def index_cell_paragraphs(
    tc: etree._Element, row_id: str, cell_index_1based: int
) -> CellParaIndex:
    """
    Build index of paragraphs inside a cell.
    Generates paragraph ids as {row_id}.cell_{cell_index}.p_{index}.
    """
    paras = [ch for ch in list(tc) if ch.tag == qn_w("p")]
    cell_id = f"{row_id}.cell_{cell_index_1based}"
    id_by_el: Dict[int, str] = {}
    el_by_id: Dict[str, etree._Element] = {}
    for i, p in enumerate(paras, start=1):
        pid = f"{cell_id}.p_{i}"
        id_by_el[id(p)] = pid
        el_by_id[pid] = p
    return CellParaIndex(
        paras_in_order=paras, id_by_el=id_by_el, el_by_id=el_by_id
    )


def replace_cell_paragraphs(
    tc: etree._Element, new_paras: List[etree._Element]
) -> None:
    """
    Replace paragraphs in a cell with new ones.
    Preserves tcPr if present.
    """
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
      - planned_modifications: list of tuples (para_id, para_el, para_json)
        for paragraphs that should be modified (pPr+runs) by caller
      - final_paragraph_elements: ordered list of paragraph elements to set into cell
    """

    # print(f"\n apply_cell_paragraph_ops для ячейки {row_id}.cell_{cell_index_1based}", flush=True)
    # print(f"   json_paragraphs получен: {json_paragraphs}", flush=True)
    # print(f"   тип json_paragraphs: {type(json_paragraphs)}", flush=True)
    # print(f"   длина json_paragraphs: {len(json_paragraphs)}", flush=True)

    idx = index_cell_paragraphs(tc, row_id, cell_index_1based)

    # print(f"   Оригинальные параграфы в ячейке: {list(idx.el_by_id.keys())}")
    # print(f"   JSON параграфов: {[p.get('id') for p in json_paragraphs]}")
    # for pj in json_paragraphs:
    #     pid = pj.get('id')
        # print(f"   Проверка {pid}: is_new_id = {is_new_id(str(pid))}", flush=True)
        # if pid == 'tbl_1.row_1.cell_1.p_5':
        #     print(f"    НАЙДЕН P_5 в JSON: {pj}")
        #     print(f"      derive_from: {pj.get('derive_from')}")
        #     print(f"      runs: {pj.get('runs', [])}")

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
        working = [
            p for p in working if idx.id_by_el.get(id(p)) not in explicitly_deleted_ids
        ]

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
                raise ValueError(
                    f"derive_from paragraph '{derive_from}' not found in cell for new paragraph '{pid}'."
                )
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
                    raise ValueError(
                        f"Cell paragraph '{pid}' anchor '{anchor}' disappeared after removal (fatal)."
                    )
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


# ----------------------------------------------------------------------
# Row processing
# ----------------------------------------------------------------------


def iter_row_cells(tr: etree._Element) -> List[etree._Element]:
    """Return all cell elements in a row, in order."""
    return [ch for ch in list(tr) if ch.tag == qn_w("tc")]


def set_table_rows(tbl: etree._Element, rows: List[etree._Element]) -> None:
    """Replace all rows in a table with new ones."""
    # Remove all existing rows
    for ch in list(tbl):
        if ch.tag == qn_w("tr"):
            tbl.remove(ch)
    # Add new rows in order
    for r in rows:
        tbl.append(r)


def clone_row(tr: etree._Element) -> etree._Element:
    """Create a deep copy of a row element."""
    return deepcopy(tr)