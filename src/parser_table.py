# parser_table.py
# Парсинг таблиц DOCX (w:tbl) в RAW JSON (схема v2.14)
# Поддержка всех свойств таблиц: границы, ширина, заливка, объединение ячеек и т.д.

import os
from typing import Any, Dict, List, Optional
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MY_NS = "https://translatefactory/schema/custom-id"


def qn(tag: str) -> str:
    """Преобразует тег вида w:name в полное имя с пространством имён."""
    pfx, local = tag.split(":")
    if pfx != "w":
        raise ValueError(f"Unsupported prefix: {pfx}")
    return f"{{{W_NS}}}{local}"


def qn_my(local: str) -> str:
    return f"{{{MY_NS}}}{local}"


def _str_attr(el: etree._Element, attr_local: str) -> Optional[str]:
    """Безопасно получает строковый атрибут элемента."""
    return el.get(f"{{{W_NS}}}{attr_local}")


def _int_attr(el: etree._Element, attr_local: str) -> Optional[int]:
    """Безопасно получает целочисленный атрибут элемента."""
    v = el.get(f"{{{W_NS}}}{attr_local}")
    if v is None:
        return None
    try:
        return int(v)
    except Exception:
        return None


def _bool_attr(el: etree._Element, attr_local: str) -> Optional[bool]:
    """Получает булев атрибут (0/1, true/false, on/off)."""
    v = el.get(f"{{{W_NS}}}{attr_local}")
    if v is None:
        return None
    if v in ("0", "false", "off"):
        return False
    if v in ("1", "true", "on"):
        return True
    return None


def _parse_border(border_el: Optional[etree._Element]) -> Optional[Dict[str, Any]]:
    """Парсит элемент границы (w:top, w:left и т.д.) в словарь."""
    if border_el is None:
        return None
    result = {
        "val": _str_attr(border_el, "val"),
        "sz": _int_attr(border_el, "sz"),
        "color": _str_attr(border_el, "color"),
    }
    # space опционален
    space = _int_attr(border_el, "space")
    if space is not None:
        result["space"] = space
    # themeColor, themeTint, themeShade могут быть добавлены позже
    return result if result["val"] is not None else None


def _parse_borders(parent: etree._Element) -> Optional[Dict[str, Any]]:
    """Парсит w:tblBorders или w:tcBorders."""
    borders_el = parent.find(qn("w:tblBorders")) or parent.find(qn("w:tcBorders"))
    if borders_el is None:
        return None
    result = {}
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        side_el = borders_el.find(qn(f"w:{side}"))
        border_dict = _parse_border(side_el)
        if border_dict:
            result[side] = border_dict
    return result if result else None


def _parse_size_type(el: Optional[etree._Element]) -> Optional[Dict[str, Any]]:
    """
    Парсит элемент, содержащий атрибуты w:w и w:type (например, w:tblW, w:tcW, w:tblInd, w:tblCellSpacing).
    Возвращает словарь {"w": int, "type": str} или None.
    """
    if el is None:
        return None
    w = _int_attr(el, "w")
    typ = _str_attr(el, "type")
    if w is None or typ is None:
        return None
    return {"w": w, "type": typ}


def _parse_cell_margins(parent: etree._Element) -> Optional[Dict[str, Any]]:
    """Парсит w:tblCellMar или w:tcMar."""
    mar_el = parent.find(qn("w:tblCellMar")) or parent.find(qn("w:tcMar"))
    if mar_el is None:
        return None
    result = {}
    for side in ("top", "left", "bottom", "right"):
        side_el = mar_el.find(qn(f"w:{side}"))
        if side_el is not None:
            size = _parse_size_type(side_el)
            if size:
                result[side] = size
    return result if result else None


def _parse_shd(el: Optional[etree._Element]) -> Optional[Dict[str, Any]]:
    """Парсит элемент w:shd."""
    if el is None:
        return None
    result = {
        "val": _str_attr(el, "val"),
        "color": _str_attr(el, "color"),
        "fill": _str_attr(el, "fill"),
    }
    # Дополнительные атрибуты темы (опционально)
    for attr in ("themeColor", "themeTint", "themeShade"):
        val = _str_attr(el, attr)
        if val:
            result[attr] = val
    return result if result.get("val") else None


def _parse_tbl_look(el: Optional[etree._Element]) -> Optional[Dict[str, Any]]:
    """Парсит элемент w:tblLook, преобразуя его в набор булевых флагов."""
    if el is None:
        return None

    result = {}

    # Сначала пробуем отдельные булевы атрибуты (современный формат)
    for flag in ("firstRow", "lastRow", "firstColumn", "lastColumn", "bandRow", "bandCol", "noHBand", "noVBand"):
        b = _bool_attr(el, flag)
        if b is not None:
            result[flag] = b

    # Если отдельные атрибуты отсутствуют, пытаемся интерпретировать hex-код
    val_hex = _str_attr(el, "val")
    if val_hex:
        try:
            mask = int(val_hex, 16)
            # Определяем флаги согласно [ISO/IEC 29500-1:2016] §17.7.6.8
            flags = {
                "firstRow":     0x0020,
                "lastRow":      0x0040,
                "firstColumn":  0x0080,
                "lastColumn":   0x0100,
                "noHBand":      0x0200,
                "noVBand":      0x0400,
            }
            for name, bit in flags.items():
                result.setdefault(name, bool(mask & bit))
        except ValueError:
            pass  # невалидный hex – игнорируем

    return result if result else None


def _parse_tbl_pr(tbl: etree._Element) -> Dict[str, Any]:
    """Парсит свойства таблицы (w:tblPr)."""
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        return {}

    result: Dict[str, Any] = {}

    # tblStyle
    tblStyle_el = tblPr.find(qn("w:tblStyle"))
    if tblStyle_el is not None:
        style = _str_attr(tblStyle_el, "val")
        if style:
            result["tblStyle"] = style

    # tblpPr (плавающие таблицы) – можно добавить позже, пока игнорируем

    # tblOverlap
    overlap_el = tblPr.find(qn("w:tblOverlap"))
    if overlap_el is not None:
        val = _str_attr(overlap_el, "val")
        if val:
            result["tblOverlap"] = val

    # tblW, jc, tblInd, tblCellSpacing
    for tag, key in (
        ("w:tblW", "tblW"),
        ("w:jc", "jc"),
        ("w:tblInd", "tblInd"),
        ("w:tblCellSpacing", "tblCellSpacing"),
    ):
        el = tblPr.find(qn(tag))
        if el is not None:
            if tag == "w:jc":
                val = _str_attr(el, "val")
                if val:
                    result[key] = val
            else:
                parsed = _parse_size_type(el)
                if parsed:
                    result[key] = parsed

    # tblBorders
    borders = _parse_borders(tblPr)
    if borders:
        result["tblBorders"] = borders

    # tblCellMar
    margins = _parse_cell_margins(tblPr)
    if margins:
        result["tblCellMar"] = margins

    # tblLayout
    layout_el = tblPr.find(qn("w:tblLayout"))
    if layout_el is not None:
        typ = _str_attr(layout_el, "type")
        if typ:
            result["tblLayout"] = typ

    # tblLook
    look_el = tblPr.find(qn("w:tblLook"))
    look = _parse_tbl_look(look_el)
    if look:
        result["tblLook"] = look

    # tblCaption, tblDescription
    for tag, key in (("w:tblCaption", "tblCaption"), ("w:tblDescription", "tblDescription")):
        el = tblPr.find(qn(tag))
        if el is not None:
            val = _str_attr(el, "val")
            if val:
                result[key] = val

    return result


def _parse_tr_pr(tr: etree._Element) -> Dict[str, Any]:
    """Парсит свойства строки (w:trPr)."""
    trPr = tr.find(qn("w:trPr"))
    if trPr is None:
        return {}

    result: Dict[str, Any] = {}

    # trHeight
    height_el = trPr.find(qn("w:trHeight"))
    if height_el is not None:
        val = _int_attr(height_el, "val")
        hRule = _str_attr(height_el, "hRule")
        if val is not None:
            height_dict = {"val": val}
            if hRule:
                height_dict["hRule"] = hRule
            result["trHeight"] = height_dict

    # tblHeader
    if trPr.find(qn("w:tblHeader")) is not None:
        result["tblHeader"] = True

    # cantSplit
    if trPr.find(qn("w:cantSplit")) is not None:
        result["cantSplit"] = True

    # jc (выравнивание строки)
    jc_el = trPr.find(qn("w:jc"))
    if jc_el is not None:
        val = _str_attr(jc_el, "val")
        if val:
            result["jc"] = val

    return result


def _parse_tc_pr(tc: etree._Element) -> Dict[str, Any]:
    """Парсит свойства ячейки (w:tcPr)."""
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        return {}

    result: Dict[str, Any] = {}

    # tcW
    w_el = tcPr.find(qn("w:tcW"))
    tcW = _parse_size_type(w_el)
    if tcW:
        result["tcW"] = tcW

    # gridSpan
    gridSpan_el = tcPr.find(qn("w:gridSpan"))
    if gridSpan_el is not None:
        val = _int_attr(gridSpan_el, "val")
        if val is not None and val > 1:
            result["gridSpan"] = val

    # vMerge
    vMerge_el = tcPr.find(qn("w:vMerge"))
    if vMerge_el is not None:
        # Согласно OOXML, элемент <w:vMerge> без атрибута означает restart
        val = _str_attr(vMerge_el, "val") or "restart"
        result["vMerge"] = val

    # tcBorders
    borders = _parse_borders(tcPr)
    if borders:
        result["tcBorders"] = borders

    # shd
    shd_el = tcPr.find(qn("w:shd"))
    shd = _parse_shd(shd_el)
    if shd:
        result["shd"] = shd

    # tcMar (отступы внутри ячейки)
    margins = _parse_cell_margins(tcPr)
    if margins:
        result["tcMar"] = margins

    # vAlign
    vAlign_el = tcPr.find(qn("w:vAlign"))
    if vAlign_el is not None:
        val = _str_attr(vAlign_el, "val")
        if val:
            result["vAlign"] = val

    # textDirection
    textDir_el = tcPr.find(qn("w:textDirection"))
    if textDir_el is not None:
        val = _str_attr(textDir_el, "val")
        if val:
            result["textDirection"] = val

    # tcFitText
    if tcPr.find(qn("w:tcFitText")) is not None:
        result["tcFitText"] = True

    # noWrap
    if tcPr.find(qn("w:noWrap")) is not None:
        result["noWrap"] = True

    return result


def _require_my_id(el: etree._Element, what: str) -> str:
    v = el.get(qn_my("id"))
    if v is None or not v.strip():
        raise ValueError(f"Missing required my:id for {what}")
    return v.strip()


def _validate_row_id(row_id: str, table_id: str) -> None:
    if not row_id.startswith(f"{table_id}.row_"):
        raise ValueError(f"Invalid row id '{row_id}': must start with '{table_id}.row_'")


def parse_table_node(parser, tbl: etree._Element, table_id: str) -> Dict[str, Any]:
    """
    Парсит элемент <w:tbl> и возвращает словарь таблицы по схеме v2.14.

    Аргументы:
        parser – экземпляр UltimateParserV43 (должен иметь методы
                 _parse_paragraph_element)
        tbl    – элемент <w:tbl> из document.xml
        table_id – идентификатор таблицы (например, "tbl_1")

    Возвращает:
        словарь с ключами type, id, tblPr, tbl_grid, rows
    """
    # Свойства таблицы
    tbl_pr = _parse_tbl_pr(tbl)

    # Сетка колонок (w:tblGrid)
    tbl_grid: List[int] = []
    tblGrid = tbl.find(qn("w:tblGrid"))
    if tblGrid is not None:
        for gridCol in tblGrid.findall(qn("w:gridCol")):
            w = _int_attr(gridCol, "w")
            if w is not None:
                tbl_grid.append(w)

    # Строки
    rows = []
    validate = os.environ.get("DOCX_PIPELINE_VALIDATE_RAW") == "1"
    for tr in tbl.findall(qn("w:tr")):
        # Контракт v2.15: id строк берём только из my:id, синтетика запрещена
        row_id = _require_my_id(tr, what=f"table row in {table_id}")
        if validate:
            _validate_row_id(row_id, table_id)

        tr_pr = _parse_tr_pr(tr)

        cells = []
        cell_counter = 1
        for tc in tr.findall(qn("w:tc")):
            cell_id = f"{row_id}.cell_{cell_counter}"
            tc_pr = _parse_tc_pr(tc)

            cell_content = []
            for idx, p in enumerate(tc.findall(qn("w:p")), start=1):
                cell_content.append(parser._parse_paragraph_element(p, parent_id=cell_id, index=idx))

            cell_obj: Dict[str, Any] = {
                "id": cell_id,
                "content": cell_content
            }
            if tc_pr:
                cell_obj["tcPr"] = tc_pr

            cells.append(cell_obj)
            cell_counter += 1

        row_obj: Dict[str, Any] = {
            "id": row_id,
            "cells": cells
        }
        if tr_pr:
            row_obj["trPr"] = tr_pr

        rows.append(row_obj)

    table_obj: Dict[str, Any] = {
        "type": "table",
        "id": table_id,
        "rows": rows
    }
    if tbl_pr:
        table_obj["tblPr"] = tbl_pr
    if tbl_grid:
        table_obj["tbl_grid"] = tbl_grid

    return table_obj