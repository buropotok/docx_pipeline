# parser_table.py
# Парсинг таблиц DOCX (w:tbl) в RAW JSON (схема v2.13)

from typing import Any, Dict
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(tag: str) -> str:
    """Преобразует тег вида w:name в полное имя с пространством имён."""
    pfx, local = tag.split(":")
    if pfx != "w":
        raise ValueError(f"Unsupported prefix: {pfx}")
    return f"{{{W_NS}}}{local}"


def _str_attr(el: etree._Element, attr_local: str) -> str | None:
    """Безопасно получает строковый атрибут элемента."""
    return el.get(f"{{{W_NS}}}{attr_local}")


def parse_table_node(parser, tbl: etree._Element) -> Dict[str, Any]:
    """
    Парсит элемент <w:tbl> и возвращает словарь таблицы по схеме v2.13.

    Аргументы:
        parser – экземпляр UltimateParserV43 (должен иметь методы
                 _parse_paragraph_element и атрибуты _table_counter, _row_counter)
        tbl    – элемент <w:tbl> из document.xml

    Возвращает:
        словарь с ключами type, id, rows (и опционально tbl_style_id)
    """
    table_id = f"tbl_{parser._table_counter}"
    parser._table_counter += 1

    tbl_style_id = None
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is not None:
        tblStyle = tblPr.find(qn("w:tblStyle"))
        if tblStyle is not None:
            tbl_style_id = _str_attr(tblStyle, "val")

    rows = []
    for tr in tbl.findall(qn("w:tr")):
        row_id = f"row_{parser._row_counter}"
        parser._row_counter += 1

        cells = []
        for tc in tr.findall(qn("w:tc")):
            cell_content = []
            for p in tc.findall(qn("w:p")):
                # Используем метод основного парсера для обработки абзаца
                cell_content.append(parser._parse_paragraph_element(p))

            # ID ячейки пока не требуется, можно добавить позже
            cell = {"content": cell_content}
            cells.append(cell)

        rows.append({
            "id": row_id,
            "cells": cells
        })

    table_obj: Dict[str, Any] = {
        "type": "table",
        "id": table_id,
        "rows": rows
    }
    if tbl_style_id:
        table_obj["tbl_style_id"] = tbl_style_id

    return table_obj