# reconstructor.py
# Minimal reconstructor with indexing and root deletion (step 2)
# Accepts donor DOCX file directly (extracts to temp dir)

from __future__ import annotations

import argparse
import json
import os
import zipfile
from copy import deepcopy
from typing import Any, Dict, List, Optional, Tuple
from lxml import etree
import sys
import re
import tempfile
import shutil

import reconstructor_table as rt
from reconstructor_picture import add_picture_to_document

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MY_NS = "https://translatefactory/schema/custom-id"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"  # for future use


def qn_w(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def qn_my(local: str) -> str:
    return f"{{{MY_NS}}}{local}"


class ReconstructorV215:
    def __init__(self, raw_json_path: str):
        self.raw_json_path = raw_json_path
        self.donor_raw_dir: Optional[str] = None
        self.package_files: Dict[str, bytes] = {}
        self.next_drawing_id: int = 1

        # JSON data
        with open(raw_json_path, "r", encoding="utf-8") as f:
            self.data = json.load(f)
        self.content: List[Dict[str, Any]] = self.data.get("content", [])

        # Indices to be built
        self.root_by_id: Dict[str, etree._Element] = {}
        self.row_by_id: Dict[str, etree._Element] = {}
        self.para_by_id: Dict[str, etree._Element] = {}
        self.original_children: List[etree._Element] = []

        self._new_id_pattern = re.compile(r'\.\d+$')

    def _copy_donor_files(self) -> None:
        """Copy all donor files except word/document.xml into package_files."""
        if not self.donor_raw_dir or not os.path.exists(self.donor_raw_dir):
            raise FileNotFoundError(f"donor_raw_dir not found: {self.donor_raw_dir}")

        exclude = {"word/document.xml"}
        for root, _, files in os.walk(self.donor_raw_dir):
            for fn in files:
                src_path = os.path.join(root, fn)
                rel_path = os.path.relpath(src_path, self.donor_raw_dir).replace("\\", "/")
                if rel_path in exclude:
                    continue
                if rel_path in self.package_files:
                    continue
                with open(src_path, "rb") as f:
                    self.package_files[rel_path] = f.read()

    def _build_indices(self, body: etree._Element) -> None:
        """
        Build indices of root elements, rows, and all paragraphs.
        Fills:
          self.root_by_id: root paragraphs and tables (by my:id)
          self.row_by_id: all table rows (by my:id)
          self.para_by_id: all paragraphs (root + nested) with generated ids
          self.original_children: list of body children (as in donor)
        """
        self.original_children = list(body)

        for elem in body:
            # Root paragraphs
            if elem.tag == qn_w("p"):
                pid = elem.get(qn_my("id"))
                if pid:
                    self.root_by_id[pid] = elem
                    self.para_by_id[pid] = elem
            # Root tables
            elif elem.tag == qn_w("tbl"):
                tid = elem.get(qn_my("id"))
                if tid:
                    self.root_by_id[tid] = elem
                    # Collect rows
                    for tr in elem.findall(qn_w("tr")):
                        row_id = tr.get(qn_my("id"))
                        if row_id:
                            self.row_by_id[row_id] = tr
                            # Process cells inside this row
                            cells = [tc for tc in tr if tc.tag == qn_w("tc")]
                            for ci, tc in enumerate(cells, start=1):
                                # Collect paragraphs inside this cell
                                paras = [p for p in tc if p.tag == qn_w("p")]
                                for pi, p in enumerate(paras, start=1):
                                    pid = f"{row_id}.cell_{ci}.p_{pi}"
                                    self.para_by_id[pid] = p

    def _apply_root_deletions(self, body: etree._Element) -> None:
        """
        Remove root elements marked as deleted in JSON.
        Modifies body children in place.
        """
        # Start with a copy of original children
        working = list(self.original_children)

        # Collect ids of elements to delete
        deleted_ids = set()
        for item in self.content:
            if item.get("deleted") is True:
                item_id = item.get("id")
                if not item_id:
                    raise ValueError("deleted item missing id")
                deleted_ids.add(item_id)

        # Remove elements from working list
        new_working = []
        for elem in working:
            # Check if this element (root) has an id and it's in deleted set
            elem_id = None
            if elem.tag in (qn_w("p"), qn_w("tbl")):
                elem_id = elem.get(qn_my("id"))
            if elem_id and elem_id in deleted_ids:
                # Skip this element (delete it)
                continue
            new_working.append(elem)

        # Replace body children with new list
        for ch in list(body):
            body.remove(ch)
        for ch in new_working:
            body.append(ch)

        # Optionally update original_children to reflect new state (for future steps)
        self.original_children = new_working

    def _needs_xml_preserve(self, text: str) -> bool:
        """Check if text needs xml:space='preserve' attribute."""
        if not text:
            return False
        return text[0] == " " or text[-1] == " " or "  " in text

    def _build_rPr(self, r: etree._Element, r_format: Optional[Dict[str, Any]]) -> None:
        """Build run properties (w:rPr) from r_format dict."""
        if not isinstance(r_format, dict) or not r_format:
            return
        
        rPr = etree.SubElement(r, qn_w("rPr"))

        # Boolean flags (presence elements)
        bool_flags = [
            ("bold", "b"),
            ("italic", "i"),
            ("strike", "strike"),
            ("double_strike", "dstrike"),
            ("caps", "caps"),
            ("small_caps", "smallCaps"),
        ]
        for json_key, xml_tag in bool_flags:
            if json_key in r_format and r_format[json_key]:
                etree.SubElement(rPr, qn_w(xml_tag))

        # Underline
        if "underline" in r_format and r_format["underline"] is not None:
            u = etree.SubElement(rPr, qn_w("u"))
            u.set(qn_w("val"), str(r_format["underline"]))

        # Color
        if "color" in r_format and r_format["color"] is not None:
            c = etree.SubElement(rPr, qn_w("color"))
            c.set(qn_w("val"), str(r_format["color"]))

        # Highlight
        if "highlight" in r_format and r_format["highlight"] is not None:
            h = etree.SubElement(rPr, qn_w("highlight"))
            h.set(qn_w("val"), str(r_format["highlight"]))

        # Vertical alignment
        if "vert_align" in r_format and r_format["vert_align"] is not None:
            va = etree.SubElement(rPr, qn_w("vertAlign"))
            va.set(qn_w("val"), str(r_format["vert_align"]))

        # Font size
        if "font_size_half_points" in r_format and r_format["font_size_half_points"] is not None:
            sz = etree.SubElement(rPr, qn_w("sz"))
            sz.set(qn_w("val"), str(int(r_format["font_size_half_points"])))

        # Fonts
        if "rFonts" in r_format and isinstance(r_format["rFonts"], dict):
            rf = etree.SubElement(rPr, qn_w("rFonts"))
            for a in ("ascii", "hAnsi", "eastAsia", "cs", 
                     "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "csTheme"):
                if a in r_format["rFonts"] and r_format["rFonts"][a] is not None:
                    rf.set(qn_w(a), str(r_format["rFonts"][a]))

        # Language
        if "lang" in r_format and r_format["lang"] is not None:
            l = etree.SubElement(rPr, qn_w("lang"))
            l.set(qn_w("val"), str(r_format["lang"]))

    def _build_run(self, run_data: Dict[str, Any], paragraph_id: str, run_index: int, next_drawing_id: int) -> Tuple[etree._Element, int]:
        """
        Build a single run (w:r) element from run data.
        
        Args:
            run_data: run data from JSON
            paragraph_id: ID of the parent paragraph
            run_index: 1-based index of the run
        
        Returns:
            etree._Element: constructed w:r element
        """
        r = etree.Element(qn_w("r"))
        
        # Set run ID if present in JSON (optional, for reference)
        if "id" in run_data:
            r.set(qn_my("id"), run_data["id"])
        
        # Build run properties
        r_format = run_data.get("r_format")
        if r_format:
            self._build_rPr(r, r_format)
        
        rtype = run_data.get("type")
        
        if rtype == "text":
            t = etree.SubElement(r, qn_w("t"))
            text = run_data.get("text", "")
            if self._needs_xml_preserve(text):
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = text
            
        elif rtype == "tab":
            etree.SubElement(r, qn_w("tab"))
            
        elif rtype == "break":
            br = etree.SubElement(r, qn_w("br"))
            br_type = run_data.get("break_type")
            if br_type:
                br.set(qn_w("type"), str(br_type))
                
        elif rtype == "cr":
            etree.SubElement(r, qn_w("cr"))
            
        elif rtype == "sym":
            sym = etree.SubElement(r, qn_w("sym"))
            if run_data.get("font") is not None:
                sym.set(qn_w("font"), str(run_data["font"]))
            if run_data.get("char") is not None:
                sym.set(qn_w("char"), str(run_data["char"]))

        elif rtype == "picture":
            next_drawing_id = add_picture_to_document(
                run_data=run_data,
                parent_element=r,
                next_drawing_id=next_drawing_id,
            )
            
        elif rtype == "shape":
            raise ValueError("Run type 'shape' is not supported")
            
        else:
            raise ValueError(f"Unsupported run type: {rtype}")
        
        return r, next_drawing_id

    def _patch_pPr(self, p: etree._Element, p_format: Optional[Dict[str, Any]]) -> None:
        """
        Хирургическое обновление свойств параграфа.
        Маппит значения из JSON в OOXML согласно схеме raw.schema.json v2.15.
        """
        if not isinstance(p_format, dict) or not p_format:
            return

        # Находим существующий pPr
        pPr = p.find(qn_w("pPr"))
        if pPr is None:
            pPr = etree.SubElement(p, qn_w("pPr"))

        # ========== ALIGNMENT (w:jc) ==========
        # Схема: alignmentEnum = ["left", "center", "right", "justify", "distribute"]
        # OOXML: left, center, right, both, distribute
        if "alignment" in p_format and p_format["alignment"] is not None:
            val = str(p_format["alignment"])
            # Маппинг: justify -> both
            if val == "justify":
                val = "both"
            # Остальные значения оставляем как есть (left, center, right, distribute)

            jc = pPr.find(qn_w("jc"))
            if jc is not None:
                jc.set(qn_w("val"), val)

        # ========== TEXT ALIGNMENT (w:textAlignment) ==========
        # Схема: textAlignmentEnum = ["auto", "baseline", "top", "center", "bottom"]
        # OOXML: auto, baseline, top, center, bottom (полное совпадение)
        if "text_alignment" in p_format and p_format["text_alignment"] is not None:
            val = str(p_format["text_alignment"])
            ta = pPr.find(qn_w("textAlignment"))
            if ta is not None:
                ta.set(qn_w("val"), val)

        # ========== INDENTS (w:ind) ==========
        # Все значения integer, маппинг не требуется
        indent_keys = {
            "indent_start_twip": "left",
            "indent_end_twip": "right",
            "indent_first_line_twip": "firstLine",
            "indent_hanging_twip": "hanging",
        }
        if any(k in p_format for k in indent_keys):
            ind = pPr.find(qn_w("ind"))
            if ind is None:
                ind = etree.SubElement(pPr, qn_w("ind"))

            for json_key, attr_name in indent_keys.items():
                if json_key in p_format and p_format[json_key] is not None:
                    value = p_format[json_key]
                    if isinstance(value, int):
                        ind.set(qn_w(attr_name), str(value))

        # ========== SPACING (w:spacing) ==========
        spacing_keys = {
            "space_before_twip": "before",
            "space_after_twip": "after",
            "line_spacing_twip": "line",
            "space_before_lines": "beforeLines",
            "space_after_lines": "afterLines",
            "before_autospacing": "beforeAutospacing",
            "after_autospacing": "afterAutospacing",
        }
        if any(k in p_format for k in spacing_keys) or "line_rule" in p_format:
            spacing = pPr.find(qn_w("spacing"))
            if spacing is None:
                spacing = etree.SubElement(pPr, qn_w("spacing"))

            # Числовые значения
            for json_key, attr_name in spacing_keys.items():
                if json_key in p_format and p_format[json_key] is not None:
                    value = p_format[json_key]
                    if isinstance(value, int):
                        spacing.set(qn_w(attr_name), str(value))
                    elif isinstance(value, bool):
                        spacing.set(qn_w(attr_name), "1" if value else "0")

            # line_rule (enum, но значения совпадают)
            if "line_rule" in p_format and p_format["line_rule"] is not None:
                val = str(p_format["line_rule"])
                # Схема: ["auto", "exact", "atLeast"] - всё совпадает с OOXML
                spacing.set(qn_w("lineRule"), val)

        # ========== BOOLEAN FLAGS ==========
        # Все boolean, только присутствие/отсутствие элемента
        bool_flags = {
            "keep_next": "keepNext",
            "keep_lines": "keepLines",
            "page_break_before": "pageBreakBefore",
            "widow_control": "widowControl",
            "contextual_spacing": "contextualSpacing",
            "snap_to_grid": "snapToGrid",
        }
        for json_key, tag in bool_flags.items():
            if json_key in p_format and p_format[json_key] is not None:
                tag_qn = qn_w(tag)
                if p_format[json_key]:
                    # Добавляем элемент если его нет
                    if pPr.find(tag_qn) is None:
                        etree.SubElement(pPr, tag_qn)
                else:
                    # Удаляем элемент если он есть
                    existing = pPr.find(tag_qn)
                    if existing is not None:
                        pPr.remove(existing)

        # ========== TABS (w:tabs) ==========
        # Схема: массив tabStop с полями posTwip, val, leader (опционально)
        if "tabs" in p_format and isinstance(p_format["tabs"], list):
            tabs_el = pPr.find(qn_w("tabs"))
            if tabs_el is None:
                tabs_el = etree.SubElement(pPr, qn_w("tabs"))
            else:
                # Очищаем существующие табуляции
                for tab in tabs_el.findall(qn_w("tab")):
                    tabs_el.remove(tab)

            for tab_item in p_format["tabs"]:
                if not isinstance(tab_item, dict):
                    continue

                tab = etree.SubElement(tabs_el, qn_w("tab"))

                # posTwip - обязательное поле
                if "posTwip" in tab_item and tab_item["posTwip"] is not None:
                    tab.set(qn_w("pos"), str(tab_item["posTwip"]))

                # val - обязательное поле (значения из OOXML: left, center, right, decimal, bar, clear, end, num, start)
                if "val" in tab_item and tab_item["val"] is not None:
                    tab.set(qn_w("val"), str(tab_item["val"]))

                # leader - опционально (none, dot, hyphen, underscore, middleDot, etc.)
                if "leader" in tab_item and tab_item["leader"] is not None:
                    tab.set(qn_w("leader"), str(tab_item["leader"]))

        # ========== NUMBERING (w:numPr) ==========
        # Схема: list_info = {"numId": string, "ilvl": string}
        if "list_info" in p_format and isinstance(p_format["list_info"], dict):
            list_info = p_format["list_info"]
            numPr = pPr.find(qn_w("numPr"))
            if numPr is None:
                numPr = etree.SubElement(pPr, qn_w("numPr"))

            # ilvl
            if "ilvl" in list_info and list_info["ilvl"] is not None:
                ilvl = numPr.find(qn_w("ilvl"))
                if ilvl is None:
                    ilvl = etree.SubElement(numPr, qn_w("ilvl"))
                ilvl.set(qn_w("val"), str(list_info["ilvl"]))

            # numId
            if "numId" in list_info and list_info["numId"] is not None:
                numId = numPr.find(qn_w("numId"))
                if numId is None:
                    numId = etree.SubElement(numPr, qn_w("numId"))
                numId.set(qn_w("val"), str(list_info["numId"]))

        # ========== ВСЁ! НИКАКОЙ ПЕРЕСТРОЙКИ ПОРЯДКА ==========

    def _process_paragraph(self, p_el: etree._Element, p_json: Dict[str, Any]) -> None:
        """
        Process a single paragraph: update pPr and rebuild runs.
        """
        # Update paragraph properties if provided
        if "p_format" in p_json:
            self._patch_pPr(p_el, p_json.get("p_format"))

        # Rebuild runs
        runs = p_json.get("runs", [])

        # Remove all existing runs
        for r in p_el.findall(qn_w("r")):
            p_el.remove(r)

        for run_idx, run_data in enumerate(runs, start=1):
            new_run, self.next_drawing_id = self._build_run(run_data, p_json.get("id"), run_idx,
                                                            self.next_drawing_id)
            p_el.append(new_run)

    def _process_table(self, tbl: etree._Element, tbl_json: Dict[str, Any]) -> None:
        """
        Process a table according to JSON order and patches.
        Applies table-level updates (tblPr/tbl_grid), row/cell property patches
        (trPr/tcPr), paragraph operations inside cells, and run rebuild for
        cell paragraphs.
        """


        rows_json = tbl_json.get("rows", [])

        # Обновляем свойства таблицы
        rt.patch_tblPr(tbl, tbl_json.get("tblPr"))
        rt.patch_tbl_grid(tbl, tbl_json.get("tbl_grid"))
        if not isinstance(rows_json, list):
            raise ValueError(f"Table '{tbl_json.get('id')}' rows must be array")

        new_rows: List[etree._Element] = []
        
        for row_item in rows_json:
            if not isinstance(row_item, dict):
                raise ValueError(f"Table row must be object")

            row_id = row_item.get("id")
            if not row_id:
                raise ValueError(f"Table row missing id")

            # Check if this is a new row (id contains .digits at end)
            is_new_row = self._is_new_id(str(row_id))
            
            if is_new_row:
                derive_from = row_item.get("derive_from")
                if not derive_from:
                    raise ValueError(f"New row '{row_id}' missing derive_from")
                    
                src_tr = self.row_by_id.get(str(derive_from))
                if src_tr is None:
                    raise ValueError(f"derive_from row '{derive_from}' not found")
                    
                tr = rt.clone_row(src_tr)
                tr.set(qn_my("id"), row_id)
                
                # Add to indices for future references
                self.row_by_id[row_id] = tr
            else:
                tr = self.row_by_id.get(row_id)
                if tr is None:
                    raise ValueError(f"Row '{row_id}' not found in donor")

            # Обновляем свойства строки
            rt.patch_trPr(tr, row_item.get("trPr"))

            # Process cells in this row
            cells_json = row_item.get("cells", [])
            if not isinstance(cells_json, list):
                raise ValueError(f"Row '{row_id}' cells must be array")

            tcs = rt.iter_row_cells(tr)
            if len(tcs) != len(cells_json):
                raise ValueError(
                    f"Row '{row_id}' cell count mismatch "
                    f"(donor={len(tcs)} json={len(cells_json)})"
                )

            for ci, (tc, cell_item) in enumerate(zip(tcs, cells_json), start=1):
                if not isinstance(cell_item, dict):
                    raise ValueError(f"Row '{row_id}' cell[{ci}] must be object")

                # Обновляем свойства ячейки
                rt.patch_tcPr(tc, cell_item.get("tcPr"))

                # Process paragraphs inside cell
                paras_json = cell_item.get("content", [])

                if not isinstance(paras_json, list):
                    raise ValueError(f"Row '{row_id}' cell[{ci}] content must be array")


                planned, final_paras = rt.apply_cell_paragraph_ops(
                    tc,
                    row_id=row_id,
                    cell_index_1based=ci,
                    json_paragraphs=paras_json,
                    clone_paragraph=lambda p: deepcopy(p),
                )

                rt.replace_cell_paragraphs(tc, final_paras)

                # Process each paragraph in the cell (rebuild runs)
                for pid, p_el, p_json in planned:
                    if p_json.get("type") != "paragraph":
                        raise ValueError(f"Cell content item '{pid}' must be paragraph")

                    # Apply paragraph formatting for nested cell paragraph first.
                    self._patch_pPr(p_el, p_json.get("p_format"))

                    # Rebuild runs for this paragraph
                    runs = p_json.get("runs", [])
                    for r in p_el.findall(qn_w("r")):
                        p_el.remove(r)

                    for run_idx, run_data in enumerate(runs, start=1):
                        new_run, self.next_drawing_id = self._build_run(run_data, pid, run_idx,
                                                                        self.next_drawing_id)
                        p_el.append(new_run)

            new_rows.append(tr)

        # Replace all rows in the table
        rt.set_table_rows(tbl, new_rows)

    def _process_content(self, body: etree._Element) -> None:
        """
        Process all content items after deletions/insertions.
        Tables are processed first (they contain paragraphs), then root paragraphs.
        """

        tables_to_process = []
        paragraphs_to_process = []

        for item in self.content:
            if item.get("deleted") is True:
                continue
                
            if item.get("type") == "table":
                tbl_el = self.root_by_id.get(item.get("id"))
                if tbl_el is not None:
                    tables_to_process.append((tbl_el, item))
            elif item.get("type") == "paragraph":
                p_el = self.para_by_id.get(item.get("id"))
                if p_el is not None:
                    paragraphs_to_process.append((p_el, item))

        # Process tables first
        for tbl_el, tbl_json in tables_to_process:
            self._process_table(tbl_el, tbl_json)

        # Then process root paragraphs
        for p_el, p_json in paragraphs_to_process:
            # _process_paragraph теперь сам перестраивает run'ы
            self._process_paragraph(p_el, p_json)


    def _is_new_id(self, elem_id: str) -> bool:
        """Check if element id indicates a newly created element (contains .digits at end)."""
        # New elements are identified by suffix pattern ".N" (digits at end),
        # e.g. p_1.1, tbl_2.3, tbl_1.row_4.1. This convention is used for
        # root paragraphs, root tables, and table rows.
        return bool(self._new_id_pattern.search(str(elem_id)))

    def _apply_root_insertions_and_moves(self, body: etree._Element) -> None:
        """
        Insert new root elements and move existing ones according to JSON.
        Modifies body children in place.
        """
        # Start with current children (after deletions)
        working = list(self.original_children)

        # Group items by (anchor, position)
        groups = {}
        ordered_items = []  # сохраняем порядок для последующей вставки

        for item in self.content:
            if item.get("deleted") is True:
                continue

            item_id = item.get("id")
            if not item_id:
                raise ValueError("content item missing id")

            anchor = item.get("anchor")
            position = item.get("position")

            # Skip items without positioning
            if anchor is None or position is None:
                continue

            key = (anchor, position)
            if key not in groups:
                groups[key] = []
                ordered_items.append(key)

            groups[key].append(item)

        # Process each group in the order they first appeared
        for anchor, position in ordered_items:
            items = groups[(anchor, position)]

            # Find anchor element
            anchor_idx = -1
            for i, elem in enumerate(working):
                if elem.tag in (qn_w("p"), qn_w("tbl")):
                    if elem.get(qn_my("id")) == anchor:
                        anchor_idx = i
                        break

            if anchor_idx == -1:
                raise ValueError(f"Anchor element '{anchor}' not found")

            # Process items in the group in JSON order
            for item in items:
                item_id = item.get("id")

                # Get or create the element to insert
                if self._is_new_id(item_id):
                    # New element - must have derive_from
                    derive_from = item.get("derive_from")
                    if not derive_from:
                        raise ValueError(f"New element '{item_id}' missing derive_from")

                    src_elem = self.root_by_id.get(derive_from)
                    if src_elem is None:
                        raise ValueError(f"Source element '{derive_from}' not found")

                    new_elem = deepcopy(src_elem)
                    new_elem.set(qn_my("id"), item_id)

                    if new_elem.tag == qn_w("p"):
                        # Paragraph path: rebuild runs from JSON
                        for r in new_elem.findall(qn_w("r")):
                            new_elem.remove(r)

                        runs_data = item.get("runs", [])
                        for run_idx, run_data in enumerate(runs_data, start=1):
                            new_run, self.next_drawing_id = self._build_run(
                                run_data,
                                item_id,
                                run_idx,
                                self.next_drawing_id,
                            )
                            new_elem.append(new_run)
                    elif new_elem.tag == qn_w("tbl"):
                        # Table path: keep cloned table structure untouched here.
                        # Table rows/cells/paragraphs are processed later in _process_table.
                        pass
                    else:
                        raise ValueError(
                            f"Unsupported root element tag for new element '{item_id}': {new_elem.tag}"
                        )

                    # Add to indices
                    self.root_by_id[item_id] = new_elem
                    if new_elem.tag == qn_w("p"):
                        self.para_by_id[item_id] = new_elem

                    elem_to_insert = new_elem
                else:
                    # Existing element - find and remove it
                    elem_to_insert = None
                    for i, elem in enumerate(working):
                        if elem.tag in (qn_w("p"), qn_w("tbl")):
                            if elem.get(qn_my("id")) == item_id:
                                elem_to_insert = elem
                                # Remove from current position
                                working.pop(i)
                                # Adjust anchor index if we removed before it
                                if i < anchor_idx:
                                    anchor_idx -= 1
                                break

                    if elem_to_insert is None:
                        raise ValueError(f"Existing element '{item_id}' not found")

                # Insert at position
                if position == "before":
                    working.insert(anchor_idx, elem_to_insert)
                    anchor_idx += 1  # сдвигаем якорь для следующих элементов в группе
                else:  # after
                    working.insert(anchor_idx + 1, elem_to_insert)
                    anchor_idx += 1  # сдвигаем для следующей вставки

        # Update body children
        for ch in list(body):
            body.remove(ch)
        for ch in working:
            body.append(ch)

        # Update original_children for future steps
        self.original_children = working

    def _init_next_drawing_id(self, doc_root: etree._Element) -> int:
        """Находит максимальный id среди wp:docPr и возвращает следующий."""
        ids: List[int] = []
        for el in doc_root.findall(f".//{{{WP_NS}}}docPr"):
            v = el.get("id")
            if v and v.isdigit():
                ids.append(int(v))
        return (max(ids) + 1) if ids else 1

    def build_docx(self, out_docx_path: str, donor_docx_path: str) -> None:
        """
        Build reconstructed DOCX from donor DOCX and JSON.

        Args:
            out_docx_path: path to output DOCX file
            donor_docx_path: path to donor DOCX file (with my:id attributes)
        """

        # Create temporary directory for donor extraction
        self.temp_dir = tempfile.mkdtemp(prefix="reconstructor_")
        try:
            # Extract donor DOCX to temp directory
            with zipfile.ZipFile(donor_docx_path, "r") as z:
                z.extractall(self.temp_dir)

            self.donor_raw_dir = self.temp_dir

            # Copy all donor files except word/document.xml into package_files
            self._copy_donor_files()

            # Load donor document.xml
            donor_doc_path = os.path.join(self.temp_dir, "word", "document.xml")
            if not os.path.exists(donor_doc_path):
                raise FileNotFoundError(f"word/document.xml not found in {donor_docx_path}")

            parser = etree.XMLParser(remove_blank_text=False, recover=False, huge_tree=True)
            doc_tree = etree.parse(donor_doc_path, parser)
            doc_root = doc_tree.getroot()
            body = doc_root.find(qn_w("body"))
            if body is None:
                raise ValueError("Document has no w:body")

            # Build indices (does not modify XML)
            self._build_indices(body)

            # Apply root deletions (modifies body)
            self._apply_root_deletions(body)

            # Инициализируем next_drawing_id из существующих картинок
            self.next_drawing_id = self._init_next_drawing_id(doc_root)

            # Apply root insertions and moves (modifies body)
            self._apply_root_insertions_and_moves(body)

            # Process content - tables first, then paragraphs
            self._process_content(body)

            # Serialize document.xml (now with deletions) and add to package
            self.package_files["word/document.xml"] = etree.tostring(
                doc_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=None,
                pretty_print=False
            )

            # Create output zip
            os.makedirs(os.path.dirname(out_docx_path), exist_ok=True)
            with zipfile.ZipFile(out_docx_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for name in sorted(self.package_files.keys()):
                    zout.writestr(name, self.package_files[name])

        finally:
            # Clean up temporary directory
            if self.temp_dir and os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir)

    def _cleanup(self) -> None:
        """Clean up temporary directory if it exists."""
        if self.temp_dir and os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
            self.temp_dir = None


def main() -> None:
    parser = argparse.ArgumentParser(description="Reconstructor v2.15 (step 2: root deletions)")
    parser.add_argument("--in-json", dest="input_json", required=True)
    parser.add_argument("--out-docx", dest="output_docx", required=True)
    parser.add_argument("--donor-docx", dest="donor_docx", required=True,
                        help="Path to donor DOCX file (with my:id attributes)")
    args = parser.parse_args()

    try:
        recon = ReconstructorV215(args.input_json)
        recon.build_docx(args.output_docx, args.donor_docx)
    except Exception:
        import traceback
        traceback.print_exc()
        if hasattr(recon, '_cleanup'):
            recon._cleanup()
        sys.exit(1)


if __name__ == "__main__":
    main()
