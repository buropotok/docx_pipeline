# UltimateReconstructorV12.py
# Reconstructor Version: v12
# Schema Version: 2.13
# Rules Version: 0.3

# RAW JSON (v2.13) -> DOCX reconstructor using lxml (NO python-docx)
# Target: visually deterministic for "forms" subset (no tables/images/fields/hyperlinks)
#
# ГИБРИДНЫЙ ПОДХОД:
# - Все файлы, кроме document.xml, копируются из донора (raw/donor)
# - document.xml модифицируется на основе JSON (список абзацев и run'ов)
# - Все остальные файлы (styles.xml, numbering.xml, settings.xml, отношения, медиа) остаются оригинальными

from __future__ import annotations

import argparse
import json
import os
import zipfile
import shutil
from typing import Any, Dict, Optional, List, Tuple

from lxml import etree
import sys

sys.path.insert(0, os.path.dirname(__file__))
from reconstructor_picture import add_picture_to_document

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

BASE_DIR = "."

# Default Normal values (same as in parser) - используются только как запасной вариант
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
    "color": "auto",
    "vertical_align": "baseline",
    "all_caps": False,
    "charSpacingTwip": 0,
    "positionHalfPoints": 0,
    "lang": {"eastAsia": "zh-CN"}
}


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


def _merge(a: Dict[str, Any], b: Dict[str, Any]) -> Dict[str, Any]:
    out = dict(a)
    for k, v in b.items():
        if v is None:
            continue
        out[k] = v
    return out


def _normalize_p_format(p_format: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    if not isinstance(p_format, dict):
        return {}
    out = dict(p_format)
    key_map = {
        # Существующие поля
        "line_spacing_twip": "lineTwip",
        "line_rule": "lineRule",
        "space_before_twip": "spaceBeforeTwip",
        "space_after_twip": "spaceAfterTwip",
        "indent_start_twip": "indentStartTwip",
        "indent_end_twip": "indentEndTwip",
        "indent_first_line_twip": "indentFirstLineTwip",
        "indent_hanging_twip": "indentHangingTwip",
        "keep_next": "keepNext",
        "keep_lines": "keepLines",
        "page_break_before": "pageBreakBefore",
        "widow_control": "widowControl",
        "text_alignment": "textAlignment",
        "list_info": "numbering",

        # Добавить эти поля
        "contextual_spacing": "contextualSpacing",
        "snap_to_grid": "snapToGrid",
        "before_autospacing": "beforeAutospacing",
        "after_autospacing": "afterAutospacing",
        "space_before_lines": "spaceBeforeLines",
        "space_after_lines": "spaceAfterLines",
    }
    line_rule_map = {"auto": "AUTO", "exact": "EXACT", "atLeast": "AT_LEAST"}
    text_alignment_map = {
        "auto": "AUTO",
        "baseline": "BASELINE",
        "top": "TOP",
        "center": "CENTER",
        "bottom": "BOTTOM",
    }
    for src, dst in key_map.items():
        if src in p_format and dst not in out:
            out[dst] = p_format[src]
    lr = out.get("lineRule")
    if isinstance(lr, str):
        out["lineRule"] = line_rule_map.get(lr, lr)
    ta = out.get("textAlignment")
    if isinstance(ta, str):
        out["textAlignment"] = text_alignment_map.get(ta, ta)
    return out


def _normalize_r_format(r_format: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    if not isinstance(r_format, dict):
        return {}
    out = dict(r_format)
    key_map = {
        "vert_align": "vertical_align",
        "caps": "all_caps",
        "spacing_twip": "charSpacingTwip",
        "position_half_points": "positionHalfPoints",
    }
    for src, dst in key_map.items():
        if src in r_format and dst not in out:
            out[dst] = r_format[src]
    return out


class UltimateReconstructorV12:
    def __init__(self, raw_json_path: str):
        self.raw_json_path = raw_json_path
        with open(raw_json_path, "r", encoding="utf-8") as f:
            self.data = json.load(f)

        # validate minimal keys (soft)
        for k in ("meta", "document_info", "numbering_definitions", "styles", "content"):
            if k not in self.data:
                raise ValueError(f"RAW JSON missing required key: {k}")

        self.meta: Dict[str, Any] = self.data.get("meta", {})
        self.document_info: Dict[str, Any] = self.data.get("document_info", {})
        self.numbering_definitions: Dict[str, Any] = self.data.get("numbering_definitions", {})
        self.styles: Dict[str, Any] = self.data.get("styles", {})  # все стили (paragraph, character, ...)
        self.doc_defaults: Dict[str, Any] = self.data.get("doc_defaults", {})
        self.latent_styles: Dict[str, Any] = self.data.get("latent_styles", {})
        self.content: List[Dict[str, Any]] = self.data.get("content", [])
        self.donor_raw_dir: Optional[str] = None  # будет установлен в build_docx

        self.default_style_id: Optional[str] = self.meta.get("default_style_id")

        # Prepare mapping: style_id -> word_style_id
        self.style_id_to_word: Dict[str, str] = {}
        for sid, st in self.styles.items():
            self.style_id_to_word[sid] = st.get("word_style_id", sid)

        # Build map for numbering pStyle replacement (if needed)
        self._prepare_numbering_pstyles()

        # Build map: style_id -> (numId, ilvl) for styles referenced in numbering levels via pStyle
        self.style_to_num: Dict[str, Tuple[str, int]] = {}
        for numId, rec in self.numbering_definitions.items():
            levels = rec.get("levels", {})
            for ilvl_str, lvl in levels.items():
                pStyle = lvl.get("pStyle")
                if pStyle and isinstance(pStyle, str):
                    try:
                        ilvl = int(ilvl_str)
                        if pStyle not in self.style_to_num:
                            self.style_to_num[pStyle] = (str(numId), ilvl)
                    except ValueError:
                        pass

        # Для работы с изображениями
        self.relationships: Dict[str, str] = {}  # будет заполняться при добавлении картинок
        self.next_rel_id = 4  # счётчик для новых отношений
        self.donor_media_path: Optional[str] = None  # установим позже

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

    def _prepare_numbering_pstyles(self) -> None:
        """
        If numbering_definitions contain pStyle values that are source_word_style_id,
        replace them with the corresponding word_style_id from our styles.
        This is idempotent: if already replaced, fine.
        """
        # Build reverse map: source_word_style_id -> word_style_id
        src_to_word = {}
        for st in self.styles.values():
            src = st.get("source_word_style_id")
            if src:
                src_to_word[src] = st.get("word_style_id", src)

        # Traverse numbering and replace
        for num_rec in self.numbering_definitions.values():
            levels = num_rec.get("levels", {})
            for lvl_rec in levels.values():
                ps = lvl_rec.get("pStyle")
                if ps and ps in src_to_word:
                    lvl_rec["pStyle"] = src_to_word[ps]

    def _copy_donor_files(self, package_files: Dict[str, bytes]) -> None:
        """
        Копирует все файлы из директории донора (raw/donor) в выходной пакет,
        за исключением document.xml, который будет сгенерирован отдельно.
        """
        if not self.donor_raw_dir or not os.path.exists(self.donor_raw_dir):
            print("[recon] Warning: donor_raw_dir not found, skipping file copy")
            return

        # Копируем все файлы, кроме document.xml
        exclude_files = {"word/document.xml"}

        for root, dirs, files in os.walk(self.donor_raw_dir):
            for file in files:
                src_path = os.path.join(root, file)
                # Относительный путь от donor_raw_dir
                rel_path = os.path.relpath(src_path, self.donor_raw_dir)
                # Нормализуем разделители для zip
                rel_path = rel_path.replace("\\", "/")

                # Пропускаем document.xml
                if rel_path in exclude_files:
                    continue

                # Пропускаем файлы, которые уже есть в package_files (на случай коллизий)
                if rel_path in package_files:
                    continue

                try:
                    with open(src_path, "rb") as f:
                        package_files[rel_path] = f.read()
                    print(f"[recon] Copied donor file: {rel_path}")
                except Exception as e:
                    print(f"[recon] Warning: failed to copy {rel_path}: {e}")

        # Копируем также медиафайлы, если они есть
        if self.donor_media_path and os.path.exists(self.donor_media_path):
            for file in os.listdir(self.donor_media_path):
                src = os.path.join(self.donor_media_path, file)
                if os.path.isfile(src):
                    with open(src, "rb") as f:
                        package_files[f"word/media/{file}"] = f.read()
                    print(f"[recon] Copied media file: {file}")

    # =========================
    # PUBLIC
    # =========================

    def build_docx(self, out_docx_path: str) -> None:
        package_files: Dict[str, bytes] = {}
        self.package_files = package_files

        # Определяем пути к донорским файлам
        run_dir = os.path.dirname(os.path.dirname(self.raw_json_path))
        donor_raw_dir = os.path.join(run_dir, "raw", "materialized")
        self.donor_raw_dir = donor_raw_dir
        self.donor_media_path = os.path.join(donor_raw_dir, "word", "media")

        print(f"[recon] donor_raw_dir = {donor_raw_dir}")
        print(f"[recon] donor_media_path = {self.donor_media_path}")

        # ШАГ 1: Копируем все файлы донора (кроме document.xml)
        self._copy_donor_files(package_files)

        # ШАГ 2: Загружаем оригинальный document.xml донора
        donor_doc_path = os.path.join(donor_raw_dir, "word", "document.xml")
        if not os.path.exists(donor_doc_path):
            raise FileNotFoundError(f"Donor document.xml not found at {donor_doc_path}")

        parser = etree.XMLParser(remove_blank_text=True)
        doc_tree = etree.parse(donor_doc_path, parser)
        doc_root = doc_tree.getroot()

        # ШАГ 3: Модифицируем document.xml согласно JSON
        self._modify_document(doc_root, self.content)

        # ШАГ 4: Сохраняем модифицированный document.xml
        package_files["word/document.xml"] = self._serialize_xml(doc_root, standalone=True)

        # ШАГ 5: Если есть новые изображения, обновляем document.xml.rels
        # (этот функционал остаётся как был, через add_picture_to_document)
        # ВАЖНО: при добавлении новых изображений нужно обновлять relationships
        # которые уже были скопированы из донора

        # Сохраняем для отладки
        raw_reconstructed_dir = os.path.join(os.path.dirname(self.raw_json_path), "raw", "reconstructed")
        self._dump_reconstructed_parts(package_files, raw_reconstructed_dir)

        # Собираем DOCX
        os.makedirs(os.path.dirname(out_docx_path), exist_ok=True)
        with zipfile.ZipFile(out_docx_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
            for name in sorted(package_files.keys()):
                z.writestr(name, package_files[name])

        print("[recon] DOCX successfully created")
        print("[recon] package_files keys:", list(package_files.keys()))

    # =========================
    # DOCUMENT MODIFICATION
    # =========================

    def _modify_document(self, root: etree._Element, content: List[Dict[str, Any]]) -> None:
        """
        Модифицирует document.xml на основе JSON-контента.
        Заменяет все абзацы в <w:body> на новые из JSON.
        Сохраняет все остальные элементы (sectPr и пр.)
        """
        # Находим body
        body = root.find(qn_w("body"))
        if body is None:
            raise ValueError("Document has no body element")

        # Индексируем оригинальные дочерние элементы
        orig_children = list(body)
        paragraph_counter = 1
        orig_paragraphs_by_id: Dict[str, etree._Element] = {}
        paragraph_index_by_id: Dict[str, int] = {}
        non_paragraphs: List[etree._Element] = []
        non_paragraph_indices: List[int] = []

        for idx, elem in enumerate(orig_children):
            if elem.tag == qn_w("p"):
                pid = f"p_{paragraph_counter}"
                orig_paragraphs_by_id[pid] = elem
                paragraph_index_by_id[pid] = idx
                paragraph_counter += 1
            else:
                non_paragraphs.append(elem)
                non_paragraph_indices.append(idx)

        # Модифицируем существующие абзацы (id без точки)
        modified_paragraphs: Dict[str, etree._Element] = {}
        for p_item in content:
            pid = p_item.get("id")
            if not pid:
                raise ValueError(f"Paragraph missing 'id' field: {p_item}")
            if '.' not in pid:
                orig_p = orig_paragraphs_by_id.get(pid)
                if orig_p is None:
                    raise ValueError(f"Original paragraph with id {pid} not found")
                self._modify_paragraph(orig_p, p_item)
                modified_paragraphs[pid] = orig_p

        # Собираем новые абзацы (id с точкой)
        new_paragraphs: List[Tuple[str, etree._Element]] = []  # (base_id, element)
        for p_item in content:
            pid = p_item.get("id")
            if '.' in pid:
                base_id = pid.split('.')[0]  # например, "p_1"
                new_p = self._build_paragraph(p_item)
                new_paragraphs.append((base_id, new_p))

        # Строим итоговый список children, сохраняя порядок
        final_children: List[etree._Element] = []
        # Проходим по оригинальным элементам в порядке
        for idx, elem in enumerate(orig_children):
            if idx in non_paragraph_indices:
                final_children.append(elem)  # не-абзацный элемент
            else:
                # Это абзац, находим его id
                # (можно было бы сохранить pid при первом проходе, но проще вычислить заново)
                # Для простоты будем использовать словарь paragraph_index_by_id, обратный индекс
                # Но нам нужен pid по индексу. Создадим mapping index -> pid
                pass

        # Упрощённый подход: сначала скопируем все оригинальные элементы,
        # затем заменим абзацы на модифицированные, а потом вставим новые.
        final_children = list(orig_children)  # копия

        # Заменяем существующие абзацы на модифицированные
        for pid, mod_p in modified_paragraphs.items():
            idx = paragraph_index_by_id[pid]
            final_children[idx] = mod_p

        # Вставляем новые абзацы после базовых
        # Группируем по base_id и сортируем по суффиксу (для стабильности)
        from collections import defaultdict
        new_by_base = defaultdict(list)
        for base_id, new_p in new_paragraphs:
            new_by_base[base_id].append(new_p)
        for base_id, new_list in new_by_base.items():
            # Определяем индекс вставки: после базового абзаца
            if base_id in paragraph_index_by_id:
                base_idx = paragraph_index_by_id[base_id]
            else:
                # Если базового нет (например, p_0), вставляем в начало
                base_idx = -1
            # Вставляем после base_idx, сдвигая индексы
            insert_pos = base_idx + 1
            for new_p in new_list:
                final_children.insert(insert_pos, new_p)
                insert_pos += 1
                # Сдвигаем индексы всех последующих элементов (не требуется, если не используем индексы далее)

        # Очищаем body и добавляем новые children
        for child in list(body):
            body.remove(child)
        for elem in final_children:
            body.append(elem)

    def _build_paragraph(self, p_item: Dict[str, Any]) -> etree._Element:
        """
        Строит элемент <w:p> из JSON-описания абзаца.
        """
        p = _w_el("p")

        style_id = p_item.get("p_style_id") or p_item.get("style_id")
        runs = p_item.get("runs", [])

        # Создаём pPr
        pPr = _w_sub(p, "pPr")

        # Добавляем ссылку на стиль
        if style_id:
            word_id = self.style_id_to_word.get(style_id, style_id)
            pStyle = _w_sub(pPr, "pStyle")
            _set_w_attr(pStyle, "val", word_id)

        # Добавляем локальное форматирование абзаца
        local_p_format = _normalize_p_format(p_item.get("p_format"))
        if local_p_format:
            local_pPr = self._build_pPr(local_p_format)
            if local_pPr is not None:
                for child in local_pPr:
                    pPr.append(child)

        # Добавляем run'ы
        for run in runs:
            r = self._build_run(run)
            if r is not None:
                p.append(r)

        return p

    def _build_run(self, run: Dict[str, Any]) -> Optional[etree._Element]:
        """
        Строит элемент <w:r> из JSON-описания run.
        """
        rtype = run.get("type")

        # Обработка разных типов run
        if rtype == "text":
            r = _w_el("r")
            txt = run.get("text", "")

            # Добавляем форматирование
            run_style_id = run.get("r_style_id") or run.get("char_style_id")
            run_local_r = _normalize_r_format(run.get("r_format"))
            if run_style_id or run_local_r:
                rPr_el = _w_sub(r, "rPr")
                if run_style_id:
                    rStyle = _w_sub(rPr_el, "rStyle")
                    word_id = self.style_id_to_word.get(run_style_id, run_style_id)
                    _set_w_attr(rStyle, "val", word_id)
                if run_local_r:
                    local_rPr = self._build_rPr(run_local_r)
                    if local_rPr is not None:
                        for child in local_rPr:
                            rPr_el.append(child)

            # Добавляем текст
            t = _w_sub(r, "t")
            if _needs_xml_preserve(txt):
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = txt
            return r

        elif rtype == "tab":
            r = _w_el("r")
            _w_sub(r, "tab")
            return r

        elif rtype == "break":
            r = _w_el("r")
            br = _w_sub(r, "br")
            bt = run.get("break_type")
            if bt in ("textWrapping", "page", "column"):
                _set_w_attr(br, "type", bt)
            return r

        elif rtype == "cr":
            r = _w_el("r")
            _w_sub(r, "cr")
            return r

        elif rtype == "sym":
            r = _w_el("r")
            sym_text = run.get("text", "")
            t = _w_sub(r, "t")
            if _needs_xml_preserve(sym_text):
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = sym_text
            return r

        elif rtype == "picture":
            if self.donor_media_path is None:
                return None
            r = _w_el("r")
            run_style_id = run.get("r_style_id") or run.get("char_style_id")
            if run_style_id:
                rPr_el = _w_sub(r, "rPr")
                rStyle = _w_sub(rPr_el, "rStyle")
                word_id = self.style_id_to_word.get(run_style_id, run_style_id)
                _set_w_attr(rStyle, "val", word_id)

            self.next_rel_id = add_picture_to_document(
                run_data=run,
                parent_element=r,
                relationships=self.relationships,
                package_files=self.package_files,
                donor_media_path=self.donor_media_path,
                next_rel_id=self.next_rel_id
            )
            return r

        else:
            # unsupported run types (softHyphen, noBreakHyphen) – ignore
            return None

    def _remove_child_if_exists(self, parent: etree._Element, tag_local: str) -> None:
        child = parent.find(qn_w(tag_local))
        if child is not None:
            parent.remove(child)

    def _modify_paragraph(self, orig_p: etree._Element, json_p: Dict[str, Any]) -> None:
        """
        Модифицирует существующий абзац: обновляет pPr согласно json_p["p_format"]
        и заменяет все run'ы на новые из json_p["runs"].
        """
        # Обновляем pPr
        pPr = orig_p.find(qn_w("pPr"))
        if pPr is None:
            pPr = etree.SubElement(orig_p, qn_w("pPr"))
        p_format = json_p.get("p_format", {})
        self._update_pPr(pPr, p_format)

        # Удаляем все старые run'ы (кроме pPr)
        for child in list(orig_p):
            if child.tag != qn_w("pPr"):
                orig_p.remove(child)

        # Добавляем новые run'ы из JSON
        for run_data in json_p.get("runs", []):
            new_run = self._build_run(run_data)
            if new_run is not None:
                orig_p.append(new_run)

    def _update_pPr(self, pPr: etree._Element, p_format: Dict[str, Any]) -> None:
        """
        Обновляет элемент <w:pPr> на основе словаря p_format.
        Для каждого поддерживаемого свойства удаляет старый соответствующий элемент
        и добавляет новый, построенный из значения в p_format.
        """
        # Alignment
        if "alignment" in p_format:
            self._remove_child_if_exists(pPr, "jc")
            align_val = p_format["alignment"]
            # Преобразуем "justify" в "both" (как в _build_pPr)
            if align_val == "justify":
                xml_val = "both"
            else:
                xml_val = align_val
            jc = etree.SubElement(pPr, qn_w("jc"))
            _set_w_attr(jc, "val", xml_val)

        # Indents
        if any(k in p_format for k in ("indent_start_twip", "indent_end_twip",
                                        "indent_first_line_twip", "indent_hanging_twip")):
            self._remove_child_if_exists(pPr, "ind")
            ind = etree.SubElement(pPr, qn_w("ind"))
            if "indent_start_twip" in p_format:
                _set_w_attr_int(ind, "left", p_format["indent_start_twip"])
            if "indent_end_twip" in p_format:
                _set_w_attr_int(ind, "right", p_format["indent_end_twip"])
            if "indent_first_line_twip" in p_format:
                _set_w_attr_int(ind, "firstLine", p_format["indent_first_line_twip"])
            if "indent_hanging_twip" in p_format:
                _set_w_attr_int(ind, "hanging", p_format["indent_hanging_twip"])

        # Spacing
        if any(k in p_format for k in ("space_before_twip", "space_after_twip",
                                        "space_before_lines", "space_after_lines",
                                        "before_autospacing", "after_autospacing",
                                        "line_spacing_twip", "line_rule")):
            self._remove_child_if_exists(pPr, "spacing")
            sp = etree.SubElement(pPr, qn_w("spacing"))
            if "space_before_twip" in p_format:
                _set_w_attr_int(sp, "before", p_format["space_before_twip"])
            if "space_after_twip" in p_format:
                _set_w_attr_int(sp, "after", p_format["space_after_twip"])
            if "space_before_lines" in p_format:
                _set_w_attr_int(sp, "beforeLines", p_format["space_before_lines"])
            if "space_after_lines" in p_format:
                _set_w_attr_int(sp, "afterLines", p_format["space_after_lines"])
            if "before_autospacing" in p_format:
                _set_w_attr(sp, "beforeAutospacing", "1" if p_format["before_autospacing"] else "0")
            if "after_autospacing" in p_format:
                _set_w_attr(sp, "afterAutospacing", "1" if p_format["after_autospacing"] else "0")
            if "line_spacing_twip" in p_format:
                _set_w_attr_int(sp, "line", p_format["line_spacing_twip"])
            lr = p_format.get("line_rule")
            if lr:
                if lr == "auto":
                    _set_w_attr(sp, "lineRule", "auto")
                elif lr == "atLeast":
                    _set_w_attr(sp, "lineRule", "atLeast")
                elif lr == "exact":
                    _set_w_attr(sp, "lineRule", "exact")

        # Tabs (более сложно, можно пока пропустить или реализовать аналогично)
        if "tabs" in p_format:
            self._remove_child_if_exists(pPr, "tabs")
            tabs = p_format["tabs"]
            if isinstance(tabs, list) and tabs:
                tabs_el = etree.SubElement(pPr, qn_w("tabs"))
                for t in tabs:
                    if not isinstance(t, dict):
                        continue
                    pos = _safe_int(t.get("posTwip"))
                    val = t.get("val")
                    if pos is None or val is None:
                        continue
                    tab_el = etree.SubElement(tabs_el, qn_w("tab"))
                    _set_w_attr(tab_el, "pos", pos)
                    _set_w_attr(tab_el, "val", val)

    # =========================
    # BUILD: pPr from p_format diff
    # =========================

    def _build_pPr(self, p_format: Dict[str, Any]) -> Optional[etree._Element]:
        """
        Строит элемент <w:pPr> из словаря форматирования (camelCase).
        """
        if not p_format:
            return None

        pPr = _w_el("pPr")

        # Alignment
        align = p_format.get("alignment", "").lower()
        if align in ("left", "center", "right", "justify", "distribute"):
            jc = _w_sub(pPr, "jc")
            if align == "justify":
                _set_w_attr(jc, "val", "both")
            else:
                _set_w_attr(jc, "val", align)

        # Indents
        ind_keys = ("indentStartTwip", "indentEndTwip", "indentFirstLineTwip", "indentHangingTwip")
        if any(k in p_format for k in ind_keys):
            ind = _w_sub(pPr, "ind")
            if "indentStartTwip" in p_format:
                _set_w_attr_int(ind, "left", p_format["indentStartTwip"])
            if "indentEndTwip" in p_format:
                _set_w_attr_int(ind, "right", p_format["indentEndTwip"])
            if "indentFirstLineTwip" in p_format:
                _set_w_attr_int(ind, "firstLine", p_format["indentFirstLineTwip"])
            if "indentHangingTwip" in p_format:
                _set_w_attr_int(ind, "hanging", p_format["indentHangingTwip"])

        # Spacing
        if any(k in p_format for k in (
                "spaceBeforeTwip", "spaceAfterTwip", "spaceBeforeLines", "spaceAfterLines",
                "beforeAutospacing", "afterAutospacing", "lineTwip", "lineRule"
        )):
            sp = _w_sub(pPr, "spacing")
            if "spaceBeforeTwip" in p_format:
                _set_w_attr_int(sp, "before", p_format["spaceBeforeTwip"])
            if "spaceAfterTwip" in p_format:
                _set_w_attr_int(sp, "after", p_format["spaceAfterTwip"])
            if "spaceBeforeLines" in p_format:
                _set_w_attr_int(sp, "beforeLines", p_format["spaceBeforeLines"])
            if "spaceAfterLines" in p_format:
                _set_w_attr_int(sp, "afterLines", p_format["spaceAfterLines"])
            if "beforeAutospacing" in p_format:
                _set_w_attr(sp, "beforeAutospacing", "1" if p_format["beforeAutospacing"] else "0")
            if "afterAutospacing" in p_format:
                _set_w_attr(sp, "afterAutospacing", "1" if p_format["afterAutospacing"] else "0")
            if "lineTwip" in p_format:
                _set_w_attr_int(sp, "line", p_format["lineTwip"])
            lr = p_format.get("lineRule")
            if lr == "AUTO":
                _set_w_attr(sp, "lineRule", "auto")
            elif lr == "AT_LEAST":
                _set_w_attr(sp, "lineRule", "atLeast")
            elif lr == "EXACT":
                _set_w_attr(sp, "lineRule", "exact")

        # Tabs
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

        # Numbering
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

        # Boolean paragraph flags
        for k, tag in (
                ("keepNext", "keepNext"),
                ("keepLines", "keepLines"),
                ("pageBreakBefore", "pageBreakBefore"),
                ("widowControl", "widowControl"),
                ("snapToGrid", "snapToGrid"),
                ("contextualSpacing", "contextualSpacing"),
        ):
            if k in p_format:
                el = _w_sub(pPr, tag)
                if p_format[k] is False:
                    _set_w_attr(el, "val", "0")

        # Text alignment
        text_alignment = p_format.get("textAlignment")
        if text_alignment in ("AUTO", "BASELINE", "TOP", "CENTER", "BOTTOM"):
            ta = _w_sub(pPr, "textAlignment")
            _set_w_attr(ta, "val", text_alignment.lower())

        return pPr

    # =========================
    # BUILD: rPr from r_format
    # =========================

    def _build_rPr(self, r_format: Dict[str, Any]) -> Optional[etree._Element]:
        """
        Строит элемент <w:rPr> из словаря форматирования.
        """
        if not r_format:
            return None

        rPr = _w_el("rPr")

        # rFonts
        rFonts = r_format.get("rFonts")
        if isinstance(rFonts, dict) and rFonts:
            rf = _w_sub(rPr, "rFonts")
            for k, v in rFonts.items():
                if v is not None:
                    _set_w_attr(rf, k, v)

        # size
        if "font_size_half_points" in r_format:
            szv = _safe_int(r_format.get("font_size_half_points"))
            if szv is not None:
                sz = _w_sub(rPr, "sz")
                _set_w_attr(sz, "val", szv)
                szCs = _w_sub(rPr, "szCs")
                _set_w_attr(szCs, "val", szv)

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
        if isinstance(lang, str) and lang:
            le = _w_sub(rPr, "lang")
            _set_w_attr(le, "val", lang)
        elif isinstance(lang, dict) and lang:
            le = _w_sub(rPr, "lang")
            for k, v in lang.items():
                if v is not None:
                    _set_w_attr(le, k, v)

        # charSpacingTwip -> w:spacing
        if "charSpacingTwip" in r_format:
            spv = _safe_int(r_format.get("charSpacingTwip"))
            if spv is not None:
                sp = _w_sub(rPr, "spacing")
                _set_w_attr(sp, "val", spv)

        # positionHalfPoints -> w:position
        if "positionHalfPoints" in r_format:
            posv = _safe_int(r_format.get("positionHalfPoints"))
            if posv is not None:
                pos = _w_sub(rPr, "position")
                _set_w_attr(pos, "val", posv)

        return rPr

    # =========================
    # HELPERS
    # =========================

    def _dump_reconstructed_parts(self, package_files: Dict[str, bytes], out_root_dir: str) -> None:
        """Сохраняет все части пакета для отладки"""
        os.makedirs(out_root_dir, exist_ok=True)
        for name, data in sorted(package_files.items()):
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            out_path = os.path.join(out_root_dir, name)
            os.makedirs(os.path.dirname(out_path), exist_ok=True)
            with open(out_path, "wb") as f:
                f.write(data)

    def _serialize_xml(self, root: etree._Element, standalone: bool = True) -> bytes:
        return etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=standalone,
            pretty_print=False
        )


if __name__ == "__main__":
    try:
        cli = argparse.ArgumentParser(description="UltimateReconstructorV12 RAW JSON -> DOCX")
        cli.add_argument("--in-json", dest="input_json", required=True,
                         help="Path to input JSON file")
        cli.add_argument("--out-docx", dest="output_docx", required=True,
                         help="Path to output DOCX file")
        args = cli.parse_args()

        recon = UltimateReconstructorV12(args.input_json)
        recon.build_docx(args.output_docx)
        print("Реконструкция v12 завершена успешно!")
    except Exception:
        import traceback

        traceback.print_exc()
        sys.exit(1)