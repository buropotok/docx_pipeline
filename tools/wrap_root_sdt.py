#!/usr/bin/env python3
"""
wrap_root_sdt.py
Обёртывает все корневые абзацы и таблицы в document.xml в SDT-блоки с уникальными идентификаторами.
Также оборачивает каждую строку таблицы в отдельный SDT с тегом table_id.row_<n>.
Создаёт новый DOCX с размеченным document.xml, сохраняя все остальные части.
Использование: python wrap_root_sdt.py --in input.docx --out output.docx
"""

import argparse
import os
import zipfile
import shutil
import tempfile
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

def qn_w(tag):
    return f"{{{W_NS}}}{tag}"

def qn_r(tag):
    return f"{{{R_NS}}}{tag}"

def create_sdt(tag_id, content_elem):
    """Создаёт элемент w:sdt с заданным тегом и содержимым."""
    sdt = etree.Element(qn_w("sdt"))
    # sdtPr
    sdtPr = etree.SubElement(sdt, qn_w("sdtPr"))
    tag_el = etree.SubElement(sdtPr, qn_w("tag"))
    tag_el.set(qn_w("val"), tag_id)
    # Скрываем визуальное отображение SDT (рамку)
    appearance = etree.SubElement(sdtPr, qn_w("appearance"))
    appearance.set(qn_w("val"), "none")
    # Также скрываем текст-заполнитель (на всякий случай)
    etree.SubElement(sdtPr, qn_w("showingPlcHdr")).set(qn_w("val"), "false")
    # sdtContent
    sdtContent = etree.SubElement(sdt, qn_w("sdtContent"))
    sdtContent.append(content_elem)
    return sdt

def wrap_table_rows(tbl, table_id):
    """
    Оборачивает каждую строку таблицы в SDT с тегом вида table_id.row_<n>.
    Возвращает новую таблицу с обёрнутыми строками (оригинальная таблица не изменяется).
    """
    # Создаём копию таблицы, чтобы не портить оригинал
    new_tbl = etree.Element(tbl.tag, tbl.attrib)
    # Копируем все дочерние элементы, кроме строк (их будем обрабатывать отдельно)
    for child in tbl:
        if child.tag != qn_w("tr"):
            new_tbl.append(etree.fromstring(etree.tostring(child)))
    # Оборачиваем строки
    row_counter = 1
    for tr in tbl.findall(qn_w("tr")):
        row_id = f"{table_id}.row_{row_counter}"
        # Копируем строку (глубокое копирование)
        tr_copy = etree.fromstring(etree.tostring(tr))
        sdt_row = create_sdt(row_id, tr_copy)
        new_tbl.append(sdt_row)
        row_counter += 1
    return new_tbl

def main():
    parser = argparse.ArgumentParser(description="Wrap root paragraphs and tables in SDT blocks with IDs")
    parser.add_argument("--in", dest="input_docx", required=True, help="Input DOCX file")
    parser.add_argument("--out", dest="output_docx", required=True, help="Output DOCX file")
    args = parser.parse_args()

    temp_dir = tempfile.mkdtemp()
    try:
        # Распаковываем исходный docx
        with zipfile.ZipFile(args.input_docx, 'r') as zin:
            zin.extractall(temp_dir)

        document_path = os.path.join(temp_dir, "word", "document.xml")
        if not os.path.exists(document_path):
            raise FileNotFoundError(f"word/document.xml not found in {args.input_docx}")

        # Парсим document.xml
        tree = etree.parse(document_path)
        root = tree.getroot()
        body = root.find(qn_w("body"))
        if body is None:
            raise ValueError("No body element")

        # Счётчики для идентификаторов
        p_counter = 1
        tbl_counter = 1

        # Список новых детей body
        new_children = []

        for child in body:
            tag = child.tag
            if tag == qn_w("p"):
                # Параграф
                sdt = create_sdt(f"p_{p_counter}", child)
                new_children.append(sdt)
                p_counter += 1
            elif tag == qn_w("tbl"):
                # Таблица – сначала оборачиваем её строки
                wrapped_tbl = wrap_table_rows(child, f"tbl_{tbl_counter}")
                sdt = create_sdt(f"tbl_{tbl_counter}", wrapped_tbl)
                new_children.append(sdt)
                tbl_counter += 1
            else:
                # Не-абзацный/не-табличный элемент – оставляем как есть
                new_children.append(child)

        # Заменяем содержимое body
        for child in list(body):
            body.remove(child)
        for elem in new_children:
            body.append(elem)

        # Сохраняем document.xml
        tree.write(document_path, xml_declaration=True, encoding="UTF-8", standalone=True)

        # Собираем новый docx
        with zipfile.ZipFile(args.output_docx, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(file_path, temp_dir).replace('\\', '/')
                    zout.write(file_path, arcname)

        print(f"Successfully wrapped root elements and table rows with SDT tags. Output: {args.output_docx}")

    finally:
        shutil.rmtree(temp_dir)

if __name__ == "__main__":
    main()