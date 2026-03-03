#!/usr/bin/env python3
"""
add_custom_attrs.py
Добавляет пользовательские атрибуты my:id к корневым абзацам (w:p),
таблицам (w:tbl) и строкам таблиц (w:tr) в word/document.xml.

Namespace:
  my -> https://translatefactory/schema/custom-id
И добавляет 'my' в mc:Ignorable.

Пишет новый DOCX, изменяя только word/document.xml внутри ZIP.
"""

import argparse
import zipfile
from lxml import etree

W_NS  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
MY_NS = "https://translatefactory/schema/custom-id"

MY_PREFIX = "my"
MC_PREFIX = "mc"

def qn(ns, local):
    return f"{{{ns}}}{local}"

def ensure_root_ns_and_ignorable(tree: etree._ElementTree) -> etree._ElementTree:
    """
    Гарантирует:
      - xmlns:my
      - xmlns:mc (если вдруг отсутствует, но нужен для mc:Ignorable)
      - mc:Ignorable содержит 'my'
    Делается через пересборку корневого элемента с корректным nsmap.
    """
    old_root = tree.getroot()

    # Собираем новый nsmap
    nsmap = dict(old_root.nsmap) if old_root.nsmap else {}
    if MY_PREFIX not in nsmap:
        nsmap[MY_PREFIX] = MY_NS
    if MC_PREFIX not in nsmap:
        nsmap[MC_PREFIX] = MC_NS

    # Новый корневой элемент с расширенным nsmap
    new_root = etree.Element(old_root.tag, nsmap=nsmap)

    # Переносим атрибуты
    for k, v in old_root.attrib.items():
        new_root.set(k, v)

    # Переносим детей
    for child in list(old_root):
        new_root.append(child)

    # Переносим текст/хвост (на всякий)
    new_root.text = old_root.text
    new_root.tail = old_root.tail

    # Обновляем mc:Ignorable
    ign_attr = qn(MC_NS, "Ignorable")
    current = new_root.get(ign_attr, "")
    tokens = current.split() if current else []
    if MY_PREFIX not in tokens:
        tokens.append(MY_PREFIX)
        new_root.set(ign_attr, " ".join(tokens))

    return etree.ElementTree(new_root)

def add_ids(tree: etree._ElementTree) -> None:
    root = tree.getroot()
    body = root.find(qn(W_NS, "body"))
    if body is None:
        raise ValueError("No w:body in document.xml")

    p_counter = 1
    tbl_counter = 1

    # Только корневые элементы body (как у тебя задумано)
    for child in list(body):
        if child.tag == qn(W_NS, "p"):
            child.set(qn(MY_NS, "id"), f"p_{p_counter}")
            p_counter += 1

        elif child.tag == qn(W_NS, "tbl"):
            tbl_id = f"tbl_{tbl_counter}"
            child.set(qn(MY_NS, "id"), tbl_id)

            row_counter = 1
            # w:tr — прямые дети w:tbl
            ns = {"w": W_NS}
            for tr in child.xpath("./w:tr", namespaces=ns):
                tr.set(qn(MY_NS, "id"), f"{tbl_id}.row_{row_counter}")
                row_counter += 1

            tbl_counter += 1

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="input_docx", required=True)
    ap.add_argument("--out", dest="output_docx", required=True)
    args = ap.parse_args()

    with zipfile.ZipFile(args.input_docx, "r") as zin:
        # читаем document.xml
        try:
            doc_xml = zin.read("word/document.xml")
        except KeyError:
            raise FileNotFoundError("word/document.xml not found in DOCX")

        # парсим без "красоты" — Word не любит лишние форматирования
        parser = etree.XMLParser(remove_blank_text=False, recover=False, huge_tree=True)
        tree = etree.fromstring(doc_xml, parser=parser)
        tree = etree.ElementTree(tree)

        # гарантируем ns + mc:Ignorable
        tree = ensure_root_ns_and_ignorable(tree)

        # добавляем id
        add_ids(tree)

        # сериализуем без standalone
        new_doc_xml = etree.tostring(
            tree,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=None,       # важно: не добавлять standalone="yes"
            pretty_print=False
        )

        # Пишем новый docx "бережно": сохраняем ZipInfo и метод сжатия каждого entry,
        # меняем только word/document.xml
        with zipfile.ZipFile(args.output_docx, "w") as zout:
            for src_info in zin.infolist():
                data = zin.read(src_info.filename)
                if src_info.filename == "word/document.xml":
                    data = new_doc_xml

                # Создаём новый ZipInfo и копируем метаданные
                dst_info = zipfile.ZipInfo(filename=src_info.filename, date_time=src_info.date_time)
                dst_info.compress_type = src_info.compress_type
                dst_info.comment = src_info.comment
                dst_info.extra = src_info.extra
                dst_info.create_system = src_info.create_system
                dst_info.create_version = src_info.create_version
                dst_info.extract_version = src_info.extract_version
                dst_info.flag_bits = src_info.flag_bits
                dst_info.internal_attr = src_info.internal_attr
                dst_info.external_attr = src_info.external_attr

                zout.writestr(dst_info, data)

    print(f"Success. Output: {args.output_docx}")

if __name__ == "__main__":
    main()