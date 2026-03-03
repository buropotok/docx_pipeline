# reconstructor.py

import argparse
import zipfile
from pathlib import Path
from collections import OrderedDict
from lxml import etree


NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_MY = "https://translatefactory/schema/custom-id"

def qn(tag):
    prefix, local = tag.split(":")
    if prefix == "w":
        return f"{{{NS_W}}}{local}"
    raise ValueError(tag)


def index_document(docx_path: Path):
    with zipfile.ZipFile(docx_path, "r") as z:
        xml_bytes = z.read("word/document.xml")

    root = etree.fromstring(xml_bytes)
    body = root.find(qn("w:body"))

    # ===============================
    # Индексационные структуры
    # ===============================

    root_elements_by_id = OrderedDict()
    row_elements_by_id = OrderedDict()
    cell_elements_by_id = OrderedDict()
    paragraphs_by_id = OrderedDict()
    runs_by_id = OrderedDict()
    passthrough = []

    orig_children = list(body)  # исходный порядок детей body
    elements_to_delete = set()  # Множество для элементов, которые нужно удалить

    for el in body:
        my_id = el.get(f"{{{NS_MY}}}id")

        # ============================================
        # КОРНЕВЫЕ АБЗАЦЫ
        # ============================================
        if el.tag == qn("w:p"):
            if my_id:
                root_elements_by_id[my_id] = el
                paragraphs_by_id[my_id] = el

                # Проверка на флаг 'deleted'
                if 'deleted' in el.attrib and el.attrib['deleted'] == 'true':
                    elements_to_delete.add(my_id)
                    continue  # Пропускаем этот элемент, он будет удалён
                # Индексация run'ов корневого абзаца
                for run_index, r in enumerate(el.findall(qn("w:r")), start=1):
                    run_id = f"{my_id}.run_{run_index}"
                    runs_by_id[run_id] = r

        # ============================================
        # ТАБЛИЦЫ
        # ============================================
        elif el.tag == qn("w:tbl"):
            if my_id:
                root_elements_by_id[my_id] = el

            # ------------------------
            # Строки таблицы
            # ------------------------
            for tr in el.findall(qn("w:tr")):
                row_id = tr.get(f"{{{NS_MY}}}id")
                if row_id:
                    row_elements_by_id[row_id] = tr

                # ------------------------
                # Ячейки
                # ------------------------
                for cell_index, tc in enumerate(tr.findall(qn("w:tc")), start=1):

                    if not row_id:
                        raise RuntimeError("Table row without my:id encountered")

                    cell_id = f"{row_id}.cell_{cell_index}"
                    cell_elements_by_id[cell_id] = tc

                    # ------------------------
                    # Вложенные абзацы
                    # ------------------------
                    for p_index, p in enumerate(tc.findall(qn("w:p")), start=1):
                        para_id = f"{cell_id}.p_{p_index}"
                        paragraphs_by_id[para_id] = p

                        # ------------------------
                        # Run'ы вложенного абзаца
                        # ------------------------
                        for run_index, r in enumerate(p.findall(qn("w:r")), start=1):
                            run_id = f"{para_id}.run_{run_index}"
                            runs_by_id[run_id] = r

        # ============================================
        # PASSTHROUGH
        # ============================================

        else:
            passthrough.append(el)

    # ============================================
    # Валидация уникальности ID
    # ============================================

    if len(root_elements_by_id) != len(set(root_elements_by_id.keys())):
        raise RuntimeError("Duplicate root element IDs detected")

    if len(row_elements_by_id) != len(set(row_elements_by_id.keys())):
        raise RuntimeError("Duplicate row IDs detected")

    if len(paragraphs_by_id) != len(set(paragraphs_by_id.keys())):
        raise RuntimeError("Duplicate paragraph IDs detected")

    if len(runs_by_id) != len(set(runs_by_id.keys())):
        raise RuntimeError("Duplicate run IDs detected")

    print(f"[INDEX] Root elements: {len(root_elements_by_id)}")
    print(f"[INDEX] Rows: {len(row_elements_by_id)}")
    print(f"[INDEX] Cells: {len(cell_elements_by_id)}")
    print(f"[INDEX] Paragraphs: {len(paragraphs_by_id)}")
    print(f"[INDEX] Runs: {len(runs_by_id)}")

    return (
        root,
        body,
        orig_children,
        root_elements_by_id,
        row_elements_by_id,
        cell_elements_by_id,
        paragraphs_by_id,
        runs_by_id,
        passthrough,
    )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--in-json", required=True)
    parser.add_argument("--out-docx", required=True)
    args = parser.parse_args()

    json_path = Path(args.in_json)
    run_dir = json_path.parent
    donor_docx = run_dir / "donor.materialized.docx"

    if not donor_docx.exists():
        raise FileNotFoundError(f"Donor not found: {donor_docx}")

    print(f"[INFO] Indexing donor: {donor_docx}")

    index_document(donor_docx)

    print("[OK] Indexing complete (no modifications yet)")

    # Пока просто копируем donor как reconstructed
    output_path = Path(args.out_docx)
    output_path.write_bytes(donor_docx.read_bytes())

    print("[OK] Reconstructed (AS IS copy)")


if __name__ == "__main__":
    main()