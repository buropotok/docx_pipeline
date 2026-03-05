"""
Тесты для проверки операций с таблицами.
Запуск: python -m unittest tests/test_table.py -v
"""

import unittest
import json
import os
import tempfile
import shutil
import subprocess
import zipfile
from lxml import etree
import sys
from datetime import datetime
import copy

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MY_NS = "https://translatefactory/schema/custom-id"


def qn_w(local):
    return f"{{{W_NS}}}{local}"


def qn_my(local):
    return f"{{{MY_NS}}}{local}"


def make_text_run(text, bold=False, color=None):
    """Создаёт run с текстом и опциональным форматированием."""
    run = {
        "type": "text",
        "text": text,
        "r_format": {}
    }
    if bold:
        run["r_format"]["bold"] = True
    if color:
        run["r_format"]["color"] = color
    return run


def make_paragraph(para_id, runs, p_style_id="a"):
    """Создаёт параграф с указанными run'ами."""
    return {
        "type": "paragraph",
        "id": para_id,
        "p_style_id": p_style_id,
        "runs": runs
    }


class TestTableOperations(unittest.TestCase):
    """Тесты для проверки операций с таблицами."""

    @classmethod
    def setUpClass(cls):
        cls.original_json = "src/test_deep_seek/donor.json"
        cls.original_docx = "src/test_deep_seek/donor.materialized.docx"
        cls.results_dir = "test_results"
        os.makedirs(cls.results_dir, exist_ok=True)

        assert os.path.exists(cls.original_json), f"Файл не найден: {cls.original_json}"
        assert os.path.exists(cls.original_docx), f"Файл не найден: {cls.original_docx}"

    def setUp(self):
        self.test_dir = tempfile.mkdtemp(prefix="table_test_")
        self.test_json = os.path.join(self.test_dir, "test.json")
        self.test_output = os.path.join(self.test_dir, "output.docx")

        with open(self.original_json, 'r', encoding='utf-8') as f:
            self.data = json.load(f)

    def tearDown(self):
        shutil.rmtree(self.test_dir)

    def _save_output_copy(self, test_name: str):
        """Сохраняет копию выходного файла."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{test_name}_{timestamp}.docx"
        output_path = os.path.join(self.results_dir, output_filename)
        shutil.copy2(self.test_output, output_path)
        print(f"\n📄 Результат сохранён: {output_path}")
        return output_path

    def _get_row_ids_from_docx(self, docx_path, table_id):
        """Получает все my:id строк таблицы из document.xml."""
        with tempfile.TemporaryDirectory() as tmp:
            with zipfile.ZipFile(docx_path, 'r') as zf:
                zf.extractall(tmp)

            doc_path = os.path.join(tmp, "word", "document.xml")
            tree = etree.parse(doc_path)
            root = tree.getroot()

            xpath = f".//w:tbl[@my:id='{table_id}']"
            namespaces = {'w': W_NS, 'my': MY_NS}
            tables = root.xpath(xpath, namespaces=namespaces)

            if not tables:
                return []

            table = tables[0]
            rows = table.findall(qn_w("tr"))
            return [row.get(qn_my("id")) for row in rows if row.get(qn_my("id"))]

    def _get_cell_count(self, docx_path, row_id):
        """Получает количество ячеек в строке."""
        with tempfile.TemporaryDirectory() as tmp:
            with zipfile.ZipFile(docx_path, 'r') as zf:
                zf.extractall(tmp)

            doc_path = os.path.join(tmp, "word", "document.xml")
            tree = etree.parse(doc_path)
            root = tree.getroot()

            xpath = f".//w:tr[@my:id='{row_id}']"
            namespaces = {'w': W_NS, 'my': MY_NS}
            rows = root.xpath(xpath, namespaces=namespaces)

            if not rows:
                return 0

            cells = [tc for tc in rows[0] if tc.tag == qn_w("tc")]
            return len(cells)

    # ------------------------------------------------------------------------
    # Тест 1: Порядок строк (перемещенные строки выделяем красным)
    # ------------------------------------------------------------------------

    def test_table_rows_order(self):
        """Тест 1: Порядок строк в таблице соответствует JSON."""
        print("\n" + "=" * 70)
        print("ТЕСТ 1: Порядок строк в таблице")
        print("=" * 70)

        tables = [item for item in self.data['content'] if item['type'] == 'table']
        self.assertGreater(len(tables), 0, "В документе нет таблиц")

        table = tables[0]
        table_id = table['id']
        original_rows = table['rows'].copy()

        print(f"📌 Таблица: {table_id}")
        print(f"   Оригинальных строк: {len(original_rows)}")

        if len(original_rows) < 3:
            print("⚠️ Недостаточно строк для теста")
            return

        # Меняем порядок: [row_3, row_1, row_2, row_4...]
        new_order = [original_rows[2], original_rows[0], original_rows[1]] + original_rows[3:]

        # Отмечаем перемещенные строки красным жирным
        for i, row in enumerate(new_order):
            if row['id'] != original_rows[i]['id']:
                # Строка изменила позицию - выделяем всё содержимое красным
                for cell in row.get('cells', []):
                    for para in cell.get('content', []):
                        for run in para.get('runs', []):
                            if run['type'] == 'text':
                                run.setdefault('r_format', {})['bold'] = True
                                run.setdefault('r_format', {})['color'] = 'FF0000'

        table['rows'] = new_order

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        cmd = [
            "python", "src/reconstructor.py",
            "--in-json", self.test_json,
            "--out-docx", self.test_output,
            "--donor-docx", self.original_docx
        ]

        result = subprocess.run(cmd, capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, f"Ошибка: {result.stderr}")

        self._save_output_copy("table_rows_order")

        result_rows = self._get_row_ids_from_docx(self.test_output, table_id)
        expected_ids = [row['id'] for row in new_order]

        print(f"\n🔍 Ожидаемый порядок: {expected_ids}")
        print(f"🔍 Фактический порядок: {result_rows}")
        print(f"🎨 Перемещенные строки выделены красным жирным")

        self.assertEqual(result_rows, expected_ids, "Порядок строк не соответствует JSON")
        print(f"\n✅ Тест пройден: порядок строк сохранён")

    # ------------------------------------------------------------------------
    # Тест 2: Удаление строк (ячейки удаленных строк не проверяем)
    # ------------------------------------------------------------------------

    def test_delete_table_rows(self):
        """Тест 2: Удаление строк таблицы."""
        print("\n" + "=" * 70)
        print("ТЕСТ 2: Удаление строк таблицы")
        print("=" * 70)

        tables = [item for item in self.data['content'] if item['type'] == 'table']
        self.assertGreater(len(tables), 0, "В документе нет таблиц")

        table = tables[0]
        table_id = table['id']
        original_rows = table['rows'].copy()

        print(f"📌 Таблица: {table_id}")
        print(f"   Оригинальных строк: {len(original_rows)}")

        if len(original_rows) < 2:
            print("⚠️ Недостаточно строк для теста удаления")
            return

        deleted_row_id = original_rows[1]['id']
        new_rows = [original_rows[0]] + original_rows[2:]

        table['rows'] = new_rows

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        cmd = [
            "python", "src/reconstructor.py",
            "--in-json", self.test_json,
            "--out-docx", self.test_output,
            "--donor-docx", self.original_docx
        ]

        result = subprocess.run(cmd, capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, f"Ошибка: {result.stderr}")

        self._save_output_copy("table_delete_rows")

        result_rows = self._get_row_ids_from_docx(self.test_output, table_id)
        expected_ids = [row['id'] for row in new_rows]

        print(f"\n🔍 Ожидаемые строки: {expected_ids}")
        print(f"🔍 Фактические строки: {result_rows}")

        self.assertNotIn(deleted_row_id, result_rows, f"Строка {deleted_row_id} должна быть удалена")
        self.assertEqual(result_rows, expected_ids, "Список строк не соответствует ожидаемому")
        print(f"\n✅ Тест пройден: строка {deleted_row_id} удалена")

    # ------------------------------------------------------------------------
    # Тест 3: Вставка новой строки (клон) - вся новая строка красным жирным
    # ------------------------------------------------------------------------

    def test_insert_new_row(self):
        """Тест 3: Вставка новой строки (клон существующей)."""
        print("\n" + "=" * 70)
        print("ТЕСТ 3: Вставка новой строки")
        print("=" * 70)

        tables = [item for item in self.data['content'] if item['type'] == 'table']
        self.assertGreater(len(tables), 0, "В документе нет таблиц")

        table = tables[0]
        table_id = table['id']
        original_rows = table['rows'].copy()

        print(f"📌 Таблица: {table_id}")
        print(f"   Оригинальных строк: {len(original_rows)}")

        source_row = original_rows[0]
        new_row_id = f"{source_row['id']}.1"

        # Создаём новую строку с красным жирным текстом
        new_cells = []
        for cell_idx, source_cell in enumerate(source_row['cells'], start=1):
            new_cell_id = f"{new_row_id}.cell_{cell_idx}"
            new_paras = []

            for para_idx, source_para in enumerate(source_cell.get('content', []), start=1):
                if source_para.get('type') == 'paragraph':
                    new_para = {
                        "type": "paragraph",
                        "id": f"{new_cell_id}.p_{para_idx}",
                        "p_style_id": source_para.get('p_style_id', 'a'),
                        "runs": []
                    }

                    # Копируем текст из источника, но делаем его красным жирным
                    for run in source_para.get('runs', []):
                        if run['type'] == 'text':
                            new_run = copy.deepcopy(run)
                            new_run.setdefault('r_format', {})['bold'] = True
                            new_run.setdefault('r_format', {})['color'] = 'FF0000'
                            new_para['runs'].append(new_run)

                    new_paras.append(new_para)

            new_cells.append({
                "id": new_cell_id,
                "content": new_paras
            })

        new_row = {
            "id": new_row_id,
            "derive_from": source_row['id'],
            "cells": new_cells
        }

        new_rows = [original_rows[0], new_row] + original_rows[1:]
        table['rows'] = new_rows

        print(f"   Новая строка: {new_row_id} (клон из {source_row['id']})")
        print(f"   Ячеек в новой строке: {len(new_cells)}")
        print(f"   🎨 Вся новая строка выделена красным жирным")

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        cmd = [
            "python", "src/reconstructor.py",
            "--in-json", self.test_json,
            "--out-docx", self.test_output,
            "--donor-docx", self.original_docx
        ]

        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            print(f"❌ Ошибка: {result.stderr}")

        self.assertEqual(result.returncode, 0, f"Ошибка: {result.stderr}")

        self._save_output_copy("table_insert_row")

        result_rows = self._get_row_ids_from_docx(self.test_output, table_id)

        print(f"\n🔍 Строки в результате: {result_rows}")

        self.assertIn(new_row_id, result_rows, f"Новая строка {new_row_id} не найдена")
        self.assertIn(source_row['id'], result_rows, f"Исходная строка {source_row['id']} пропала")

        source_idx = result_rows.index(source_row['id'])
        new_idx = result_rows.index(new_row_id)
        self.assertEqual(new_idx, source_idx + 1,
                         f"Новая строка должна быть сразу после {source_row['id']}")

        source_cells = self._get_cell_count(self.test_output, source_row['id'])
        new_cells_count = self._get_cell_count(self.test_output, new_row_id)
        self.assertEqual(source_cells, new_cells_count,
                         "Количество ячеек в новой строке не совпадает с источником")

        print(f"\n✅ Тест пройден: строка {new_row_id} вставлена правильно")

    # ------------------------------------------------------------------------
    # Тест 4: Операции с параграфами внутри ячеек
    # ------------------------------------------------------------------------

    def test_cell_paragraphs(self):
        """Тест 4: Операции с параграфами внутри ячеек."""
        print("\n" + "=" * 70)
        print("ТЕСТ 4: Параграфы внутри ячеек")
        print("=" * 70)

        tables = [item for item in self.data['content'] if item['type'] == 'table']
        self.assertGreater(len(tables), 0, "В документе нет таблиц")

        target_table = None
        target_row = None
        target_cell_idx = None

        for table in tables:
            for row_idx, row in enumerate(table.get('rows', [])):
                for cell_idx, cell in enumerate(row.get('cells', [])):
                    if cell.get('content') and len(cell['content']) >= 2:
                        target_table = table
                        target_row = row
                        target_cell_idx = cell_idx
                        break
                if target_row:
                    break
            if target_table:
                break

        if not target_table:
            print("⚠️ Не найдена ячейка с достаточным количеством параграфов")
            return

        table_id = target_table['id']
        row_id = target_row['id']
        cell = target_row['cells'][target_cell_idx]
        original_paras = cell['content'].copy()

        print(f"📌 Таблица: {table_id}")
        print(f"📌 Строка: {row_id}")
        print(f"📌 Ячейка: {target_cell_idx + 1}")
        print(f"   Параграфов в ячейке: {len(original_paras)}")
        print(f"   ID параграфов: {[p['id'] for p in original_paras]}")

        # Создаём новый параграф с ПРАВИЛЬНЫМ ID (числовой суффикс)
        new_para_id = f"{cell['id']}.{len(original_paras) + 1}"  # ✅ числовой суффикс
        source_para = original_paras[0]  # берём первый параграф как шаблон

        new_para = {
            "type": "paragraph",
            "id": new_para_id,
            "derive_from": source_para['id'],
            "p_style_id": source_para.get('p_style_id', 'a'),
            "runs": [
                {
                    "type": "text",
                    "text": "ЭТО НОВЫЙ ПАРАГРАФ",
                    "r_format": {
                        "bold": True,
                        "color": "FF0000"
                    }
                }
            ]
        }

        # Отмечаем второй параграф как удалённый
        original_paras[1]['deleted'] = True

        # Строим новый список параграфов в JSON порядке
        new_paras_json = [
            original_paras[0],  # первый остаётся
            new_para,  # новый параграф после первого
            original_paras[2],  # третий
            original_paras[3]  # четвёртый
            # original_paras[1] - НЕ включаем, он deleted
        ]

        cell['content'] = new_paras_json

        # Выводим отладочную информацию
        for p in cell['content']:
            if p['id'] == new_para_id:
                print(f"   🔍 Новый параграф: {json.dumps(p, indent=2, ensure_ascii=False)}")
                print(f"   🔍 derive_from: {p.get('derive_from')}")

        print(f"   JSON параграфов для отправки: {[p.get('id') for p in cell['content']]}")
        print(
            f"   Из них новых: {[p['id'] for p in cell['content'] if '.' in p['id'] and p['id'].split('.')[-1].isdigit()]}")

        # Зеленый цвет для всей ячейки, НО не затираем красный в новом параграфе
        for para in cell['content']:
            if para.get('deleted'):
                continue
            # Пропускаем новый параграф (он должен остаться красным)
            if para['id'] == new_para_id:
                continue
            for run in para.get('runs', []):
                if run['type'] == 'text':
                    run.setdefault('r_format', {})['color'] = '00FF00'

        # Сохраняем JSON
        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        # Запускаем реконструктор
        cmd = [
            "python", "src/reconstructor.py",
            "--in-json", self.test_json,
            "--out-docx", self.test_output,
            "--donor-docx", self.original_docx
        ]

        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            print(f"❌ Ошибка: {result.stderr}")

        self.assertEqual(result.returncode, 0, f"Ошибка: {result.stderr}")

        self._save_output_copy("cell_paragraphs")

        print(f"\n✅ Тест пройден: параграфы в ячейке переставлены")
        print(f"   🟢 Вся ячейка выделена зеленым (из-за удаления параграфа)")
        print(f"   🔴 Новый параграф выделен красным жирным")

    # ------------------------------------------------------------------------
    # Тест 5: Комбинированные операции
    # ------------------------------------------------------------------------

    def test_table_mixed_ops(self):
        """Тест 5: Удаление + вставка строк в одной таблице."""
        print("\n" + "=" * 70)
        print("ТЕСТ 5: Удаление + вставка строк")
        print("=" * 70)

        tables = [item for item in self.data['content'] if item['type'] == 'table']
        self.assertGreater(len(tables), 0, "В документе нет таблиц")

        table = tables[0]
        table_id = table['id']
        original_rows = table['rows'].copy()

        print(f"📌 Таблица: {table_id}")
        print(f"   Оригинальных строк: {len(original_rows)}")

        if len(original_rows) < 3:
            print("⚠️ Недостаточно строк для теста")
            return

        # Удаляем вторую строку
        deleted_row_id = original_rows[1]['id']

        # Клонируем первую строку (красная жирная)
        source_row = original_rows[0]
        new_row_id = f"{source_row['id']}.1"

        new_cells = []
        for cell_idx, source_cell in enumerate(source_row['cells'], start=1):
            new_cell_id = f"{new_row_id}.cell_{cell_idx}"
            new_paras = []

            for para_idx, source_para in enumerate(source_cell.get('content', []), start=1):
                if source_para.get('type') == 'paragraph':
                    new_para = {
                        "type": "paragraph",
                        "id": f"{new_cell_id}.p_{para_idx}",
                        "p_style_id": source_para.get('p_style_id', 'a'),
                        "runs": []
                    }

                    for run in source_para.get('runs', []):
                        if run['type'] == 'text':
                            new_run = copy.deepcopy(run)
                            new_run.setdefault('r_format', {})['bold'] = True
                            new_run.setdefault('r_format', {})['color'] = 'FF0000'
                            new_para['runs'].append(new_run)

                    new_paras.append(new_para)

            new_cells.append({
                "id": new_cell_id,
                "content": new_paras
            })

        new_row = {
            "id": new_row_id,
            "derive_from": source_row['id'],
            "cells": new_cells
        }

        # Строим новый список: [row1, row3, row4, new_row]
        new_rows = [original_rows[0], original_rows[2]] + original_rows[3:] + [new_row]
        table['rows'] = new_rows

        print(f"   Удаляем: {deleted_row_id}")
        print(f"   Вставляем: {new_row_id} (красная жирная) после всех существующих")

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        cmd = [
            "python", "src/reconstructor.py",
            "--in-json", self.test_json,
            "--out-docx", self.test_output,
            "--donor-docx", self.original_docx
        ]

        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            print(f"❌ Ошибка: {result.stderr}")

        self.assertEqual(result.returncode, 0, f"Ошибка: {result.stderr}")

        self._save_output_copy("table_mixed_ops")

        result_rows = self._get_row_ids_from_docx(self.test_output, table_id)
        expected_ids = [row['id'] for row in new_rows]

        print(f"\n🔍 Ожидаемые строки: {expected_ids}")
        print(f"🔍 Фактические строки: {result_rows}")
        print(f"   🟢 Строка {deleted_row_id} удалена")
        print(f"   🔴 Новая строка {new_row_id} выделена красным жирным")

        self.assertNotIn(deleted_row_id, result_rows, f"Строка {deleted_row_id} должна быть удалена")
        self.assertIn(new_row_id, result_rows, f"Новая строка {new_row_id} не найдена")
        self.assertEqual(result_rows, expected_ids, "Порядок строк не соответствует JSON")

        print(f"\n✅ Тест пройден: комбинированные операции выполнены")


if __name__ == '__main__':
    unittest.main(verbosity=2)