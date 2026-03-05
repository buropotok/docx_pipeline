"""
Тесты для проверки модификации свойств параграфа (p_format).
Запуск: python -m unittest tests/test_patch_pformat.py -v
Результаты сохраняются в docx_pipeline/test_results/
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
import random

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MY_NS = "https://translatefactory/schema/custom-id"
RED_COLOR = "FF0000"  # Красный цвет для выделения изменяемых параграфов


def qn_w(local):
    return f"{{{W_NS}}}{local}"


def qn_my(local):
    return f"{{{MY_NS}}}{local}"


class TestPatchPFormat(unittest.TestCase):
    """Тесты для проверки хирургического патча p_format."""

    @classmethod
    def setUpClass(cls):
        cls.original_json = "src/test_deep_seek/donor.json"
        cls.original_docx = "src/test_deep_seek/donor.materialized.docx"
        cls.results_dir = os.path.join("test_results", "pformat_tests")
        os.makedirs(cls.results_dir, exist_ok=True)

        assert os.path.exists(cls.original_json), f"Файл не найден: {cls.original_json}"
        assert os.path.exists(cls.original_docx), f"Файл не найден: {cls.original_docx}"

    def setUp(self):
        self.test_dir = tempfile.mkdtemp(prefix="pformat_test_")
        self.test_json = os.path.join(self.test_dir, "test.json")
        self.test_output = os.path.join(self.test_dir, "output.docx")

        with open(self.original_json, 'r', encoding='utf-8') as f:
            self.data = json.load(f)

    def tearDown(self):
        shutil.rmtree(self.test_dir)

    def _save_output_copy(self, test_name: str):
        """Сохраняет копию выходного файла с временной меткой."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{test_name}_{timestamp}.docx"
        output_path = os.path.join(self.results_dir, output_filename)
        shutil.copy2(self.test_output, output_path)
        print(f"\n📄 Результат сохранён: {output_path}")
        return output_path

    def _mark_paragraph_red(self, para_id: str):
        """Находит параграф по ID и делает весь его текст красным."""
        for item in self.data['content']:
            if item.get('type') == 'paragraph' and item.get('id') == para_id:
                for run in item.get('runs', []):
                    if run.get('type') == 'text':
                        if 'r_format' not in run:
                            run['r_format'] = {}
                        run['r_format']['color'] = RED_COLOR
                return True

            # Поиск в таблицах
            if item.get('type') == 'table':
                for row in item.get('rows', []):
                    for cell in row.get('cells', []):
                        for para in cell.get('content', []):
                            if para.get('id') == para_id:
                                for run in para.get('runs', []):
                                    if run.get('type') == 'text':
                                        if 'r_format' not in run:
                                            run['r_format'] = {}
                                        run['r_format']['color'] = RED_COLOR
                                return True
        return False

    def _run_reconstructor(self):
        """Запускает реконструктор с текущим JSON."""
        cmd = [
            sys.executable, "src/reconstructor.py",
            "--in-json", self.test_json,
            "--out-docx", self.test_output,
            "--donor-docx", self.original_docx
        ]

        result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
        self.assertEqual(result.returncode, 0, f"Ошибка реконструктора: {result.stderr}")
        return result

    def _verify_docx_opens(self, docx_path):
        """Проверяет, что DOCX файл открывается (валидный ZIP)."""
        try:
            with zipfile.ZipFile(docx_path, 'r') as zf:
                required = ["[Content_Types].xml", "word/document.xml"]
                for req in required:
                    self.assertIn(req, zf.namelist(), f"Missing {req}")
            return True
        except Exception as e:
            self.fail(f"DOCX file is corrupted: {e}")

    # =========================================================================
    # 1. БАЗОВЫЕ ТЕСТЫ (каждое свойство отдельно)
    # =========================================================================

    def test_alignment(self):
        """Тест 1.1: Изменение выравнивания на разных параграфах"""
        print("\n" + "=" * 70)
        print("ТЕСТ 1.1: Изменение выравнивания")
        print("=" * 70)

        # Берем параграф p_10 (после таблицы) и меняем выравнивание
        test_para_id = "p_10"
        self._mark_paragraph_red(test_para_id)

        # Находим параграф и применяем изменения
        for item in self.data['content']:
            if item.get('type') == 'paragraph' and item.get('id') == test_para_id:
                if 'p_format' not in item:
                    item['p_format'] = {}
                # Меняем выравнивание на justify (проверяем маппинг justify -> both)
                item['p_format']['alignment'] = "justify"
                break

        # Сохраняем JSON и запускаем реконструктор
        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_alignment")

        print(f"✅ Параграф {test_para_id} изменен: alignment='justify' (маппинг в 'both')")

    def test_text_alignment(self):
        """Тест 1.2: Изменение вертикального выравнивания текста"""
        print("\n" + "=" * 70)
        print("ТЕСТ 1.2: Вертикальное выравнивание текста")
        print("=" * 70)

        test_para_id = "p_15"
        self._mark_paragraph_red(test_para_id)

        for item in self.data['content']:
            if item.get('type') == 'paragraph' and item.get('id') == test_para_id:
                if 'p_format' not in item:
                    item['p_format'] = {}
                item['p_format']['text_alignment'] = "center"
                break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_text_alignment")

        print(f"✅ Параграф {test_para_id}: text_alignment='center'")

    def test_indents(self):
        """Тест 1.3: Все виды отступов"""
        print("\n" + "=" * 70)
        print("ТЕСТ 1.3: Отступы")
        print("=" * 70)

        test_para_id = "p_16"
        self._mark_paragraph_red(test_para_id)

        for item in self.data['content']:
            if item.get('type') == 'paragraph' and item.get('id') == test_para_id:
                if 'p_format' not in item:
                    item['p_format'] = {}
                item['p_format'].update({
                    "indent_start_twip": 1440,  # 2.5 см
                    "indent_end_twip": 720,  # 1.25 см
                    "indent_first_line_twip": 360,  # 0.625 см
                    "indent_hanging_twip": 540  # ~0.94 см
                })
                break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_indents")

        print(f"✅ Параграф {test_para_id}: все виды отступов")

    def test_spacing(self):
        """Тест 1.4: Интервалы"""
        print("\n" + "=" * 70)
        print("ТЕСТ 1.4: Интервалы")
        print("=" * 70)

        test_para_id = "p_17"
        self._mark_paragraph_red(test_para_id)

        for item in self.data['content']:
            if item.get('type') == 'paragraph' and item.get('id') == test_para_id:
                if 'p_format' not in item:
                    item['p_format'] = {}
                item['p_format'].update({
                    "space_before_twip": 480,  # интервал перед
                    "space_after_twip": 480,  # интервал после
                    "line_spacing_twip": 480,  # межстрочный
                    "line_rule": "exact"  # точный интервал
                })
                break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_spacing")

        print(f"✅ Параграф {test_para_id}: интервалы настроены")

    def test_line_rule(self):
        """Тест 1.5: Разные правила межстрочного интервала"""
        print("\n" + "=" * 70)
        print("ТЕСТ 1.5: Правила межстрочного интервала")
        print("=" * 70)

        test_para_id = "p_18"
        self._mark_paragraph_red(test_para_id)

        for item in self.data['content']:
            if item.get('type') == 'paragraph' and item.get('id') == test_para_id:
                if 'p_format' not in item:
                    item['p_format'] = {}
                # Проверяем все три возможных значения
                item['p_format'].update({
                    "line_spacing_twip": 480,
                    "line_rule": "atLeast"  # auto, exact, atLeast
                })
                break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_line_rule")

        print(f"✅ Параграф {test_para_id}: line_rule='atLeast'")

    def test_boolean_flags(self):
        """Тест 1.6: Булевы флаги"""
        print("\n" + "=" * 70)
        print("ТЕСТ 1.6: Булевы флаги")
        print("=" * 70)

        test_para_id = "p_19"
        self._mark_paragraph_red(test_para_id)

        for item in self.data['content']:
            if item.get('type') == 'paragraph' and item.get('id') == test_para_id:
                if 'p_format' not in item:
                    item['p_format'] = {}
                # Устанавливаем все булевы флаги в True
                item['p_format'].update({
                    "keep_next": True,
                    "keep_lines": True,
                    "page_break_before": True,
                    "widow_control": True,
                    "contextual_spacing": True,
                    "snap_to_grid": True
                })
                break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_boolean_flags")

        print(f"✅ Параграф {test_para_id}: все булевы флаги установлены в True")

    def test_tabs(self):
        """Тест 1.7: Табуляция"""
        print("\n" + "=" * 70)
        print("ТЕСТ 1.7: Табуляция")
        print("=" * 70)

        test_para_id = "p_20"
        self._mark_paragraph_red(test_para_id)

        # Добавляем текст с табуляцией в параграф, если его нет
        for item in self.data['content']:
            if item.get('type') == 'paragraph' and item.get('id') == test_para_id:
                # Убедимся, что есть runs с текстом
                if not item.get('runs'):
                    item['runs'] = [{
                        "type": "text",
                        "text": "Колонка 1\tКолонка 2\tКолонка 3",
                        "id": f"{test_para_id}.run_1",
                        "parent_id": test_para_id,
                        "meta": {"preserve": True}
                    }]

                if 'p_format' not in item:
                    item['p_format'] = {}

                item['p_format']['tabs'] = [
                    {"posTwip": 1440, "val": "left"},  # 2.5 см, по левому краю
                    {"posTwip": 2880, "val": "center"},  # 5 см, по центру
                    {"posTwip": 4320, "val": "right"},  # 7.5 см, по правому краю
                    {"posTwip": 5760, "val": "decimal"},  # 10 см, десятичная
                    {"posTwip": 7200, "val": "bar"}  # 12.5 см, вертикальная черта
                ]
                break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_tabs")

        print(f"✅ Параграф {test_para_id}: добавлены табуляции")

    def test_list_info(self):
        """Тест 1.8: Нумерация"""
        print("\n" + "=" * 70)
        print("ТЕСТ 1.8: Нумерация")
        print("=" * 70)

        test_para_id = "p_21"
        self._mark_paragraph_red(test_para_id)

        for item in self.data['content']:
            if item.get('type') == 'paragraph' and item.get('id') == test_para_id:
                if 'p_format' not in item:
                    item['p_format'] = {}

                # Привязываем к существующей нумерации (numId="1", уровень 0)
                item['p_format']['list_info'] = {
                    "numId": "1",
                    "ilvl": "0"
                }
                break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_list_info")

        print(f"✅ Параграф {test_para_id}: добавлена нумерация")

    # =========================================================================
    # 2. КОМБИНИРОВАННЫЕ ТЕСТЫ
    # =========================================================================

    def test_combined_simple(self):
        """Тест 2.1: Комбинация 2-3 свойств"""
        print("\n" + "=" * 70)
        print("ТЕСТ 2.1: Комбинация 2-3 свойств")
        print("=" * 70)

        test_para_id = "p_22"
        self._mark_paragraph_red(test_para_id)

        for item in self.data['content']:
            if item.get('type') == 'paragraph' and item.get('id') == test_para_id:
                if 'p_format' not in item:
                    item['p_format'] = {}

                item['p_format'].update({
                    "alignment": "center",
                    "indent_start_twip": 720,
                    "space_before_twip": 240,
                    "keep_next": True
                })
                break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_combined_simple")

        print(f"✅ Параграф {test_para_id}: применены center + отступ 720 + интервал 240")

    def test_combined_all(self):
        """Тест 2.2: Все свойства сразу"""
        print("\n" + "=" * 70)
        print("ТЕСТ 2.2: Все свойства сразу")
        print("=" * 70)

        test_para_id = "p_23"
        self._mark_paragraph_red(test_para_id)

        for item in self.data['content']:
            if item.get('type') == 'paragraph' and item.get('id') == test_para_id:
                if 'p_format' not in item:
                    item['p_format'] = {}

                item['p_format'].update({
                    "alignment": "justify",
                    "text_alignment": "center",
                    "contextual_spacing": True,
                    "line_spacing_twip": 480,
                    "line_rule": "atLeast",
                    "space_before_twip": 240,
                    "space_after_twip": 240,
                    "indent_start_twip": 720,
                    "indent_end_twip": 360,
                    "indent_first_line_twip": 180,
                    "keep_next": True,
                    "keep_lines": True,
                    "page_break_before": False,
                    "widow_control": True,
                    "snap_to_grid": True,
                    "tabs": [
                        {"posTwip": 1440, "val": "left"},
                        {"posTwip": 2880, "val": "center"}
                    ],
                    "list_info": {
                        "numId": "1",
                        "ilvl": "1"
                    }
                })
                break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_combined_all")

        print(f"✅ Параграф {test_para_id}: применены все свойства")

    # =========================================================================
    # 3. ТЕСТЫ НА ПАРАГРАФАХ ВНУТРИ ТАБЛИЦ
    # =========================================================================

    def test_table_cell_alignment(self):
        """Тест 3.1: Изменение выравнивания в ячейке таблицы"""
        print("\n" + "=" * 70)
        print("ТЕСТ 3.1: Выравнивание в ячейке таблицы")
        print("=" * 70)

        test_para_id = "tbl_1.row_2.cell_2.p_1"
        self._mark_paragraph_red(test_para_id)

        # Находим параграф в таблице
        for item in self.data['content']:
            if item.get('type') == 'table' and item.get('id') == 'tbl_1':
                for row in item.get('rows', []):
                    for cell in row.get('cells', []):
                        for para in cell.get('content', []):
                            if para.get('id') == test_para_id:
                                if 'p_format' not in para:
                                    para['p_format'] = {}
                                para['p_format']['alignment'] = "right"
                                break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_table_cell_alignment")

        print(f"✅ Параграф {test_para_id} в таблице: alignment='right'")

    def test_table_cell_indents(self):
        """Тест 3.2: Отступы в ячейке таблицы"""
        print("\n" + "=" * 70)
        print("ТЕСТ 3.2: Отступы в ячейке таблицы")
        print("=" * 70)

        test_para_id = "tbl_1.row_3.cell_1.p_1"
        self._mark_paragraph_red(test_para_id)

        for item in self.data['content']:
            if item.get('type') == 'table' and item.get('id') == 'tbl_1':
                for row in item.get('rows', []):
                    for cell in row.get('cells', []):
                        for para in cell.get('content', []):
                            if para.get('id') == test_para_id:
                                if 'p_format' not in para:
                                    para['p_format'] = {}
                                para['p_format'].update({
                                    "indent_start_twip": 360,
                                    "indent_end_twip": 360,
                                    "indent_first_line_twip": 180
                                })
                                break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_table_cell_indents")

        print(f"✅ Параграф {test_para_id} в таблице: отступы")

    def test_table_cell_spacing(self):
        """Тест 3.3: Интервалы в ячейке таблицы"""
        print("\n" + "=" * 70)
        print("ТЕСТ 3.3: Интервалы в ячейке таблицы")
        print("=" * 70)

        test_para_id = "tbl_1.row_4.cell_2.p_2"
        self._mark_paragraph_red(test_para_id)

        for item in self.data['content']:
            if item.get('type') == 'table' and item.get('id') == 'tbl_1':
                for row in item.get('rows', []):
                    for cell in row.get('cells', []):
                        for para in cell.get('content', []):
                            if para.get('id') == test_para_id:
                                if 'p_format' not in para:
                                    para['p_format'] = {}
                                para['p_format'].update({
                                    "space_before_twip": 240,
                                    "space_after_twip": 240,
                                    "line_spacing_twip": 360
                                })
                                break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_table_cell_spacing")

        print(f"✅ Параграф {test_para_id} в таблице: интервалы")

    def test_table_cell_boolean(self):
        """Тест 3.4: Булевы флаги в ячейке таблицы"""
        print("\n" + "=" * 70)
        print("ТЕСТ 3.4: Булевы флаги в ячейке таблицы")
        print("=" * 70)

        test_para_id = "tbl_2.row_1.cell_1.p_1"
        self._mark_paragraph_red(test_para_id)

        for item in self.data['content']:
            if item.get('type') == 'table' and item.get('id') == 'tbl_2':
                for row in item.get('rows', []):
                    for cell in row.get('cells', []):
                        for para in cell.get('content', []):
                            if para.get('id') == test_para_id:
                                if 'p_format' not in para:
                                    para['p_format'] = {}
                                para['p_format'].update({
                                    "keep_next": True,
                                    "keep_lines": True,
                                    "widow_control": False
                                })
                                break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_table_cell_boolean")

        print(f"✅ Параграф {test_para_id} в таблице: булевы флаги")

    def test_table_cell_tabs(self):
        """Тест 3.5: Табуляция в ячейке таблицы"""
        print("\n" + "=" * 70)
        print("ТЕСТ 3.5: Табуляция в ячейке таблицы")
        print("=" * 70)

        test_para_id = "tbl_3.row_2.cell_3.p_1"
        self._mark_paragraph_red(test_para_id)

        # Добавляем текст с табуляцией
        for item in self.data['content']:
            if item.get('type') == 'table' and item.get('id') == 'tbl_3':
                for row in item.get('rows', []):
                    for cell in row.get('cells', []):
                        for para in cell.get('content', []):
                            if para.get('id') == test_para_id:
                                # Добавляем текст с табуляцией
                                if not para.get('runs'):
                                    para['runs'] = [{
                                        "type": "text",
                                        "text": "Данные 1\tДанные 2\tДанные 3",
                                        "id": f"{test_para_id}.run_1",
                                        "parent_id": test_para_id
                                    }]

                                if 'p_format' not in para:
                                    para['p_format'] = {}

                                para['p_format']['tabs'] = [
                                    {"posTwip": 1000, "val": "left"},
                                    {"posTwip": 2000, "val": "center"},
                                    {"posTwip": 3000, "val": "right"}
                                ]
                                break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_table_cell_tabs")

        print(f"✅ Параграф {test_para_id} в таблице: табуляция")

    def test_table_cell_list(self):
        """Тест 3.6: Нумерация в ячейке таблицы"""
        print("\n" + "=" * 70)
        print("ТЕСТ 3.6: Нумерация в ячейке таблицы")
        print("=" * 70)

        test_para_id = "tbl_3.row_3.cell_2.p_1"
        self._mark_paragraph_red(test_para_id)

        for item in self.data['content']:
            if item.get('type') == 'table' and item.get('id') == 'tbl_3':
                for row in item.get('rows', []):
                    for cell in row.get('cells', []):
                        for para in cell.get('content', []):
                            if para.get('id') == test_para_id:
                                if 'p_format' not in para:
                                    para['p_format'] = {}

                                para['p_format']['list_info'] = {
                                    "numId": "1",
                                    "ilvl": "0"
                                }
                                break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_table_cell_list")

        print(f"✅ Параграф {test_para_id} в таблице: нумерация")

    # =========================================================================
    # 4. КРАЕВЫЕ СЛУЧАИ
    # =========================================================================

    def test_empty_paragraph(self):
        """Тест 4.1: Пустой параграф"""
        print("\n" + "=" * 70)
        print("ТЕСТ 4.1: Пустой параграф (p_2)")
        print("=" * 70)

        test_para_id = "p_2"  # пустой параграф в доноре
        self._mark_paragraph_red(test_para_id)

        for item in self.data['content']:
            if item.get('type') == 'paragraph' and item.get('id') == test_para_id:
                if 'p_format' not in item:
                    item['p_format'] = {}
                item['p_format']['alignment'] = "center"
                break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_empty_paragraph")

        print(f"✅ Пустой параграф {test_para_id}: alignment='center'")

    def test_paragraph_without_jc(self):
        """Тест 4.2: Параграф без элемента jc"""
        print("\n" + "=" * 70)
        print("ТЕСТ 4.2: Параграф без jc")
        print("=" * 70)

        # Ищем параграф без jc или создаем его
        test_para_id = "p_24"  # предположительно без jc

        # Удаляем jc из p_format, если есть
        for item in self.data['content']:
            if item.get('type') == 'paragraph' and item.get('id') == test_para_id:
                if 'p_format' in item:
                    item['p_format'].pop('alignment', None)

                # Теперь добавляем jc через патч
                if 'p_format' not in item:
                    item['p_format'] = {}
                item['p_format']['alignment'] = "right"
                break

        self._mark_paragraph_red(test_para_id)

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_paragraph_without_jc")

        print(f"✅ Параграф {test_para_id}: добавлен jc")

    def test_negative_indents(self):
        """Тест 4.3: Отрицательные отступы"""
        print("\n" + "=" * 70)
        print("ТЕСТ 4.3: Отрицательные отступы")
        print("=" * 70)

        test_para_id = "p_25"
        self._mark_paragraph_red(test_para_id)

        for item in self.data['content']:
            if item.get('type') == 'paragraph' and item.get('id') == test_para_id:
                if 'p_format' not in item:
                    item['p_format'] = {}

                # Отрицательные отступы (выступы)
                item['p_format'].update({
                    "indent_start_twip": -360,  # выступ влево
                    "indent_end_twip": -360,  # выступ вправо
                    "indent_hanging_twip": 180  # висячий отступ
                })
                break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_negative_indents")

        print(f"✅ Параграф {test_para_id}: отрицательные отступы")

    def test_zero_spacing(self):
        """Тест 4.4: Нулевые интервалы"""
        print("\n" + "=" * 70)
        print("ТЕСТ 4.4: Нулевые интервалы")
        print("=" * 70)

        test_para_id = "p_26"
        self._mark_paragraph_red(test_para_id)

        for item in self.data['content']:
            if item.get('type') == 'paragraph' and item.get('id') == test_para_id:
                if 'p_format' not in item:
                    item['p_format'] = {}

                item['p_format'].update({
                    "space_before_twip": 0,
                    "space_after_twip": 0,
                    "line_spacing_twip": 240
                })
                break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_zero_spacing")

        print(f"✅ Параграф {test_para_id}: нулевые интервалы")

    def test_tabs_clear(self):
        """Тест 4.5: Очистка существующих табуляций"""
        print("\n" + "=" * 70)
        print("ТЕСТ 4.5: Очистка и установка новых табуляций")
        print("=" * 70)

        test_para_id = "p_33"  # уже есть табуляции в доноре
        self._mark_paragraph_red(test_para_id)

        for item in self.data['content']:
            if item.get('type') == 'paragraph' and item.get('id') == test_para_id:
                if 'p_format' not in item:
                    item['p_format'] = {}

                # Заменяем существующие табуляции новыми
                item['p_format']['tabs'] = [
                    {"posTwip": 1000, "val": "left", "leader": "dot"},
                    {"posTwip": 3000, "val": "center", "leader": "hyphen"},
                    {"posTwip": 5000, "val": "right", "leader": "underscore"}
                ]
                break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_tabs_clear")

        print(f"✅ Параграф {test_para_id}: табуляции заменены")

    # =========================================================================
    # 5. ТЕСТЫ НА НЕСКОЛЬКИХ ПАРАГРАФАХ
    # =========================================================================

    def test_multiple_paragraphs_diff(self):
        """Тест 5.1: Разные параграфы с разными изменениями"""
        print("\n" + "=" * 70)
        print("ТЕСТ 5.1: Несколько параграфов с разными изменениями")
        print("=" * 70)

        test_paras = [
            {"id": "p_27", "changes": {"alignment": "center", "indent_start_twip": 360}},
            {"id": "p_28", "changes": {"space_before_twip": 480, "keep_next": True}},
            {"id": "p_29", "changes": {"tabs": [{"posTwip": 2000, "val": "left"}]}},
            {"id": "p_30", "changes": {"list_info": {"numId": "1", "ilvl": "2"}}}
        ]

        for para_info in test_paras:
            self._mark_paragraph_red(para_info["id"])

            for item in self.data['content']:
                if item.get('type') == 'paragraph' and item.get('id') == para_info["id"]:
                    if 'p_format' not in item:
                        item['p_format'] = {}
                    item['p_format'].update(para_info["changes"])
                    break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_multiple_paragraphs_diff")

        print(f"✅ {len(test_paras)} параграфов изменены")

    def test_multiple_paragraphs_same(self):
        """Тест 5.2: Несколько параграфов с одинаковыми изменениями"""
        print("\n" + "=" * 70)
        print("ТЕСТ 5.2: Несколько параграфов с одинаковыми изменениями")
        print("=" * 70)

        para_ids = ["p_34", "p_35", "p_36"]
        common_changes = {
            "alignment": "justify",
            "indent_start_twip": 720,
            "space_before_twip": 240,
            "contextual_spacing": True
        }

        for para_id in para_ids:
            self._mark_paragraph_red(para_id)

            for item in self.data['content']:
                if item.get('type') == 'paragraph' and item.get('id') == para_id:
                    if 'p_format' not in item:
                        item['p_format'] = {}
                    item['p_format'].update(common_changes)
                    break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        self._run_reconstructor()
        self._verify_docx_opens(self.test_output)
        self._save_output_copy("test_multiple_paragraphs_same")

        print(f"✅ {len(para_ids)} параграфов с одинаковыми изменениями")


if __name__ == '__main__':
    unittest.main(verbosity=2)