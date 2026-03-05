"""
Тесты для проверки модификации свойств параграфа (p_format).
Запуск: python -m unittest tests/test_p_format.py -v
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
import sys
import importlib
importlib.reload(sys)  # перезагружаем модуль sys
# sys.setdefaultencoding не существует в Python 3, но можно:
import os
os.environ["PYTHONIOENCODING"] = "utf-8"

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MY_NS = "https://translatefactory/schema/custom-id"


def qn_w(local):
    return f"{{{W_NS}}}{local}"


def qn_my(local):
    return f"{{{MY_NS}}}{local}"


class TestPFormat(unittest.TestCase):
    """Тесты для проверки модификации свойств параграфа."""

    @classmethod
    def setUpClass(cls):
        cls.original_json = "src/test_deep_seek/donor.json"
        cls.original_docx = "src/test_deep_seek/donor.materialized.docx"
        cls.results_dir = "test_results"
        os.makedirs(cls.results_dir, exist_ok=True)

    def setUp(self):
        self.test_dir = tempfile.mkdtemp(prefix="pformat_test_")
        self.test_json = os.path.join(self.test_dir, "test.json")
        self.test_output = os.path.join(self.test_dir, "output.docx")

        with open(self.original_json, 'r', encoding='utf-8') as f:
            self.data = json.load(f)

    def tearDown(self):
        shutil.rmtree(self.test_dir)

    def _save_output_copy(self, test_name: str):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{test_name}_{timestamp}.docx"
        output_path = os.path.join(self.results_dir, output_filename)
        shutil.copy2(self.test_output, output_path)
        print(f"\n📄 Результат сохранён: {output_path}")
        return output_path

    def _get_paragraph_property(self, docx_path, para_id, property_name):
        """Извлекает значение свойства параграфа из XML."""
        with tempfile.TemporaryDirectory() as tmp:
            with zipfile.ZipFile(docx_path, 'r') as zf:
                zf.extractall(tmp)

            doc_path = os.path.join(tmp, "word", "document.xml")
            tree = etree.parse(doc_path)
            root = tree.getroot()

            xpath = f".//w:p[@my:id='{para_id}']"
            namespaces = {'w': W_NS, 'my': MY_NS}
            paras = root.xpath(xpath, namespaces=namespaces)

            if not paras:
                return None

            p = paras[0]
            pPr = p.find(qn_w("pPr"))
            if pPr is None:
                return None

            # Для разных типов свойств разная логика извлечения
            if property_name == "alignment":
                jc = pPr.find(qn_w("jc"))
                return jc.get(qn_w("val")) if jc is not None else None

            elif property_name.startswith("indent_"):
                ind = pPr.find(qn_w("ind"))
                if ind is None:
                    return None
                attr_map = {
                    "indent_start_twip": "left",
                    "indent_end_twip": "right",
                    "indent_first_line_twip": "firstLine",
                    "indent_hanging_twip": "hanging"
                }
                return ind.get(qn_w(attr_map[property_name]))

            elif property_name in ["space_before_twip", "space_after_twip",
                                  "line_spacing_twip", "line_rule"]:
                spacing = pPr.find(qn_w("spacing"))
                if spacing is None:
                    return None
                attr_map = {
                    "space_before_twip": "before",
                    "space_after_twip": "after",
                    "line_spacing_twip": "line",
                    "line_rule": "lineRule"
                }
                return spacing.get(qn_w(attr_map[property_name]))

            elif property_name in ["keep_next", "keep_lines", "page_break_before",
                                  "widow_control", "contextual_spacing", "snap_to_grid"]:
                tag_map = {
                    "keep_next": "keepNext",
                    "keep_lines": "keepLines",
                    "page_break_before": "pageBreakBefore",
                    "widow_control": "widowControl",
                    "contextual_spacing": "contextualSpacing",
                    "snap_to_grid": "snapToGrid"
                }
                el = pPr.find(qn_w(tag_map[property_name]))
                return el is not None

            return None

    # ------------------------------------------------------------------------
    # Тест 1: Выравнивание
    # ------------------------------------------------------------------------

    def test_alignment(self):
        """Тест 1: Изменение выравнивания параграфа."""
        print("\n" + "=" * 70)
        print("ТЕСТ 1: Изменение выравнивания")
        print("=" * 70)

        # Берём первый параграф
        paragraphs = [item for item in self.data['content']
                     if item['type'] == 'paragraph']
        target = paragraphs[0]
        para_id = target['id']

        print(f"📌 Параграф: {para_id}")

        # Меняем выравнивание на center и делаем текст красным
        target['p_format'] = {
            "alignment": "center"
        }

        # Делаем весь текст красным, чтобы было видно изменения
        for run in target.get('runs', []):
            if run['type'] == 'text':
                run.setdefault('r_format', {})['color'] = 'FF0000'

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
        self.assertEqual(result.returncode, 0, f"Ошибка: {result.stderr}")

        saved_path = self._save_output_copy("alignment_center")

        # Проверяем выравнивание
        alignment = self._get_paragraph_property(self.test_output, para_id, "alignment")
        self.assertEqual(alignment, "center", "Выравнивание должно быть center")

        print(f"\n✅ Выравнивание изменено на center")
        print(f"   🔴 Текст выделен красным")
        print(f"   Результат: {saved_path}")

    # ------------------------------------------------------------------------
    # Тест 2: Отступы
    # ------------------------------------------------------------------------

    def test_indents(self):
        """Тест 2: Изменение отступов параграфа."""
        print("\n" + "=" * 70)
        print("ТЕСТ 2: Изменение отступов")
        print("=" * 70)

        paragraphs = [item for item in self.data['content']
                     if item['type'] == 'paragraph']
        target = paragraphs[0]
        para_id = target['id']

        print(f"📌 Параграф: {para_id}")

        # Устанавливаем все виды отступов
        target['p_format'] = {
            "indent_start_twip": 720,      # 0.5 см
            "indent_end_twip": 720,
            "indent_first_line_twip": 360,  # 0.25 см
            "indent_hanging_twip": 360
        }

        # Красный текст для визуализации
        for run in target.get('runs', []):
            if run['type'] == 'text':
                run.setdefault('r_format', {})['color'] = 'FF0000'

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

        saved_path = self._save_output_copy("indents")

        # Проверяем отступы
        start = self._get_paragraph_property(self.test_output, para_id, "indent_start_twip")
        end = self._get_paragraph_property(self.test_output, para_id, "indent_end_twip")
        first = self._get_paragraph_property(self.test_output, para_id, "indent_first_line_twip")
        hanging = self._get_paragraph_property(self.test_output, para_id, "indent_hanging_twip")

        self.assertEqual(start, "720", "Отступ слева должен быть 720")
        self.assertEqual(end, "720", "Отступ справа должен быть 720")
        self.assertEqual(first, "360", "Отступ первой строки должен быть 360")
        self.assertEqual(hanging, "360", "Висячий отступ должен быть 360")

        print(f"\n✅ Отступы изменены")
        print(f"   🔴 Текст выделен красным")
        print(f"   Результат: {saved_path}")

    # ------------------------------------------------------------------------
    # Тест 3: Интервалы
    # ------------------------------------------------------------------------

    def test_spacing(self):
        """Тест 3: Изменение интервалов."""
        print("\n" + "=" * 70)
        print("ТЕСТ 3: Изменение интервалов")
        print("=" * 70)

        paragraphs = [item for item in self.data['content']
                     if item['type'] == 'paragraph']
        target = paragraphs[0]
        para_id = target['id']

        print(f"📌 Параграф: {para_id}")

        target['p_format'] = {
            "space_before_twip": 240,      # интервал перед
            "space_after_twip": 240,       # интервал после
            "line_spacing_twip": 360,       # межстрочный интервал
            "line_rule": "exact"            # точный интервал
        }

        for run in target.get('runs', []):
            if run['type'] == 'text':
                run.setdefault('r_format', {})['color'] = 'FF0000'

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

        saved_path = self._save_output_copy("spacing")

        # Проверяем интервалы
        before = self._get_paragraph_property(self.test_output, para_id, "space_before_twip")
        after = self._get_paragraph_property(self.test_output, para_id, "space_after_twip")
        line = self._get_paragraph_property(self.test_output, para_id, "line_spacing_twip")
        rule = self._get_paragraph_property(self.test_output, para_id, "line_rule")

        self.assertEqual(before, "240", "Интервал перед должен быть 240")
        self.assertEqual(after, "240", "Интервал после должен быть 240")
        self.assertEqual(line, "360", "Межстрочный интервал должен быть 360")
        self.assertEqual(rule, "exact", "Правило должно быть exact")

        print(f"\n✅ Интервалы изменены")
        print(f"   🔴 Текст выделен красным")
        print(f"   Результат: {saved_path}")

    # ------------------------------------------------------------------------
    # Тест 4: Булевы флаги
    # ------------------------------------------------------------------------

    def test_boolean_flags(self):
        """Тест 4: Установка булевых флагов."""
        print("\n" + "=" * 70)
        print("ТЕСТ 4: Булевы флаги")
        print("=" * 70)

        paragraphs = [item for item in self.data['content']
                     if item['type'] == 'paragraph']
        target = paragraphs[0]
        para_id = target['id']

        print(f"📌 Параграф: {para_id}")

        target['p_format'] = {
            "keep_next": True,
            "keep_lines": True,
            "page_break_before": True,
            "widow_control": True,
            "contextual_spacing": True,
            "snap_to_grid": True
        }

        for run in target.get('runs', []):
            if run['type'] == 'text':
                run.setdefault('r_format', {})['color'] = 'FF0000'

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

        saved_path = self._save_output_copy("boolean_flags")

        # Проверяем флаги
        keep_next = self._get_paragraph_property(self.test_output, para_id, "keep_next")
        keep_lines = self._get_paragraph_property(self.test_output, para_id, "keep_lines")
        page_break = self._get_paragraph_property(self.test_output, para_id, "page_break_before")
        widow = self._get_paragraph_property(self.test_output, para_id, "widow_control")
        contextual = self._get_paragraph_property(self.test_output, para_id, "contextual_spacing")
        snap = self._get_paragraph_property(self.test_output, para_id, "snap_to_grid")

        self.assertTrue(keep_next, "keep_next должен быть True")
        self.assertTrue(keep_lines, "keep_lines должен быть True")
        self.assertTrue(page_break, "page_break_before должен быть True")
        self.assertTrue(widow, "widow_control должен быть True")
        self.assertTrue(contextual, "contextual_spacing должен быть True")
        self.assertTrue(snap, "snap_to_grid должен быть True")

        print(f"\n✅ Булевы флаги установлены")
        print(f"   🔴 Текст выделен красным")
        print(f"   Результат: {saved_path}")

    # ------------------------------------------------------------------------
    # Тест 5: Табуляция
    # ------------------------------------------------------------------------

    def test_tabs(self):
        """Тест 5: Добавление табуляции."""
        print("\n" + "=" * 70)
        print("ТЕСТ 5: Табуляция")
        print("=" * 70)

        paragraphs = [item for item in self.data['content']
                     if item['type'] == 'paragraph']
        target = paragraphs[0]
        para_id = target['id']

        print(f"📌 Параграф: {para_id}")

        target['p_format'] = {
            "tabs": [
                {"posTwip": 1440, "val": "left"},      # 2.5 см
                {"posTwip": 2880, "val": "center"},    # 5 см
                {"posTwip": 4320, "val": "right"},     # 7.5 см
                {"posTwip": 5760, "val": "decimal"}    # 10 см
            ]
        }

        for run in target.get('runs', []):
            if run['type'] == 'text':
                run.setdefault('r_format', {})['color'] = 'FF0000'

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

        saved_path = self._save_output_copy("tabs")

        print(f"\n✅ Табуляция добавлена")
        print(f"   🔴 Текст выделен красным")
        print(f"   Результат: {saved_path}")

    # ------------------------------------------------------------------------
    # Тест 6: Нумерация
    # ------------------------------------------------------------------------

    def test_list_info(self):
        """Тест 6: Привязка к нумерации."""
        print("\n" + "=" * 70)
        print("ТЕСТ 6: Нумерация")
        print("=" * 70)

        paragraphs = [item for item in self.data['content']
                     if item['type'] == 'paragraph']
        target = paragraphs[0]
        para_id = target['id']

        print(f"📌 Параграф: {para_id}")

        # Добавляем нумерацию (numId должен существовать в документе)
        target['p_format'] = {
            "list_info": {
                "numId": "1",
                "ilvl": "0"
            }
        }

        for run in target.get('runs', []):
            if run['type'] == 'text':
                run.setdefault('r_format', {})['color'] = 'FF0000'

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

        saved_path = self._save_output_copy("list_info")

        print(f"\n✅ Нумерация добавлена")
        print(f"   🔴 Текст выделен красным")
        print(f"   Результат: {saved_path}")

    # ------------------------------------------------------------------------
    # Тест 7: Комбинированный
    # ------------------------------------------------------------------------

    def test_combined(self):
        """Тест 7: Все параметры вместе."""
        print("\n" + "=" * 70)
        print("ТЕСТ 7: Все параметры вместе")
        print("=" * 70)

        paragraphs = [item for item in self.data['content']
                     if item['type'] == 'paragraph']
        target = paragraphs[0]
        para_id = target['id']

        print(f"📌 Параграф: {para_id}")

        target['p_format'] = {
            "alignment": "justify",
            "text_alignment": "center",
            "contextual_spacing": True,
            "line_spacing_twip": 480,
            "line_rule": "atLeast",
            "space_before_twip": 240,
            "space_after_twip": 240,
            "indent_start_twip": 720,
            "indent_end_twip": 720,
            "indent_first_line_twip": 360,
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
                "ilvl": "0"
            }
        }

        for run in target.get('runs', []):
            if run['type'] == 'text':
                run.setdefault('r_format', {})['color'] = 'FF0000'

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

        saved_path = self._save_output_copy("combined")

        print(f"\n✅ Все параметры применены")
        print(f"   🔴 Текст выделен красным")
        print(f"   Результат: {saved_path}")


if __name__ == '__main__':
    unittest.main(verbosity=2)