"""
Тесты для проверки операции удаления корневых элементов.
Запуск: python -m pytest tests/test_deletion.py -v
или: python -m unittest tests/test_deletion.py -v
"""

import unittest
import json
import random
import os
import tempfile
import shutil
import subprocess
import zipfile
from lxml import etree
import sys
from datetime import datetime

# Добавляем путь к src для импорта, если нужно
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

# Пространства имён для XML
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MY_NS = "https://translatefactory/schema/custom-id"


def qn_w(local):
    return f"{{{W_NS}}}{local}"


def qn_my(local):
    return f"{{{MY_NS}}}{local}"


class TestRootDeletion(unittest.TestCase):
    """Тесты для проверки удаления корневых элементов."""

    @classmethod
    def setUpClass(cls):
        """Выполняется один раз перед всеми тестами."""
        # Пути к исходным файлам (относительно корня проекта)
        cls.original_json = "src/test_deep_seek/donor.json"
        cls.original_docx = "src/test_deep_seek/donor.materialized.docx"

        # Создаём папку для результатов тестов
        cls.results_dir = "test_results"
        os.makedirs(cls.results_dir, exist_ok=True)

        # Проверяем, что файлы существуют
        assert os.path.exists(cls.original_json), f"Файл не найден: {cls.original_json}"
        assert os.path.exists(cls.original_docx), f"Файл не найден: {cls.original_docx}"

    def setUp(self):
        """Выполняется перед каждым тестом."""
        # Создаём временную директорию для теста
        self.test_dir = tempfile.mkdtemp(prefix="reconstructor_test_")
        self.test_json = os.path.join(self.test_dir, "test.json")
        self.test_output = os.path.join(self.test_dir, "output.docx")

        # Загружаем оригинальный JSON
        with open(self.original_json, 'r', encoding='utf-8') as f:
            self.data = json.load(f)

    def tearDown(self):
        """Выполняется после каждого теста."""
        # Удаляем временную директорию
        shutil.rmtree(self.test_dir)

    def _save_output_copy(self, test_name: str):
        """Сохраняет копию выходного файла для наглядной проверки."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{test_name}_{timestamp}.docx"
        output_path = os.path.join(self.results_dir, output_filename)
        shutil.copy2(self.test_output, output_path)
        print(f"\n📄 Результат сохранён: {output_path}")
        return output_path

    def _print_xml_preview(self, docx_path: str, max_lines: int = 20):
        """Выводит превью XML для наглядной проверки."""
        with tempfile.TemporaryDirectory() as tmp:
            with zipfile.ZipFile(docx_path, 'r') as zf:
                zf.extractall(tmp)

            doc_path = os.path.join(tmp, "word", "document.xml")
            with open(doc_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()

            #print("\n📄 Превью document.xml (первые {} строк):".format(max_lines))
            # print("-" * 60)
            # for i, line in enumerate(lines[:max_lines]):
            #     print(f"{i+1:3d}: {line.rstrip()}")
            # if len(lines) > max_lines:
            #     print(f"... и ещё {len(lines) - max_lines} строк")
            # print("-" * 60)

    def _count_root_elements_in_docx(self, docx_path):
        """Вспомогательный метод: подсчитывает корневые элементы в document.xml."""
        with tempfile.TemporaryDirectory() as tmp:
            with zipfile.ZipFile(docx_path, 'r') as zf:
                zf.extractall(tmp)

            doc_path = os.path.join(tmp, "word", "document.xml")
            tree = etree.parse(doc_path)
            root = tree.getroot()
            body = root.find(qn_w("body"))

            paragraphs = [elem for elem in body if elem.tag == qn_w("p")]
            tables = [elem for elem in body if elem.tag == qn_w("tbl")]

            return len(paragraphs), len(tables)

    def _get_element_ids_from_docx(self, docx_path):
        """Вспомогательный метод: получает все my:id из document.xml."""
        with tempfile.TemporaryDirectory() as tmp:
            with zipfile.ZipFile(docx_path, 'r') as zf:
                zf.extractall(tmp)

            doc_path = os.path.join(tmp, "word", "document.xml")
            tree = etree.parse(doc_path)
            root = tree.getroot()
            body = root.find(qn_w("body"))

            ids = []
            for elem in body:
                if elem.tag in (qn_w("p"), qn_w("tbl")):
                    elem_id = elem.get(qn_my("id"))
                    if elem_id:
                        ids.append(elem_id)
            return ids

    def _get_element_text_preview(self, docx_path, elem_id: str) -> str:
        """Получает превью текста элемента для наглядности."""
        with tempfile.TemporaryDirectory() as tmp:
            with zipfile.ZipFile(docx_path, 'r') as zf:
                zf.extractall(tmp)

            doc_path = os.path.join(tmp, "word", "document.xml")
            tree = etree.parse(doc_path)
            root = tree.getroot()

            # Ищем элемент по ID
            xpath = f".//w:p[@my:id='{elem_id}']//w:t"
            namespaces = {'w': W_NS, 'my': MY_NS}
            texts = root.xpath(xpath, namespaces=namespaces)

            if texts:
                return ' '.join([t.text for t in texts if t.text])[:100]
            return "<нет текста>"

    def test_delete_random_paragraphs_with_preview(self):
        """Тест 1: Удаление случайных параграфов с наглядным выводом."""
        print("\n" + "=" * 70)
        print("ТЕСТ 1: Удаление случайных параграфов (с превью)")
        print("=" * 70)

        # Получаем все параграфы из content
        paragraphs = [item for item in self.data['content']
                      if item['type'] == 'paragraph']

        # Показываем первые несколько параграфов для контекста
        print("\n📋 Первые 5 параграфов в оригинале:")
        for i, p in enumerate(paragraphs[:5]):
            preview = self._get_element_text_preview(self.original_docx, p['id'])
            print(f"  {p['id']}: {preview}")

        # Выбираем случайные 30% для удаления
        random.seed(42)
        to_delete = random.sample(paragraphs, k=max(1, len(paragraphs) // 3))

        # Устанавливаем флаг deleted
        deleted_ids = []
        for p in to_delete:
            p['deleted'] = True
            deleted_ids.append(p['id'])

        #print(f"\n🗑️ Помечено на удаление ({len(deleted_ids)} параграфов):")
        for pid in deleted_ids[:5]:  # Показываем первые 5
            preview = self._get_element_text_preview(self.original_docx, pid)
            print(f"  {pid}: {preview}")
        if len(deleted_ids) > 5:
            print(f"  ... и ещё {len(deleted_ids) - 5}")

        # Сохраняем модифицированный JSON
        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        # Запускаем реконструктор
        cmd = [
            "python", "src/reconstructor.py",
            "--in-json", self.test_json,
            "--out-docx", self.test_output,
            "--donor-docx", self.original_docx
        ]

        print(f"\n🚀 Запуск реконструктора...")
        result = subprocess.run(cmd, capture_output=True, text=True)

        # Проверяем успешность
        self.assertEqual(result.returncode, 0,
                         f"Реконструктор завершился с ошибкой:\n{result.stderr}")

        # Сохраняем копию результата
        saved_path = self._save_output_copy("deleted_paragraphs")

        # Показываем превью результата
        self._print_xml_preview(self.test_output)

        # Получаем ID из выходного DOCX
        output_ids = self._get_element_ids_from_docx(self.test_output)

        # Проверяем удаление
        print("\n🔍 Проверка результатов:")
        for pid in deleted_ids[:3]:  # Проверяем первые 3 удалённых
            self.assertNotIn(pid, output_ids)
            print(f"  ✅ {pid} успешно удалён")

        # Проверяем, что остальные присутствуют
        remaining_ids = [p['id'] for p in paragraphs if p['id'] not in deleted_ids]
        for pid in remaining_ids[:3]:  # Проверяем первые 3 оставшихся
            self.assertIn(pid, output_ids)
            preview = self._get_element_text_preview(self.test_output, pid)
            print(f"  ✅ {pid} сохранён: {preview}")

        print(f"\n✅ Тест пройден: удалено {len(deleted_ids)} параграфов")
        print(f"   Результат сохранён в: {saved_path}")

    def test_delete_specific_paragraphs_by_position(self):
        """Тест 2: Удаление конкретных параграфов (первый, последний, средний)."""
        print("\n" + "=" * 70)
        print("ТЕСТ 2: Удаление конкретных параграфов по позициям")
        print("=" * 70)

        # Получаем все параграфы
        paragraphs = [item for item in self.data['content']
                      if item['type'] == 'paragraph']

        if len(paragraphs) < 3:
            print("⚠️ Недостаточно параграфов для теста")
            return

        # Удаляем первый, последний и средний параграф
        to_delete = [
            paragraphs[0],                    # первый
            paragraphs[-1],                    # последний
            paragraphs[len(paragraphs)//2]      # средний
        ]

        deleted_ids = []
        for p in to_delete:
            p['deleted'] = True
            deleted_ids.append(p['id'])
            preview = self._get_element_text_preview(self.original_docx, p['id'])
            print(f"🗑️ Помечен на удаление: {p['id']} - {preview}")

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
        self.assertEqual(result.returncode, 0)

        # Сохраняем результат
        saved_path = self._save_output_copy("deleted_specific")

        # Проверяем
        output_ids = self._get_element_ids_from_docx(self.test_output)

        for pid in deleted_ids:
            self.assertNotIn(pid, output_ids)
            print(f"✅ {pid} успешно удалён")

        print(f"\n✅ Тест пройден, результат: {saved_path}")

    def test_delete_paragraphs_with_visual_diff(self):
        """Тест 3: Удаление с визуальным сравнением (выводит diff)."""
        print("\n" + "=" * 70)
        print("ТЕСТ 3: Удаление с визуальным сравнением")
        print("=" * 70)

        # Создаём JSON с удалением нескольких параграфов
        test_json = {
            "meta": {"schema_version": "2.15"},
            "document_info": {},
            "numbering_definitions": {},
            "doc_defaults": {"p_format": {}, "r_format": {}},
            "latent_styles": {},
            "styles": {},
            "content": []
        }

        # Берём оригинальный контент и помечаем некоторые на удаление
        for i, item in enumerate(self.data['content']):
            if item['type'] == 'paragraph':
                # Удаляем каждый 3-й параграф
                if i % 3 == 0:
                    item['deleted'] = True
                    #print(f"🗑️ Параграф {item['id']} помечен на удаление")
            test_json['content'].append(item)

        # Сохраняем тестовый JSON
        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(test_json, f, indent=2, ensure_ascii=False)

        # Запускаем реконструктор
        cmd = [
            "python", "src/reconstructor.py",
            "--in-json", self.test_json,
            "--out-docx", self.test_output,
            "--donor-docx", self.original_docx
        ]

        result = subprocess.run(cmd, capture_output=True, text=True)
        self.assertEqual(result.returncode, 0)

        # Сохраняем результат
        saved_path = self._save_output_copy("visual_diff")

        # Показываем статистику
        original_paras, original_tables = self._count_root_elements_in_docx(self.original_docx)
        new_paras, new_tables = self._count_root_elements_in_docx(self.test_output)

        print(f"\n📊 Статистика:")
        print(f"  Оригинал: {original_paras} параграфов, {original_tables} таблиц")
        print(f"  Результат: {new_paras} параграфов, {new_tables} таблиц")
        print(f"  Удалено: {original_paras - new_paras} параграфов")

        print(f"\n✅ Тест пройден, результат: {saved_path}")


if __name__ == '__main__':
    unittest.main(verbosity=2)