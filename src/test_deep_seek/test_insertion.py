"""
Тесты для проверки операции вставки новых параграфов.
Запуск: python -m unittest tests/test_insertion.py -v
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
import logging
logging.basicConfig(level=logging.DEBUG, format='%(message)s')
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MY_NS = "https://translatefactory/schema/custom-id"


def qn_w(local):
    return f"{{{W_NS}}}{local}"


def qn_my(local):
    return f"{{{MY_NS}}}{local}"


class TestParagraphInsertion(unittest.TestCase):
    """Тесты для проверки вставки новых параграфов."""

    @classmethod
    def setUpClass(cls):
        cls.original_json = "src/test_deep_seek/donor.json"
        cls.original_docx = "src/test_deep_seek/donor.materialized.docx"
        cls.results_dir = "test_results"
        os.makedirs(cls.results_dir, exist_ok=True)

        assert os.path.exists(cls.original_json), f"Файл не найден: {cls.original_json}"
        assert os.path.exists(cls.original_docx), f"Файл не найден: {cls.original_docx}"

    def setUp(self):
        self.test_dir = tempfile.mkdtemp(prefix="insertion_test_")
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

    def _get_element_ids_from_docx(self, docx_path):
        """Получает все my:id из document.xml."""
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

    def _verify_element_order(self, docx_path, expected_ids):
        """Проверяет порядок элементов."""
        with tempfile.TemporaryDirectory() as tmp:
            with zipfile.ZipFile(docx_path, 'r') as zf:
                zf.extractall(tmp)

            doc_path = os.path.join(tmp, "word", "document.xml")
            tree = etree.parse(doc_path)
            root = tree.getroot()
            body = root.find(qn_w("body"))

            actual_ids = []
            for elem in body:
                if elem.tag in (qn_w("p"), qn_w("tbl")):
                    elem_id = elem.get(qn_my("id"))
                    if elem_id:
                        actual_ids.append(elem_id)

            # Проверяем, что ожидаемые ID присутствуют в правильном порядке
            filtered_actual = [id for id in actual_ids if id in expected_ids]
            self.assertEqual(filtered_actual, expected_ids)
            return actual_ids

    def test_insert_paragraph_before_first(self):
        """Тест 1: Вставка нового параграфа перед первым."""
        print("\n" + "=" * 70)
        print("ТЕСТ 1: Вставка нового параграфа перед первым")
        print("=" * 70)

        # Находим первый параграф
        paragraphs = [item for item in self.data['content']
                      if item['type'] == 'paragraph']
        first_para = paragraphs[0]

        print(f"📌 Якорь: {first_para['id']} (первый параграф)")

        # Создаём новый параграф
        new_para = {
            "type": "paragraph",
            "id": f"{first_para['id']}.1",  # новый ID с точкой
            "p_style_id": first_para['p_style_id'],
            "p_format": first_para.get('p_format', {}),
            "runs": first_para.get('runs', []),
            "anchor": first_para['id'],
            "position": "before",
            "derive_from": first_para['id']
        }

        # Вставляем в начало списка
        self.data['content'].insert(0, new_para)

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
        saved_path = self._save_output_copy("insert_before_first")

        # Проверяем, что новый элемент появился
        output_ids = self._get_element_ids_from_docx(self.test_output)
        self.assertIn(new_para['id'], output_ids)

        # Проверяем порядок
        all_para_ids = [p['id'] for p in paragraphs]
        expected_order = [new_para['id']] + all_para_ids
        self._verify_element_order(self.test_output, expected_order)

        print(f"\n✅ Новый параграф {new_para['id']} вставлен перед {first_para['id']}")
        print(f"   Результат: {saved_path}")

    def test_insert_paragraph_after_last(self):
        """Тест 2: Вставка нового параграфа после последнего."""
        print("\n" + "=" * 70)
        print("ТЕСТ 2: Вставка нового параграфа после последнего")
        print("=" * 70)

        # Находим последний параграф
        paragraphs = [item for item in self.data['content']
                      if item['type'] == 'paragraph']
        last_para = paragraphs[-1]

        print(f"📌 Якорь: {last_para['id']} (последний параграф)")

        # Создаём новый параграф
        new_para = {
            "type": "paragraph",
            "id": f"{last_para['id']}.1",
            "p_style_id": last_para['p_style_id'],
            "p_format": last_para.get('p_format', {}),
            "runs": last_para.get('runs', []),
            "anchor": last_para['id'],
            "position": "after",
            "derive_from": last_para['id']
        }

        # Вставляем в конец
        self.data['content'].append(new_para)

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        cmd = [
            "python", "src/reconstructor.py",
            "--in-json", self.test_json,
            "--out-docx", self.test_output,
            "--donor-docx", self.original_docx
        ]

        result = subprocess.run(cmd, capture_output=True, text=True)
        self.assertEqual(result.returncode, 0)

        saved_path = self._save_output_copy("insert_after_last")

        output_ids = self._get_element_ids_from_docx(self.test_output)
        self.assertIn(new_para['id'], output_ids)

        print(f"\n✅ Новый параграф {new_para['id']} вставлен после {last_para['id']}")
        print(f"   Результат: {saved_path}")

    def test_insert_multiple_paragraphs(self):
        """Тест 3: Вставка нескольких новых параграфов в разные места."""
        print("\n" + "=" * 70)
        print("ТЕСТ 3: Вставка нескольких новых параграфов")
        print("=" * 70)

        # Находим опорные точки
        paragraphs = [item for item in self.data['content']
                      if item['type'] == 'paragraph']

        if len(paragraphs) < 3:
            print("⚠️ Недостаточно параграфов для теста")
            return

        anchor1 = paragraphs[0]  # первый
        anchor2 = paragraphs[2]  # третий

        print(f"📌 Якорь 1: {anchor1['id']}")
        print(f"📌 Якорь 2: {anchor2['id']}")

        # Создаём два новых параграфа
        new_para1 = {
            "type": "paragraph",
            "id": f"{anchor1['id']}.1",
            "p_style_id": anchor1['p_style_id'],
            "p_format": anchor1.get('p_format', {}),
            "runs": anchor1.get('runs', []),
            "anchor": anchor1['id'],
            "position": "after",
            "derive_from": anchor1['id']
        }

        new_para2 = {
            "type": "paragraph",
            "id": f"{anchor2['id']}.1",
            "p_style_id": anchor2['p_style_id'],
            "p_format": anchor2.get('p_format', {}),
            "runs": anchor2.get('runs', []),
            "anchor": anchor2['id'],
            "position": "before",
            "derive_from": anchor2['id']
        }

        # Вставляем в правильном порядке (порядок в JSON важен!)
        new_content = []
        for item in self.data['content']:
            if item['id'] == anchor1['id']:
                new_content.append(item)
                new_content.append(new_para1)
            elif item['id'] == anchor2['id']:
                new_content.append(new_para2)
                new_content.append(item)
            else:
                new_content.append(item)

        self.data['content'] = new_content

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        cmd = [
            "python", "src/reconstructor.py",
            "--in-json", self.test_json,
            "--out-docx", self.test_output,
            "--donor-docx", self.original_docx
        ]

        result = subprocess.run(cmd, capture_output=True, text=True)
        self.assertEqual(result.returncode, 0)

        saved_path = self._save_output_copy("insert_multiple")

        # Проверяем, что оба новых параграфа появились
        output_ids = self._get_element_ids_from_docx(self.test_output)
        self.assertIn(new_para1['id'], output_ids)
        self.assertIn(new_para2['id'], output_ids)

        print(f"\n✅ Новые параграфы вставлены:")
        print(f"   {new_para1['id']} после {anchor1['id']}")
        print(f"   {new_para2['id']} перед {anchor2['id']}")
        print(f"   Результат: {saved_path}")

    def test_insert_with_custom_content(self):
        """Тест 4: Вставка параграфа с изменённым содержимым."""
        print("\n" + "=" * 70)
        print("ТЕСТ 4: Вставка параграфа с изменённым содержимым")
        print("=" * 70)

        # Находим параграф-донор
        paragraphs = [item for item in self.data['content']
                      if item['type'] == 'paragraph']
        donor = paragraphs[0]

        print(f"📌 Донор: {donor['id']}")

        # Создаём новый параграф с изменённым текстом
        new_para_id = f"{donor['id']}.1"
        new_para = {
            "type": "paragraph",
            "id": new_para_id,
            "p_style_id": donor['p_style_id'],
            "p_format": donor.get('p_format', {}),
            "runs": [
                {
                    "type": "text",
                    "text": "ЭТО НОВЫЙ ПАРАГРАФ, ВСТАВЛЕННЫЙ ТЕСТОМ!",
                    "id": f"{new_para_id}.run_1",
                    "parent_id": new_para_id,
                    "r_format": {
                        "bold": True,
                        "color": "FF0000"
                    }
                }
            ],
            "anchor": donor['id'],
            "position": "after",
            "derive_from": donor['id']
        }

        # Вставляем после донора
        new_content = []
        for item in self.data['content']:
            new_content.append(item)
            if item['id'] == donor['id']:
                new_content.append(new_para)

        self.data['content'] = new_content

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

        print(f"\n🚀 Запуск реконструктора...")
        result = subprocess.run(cmd, capture_output=True, text=True)

        # Проверяем успешность
        self.assertEqual(result.returncode, 0,
                         f"Реконструктор завершился с ошибкой:\n{result.stderr}")

        # ✅ Сохраняем DOCX в папку с результатами
        saved_path = self._save_output_copy("insert_custom")

        print(f"\n✅ Новый параграф с кастомным текстом вставлен")
        print(f"   Результат сохранён в: {saved_path}")
        print("\n💡 Откройте файл в Word и найдите красный жирный текст!")


    def test_insert_two_paragraphs_after_anchor(self):
        """Тест 5: Вставка двух новых параграфов ПОСЛЕ якорного."""
        print("\n" + "=" * 70)
        print("ТЕСТ 5: Вставка 2 параграфов ПОСЛЕ якорного")
        print("=" * 70)

        # Находим якорный параграф (например, p_1)
        paragraphs = [item for item in self.data['content']
                      if item['type'] == 'paragraph']
        anchor = paragraphs[0]
        anchor_id = anchor['id']

        print(f"📌 Якорь: {anchor_id}")

        # Создаём два новых параграфа
        new_para1 = {
            "type": "paragraph",
            "id": f"{anchor_id}.1",
            "p_style_id": anchor['p_style_id'],
            "p_format": anchor.get('p_format', {}),
            "runs": [
                {
                    "type": "text",
                    "text": "Новый параграф 1",
                    "id": f"{anchor_id}.1.run_1",
                    "parent_id": f"{anchor_id}.1",
                    "r_format": {
                        "bold": True,
                        "color": "FF0000",
                        "underline": "single"
                    }
                }
            ],
            "anchor": anchor_id,
            "position": "after",
            "derive_from": anchor_id
        }

        new_para2 = {
            "type": "paragraph",
            "id": f"{anchor_id}.2",
            "p_style_id": anchor['p_style_id'],
            "p_format": anchor.get('p_format', {}),
            "runs": [
                {
                    "type": "text",
                    "text": "Новый параграф 2",
                    "id": f"{anchor_id}.2.run_1",
                    "parent_id": f"{anchor_id}.2",
                    "r_format": {
                        "bold": True,
                        "color": "FF0000",
                        "underline": "single"
                    }
                }
            ],
            "anchor": anchor_id,
            "position": "after",
            "derive_from": anchor_id
        }

        # Сохраняем оригинальный контент для сравнения
        original_content = self.data['content'].copy()

        # Вставляем новые параграфы в правильном порядке
        new_content = []
        for item in original_content:
            new_content.append(item)
            if item['id'] == anchor_id:
                new_content.append(new_para1)  # сначала 1
                new_content.append(new_para2)  # потом 2

        self.data['content'] = new_content

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
        saved_path = self._save_output_copy("insert_two_after")

        # ПРОВЕРКА 1: Порядок элементов в XML
        with tempfile.TemporaryDirectory() as tmp:
            with zipfile.ZipFile(self.test_output, 'r') as zf:
                zf.extractall(tmp)

            doc_path = os.path.join(tmp, "word", "document.xml")
            tree = etree.parse(doc_path)
            root = tree.getroot()
            body = root.find(qn_w("body"))

            # Собираем все ID корневых элементов в порядке их следования
            actual_ids = []
            print("\n📋 ПОЛНЫЙ СПИСОК ЭЛЕМЕНТОВ BODY (с типами):")
            for i, elem in enumerate(body):
                elem_id = elem.get(qn_my("id")) if elem.tag in (qn_w("p"), qn_w("tbl")) else "None"
                elem_tag = elem.tag
                print(f"  {i}: {elem_tag} - {elem_id}")
            for elem in body:
                if elem.tag in (qn_w("p"), qn_w("tbl")):
                    elem_id = elem.get(qn_my("id"))
                    if elem_id:
                        actual_ids.append(elem_id)
            print("\n📋 Полный список элементов в body:")
            for i, eid in enumerate(actual_ids):
                print(f"  {i}: {eid}")

            # Найдите все passthrough-элементы между якорем и параграфом 1
            print("\n🔍 Элементы между якорем и параграфом 1:")
            anchor_found = False
            for elem in body:
                if elem.tag in (qn_w("p"), qn_w("tbl")):
                    if elem.get(qn_my("id")) == anchor_id:
                        anchor_found = True
                        continue
                if anchor_found and elem.tag not in (qn_w("p"), qn_w("tbl")):
                    print(f"  {elem.tag}: {etree.tostring(elem)[:100]}...")

            # Находим позиции элементов
            anchor_pos = actual_ids.index(anchor_id)
            para1_pos = actual_ids.index(new_para1['id'])
            para2_pos = actual_ids.index(new_para2['id'])

            print(f"\n🔍 Проверка порядка следования:")
            print(f"   Позиция якоря: {anchor_pos}")
            print(f"   Позиция параграфа 1: {para1_pos}")
            print(f"   Позиция параграфа 2: {para2_pos}")

            # Проверяем, что оба новых параграфа идут сразу после якоря
            self.assertEqual(para1_pos, anchor_pos + 1,
                             f"Параграф 1 должен быть сразу после якоря")
            self.assertEqual(para2_pos, anchor_pos + 2,
                             f"Параграф 2 должен быть сразу после параграфа 1")

            # Проверяем порядок между собой
            self.assertLess(para1_pos, para2_pos,
                            "Параграф 1 должен быть перед параграфом 2")

        print(f"\n✅ Порядок следования подтверждён")
        print(f"   Результат сохранён в: {saved_path}")


    def test_insert_two_paragraphs_before_anchor(self):
        """Тест 6: Вставка двух новых параграфов ПЕРЕД якорным."""
        print("\n" + "=" * 70)
        print("ТЕСТ 6: Вставка 2 параграфов ПЕРЕД якорным")
        print("=" * 70)

        # Находим якорный параграф
        paragraphs = [item for item in self.data['content']
                      if item['type'] == 'paragraph']
        anchor = paragraphs[2]  # берём третий параграф, чтобы было место перед ним
        anchor_id = anchor['id']

        print(f"📌 Якорь: {anchor_id}")

        # Создаём два новых параграфа
        new_para1 = {
            "type": "paragraph",
            "id": f"{anchor_id}.before1",
            "p_style_id": anchor['p_style_id'],
            "p_format": anchor.get('p_format', {}),
            "runs": [
                {
                    "type": "text",
                    "text": "Новый параграф 1 (перед)",
                    "id": f"{anchor_id}.before1.run_1",
                    "parent_id": f"{anchor_id}.before1",
                    "r_format": {
                        "bold": True,
                        "color": "FF0000",
                        "underline": "single"
                    }
                }
            ],
            "anchor": anchor_id,
            "position": "before",
            "derive_from": anchor_id
        }

        new_para1 = {
            "type": "paragraph",
            "id": f"{anchor_id}.1",  # числовой суффикс
            "p_style_id": anchor['p_style_id'],
            "p_format": anchor.get('p_format', {}),
            "runs": [
                {
                    "type": "text",
                    "text": "Новый параграф 1 (перед)",
                    "id": f"{anchor_id}.1.run_1",  # обновляем и run ID
                    "parent_id": f"{anchor_id}.1",
                    "r_format": {
                        "bold": True,
                        "color": "FF0000",
                        "underline": "single"
                    }
                }
            ],
            "anchor": anchor_id,
            "position": "before",
            "derive_from": anchor_id
        }

        new_para2 = {
            "type": "paragraph",
            "id": f"{anchor_id}.2",  # числовой суффикс
            "p_style_id": anchor['p_style_id'],
            "p_format": anchor.get('p_format', {}),
            "runs": [
                {
                    "type": "text",
                    "text": "Новый параграф 2 (перед)",
                    "id": f"{anchor_id}.2.run_1",  # обновляем и run ID
                    "parent_id": f"{anchor_id}.2",
                    "r_format": {
                        "bold": True,
                        "color": "FF0000",
                        "underline": "single"
                    }
                }
            ],
            "anchor": anchor_id,
            "position": "before",
            "derive_from": anchor_id
        }

        # Сохраняем оригинальный контент
        original_content = self.data['content'].copy()

        # Вставляем новые параграфы перед якорем
        # ВАЖНО: порядок в JSON определяет порядок в документе!
        new_content = []
        for item in original_content:
            if item['id'] == anchor_id:
                # Сначала вставляем оба новых (в порядке JSON)
                new_content.append(new_para1)
                new_content.append(new_para2)
                new_content.append(item)
            else:
                new_content.append(item)

        self.data['content'] = new_content

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

        print(f"\n🚀 Запуск команды: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True)

        # Выводим всегда, не только при ошибке
        print("\n📋 РЕЗУЛЬТАТ РЕКОНСТРУКТОРА:")
        print(f"Return code: {result.returncode}")
        print(f"STDERR: {result.stderr}")
        print(f"STDOUT: {result.stdout}")

        # Сохраняем лог в файл на всякий случай
        with open(os.path.join(self.test_dir, "reconstructor.log"), "w", encoding='utf-8') as f:
            f.write(f"Return code: {result.returncode}\n")
            f.write(f"STDOUT:\n{result.stdout}\n")
            f.write(f"STDERR:\n{result.stderr}\n")

        self.assertEqual(result.returncode, 0)

        # Сохраняем результат
        saved_path = self._save_output_copy("insert_two_before")

        # ПРОВЕРКА 1: Порядок элементов в XML
        with tempfile.TemporaryDirectory() as tmp:
            with zipfile.ZipFile(self.test_output, 'r') as zf:
                zf.extractall(tmp)

            doc_path = os.path.join(tmp, "word", "document.xml")
            tree = etree.parse(doc_path)
            root = tree.getroot()
            body = root.find(qn_w("body"))

            # Собираем все ID корневых элементов в порядке их следования
            actual_ids = []
            for elem in body:
                if elem.tag in (qn_w("p"), qn_w("tbl")):
                    elem_id = elem.get(qn_my("id"))
                    if elem_id:
                        actual_ids.append(elem_id)

            # Находим позиции элементов
            anchor_pos = actual_ids.index(anchor_id)
            para1_pos = actual_ids.index(new_para1['id'])
            para2_pos = actual_ids.index(new_para2['id'])

            print(f"\n🔍 Проверка порядка следования:")
            print(f"   Позиция параграфа 1: {para1_pos}")
            print(f"   Позиция параграфа 2: {para2_pos}")
            print(f"   Позиция якоря: {anchor_pos}")

            # Проверяем, что оба новых параграфа идут перед якорем
            self.assertLess(para1_pos, anchor_pos,
                            "Параграф 1 должен быть перед якорем")
            self.assertLess(para2_pos, anchor_pos,
                            "Параграф 2 должен быть перед якорем")

            # Проверяем порядок между собой (должен соответствовать JSON)
            self.assertLess(para1_pos, para2_pos,
                            "Параграф 1 должен быть перед параграфом 2 (согласно JSON)")

        print(f"\n✅ Порядок следования подтверждён")
        print(f"   Результат сохранён в: {saved_path}")


    def test_verify_xml_structure_preserved(self):
        """Тест 7: Проверка, что операции не повредили XML структуру."""
        print("\n" + "=" * 70)
        print("ТЕСТ 7: Проверка сохранности XML структуры")
        print("=" * 70)

        # Берём первый параграф как якорь
        paragraphs = [item for item in self.data['content']
                      if item['type'] == 'paragraph']
        anchor = paragraphs[0]
        anchor_id = anchor['id']

        # Создаём один новый параграф (минимальные изменения)
        # Создаём один новый параграф с правильными ID run'ов
        new_para_id = f"{anchor_id}.1"
        new_para = {
            "type": "paragraph",
            "id": new_para_id,
            "p_style_id": anchor['p_style_id'],
            "p_format": anchor.get('p_format', {}),
            "runs": [
                {
                    "type": "text",
                    "text": "Тестовый параграф для проверки структуры",
                    "id": f"{new_para_id}.run_1",
                    "parent_id": new_para_id,
                    "r_format": {}
                }
            ],
            "anchor": anchor_id,
            "position": "after",
            "derive_from": anchor_id
        }

        # Вставляем новый параграф
        new_content = []
        for item in self.data['content']:
            new_content.append(item)
            if item['id'] == anchor_id:
                new_content.append(new_para)

        self.data['content'] = new_content

        # Сохраняем оригинальный XML для сравнения
        with tempfile.TemporaryDirectory() as tmp_orig:
            with zipfile.ZipFile(self.original_docx, 'r') as zf:
                zf.extractall(tmp_orig)
            orig_doc_path = os.path.join(tmp_orig, "word", "document.xml")
            with open(orig_doc_path, 'r', encoding='utf-8') as f:
                orig_xml = f.read()

        # Сохраняем JSON и запускаем
        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        cmd = [
            "python", "src/reconstructor.py",
            "--in-json", self.test_json,
            "--out-docx", self.test_output,
            "--donor-docx", self.original_docx
        ]

        print(f"\n🚀 Запуск команды: {' '.join(cmd)}")
        print(f"📁 JSON файл: {self.test_json}")

        # Проверим JSON
        with open(self.test_json, 'r', encoding='utf-8') as f:
            json_data = json.load(f)
            print(f"📄 JSON содержит {len(json_data.get('content', []))} элементов")
            # Покажем первые 2 элемента для проверки
            for i, item in enumerate(json_data.get('content', [])[:2]):
                print(f"   Элемент {i}: id={item.get('id')}, type={item.get('type')}")

        result = subprocess.run(cmd, capture_output=True, text=True)

        print(f"\n📊 РЕЗУЛЬТАТ:")
        print(f"  Return code: {result.returncode}")
        if result.stdout:
            print(f"  STDOUT: {result.stdout}")
        if result.stderr:
            print(f"  STDERR: {result.stderr}")
        self.assertEqual(result.returncode, 0)

        # Сохраняем результат
        saved_path = self._save_output_copy("structure_check")

        # ПРОВЕРКА: Сравниваем XML всех параграфов, которые не должны были измениться
        with tempfile.TemporaryDirectory() as tmp_result:
            with zipfile.ZipFile(self.test_output, 'r') as zf:
                zf.extractall(tmp_result)

            result_doc_path = os.path.join(tmp_result, "word", "document.xml")
            result_tree = etree.parse(result_doc_path)
            result_root = result_tree.getroot()
            result_body = result_root.find(qn_w("body"))

            # Для каждого оригинального параграфа (кроме тех, что были изменены)
            unchanged_paragraphs = [p for p in paragraphs if p['id'] != anchor_id]

            print(f"\n🔍 Проверка {len(unchanged_paragraphs)} неизменённых параграфов...")

            # Загружаем оригинальный XML для сравнения
            with tempfile.TemporaryDirectory() as tmp_orig:
                with zipfile.ZipFile(self.original_docx, 'r') as zf:
                    zf.extractall(tmp_orig)
                orig_doc_path = os.path.join(tmp_orig, "word", "document.xml")
                orig_tree = etree.parse(orig_doc_path)
                orig_root = orig_tree.getroot()

            namespaces = {'w': W_NS, 'my': MY_NS}

            for orig_para in unchanged_paragraphs:
                para_id = orig_para['id']

                # Находим параграф в оригинале
                orig_xpath = f".//w:p[@my:id='{para_id}']"
                orig_para_elem = orig_root.xpath(orig_xpath, namespaces=namespaces)
                self.assertEqual(len(orig_para_elem), 1,
                                 f"Параграф {para_id} не найден в оригинале")

                # Находим тот же параграф в результате
                result_xpath = f".//w:p[@my:id='{para_id}']"
                result_para_elem = result_body.xpath(result_xpath, namespaces=namespaces)
                self.assertEqual(len(result_para_elem), 1,
                                 f"Параграф {para_id} не найден в результате")

                # Сравниваем XML строки
                orig_xml_str = etree.tostring(orig_para_elem[0], encoding='unicode')
                result_xml_str = etree.tostring(result_para_elem[0], encoding='unicode')

                self.assertEqual(orig_xml_str, result_xml_str,
                                 f"Параграф {para_id} изменился")

                print(f"   ✓ {para_id} идентичен оригиналу")

        print(f"\n✅ Все неизменённые параграфы сохранили свою структуру")
        print(f"   Результат сохранён в: {saved_path}")


    def test_combined_deletion_and_insertion(self):
            """Тест 8: Комбинированная операция - удаление одних и вставка других."""
            print("\n" + "=" * 70)
            print("ТЕСТ 8: Удаление + вставка")
            print("=" * 70)

            paragraphs = [item for item in self.data['content']
                          if item['type'] == 'paragraph']

            # Берём три параграфа для теста
            p1 = paragraphs[0]  # будет удалён
            p2 = paragraphs[1]  # останется
            p3 = paragraphs[2]  # будет якорем для вставки

            print(f"📌 p1 ({p1['id']}) - будет УДАЛЁН")
            print(f"📌 p2 ({p2['id']}) - останется")
            print(f"📌 p3 ({p3['id']}) - якорь для вставки")

            # Помечаем p1 на удаление
            p1['deleted'] = True

            # Создаём новый параграф после p3
            new_para = {
                "type": "paragraph",
                "id": f"{p3['id']}.1",
                "p_style_id": p3['p_style_id'],
                "p_format": p3.get('p_format', {}),
                "runs": [
                    {
                        "type": "text",
                        "text": "НОВЫЙ ПАРАГРАФ ПОСЛЕ УДАЛЕНИЯ",
                        "id": f"{p3['id']}.1.run_1",
                        "parent_id": f"{p3['id']}.1",
                        "r_format": {
                            "bold": True,
                            "color": "FF0000"
                        }
                    }
                ],
                "anchor": p3['id'],
                "position": "after",
                "derive_from": p3['id']
            }

            # Строим новый контент
            new_content = []
            for item in self.data['content']:
                if item['id'] == p3['id']:
                    new_content.append(item)
                    new_content.append(new_para)
                else:
                    new_content.append(item)

            self.data['content'] = new_content

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

            print(f"\n🚀 Запуск команды: {' '.join(cmd)}")
            result = subprocess.run(cmd, capture_output=True, text=True)

            print(f"📋 РЕЗУЛЬТАТ РЕКОНСТРУКТОРА:")
            print(f"Return code: {result.returncode}")
            print(f"STDERR: {result.stderr}")
            print(f"STDOUT: {result.stdout}")
            self.assertEqual(result.returncode, 0)

            # Сохраняем результат
            saved_path = self._save_output_copy("combined")

            # Проверяем результат
            with tempfile.TemporaryDirectory() as tmp:
                with zipfile.ZipFile(self.test_output, 'r') as zf:
                    zf.extractall(tmp)

                doc_path = os.path.join(tmp, "word", "document.xml")
                tree = etree.parse(doc_path)
                root = tree.getroot()
                body = root.find(qn_w("body"))

                # Собираем все ID
                actual_ids = []
                for elem in body:
                    if elem.tag in (qn_w("p"), qn_w("tbl")):
                        elem_id = elem.get(qn_my("id"))
                        if elem_id:
                            actual_ids.append(elem_id)

                print(f"\n🔍 Элементы в результате:")
                for i, pid in enumerate(actual_ids):
                    print(f"   {i}: {pid}")

                # Проверяем, что p1 удалён
                self.assertNotIn(p1['id'], actual_ids,
                                 f"Параграф {p1['id']} должен быть удалён")

                # Проверяем, что новый параграф вставлен после p3
                p3_pos = actual_ids.index(p3['id'])
                new_pos = actual_ids.index(new_para['id'])
                self.assertEqual(new_pos, p3_pos + 1,
                                 "Новый параграф должен быть сразу после p3")

            print(f"\n✅ Комбинированная операция выполнена успешно")
            print(f"   Результат сохранён в: {saved_path}")

if __name__ == '__main__':
    unittest.main(verbosity=2)