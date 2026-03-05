import unittest
import json
import os
import tempfile
import shutil
import subprocess
import zipfile
from lxml import etree
import sys

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


class TestMinimalPFormat(unittest.TestCase):
    def setUp(self):
        self.original_docx = "src/test_deep_seek/donor.materialized.docx"
        self.test_dir = tempfile.mkdtemp()
        self.test_json = os.path.join(self.test_dir, "test.json")
        self.test_output = os.path.join(self.test_dir, "output.docx")

    def tearDown(self):
        shutil.rmtree(self.test_dir)

    def test_single_property_alignment(self):
        """Тест только с одним свойством - alignment"""
        with open("src/test_deep_seek/donor.json", 'r', encoding='utf-8') as f:
            data = json.load(f)

        # Найти первый параграф
        for item in data['content']:
            if item['type'] == 'paragraph':
                item['p_format'] = {"alignment": "center"}
                break

        with open(self.test_json, 'w', encoding='utf-8') as f:
            json.dump(data, f)

        cmd = [
            sys.executable, "src/reconstructor.py",
            "--in-json", self.test_json,
            "--out-docx", self.test_output,
            "--donor-docx", self.original_docx
        ]

        result = subprocess.run(cmd, capture_output=True, text=True)
        print(f"STDOUT: {result.stdout}")
        print(f"STDERR: {result.stderr}")

        # Проверяем что файл создался
        self.assertTrue(os.path.exists(self.test_output))

        # Проверяем что это валидный ZIP
        with zipfile.ZipFile(self.test_output, 'r') as zf:
            self.assertIn("word/document.xml", zf.namelist())


if __name__ == '__main__':
    unittest.main()

