import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MY_NS = "https://translatefactory/schema/custom-id"


def qn_w(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def qn_my(local: str) -> str:
    return f"{{{MY_NS}}}{local}"


class TestPPrGetOrCreate(unittest.TestCase):
    def setUp(self):
        self.base_json = "src/test_deep_seek/donor.json"
        self.base_docx = "src/test_deep_seek/donor.materialized.docx"
        self.tmp_dir = tempfile.mkdtemp(prefix="ppr_get_or_create_")
        self.test_json = os.path.join(self.tmp_dir, "test.json")
        self.test_donor = os.path.join(self.tmp_dir, "donor_without_jc_ta.docx")
        self.test_output = os.path.join(self.tmp_dir, "output.docx")

    def tearDown(self):
        shutil.rmtree(self.tmp_dir)

    def _first_paragraph_id(self, data):
        for item in data.get("content", []):
            if item.get("type") == "paragraph":
                return item.get("id")
        return None

    def _remove_jc_and_text_alignment(self, input_docx: str, para_id: str, out_docx: str) -> None:
        with tempfile.TemporaryDirectory(prefix="docx_patch_") as unzip_dir:
            with zipfile.ZipFile(input_docx, "r") as zin:
                zin.extractall(unzip_dir)

            doc_xml = os.path.join(unzip_dir, "word", "document.xml")
            tree = etree.parse(doc_xml)
            root = tree.getroot()
            namespaces = {"w": W_NS, "my": MY_NS}
            paras = root.xpath(f".//w:p[@my:id='{para_id}']", namespaces=namespaces)
            self.assertTrue(paras, f"Paragraph {para_id} not found in donor DOCX")

            p = paras[0]
            pPr = p.find(qn_w("pPr"))
            if pPr is None:
                pPr = etree.SubElement(p, qn_w("pPr"))

            jc = pPr.find(qn_w("jc"))
            if jc is not None:
                pPr.remove(jc)

            ta = pPr.find(qn_w("textAlignment"))
            if ta is not None:
                pPr.remove(ta)

            tree.write(doc_xml, encoding="utf-8", xml_declaration=True)

            with zipfile.ZipFile(out_docx, "w", zipfile.ZIP_DEFLATED) as zout:
                for dirpath, _, files in os.walk(unzip_dir):
                    for filename in files:
                        abs_path = os.path.join(dirpath, filename)
                        rel_path = os.path.relpath(abs_path, unzip_dir).replace("\\", "/")
                        zout.write(abs_path, rel_path)

    def test_alignment_and_text_alignment_are_created_when_absent(self):
        with open(self.base_json, "r", encoding="utf-8") as f:
            data = json.load(f)

        para_id = self._first_paragraph_id(data)
        self.assertIsNotNone(para_id, "No paragraph found in donor JSON")

        # Minimal RAW patch payload for pPr fields under test.
        for item in data["content"]:
            if item.get("id") == para_id:
                item["p_format"] = {
                    "alignment": "center",
                    "text_alignment": "bottom",
                }
                break

        with open(self.test_json, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        self._remove_jc_and_text_alignment(self.base_docx, para_id, self.test_donor)

        cmd = [
            sys.executable,
            "src/reconstructor.py",
            "--in-json",
            self.test_json,
            "--out-docx",
            self.test_output,
            "--donor-docx",
            self.test_donor,
        ]
        result = subprocess.run(cmd, capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, msg=f"stdout={result.stdout}\nstderr={result.stderr}")

        with tempfile.TemporaryDirectory(prefix="docx_out_") as unzip_dir:
            with zipfile.ZipFile(self.test_output, "r") as zf:
                zf.extractall(unzip_dir)

            doc_xml = os.path.join(unzip_dir, "word", "document.xml")
            tree = etree.parse(doc_xml)
            root = tree.getroot()
            namespaces = {"w": W_NS, "my": MY_NS}
            p = root.xpath(f".//w:p[@my:id='{para_id}']", namespaces=namespaces)[0]
            pPr = p.find(qn_w("pPr"))
            self.assertIsNotNone(pPr)

            jc = pPr.find(qn_w("jc"))
            self.assertIsNotNone(jc, "w:jc must be created when alignment is present in RAW")
            self.assertEqual(jc.get(qn_w("val")), "center")

            ta = pPr.find(qn_w("textAlignment"))
            self.assertIsNotNone(ta, "w:textAlignment must be created when text_alignment is present in RAW")
            self.assertEqual(ta.get(qn_w("val")), "bottom")


if __name__ == "__main__":
    unittest.main()
