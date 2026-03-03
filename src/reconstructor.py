# reconstructor.py
# Minimal reconstructor with indexing and root deletion (step 2)
# Accepts donor DOCX file directly (extracts to temp dir)gpt

from __future__ import annotations

import argparse
import json
import os
import zipfile
from copy import deepcopy
from typing import Any, Dict, List, Optional
from lxml import etree
import sys
import re
import tempfile
import shutil

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MY_NS = "https://translatefactory/schema/custom-id"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"  # for future use


def qn_w(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def qn_my(local: str) -> str:
    return f"{{{MY_NS}}}{local}"


class ReconstructorV215:
    def __init__(self, raw_json_path: str):
        self.raw_json_path = raw_json_path
        self.donor_raw_dir: Optional[str] = None
        self.package_files: Dict[str, bytes] = {}

        # JSON data
        with open(raw_json_path, "r", encoding="utf-8") as f:
            self.data = json.load(f)
        self.content: List[Dict[str, Any]] = self.data.get("content", [])

        # Indices to be built
        self.root_by_id: Dict[str, etree._Element] = {}
        self.row_by_id: Dict[str, etree._Element] = {}
        self.para_by_id: Dict[str, etree._Element] = {}
        self.original_children: List[etree._Element] = []

        self.temp_dir: Optional[str] = None

    def _copy_donor_files(self) -> None:
        """Copy all donor files except word/document.xml into package_files."""
        if not self.donor_raw_dir or not os.path.exists(self.donor_raw_dir):
            raise FileNotFoundError(f"donor_raw_dir not found: {self.donor_raw_dir}")

        exclude = {"word/document.xml"}
        for root, _, files in os.walk(self.donor_raw_dir):
            for fn in files:
                src_path = os.path.join(root, fn)
                rel_path = os.path.relpath(src_path, self.donor_raw_dir).replace("\\", "/")
                if rel_path in exclude:
                    continue
                if rel_path in self.package_files:
                    continue
                with open(src_path, "rb") as f:
                    self.package_files[rel_path] = f.read()

    def _build_indices(self, body: etree._Element) -> None:
        """
        Build indices of root elements, rows, and all paragraphs.
        Fills:
          self.root_by_id: root paragraphs and tables (by my:id)
          self.row_by_id: all table rows (by my:id)
          self.para_by_id: all paragraphs (root + nested) with generated ids
          self.original_children: list of body children (as in donor)
        """
        self.original_children = list(body)

        for elem in body:
            # Root paragraphs
            if elem.tag == qn_w("p"):
                pid = elem.get(qn_my("id"))
                if pid:
                    self.root_by_id[pid] = elem
                    self.para_by_id[pid] = elem
            # Root tables
            elif elem.tag == qn_w("tbl"):
                tid = elem.get(qn_my("id"))
                if tid:
                    self.root_by_id[tid] = elem
                    # Collect rows
                    for tr in elem.findall(qn_w("tr")):
                        row_id = tr.get(qn_my("id"))
                        if row_id:
                            self.row_by_id[row_id] = tr
                            # Process cells inside this row
                            cells = [tc for tc in tr if tc.tag == qn_w("tc")]
                            for ci, tc in enumerate(cells, start=1):
                                # Collect paragraphs inside this cell
                                paras = [p for p in tc if p.tag == qn_w("p")]
                                for pi, p in enumerate(paras, start=1):
                                    pid = f"{row_id}.cell_{ci}.p_{pi}"
                                    self.para_by_id[pid] = p

    def _apply_root_deletions(self, body: etree._Element) -> None:
        """
        Remove root elements marked as deleted in JSON.
        Modifies body children in place.
        """
        # Start with a copy of original children
        working = list(self.original_children)

        # Collect ids of elements to delete
        deleted_ids = set()
        for item in self.content:
            if item.get("deleted") is True:
                item_id = item.get("id")
                if not item_id:
                    raise ValueError("deleted item missing id")
                deleted_ids.add(item_id)

        # Remove elements from working list
        new_working = []
        for elem in working:
            # Check if this element (root) has an id and it's in deleted set
            elem_id = None
            if elem.tag in (qn_w("p"), qn_w("tbl")):
                elem_id = elem.get(qn_my("id"))
            if elem_id and elem_id in deleted_ids:
                # Skip this element (delete it)
                continue
            new_working.append(elem)

        # Replace body children with new list
        for ch in list(body):
            body.remove(ch)
        for ch in new_working:
            body.append(ch)

        # Optionally update original_children to reflect new state (for future steps)
        self.original_children = new_working

    def build_docx(self, out_docx_path: str, donor_docx_path: str) -> None:
        """
        Build reconstructed DOCX from donor DOCX and JSON.

        Args:
            out_docx_path: path to output DOCX file
            donor_docx_path: path to donor DOCX file (with my:id attributes)
        """
        # Create temporary directory for donor extraction
        self.temp_dir = tempfile.mkdtemp(prefix="reconstructor_")
        try:
            # Extract donor DOCX to temp directory
            print(f"Extracting {donor_docx_path} to {self.temp_dir}")
            with zipfile.ZipFile(donor_docx_path, "r") as z:
                z.extractall(self.temp_dir)

            self.donor_raw_dir = self.temp_dir

            # Copy all donor files except word/document.xml into package_files
            self._copy_donor_files()

            # Load donor document.xml
            donor_doc_path = os.path.join(self.temp_dir, "word", "document.xml")
            if not os.path.exists(donor_doc_path):
                raise FileNotFoundError(f"word/document.xml not found in {donor_docx_path}")

            parser = etree.XMLParser(remove_blank_text=False, recover=False, huge_tree=True)
            doc_tree = etree.parse(donor_doc_path, parser)
            doc_root = doc_tree.getroot()
            body = doc_root.find(qn_w("body"))
            if body is None:
                raise ValueError("Document has no w:body")

            # Build indices (does not modify XML)
            self._build_indices(body)

            # Apply root deletions (modifies body)
            self._apply_root_deletions(body)

            # Serialize document.xml (now with deletions) and add to package
            self.package_files["word/document.xml"] = etree.tostring(
                doc_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=None,
                pretty_print=False
            )

            # Create output zip
            os.makedirs(os.path.dirname(out_docx_path), exist_ok=True)
            with zipfile.ZipFile(out_docx_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for name in sorted(self.package_files.keys()):
                    zout.writestr(name, self.package_files[name])

            print(f"Reconstruction (step 2: root deletions) completed. Output: {out_docx_path}")

        finally:
            # Clean up temporary directory
            if self.temp_dir and os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir)
                print(f"Cleaned up temporary directory: {self.temp_dir}")

    def _cleanup(self) -> None:
        """Clean up temporary directory if it exists."""
        if self.temp_dir and os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
            self.temp_dir = None


def main() -> None:
    parser = argparse.ArgumentParser(description="Reconstructor v2.15 (step 2: root deletions)")
    parser.add_argument("--in-json", dest="input_json", required=True)
    parser.add_argument("--out-docx", dest="output_docx", required=True)
    parser.add_argument("--donor-docx", dest="donor_docx", required=True,
                        help="Path to donor DOCX file (with my:id attributes)")
    args = parser.parse_args()

    try:
        recon = ReconstructorV215(args.input_json)
        recon.build_docx(args.output_docx, args.donor_docx)
    except Exception:
        import traceback
        traceback.print_exc()
        if hasattr(recon, '_cleanup'):
            recon._cleanup()
        sys.exit(1)


if __name__ == "__main__":
    main()