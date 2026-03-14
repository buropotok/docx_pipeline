import unittest
from lxml import etree

from docx_pipeline.pipeline.parse.parser_shape import parse_shape_node
from docx_pipeline.pipeline.parse.parser_picture import parse_picture_node
from docx_pipeline.pipeline.parse.parser_media_classifier import classify_media_node

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
V_NS = "urn:schemas-microsoft-com:vml"


class _StubParser:
    def _parse_paragraph_element(self, p, parent_id, index):
        return {
            "type": "paragraph",
            "id": f"{parent_id}.p_{index}",
            "parent_id": parent_id,
            "runs": []
        }


def _node(xml: str):
    return etree.fromstring(xml.encode("utf-8"))


class TestMediaNodeClassification(unittest.TestCase):
    def setUp(self):
        self.parser = _StubParser()
        self.relationships = {"rId1": "media/image1.png"}
        self.run_id = "p_1.run_1"

    def test_case_table(self):
        """
        Coverage table:
        - VML image -> picture
        - VML textbox -> shape
        - DrawingML inline picture -> picture
        - DrawingML anchor shape -> shape
        - Shape with text -> shape
        """
        cases = [
            (
                "VML image",
                f'''<w:pict xmlns:w="{W_NS}" xmlns:v="{V_NS}" xmlns:r="{R_NS}">
                        <v:shape style="width:10pt;height:20pt;">
                            <v:imagedata r:id="rId1"/>
                        </v:shape>
                    </w:pict>''',
                "picture",
            ),
            (
                "VML textbox",
                f'''<w:pict xmlns:w="{W_NS}" xmlns:v="{V_NS}">
                        <v:shape style="position:absolute;width:20pt;height:10pt;">
                            <v:textbox>
                              <w:txbxContent><w:p/></w:txbxContent>
                            </v:textbox>
                        </v:shape>
                    </w:pict>''',
                "shape",
            ),
            (
                "DrawingML inline picture",
                f'''<w:drawing xmlns:w="{W_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:pic="{PIC_NS}" xmlns:r="{R_NS}">
                        <wp:inline>
                          <wp:extent cx="100" cy="200"/>
                          <a:graphic>
                            <a:graphicData>
                              <pic:pic>
                                <pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill>
                              </pic:pic>
                            </a:graphicData>
                          </a:graphic>
                        </wp:inline>
                    </w:drawing>''',
                "picture",
            ),
            (
                "DrawingML anchor shape",
                f'''<w:drawing xmlns:w="{W_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}">
                        <wp:anchor>
                          <wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>
                          <wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>
                          <a:graphic>
                            <a:graphicData>
                              <a:sp>
                                <a:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></a:spPr>
                              </a:sp>
                            </a:graphicData>
                          </a:graphic>
                        </wp:anchor>
                    </w:drawing>''',
                "shape",
            ),
            (
                "Shape with text",
                f'''<w:drawing xmlns:w="{W_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}">
                        <wp:anchor>
                          <a:graphic>
                            <a:graphicData>
                              <a:sp>
                                <a:txBody>
                                  <a:p/>
                                </a:txBody>
                              </a:sp>
                            </a:graphicData>
                          </a:graphic>
                        </wp:anchor>
                    </w:drawing>''',
                "shape",
            ),
        ]

        for name, xml, expected_kind in cases:
            with self.subTest(case=name):
                node = _node(xml)
                classification = classify_media_node(node)
                self.assertEqual(expected_kind, classification.kind)

                shape_run = parse_shape_node(node, self.run_id, self.parser)
                picture_run = parse_picture_node(node, self.run_id, self.relationships)

                if expected_kind == "shape":
                    self.assertIsNotNone(shape_run)
                    self.assertEqual("shape", shape_run["type"])
                    self.assertIsNone(picture_run)
                else:
                    self.assertIsNone(shape_run)
                    self.assertIsNotNone(picture_run)
                    self.assertEqual("picture", picture_run["type"])


if __name__ == "__main__":
    unittest.main()
