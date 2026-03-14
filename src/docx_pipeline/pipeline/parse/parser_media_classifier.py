from __future__ import annotations

from dataclasses import dataclass
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
V_NS = "urn:schemas-microsoft-com:vml"


def qn(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


@dataclass(frozen=True)
class MediaNodeClassification:
    kind: str
    has_text_content: bool
    has_vml_imagedata: bool
    has_drawingml_picture: bool
    has_drawingml_shape_container: bool
    has_wp_anchor: bool
    has_wp_inline: bool


def classify_media_node(node: etree._Element) -> MediaNodeClassification:
    """
    Classify media run node (w:drawing / w:pict) as picture or shape.

    Case table:
      - VML image (v:imagedata without textbox/text)         -> picture
      - VML textbox (v:textbox / w:txbxContent)              -> shape
      - DrawingML inline picture (wp:inline + pic:pic)       -> picture
      - DrawingML anchor shape (wp:anchor + a:sp/prstGeom)   -> shape
      - Any shape with text (w:txbxContent / a:txBody)       -> shape
    """
    if node is None:
        return MediaNodeClassification("unsupported", False, False, False, False, False, False)

    has_vml_imagedata = node.find(".//" + qn(V_NS, "imagedata")) is not None
    has_vml_textbox = node.find(".//" + qn(V_NS, "textbox")) is not None
    has_word_txbx_content = node.find(".//" + qn(W_NS, "txbxContent")) is not None
    has_drawingml_text_body = node.find(".//" + qn(A_NS, "txBody")) is not None
    has_text_content = has_vml_textbox or has_word_txbx_content or has_drawingml_text_body

    has_drawingml_picture = node.find(".//" + qn(PIC_NS, "pic")) is not None
    has_drawingml_shape_container = (
        node.find(".//" + qn(A_NS, "sp")) is not None or
        node.find(".//" + qn(A_NS, "prstGeom")) is not None
    )

    has_vml_shape_container = any(
        node.find(".//" + qn(V_NS, tag_name)) is not None
        for tag_name in ("rect", "roundrect", "oval", "line", "shape")
    )

    has_wp_anchor = node.find(".//" + qn(WP_NS, "anchor")) is not None
    has_wp_inline = node.find(".//" + qn(WP_NS, "inline")) is not None

    if has_text_content:
        kind = "shape"
    elif has_vml_imagedata:
        kind = "picture"
    elif has_drawingml_shape_container and not has_drawingml_picture:
        kind = "shape"
    elif has_wp_anchor and not has_drawingml_picture:
        kind = "shape"
    elif has_drawingml_picture:
        kind = "picture"
    elif has_vml_shape_container:
        kind = "shape"
    else:
        kind = "unsupported"

    return MediaNodeClassification(
        kind=kind,
        has_text_content=has_text_content,
        has_vml_imagedata=has_vml_imagedata,
        has_drawingml_picture=has_drawingml_picture,
        has_drawingml_shape_container=has_drawingml_shape_container,
        has_wp_anchor=has_wp_anchor,
        has_wp_inline=has_wp_inline,
    )

