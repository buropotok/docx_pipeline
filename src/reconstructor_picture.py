# reconstructor_picture.py
# Вставка изображений в document.xml
# Использует существующие relation_id из JSON, не создаёт новые отношения

from __future__ import annotations

from typing import Dict, Any, Tuple
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def add_picture_to_document(
        run_data: Dict[str, Any],
        parent_element: etree._Element,  # w:r
        next_drawing_id: int,
) -> int:
    """
    Добавляет элемент <w:drawing> в parent_element, используя существующий relation_id.
    Возвращает обновлённый next_drawing_id.
    """
    # Проверяем обязательные поля
    relation_id = run_data.get("relation_id")
    if not relation_id:
        raise ValueError("Picture run missing relation_id")

    extent = run_data.get("extent")
    if not isinstance(extent, dict):
        raise ValueError("Picture run missing extent")

    cx = extent.get("cx")
    cy = extent.get("cy")
    if cx is None or cy is None:
        raise ValueError("Picture run extent missing cx/cy")

    # Создаём w:drawing
    drawing = etree.SubElement(parent_element, f"{{{W_NS}}}drawing")

    # wp:inline
    inline = etree.SubElement(
        drawing,
        f"{{{WP_NS}}}inline",
        nsmap={"wp": WP_NS, "a": A_NS, "pic": PIC_NS, "r": R_NS},
    )
    inline.set("distT", "0")
    inline.set("distB", "0")
    inline.set("distL", "0")
    inline.set("distR", "0")

    # wp:extent
    ext_el = etree.SubElement(inline, f"{{{WP_NS}}}extent")
    ext_el.set("cx", str(cx))
    ext_el.set("cy", str(cy))

    # wp:effectExtent
    eff = etree.SubElement(inline, f"{{{WP_NS}}}effectExtent")
    eff.set("l", "0")
    eff.set("t", "0")
    eff.set("r", "0")
    eff.set("b", "0")

    # wp:docPr (уникальный id)
    doc_pr = etree.SubElement(inline, f"{{{WP_NS}}}docPr")
    doc_pr.set("id", str(next_drawing_id))
    doc_pr.set("name", f"Picture {next_drawing_id}")
    next_drawing_id += 1

    # wp:cNvGraphicFramePr
    c_nv = etree.SubElement(inline, f"{{{WP_NS}}}cNvGraphicFramePr")
    locks = etree.SubElement(c_nv, f"{{{A_NS}}}graphicFrameLocks")
    locks.set("noChangeAspect", "1")

    # a:graphic
    graphic = etree.SubElement(inline, f"{{{A_NS}}}graphic")
    graphic_data = etree.SubElement(graphic, f"{{{A_NS}}}graphicData")
    graphic_data.set("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture")

    # pic:pic
    pic = etree.SubElement(graphic_data, f"{{{PIC_NS}}}pic")

    # pic:nvPicPr
    nv_pic_pr = etree.SubElement(pic, f"{{{PIC_NS}}}nvPicPr")
    c_nv_pr = etree.SubElement(nv_pic_pr, f"{{{PIC_NS}}}cNvPr")
    c_nv_pr.set("id", str(next_drawing_id - 1))  # тот же id, что и у docPr
    c_nv_pr.set("name", run_data.get("file", ""))

    c_nv_pic_pr = etree.SubElement(nv_pic_pr, f"{{{PIC_NS}}}cNvPicPr")
    pic_locks = etree.SubElement(c_nv_pic_pr, f"{{{A_NS}}}picLocks")
    pic_locks.set("noChangeAspect", "1")

    # pic:blipFill
    blip_fill = etree.SubElement(pic, f"{{{PIC_NS}}}blipFill")
    blip = etree.SubElement(blip_fill, f"{{{A_NS}}}blip")
    blip.set(f"{{{R_NS}}}embed", relation_id)  # используем существующий rId

    stretch = etree.SubElement(blip_fill, f"{{{A_NS}}}stretch")
    etree.SubElement(stretch, f"{{{A_NS}}}fillRect")

    # pic:spPr
    sp_pr = etree.SubElement(pic, f"{{{PIC_NS}}}spPr")
    xfrm = etree.SubElement(sp_pr, f"{{{A_NS}}}xfrm")
    off = etree.SubElement(xfrm, f"{{{A_NS}}}off")
    off.set("x", "0")
    off.set("y", "0")
    aext = etree.SubElement(xfrm, f"{{{A_NS}}}ext")
    aext.set("cx", str(cx))
    aext.set("cy", str(cy))

    prst = etree.SubElement(sp_pr, f"{{{A_NS}}}prstGeom")
    prst.set("prst", "rect")
    etree.SubElement(prst, f"{{{A_NS}}}avLst")

    return next_drawing_id