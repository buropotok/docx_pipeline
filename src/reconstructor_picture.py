# reconstructor_picture.py
# Вставка изображений в document.xml + регистрация rels + копирование media part в пакет.
# Универсально и по контракту: без синтетических дефолтов.

from __future__ import annotations

from typing import Dict, Any, Optional, Tuple
import os
from lxml import etree

W_NS   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS   = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP_NS  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
R_NS   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _normalize_target(file_target: str) -> str:
    """
    Нормализуем file из RAW к виду 'media/<name.ext>'.
    Допускаем вход:
      - 'media/x.jpeg'
      - 'word/media/x.jpeg'
      - '/word/media/x.jpeg'
      - '../media/x.jpeg'
    """
    t = (file_target or "").strip().lstrip("/")
    if t.startswith("word/"):
        t = t[len("word/"):]
    t = os.path.normpath(t).replace("\\", "/")
    return t


def _parse_rid_num(rid: str) -> Optional[int]:
    # 'rId12' -> 12
    if not rid:
        return None
    if rid.startswith("rId"):
        s = rid[3:]
        if s.isdigit():
            return int(s)
    return None


def _pick_docpr_id(run_id: str, rel_id: str) -> int:
    """
    Дет. выбор wp:docPr/@id (должен быть >0).
    Лучше привязать к rel_id числу, чтобы было стабильно при одинаковом JSON.
    """
    n = _parse_rid_num(rel_id)
    if n is not None and n > 0:
        return n
    # fallback: попробуем извлечь число из run_id 'run_42'
    digits = "".join(ch for ch in (run_id or "") if ch.isdigit())
    if digits.isdigit():
        v = int(digits)
        return v if v > 0 else 1
    return 1


def _require_extent(run_data: Dict[str, Any]) -> Tuple[int, int]:
    extent = run_data.get("extent")
    if not isinstance(extent, dict):
        raise ValueError("Contract violation: picture run missing extent (no synthesis allowed).")
    cx = extent.get("cx")
    cy = extent.get("cy")
    if cx is None or cy is None:
        raise ValueError("Contract violation: picture run extent missing cx/cy (no synthesis allowed).")
    try:
        icx = int(cx)
        icy = int(cy)
    except Exception:
        raise ValueError("Contract violation: picture run extent cx/cy must be integers (EMU).")
    if icx <= 0 or icy <= 0:
        raise ValueError("Contract violation: picture run extent cx/cy must be > 0 (EMU).")
    return icx, icy


def add_picture_to_document(
    run_data: Dict[str, Any],
    parent_element: etree._Element,        # w:r
    relationships: Dict[str, str],         # map {rId: target} to be serialized into document.xml.rels
    package_files: Dict[str, bytes],       # zip path -> bytes
    donor_media_path: str,                # .../raw/donor/word/media
    next_rel_id: int,
) -> int:
    """
    Добавляет картинку:
      - DrawingML в parent_element (w:r)
      - relationship: Id=relation_id, Target=file (relative, 'media/..')
      - binary in zip at 'word/' + file ('word/media/..')

    Возвращает обновлённый next_rel_id.
    """
    file_target_raw = run_data.get("file")
    if not file_target_raw:
        raise ValueError("Contract violation: picture run missing file.")

    file_target = _normalize_target(str(file_target_raw))
    # Строго контрактно ожидаем media/...
    if not file_target.startswith("media/"):
        raise ValueError(f"Contract violation: picture file must be relative to word/document.xml (expected 'media/...'), got: {file_target}")

    # extent обязателен (у тебя он есть в старом JSON)
    cx, cy = _require_extent(run_data)

    # relation_id: если задан — используем; иначе выделяем новый
    rel_id = run_data.get("relation_id")
    if rel_id:
        rel_id = str(rel_id)
    else:
        rel_id = f"rId{next_rel_id}"
        next_rel_id += 1

    # если rel_id задан, двигаем next_rel_id вперёд, чтобы не было коллизий
    n = _parse_rid_num(rel_id)
    if n is not None and (n + 1) > next_rel_id:
        next_rel_id = n + 1

    # 1) кладём файл в пакет
    # donor_media_path хранит только basename, поэтому берем basename
    basename = os.path.basename(file_target)
    src_file = os.path.join(donor_media_path, basename)
    if not os.path.exists(src_file):
        raise ValueError(f"Contract violation: donor image not found: {src_file}")

    with open(src_file, "rb") as f:
        image_data = f.read()

    zip_path = "word/" + file_target  # word/media/<name>
    package_files[zip_path] = image_data

    # 2) relationship (Target должен быть 'media/<name>')
    relationships[rel_id] = file_target

    # 3) DrawingML
    drawing = etree.SubElement(parent_element, f"{{{W_NS}}}drawing")

    inline = etree.SubElement(
        drawing,
        f"{{{WP_NS}}}inline",
        nsmap={"wp": WP_NS, "a": A_NS, "pic": PIC_NS, "r": R_NS},
    )
    inline.set("distT", "0")
    inline.set("distB", "0")
    inline.set("distL", "0")
    inline.set("distR", "0")

    ext_el = etree.SubElement(inline, f"{{{WP_NS}}}extent")
    ext_el.set("cx", str(cx))
    ext_el.set("cy", str(cy))

    eff = etree.SubElement(inline, f"{{{WP_NS}}}effectExtent")
    eff.set("l", "0")
    eff.set("t", "0")
    eff.set("r", "0")
    eff.set("b", "0")

    doc_pr = etree.SubElement(inline, f"{{{WP_NS}}}docPr")
    docpr_id = _pick_docpr_id(run_data.get("id", ""), rel_id)
    doc_pr.set("id", str(docpr_id))
    doc_pr.set("name", f"Picture {docpr_id}")

    c_nv = etree.SubElement(inline, f"{{{WP_NS}}}cNvGraphicFramePr")
    locks = etree.SubElement(c_nv, f"{{{A_NS}}}graphicFrameLocks")
    locks.set("noChangeAspect", "1")

    graphic = etree.SubElement(inline, f"{{{A_NS}}}graphic")
    graphic_data = etree.SubElement(graphic, f"{{{A_NS}}}graphicData")
    graphic_data.set("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture")

    pic = etree.SubElement(graphic_data, f"{{{PIC_NS}}}pic")

    nv_pic_pr = etree.SubElement(pic, f"{{{PIC_NS}}}nvPicPr")
    c_nv_pr = etree.SubElement(nv_pic_pr, f"{{{PIC_NS}}}cNvPr")
    c_nv_pr.set("id", "0")
    c_nv_pr.set("name", basename)

    c_nv_pic_pr = etree.SubElement(nv_pic_pr, f"{{{PIC_NS}}}cNvPicPr")
    pic_locks = etree.SubElement(c_nv_pic_pr, f"{{{A_NS}}}picLocks")
    pic_locks.set("noChangeAspect", "1")

    blip_fill = etree.SubElement(pic, f"{{{PIC_NS}}}blipFill")
    blip = etree.SubElement(blip_fill, f"{{{A_NS}}}blip")
    blip.set(f"{{{R_NS}}}embed", rel_id)

    stretch = etree.SubElement(blip_fill, f"{{{A_NS}}}stretch")
    etree.SubElement(stretch, f"{{{A_NS}}}fillRect")

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
    etree.SubElement(prst, f"{{{A_NS}}}avLst")  # важно для валидности

    return next_rel_id