# parser_picture.py
# Разбор изображений из DOCX (w:drawing, w:pict) в RAW JSON по контракту:
#   file: "media/<name.ext>" (Target из rels, нормализованный)
#   relation_id: "rIdX"
#   extent: {cx, cy} в EMU (только если извлечено, без синтетики)

from __future__ import annotations

from typing import Dict, Any, Optional, Tuple
import os
import re
from lxml import etree

W_NS  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS  = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
R_NS  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
V_NS  = "urn:schemas-microsoft-com:vml"

EMU_PER_PT = 12700
EMU_PER_IN = 914400
EMU_PER_PX_96DPI = 9525


def _safe_int(x: Optional[str]) -> Optional[int]:
    if x is None:
        return None
    try:
        return int(x)
    except Exception:
        return None


def _to_emu(val: float, unit: str) -> int:
    u = (unit or "pt").lower()
    if u == "pt":
        return int(val * EMU_PER_PT)
    if u == "in":
        return int(val * EMU_PER_IN)
    if u == "px":
        return int(val * EMU_PER_PX_96DPI)
    # неизвестная единица — трактуем как pt (детерминированно)
    return int(val * EMU_PER_PT)


def _normalize_rels_target(target: str) -> str:
    """
    Нормализуем Target из word/_rels/document.xml.rels к каноническому виду относительно word/document.xml:
    - убираем ведущие '/'
    - убираем префикс 'word/' если кто-то ошибочно его записал
    - нормализуем '../'
    Результат ожидаем вида 'media/<file>'.
    """
    t = (target or "").strip()
    t = t.lstrip("/")                 # "/word/media/x" -> "word/media/x"
    if t.startswith("word/"):
        t = t[len("word/"):]          # "media/x"
    t = os.path.normpath(t).replace("\\", "/")  # "../media/x" -> "media/x" (относительно)
    return t


def _extract_extent_from_drawing(node: etree._Element) -> Optional[Dict[str, int]]:
    """
    Для w:drawing: сначала wp:extent, если нет — пробуем a:xfrm/a:ext.
    Все значения в этих местах уже EMU.
    """
    extent = node.find(f".//{{{WP_NS}}}extent")
    if extent is not None:
        cx = _safe_int(extent.get("cx"))
        cy = _safe_int(extent.get("cy"))
        if cx is not None and cy is not None:
            return {"cx": cx, "cy": cy}

    # fallback: a:xfrm/a:ext (тоже EMU)
    aext = node.find(f".//{{{A_NS}}}xfrm/{{{A_NS}}}ext")
    if aext is not None:
        cx = _safe_int(aext.get("cx"))
        cy = _safe_int(aext.get("cy"))
        if cx is not None and cy is not None:
            return {"cx": cx, "cy": cy}

    return None


def _extract_rid_from_drawing(node: etree._Element) -> Optional[str]:
    blip = node.find(f".//{{{A_NS}}}blip")
    if blip is None:
        return None
    rid = blip.get(f"{{{R_NS}}}embed")
    return rid or None


def _extract_from_pict(node: etree._Element) -> Tuple[Optional[str], Optional[Dict[str, int]]]:
    """
    Для w:pict:
    - rid: v:imagedata/@r:id
    - extent: v:shape/@style (width/height)
    """
    rid = None
    imagedata = node.find(f".//{{{V_NS}}}imagedata")
    if imagedata is not None:
        rid = imagedata.get(f"{{{R_NS}}}id") or None

    extent = None
    shape = node.find(f".//{{{V_NS}}}shape")
    if shape is not None:
        style = shape.get("style") or ""
        w_match = re.search(r"width:\s*([\d.]+)\s*(pt|in|px)?", style, flags=re.IGNORECASE)
        h_match = re.search(r"height:\s*([\d.]+)\s*(pt|in|px)?", style, flags=re.IGNORECASE)
        if w_match and h_match:
            vw = float(w_match.group(1))
            uw = (w_match.group(2) or "pt").lower()
            vh = float(h_match.group(1))
            uh = (h_match.group(2) or "pt").lower()
            extent = {"cx": _to_emu(vw, uw), "cy": _to_emu(vh, uh)}

    return rid, extent


def parse_picture_node(
    node: etree._Element,
    run_id: str,
    relationships: Dict[str, str],
) -> Optional[Dict[str, Any]]:
    """
    Возвращает run-словарь для picture согласно контракту или None.

    Обязательные поля при успехе:
      id, type="picture", file, relation_id
    Опционально:
      extent {cx, cy} (EMU), если удалось извлечь из OOXML.
    """
    if node is None:
        return None

    rid: Optional[str] = None
    extent: Optional[Dict[str, int]] = None

    if node.tag.endswith("drawing"):
        rid = _extract_rid_from_drawing(node)
        extent = _extract_extent_from_drawing(node)
    elif node.tag.endswith("pict"):
        rid, extent = _extract_from_pict(node)
    else:
        return None

    if not rid:
        return None

    target = relationships.get(rid)
    if not target:
        return None

    norm_target = _normalize_rels_target(target)

    # Контракт: храним ровно target относительно word/document.xml, обычно "media/<file>"
    result: Dict[str, Any] = {
        "id": run_id,
        "type": "picture",
        "relation_id": rid,
        "file": norm_target,
    }
    if extent is not None:
        result["extent"] = extent

    return result