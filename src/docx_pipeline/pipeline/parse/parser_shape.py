
from lxml import etree
import re

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
V_NS = "urn:schemas-microsoft-com:vml"
W10_NS = "urn:schemas-microsoft-com:office:word"


def qn(ns, tag):
    return f"{{{ns}}}{tag}"


def _pt_to_emu(pt_value):
    return int(round(float(pt_value) * 12700))


def _parse_vml_style_extent(node):
    """
    Extract width/height from VML style="...width:74.5pt;height:97.65pt;..."
    and convert to EMU.
    """
    v_nodes = (
        node.findall(".//" + qn(V_NS, "rect")) +
        node.findall(".//" + qn(V_NS, "roundrect")) +
        node.findall(".//" + qn(V_NS, "oval")) +
        node.findall(".//" + qn(V_NS, "line")) +
        node.findall(".//" + qn(V_NS, "shape"))
    )

    for v_node in v_nodes:
        style = v_node.get("style") or ""
        if not style:
            continue

        m_w = re.search(r"width\s*:\s*([0-9]+(?:\.[0-9]+)?)pt", style)
        m_h = re.search(r"height\s*:\s*([0-9]+(?:\.[0-9]+)?)pt", style)

        if m_w and m_h:
            return {
                "cx": _pt_to_emu(m_w.group(1)),
                "cy": _pt_to_emu(m_h.group(1)),
            }

    return None


def _has_vml_textbox(node):
    return node.find(".//" + qn(V_NS, "textbox")) is not None


def _has_vml_imagedata(node):
    return node.find(".//" + qn(V_NS, "imagedata")) is not None


def _has_drawingml_picture(node):
    """
    True for DrawingML picture payload:
      w:drawing/wp:inline|anchor/a:graphic/a:graphicData/pic:pic
    """
    return node.find(".//" + qn(PIC_NS, "pic")) is not None


def _parse_vml_style_map(style: str):
    out = {}
    for part in (style or "").split(";"):
        if ":" not in part:
            continue
        key, value = part.split(":", 1)
        key = key.strip()
        value = value.strip()
        if key:
            out[key] = value
    return out


def _pt_style_to_emu(value):
    if not value:
        return None
    m = re.search(r"([0-9]+(?:\.[0-9]+)?)pt", str(value))
    if not m:
        return None
    try:
        return _pt_to_emu(m.group(1))
    except Exception:
        return None



def _int_or_none(value):
    if value is None:
        return None
    try:
        return int(str(value).strip())
    except Exception:
        return None


def _parse_vml_positioning(node):
    """
    Parse basic positioning / wrap from VML shape style.
    This is heuristic, but enough to distinguish inline-safe vs floating wrap.
    """
    v_node = (
        node.find(".//" + qn(V_NS, "rect")) or
        node.find(".//" + qn(V_NS, "roundrect")) or
        node.find(".//" + qn(V_NS, "oval")) or
        node.find(".//" + qn(V_NS, "line")) or
        node.find(".//" + qn(V_NS, "shape"))
    )
    if v_node is None:
        return None

    style_map = _parse_vml_style_map(v_node.get("style") or "")
    if not style_map and not _has_vml_textbox(node):
        return None

    out = {}

    position_mode = (style_map.get("position") or "").strip().lower()
    out["layout"] = "anchor" if position_mode == "absolute" else "inline"

    wrap_el = node.find(".//" + qn(W10_NS, "wrap"))
    wrap_type = None
    if wrap_el is not None:
        wrap_type = (wrap_el.get("type") or "").strip().lower()

        anchor_x = (wrap_el.get("anchorx") or "").strip()
        anchor_y = (wrap_el.get("anchory") or "").strip()
        if anchor_x and "relativeFromH" not in out:
            out["relativeFromH"] = anchor_x
        if anchor_y and "relativeFromV" not in out:
            out["relativeFromV"] = anchor_y

    if not wrap_type:
        wrap_type = (style_map.get("mso-wrap-style") or "").strip().lower()

    z_index = _int_or_none(style_map.get("z-index"))

    wrap_map = {
        "square": "square",
        "tight": "tight",
        "through": "through",
        "top-and-bottom": "topAndBottom",
        "top-bottom": "topAndBottom",
    }

    if wrap_type in wrap_map:
        out["wrap"] = wrap_map[wrap_type]
    elif wrap_type == "none":
        # Heuristic for VML:
        # no-wrap + negative z-index -> behind text
        # no-wrap + non-negative/unknown z-index -> in front of text
        if z_index is not None and z_index < 0:
            out["wrap"] = "behindText"
        else:
            out["wrap"] = "inFrontOfText"

    left_emu = _pt_style_to_emu(style_map.get("margin-left") or style_map.get("left"))
    top_emu = _pt_style_to_emu(style_map.get("margin-top") or style_map.get("top"))
    if left_emu is not None:
        out["horizontalPos"] = left_emu
    if top_emu is not None:
        out["verticalPos"] = top_emu

    rel_h = (style_map.get("mso-position-horizontal-relative") or "").strip()
    rel_v = (style_map.get("mso-position-vertical-relative") or "").strip()
    if rel_h:
        out["relativeFromH"] = rel_h
    if rel_v:
        out["relativeFromV"] = rel_v

    dist_l = _pt_style_to_emu(style_map.get("mso-wrap-distance-left"))
    dist_r = _pt_style_to_emu(style_map.get("mso-wrap-distance-right"))
    dist_t = _pt_style_to_emu(style_map.get("mso-wrap-distance-top"))
    dist_b = _pt_style_to_emu(style_map.get("mso-wrap-distance-bottom"))
    if dist_l is not None:
        out["distL"] = dist_l
    if dist_r is not None:
        out["distR"] = dist_r
    if dist_t is not None:
        out["distT"] = dist_t
    if dist_b is not None:
        out["distB"] = dist_b

    if z_index is not None:
        out["zIndex"] = z_index

    return out if out else None



def _emu_int(value):
    if value is None:
        return None
    try:
        return int(value)
    except Exception:
        return None


def _parse_shape_positioning(node):
    """
    Parse positioning/wrap info for DrawingML floating or inline shapes.
    Returns dict compatible with shape.positioning or None.
    """
    inline = node.find(".//" + qn(WP_NS, "inline"))
    anchor = node.find(".//" + qn(WP_NS, "anchor"))

    # VML branch
    if inline is None and anchor is None:
        return _parse_vml_positioning(node)

    container = anchor if anchor is not None else inline
    if container is None:
        return None


    out = {}
    out["layout"] = "anchor" if anchor is not None else "inline"

    if anchor is not None:
        simple_pos = anchor.find(qn(WP_NS, "simplePos"))
        if simple_pos is not None:
            x = _emu_int(simple_pos.get("x"))
            y = _emu_int(simple_pos.get("y"))
            if x is not None:
                out["horizontalPos"] = x
            if y is not None:
                out["verticalPos"] = y

        pos_h = anchor.find(qn(WP_NS, "positionH"))
        if pos_h is not None:
            rel = pos_h.get("relativeFrom")
            if rel:
                out["relativeFromH"] = rel
            pos_off = pos_h.find(qn(WP_NS, "posOffset"))
            if pos_off is not None and pos_off.text:
                val = _emu_int(pos_off.text)
                if val is not None:
                    out["horizontalPos"] = val

        pos_v = anchor.find(qn(WP_NS, "positionV"))
        if pos_v is not None:
            rel = pos_v.get("relativeFrom")
            if rel:
                out["relativeFromV"] = rel
            pos_off = pos_v.find(qn(WP_NS, "posOffset"))
            if pos_off is not None and pos_off.text:
                val = _emu_int(pos_off.text)
                if val is not None:
                    out["verticalPos"] = val

        wrap_map = (
            ("wrapNone", "none"),
            ("wrapSquare", "square"),
            ("wrapTight", "tight"),
            ("wrapThrough", "through"),
            ("wrapTopAndBottom", "topAndBottom"),
        )
        for tag_name, wrap_value in wrap_map:
            wrap_el = anchor.find(qn(WP_NS, tag_name))
            if wrap_el is not None:
                out["wrap"] = wrap_value
                break

        for attr_name in ("distL", "distR", "distT", "distB"):
            val = _emu_int(anchor.get(attr_name))
            if val is not None:
                out[attr_name] = val

        allow_overlap = anchor.get("allowOverlap")
        if allow_overlap is not None:
            out["allowOverlap"] = allow_overlap in ("1", "true", "on")

        behind_doc = anchor.get("behindDoc")
        if behind_doc is not None:
            out["behindDoc"] = behind_doc in ("1", "true", "on")
            if out["behindDoc"]:
                out["wrap"] = "behindText"

    return out if out else None


def _parse_extent(node):
    """
    Extract extent (cx, cy) from:
    - wp:extent
    - a:ext
    - VML style width/height
    """
    wp_extent = node.find(".//" + qn(WP_NS, "extent"))
    if wp_extent is not None:
        cx = wp_extent.get("cx")
        cy = wp_extent.get("cy")
        if cx and cy:
            return {"cx": int(cx), "cy": int(cy)}

    a_ext = node.find(".//" + qn(A_NS, "ext"))
    if a_ext is not None:
        cx = a_ext.get("cx")
        cy = a_ext.get("cy")
        if cx and cy:
            return {"cx": int(cx), "cy": int(cy)}

    vml_extent = _parse_vml_style_extent(node)
    if vml_extent is not None:
        return vml_extent

    return None


def _parse_shape_type(node):
    """
    Detect shape type from DrawingML or VML.
    """
    prst = node.find(".//" + qn(A_NS, "prstGeom"))
    if prst is not None:
        val = prst.get("prst")
        if val:
            return val

    for tag_name, shape_type in (
        ("rect", "rect"),
        ("roundrect", "roundRect"),
        ("oval", "ellipse"),
        ("line", "line"),
    ):
        if node.find(".//" + qn(V_NS, tag_name)) is not None:
            return shape_type

    v_shape = node.find(".//" + qn(V_NS, "shape"))
    if v_shape is not None:
        t = v_shape.get("type")
        if t:
            return t

    if _has_vml_textbox(node):
        return "textBox"

    return "unknown"


def _parse_textbox_content(node, parser, run_id):
    """
    Parse textbox paragraphs inside:
    - w:txbxContent (VML)
    - a:txBody (DrawingML)
    """
    content = []

    # VML textbox
    txbx = node.find(".//" + qn(W_NS, "txbxContent"))

    # DrawingML textbox
    if txbx is None:
        txbx = node.find(".//" + qn(A_NS, "txBody"))

    if txbx is None:
        return content

    p_index = 1

    paragraphs = txbx.findall(qn(W_NS, "p")) + txbx.findall(qn(A_NS, "p"))
    for p in paragraphs:
        paragraph = parser._parse_paragraph_element(
            p,
            parent_id=run_id,
            index=p_index
        )

        paragraph["id"] = f"{run_id}.p_{p_index}"
        content.append(paragraph)

        p_index += 1

    return content


def parse_shape_node(node, run_id, parser):
    """
    Parse shape from w:drawing or w:pict.
    Return None for picture-only payloads so parser_picture fallback can handle them.
    """
    # VML image without textbox -> not a shape, let picture parser handle it
    if _has_vml_imagedata(node) and not _has_vml_textbox(node):
        return None

    # DrawingML picture without textbox/content -> not a shape, let picture parser handle it
    # Important: DrawingML may represent both pictures and real shapes.
    # We only deflect pure picture payloads (pic:pic). Text boxes / actual shapes
    # will continue through normal shape parsing.
    if _has_drawingml_picture(node) and not _has_vml_textbox(node):
        return None

    shape_type = _parse_shape_type(node)
    extent = _parse_extent(node)

    shape = {
        "shapeType": shape_type
    }

    if extent:
        shape["extent"] = extent

    positioning = _parse_shape_positioning(node)
    if positioning:
        shape["positioning"] = positioning

    content = _parse_textbox_content(node, parser, run_id)

    if content:
        shape["content"] = content

    parent_id = run_id.rsplit(".run_", 1)[0]

    return {
        "type": "shape",
        "id": run_id,
        "parent_id": parent_id,
        "shape": shape
    }
