import argparse
import json
import os
import sys
from typing import Any, Dict, Optional


WD_LINE_SPACE_SINGLE = 0
WD_LINE_SPACE_1PT5 = 1
WD_LINE_SPACE_DOUBLE = 2
WD_LINE_SPACE_AT_LEAST = 3
WD_LINE_SPACE_EXACTLY = 4
WD_LINE_SPACE_MULTIPLE = 5


def _points_to_twips(points: Any) -> Optional[int]:
    try:
        return round(float(points) * 20)
    except Exception:
        return None


def _extract_line_info(line_rule: Any, line_spacing: Any) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    try:
        rule = int(line_rule)
    except Exception:
        return out

    if rule == WD_LINE_SPACE_SINGLE:
        out["lineRule"] = "AUTO"
        out["lineTwip"] = 240
        return out

    if rule in (WD_LINE_SPACE_1PT5, WD_LINE_SPACE_DOUBLE):
        try:
            multiple = float(line_spacing) / 12.0
            out["lineRule"] = "AUTO"
            out["lineTwip"] = round(240 * multiple)
        except Exception:
            pass
        return out

    if rule == WD_LINE_SPACE_MULTIPLE:
        try:
            multiple = float(line_spacing) / 12.0
            out["lineRule"] = "AUTO"
            out["lineTwip"] = round(240 * multiple)
        except Exception:
            pass
        return out

    if rule == WD_LINE_SPACE_EXACTLY:
        tw = _points_to_twips(line_spacing)
        if tw is not None:
            out["lineRule"] = "EXACT"
            out["lineTwip"] = tw
        return out

    if rule == WD_LINE_SPACE_AT_LEAST:
        tw = _points_to_twips(line_spacing)
        if tw is not None:
            out["lineRule"] = "AT_LEAST"
            out["lineTwip"] = tw
        return out

    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="Fill missing effective spacing/line fields from Word COM Normal style")
    ap.add_argument("--docx", required=True)
    ap.add_argument("--in-json", required=True, dest="in_json")
    ap.add_argument("--out-json", required=True, dest="out_json")
    args = ap.parse_args()

    in_json = os.path.abspath(args.in_json)
    out_json = os.path.abspath(args.out_json)
    docx_path = os.path.abspath(args.docx)

    with open(in_json, "r", encoding="utf-8") as f:
        data = json.load(f)

    meta = data.get("meta") or {}
    styles = data.get("styles") or {}
    default_style_id = meta.get("default_style_id")

    if not isinstance(default_style_id, str) or default_style_id not in styles:
        print("[materialize_effective] warning: meta.default_style_id missing/invalid, writing copy")
        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return 0

    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        print(f"[materialize_effective] pywin32 import failed: {exc}")
        return 2

    app = None
    doc = None
    extracted: Dict[str, Any] = {}

    try:
        print(f"[materialize_effective] opening docx: {docx_path}")
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        doc = app.Documents.Open(docx_path, ReadOnly=True)

        style = None
        try:
            style = doc.Styles("Normal")
        except Exception:
            try:
                wd_style_normal = getattr(win32com.client.constants, "wdStyleNormal", -1)
                if wd_style_normal != -1:
                    style = doc.Styles(wd_style_normal)
            except Exception:
                style = None

        if style is None:
            print("[materialize_effective] warning: cannot resolve Normal style, writing copy")
            with open(out_json, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return 0

        pf = style.ParagraphFormat

        try:
            sb = pf.SpaceBefore
            tw = _points_to_twips(sb)
            if tw is not None:
                extracted["spaceBeforeTwip"] = tw
            print(f"[materialize_effective] SpaceBefore={sb}pt -> {tw}")
        except Exception as exc:
            print(f"[materialize_effective] SpaceBefore unavailable: {exc}")

        try:
            sa = pf.SpaceAfter
            tw = _points_to_twips(sa)
            if tw is not None:
                extracted["spaceAfterTwip"] = tw
            print(f"[materialize_effective] SpaceAfter={sa}pt -> {tw}")
        except Exception as exc:
            print(f"[materialize_effective] SpaceAfter unavailable: {exc}")

        line_rule = None
        line_spacing = None
        try:
            line_rule = pf.LineSpacingRule
            print(f"[materialize_effective] LineSpacingRule={line_rule}")
        except Exception as exc:
            print(f"[materialize_effective] LineSpacingRule unavailable: {exc}")

        try:
            line_spacing = pf.LineSpacing
            print(f"[materialize_effective] LineSpacing={line_spacing}")
        except Exception as exc:
            print(f"[materialize_effective] LineSpacing unavailable: {exc}")

        line_vals = _extract_line_info(line_rule, line_spacing)
        if line_vals:
            extracted.update(line_vals)
            print(f"[materialize_effective] line fields extracted: {line_vals}")
        else:
            print("[materialize_effective] line fields not extracted")
    finally:
        if doc is not None:
            try:
                print("[materialize_effective] closing doc")
                doc.Close(False)
            except Exception as exc:
                print(f"[materialize_effective] close warning: {exc}")
        if app is not None:
            try:
                print("[materialize_effective] quitting Word")
                app.Quit()
            except Exception as exc:
                print(f"[materialize_effective] quit warning: {exc}")

    p_format = styles[default_style_id].setdefault("p_format", {})
    filled = []
    for key in ("spaceBeforeTwip", "spaceAfterTwip", "lineTwip", "lineRule"):
        if key in extracted and key not in p_format:
            p_format[key] = extracted[key]
            filled.append(key)

    print(f"[materialize_effective] filled fields: {filled}")

    os.makedirs(os.path.dirname(out_json) or ".", exist_ok=True)
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return 0


if __name__ == "__main__":
    sys.exit(main())
