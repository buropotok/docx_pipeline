import argparse
import json
import os
import sys
import traceback
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
        print("[effective] warning: meta.default_style_id missing/invalid, writing copy")
        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return 0

    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        print(f"[effective] pywin32 import failed: {exc}")
        return 2

    app = None
    doc = None
    extracted: Dict[str, Any] = {}

    try:
        print(f"[effective] opening docx: {docx_path}")
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        doc = app.Documents.Open(docx_path, ReadOnly=True)

        style = None
        try:
            style = doc.Styles("Normal")
            print('[materialize_effective] resolve_normal: method=name("Normal") ok')
        except Exception as e1:
            print(f'[materialize_effective] resolve_normal: method=name("Normal") failed: {e1!r}')
            wd_style_normal = -1
            source = "fallback:-1"
            try:
                wd_style_normal = getattr(win32com.client.constants, "wdStyleNormal")
                source = "constants.wdStyleNormal"
            except Exception as e_const:
                print(f'[materialize_effective] resolve_normal: constants.wdStyleNormal unavailable: {e_const!r}; using -1')

            try:
                style = doc.Styles(wd_style_normal)
                print(f'[materialize_effective] resolve_normal: method={source}({wd_style_normal}) ok')
            except Exception as e2:
                print(f'[materialize_effective] resolve_normal: method=wdStyleNormal({wd_style_normal}) failed: {e2!r}')
                style = None

        if style is None:
            print("[effective] warning: cannot resolve Normal style, writing copy")
            with open(out_json, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return 0

        try:
            print(f"[effective] resolve_normal: style_name={style.Name} style_name_local={style.NameLocal}")
        except Exception:
            print("[effective] resolve_normal: style_name unavailable")

        pf = style.ParagraphFormat

        try:
            sb = pf.SpaceBefore
            tw = _points_to_twips(sb)
            if tw is not None:
                extracted["spaceBeforeTwip"] = tw
            print(f"[effective] SpaceBefore={sb}pt -> {tw}")
        except Exception as exc:
            print(f"[effective] SpaceBefore unavailable: {exc}")

        try:
            sa = pf.SpaceAfter
            tw = _points_to_twips(sa)
            if tw is not None:
                extracted["spaceAfterTwip"] = tw
            print(f"[effective] SpaceAfter={sa}pt -> {tw}")
        except Exception as exc:
            print(f"[effective] SpaceAfter unavailable: {exc}")

        line_rule = None
        line_spacing = None
        try:
            line_rule = pf.LineSpacingRule
            print(f"[effective] LineSpacingRule={line_rule}")
        except Exception as exc:
            print(f"[effective] LineSpacingRule unavailable: {exc}")

        try:
            line_spacing = pf.LineSpacing
            print(f"[effective] LineSpacing={line_spacing}")
        except Exception as exc:
            print(f"[effective] LineSpacing unavailable: {exc}")

        line_vals = _extract_line_info(line_rule, line_spacing)
        if line_vals:
            extracted.update(line_vals)
            print(f"[effective] line fields extracted: {line_vals}")
        else:
            print("[effective] line fields not extracted")
    except Exception as exc:
        print(f"[effective] ERROR type={type(exc).__name__} message={exc}")
        print(traceback.format_exc())
        return 1
    finally:
        if doc is not None:
            try:
                print("[effective] closing doc")
                doc.Close(False)
            except Exception as exc:
                print(f"[effective] close warning: {exc}")
        if app is not None:
            try:
                print("[effective] quitting Word")
                app.Quit()
            except Exception as exc:
                print(f"[effective] quit warning: {exc}")

    p_format = styles[default_style_id].setdefault("p_format", {})
    target_fields = ["spaceBeforeTwip", "spaceAfterTwip", "lineTwip", "lineRule"]
    missing_before = sorted([k for k in target_fields if k not in p_format])
    skipped_existing = sorted([k for k in target_fields if k in p_format])
    filled_fields = []

    print(f"[effective] target_style_id={default_style_id}")
    print(f"[effective] missing_fields_before={missing_before}")

    for key in sorted(target_fields):
        if key in extracted and key not in p_format:
            p_format[key] = extracted[key]
            filled_fields.append(f"{key}: old=None -> new={extracted[key]}")

    print(f"[effective] filled_fields={filled_fields}")
    print(f"[effective] skipped_existing_fields={skipped_existing}")
    print(f"[effective] enriched={'yes' if filled_fields else 'no'} filled_count={len(filled_fields)} skipped_count={len(skipped_existing)}")

    os.makedirs(os.path.dirname(out_json) or ".", exist_ok=True)
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return 0


if __name__ == "__main__":
    sys.exit(main())
