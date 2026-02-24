import argparse
import os
import sys
import traceback


def main() -> int:
    parser = argparse.ArgumentParser(description="Materialize DOCX via Word SaveAs2")
    parser.add_argument("--in", dest="in_docx", required=True, help="Input DOCX path")
    parser.add_argument("--out", dest="out_docx", required=True, help="Output materialized DOCX path")
    args = parser.parse_args()

    in_docx = os.path.abspath(args.in_docx)
    out_docx = os.path.abspath(args.out_docx)

    print(f"[saveas] input: {in_docx}")
    print(f"[saveas] output: {out_docx}")

    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        print(f"[saveas] pywin32 import failed: {exc}")
        return 2

    os.makedirs(os.path.dirname(out_docx) or ".", exist_ok=True)

    app = None
    doc = None
    try:
        print("[saveas] starting Word.Application")
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        try:
            print(f"[saveas] word_version={app.Version}")
        except Exception:
            print("[saveas] word_version=unavailable")
        print(f"[saveas] visible={app.Visible} display_alerts={app.DisplayAlerts}")

        print(f"[saveas] opening document path={in_docx} ReadOnly=True")
        doc = app.Documents.Open(in_docx, ReadOnly=True)

        print(f"[saveas] saveas path={out_docx} file_format=16")
        doc.SaveAs2(out_docx, FileFormat=16)

        if not os.path.exists(out_docx):
            print("[saveas] SaveAs2 finished but output file not found")
            return 3

        print("[saveas] done")
        return 0
    except Exception as exc:
        hresult = getattr(exc, "hresult", None)
        print(f"[saveas] ERROR type={type(exc).__name__} hresult={hresult} message={exc}")
        print(traceback.format_exc())
        return 1
    finally:
        if doc is not None:
            try:
                print("[saveas] closing document")
                doc.Close(False)
            except Exception as exc:
                print(f"[saveas] close doc warning: {exc}")
        if app is not None:
            try:
                print("[saveas] quitting Word")
                app.Quit()
            except Exception as exc:
                print(f"[saveas] quit warning: {exc}")


if __name__ == "__main__":
    sys.exit(main())
