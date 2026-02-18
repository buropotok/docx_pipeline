import argparse
import os
import sys


def main() -> int:
    parser = argparse.ArgumentParser(description="Materialize DOCX via Word SaveAs2")
    parser.add_argument("--in", dest="in_docx", required=True, help="Input DOCX path")
    parser.add_argument("--out", dest="out_docx", required=True, help="Output materialized DOCX path")
    args = parser.parse_args()

    in_docx = os.path.abspath(args.in_docx)
    out_docx = os.path.abspath(args.out_docx)

    print(f"[word_saveas] input: {in_docx}")
    print(f"[word_saveas] output: {out_docx}")

    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        print(f"[word_saveas] pywin32 import failed: {exc}")
        return 2

    os.makedirs(os.path.dirname(out_docx) or ".", exist_ok=True)

    app = None
    doc = None
    try:
        print("[word_saveas] starting Word.Application")
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0

        print("[word_saveas] opening document (ReadOnly=True)")
        doc = app.Documents.Open(in_docx, ReadOnly=True)

        print("[word_saveas] saving as DOCX (FileFormat=16)")
        doc.SaveAs2(out_docx, FileFormat=16)

        if not os.path.exists(out_docx):
            print("[word_saveas] SaveAs2 finished but output file not found")
            return 3

        print("[word_saveas] done")
        return 0
    finally:
        if doc is not None:
            try:
                print("[word_saveas] closing document")
                doc.Close(False)
            except Exception as exc:
                print(f"[word_saveas] close doc warning: {exc}")
        if app is not None:
            try:
                print("[word_saveas] quitting Word")
                app.Quit()
            except Exception as exc:
                print(f"[word_saveas] quit warning: {exc}")


if __name__ == "__main__":
    sys.exit(main())
