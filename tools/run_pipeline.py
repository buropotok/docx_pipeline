import shutil
import zipfile
from pathlib import Path
import subprocess
import sys


WORK_DIR = Path(__file__).resolve().parent.parent / "data"



def run_cmd(cmd):
    print("[run_pipeline]", " ".join(str(x) for x in cmd))
    subprocess.run(cmd, check=True)


def extract_docx_raw(docx_path: Path, out_dir: Path) -> None:
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(docx_path, "r") as z:
        z.extractall(out_dir)
    print(f"[run_pipeline] raw extracted: {docx_path} -> {out_dir}")


def resolve_input(arg: str) -> Path:
    p = Path(arg)
    if p.is_absolute():
        return p
    return WORK_DIR / p.name


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: python tools/run_pipeline.py <donor.docx | /data/donor.docx>")
        return 2

    WORK_DIR.mkdir(parents=True, exist_ok=True)

    input_docx = resolve_input(sys.argv[1])
    if not input_docx.exists():
        print(f"[run_pipeline] input DOCX not found: {input_docx}")
        print("[run_pipeline] Put donor file in /data or pass absolute /data/<name>.docx path")
        return 2

    stem = input_docx.stem
    materialized_docx = WORK_DIR / f"{stem}.materialized.docx"
    raw_json = WORK_DIR / f"{stem}.json"
    effective_json = WORK_DIR / f"{stem}.effective.json"
    reconstructed_docx = WORK_DIR / f"{stem}.reconstructed.docx"

    raw_root = WORK_DIR / "raw"
    raw_donor = raw_root / "donor"
    raw_materialized = raw_root / "materialized"
    raw_reconstructed = raw_root / "reconstructed"

    print(f"[run_pipeline] work dir: {WORK_DIR}")
    print(f"[run_pipeline] input: {input_docx}")
    print(f"[run_pipeline] materialized: {materialized_docx}")
    print(f"[run_pipeline] raw json: {raw_json}")
    print(f"[run_pipeline] effective json: {effective_json}")
    print(f"[run_pipeline] reconstructed: {reconstructed_docx}")
    print(f"[run_pipeline] raw dirs: {raw_donor}, {raw_materialized}, {raw_reconstructed}")

    py = sys.executable

    try:
        run_cmd([py, "tools/word_saveas.py", "--in", str(input_docx), "--out", str(materialized_docx)])
        extract_docx_raw(input_docx, raw_donor)
        extract_docx_raw(materialized_docx, raw_materialized)

        run_cmd([py, "src/parser.py", "--in", str(materialized_docx), "--out", str(raw_json)])
        run_cmd([py, "tools/materialize_effective.py", "--docx", str(materialized_docx), "--in-json", str(raw_json), "--out-json", str(effective_json)])
        run_cmd([py, "src/reconstructor.py", "--in-json", str(effective_json), "--out-docx", str(reconstructed_docx)])

        extract_docx_raw(reconstructed_docx, raw_reconstructed)
    except subprocess.CalledProcessError as exc:
        print(f"[run_pipeline] step failed with code {exc.returncode}: {exc.cmd}")
        return exc.returncode or 1

    print("[run_pipeline] done")
    return 0


if __name__ == "__main__":
    sys.exit(main())
