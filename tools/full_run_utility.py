# tools/full_run_utility.py
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]  # .../docx_pipeline
DATA_DIR = PROJECT_ROOT / "data"
RUNS_DIR = DATA_DIR / "runs"


@dataclass(frozen=True)
class RunPaths:
    run_dir: Path
    input_dir: Path
    out_dir: Path
    raw_dir: Path

    input_docx: Path
    materialized_docx: Path
    raw_json: Path
    reconstructed_docx: Path

    raw_donor: Path
    raw_materialized: Path
    raw_reconstructed: Path


def _mk_run_paths(run_id: str, input_docx: Path) -> RunPaths:
    run_dir = RUNS_DIR / run_id
    input_dir = run_dir / "input"
    out_dir = run_dir / "out"
    raw_dir = run_dir / "raw"

    input_dir.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)
    raw_dir.mkdir(parents=True, exist_ok=True)

    run_input_docx = input_dir / input_docx.name
    if not run_input_docx.exists():
        shutil.copy2(input_docx, run_input_docx)

    stem = run_input_docx.stem
    materialized_docx = out_dir / f"{stem}.materialized.docx"
    raw_json = out_dir / f"{stem}.json"
    reconstructed_docx = out_dir / f"{stem}.reconstructed.docx"

    raw_donor = raw_dir / "donor"
    raw_materialized = raw_dir / "materialized"
    raw_reconstructed = raw_dir / "reconstructed"

    return RunPaths(
        run_dir=run_dir,
        input_dir=input_dir,
        out_dir=out_dir,
        raw_dir=raw_dir,
        input_docx=run_input_docx,
        materialized_docx=materialized_docx,
        raw_json=raw_json,
        reconstructed_docx=reconstructed_docx,
        raw_donor=raw_donor,
        raw_materialized=raw_materialized,
        raw_reconstructed=raw_reconstructed,
    )


def _run_cmd(cmd: list[str], step: str) -> None:
    print(f"[{step}] cmd={' '.join(cmd)}")
    proc = subprocess.run(
        cmd,
        cwd=str(PROJECT_ROOT),
        text=True,
        capture_output=True,
    )
    if proc.stdout:
        for line in proc.stdout.splitlines():
            print(f"[{step}] stdout: {line}")
    if proc.stderr:
        for line in proc.stderr.splitlines():
            print(f"[{step}] stderr: {line}", file=sys.stderr)

    print(f"[{step}] exit_code={proc.returncode}")
    if proc.returncode != 0:
        raise subprocess.CalledProcessError(proc.returncode, cmd, output=proc.stdout, stderr=proc.stderr)


def _extract_docx_raw(docx_path: Path, out_dir: Path, step: str) -> None:
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(docx_path, "r") as z:
        z.extractall(out_dir)

    files_count = sum(1 for p in out_dir.rglob("*") if p.is_file())
    key_styles = (out_dir / "word" / "styles.xml").exists()
    key_numbering = (out_dir / "word" / "numbering.xml").exists()
    key_settings = (out_dir / "word" / "settings.xml").exists()
    key_doc = (out_dir / "word" / "document.xml").exists()
    key_rels = (out_dir / "word" / "_rels" / "document.xml.rels").exists()

    print(f"[{step}] extracted_from={docx_path} to={out_dir} files_count={files_count}")
    print(f"[{step}] key_xml document={key_doc} rels={key_rels} styles={key_styles} numbering={key_numbering} settings={key_settings}")


def main() -> None:
    cli = argparse.ArgumentParser(description="Monolithic DOCX pipeline runner (Dagster-friendly)")
    cli.add_argument("--run-id", required=True, help="Dagster run_id (used as data/runs/<run_id>)")
    cli.add_argument("--input-docx", required=False, help="Path to input donor.docx (default: data/donor.docx)")
    args = cli.parse_args()

    input_docx = Path(args.input_docx) if args.input_docx else (DATA_DIR / "donor.docx")
    if not input_docx.exists():
        raise FileNotFoundError(str(input_docx))

    p = _mk_run_paths(args.run_id, input_docx)
    py = sys.executable

    # 1) saveas -> out/<stem>.materialized.docx
    _run_cmd([py, "-m", "tools.word_saveas", "--in", str(p.input_docx), "--out", str(p.materialized_docx)], "saveas")

    # keep "before attrs" artifact
    before_attrs = p.materialized_docx.with_name(p.materialized_docx.name.replace(".materialized.docx", ".materialized.before_attrs.docx"))
    shutil.copy2(p.materialized_docx, before_attrs)

    # 2) add_custom_attrs: overwrite materialized_docx so downstream paths remain identical
    tmp_tagged = p.materialized_docx.with_suffix(".tmp.tagged.docx")
    _run_cmd([py, "tools/add_custom_attrs.py", "--in", str(p.materialized_docx), "--out", str(tmp_tagged)], "add_custom_attrs")
    tmp_tagged.replace(p.materialized_docx)

    # 3) raw extractions (donor + materialized-with-myid)
    _extract_docx_raw(p.input_docx, p.raw_donor, "raw_donor")
    _extract_docx_raw(p.materialized_docx, p.raw_materialized, "raw_materialized")

    # 4) parser -> out/<stem>.json
    _run_cmd([py, "src/parser.py", "--in", str(p.materialized_docx), "--out", str(p.raw_json)], "parser")

    # 5) reconstructor -> out/<stem>.reconstructed.docx
    _run_cmd([py, "src/reconstructor.py", "--in-json", str(p.raw_json), "--out-docx", str(p.reconstructed_docx),
              "--donor-docx", str(p.materialized_docx)], "reconstructor")

    # 6) raw reconstruction extraction
    _extract_docx_raw(p.reconstructed_docx, p.raw_reconstructed, "raw_reconstructed")

    print(f"[done] run_dir={p.run_dir}")


if __name__ == "__main__":
    main()