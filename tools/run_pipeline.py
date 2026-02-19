import shutil
import subprocess
import sys
import time
import zipfile
from pathlib import Path

from tools.logging_utils import setup_run_logger


WORK_DIR = Path("/data")


def parse_versions_from_version_md() -> dict:
    out = {"Schema": "unknown", "Rules": "unknown", "Parser": "unknown", "Reconstructor": "unknown"}
    vpath = Path("VERSION.md")
    if not vpath.exists():
        return out
    for line in vpath.read_text(encoding="utf-8").splitlines():
        if ":" not in line:
            continue
        k, v = [x.strip() for x in line.split(":", 1)]
        if k in out:
            out[k] = v
    return out


def run_cmd(cmd, logger, step: str):
    logger.info(f"[{step}] cmd={' '.join(str(x) for x in cmd)}")
    started = time.time()
    proc = subprocess.run(cmd, capture_output=True, text=True)
    duration = time.time() - started

    if proc.stdout:
        for line in proc.stdout.splitlines():
            logger.info(f"[{step}] stdout: {line}")
    if proc.stderr:
        for line in proc.stderr.splitlines():
            logger.warning(f"[{step}] stderr: {line}")

    logger.info(f"[{step}] exit_code={proc.returncode} duration_s={duration:.3f}")
    if proc.returncode != 0:
        raise subprocess.CalledProcessError(proc.returncode, cmd, output=proc.stdout, stderr=proc.stderr)


def extract_docx_raw(docx_path: Path, out_dir: Path, logger, step: str) -> None:
    if out_dir.exists():
        logger.info(f"[{step}] clear_dir={out_dir}")
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(docx_path, "r") as z:
        z.extractall(out_dir)

    files = sorted([p for p in out_dir.rglob("*") if p.is_file()])
    key_styles = (out_dir / "word" / "styles.xml").exists()
    key_numbering = (out_dir / "word" / "numbering.xml").exists()
    key_settings = (out_dir / "word" / "settings.xml").exists()

    logger.info(f"[{step}] extracted_from={docx_path} to={out_dir} files_count={len(files)}")
    logger.info(f"[{step}] key_xml styles={key_styles} numbering={key_numbering} settings={key_settings}")


def resolve_input(arg: str) -> Path:
    p = Path(arg)
    if p.is_absolute():
        return p
    return WORK_DIR / p.name


def file_size_if_exists(path: Path) -> int:
    if path.exists():
        return path.stat().st_size
    return -1


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: python tools/run_pipeline.py <donor.docx | /data/donor.docx>")
        return 2

    WORK_DIR.mkdir(parents=True, exist_ok=True)
    logger, log_path = setup_run_logger(WORK_DIR)
    logger.name = "run"

    run_started = time.time()
    logger.info("[run] start")

    versions = parse_versions_from_version_md()
    logger.info(
        "[run] versions schema=%s rules=%s parser=%s reconstructor=%s",
        versions["Schema"], versions["Rules"], versions["Parser"], versions["Reconstructor"]
    )

    input_docx = resolve_input(sys.argv[1])
    if not input_docx.exists():
        logger.error(f"[run] input DOCX not found: {input_docx}")
        logger.error("[run] Put donor file in /data or pass absolute /data/<name>.docx path")
        logger.info(f"[run] log_file={log_path}")
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

    logger.info(f"[run] work_dir={WORK_DIR}")
    logger.info(f"[run] input={input_docx}")
    logger.info(f"[run] outputs materialized={materialized_docx} raw_json={raw_json} effective_json={effective_json} reconstructed={reconstructed_docx}")
    logger.info(f"[run] raw_dirs donor={raw_donor} materialized={raw_materialized} reconstructed={raw_reconstructed}")

    py = sys.executable

    try:
        run_cmd([py, "tools/word_saveas.py", "--in", str(input_docx), "--out", str(materialized_docx)], logger, "saveas")
        extract_docx_raw(input_docx, raw_donor, logger, "raw")
        extract_docx_raw(materialized_docx, raw_materialized, logger, "raw")

        run_cmd([py, "src/parser.py", "--in", str(materialized_docx), "--out", str(raw_json)], logger, "parser")
        run_cmd([py, "tools/materialize_effective.py", "--docx", str(materialized_docx), "--in-json", str(raw_json), "--out-json", str(effective_json)], logger, "effective")
        run_cmd([py, "src/reconstructor.py", "--in-json", str(effective_json), "--out-docx", str(reconstructed_docx)], logger, "recon")

        extract_docx_raw(reconstructed_docx, raw_reconstructed, logger, "raw")
    except subprocess.CalledProcessError as exc:
        logger.error(f"[run] step failed: cmd={' '.join(str(x) for x in exc.cmd)} exit_code={exc.returncode}")
        if exc.stderr:
            logger.error(f"[run] stderr={exc.stderr.strip()}")
        if exc.output:
            logger.error(f"[run] stdout={exc.output.strip()}")
        logger.info(f"[run] log_file={log_path}")
        return exc.returncode or 1

    logger.info(
        "[run] file_sizes donor=%s materialized=%s reconstructed=%s",
        file_size_if_exists(input_docx), file_size_if_exists(materialized_docx), file_size_if_exists(reconstructed_docx)
    )
    logger.info(f"[run] done duration_s={time.time() - run_started:.3f}")
    logger.info(f"[run] log_file={log_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
