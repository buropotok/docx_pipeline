import shutil
import subprocess
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path

from dagster import asset, AssetExecutionContext, Output, MetadataValue, Nothing


PROJECT_ROOT = Path(__file__).resolve().parents[1]            # ...\docx_pipeline
DATA_DIR = PROJECT_ROOT / "data"
RUNS_DIR = DATA_DIR / "runs"

DEFAULT_INPUT_DOCX = DATA_DIR / "donor.docx"


@dataclass(frozen=True)
class RunPaths:
    run_dir: Path
    input_dir: Path
    out_dir: Path
    raw_dir: Path

    input_docx: Path
    materialized_docx: Path
    raw_json: Path
    effective_json: Path
    optimized_json: Path
    reconstructed_docx: Path

    raw_donor: Path
    raw_materialized: Path
    raw_reconstructed: Path


def _mk_run_paths(context: AssetExecutionContext, input_docx: Path) -> RunPaths:
    run_dir = RUNS_DIR / context.run_id
    input_dir = run_dir / "input"
    out_dir = run_dir / "out"
    raw_dir = run_dir / "raw"

    input_dir.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)
    raw_dir.mkdir(parents=True, exist_ok=True)

    # Копируем вход в run_dir, чтобы run был самодостаточным
    run_input_docx = input_dir / input_docx.name
    if not run_input_docx.exists():
        shutil.copy2(input_docx, run_input_docx)

    stem = run_input_docx.stem
    materialized_docx = out_dir / f"{stem}.materialized.docx"
    raw_json = out_dir / f"{stem}.json"
    effective_json = out_dir / f"{stem}.effective.json"
    optimized_json = out_dir / f"{stem}.optimized.json"  # новая строка
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
        effective_json=effective_json,
        optimized_json=optimized_json,  # добавлено
        reconstructed_docx=reconstructed_docx,
        raw_donor=raw_donor,
        raw_materialized=raw_materialized,
        raw_reconstructed=raw_reconstructed,
    )


def _run_cmd(context: AssetExecutionContext, cmd: list[str], step: str) -> None:
    context.log.info(f"[{step}] cmd={' '.join(cmd)}")
    proc = subprocess.run(
        cmd,
        cwd=str(PROJECT_ROOT),
        text=True,
        capture_output=True,
    )
    if proc.stdout:
        for line in proc.stdout.splitlines():
            context.log.info(f"[{step}] stdout: {line}")
    if proc.stderr:
        for line in proc.stderr.splitlines():
            context.log.warning(f"[{step}] stderr: {line}")

    context.log.info(f"[{step}] exit_code={proc.returncode}")
    if proc.returncode != 0:
        raise subprocess.CalledProcessError(proc.returncode, cmd, output=proc.stdout, stderr=proc.stderr)


def _extract_docx_raw(context: AssetExecutionContext, docx_path: Path, out_dir: Path, step: str) -> None:
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(docx_path, "r") as z:
        z.extractall(out_dir)

    files_count = sum(1 for p in out_dir.rglob("*") if p.is_file())
    key_styles = (out_dir / "word" / "styles.xml").exists()
    key_numbering = (out_dir / "word" / "numbering.xml").exists()
    key_settings = (out_dir / "word" / "settings.xml").exists()

    context.log.info(f"[{step}] extracted_from={docx_path} to={out_dir} files_count={files_count}")
    context.log.info(f"[{step}] key_xml styles={key_styles} numbering={key_numbering} settings={key_settings}")


# --- Assets ---

@asset
def input_docx_path() -> str:
    # Пока хардкодим donor, позже сделаем config / input from email
    if not DEFAULT_INPUT_DOCX.exists():
        raise FileNotFoundError(str(DEFAULT_INPUT_DOCX))
    return str(DEFAULT_INPUT_DOCX)



@asset
def saveas_materialized(context: AssetExecutionContext, input_docx_path: str) -> Output[Nothing]:
    p = _mk_run_paths(context, Path(input_docx_path))
    py = sys.executable

    _run_cmd(context, [py, "-m", "tools.word_saveas", "--in", str(p.input_docx), "--out", str(p.materialized_docx)], "saveas")
    _extract_docx_raw(context, p.input_docx, p.raw_donor, "raw_donor")
    _extract_docx_raw(context, p.materialized_docx, p.raw_materialized, "raw_materialized")

    return Output(
        value=None,
        metadata={
            "run_dir": MetadataValue.path(str(p.run_dir)),
            "input_docx": MetadataValue.path(str(p.input_docx)),
            "materialized_docx": MetadataValue.path(str(p.materialized_docx)),
        },
    )


@asset(non_argument_deps={"saveas_materialized"})
def parse_raw_json(context: AssetExecutionContext, input_docx_path: str) -> Output[Nothing]:
    p = _mk_run_paths(context, Path(input_docx_path))
    py = sys.executable

    _run_cmd(context, [py, "src/parser.py", "--in", str(p.materialized_docx), "--out", str(p.raw_json)], "parser")

    return Output(
        value=None,
        metadata={"raw_json": MetadataValue.path(str(p.raw_json))},
    )


@asset(non_argument_deps={"parse_raw_json"})
def materialize_effective(context: AssetExecutionContext, input_docx_path: str) -> Output[Nothing]:
    p = _mk_run_paths(context, Path(input_docx_path))
    py = sys.executable

    _run_cmd(
        context,
        [py, "tools/materialize_effective.py", "--docx", str(p.materialized_docx), "--in-json", str(p.raw_json), "--out-json", str(p.effective_json)],
        "effective",
    )

    return Output(
        value=None,
        metadata={"effective_json": MetadataValue.path(str(p.effective_json))},
    )

@asset(non_argument_deps={"materialize_effective"})
def optimize_tabs(context: AssetExecutionContext, input_docx_path: str) -> Output[Nothing]:
    p = _mk_run_paths(context, Path(input_docx_path))
    py = sys.executable

    # Запускаем скрипт оптимизации, который использует Word COM для точного измерения
    _run_cmd(
        context,
        [py, "tools/optimization/optimize_tabs.py", "--in-json", str(p.effective_json), "--out-json",
         str(p.optimized_json)],
        "optimize_tabs",
    )

    return Output(
        value=None,
        metadata={"optimized_json": MetadataValue.path(str(p.optimized_json))},
    )

@asset(non_argument_deps={"optimize_tabs"})
def reconstruct_docx_opt(context: AssetExecutionContext, input_docx_path: str) -> Output[Nothing]:
    p = _mk_run_paths(context, Path(input_docx_path))
    py = sys.executable

    _run_cmd(context, [py, "src/reconstructor.py", "--in-json", str(p.optimized_json), "--out-docx", str(p.reconstructed_docx)], "recon")
    _extract_docx_raw(context, p.reconstructed_docx, p.raw_reconstructed, "raw_reconstructed")

    return Output(
        value=None,
        metadata={
            "reconstructed_docx": MetadataValue.path(str(p.reconstructed_docx)),
            "raw_reconstructed_dir": MetadataValue.path(str(p.raw_reconstructed)),
        },
    )

@asset(non_argument_deps={"parse_raw_json"})
def reconstruct_docx(context: AssetExecutionContext, input_docx_path: str) -> Output[Nothing]:
    p = _mk_run_paths(context, Path(input_docx_path))
    py = sys.executable

    _run_cmd(context, [py, "src/reconstructor.py", "--in-json", str(p.raw_json), "--out-docx", str(p.reconstructed_docx)], "recon")
    _extract_docx_raw(context, p.reconstructed_docx, p.raw_reconstructed, "raw_reconstructed")

    return Output(
        value=None,
        metadata={
            "reconstructed_docx": MetadataValue.path(str(p.reconstructed_docx)),
            "raw_reconstructed_dir": MetadataValue.path(str(p.raw_reconstructed)),
        },
    )

@asset
def full_run_pipeline(context: AssetExecutionContext) -> Output[Nothing]:
    """
    One-step Dagster asset: runs saveas -> add_custom_attrs -> parser -> reconstructor
    and writes artifacts into data/runs/<run_id>/... (same structure as before).
    """
    py = sys.executable

    cmd = [
        py,
        "tools/full_run_utility.py",
        "--run-id",
        context.run_id,
        # "--input-docx", str(DEFAULT_INPUT_DOCX),  # можно явно, но утилита и так берёт data/donor.docx по умолчанию
    ]

    context.log.info(f"[full_run_pipeline] cmd={' '.join(cmd)}")
    proc = subprocess.run(
        cmd,
        cwd=str(PROJECT_ROOT),
        text=True,
        capture_output=True,
    )
    if proc.stdout:
        for line in proc.stdout.splitlines():
            context.log.info(f"[full_run_pipeline] stdout: {line}")
    if proc.stderr:
        for line in proc.stderr.splitlines():
            context.log.warning(f"[full_run_pipeline] stderr: {line}")

    context.log.info(f"[full_run_pipeline] exit_code={proc.returncode}")
    if proc.returncode != 0:
        raise subprocess.CalledProcessError(proc.returncode, cmd, output=proc.stdout, stderr=proc.stderr)

    run_dir = RUNS_DIR / context.run_id
    return Output(
        value=None,
        metadata={
            "run_dir": MetadataValue.path(str(run_dir)),
            "out_dir": MetadataValue.path(str(run_dir / "out")),
            "raw_dir": MetadataValue.path(str(run_dir / "raw")),
        },
    )


@asset
def full_run_pipeline_gpt(context: AssetExecutionContext) -> Output[Nothing]:
    """
    One-step Dagster asset: runs saveas -> add_custom_attrs -> parser -> reconstructor
    and writes artifacts into data/runs/<run_id>/... (same structure as before).
    """
    py = sys.executable

    cmd = [
        py,
        "tools/full_run_utility_gpt.py",
        "--run-id",
        context.run_id,
        # "--input-docx", str(DEFAULT_INPUT_DOCX),  # можно явно, но утилита и так берёт data/donor.docx по умолчанию
    ]

    context.log.info(f"[full_run_pipeline_gpt] cmd={' '.join(cmd)}")
    proc = subprocess.run(
        cmd,
        cwd=str(PROJECT_ROOT),
        text=True,
        capture_output=True,
    )
    if proc.stdout:
        for line in proc.stdout.splitlines():
            context.log.info(f"[full_run_pipeline_gpt] stdout: {line}")
    if proc.stderr:
        for line in proc.stderr.splitlines():
            context.log.warning(f"[_gpt] stderr: {line}")

    context.log.info(f"[full_run_pipeline_gpt] exit_code={proc.returncode}")
    if proc.returncode != 0:
        raise subprocess.CalledProcessError(proc.returncode, cmd, output=proc.stdout, stderr=proc.stderr)

    run_dir = RUNS_DIR / context.run_id
    return Output(
        value=None,
        metadata={
            "run_dir": MetadataValue.path(str(run_dir)),
            "out_dir": MetadataValue.path(str(run_dir / "out")),
            "raw_dir": MetadataValue.path(str(run_dir / "raw")),
        },
    )