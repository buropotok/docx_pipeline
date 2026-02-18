import os
import subprocess
import sys


WORK_DIR = "."


def run_cmd(cmd):
    print("[run_pipeline]", " ".join(cmd))
    subprocess.run(cmd, check=True)


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: python tools/run_pipeline.py <donor.docx>")
        return 2

    filename = sys.argv[1]
    stem, _ = os.path.splitext(filename)

    input_docx = os.path.join(WORK_DIR, filename)
    materialized_docx = os.path.join(WORK_DIR, f"{stem}.materialized.docx")
    raw_json = os.path.join(WORK_DIR, f"{stem}.json")
    effective_json = os.path.join(WORK_DIR, f"{stem}.effective.json")
    reconstructed_docx = os.path.join(WORK_DIR, f"{stem}.reconstructed.docx")

    py = sys.executable

    try:
        run_cmd([py, "tools/word_saveas.py", "--in", input_docx, "--out", materialized_docx])
        run_cmd([py, "src/parser.py", "--in", materialized_docx, "--out", raw_json])
        run_cmd([py, "tools/materialize_effective.py", "--docx", materialized_docx, "--in-json", raw_json, "--out-json", effective_json])
        run_cmd([py, "src/reconstructor.py", "--in-json", effective_json, "--out-docx", reconstructed_docx])
    except subprocess.CalledProcessError as exc:
        print(f"[run_pipeline] step failed with code {exc.returncode}: {exc.cmd}")
        return exc.returncode or 1

    print("[run_pipeline] done")
    return 0


if __name__ == "__main__":
    sys.exit(main())
