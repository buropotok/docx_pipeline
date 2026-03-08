from pathlib import Path

# Укажите путь к корню проекта
PROJECT_ROOT = Path(r"C:\Users\sokol\PycharmProjects\Buropotok\docx_pipeline")

DIRS = [
    "src/docx_pipeline/config",
    "src/docx_pipeline/db",
    "src/docx_pipeline/bootstrap",
    "src/docx_pipeline/pipeline/saveas",
    "src/docx_pipeline/pipeline/parse",
    "src/docx_pipeline/pipeline/fingerprint_input",
    "src/docx_pipeline/pipeline/classify",
    "src/docx_pipeline/utils",
    "scripts",
]

FILES = [
    "src/docx_pipeline/config/settings.py",
    "src/docx_pipeline/db/__init__.py",
    "src/docx_pipeline/db/db.py",
    "src/docx_pipeline/db/schema.sql",
    "src/docx_pipeline/db/repositories.py",
    "src/docx_pipeline/db/init_reference_data.py",
    "src/docx_pipeline/bootstrap/__init__.py",
    "src/docx_pipeline/bootstrap/bootstrap_documents.py",
    "src/docx_pipeline/bootstrap/scan_documents.py",
    "src/docx_pipeline/utils/__init__.py",
    "src/docx_pipeline/utils/fs.py",
    "src/docx_pipeline/utils/logging.py",
    "scripts/bootstrap_db.py",
    "scripts/bootstrap_documents.py",
]

def main() -> None:
    PROJECT_ROOT.mkdir(parents=True, exist_ok=True)

    for rel_dir in DIRS:
        path = PROJECT_ROOT / rel_dir
        path.mkdir(parents=True, exist_ok=True)
        print(f"[DIR ] {path}")

    for rel_file in FILES:
        path = PROJECT_ROOT / rel_file
        path.parent.mkdir(parents=True, exist_ok=True)
        if not path.exists():
            path.touch()
            print(f"[FILE] {path}")
        else:
            print(f"[SKIP] {path} already exists")

    print("\nDone.")

if __name__ == "__main__":
    main()