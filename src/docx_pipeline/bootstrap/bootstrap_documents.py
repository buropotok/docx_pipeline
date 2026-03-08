from __future__ import annotations

from pathlib import Path

from docx_pipeline.config.settings import get_settings
from docx_pipeline.db.repositories import (
    get_document_by_mail_uid_and_attachment_index,
    insert_document,
)


ALLOWED_EXTENSIONS = {".doc", ".docx"}


def find_source_files(mail_uid_dir: Path) -> list[Path]:
    candidates = [
        p for p in mail_uid_dir.iterdir()
        if p.is_file() and p.suffix.lower() in ALLOWED_EXTENSIONS
    ]
    return sorted(candidates, key=lambda p: (p.name.lower(), p.suffix.lower()))


def bootstrap_documents() -> None:
    settings = get_settings()
    settings.artifacts_dir.mkdir(parents=True, exist_ok=True)

    if not settings.base_documents_dir.exists():
        raise FileNotFoundError(
            f"Base documents directory not found: {settings.base_documents_dir}"
        )

    mail_uid_dirs = sorted(
        [p for p in settings.base_documents_dir.iterdir() if p.is_dir() and p.name.startswith("uid_")],
        key=lambda p: p.name.lower(),
    )

    print(f"Found {len(mail_uid_dirs)} mail uid directories")

    inserted = 0
    skipped = 0
    missing_source = 0

    for mail_uid_dir in mail_uid_dirs:
        mail_uid = mail_uid_dir.name
        source_files = find_source_files(mail_uid_dir)

        if not source_files:
            print(f"[WARN] {mail_uid} has no .doc/.docx source files")
            missing_source += 1
            continue

        for attachment_index, source_file in enumerate(source_files, start=1):
            document_uid = f"{mail_uid}_{attachment_index}"

            existing = get_document_by_mail_uid_and_attachment_index(
                mail_uid=mail_uid,
                attachment_index=attachment_index,
            )
            if existing:
                print(f"[SKIP] {document_uid} already exists in DB")
                skipped += 1
                continue

            artifact_dir = settings.artifacts_dir / document_uid
            artifact_dir.mkdir(parents=True, exist_ok=True)

            insert_document(
                uid=document_uid,
                mail_uid=mail_uid,
                attachment_index=attachment_index,
                source_filename=source_file.name,
                source_ext=source_file.suffix.lower(),
                source_abs_path=str(source_file.resolve()),
                artifacts_abs_path=str(artifact_dir.resolve()),
                processing_status="discovered",
            )

            # print(f"[OK] inserted {document_uid} -> {source_file.name}")
            inserted += 1

    print()
    print(f"Inserted: {inserted}")
    print(f"Skipped : {skipped}")
    print(f"Missing : {missing_source}")


if __name__ == "__main__":
    bootstrap_documents()
