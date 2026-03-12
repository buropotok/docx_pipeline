from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class Settings:
    project_root: Path
    home_path: Path
    base_documents_dir: Path
    artifacts_dir: Path
    db_path: Path


def get_settings() -> Settings:
    project_root = Path(r"C:\Users\sokol\PycharmProjects\Buropotok\docx_pipeline")
    home_path = Path(r"E:\Buro_potok")

    return Settings(
        project_root=project_root,
        home_path=home_path,
        base_documents_dir=home_path / "base_documents",
        artifacts_dir=home_path / "artifacts",
        db_path=home_path / "db" / "translate_factory.sqlite",
    )