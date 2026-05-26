import os
import sqlite3
from pathlib import Path

from sqlalchemy import create_engine
from sqlalchemy.orm import declarative_base, sessionmaker


def _local_database_url():
    local_root = Path(
        os.getenv("LOCALAPPDATA") or Path.home() / "AppData" / "Local"
    )
    database_dir = local_root / "AIResumeCoach"
    database_path = database_dir / "resume_app.db"
    legacy_path = Path(__file__).resolve().parent / "resume_app.db"

    database_dir.mkdir(parents=True, exist_ok=True)
    if not database_path.exists() and legacy_path.exists():
        source = sqlite3.connect(
            f"{legacy_path.as_uri()}?mode=ro&immutable=1",
            uri=True
        )
        destination = sqlite3.connect(str(database_path))
        try:
            source.backup(destination)
        finally:
            destination.close()
            source.close()

    return f"sqlite:///{database_path.as_posix()}"


DATABASE_URL = os.getenv("DATABASE_URL") or _local_database_url()

engine_options = {"pool_pre_ping": True}
if DATABASE_URL.startswith("sqlite"):
    engine_options["connect_args"] = {
        "check_same_thread": False,
        "timeout": 15
    }

engine = create_engine(DATABASE_URL, **engine_options)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)
Base = declarative_base()
