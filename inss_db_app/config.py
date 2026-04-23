from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent


def build_database_url() -> str:
    explicit_url = os.getenv("DATABASE_URL", "").strip()
    if explicit_url:
        return explicit_url

    pg_host = os.getenv("PGHOST", "").strip()
    pg_database = os.getenv("PGDATABASE", "").strip()
    if pg_host and pg_database:
        pg_port = os.getenv("PGPORT", "5432").strip() or "5432"
        pg_user = os.getenv("PGUSER", "postgres").strip() or "postgres"
        pg_password = os.getenv("PGPASSWORD", "").strip()
        auth = pg_user
        if pg_password:
            auth += f":{pg_password}"
        return f"postgresql+psycopg://{auth}@{pg_host}:{pg_port}/{pg_database}"

    return f"sqlite:///{(PROJECT_DIR / 'inss_clientes.db').as_posix()}"


@dataclass(frozen=True)
class Settings:
    database_url: str = build_database_url()
    uploads_dir: Path = Path(os.getenv("UPLOADS_DIR", str(PROJECT_DIR / "data" / "uploads")))
    archive_dir: Path = Path(os.getenv("ARCHIVE_DIR", str(PROJECT_DIR / "data" / "archive")))
    exports_dir: Path = Path(os.getenv("EXPORTS_DIR", str(PROJECT_DIR / "data" / "exports")))
    host: str = os.getenv("APP_HOST", "0.0.0.0")
    port: int = int(os.getenv("APP_PORT", "8000"))
    session_secret_key: str = os.getenv("APP_SESSION_SECRET", "inss-clientes-session-secret")
    default_admin_username: str = os.getenv("APP_DEFAULT_ADMIN_USERNAME", "admin")
    default_admin_password: str = os.getenv("APP_DEFAULT_ADMIN_PASSWORD", "admin123")


settings = Settings()
