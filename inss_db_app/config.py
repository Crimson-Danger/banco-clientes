from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent


def env_flag(name: str, default: bool = False) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on", "sim"}


def build_database_url() -> str:
    explicit_url = os.getenv("DATABASE_URL", "").strip()
    if explicit_url:
        if not explicit_url.startswith("postgresql+psycopg://"):
            raise RuntimeError("DATABASE_URL invalida: apenas PostgreSQL (postgresql+psycopg://) e suportado.")
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
    raise RuntimeError(
        "Configuracao de banco ausente. Defina DATABASE_URL (postgresql+psycopg://...) "
        "ou PGHOST/PGPORT/PGUSER/PGPASSWORD/PGDATABASE."
    )


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
    enable_cep_enrichment: bool = env_flag("APP_ENABLE_CEP_ENRICHMENT", default=False)
    viacep_timeout_seconds: float = float(os.getenv("APP_VIACEP_TIMEOUT_SECONDS", "3"))
    viacep_base_url: str = os.getenv("APP_VIACEP_BASE_URL", "https://viacep.com.br/ws").rstrip("/")
    enable_ddd_enrichment: bool = env_flag("APP_ENABLE_DDD_ENRICHMENT", default=False)
    brasilapi_timeout_seconds: float = float(os.getenv("APP_BRASILAPI_TIMEOUT_SECONDS", "3"))
    brasilapi_base_url: str = os.getenv("APP_BRASILAPI_BASE_URL", "https://brasilapi.com.br/api").rstrip("/")
    enable_startup_address_backfill: bool = env_flag("APP_ENABLE_STARTUP_ADDRESS_BACKFILL", default=True)
    startup_address_backfill_batch_size: int = int(os.getenv("APP_STARTUP_ADDRESS_BACKFILL_BATCH_SIZE", "250"))
    enable_startup_schema_maintenance: bool = env_flag("APP_ENABLE_STARTUP_SCHEMA_MAINTENANCE", default=True)
    facta_homolog_base: str = os.getenv("FACTA_HOMOLOG_BASE", "https://webservice-homol.facta.com.br").rstrip("/")
    facta_offline_base: str = os.getenv("FACTA_OFFLINE_BASE", "https://cltoff.facta.com.br").rstrip("/")
    facta_username: str = os.getenv("FACTA_USERNAME", "")
    facta_password: str = os.getenv("FACTA_PASSWORD", "")
    c6_worker_enabled: bool = env_flag("C6_WORKER_ENABLED", default=False)
    c6_base_url: str = os.getenv("C6_BASE_URL", "https://marketplace-proposal-service-api-p.c6bank.info").rstrip("/")
    c6_username: str = os.getenv("C6_USERNAME", "").strip()
    c6_password: str = os.getenv("C6_PASSWORD", "").strip()
    c6_timeout_seconds: float = float(os.getenv("C6_TIMEOUT_SECONDS", "20"))


settings = Settings()
