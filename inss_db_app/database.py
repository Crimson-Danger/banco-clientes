from __future__ import annotations

from contextlib import contextmanager
from threading import RLock

from sqlalchemy import create_engine, text
from sqlalchemy import event
from sqlalchemy.orm import DeclarativeBase, sessionmaker

from .config import settings


connect_args = {}
engine_kwargs = {"future": True}
if settings.database_url.startswith("sqlite"):
    connect_args["check_same_thread"] = False
    connect_args["timeout"] = 30
else:
    engine_kwargs["pool_pre_ping"] = True

if connect_args:
    engine_kwargs["connect_args"] = connect_args

engine = create_engine(settings.database_url, **engine_kwargs)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False, future=True)
write_lock = RLock()


if settings.database_url.startswith("sqlite"):
    @event.listens_for(engine, "connect")
    def configure_sqlite(dbapi_connection, _connection_record) -> None:
        cursor = dbapi_connection.cursor()
        cursor.execute("PRAGMA foreign_keys=ON")
        cursor.execute("PRAGMA journal_mode=WAL")
        cursor.execute("PRAGMA synchronous=NORMAL")
        cursor.execute("PRAGMA busy_timeout=30000")
        cursor.close()


@contextmanager
def serialized_write():
    with write_lock:
        yield


def ensure_runtime_indexes() -> None:
    if not settings.database_url.startswith("postgresql"):
        return
    statements = [
        "ALTER TABLE import_batches ADD COLUMN IF NOT EXISTS base_segment VARCHAR(12) NOT NULL DEFAULT 'INSS'",
        "ALTER TABLE import_batches ADD COLUMN IF NOT EXISTS base_source VARCHAR(120) NOT NULL DEFAULT ''",
        "ALTER TABLE source_files ADD COLUMN IF NOT EXISTS base_segment VARCHAR(12) NOT NULL DEFAULT 'INSS'",
        "ALTER TABLE source_files ADD COLUMN IF NOT EXISTS base_source VARCHAR(120) NOT NULL DEFAULT ''",
        "ALTER TABLE client_occurrences ADD COLUMN IF NOT EXISTS base_segment VARCHAR(12) NOT NULL DEFAULT 'INSS'",
        "ALTER TABLE client_occurrences ADD COLUMN IF NOT EXISTS base_source VARCHAR(120) NOT NULL DEFAULT ''",
        "ALTER TABLE client_occurrences ADD COLUMN IF NOT EXISTS cpf_original VARCHAR(20) NOT NULL DEFAULT ''",
        "ALTER TABLE client_occurrences ADD COLUMN IF NOT EXISTS nome_higienizado VARCHAR(255) NOT NULL DEFAULT ''",
        "ALTER TABLE client_occurrences ADD COLUMN IF NOT EXISTS matricula VARCHAR(80) NOT NULL DEFAULT ''",
        "ALTER TABLE client_occurrences ADD COLUMN IF NOT EXISTS servico VARCHAR(255) NOT NULL DEFAULT ''",
        "ALTER TABLE client_occurrences ADD COLUMN IF NOT EXISTS situacao VARCHAR(120) NOT NULL DEFAULT ''",
        "ALTER TABLE client_occurrences ADD COLUMN IF NOT EXISTS pis VARCHAR(20) NOT NULL DEFAULT ''",
        "ALTER TABLE client_occurrences ADD COLUMN IF NOT EXISTS entidade VARCHAR(255) NOT NULL DEFAULT ''",
        "ALTER TABLE client_occurrences ADD COLUMN IF NOT EXISTS secretaria VARCHAR(255) NOT NULL DEFAULT ''",
        "ALTER TABLE client_occurrences ADD COLUMN IF NOT EXISTS convenio VARCHAR(255) NOT NULL DEFAULT ''",
        "ALTER TABLE client_occurrences ADD COLUMN IF NOT EXISTS cbo_titulo VARCHAR(255) NOT NULL DEFAULT ''",
        "ALTER TABLE client_occurrences ADD COLUMN IF NOT EXISTS salario NUMERIC(14, 2)",
        "ALTER TABLE client_occurrences ADD COLUMN IF NOT EXISTS margem_liquida NUMERIC(14, 2)",
        "ALTER TABLE client_occurrences ADD COLUMN IF NOT EXISTS margem_bruta NUMERIC(14, 2)",
        "ALTER TABLE client_occurrences ADD COLUMN IF NOT EXISTS margem_utilizada NUMERIC(14, 2)",
        "ALTER TABLE client_occurrences ADD COLUMN IF NOT EXISTS margem_total NUMERIC(14, 2)",
        "ALTER TABLE client_occurrences ADD COLUMN IF NOT EXISTS extras_json TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE clients ADD COLUMN IF NOT EXISTS tem_inss BOOLEAN NOT NULL DEFAULT FALSE",
        "ALTER TABLE clients ADD COLUMN IF NOT EXISTS tem_governo BOOLEAN NOT NULL DEFAULT FALSE",
        "ALTER TABLE clients ADD COLUMN IF NOT EXISTS pis_atual VARCHAR(20) NOT NULL DEFAULT ''",
        "ALTER TABLE clients ADD COLUMN IF NOT EXISTS matricula_atual VARCHAR(80) NOT NULL DEFAULT ''",
        "ALTER TABLE clients ADD COLUMN IF NOT EXISTS entidade_atual VARCHAR(255) NOT NULL DEFAULT ''",
        "ALTER TABLE app_users ADD COLUMN IF NOT EXISTS last_login_at TIMESTAMP",
        "ALTER TABLE app_users ADD COLUMN IF NOT EXISTS last_password_reset_at TIMESTAMP",
        "ALTER TABLE app_users ADD COLUMN IF NOT EXISTS locked_until TIMESTAMP",
        "CREATE TABLE IF NOT EXISTS enrichment_import_items (id SERIAL PRIMARY KEY, cpf VARCHAR(11) NOT NULL, fonte VARCHAR(40) NOT NULL DEFAULT '', arquivo_origem VARCHAR(255) NOT NULL DEFAULT '', importado_por VARCHAR(120) NOT NULL DEFAULT 'operador', criado_em TIMESTAMP NOT NULL DEFAULT NOW())",
        "CREATE UNIQUE INDEX IF NOT EXISTS uq_enrichment_import_item ON enrichment_import_items (cpf, arquivo_origem, fonte)",
        "CREATE INDEX IF NOT EXISTS ix_enrichment_import_items_cpf ON enrichment_import_items (cpf)",
        "CREATE INDEX IF NOT EXISTS ix_enrichment_import_items_arquivo_origem ON enrichment_import_items (arquivo_origem)",
        "CREATE TABLE IF NOT EXISTS saved_filters (id SERIAL PRIMARY KEY, user_id INTEGER NOT NULL REFERENCES app_users(id), page_key VARCHAR(40) NOT NULL DEFAULT 'clients', name VARCHAR(120) NOT NULL, query_string TEXT NOT NULL DEFAULT '', created_at TIMESTAMP NOT NULL DEFAULT NOW())",
        "CREATE INDEX IF NOT EXISTS ix_saved_filters_user_id ON saved_filters (user_id)",
        "CREATE INDEX IF NOT EXISTS ix_saved_filters_page_key ON saved_filters (page_key)",
        "CREATE TABLE IF NOT EXISTS audit_logs (id SERIAL PRIMARY KEY, actor_username VARCHAR(120) NOT NULL DEFAULT 'sistema', action VARCHAR(80) NOT NULL, target_type VARCHAR(60) NOT NULL DEFAULT '', target_id VARCHAR(120) NOT NULL DEFAULT '', message VARCHAR(255) NOT NULL DEFAULT '', metadata_json TEXT NOT NULL DEFAULT '', created_at TIMESTAMP NOT NULL DEFAULT NOW())",
        "CREATE INDEX IF NOT EXISTS ix_audit_logs_action ON audit_logs (action)",
        "CREATE INDEX IF NOT EXISTS ix_audit_logs_actor_username ON audit_logs (actor_username)",
        "CREATE INDEX IF NOT EXISTS ix_audit_logs_created_at ON audit_logs (created_at)",
        "CREATE TABLE IF NOT EXISTS export_jobs (id SERIAL PRIMARY KEY, requested_by VARCHAR(120) NOT NULL DEFAULT 'operador', job_type VARCHAR(40) NOT NULL DEFAULT 'campaign', status VARCHAR(40) NOT NULL DEFAULT 'PENDENTE', filters_json TEXT NOT NULL DEFAULT '', include_audit BOOLEAN NOT NULL DEFAULT FALSE, file_format VARCHAR(10) NOT NULL DEFAULT 'xlsx', output_file VARCHAR(500) NOT NULL DEFAULT '', total_records INTEGER NOT NULL DEFAULT 0, error_message TEXT NOT NULL DEFAULT '', created_at TIMESTAMP NOT NULL DEFAULT NOW(), started_at TIMESTAMP, finished_at TIMESTAMP)",
        "CREATE INDEX IF NOT EXISTS ix_export_jobs_status ON export_jobs (status)",
        "CREATE INDEX IF NOT EXISTS ix_export_jobs_requested_by ON export_jobs (requested_by)",
        "CREATE INDEX IF NOT EXISTS ix_export_jobs_created_at ON export_jobs (created_at)",
        "CREATE INDEX IF NOT EXISTS ix_client_occurrences_nu_nb ON client_occurrences (nu_nb)",
        "CREATE INDEX IF NOT EXISTS ix_client_occurrences_esp ON client_occurrences (esp)",
        "CREATE INDEX IF NOT EXISTS ix_client_occurrences_ddb ON client_occurrences (ddb)",
        "CREATE INDEX IF NOT EXISTS ix_client_occurrences_matricula ON client_occurrences (matricula)",
        "CREATE INDEX IF NOT EXISTS ix_client_occurrences_servico ON client_occurrences (servico)",
    ]
    with engine.begin() as connection:
        for statement in statements:
            connection.execute(text(statement))


class Base(DeclarativeBase):
    pass
