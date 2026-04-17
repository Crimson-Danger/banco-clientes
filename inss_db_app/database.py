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
        "CREATE INDEX IF NOT EXISTS ix_client_occurrences_nu_nb ON client_occurrences (nu_nb)",
        "CREATE INDEX IF NOT EXISTS ix_client_occurrences_esp ON client_occurrences (esp)",
        "CREATE INDEX IF NOT EXISTS ix_client_occurrences_ddb ON client_occurrences (ddb)",
    ]
    with engine.begin() as connection:
        for statement in statements:
            connection.execute(text(statement))


class Base(DeclarativeBase):
    pass
