from __future__ import annotations

import logging
import threading

from ..auth import hash_password
from ..config import settings
from ..database import Base, SessionLocal, engine, ensure_runtime_indexes, validate_runtime_schema
from ..models import AppUser
from .job_orchestration import recover_stuck_import_batches, run_startup_address_backfill


def ensure_default_admin() -> None:
    session = SessionLocal()
    try:
        existing_user = session.query(AppUser).count()
        if existing_user:
            return
        password_hash, password_salt = hash_password(settings.default_admin_password)
        session.add(
            AppUser(
                username=settings.default_admin_username,
                full_name="Administrador",
                password_hash=password_hash,
                password_salt=password_salt,
                is_active=True,
                is_admin=True,
                can_view_dashboard=True,
                can_import=True,
                can_search=True,
                can_access_consignado_inss=True,
                can_export=True,
                can_view_history=True,
                can_delete_batches=True,
                can_manage_users=True,
            )
        )
        session.commit()
    finally:
        session.close()


def create_schema() -> None:
    Base.metadata.create_all(bind=engine)
    ensure_runtime_indexes()
    ensure_default_admin()


def run_startup(
    log_event,
) -> None:
    if settings.enable_startup_schema_maintenance:
        create_schema()
        log_event(logging.INFO, "startup_schema_ready", "Startup maintenance: schema ready", user="sistema", request_id="-")
    else:
        if not settings.skip_runtime_schema_validation:
            validate_runtime_schema()
        else:
            log_event(logging.INFO, "startup_schema_validation_skipped", "Startup maintenance: runtime schema validation skipped", user="sistema", request_id="-")
        ensure_default_admin()
        log_event(logging.INFO, "startup_schema_validated", "Startup maintenance: schema validated", user="sistema", request_id="-")

    recovered = recover_stuck_import_batches()
    if recovered:
        log_event(logging.WARNING, "startup_recovery_batches", "Lotes PROCESSANDO foram marcados como ERRO no startup.", user="sistema", request_id="-", data={"recovered_batches": recovered})
    if settings.enable_startup_address_backfill:
        threading.Thread(target=run_startup_address_backfill, daemon=True).start()
