from __future__ import annotations

import importlib
import sys
from pathlib import Path
from types import SimpleNamespace

from sqlalchemy import create_engine, select
from sqlalchemy.orm import Session, sessionmaker

from inss_db_app import main
from inss_db_app.database import Base
from inss_db_app.importer import FilterSet
from inss_db_app.models import ExportJob


def make_session() -> Session:
    engine = create_engine("sqlite:///:memory:", future=True)
    Base.metadata.create_all(engine)
    return sessionmaker(bind=engine, autoflush=False, autocommit=False, future=True)()


def reload_config_module():
    sys.modules.pop("inss_db_app.config", None)
    return importlib.import_module("inss_db_app.config")


def test_config_blocks_weak_secret_in_production(monkeypatch) -> None:
    monkeypatch.setenv("DATABASE_URL", "postgresql+psycopg://user:pass@localhost:5432/db")
    monkeypatch.setenv("APP_ENV", "production")
    monkeypatch.setenv("APP_SESSION_SECRET", "inss-clientes-session-secret")
    monkeypatch.setenv("APP_DEFAULT_ADMIN_PASSWORD", "SenhaMuitoForte123!")

    try:
        reload_config_module()
        assert False, "Era esperado RuntimeError para APP_SESSION_SECRET fraco em producao."
    except RuntimeError as exc:
        assert "APP_SESSION_SECRET" in str(exc)


def test_config_blocks_weak_admin_password_in_production(monkeypatch) -> None:
    monkeypatch.setenv("DATABASE_URL", "postgresql+psycopg://user:pass@localhost:5432/db")
    monkeypatch.setenv("APP_ENV", "production")
    monkeypatch.setenv("APP_SESSION_SECRET", "segredo_super_forte_com_mais_de_32_caracteres_123")
    monkeypatch.setenv("APP_DEFAULT_ADMIN_PASSWORD", "admin123")

    try:
        reload_config_module()
        assert False, "Era esperado RuntimeError para APP_DEFAULT_ADMIN_PASSWORD fraca em producao."
    except RuntimeError as exc:
        assert "APP_DEFAULT_ADMIN_PASSWORD" in str(exc)


def test_ready_endpoint_returns_ok_when_db_and_directories_are_ready(monkeypatch, tmp_path: Path) -> None:
    uploads = tmp_path / "uploads"
    archive = tmp_path / "archive"
    exports = tmp_path / "exports"
    uploads.mkdir()
    archive.mkdir()
    exports.mkdir()

    monkeypatch.setattr(
        main,
        "settings",
        SimpleNamespace(uploads_dir=uploads, archive_dir=archive, exports_dir=exports),
    )

    session = make_session()
    response = main.readiness_check(session=session)
    assert response.status_code == 200
    assert b'"status":"ok"' in response.body


def test_ready_endpoint_returns_degraded_when_directory_is_missing(monkeypatch, tmp_path: Path) -> None:
    uploads = tmp_path / "uploads"
    archive = tmp_path / "archive"
    missing_exports = tmp_path / "exports_missing"
    uploads.mkdir()
    archive.mkdir()

    monkeypatch.setattr(
        main,
        "settings",
        SimpleNamespace(uploads_dir=uploads, archive_dir=archive, exports_dir=missing_exports),
    )

    session = make_session()
    response = main.readiness_check(session=session)
    assert response.status_code == 503
    assert b'"status":"degraded"' in response.body
    assert b"directory_exports" in response.body


def test_enqueue_export_job_creates_pending_job_and_starts_worker(monkeypatch) -> None:
    engine = create_engine("sqlite:///:memory:", future=True)
    Base.metadata.create_all(engine)
    session_factory = sessionmaker(bind=engine, autoflush=False, autocommit=False, future=True)
    monkeypatch.setattr(main, "SessionLocal", session_factory)

    started: dict[str, object] = {}

    class FakeThread:
        def __init__(self, target, args, daemon):
            started["target"] = target
            started["args"] = args
            started["daemon"] = daemon

        def start(self):
            started["started"] = True

    monkeypatch.setattr(main.threading, "Thread", FakeThread)

    session = session_factory()
    job = main.enqueue_export_job(
        session=session,
        requested_by="tester",
        job_type="campaign",
        filters=FilterSet(base_segment="INSS"),
        request_id="req-enqueue-1",
    )
    persisted = session.get(ExportJob, job.id)
    assert persisted is not None
    assert persisted.status == "PENDENTE"
    assert persisted.queue_backend in {"thread", "rq"}
    assert persisted.queue_job_id != ""
    assert started.get("started") is True
    assert started.get("args") == (job.id, "req-enqueue-1")


def test_run_export_job_executes_and_marks_as_completed(monkeypatch, tmp_path: Path) -> None:
    engine = create_engine("sqlite:///:memory:", future=True)
    Base.metadata.create_all(engine)
    session_factory = sessionmaker(bind=engine, autoflush=False, autocommit=False, future=True)
    monkeypatch.setattr(main, "SessionLocal", session_factory)
    monkeypatch.setattr(
        main.threading,
        "Thread",
        lambda target, args, daemon: SimpleNamespace(start=lambda: None),
    )

    session = session_factory()
    job = main.enqueue_export_job(
        session=session,
        requested_by="tester",
        job_type="campaign",
        filters=FilterSet(base_segment="INSS"),
        request_id="req-run-1",
    )

    output_file = tmp_path / "export_campaign.xlsx"
    output_file.write_text("ok", encoding="utf-8")

    monkeypatch.setattr(main, "count_client_results", lambda _session, _filters: 5)
    monkeypatch.setattr(main, "export_clients", lambda _session, _filters, include_audit=False, file_format="xlsx": output_file)

    main.run_export_job(job.id, request_id="req-run-1")

    verification_session = session_factory()
    persisted = verification_session.scalar(select(ExportJob).where(ExportJob.id == job.id))
    assert persisted is not None
    assert persisted.status == "CONCLUIDO"
    assert persisted.output_file == str(output_file)
