from __future__ import annotations

from contextvars import ContextVar
import json
import logging
import threading
import traceback
from datetime import datetime

from sqlalchemy import select
from sqlalchemy.orm import Session

from ..database import SessionLocal
from ..config import settings
from ..importer import (
    FilterSet,
    backfill_missing_addresses,
    count_client_results,
    export_clients,
    export_cpfs_for_enrichment,
    export_wrong_number_clients_for_novavida,
    export_updated_clients_from_latest_enrichment,
    normalize_base_segment,
    process_import_batch,
)
from ..models import AuditLog, ExportHistory, ExportJob, ImportBatch
from ..schemas import ImportRequest
from ..web_helpers import current_utc
from .job_queue import enqueue_background_job


request_id_ctx: ContextVar[str] = ContextVar("request_id", default="-")


def log_event(level: int, event: str, message: str, *, user: str = "-", request_id: str | None = None, data: dict[str, object] | None = None, exc_info=None) -> None:
    logger = logging.getLogger("inss_app")
    logger.log(
        level,
        message,
        extra={
            "event": event,
            "user": user or "-",
            "request_id": request_id if request_id is not None else request_id_ctx.get("-"),
            "extra_data": data or {},
        },
        exc_info=exc_info,
    )


def log_audit(session: Session, actor_username: str, action: str, target_type: str = "", target_id: str = "", message: str = "", metadata: dict[str, object] | None = None) -> None:
    session.add(
        AuditLog(
            actor_username=actor_username or "sistema",
            action=action,
            target_type=target_type,
            target_id=str(target_id or ""),
            message=message,
            metadata_json=json.dumps(metadata or {}, ensure_ascii=False, default=str),
        )
    )


def log_audit_with_new_session(actor_username: str, action: str, target_type: str = "", target_id: str = "", message: str = "", metadata: dict[str, object] | None = None) -> None:
    session = SessionLocal()
    try:
        log_audit(session, actor_username, action, target_type=target_type, target_id=target_id, message=message, metadata=metadata)
        session.commit()
    finally:
        session.close()


def serialize_filters(filters: FilterSet) -> str:
    return json.dumps(filters.__dict__, ensure_ascii=False, default=str)


def deserialize_filters(payload: str) -> FilterSet:
    try:
        raw = json.loads(payload or "{}")
    except json.JSONDecodeError:
        raw = {}
    return FilterSet(**raw)


def summarize_export_filters(filters: FilterSet) -> str:
    parts: list[str] = []
    if filters.base_segment:
        parts.append(f"segmento={filters.base_segment}")
    if filters.year:
        parts.append(f"ano={filters.year}")
    if filters.month:
        parts.append(f"mes={filters.month}")
    if filters.uf:
        parts.append(f"uf={filters.uf.upper()}")
    if filters.city:
        parts.append(f"cidade={filters.city}")
    if filters.ddd:
        parts.append(f"ddd={filters.ddd}")
    if filters.base_source:
        parts.append(f"fonte={filters.base_source}")
    if filters.campaign_profile:
        parts.append(f"perfil={filters.campaign_profile}")
    if filters.has_phone is True:
        parts.append("com_telefone=sim")
    if filters.has_phone is False:
        parts.append("com_telefone=nao")
    return ", ".join(parts) if parts else "sem filtros"


def enqueue_export_job(
    session: Session,
    requested_by: str,
    job_type: str,
    filters: FilterSet | None = None,
    file_format: str = "xlsx",
    include_audit: bool = False,
    message: str = "Exportacao colocada em fila.",
    metadata: dict[str, object] | None = None,
    request_id: str | None = None,
) -> ExportJob:
    job = ExportJob(
        requested_by=requested_by,
        job_type=job_type,
        status="PENDENTE",
        filters_json=serialize_filters(filters) if filters is not None else "{}",
        include_audit=include_audit,
        file_format=file_format,
    )
    session.add(job)
    session.flush()
    log_audit(
        session,
        requested_by,
        "export_job_created",
        target_type="export_job",
        target_id=str(job.id),
        message=message,
        metadata=metadata or {},
    )
    session.commit()
    current_request_id = request_id if request_id is not None else request_id_ctx.get("-")
    log_event(
        logging.INFO,
        "export_job_created",
        "Exportacao colocada em fila.",
        user=requested_by,
        request_id=current_request_id,
        data={"job_id": job.id, "job_type": job_type, "file_format": file_format},
    )
    enqueue_result = enqueue_background_job(
        queue_name=settings.queue_exports_name,
        callable_path="inss_db_app.services.job_orchestration.run_export_job",
        fallback_target=run_export_job,
        fallback_args=(job.id, current_request_id),
        fallback_daemon=True,
        kwargs={"job_id": job.id, "request_id": current_request_id},
        rq_job_timeout=7200,
    )
    job.queue_backend = enqueue_result.get("enqueue_mode", "thread")
    job.queue_name = enqueue_result.get("queue_name", settings.queue_exports_name)
    job.queue_job_id = enqueue_result.get("queue_job_id", "")
    session.commit()
    log_event(
        logging.INFO,
        "export_job_enqueued",
        "Job de exportacao enfileirado.",
        user=requested_by,
        request_id=current_request_id,
        data={"job_id": job.id, "enqueue_mode": job.queue_backend, "queue_backend": settings.queue_backend, "queue_job_id": job.queue_job_id},
    )
    return job


def enqueue_import_job(batch_id: int, request_data: ImportRequest, request_id: str = "-") -> dict[str, str]:
    return enqueue_background_job(
        queue_name=settings.queue_imports_name,
        callable_path="inss_db_app.services.job_orchestration.run_import_job",
        fallback_target=run_import_job,
        fallback_args=(batch_id, request_data, request_id),
        fallback_daemon=False,
        kwargs={"batch_id": batch_id, "request_data": request_data, "request_id": request_id},
        rq_job_timeout=21600,
    )


def run_export_job(job_id: int, request_id: str = "-") -> None:
    request_id_ctx.set(request_id or "-")
    session = SessionLocal()
    try:
        job = session.get(ExportJob, job_id)
        if job is None:
            return
        log_event(
            logging.INFO,
            "export_job_started",
            "Processamento de exportacao iniciado.",
            user=job.requested_by,
            request_id=request_id,
            data={"job_id": job.id, "job_type": job.job_type},
        )
        job.status = "PROCESSANDO"
        job.started_at = current_utc()
        session.commit()

        filters = deserialize_filters(job.filters_json)
        if job.job_type == "campaign":
            total_rows = count_client_results(session, filters)
            if total_rows <= 0:
                raise ValueError(f"Nenhum registro encontrado para os filtros informados ({summarize_export_filters(filters)}).")
            export_path = export_clients(session, filters, include_audit=job.include_audit, file_format=job.file_format)
        elif job.job_type == "cpf_enrichment":
            total_rows = count_client_results(session, filters)
            if total_rows <= 0:
                raise ValueError(f"Nenhum CPF encontrado para os filtros informados ({summarize_export_filters(filters)}).")
            export_path = export_cpfs_for_enrichment(session, filters, file_format=job.file_format)
        elif job.job_type == "novavida_wrong_number":
            total_rows = count_client_results(session, filters)
            if total_rows <= 0:
                raise ValueError(f"Nenhum registro encontrado para os filtros informados ({summarize_export_filters(filters)}).")
            export_path = export_wrong_number_clients_for_novavida(session, filters, file_format=job.file_format)
        elif job.job_type == "updated_base":
            export_path = export_updated_clients_from_latest_enrichment(
                session,
                base_segment=filters.base_segment or "INSS",
                file_format=job.file_format,
            )
        else:
            raise ValueError(f"Tipo de job de exportacao nao suportado: {job.job_type}")

        job.output_file = str(export_path)
        job.total_records = session.scalar(select(ExportHistory.total_registros).where(ExportHistory.arquivo_saida == str(export_path)).order_by(ExportHistory.id.desc()).limit(1)) or 0
        job.status = "CONCLUIDO"
        job.finished_at = current_utc()
        log_audit(session, job.requested_by, "export_job_completed", target_type="export_job", target_id=str(job.id), message="Exportacao assincrona concluida.", metadata={"arquivo": job.output_file, "total_registros": job.total_records})
        session.commit()
        log_event(
            logging.INFO,
            "export_job_completed",
            "Exportacao concluida.",
            user=job.requested_by,
            request_id=request_id,
            data={"job_id": job.id, "job_type": job.job_type, "records": job.total_records, "output_file": job.output_file},
        )
    except Exception as exc:
        session.rollback()
        failed_job = session.get(ExportJob, job_id)
        if failed_job is not None:
            failed_job.status = "ERRO"
            failed_job.error_message = str(exc)
            failed_job.finished_at = current_utc()
            log_audit(session, failed_job.requested_by, "export_job_failed", target_type="export_job", target_id=str(failed_job.id), message="Falha na exportacao assincrona.", metadata={"erro": str(exc)})
            session.commit()
            log_event(
                logging.ERROR,
                "export_job_failed",
                "Falha na exportacao assincrona.",
                user=failed_job.requested_by,
                request_id=request_id,
                data={"job_id": failed_job.id, "job_type": failed_job.job_type, "error": str(exc)},
                exc_info=True,
            )
    finally:
        session.close()


def run_import_job(batch_id: int, request_data: ImportRequest, request_id: str = "-") -> None:
    request_id_ctx.set(request_id or "-")
    session = SessionLocal()
    try:
        batch = session.get(ImportBatch, batch_id)
        if batch is None:
            return
        log_event(
            logging.INFO,
            "import_batch_started",
            "Processamento de importacao iniciado.",
            user=request_data.user_name,
            request_id=request_id,
            data={"batch_id": batch_id, "base_segment": request_data.base_segment, "files": len(request_data.files)},
        )
        summary = process_import_batch(session, batch, request_data)
        log_event(
            logging.INFO,
            "import_batch_completed",
            "Importacao concluida.",
            user=request_data.user_name,
            request_id=request_id,
            data={"batch_id": batch_id, "rows_imported": summary.rows_imported, "files_imported": summary.files_imported},
        )
        log_audit_with_new_session(
            request_data.user_name,
            "import_batch_completed",
            target_type="batch",
            target_id=str(batch_id),
            message="Importacao concluida.",
            metadata={"rows_imported": summary.rows_imported, "files_imported": summary.files_imported},
        )
        if normalize_base_segment(request_data.base_segment) == "CREFAZ":
            threading.Thread(
                target=run_post_import_enrichment,
                args=(batch_id, request_data.user_name),
                daemon=True,
            ).start()
    except Exception as exc:
        traceback.print_exc()
        log_event(
            logging.ERROR,
            "import_batch_failed",
            "Falha na importacao.",
            user=request_data.user_name,
            request_id=request_id,
            data={"batch_id": batch_id, "error": str(exc)},
            exc_info=True,
        )
        try:
            failed_batch = session.get(ImportBatch, batch_id)
            if failed_batch is not None and failed_batch.status == "PROCESSANDO":
                failed_batch.status = "ERRO"
                failed_batch.resumo = f"Falha na importacao (worker): {exc}"
                session.commit()
        except Exception:
            session.rollback()
        log_audit_with_new_session(
            request_data.user_name,
            "import_batch_failed",
            target_type="batch",
            target_id=str(batch_id),
            message="Falha na importacao.",
            metadata={"erro": str(exc), "traceback": traceback.format_exc(limit=20)},
        )
    finally:
        session.close()


def run_startup_address_backfill() -> None:
    session = SessionLocal()
    try:
        summary = backfill_missing_addresses(session)
        log_event(
            logging.INFO,
            "startup_address_backfill_completed",
            "Backfill de endereco no startup concluido.",
            user="sistema",
            request_id="-",
            data={
                "scanned": summary.scanned_rows,
                "updated": summary.updated_rows,
                "rebuilt_clients": summary.rebuilt_clients,
                "skipped_invalid_cep": summary.skipped_invalid_cep,
            },
        )
    except Exception as exc:
        log_event(
            logging.ERROR,
            "startup_address_backfill_failed",
            "Falha no backfill de endereco no startup.",
            user="sistema",
            request_id="-",
            data={"error": str(exc)},
            exc_info=True,
        )
    finally:
        session.close()


def run_post_import_enrichment(batch_id: int, actor_username: str) -> None:
    session = SessionLocal()
    try:
        summary = backfill_missing_addresses(session)
        log_audit(
            session,
            actor_username or "sistema",
            "post_import_enrichment_completed",
            target_type="batch",
            target_id=str(batch_id),
            message="Enriquecimento pos-importacao concluido.",
            metadata={
                "scanned_rows": summary.scanned_rows,
                "updated_rows": summary.updated_rows,
                "rebuilt_clients": summary.rebuilt_clients,
                "skipped_invalid_cep": summary.skipped_invalid_cep,
            },
        )
        session.commit()
    except Exception as exc:
        session.rollback()
        log_audit_with_new_session(
            actor_username or "sistema",
            "post_import_enrichment_failed",
            target_type="batch",
            target_id=str(batch_id),
            message="Falha no enriquecimento pos-importacao.",
            metadata={"erro": str(exc), "traceback": traceback.format_exc(limit=20)},
        )
    finally:
        session.close()


def recover_stuck_import_batches() -> int:
    session = SessionLocal()
    try:
        stuck_batches = (
            session.execute(select(ImportBatch).where(ImportBatch.status == "PROCESSANDO"))
            .scalars()
            .all()
        )
        if not stuck_batches:
            return 0
        stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        for batch in stuck_batches:
            batch.status = "ERRO"
            previous = (batch.resumo or "").strip()
            suffix = f" Marcado automaticamente como ERRO em {stamp}: processamento interrompido (worker inativo). Reprocessar lote."
            batch.resumo = f"{previous} | {suffix}".strip(" |")
        session.commit()
        return len(stuck_batches)
    finally:
        session.close()
