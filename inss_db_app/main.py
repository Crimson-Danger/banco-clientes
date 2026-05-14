from __future__ import annotations

from contextlib import asynccontextmanager
from contextvars import ContextVar
import json
import logging
import mimetypes
import re
from pathlib import Path
import threading
import traceback
import uuid
from types import SimpleNamespace
from urllib.parse import urlencode
from urllib.parse import quote_plus
from datetime import UTC, datetime
from datetime import timedelta

from fastapi import Body, Depends, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy import case, func, or_, select
from sqlalchemy.orm import Session

from .auth import hash_password, verify_password
from .c6_worker_loan import C6WorkerLoanError, c6_worker_loan_client
from .config import settings
from .database import SessionLocal
from .importer import (
    backfill_missing_addresses,
    rebuild_clients_for_cpfs,
    FilterSet,
    count_client_results,
    create_import_batch,
    delete_batch,
    export_clients,
    export_cpfs_for_enrichment,
    export_wrong_number_clients_for_novavida,
    export_updated_clients_from_latest_enrichment,
    format_money,
    import_public_campaign_updates,
    import_phone_enrichments,
    normalize_base_segment,
    normalize_cpf,
    parse_decimal,
    process_import_batch,
    query_clients,
    save_uploaded_files,
)
from .models import AppUser, AuditLog, Client, ClientOccurrence, ExportHistory, ExportJob, ImportBatch, PhoneEnrichment, PublicCampaignUpdate, SavedFilter, SourceFile
from .security_helpers import PERMISSION_LABELS, build_permissions, create_auth_token, parse_auth_token, user_can
from .schemas import ImportRequest
from .services.job_orchestration import (
    enqueue_import_job as enqueue_import_job_service,
    enqueue_export_job as enqueue_export_job_service,
    run_export_job as run_export_job_service,
)
from .services.job_queue import cancel_rq_job
from .services.audit_service import log_audit, log_audit_with_new_session
from .services.startup_service import create_schema, ensure_default_admin, run_startup
from .web_helpers import (
    build_filter_query_params,
    build_pagination,
    current_utc,
    format_dashboard_datetime,
    format_phone,
    format_reference,
    is_future_timestamp,
    parse_optional_int,
    safe_next_path,
)


BASE_DIR = Path(__file__).resolve().parent
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))
AUTH_COOKIE_NAME = "inss_auth"
BASE_SEGMENT_OPTIONS = [
    ("INSS", "INSS"),
    ("CREFAZ", "Crefaz"),
    ("GOVERNO", "Governo"),
    ("PREFEITURA", "Prefeitura"),
]
CLIENTS_PAGE_SIZE = 25
HISTORY_PAGE_SIZE = 20
request_id_ctx: ContextVar[str] = ContextVar("request_id", default="-")
HOME_DASHBOARD_CACHE_TTL_SECONDS = 90
home_dashboard_cache_lock = threading.Lock()
home_dashboard_cache: dict[str, object] = {"expires_at": None, "context": None}


class JsonFormatter(logging.Formatter):
    def format(self, record: logging.LogRecord) -> str:
        payload = {
            "timestamp": datetime.now(UTC).isoformat(),
            "level": record.levelname,
            "event": getattr(record, "event", record.getMessage()),
            "request_id": getattr(record, "request_id", request_id_ctx.get("-")),
            "user": getattr(record, "user", "-"),
            "message": record.getMessage(),
        }
        if getattr(record, "extra_data", None):
            payload["data"] = record.extra_data
        if record.exc_info:
            payload["error"] = self.formatException(record.exc_info)
        return json.dumps(payload, ensure_ascii=False, default=str)


def configure_json_logging() -> logging.Logger:
    logger_instance = logging.getLogger("inss_app")
    if logger_instance.handlers:
        return logger_instance
    logger_instance.setLevel(logging.INFO)
    handler = logging.StreamHandler()
    handler.setFormatter(JsonFormatter())
    logger_instance.addHandler(handler)
    logger_instance.propagate = False
    return logger_instance


logger = configure_json_logging()


def log_event(level: int, event: str, message: str, *, user: str = "-", request_id: str | None = None, data: dict[str, object] | None = None, exc_info=None) -> None:
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


templates.env.filters["brl"] = format_money
templates.env.filters["ref_br"] = format_reference
templates.env.filters["phone_br"] = format_phone


def get_session() -> Session:
    session = SessionLocal()
    try:
        yield session
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
    return enqueue_export_job_service(
        session=session,
        requested_by=requested_by,
        job_type=job_type,
        filters=filters,
        file_format=file_format,
        include_audit=include_audit,
        message=message,
        metadata=metadata,
        request_id=request_id,
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
        log_audit(
            session,
            job.requested_by,
            "export_job_completed",
            target_type="export_job",
            target_id=str(job.id),
            message="Exportacao assincrona concluida.",
            metadata={"arquivo": export_path.name, "registros": job.total_records, "job_type": job.job_type},
        )
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


def render_template(request: Request, template_name: str, context: dict[str, object], status_code: int = 200) -> HTMLResponse:
    path = str(getattr(request.url, "path", "") or "")
    if path == "/":
        active_nav = "home"
    elif path.startswith("/import"):
        active_nav = "import"
    elif path.startswith("/clients"):
        active_nav = "clients"
    elif path.startswith("/c6"):
        active_nav = "c6"
    elif path.startswith("/exports") or path.startswith("/downloads"):
        active_nav = "exports"
    elif path.startswith("/history"):
        active_nav = "history"
    elif path.startswith("/users"):
        active_nav = "users"
    else:
        active_nav = ""

    payload = {
        "current_user": getattr(request.state, "current_user", None),
        "permissions": build_permissions(getattr(request.state, "current_user", None)),
        "base_segment_options": BASE_SEGMENT_OPTIONS,
        "active_nav": active_nav,
    }
    payload.update(context)
    return templates.TemplateResponse(request, template_name, payload, status_code=status_code)


def build_c6_user_friendly_error(error: object) -> str:
    raw = str(error or "").strip()
    normalized = raw.lower()
    if any(term in normalized for term in ("enrollment", "beneficio", "benefício", "deve conter 10 caracteres")):
        return "Numero do beneficio invalido. Informe exatamente 10 digitos."
    if any(term in normalized for term in ("margem", "insufficient", "insuficiente", "sem margem", "saldo indisponivel")):
        return "Cliente sem margem disponivel para esta simulacao."
    if any(term in normalized for term in ("cpf", "tax_identifier", "documento invalido")):
        return "CPF invalido. Confira o numero informado."
    if any(term in normalized for term in ("birth_date", "data de nascimento", "nascimento invalido")):
        return "Data de nascimento invalida. Use o formato correto."
    if any(term in normalized for term in ("timeout", "timed out", "temporarily unavailable", "service unavailable", "502", "503", "504")):
        return "Servico de consignado indisponivel no momento. Tente novamente em instantes."
    return "Nao foi possivel concluir a operacao de consignado. Revise os dados e tente novamente."


def base_segment_label(value: str) -> str:
    normalized = normalize_base_segment(value)
    return dict(BASE_SEGMENT_OPTIONS).get(normalized, normalized)


def _norm_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip().replace("\ufeff", "")


def _parse_extras_json(payload: str) -> dict[str, object]:
    if not payload:
        return {}
    try:
        data = json.loads(payload)
    except json.JSONDecodeError:
        return {}
    return data if isinstance(data, dict) else {}


def _first_extra_value(extras: dict[str, object], *keys: str) -> str:
    for key in keys:
        value = extras.get(key)
        if _norm_text(value):
            return _norm_text(value)
    return ""


def _build_public_filter_options(session: Session, base_segment: str, base_source: str) -> dict[str, list[str]]:
    options = {
        "entidades": [],
        "convenios": [],
        "servicos": [],
        "situacoes": [],
        "cargos": [],
        "tipos_vinculo": [],
    }
    if base_segment not in {"GOVERNO", "PREFEITURA"} or not _norm_text(base_source):
        return options

    rows = session.execute(
        select(
            ClientOccurrence.entidade,
            ClientOccurrence.secretaria,
            ClientOccurrence.convenio,
            ClientOccurrence.servico,
            ClientOccurrence.situacao,
            ClientOccurrence.cbo_titulo,
            ClientOccurrence.extras_json,
        ).where(
            ClientOccurrence.base_segment == base_segment,
            ClientOccurrence.base_source == base_source,
        )
    ).all()

    entidades: set[str] = set()
    convenios: set[str] = set()
    servicos: set[str] = set()
    situacoes: set[str] = set()
    cargos: set[str] = set()
    tipos_vinculo: set[str] = set()

    for entidade, secretaria, convenio, servico, situacao, cbo_titulo, extras_json in rows:
        if _norm_text(entidade):
            entidades.add(_norm_text(entidade))
        if _norm_text(secretaria):
            entidades.add(_norm_text(secretaria))
        if _norm_text(convenio):
            convenios.add(_norm_text(convenio))
        if _norm_text(servico):
            servicos.add(_norm_text(servico))
        if _norm_text(situacao):
            situacoes.add(_norm_text(situacao))
        if _norm_text(cbo_titulo):
            cargos.add(_norm_text(cbo_titulo))

        extras = _parse_extras_json(extras_json or "")
        entidade_extra = _first_extra_value(
            extras,
            "ENTIDADE",
            "ORGAO",
            "ORGÃƒO",
            "SECRETARIA",
            "DESCRIÃ‡AO CNPJ",
            "DESCRIÃ‡ÃƒO CNPJ",
            "DESCRICAO CNPJ",
        )
        convenio_extra = _first_extra_value(extras, "CONVENIO", "CONVÃŠNIO")
        servico_extra = _first_extra_value(
            extras,
            "SERVICO",
            "SERVIÃ‡O",
            "SERVICO (SERVIDOR)",
            "SERVIÃ‡O (SERVIDOR)",
            "TIPO SERVICO (SERVIDOR)",
            "TIPO SERVIÃ‡O (SERVIDOR)",
        )
        situacao_extra = _first_extra_value(extras, "SITUACAO", "SITUAÃ‡ÃƒO", "STATUS")
        cargo_extra = _first_extra_value(extras, "CARGO", "FUNCAO", "FUNÃ‡ÃƒO", "FUNÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢O", "CARGO/FUNCAO", "CARGO FUNCAO", "CBO TITULO", "CBO_TITULO")
        tipo_vinculo = _first_extra_value(extras, "REGIME_CONTRATACAO", "TIPO DE VINCULO", "TIPO_VINCULO", "VINCULO", "REGIME DE CONTRATACAO")
        if entidade_extra:
            entidades.add(entidade_extra)
        if convenio_extra:
            convenios.add(convenio_extra)
        if servico_extra:
            servicos.add(servico_extra)
        if situacao_extra:
            situacoes.add(situacao_extra)
        if cargo_extra:
            cargos.add(cargo_extra)
        if tipo_vinculo:
            tipos_vinculo.add(tipo_vinculo)

    options["entidades"] = sorted(entidades)
    options["convenios"] = sorted(convenios)
    options["servicos"] = sorted(servicos)
    options["situacoes"] = sorted(situacoes)
    options["cargos"] = sorted(cargos)
    options["tipos_vinculo"] = sorted(tipos_vinculo)
    return options


def ensure_default_admin() -> None:
    session = SessionLocal()
    try:
        existing_user = session.scalar(select(func.count()).select_from(AppUser)) or 0
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


def get_authenticated_user(request: Request, session: Session) -> AppUser:
    user_id = parse_auth_token(request.cookies.get(AUTH_COOKIE_NAME, ""))
    if not user_id:
        raise HTTPException(status_code=401)
    user = session.get(AppUser, user_id)
    if user is None or not user.is_active:
        raise HTTPException(status_code=401)
    if is_future_timestamp(user.locked_until):
        raise HTTPException(status_code=403)
    return user


def require_permission(permission: str):
    def dependency(request: Request, session: Session = Depends(get_session)) -> AppUser:
        user = get_authenticated_user(request, session)
        if not user_can(user, permission):
            raise HTTPException(status_code=403)
        return user

    return dependency


@asynccontextmanager
async def lifespan(_: FastAPI):
    run_startup(log_event)
    yield


app = FastAPI(title="Banco de Clientes INSS", lifespan=lifespan)
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")


@app.middleware("http")
async def attach_current_user(request: Request, call_next):
    request_id = (request.headers.get("x-request-id") or "").strip() or str(uuid.uuid4())
    request_id_ctx.set(request_id)
    request.state.request_id = request_id
    request.state.current_user = None
    user_id = parse_auth_token(request.cookies.get(AUTH_COOKIE_NAME, ""))
    if user_id:
        session = SessionLocal()
        try:
            user = session.get(AppUser, user_id)
            if user is not None and user.is_active:
                request.state.current_user = user
        finally:
            session.close()
    started_at = datetime.now(UTC)
    response = await call_next(request)
    elapsed_ms = int((datetime.now(UTC) - started_at).total_seconds() * 1000)
    user_name = getattr(getattr(request.state, "current_user", None), "username", "-")
    response.headers["X-Request-ID"] = request_id
    log_event(
        logging.INFO,
        "http_request_completed",
        f"{request.method} {request.url.path}",
        user=user_name,
        request_id=request_id,
        data={"method": request.method, "path": request.url.path, "status_code": response.status_code, "elapsed_ms": elapsed_ms},
    )
    return response


@app.exception_handler(401)
async def unauthorized_handler(request: Request, _exc: HTTPException):
    target = quote_plus(str(request.url.path))
    return RedirectResponse(url=f"/login?next={target}", status_code=303)


@app.exception_handler(403)
async def forbidden_handler(request: Request, _exc: HTTPException):
    return render_template(
        request,
        "error.html",
        {"title": "Acesso negado", "message": "Voce nao tem permissao para acessar esta funcionalidade."},
        status_code=403,
    )


@app.exception_handler(Exception)
async def unhandled_exception_handler(request: Request, exc: Exception):
    user_name = getattr(getattr(request.state, "current_user", None), "username", "-")
    request_id = getattr(request.state, "request_id", request_id_ctx.get("-"))
    log_event(
        logging.ERROR,
        "unhandled_exception",
        "Erro nao tratado durante processamento da requisicao.",
        user=user_name,
        request_id=request_id,
        data={"method": request.method, "path": request.url.path},
        exc_info=True,
    )
    return JSONResponse(
        {"detail": "Erro interno do servidor.", "request_id": request_id},
        status_code=500,
    )


def login_page(request: Request, next: str = "/clients") -> HTMLResponse:
    if getattr(request.state, "current_user", None):
        return RedirectResponse(url=safe_next_path(next), status_code=303)
    return render_template(request, "login.html", {"error": "", "next": safe_next_path(next), "title": "Login"})


def login_submit(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    next: str = Form("/clients"),
    session: Session = Depends(get_session),
) -> HTMLResponse:
    user = session.scalar(select(AppUser).where(AppUser.username == username.strip()))
    redirect_target = safe_next_path(next)
    if user is not None and is_future_timestamp(user.locked_until):
        return render_template(
            request,
            "login.html",
            {"error": "Usuario temporariamente bloqueado.", "next": redirect_target, "title": "Login"},
            status_code=403,
        )
    if user is None or not user.is_active or not verify_password(password, user.password_hash, user.password_salt):
        return render_template(
            request,
            "login.html",
            {"error": "Usuario ou senha invalidos.", "next": redirect_target, "title": "Login"},
            status_code=400,
        )
    user.last_login_at = current_utc()
    log_audit(session, user.username, "login_success", target_type="user", target_id=str(user.id), message="Login efetuado com sucesso.")
    session.commit()
    response = RedirectResponse(url=redirect_target, status_code=303)
    response.set_cookie(AUTH_COOKIE_NAME, create_auth_token(user.id), httponly=True, samesite="lax")
    return response


def logout() -> RedirectResponse:
    response = RedirectResponse(url="/login", status_code=303)
    response.delete_cookie(AUTH_COOKIE_NAME)
    return response


def c6_entry(_current_user: AppUser = Depends(require_permission("can_access_consignado_inss"))) -> RedirectResponse:
    return RedirectResponse(url="/c6-inss", status_code=302)


def c6_inss_page(
    request: Request,
    _current_user: AppUser = Depends(require_permission("can_access_consignado_inss")),
) -> HTMLResponse:
    return render_template(
        request,
        "c6_inss.html",
        {
            "title": "Consignado Tradicional INSS C6",
            "result": None,
            "error": "",
            "active_action": "simulation",
            "tax_identifier": "",
            "enrollment": "",
            "birth_date": "",
            "income_amount": "4000",
            "requested_amount": "3000",
            "installment_amount": "",
            "installment_quantity": "84",
            "simulation_type": "POR_VALOR_SOLICITADO",
            "operation_type": "NOVA",
            "product_type_code": "0001",
            "formalization_subtype": "DIGITAL_WEB",
            "promoter_code": "003238",
            "covenant_group": "INSS",
            "public_agency": "000001",
            "bank_code": "",
            "agency_number": "",
            "agency_digit": "",
            "account_type": "01",
            "account_number": "",
            "account_digit": "",
            "account_holder_name": "",
            "account_holder_tax_identifier": "",
        },
    )


def c6_inss_submit(
    request: Request,
    tax_identifier: str = Form(""),
    enrollment: str = Form(""),
    birth_date: str = Form(""),
    income_amount: str = Form(""),
    requested_amount: str = Form(""),
    installment_amount: str = Form(""),
    installment_quantity: str = Form(""),
    simulation_type: str = Form("POR_VALOR_SOLICITADO"),
    operation_type: str = Form("NOVA"),
    product_type_code: str = Form("0001"),
    formalization_subtype: str = Form("DIGITAL_WEB"),
    promoter_code: str = Form("003238"),
    covenant_group: str = Form("INSS"),
    public_agency: str = Form("000001"),
    current_user: AppUser = Depends(require_permission("can_access_consignado_inss")),
    session: Session = Depends(get_session),
) -> HTMLResponse:
    def to_float(value: str, field_name: str) -> float:
        raw = (value or "").strip()
        if not raw:
            return 0.0
        if "," in raw and "." in raw:
            if raw.rfind(",") > raw.rfind("."):
                raw = raw.replace(".", "").replace(",", ".")
            else:
                raw = raw.replace(",", "")
        elif "," in raw:
            raw = raw.replace(".", "").replace(",", ".")
        try:
            return float(raw)
        except ValueError as exc:
            raise ValueError(f"Valor invalido para {field_name}: {value}") from exc

    def to_int(value: str) -> int:
        raw = "".join(ch for ch in (value or "") if ch.isdigit())
        return int(raw or "0")

    def to_enrollment(value: str) -> str:
        digits = "".join(ch for ch in (value or "") if ch.isdigit())
        if len(digits) != 10:
            raise ValueError("Numero do beneficio invalido. Informe exatamente 10 digitos.")
        return digits

    def to_birth_date(value: str) -> str:
        raw = (value or "").strip()
        if not raw:
            raise ValueError("Data de nascimento obrigatoria.")
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", raw):
            normalized = raw
        elif re.fullmatch(r"\d{2}/\d{2}/\d{4}", raw):
            day, month, year = raw.split("/")
            normalized = f"{year}-{month}-{day}"
        else:
            raise ValueError("Data de nascimento invalida. Use dd/mm/aaaa ou aaaa-mm-dd.")
        parsed = datetime.strptime(normalized, "%Y-%m-%d")
        if parsed.date() > datetime.now().date():
            raise ValueError("Data de nascimento nao pode ser futura.")
        return normalized

    def to_money_br(value: object) -> str:
        number = float(value or 0)
        return f"{number:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    error = ""
    api_status_note = ""
    result: dict[str, object] | None = None
    simulation_conditions: list[dict[str, object]] = []
    continue_url = ""
    best_offer: dict[str, object] | None = None
    payload: dict[str, object] = {}
    try:
        simulation_type_normalized = simulation_type.strip().upper()
        payload = {
            "operation_type": operation_type.strip(),
            "product_type_code": product_type_code.strip(),
            "simulation_type": simulation_type_normalized,
            "formalization_subtype": formalization_subtype.strip(),
            "promoter_code": promoter_code.strip().zfill(6),
            "covenant_group": covenant_group.strip(),
            "public_agency": public_agency.strip(),
            "installment_quantity": to_int(installment_quantity),
                "client": {
                    "tax_identifier": tax_identifier.strip(),
                    "enrollment": to_enrollment(enrollment),
                    "birth_date": to_birth_date(birth_date),
                    "income_amount": to_float(income_amount, "income_amount"),
                },
        }
        if simulation_type_normalized == "POR_VALOR_PARCELA":
            payload["installment_amount"] = to_float(installment_amount, "installment_amount")
        else:
            payload["requested_amount"] = to_float(requested_amount, "requested_amount")
    except ValueError as exc:
        error = str(exc)

    if not error:
        try:
            response = c6_worker_loan_client.simulate_inss_proposal(payload)
            result = {"status_code": response.status_code, "data": response.payload, "action": "simulation", "payload_enviado": payload}
            if response.status_code == 200:
                api_status_note = "Consulta concluida com sucesso."
            elif response.status_code in (400, 422):
                api_status_note = "API recebeu a requisicao, mas retornou validacao de dados."
            elif response.status_code in (401, 403):
                api_status_note = "API recusou autenticacao/autorizacao."
            else:
                api_status_note = f"API retornou HTTP {response.status_code}."
            for item in (response.payload or {}).get("credit_conditions", []):
                covenant = item.get("covenant", {}) if isinstance(item, dict) else {}
                observation = covenant.get("observation", "")
                installment_value = item.get("installment_amount", 0) if isinstance(item, dict) else 0
                client_amount_value = item.get("client_amount", 0) if isinstance(item, dict) else 0
                monthly_rate_value = item.get("monthly_customer_rate", 0) if isinstance(item, dict) else 0
                row = {
                    "covenant_code": covenant.get("code", ""),
                    "covenant_description": covenant.get("description", ""),
                    "observation": observation,
                    "requested_amount": item.get("requested_amount", 0),
                    "installment_amount": installment_value,
                    "installment_quantity": item.get("installment_quantity", 0),
                    "client_amount": client_amount_value,
                    "monthly_customer_rate": monthly_rate_value,
                    "product_code": (item.get("product", {}) or {}).get("code", "") if isinstance(item, dict) else "",
                    "product_description": (item.get("product", {}) or {}).get("description", "") if isinstance(item, dict) else "",
                    "is_valid": not observation and float(item.get("requested_amount", 0) or 0) > 0 and float(installment_value or 0) > 0,
                }
                if row["is_valid"]:
                    requested_amount_for_include = row["requested_amount"]
                    if simulation_type_normalized == "POR_VALOR_PARCELA":
                        requested_amount_for_include = row["client_amount"] or row["requested_amount"]
                    row["select_url"] = "/c6-inss/include?" + urlencode(
                        {
                            "tax_identifier": tax_identifier,
                            "enrollment": enrollment,
                            "birth_date": birth_date,
                            "income_amount": income_amount,
                            "requested_amount": to_money_br(requested_amount_for_include),
                            "installment_amount": installment_value,
                            "installment_quantity": str(row["installment_quantity"]),
                            "simulation_type": simulation_type,
                            "operation_type": operation_type,
                            "product_type_code": product_type_code,
                            "formalization_subtype": formalization_subtype,
                            "promoter_code": promoter_code,
                            "covenant_group": covenant_group,
                            "public_agency": public_agency,
                            "selected_covenant_code": row["covenant_code"],
                            "selected_covenant_description": row["covenant_description"],
                            "selected_product_code": row["product_code"],
                            "selected_product_description": row["product_description"],
                            "selected_installment_amount": str(row["installment_amount"]),
                            "selected_monthly_rate": str(row["monthly_customer_rate"]),
                        }
                    )
                simulation_conditions.append(row)
            valid_rows = [row for row in simulation_conditions if row["is_valid"]]
            invalid_rows = [row for row in simulation_conditions if not row["is_valid"]]
            valid_rows.sort(
                key=lambda row: (
                    -float(row.get("client_amount") or 0),
                    float(row.get("monthly_customer_rate") or 999),
                    float(row.get("installment_amount") or 999999),
                )
            )
            if valid_rows:
                best_offer = valid_rows[0]
                best_offer["highlight_badge"] = "Maior valor"
                best_rate_row = min(valid_rows, key=lambda row: float(row.get("monthly_customer_rate") or 999))
                if best_rate_row is not best_offer:
                    best_rate_row["highlight_badge"] = "Menor taxa"
                else:
                    best_offer["highlight_badge"] = "Maior valor e menor taxa"
            simulation_conditions = valid_rows + invalid_rows
            requested_amount_for_continue = requested_amount
            if simulation_type_normalized == "POR_VALOR_PARCELA" and best_offer is not None:
                requested_amount_for_continue = to_money_br(best_offer.get("client_amount") or requested_amount)
            continue_url = "/c6-inss/include?" + urlencode(
                {
                    "tax_identifier": tax_identifier,
                    "enrollment": enrollment,
                    "birth_date": birth_date,
                    "income_amount": income_amount,
                    "requested_amount": requested_amount_for_continue,
                    "installment_amount": installment_amount,
                    "installment_quantity": installment_quantity,
                    "simulation_type": simulation_type,
                    "operation_type": operation_type,
                    "product_type_code": product_type_code,
                    "formalization_subtype": formalization_subtype,
                    "promoter_code": promoter_code,
                    "covenant_group": covenant_group,
                    "public_agency": public_agency,
                }
            )
            log_audit(
                session,
                current_user.username,
                "c6_inss_simulation_requested",
                target_type="integration",
                message="Operacao C6 INSS executada pela tela de consignado tradicional.",
                metadata={"path": str(request.url.path), "action": "simulation"},
            )
            session.commit()
        except C6WorkerLoanError as exc:
            error = build_c6_user_friendly_error(exc)
            api_status_note = "Falha ao consultar API de consignado."
            log_audit(
                session,
                current_user.username,
                "c6_inss_simulation_failed",
                target_type="integration",
                message=error,
                metadata={"path": str(request.url.path), "action": "simulation", "technical_error": str(exc), "payload_enviado": payload},
            )
            session.commit()

    return render_template(
        request,
        "c6_inss.html",
        {
            "title": "Consignado Tradicional INSS C6",
            "result": result,
            "error": error,
            "api_status_note": api_status_note,
            "simulation_conditions": simulation_conditions,
            "best_offer": best_offer,
            "continue_url": continue_url,
            "tax_identifier": tax_identifier,
            "enrollment": enrollment,
            "birth_date": birth_date,
            "income_amount": income_amount,
            "requested_amount": requested_amount,
            "installment_amount": installment_amount,
            "installment_quantity": installment_quantity,
            "simulation_type": simulation_type,
            "operation_type": operation_type,
            "product_type_code": product_type_code,
            "formalization_subtype": formalization_subtype,
            "promoter_code": promoter_code.strip().zfill(6),
            "covenant_group": covenant_group,
            "public_agency": public_agency,
        },
        status_code=400 if error else 200,
    )


def c6_latest_inss_client_by_cpf(
    cpf: str,
    _current_user: AppUser = Depends(require_permission("can_access_consignado_inss")),
    session: Session = Depends(get_session),
) -> JSONResponse:
    normalized_cpf = normalize_cpf(cpf)
    if len(normalized_cpf) != 11:
        raise HTTPException(status_code=422, detail="CPF invalido.")

    latest_occurrence = session.execute(
        select(ClientOccurrence)
        .where(
            ClientOccurrence.cpf == normalized_cpf,
            ClientOccurrence.base_segment == "INSS",
        )
        .order_by(ClientOccurrence.criado_em.desc(), ClientOccurrence.id.desc())
        .limit(1)
    ).scalar_one_or_none()

    if latest_occurrence is None:
        raise HTTPException(status_code=404, detail="Nenhuma ocorrencia INSS encontrada para este CPF.")

    birth_date = (latest_occurrence.dt_nasc or "").strip()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", birth_date):
        year, month, day = birth_date.split("-")
        birth_date = f"{day}/{month}/{year}"

    benefit = (latest_occurrence.nu_nb or latest_occurrence.matricula or "").strip()
    income = latest_occurrence.vl_rmi if latest_occurrence.vl_rmi is not None else latest_occurrence.salario

    return JSONResponse(
        {
            "cpf": normalized_cpf,
            "benefit": benefit,
            "birth_date": birth_date,
            "income_amount": float(income) if income is not None else 0.0,
            "updated_at": latest_occurrence.criado_em.isoformat() if latest_occurrence.criado_em else "",
        }
    )


def c6_inss_include_page(
    request: Request,
    _current_user: AppUser = Depends(require_permission("can_access_consignado_inss")),
) -> HTMLResponse:
    q = request.query_params
    return render_template(
        request,
        "c6_inss_include.html",
        {
            "title": "Inclusao Proposta INSS C6",
            "result": None,
            "error": "",
            "tax_identifier": q.get("tax_identifier", ""),
            "enrollment": q.get("enrollment", ""),
            "birth_date": q.get("birth_date", ""),
            "income_amount": q.get("income_amount", ""),
            "requested_amount": q.get("requested_amount", ""),
            "installment_quantity": q.get("installment_quantity", ""),
            "simulation_type": q.get("simulation_type", "POR_VALOR_SOLICITADO"),
            "operation_type": q.get("operation_type", "NOVA"),
            "product_type_code": q.get("product_type_code", "0001"),
            "formalization_subtype": q.get("formalization_subtype", "DIGITAL_WEB"),
            "promoter_code": q.get("promoter_code", "003238"),
            "covenant_group": q.get("covenant_group", "INSS"),
            "public_agency": q.get("public_agency", "000001"),
            "bank_code": "",
            "agency_number": "",
            "agency_digit": "",
            "account_type": "01",
            "account_number": "",
            "account_digit": "",
            "account_holder_name": "",
            "account_holder_tax_identifier": "",
            "selected_covenant_code": q.get("selected_covenant_code", ""),
            "selected_covenant_description": q.get("selected_covenant_description", ""),
            "selected_product_code": q.get("selected_product_code", ""),
            "selected_product_description": q.get("selected_product_description", ""),
            "selected_installment_amount": q.get("selected_installment_amount", ""),
            "selected_monthly_rate": q.get("selected_monthly_rate", ""),
        },
    )


def c6_inss_include_submit(
    request: Request,
    tax_identifier: str = Form(""),
    enrollment: str = Form(""),
    birth_date: str = Form(""),
    income_amount: str = Form(""),
    requested_amount: str = Form(""),
    installment_quantity: str = Form(""),
    simulation_type: str = Form("POR_VALOR_SOLICITADO"),
    operation_type: str = Form("NOVA"),
    product_type_code: str = Form("0001"),
    formalization_subtype: str = Form("DIGITAL_WEB"),
    promoter_code: str = Form("003238"),
    covenant_group: str = Form("INSS"),
    public_agency: str = Form("000001"),
    bank_code: str = Form(""),
    agency_number: str = Form(""),
    agency_digit: str = Form(""),
    account_type: str = Form("01"),
    account_number: str = Form(""),
    account_digit: str = Form(""),
    account_holder_name: str = Form(""),
    account_holder_tax_identifier: str = Form(""),
    selected_covenant_code: str = Form(""),
    selected_covenant_description: str = Form(""),
    selected_product_code: str = Form(""),
    selected_product_description: str = Form(""),
    selected_installment_amount: str = Form(""),
    selected_monthly_rate: str = Form(""),
    current_user: AppUser = Depends(require_permission("can_access_consignado_inss")),
    session: Session = Depends(get_session),
) -> HTMLResponse:
    try:
        payload = {
            "operation_type": operation_type.strip(),
            "product_type_code": product_type_code.strip(),
            "simulation_type": simulation_type.strip(),
            "formalization_subtype": formalization_subtype.strip(),
            "promoter_code": promoter_code.strip().zfill(6),
            "covenant_group": covenant_group.strip(),
            "public_agency": public_agency.strip(),
            "requested_amount": float(str(requested_amount).replace(".", "").replace(",", ".")),
            "installment_quantity": int("".join(ch for ch in str(installment_quantity) if ch.isdigit()) or "0"),
            "client": {
                "tax_identifier": tax_identifier.strip(),
                "enrollment": enrollment.strip(),
                "birth_date": datetime.strptime(birth_date.strip(), "%d/%m/%Y").strftime("%Y-%m-%d")
                if "/" in birth_date
                else birth_date.strip(),
                "income_amount": float(str(income_amount).replace(".", "").replace(",", ".")),
            },
            "banking_data": {
                "bank_code": bank_code.strip(),
                "agency_number": agency_number.strip(),
                "agency_digit": agency_digit.strip(),
                "account_type": account_type.strip(),
                "account_number": account_number.strip(),
                "account_digit": account_digit.strip(),
                "account_holder_name": account_holder_name.strip(),
                "account_holder_tax_identifier": account_holder_tax_identifier.strip() or tax_identifier.strip(),
            },
        }
        response = c6_worker_loan_client.include_inss_proposal(payload)
        log_audit(
            session,
            current_user.username,
            "c6_inss_inclusion_requested",
            target_type="integration",
            message="Inclusao C6 INSS executada.",
            metadata={"path": str(request.url.path), "action": "include"},
        )
        session.commit()
        result = {"status_code": response.status_code, "data": response.payload, "action": "include", "payload_enviado": payload}
        error = ""
    except Exception as exc:
        result = None
        error = build_c6_user_friendly_error(exc)
        log_audit(
            session,
            current_user.username,
            "c6_inss_inclusion_failed",
            target_type="integration",
            message=error,
            metadata={"path": str(request.url.path), "action": "include", "technical_error": str(exc)},
        )
        session.commit()

    return render_template(
        request,
        "c6_inss_include.html",
        {
            "title": "Inclusao Proposta INSS C6",
            "result": result,
            "error": error,
            "tax_identifier": tax_identifier,
            "enrollment": enrollment,
            "birth_date": birth_date,
            "income_amount": income_amount,
            "requested_amount": requested_amount,
            "installment_quantity": installment_quantity,
            "simulation_type": simulation_type,
            "operation_type": operation_type,
            "product_type_code": product_type_code,
            "formalization_subtype": formalization_subtype,
            "promoter_code": promoter_code,
            "covenant_group": covenant_group,
            "public_agency": public_agency,
            "bank_code": bank_code,
            "agency_number": agency_number,
            "agency_digit": agency_digit,
            "account_type": account_type,
            "account_number": account_number,
            "account_digit": account_digit,
            "account_holder_name": account_holder_name,
            "account_holder_tax_identifier": account_holder_tax_identifier,
            "selected_covenant_code": selected_covenant_code,
            "selected_covenant_description": selected_covenant_description,
            "selected_product_code": selected_product_code,
            "selected_product_description": selected_product_description,
            "selected_installment_amount": selected_installment_amount,
            "selected_monthly_rate": selected_monthly_rate,
        },
        status_code=400 if error else 200,
    )


def home(
    request: Request,
    current_user: AppUser = Depends(require_permission("can_view_dashboard")),
    session: Session = Depends(get_session),
) -> HTMLResponse:
    now_utc = current_utc()
    with home_dashboard_cache_lock:
        cached_expires_at = home_dashboard_cache.get("expires_at")
        cached_context = home_dashboard_cache.get("context")
        if isinstance(cached_expires_at, datetime) and cached_expires_at > now_utc and isinstance(cached_context, dict):
            return render_template(
                request,
                "home.html",
                dict(cached_context),
            )

    batch_stats = session.execute(
        select(
            func.count(ImportBatch.id),
            func.sum(case((ImportBatch.status == "CONCLUIDO", 1), else_=0)),
        )
    ).one()
    total_batches = int(batch_stats[0] or 0)
    completed_batches = int(batch_stats[1] or 0)

    total_files = int(session.scalar(select(func.count(SourceFile.id))) or 0)
    total_occurrences = int(session.scalar(select(func.count(ClientOccurrence.id))) or 0)

    client_stats = session.execute(
        select(
            func.count(Client.cpf),
            func.sum(case((Client.tem_telefone.is_(True), 1), else_=0)),
            func.sum(case((Client.do_not_call.is_(True), 1), else_=0)),
            func.sum(case((Client.tem_telefone.is_(True), case((Client.do_not_call.is_(False), 1), else_=0)), else_=0)),
        )
    ).one()
    total_clients = int(client_stats[0] or 0)
    total_with_phone = int(client_stats[1] or 0)
    total_do_not_call = int(client_stats[2] or 0)
    total_actionable = int(client_stats[3] or 0)

    latest_exports = session.execute(select(ExportHistory).order_by(ExportHistory.criado_em.desc()).limit(5)).scalars().all()
    latest_audits = session.execute(select(AuditLog).order_by(AuditLog.created_at.desc()).limit(5)).scalars().all()
    has_phone_occurrence = or_(
        ClientOccurrence.telefone1 != "",
        ClientOccurrence.telefone2 != "",
        ClientOccurrence.telefone3 != "",
    )
    segment_rows = session.execute(
        select(
            ClientOccurrence.base_segment,
            func.count(ClientOccurrence.id),
            func.count(func.distinct(ClientOccurrence.cpf)),
            func.count(func.distinct(case((has_phone_occurrence, ClientOccurrence.cpf), else_=None))),
        )
        .where(ClientOccurrence.base_segment.in_([segment for segment, _ in BASE_SEGMENT_OPTIONS]))
        .group_by(ClientOccurrence.base_segment)
    ).all()
    segment_map = {
        row[0]: {"occurrences": int(row[1] or 0), "clients": int(row[2] or 0), "with_phone": int(row[3] or 0)}
        for row in segment_rows
    }
    segment_stats = []
    for segment, _label in BASE_SEGMENT_OPTIONS:
        stats = segment_map.get(segment, {"occurrences": 0, "clients": 0, "with_phone": 0})
        segment_stats.append(
            {
                "segment": segment,
                "label": base_segment_label(segment),
                "occurrences": stats["occurrences"],
                "clients": stats["clients"],
                "with_phone": stats["with_phone"],
            }
        )

    top_states = session.execute(
        select(Client.uf_atual, func.count())
        .where(Client.uf_atual != "")
        .group_by(Client.uf_atual)
        .order_by(func.count().desc())
        .limit(5)
    ).all()

    if total_clients == 0 and total_occurrences > 0:
        fallback_stats = session.execute(
            select(
                func.count(func.distinct(ClientOccurrence.cpf)),
                func.count(func.distinct(case((has_phone_occurrence, ClientOccurrence.cpf), else_=None))),
            )
        ).one()
        total_clients = int(fallback_stats[0] or 0)
        total_with_phone = int(fallback_stats[1] or 0)
        total_actionable = total_with_phone
        total_do_not_call = 0
        top_states = session.execute(
            select(ClientOccurrence.uf, func.count(func.distinct(ClientOccurrence.cpf)))
            .where(ClientOccurrence.uf != "")
            .group_by(ClientOccurrence.uf)
            .order_by(func.count(func.distinct(ClientOccurrence.cpf)).desc())
            .limit(5)
        ).all()

    latest_batches = session.execute(select(ImportBatch).order_by(ImportBatch.data_importacao.desc()).limit(4)).scalars().all()

    phone_ratio = round((total_with_phone / total_clients) * 100, 1) if total_clients else 0
    actionable_ratio = round((total_actionable / total_clients) * 100, 1) if total_clients else 0
    do_not_call_ratio = round((total_do_not_call / total_clients) * 100, 1) if total_clients else 0
    no_phone_ratio = round(100 - phone_ratio, 1) if total_clients else 0
    total_without_phone = max(total_clients - total_with_phone, 0)
    duplicate_records = max(total_occurrences - total_clients, 0)
    duplicate_ratio = round((duplicate_records / total_occurrences) * 100, 1) if total_occurrences else 0
    state_total = sum(item[1] for item in top_states) or 1
    states_chart = [
        {"uf": uf or "N/D", "count": count, "pct": round((count / state_total) * 100, 1)}
        for uf, count in top_states
    ]
    latest_batch = latest_batches[0] if latest_batches else None
    latest_export = latest_exports[0] if latest_exports else None
    top_state = states_chart[0] if states_chart else None
    avg_records_per_file = round(total_occurrences / total_files) if total_files else 0
    avg_clients_per_campaign = round(total_clients / total_batches) if total_batches else 0
    file_per_campaign = round(total_files / total_batches, 1) if total_batches else 0
    fresh_stats = session.execute(
        select(
            func.sum(case((Client.atualizado_em >= now_utc - timedelta(days=30), 1), else_=0)),
            func.sum(case((Client.atualizado_em >= now_utc - timedelta(days=60), 1), else_=0)),
            func.sum(case((Client.atualizado_em >= now_utc - timedelta(days=90), 1), else_=0)),
        )
        .select_from(Client)
    ).one()
    fresh_30 = int(fresh_stats[0] or 0)
    fresh_60 = int(fresh_stats[1] or 0)
    fresh_90 = int(fresh_stats[2] or 0)
    fresh_30_ratio = round((fresh_30 / total_clients) * 100, 1) if total_clients else 0
    fresh_60_ratio = round((fresh_60 / total_clients) * 100, 1) if total_clients else 0
    fresh_90_ratio = round((fresh_90 / total_clients) * 100, 1) if total_clients else 0
    actionable_states = session.execute(
        select(Client.uf_atual, func.count())
        .where(
            Client.uf_atual != "",
            Client.tem_telefone.is_(True),
            Client.do_not_call.is_(False),
        )
        .group_by(Client.uf_atual)
        .order_by(func.count().desc())
        .limit(5)
    ).all()
    actionable_states_total = sum(item[1] for item in actionable_states) or 1
    actionable_states_chart = [
        {"uf": uf or "N/D", "count": count, "pct": round((count / actionable_states_total) * 100, 1)}
        for uf, count in actionable_states
    ]
    avg_export_hours_from_import = 0.0
    if latest_exports and latest_batches:
        batch_times = sorted(
            [item.data_importacao for item in latest_batches if item.data_importacao is not None]
        )
        deltas_in_hours: list[float] = []
        for export in latest_exports:
            if export.criado_em is None:
                continue
            candidates = [timestamp for timestamp in batch_times if timestamp <= export.criado_em]
            if not candidates:
                continue
            nearest_batch = candidates[-1]
            delta_hours = (export.criado_em - nearest_batch).total_seconds() / 3600
            if delta_hours >= 0:
                deltas_in_hours.append(delta_hours)
        if deltas_in_hours:
            avg_export_hours_from_import = round(sum(deltas_in_hours) / len(deltas_in_hours), 1)
    dashboard_context = {
        "title": "Inicio",
        "total_batches": total_batches,
        "total_files": total_files,
        "total_occurrences": total_occurrences,
        "total_clients": total_clients,
        "total_with_phone": total_with_phone,
        "total_without_phone": total_without_phone,
        "phone_ratio": phone_ratio,
        "actionable_ratio": actionable_ratio,
        "do_not_call_ratio": do_not_call_ratio,
        "no_phone_ratio": no_phone_ratio,
        "total_actionable": total_actionable,
        "total_do_not_call": total_do_not_call,
        "duplicate_records": duplicate_records,
        "duplicate_ratio": duplicate_ratio,
        "fresh_30": fresh_30,
        "fresh_60": fresh_60,
        "fresh_90": fresh_90,
        "fresh_30_ratio": fresh_30_ratio,
        "fresh_60_ratio": fresh_60_ratio,
        "fresh_90_ratio": fresh_90_ratio,
        "actionable_states_chart": actionable_states_chart,
        "avg_export_hours_from_import": avg_export_hours_from_import,
        "completed_batches": completed_batches,
        "avg_records_per_file": avg_records_per_file,
        "avg_clients_per_campaign": avg_clients_per_campaign,
        "file_per_campaign": file_per_campaign,
        "latest_batch_label": format_dashboard_datetime(latest_batch.data_importacao) if latest_batch else "-",
        "latest_export_label": format_dashboard_datetime(latest_export.criado_em) if latest_export else "-",
        "top_state": top_state,
        "states_chart": states_chart,
        "latest_batches": latest_batches,
        "latest_exports": latest_exports,
        "latest_audits": latest_audits,
        "segment_stats": segment_stats,
    }
    with home_dashboard_cache_lock:
        home_dashboard_cache["context"] = dict(dashboard_context)
        home_dashboard_cache["expires_at"] = now_utc + timedelta(seconds=HOME_DASHBOARD_CACHE_TTL_SECONDS)
    return render_template(request, "home.html", dashboard_context)


def import_page(request: Request, _current_user: AppUser = Depends(require_permission("can_import"))) -> HTMLResponse:
    session = SessionLocal()
    try:
        available_public_sources = [
            value
            for (value,) in session.execute(
                select(ClientOccurrence.base_source)
                .where(
                    ClientOccurrence.base_segment.in_(["GOVERNO", "PREFEITURA"]),
                    ClientOccurrence.base_source != "",
                )
                .distinct()
                .order_by(ClientOccurrence.base_source.asc())
            ).all()
            if value
        ]
    finally:
        session.close()
    return render_template(
        request,
        "import.html",
        {
            "summary": None,
            "error": request.query_params.get("error", ""),
            "message": request.query_params.get("message", ""),
            "title": "Importacao",
            "selected_base_segment": request.query_params.get("base_segment", "INSS"),
            "available_public_sources": available_public_sources,
        },
    )


async def run_import(
    request: Request,
    year: str = Form(""),
    month: str = Form(""),
    base_segment: str = Form("INSS"),
    base_source: str = Form(""),
    user_name: str = Form("operador"),
    origin_folder: str = Form("upload_web"),
    allowed_species: list[str] = Form(default=[]),
    custom_species: str = Form(""),
    server_folder: str = Form(""),
    files: list[UploadFile] = File(default=[]),
    current_user: AppUser = Depends(require_permission("can_import")),
    session: Session = Depends(get_session),
) -> HTMLResponse:
    normalized_segment = normalize_base_segment(base_segment)
    now = datetime.now()
    if normalized_segment == "INSS":
        try:
            year_value = int(str(year or "").strip())
            month_value = int(str(month or "").strip())
        except ValueError:
            return render_template(
                request,
                "import.html",
                {"summary": None, "error": "Informe ano e mes validos para importar INSS.", "title": "Importacao", "selected_base_segment": normalized_segment},
                status_code=400,
            )
        if year_value < 2020 or year_value > 2100 or month_value < 1 or month_value > 12:
            return render_template(
                request,
                "import.html",
                {"summary": None, "error": "Ano/mes fora do intervalo permitido para importacao.", "title": "Importacao", "selected_base_segment": normalized_segment},
                status_code=400,
            )
    else:
        year_value = now.year
        month_value = now.month

    selected_paths: list[Path] = []
    if server_folder.strip():
        folder = Path(server_folder.strip())
        if not folder.exists() or not folder.is_dir():
            return render_template(request, "import.html", {"summary": None, "error": f"Pasta nao encontrada: {folder}", "title": "Importacao"}, status_code=400)
        selected_paths = [path for path in sorted(folder.iterdir()) if path.suffix.lower() in {".csv", ".xlsx", ".xlsm", ".xls"}]
        origin_folder = str(folder)
    elif files:
        selected_paths = save_uploaded_files([(item.filename or "arquivo.csv", await item.read()) for item in files])
    else:
        return render_template(request, "import.html", {"summary": None, "error": "Envie arquivos ou informe uma pasta do servidor.", "title": "Importacao"}, status_code=400)

    import_request = ImportRequest(
        year=year_value,
        month=month_value,
        user_name=current_user.username,
        origin_folder=origin_folder,
        files=selected_paths,
        base_segment=normalized_segment,
        base_source=base_source,
        allowed_species=[*allowed_species, custom_species] if normalized_segment == "INSS" else [],
    )
    batch = create_import_batch(session, import_request)
    log_audit(
        session,
        current_user.username,
        "import_batch_created",
        target_type="batch",
        target_id=str(batch.id),
        message="Importacao iniciada.",
        metadata={"segmento": normalized_segment, "fonte": base_source, "arquivos": [path.name for path in selected_paths]},
    )
    session.commit()
    current_request_id = getattr(request.state, "request_id", request_id_ctx.get("-"))
    log_event(
        logging.INFO,
        "import_batch_created",
        "Importacao colocada em fila.",
        user=current_user.username,
        request_id=current_request_id,
        data={"batch_id": batch.id, "segmento": normalized_segment, "files": len(selected_paths)},
    )
    enqueue_result = enqueue_import_job_service(batch.id, import_request, current_request_id)
    batch.queue_backend = enqueue_result.get("enqueue_mode", "thread")
    batch.queue_name = enqueue_result.get("queue_name", settings.queue_imports_name)
    batch.queue_job_id = enqueue_result.get("queue_job_id", "")
    session.commit()
    log_event(
        logging.INFO,
        "import_batch_enqueued",
        "Job de importacao enfileirado.",
        user=current_user.username,
        request_id=current_request_id,
        data={
            "batch_id": batch.id,
            "enqueue_mode": enqueue_result.get("enqueue_mode", "thread"),
            "queue_backend": settings.queue_backend,
            "queue_job_id": enqueue_result.get("queue_job_id", ""),
        },
    )
    return RedirectResponse(url=f"/history/{batch.id}", status_code=303)


async def run_public_update_import_from_import_page(
    base_segment: str = Form("GOVERNO"),
    source_name: str = Form(""),
    files: list[UploadFile] = File(default=[]),
    current_user: AppUser = Depends(require_permission("can_import")),
    session: Session = Depends(get_session),
) -> RedirectResponse:
    normalized_segment = normalize_base_segment(base_segment, default="GOVERNO")
    if normalized_segment not in {"GOVERNO", "PREFEITURA"}:
        return RedirectResponse(url="/import?error=" + quote_plus("A atualizacao complementar e exclusiva para GOV/PREF."), status_code=303)

    uploaded_files: list[tuple[str, bytes]] = []
    for item in files:
        if item.filename:
            uploaded_files.append((item.filename, await item.read()))
    if not uploaded_files:
        return RedirectResponse(
            url=f"/import?base_segment={normalized_segment}&error=" + quote_plus("Selecione ao menos um arquivo complementar."),
            status_code=303,
        )

    saved_paths = save_uploaded_files(uploaded_files)
    summary = import_public_campaign_updates(
        session,
        saved_paths,
        base_segment=normalized_segment,
        source_name=source_name,
        imported_by=current_user.username,
    )
    log_audit(
        session,
        current_user.username,
        "public_update_imported",
        target_type="public_update",
        target_id=normalized_segment,
        message="Atualizacao complementar GOV/PREF importada.",
        metadata={"segmento": normalized_segment, "fonte": source_name, "arquivos": [path.name for path in saved_paths], "clientes_atualizados": summary.clients_updated},
    )
    session.commit()
    message = (
        f"Atualizacao GOV/PREF importada: {summary.rows_imported} linha(s), "
        f"{summary.contracts_saved} contrato(s) gravado(s), "
        f"{summary.clients_updated} cliente(s) atualizados."
    )
    return RedirectResponse(
        url=f"/import?base_segment={normalized_segment}&message=" + quote_plus(message),
        status_code=303,
    )


def clients_page(
    request: Request,
    base_segment: str = "INSS",
    base_source: str = "",
    selected_matricula: str = "",
    selected_servico: str = "",
    page: int = 1,
    quick_mode: str = "cpf",
    quick_value: str = "",
    year: str = "",
    month: str = "",
    week: str = "",
    uf: str = "",
    city: str = "",
    esp: str = "",
    margin_min: str = "",
    margin_max: str = "",
    age_min: str = "",
    age_max: str = "",
    has_phone: str = "",
    source_file: str = "",
    cpf: str = "",
    phone: str = "",
    ddd: str = "",
    name: str = "",
    current_user: AppUser = Depends(require_permission("can_search")),
    session: Session = Depends(get_session),
) -> HTMLResponse:
    def has_value(value: object) -> bool:
        return bool(str(value or "").strip())

    safe_page = parse_optional_int(str(page)) or 1
    normalized_segment = normalize_base_segment(base_segment)
    filters = FilterSet(
        base_segment=normalized_segment,
        quick_mode="telefone" if quick_mode == "telefone" else "cpf",
        quick_value=quick_value,
        year=parse_optional_int(year),
        month=parse_optional_int(month),
        week=week,
        uf=uf,
        city=city,
        esp=esp,
        margin_min=parse_decimal(margin_min),
        margin_max=parse_decimal(margin_max),
        age_min=parse_optional_int(age_min),
        age_max=parse_optional_int(age_max),
        has_phone=True if has_phone == "sim" else False if has_phone == "nao" else None,
        source_file=source_file,
        base_source=base_source,
        cpf=cpf,
        phone=phone,
        ddd=ddd,
        name=name,
    )
    has_active_lookup = any(
        [
            bool(filters.quick_value.strip()),
            filters.year is not None,
            filters.month is not None,
            bool(filters.week.strip()),
            bool(filters.uf.strip()),
            bool(filters.city.strip()),
            bool(filters.esp.strip()),
            filters.margin_min is not None,
            filters.margin_max is not None,
            filters.age_min is not None,
            filters.age_max is not None,
            filters.has_phone is not None,
            bool(filters.source_file.strip()),
            bool(filters.base_source.strip()),
            bool(filters.cpf.strip()),
            bool(filters.phone.strip()),
            bool(filters.ddd.strip()),
            bool(filters.name.strip()),
            bool(selected_matricula.strip()),
            bool(selected_servico.strip()),
        ]
    )
    quick_only_lookup = bool(filters.quick_value.strip()) and not any(
        [
            filters.year is not None,
            filters.month is not None,
            bool(filters.week.strip()),
            bool(filters.uf.strip()),
            bool(filters.city.strip()),
            bool(filters.esp.strip()),
            filters.margin_min is not None,
            filters.margin_max is not None,
            filters.age_min is not None,
            filters.age_max is not None,
            filters.has_phone is not None,
            bool(filters.source_file.strip()),
            bool(filters.base_source.strip()),
            bool(filters.cpf.strip()),
            bool(filters.phone.strip()),
            bool(filters.ddd.strip()),
            bool(filters.name.strip()),
            bool(selected_matricula.strip()),
            bool(selected_servico.strip()),
        ]
    )
    total_results = 0
    pagination = build_pagination(total_results, 1, CLIENTS_PAGE_SIZE)
    results: list[dict[str, object]] = []
    if has_active_lookup:
        if quick_only_lookup and safe_page == 1:
            # Fast path for quick search: avoid expensive count on first render.
            prefetched = query_clients(session, filters, limit=CLIENTS_PAGE_SIZE + 1, offset=0)
            has_next_page = len(prefetched) > CLIENTS_PAGE_SIZE
            results = prefetched[:CLIENTS_PAGE_SIZE]
            total_results = len(results) + (1 if has_next_page else 0)
            pagination = build_pagination(total_results, 1, CLIENTS_PAGE_SIZE)
        else:
            total_results = count_client_results(session, filters)
            pagination = build_pagination(total_results, safe_page, CLIENTS_PAGE_SIZE)
            results = query_clients(session, filters, limit=CLIENTS_PAGE_SIZE, offset=int(pagination["offset"]))
    highlight = results[0] if results and has_active_lookup else None
    public_rows = results if normalized_segment in {"GOVERNO", "PREFEITURA"} else []
    has_public_matriculas = False
    if public_rows:
        rows_with_matricula = [row for row in public_rows if has_value(row.get("matricula"))]
        unique_rows_with_matricula: list[dict[str, object]] = []
        seen_matricula_keys: set[tuple[str, str, str]] = set()
        for row in rows_with_matricula:
            row_key = (
                str(row.get("matricula", "") or "").strip().upper(),
                str(row.get("servico", "") or "").strip().upper(),
                str(row.get("situacao", "") or "").strip().upper(),
            )
            if row_key in seen_matricula_keys:
                continue
            seen_matricula_keys.add(row_key)
            unique_rows_with_matricula.append(row)
        if rows_with_matricula:
            has_public_matriculas = True
            public_rows = unique_rows_with_matricula
            results = unique_rows_with_matricula
            if has_active_lookup:
                highlight = results[0]
    if public_rows and (selected_matricula or selected_servico):
        highlight = next(
            (
                row
                for row in public_rows
                if (not selected_matricula or row.get("matricula") == selected_matricula)
                and (not selected_servico or row.get("servico") == selected_servico)
            ),
            highlight,
        )
    selection_query = urlencode(build_filter_query_params(filters, normalized_segment, page=int(pagination["page"])))
    if public_rows:
        for row in public_rows:
            row["selection_url"] = "/clients?" + selection_query + "&" + urlencode(
                {
                    "selected_matricula": row.get("matricula", "") or "",
                    "selected_servico": row.get("servico", "") or "",
                }
            )
    saved_filters = session.execute(
        select(SavedFilter).where(SavedFilter.user_id == current_user.id, SavedFilter.page_key == "clients").order_by(SavedFilter.created_at.desc()).limit(10)
    ).scalars().all()
    crefaz_related_accounts: list[dict[str, str]] = []
    if normalized_segment == "CREFAZ" and highlight:
        related_rows = session.execute(
            select(
                ClientOccurrence.nome,
                ClientOccurrence.cpf,
                ClientOccurrence.cpf_original,
                ClientOccurrence.telefone1,
                ClientOccurrence.telefone2,
                ClientOccurrence.telefone3,
                ClientOccurrence.cidade,
                ClientOccurrence.uf,
                ClientOccurrence.extras_json,
            )
            .where(
                ClientOccurrence.base_segment == "CREFAZ",
                ClientOccurrence.cpf == str(highlight.get("cpf") or ""),
            )
            .order_by(ClientOccurrence.id.desc())
            .limit(500)
        ).all()
        seen_related_keys: set[tuple[str, str, str]] = set()
        for item in related_rows:
            extras: dict[str, object] = {}
            if item.extras_json:
                try:
                    extras = json.loads(item.extras_json)
                except json.JSONDecodeError:
                    extras = {}
            numero_cliente = str(
                extras.get("NÂº DO CLIENTE")
                or extras.get("N DO CLIENTE")
                or extras.get("NRO CLIENTE")
                or extras.get("NUMERO CLIENTE")
                or ""
            ).strip()
            conta_contrato = str(extras.get("CONTA CONTRATO") or extras.get("CONTA_CONTRATO") or "").strip()
            numero_instalacao = str(
                extras.get("NUMERO INSTALAÃ‡ÃƒO")
                or extras.get("NUMERO INSTALACAO")
                or extras.get("NÂº INSTALACAO")
                or extras.get("N INSTALACAO")
                or ""
            ).strip()
            dedupe_key = (numero_cliente, conta_contrato, numero_instalacao)
            if dedupe_key in seen_related_keys:
                continue
            seen_related_keys.add(dedupe_key)
            crefaz_related_accounts.append(
                {
                    "nome": str(item.nome or "").strip(),
                    "cpf_original": str(item.cpf_original or item.cpf or "").strip(),
                    "telefone": str(item.telefone1 or item.telefone2 or item.telefone3 or "").strip(),
                    "numero_cliente": numero_cliente,
                    "conta_contrato": conta_contrato,
                    "numero_instalacao": numero_instalacao,
                    "cidade": str(item.cidade or "").strip(),
                    "uf": str(item.uf or "").strip(),
                }
            )
    return render_template(
        request,
        "clients.html",
        {
            "title": f"Consulta {base_segment_label(normalized_segment)}",
            "results": results,
            "public_rows": public_rows,
            "highlight": highlight,
            "filters": filters,
            "result_count": total_results,
            "has_active_lookup": has_active_lookup,
            "selected_base_segment": normalized_segment,
            "selected_matricula": selected_matricula,
            "selected_servico": selected_servico,
            "selection_query": selection_query,
            "has_public_matriculas": has_public_matriculas,
            "pagination": pagination,
            "saved_filters": saved_filters,
            "crefaz_related_accounts": crefaz_related_accounts,
            "favorite_query": urlencode(build_filter_query_params(filters, normalized_segment, selected_matricula=selected_matricula, selected_servico=selected_servico)),
            "current_query_url": str(request.url.path) + (("?" + request.url.query) if request.url.query else ""),
            "message": request.query_params.get("message", ""),
            "error": request.query_params.get("error", ""),
        },
    )


def exports_page(
    request: Request,
    base_segment: str = "INSS",
    base_source: str = "",
    job_status: str = "",
    job_backend: str = "",
    message: str = "",
    error: str = "",
    current_user: AppUser = Depends(require_permission("can_export")),
    session: Session = Depends(get_session),
) -> HTMLResponse:
    normalized_segment = normalize_base_segment(base_segment)
    normalized_source = base_source.strip()
    normalized_job_status = (job_status or "").strip().upper()
    normalized_job_backend = (job_backend or "").strip().lower()
    exports = session.execute(select(ExportHistory).order_by(ExportHistory.criado_em.desc()).limit(50)).scalars().all()
    jobs_stmt = select(ExportJob).where(ExportJob.requested_by == current_user.username)
    if normalized_job_status in {"PENDENTE", "PROCESSANDO", "CONCLUIDO", "ERRO"}:
        jobs_stmt = jobs_stmt.where(ExportJob.status == normalized_job_status)
    if normalized_job_backend in {"thread", "rq"}:
        jobs_stmt = jobs_stmt.where(ExportJob.queue_backend == normalized_job_backend)
    export_jobs = session.execute(jobs_stmt.order_by(ExportJob.created_at.desc()).limit(20)).scalars().all()
    phone_enrichments = session.execute(select(PhoneEnrichment).order_by(PhoneEnrichment.criado_em.desc()).limit(20)).scalars().all()
    available_sources = [
        value
        for (value,) in session.execute(
            select(ClientOccurrence.base_source)
            .where(
                ClientOccurrence.base_segment == normalized_segment,
                ClientOccurrence.base_source != "",
            )
            .distinct()
            .order_by(ClientOccurrence.base_source.asc())
        ).all()
        if value
    ]
    public_filter_options = _build_public_filter_options(session, normalized_segment, normalized_source)
    return render_template(
        request,
        "exports.html",
        {
            "exports": exports,
            "export_jobs": export_jobs,
            "phone_enrichments": phone_enrichments,
            "available_sources": available_sources,
            "public_filter_options": public_filter_options,
            "message": message,
            "error": error,
            "title": f"Listas {base_segment_label(normalized_segment)}",
            "selected_base_segment": normalized_segment,
            "selected_base_source": normalized_source,
            "selected_job_status": normalized_job_status,
            "selected_job_backend": normalized_job_backend,
        },
    )


def run_export(
    request: Request,
    base_segment: str = Form("INSS"),
    base_source: str = Form(""),
    year: str = Form(""),
    month: str = Form(""),
    ddb_start: str = Form(""),
    ddb_end: str = Form(""),
    uf: str = Form(""),
    city: str = Form(""),
    entidade: list[str] = Form(default=[]),
    convenio: list[str] = Form(default=[]),
    servico_publico: list[str] = Form(default=[]),
    situacao: list[str] = Form(default=[]),
    cargo: list[str] = Form(default=[]),
    tipo_vinculo: list[str] = Form(default=[]),
    salary_min: str = Form(""),
    salary_max: str = Form(""),
    campaign_profile: str = Form(""),
    esp: list[str] = Form(default=[]),
    margin_min: str = Form(""),
    margin_max: str = Form(""),
    age_min: str = Form(""),
    age_max: str = Form(""),
    has_phone: str = Form(""),
    source_file: str = Form(""),
    cpf: str = Form(""),
    ddd: str = Form(""),
    name: str = Form(""),
    include_audit: bool = Form(False),
    file_format: str = Form("xlsx"),
    current_user: AppUser = Depends(require_permission("can_export")),
    session: Session = Depends(get_session),
) -> RedirectResponse:
    filters = FilterSet(
        base_segment=normalize_base_segment(base_segment),
        base_source=base_source.strip(),
        year=int(year) if year else None,
        month=int(month) if month else None,
        ddb_start=ddb_start,
        ddb_end=ddb_end,
        uf=uf,
        city=city,
        entidade="|".join(item.strip() for item in entidade if item.strip()),
        convenio="|".join(item.strip() for item in convenio if item.strip()),
        servico_publico="|".join(item.strip() for item in servico_publico if item.strip()),
        situacao="|".join(item.strip() for item in situacao if item.strip()),
        cargo="|".join(item.strip() for item in cargo if item.strip()),
        tipo_vinculo="|".join(item.strip() for item in tipo_vinculo if item.strip()),
        salary_min=parse_decimal(salary_min),
        salary_max=parse_decimal(salary_max),
        campaign_profile=campaign_profile,
        esp=",".join(item for item in esp if item),
        margin_min=parse_decimal(margin_min),
        margin_max=parse_decimal(margin_max),
        age_min=int(age_min) if age_min else None,
        age_max=int(age_max) if age_max else None,
        has_phone=True if has_phone == "sim" else False if has_phone == "nao" else None,
        source_file=source_file,
        cpf=cpf,
        ddd=ddd,
        name=name,
    )
    enqueue_export_job(
        session,
        current_user.username,
        "campaign",
        filters=filters,
        file_format=file_format,
        include_audit=include_audit,
        message="Exportacao colocada em fila.",
        metadata={"segmento": filters.base_segment, "formato": file_format, "auditoria": include_audit},
        request_id=getattr(request.state, "request_id", request_id_ctx.get("-")),
    )
    return RedirectResponse(url="/exports?message=" + quote_plus("Exportacao colocada em fila. Atualize a pagina para acompanhar."), status_code=303)


def run_cpf_enrichment_export(
    request: Request,
    base_segment: str = Form("INSS"),
    year: str = Form(""),
    month: str = Form(""),
    ddb_start: str = Form(""),
    ddb_end: str = Form(""),
    uf: str = Form(""),
    city: str = Form(""),
    esp: list[str] = Form(default=[]),
    margin_min: str = Form(""),
    margin_max: str = Form(""),
    age_min: str = Form(""),
    age_max: str = Form(""),
    has_phone: str = Form(""),
    source_file: str = Form(""),
    cpf: str = Form(""),
    ddd: str = Form(""),
    name: str = Form(""),
    file_format: str = Form("xlsx"),
    current_user: AppUser = Depends(require_permission("can_export")),
    session: Session = Depends(get_session),
) -> RedirectResponse:
    filters = FilterSet(
        base_segment=normalize_base_segment(base_segment),
        year=int(year) if year else None,
        month=int(month) if month else None,
        ddb_start=ddb_start,
        ddb_end=ddb_end,
        uf=uf,
        city=city,
        esp=",".join(item for item in esp if item),
        margin_min=parse_decimal(margin_min),
        margin_max=parse_decimal(margin_max),
        age_min=int(age_min) if age_min else None,
        age_max=int(age_max) if age_max else None,
        has_phone=True if has_phone == "sim" else False if has_phone == "nao" else None,
        source_file=source_file,
        cpf=cpf,
        ddd=ddd,
        name=name,
    )
    enqueue_export_job(
        session,
        current_user.username,
        "cpf_enrichment",
        filters=filters,
        file_format=file_format,
        message="Lista de CPFs colocada em fila.",
        metadata={"segmento": filters.base_segment, "formato": file_format},
        request_id=getattr(request.state, "request_id", request_id_ctx.get("-")),
    )
    return RedirectResponse(url="/exports?message=" + quote_plus("Lista de CPFs colocada em fila. Atualize a pagina para acompanhar."), status_code=303)


def run_updated_base_export(
    request: Request,
    base_segment: str = Form("INSS"),
    file_format: str = Form("xlsx"),
    current_user: AppUser = Depends(require_permission("can_export")),
    session: Session = Depends(get_session),
) -> RedirectResponse:
    filters = FilterSet(base_segment=normalize_base_segment(base_segment))
    enqueue_export_job(
        session,
        current_user.username,
        "updated_base",
        filters=filters,
        file_format=file_format,
        message="Base atualizada colocada em fila.",
        metadata={"segmento": filters.base_segment, "formato": file_format},
        request_id=getattr(request.state, "request_id", request_id_ctx.get("-")),
    )
    return RedirectResponse(url="/exports?message=" + quote_plus("Base atualizada colocada em fila. Atualize a pagina para acompanhar."), status_code=303)


def run_novavida_wrong_number_export(
    request: Request,
    base_segment: str = Form("INSS"),
    base_source: str = Form(""),
    file_format: str = Form("xlsx"),
    current_user: AppUser = Depends(require_permission("can_export")),
    session: Session = Depends(get_session),
) -> RedirectResponse:
    filters = FilterSet(
        base_segment=normalize_base_segment(base_segment),
        base_source=base_source.strip(),
    )
    enqueue_export_job(
        session,
        current_user.username,
        "novavida_wrong_number",
        filters=filters,
        file_format=file_format,
        message="Lista Nova Vida (numero errado) colocada em fila.",
        metadata={"segmento": filters.base_segment, "formato": file_format, "motivo": "nao_e_o_cliente"},
        request_id=getattr(request.state, "request_id", request_id_ctx.get("-")),
    )
    return RedirectResponse(url="/exports?message=" + quote_plus("Lista Nova Vida (numero errado) colocada em fila. Atualize a pagina para acompanhar."), status_code=303)


async def run_phone_enrichment_import(
    source_name: str = Form("LEMIT"),
    files: list[UploadFile] = File(default=[]),
    current_user: AppUser = Depends(require_permission("can_export")),
    session: Session = Depends(get_session),
) -> RedirectResponse:
    uploaded_files: list[tuple[str, bytes]] = []
    for item in files:
        if item.filename:
            uploaded_files.append((item.filename, await item.read()))
    if not uploaded_files:
        return RedirectResponse(url="/exports?error=" + quote_plus("Selecione ao menos um arquivo de retorno."), status_code=303)

    saved_paths = save_uploaded_files(uploaded_files)
    summary = import_phone_enrichments(session, saved_paths, source_name=source_name, imported_by=current_user.username)
    log_audit(
        session,
        current_user.username,
        "phone_enrichment_imported",
        target_type="phone_enrichment",
        target_id=source_name,
        message="Retorno telefonico importado.",
        metadata={"fonte": source_name, "arquivos": [path.name for path in saved_paths], "clientes_atualizados": summary.clients_updated},
    )
    session.commit()
    message = (
        f"Retorno importado: {summary.clients_updated} clientes atualizados, "
        f"{summary.data_fields_updated} atualizacao(oes) de dados complementares, "
        f"{summary.phones_added} telefone(s) novo(s) e {summary.skipped_not_found} CPF(s) sem cadastro."
    )
    return RedirectResponse(url="/exports?message=" + quote_plus(message), status_code=303)


async def run_public_update_import(
    base_segment: str = Form("GOVERNO"),
    source_name: str = Form(""),
    files: list[UploadFile] = File(default=[]),
    current_user: AppUser = Depends(require_permission("can_export")),
    session: Session = Depends(get_session),
) -> RedirectResponse:
    normalized_segment = normalize_base_segment(base_segment, default="GOVERNO")
    if normalized_segment not in {"GOVERNO", "PREFEITURA"}:
        return RedirectResponse(url="/exports?error=" + quote_plus("A atualizacao complementar e exclusiva para GOV/PREF."), status_code=303)

    uploaded_files: list[tuple[str, bytes]] = []
    for item in files:
        if item.filename:
            uploaded_files.append((item.filename, await item.read()))
    if not uploaded_files:
        return RedirectResponse(url=f"/exports?base_segment={normalized_segment}&error=" + quote_plus("Selecione ao menos um arquivo complementar."), status_code=303)

    saved_paths = save_uploaded_files(uploaded_files)
    summary = import_public_campaign_updates(
        session,
        saved_paths,
        base_segment=normalized_segment,
        source_name=source_name,
        imported_by=current_user.username,
    )
    log_audit(
        session,
        current_user.username,
        "public_update_imported",
        target_type="public_update",
        target_id=normalized_segment,
        message="Atualizacao complementar GOV/PREF importada.",
        metadata={"segmento": normalized_segment, "fonte": source_name, "arquivos": [path.name for path in saved_paths], "clientes_atualizados": summary.clients_updated},
    )
    session.commit()
    message = (
        f"Atualizacao GOV/PREF importada: {summary.rows_imported} linha(s), "
        f"{summary.contracts_saved} contrato(s) gravado(s), "
        f"{summary.clients_updated} cliente(s) atualizados e "
        f"{summary.skipped_not_found} CPF(s) sem base correspondente."
    )
    return RedirectResponse(url=f"/exports?base_segment={normalized_segment}&message=" + quote_plus(message), status_code=303)


def download_export(filename: str, current_user: AppUser = Depends(require_permission("can_export")), session: Session = Depends(get_session)) -> FileResponse:
    path = settings.exports_dir / filename
    if not path.exists():
        raise HTTPException(status_code=404, detail="Arquivo nao encontrado.")
    media_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
    log_audit(session, current_user.username, "download_export", target_type="export_file", target_id=filename, message="Arquivo de exportacao baixado.")
    session.commit()
    return FileResponse(path=path, filename=filename, media_type=media_type)


def history_page(
    request: Request,
    message: str = "",
    page: int = 1,
    _current_user: AppUser = Depends(require_permission("can_view_history")),
    session: Session = Depends(get_session),
) -> HTMLResponse:
    total_batches = session.scalar(select(func.count()).select_from(ImportBatch)) or 0
    pagination = build_pagination(total_batches, page, HISTORY_PAGE_SIZE)
    batches = (
        session.execute(
            select(ImportBatch)
            .order_by(ImportBatch.data_importacao.desc())
            .offset(int(pagination["offset"]))
            .limit(HISTORY_PAGE_SIZE)
        )
        .scalars()
        .all()
    )
    batch_ids = [batch.id for batch in batches]
    file_counts = {batch_id: 0 for batch_id in batch_ids}
    if batch_ids:
        file_counts.update(
            {
                batch_id: total
                for batch_id, total in session.execute(
                    select(SourceFile.batch_id, func.count())
                    .where(SourceFile.batch_id.in_(batch_ids))
                    .group_by(SourceFile.batch_id)
                ).all()
            }
        )
    return render_template(request, "history.html", {"batches": batches, "file_counts": file_counts, "message": message, "title": "Historico", "pagination": pagination})


def history_batch_detail(
    batch_id: int,
    request: Request,
    _current_user: AppUser = Depends(require_permission("can_view_history")),
    session: Session = Depends(get_session),
) -> HTMLResponse:
    batch = session.get(ImportBatch, batch_id)
    if batch is None:
        raise HTTPException(status_code=404, detail="Lote nao encontrado.")

    file_rows = (
        session.execute(
            select(
                SourceFile.id,
                SourceFile.nome_arquivo,
                SourceFile.semana_label,
                SourceFile.tipo_arquivo,
                SourceFile.total_linhas,
                func.count(ClientOccurrence.id).label("importadas"),
            )
            .outerjoin(ClientOccurrence, ClientOccurrence.source_file_id == SourceFile.id)
            .where(SourceFile.batch_id == batch_id)
            .group_by(
                SourceFile.id,
                SourceFile.nome_arquivo,
                SourceFile.semana_label,
                SourceFile.tipo_arquivo,
                SourceFile.total_linhas,
            )
            .order_by(SourceFile.nome_arquivo.asc())
        )
        .all()
    )
    imported_rows = sum(row.importadas for row in file_rows)
    total_rows = sum(row.total_linhas for row in file_rows)
    skipped_rows = max(total_rows - imported_rows, 0)
    source_file_ids = [row.id for row in file_rows]
    unique_contacts = 0
    contacts_with_phone = 0
    updated_ratio = 0
    if source_file_ids:
        unique_contacts = session.scalar(
            select(func.count(func.distinct(ClientOccurrence.cpf))).where(ClientOccurrence.source_file_id.in_(source_file_ids))
        ) or 0
        contacts_with_phone = session.scalar(
            select(func.count(func.distinct(ClientOccurrence.cpf))).where(
                ClientOccurrence.source_file_id.in_(source_file_ids),
                or_(
                    ClientOccurrence.telefone1 != "",
                    ClientOccurrence.telefone2 != "",
                    ClientOccurrence.telefone3 != "",
                ),
            )
        ) or 0
    if total_rows:
        updated_ratio = round((imported_rows / total_rows) * 100)
    return render_template(
        request,
        "history_detail.html",
        {
            "title": "Detalhes do lote",
            "batch": batch,
            "file_rows": file_rows,
            "imported_rows": imported_rows,
            "total_rows": total_rows,
            "skipped_rows": skipped_rows,
            "unique_contacts": unique_contacts,
            "contacts_with_phone": contacts_with_phone,
            "updated_ratio": updated_ratio,
        },
    )


def delete_history_batch(
    batch_id: int,
    current_user: AppUser = Depends(require_permission("can_delete_batches")),
    session: Session = Depends(get_session),
) -> RedirectResponse:
    ok, message = delete_batch(session, batch_id)
    log_audit(
        session,
        current_user.username,
        "batch_deleted" if ok else "batch_delete_failed",
        target_type="batch",
        target_id=str(batch_id),
        message=message or "",
    )
    session.commit()
    target = "/history?message=" + quote_plus(message if message else ("Campanha apagada com sucesso." if ok else "Falha ao apagar lote."))
    return RedirectResponse(url=target, status_code=303)


def cancel_history_batch(
    batch_id: int,
    current_user: AppUser = Depends(require_permission("can_delete_batches")),
    session: Session = Depends(get_session),
) -> RedirectResponse:
    batch = session.get(ImportBatch, batch_id)
    if batch is None:
        target = "/history?message=" + quote_plus("Lote nao encontrado.")
        return RedirectResponse(url=target, status_code=303)

    if batch.status not in {"PENDENTE", "PROCESSANDO"}:
        target = "/history?message=" + quote_plus(f"Lote {batch_id} nao esta ativo para cancelamento.")
        return RedirectResponse(url=target, status_code=303)

    cancelled = False
    cancel_error = ""
    if (batch.queue_backend or "").lower() == "rq" and (batch.queue_job_id or "").strip():
        try:
            cancelled = cancel_rq_job(batch.queue_job_id.strip())
        except Exception as exc:
            cancel_error = str(exc)

    batch.status = "CANCELADO"
    previous_summary = (batch.resumo or "").strip()
    reason = "Cancelado manualmente pelo painel."
    if cancel_error:
        reason = f"{reason} Falha ao enviar comando de cancelamento da fila: {cancel_error}"
    elif not cancelled and (batch.queue_backend or "").lower() == "rq":
        reason = f"{reason} Job de fila nao estava mais cancelavel (id={batch.queue_job_id})."
    batch.resumo = f"{previous_summary} | {reason}".strip(" |")
    session.commit()

    log_audit(
        session,
        current_user.username,
        "import_batch_cancelled",
        target_type="batch",
        target_id=str(batch_id),
        message="Cancelamento solicitado pelo painel.",
        metadata={
            "queue_backend": batch.queue_backend,
            "queue_name": batch.queue_name,
            "queue_job_id": batch.queue_job_id,
            "cancelled_in_queue": cancelled,
            "cancel_error": cancel_error,
        },
    )
    session.commit()

    message = f"Lote {batch_id} cancelado pelo painel."
    return RedirectResponse(url="/history?message=" + quote_plus(message), status_code=303)


def client_detail_page(
    cpf: str,
    request: Request,
    current_user: AppUser = Depends(require_permission("can_search")),
    session: Session = Depends(get_session),
) -> HTMLResponse:
    normalized_cpf = "".join(char for char in cpf if char.isdigit())
    if len(normalized_cpf) != 11:
        raise HTTPException(status_code=404, detail="Cliente nao encontrado.")

    client = session.get(Client, normalized_cpf)
    occurrences = (
        session.execute(
            select(ClientOccurrence, SourceFile.nome_arquivo, ImportBatch.id, ImportBatch.ano_referencia, ImportBatch.mes_referencia)
            .join(SourceFile, SourceFile.id == ClientOccurrence.source_file_id)
            .join(ImportBatch, ImportBatch.id == SourceFile.batch_id)
            .where(ClientOccurrence.cpf == normalized_cpf)
            .order_by(
                ImportBatch.ano_referencia.desc(),
                ImportBatch.mes_referencia.desc(),
                SourceFile.id.desc(),
                ClientOccurrence.id.desc(),
            )
            .limit(100)
        )
        .all()
    )
    if client is None and not occurrences:
        raise HTTPException(status_code=404, detail="Cliente nao encontrado.")

    latest_occurrence = occurrences[0][0] if occurrences else None
    if client is None and latest_occurrence is not None:
        highest_margin = max((occurrence.vl_margem for occurrence, *_rest in occurrences if occurrence.vl_margem is not None), default=None)
        latest_reference = f"{occurrences[0][4]:02d}/{occurrences[0][3]}" if occurrences[0][4] and occurrences[0][3] else "-"
        client = SimpleNamespace(
            cpf=normalized_cpf,
            nome_atual=latest_occurrence.nome or "",
            cidade_atual=latest_occurrence.cidade or "",
            uf_atual=latest_occurrence.uf or "",
            melhor_telefone=latest_occurrence.telefone1 or latest_occurrence.telefone2 or latest_occurrence.telefone3 or "",
            tem_telefone=bool(latest_occurrence.telefone1 or latest_occurrence.telefone2 or latest_occurrence.telefone3),
            tem_inss=any(normalize_base_segment(item[0].base_segment) == "INSS" for item in occurrences),
            tem_governo=any(normalize_base_segment(item[0].base_segment) == "GOVERNO" for item in occurrences),
            pis_atual=latest_occurrence.pis or "",
            matricula_atual=latest_occurrence.matricula or "",
            entidade_atual=latest_occurrence.entidade or "",
            maior_margem=highest_margin,
            ultima_referencia=latest_reference,
            qtd_ocorrencias=len(occurrences),
            atualizado_em=latest_occurrence.criado_em,
        )

    segment_details: dict[str, dict[str, object]] = {}
    public_occurrences_by_key: dict[tuple[str, str], tuple[ClientOccurrence, str, int, int, int]] = {}
    public_fallback_by_segment: dict[str, tuple[ClientOccurrence, str, int, int, int]] = {}
    for occurrence, file_name, batch_id, ano_referencia, mes_referencia in occurrences:
        segment = normalize_base_segment(occurrence.base_segment)
        if segment in {"GOVERNO", "PREFEITURA"} and occurrence.matricula:
            public_occurrences_by_key[(segment, occurrence.matricula)] = (occurrence, file_name, batch_id, ano_referencia, mes_referencia)
            public_fallback_by_segment.setdefault(segment, (occurrence, file_name, batch_id, ano_referencia, mes_referencia))
        if segment in segment_details:
            continue
        address_parts = [part for part in [occurrence.logradouro, occurrence.numero, occurrence.complemento, occurrence.bairro] if part]
        address_text = ", ".join(str(part) for part in address_parts)
        extras = _parse_extras_json(occurrence.extras_json or "")
        segment_details[segment] = {
            "segment": segment,
            "label": base_segment_label(segment),
            "source": occurrence.base_source or "-",
            "phone": occurrence.telefone1 or occurrence.telefone2 or occurrence.telefone3 or "",
            "phone_1": occurrence.telefone1 or "",
            "phone_2": occurrence.telefone2 or "",
            "phone_3": occurrence.telefone3 or "",
            "email": occurrence.email or "",
            "address": address_text,
            "bairro": occurrence.bairro or "",
            "cep": occurrence.cep or "",
            "city": occurrence.cidade or client.cidade_atual or "-",
            "uf": occurrence.uf or client.uf_atual or "-",
            "reference": f"{mes_referencia:02d}/{ano_referencia}" if mes_referencia and ano_referencia else "-",
            "margin": occurrence.vl_margem,
            "rmi": occurrence.vl_rmi,
            "rmc": occurrence.vl_rmc,
            "margem_total": occurrence.margem_total,
            "margem_liquida": occurrence.margem_liquida,
            "margem_bruta": occurrence.margem_bruta,
            "margem_utilizada": occurrence.margem_utilizada,
            "benefit": occurrence.nu_nb or "",
            "species": occurrence.esp or "",
            "matricula": occurrence.matricula or "",
            "entity": occurrence.entidade or "",
            "cbo_title": occurrence.cbo_titulo or "",
            "service": occurrence.servico or "",
            "status": occurrence.situacao or "",
            "pis": occurrence.pis or client.pis_atual or "",
            "salary": occurrence.salario,
            "file_name": file_name,
            "updated_at": occurrence.criado_em,
            "regional_alert": _first_extra_value(extras, "ALERTA_REGIONAL"),
            "ddd_phone": _first_extra_value(extras, "DDD_TELEFONE"),
            "ddd_uf": _first_extra_value(extras, "DDD_UF_BRASILAPI"),
            "numero_cliente": _first_extra_value(extras, "N DO CLIENTE", "Nº DO CLIENTE", "NRO CLIENTE", "NUMERO CLIENTE"),
            "conta_contrato": _first_extra_value(extras, "CONTA CONTRATO", "CONTA_CONTRATO"),
            "numero_instalacao": _first_extra_value(extras, "NUMERO INSTALACAO", "NUMERO INSTALAÇÃO", "N INSTALACAO", "Nº INSTALACAO"),
        }

    public_matricula_details: list[dict[str, object]] = []
    public_updates = (
        session.execute(
            select(PublicCampaignUpdate)
            .where(PublicCampaignUpdate.cpf == normalized_cpf)
            .order_by(
                PublicCampaignUpdate.base_segment.asc(),
                PublicCampaignUpdate.matricula.asc(),
                PublicCampaignUpdate.criado_em.desc(),
                PublicCampaignUpdate.id.desc(),
            )
        )
        .scalars()
        .all()
    )
    seen_public_keys: set[tuple[str, str]] = set()
    for update in public_updates:
        segment = normalize_base_segment(update.base_segment)
        if segment not in {"GOVERNO", "PREFEITURA"} or not update.matricula:
            continue
        key = (segment, update.matricula)
        if key in seen_public_keys:
            continue
        seen_public_keys.add(key)
        occurrence_bundle = public_occurrences_by_key.get(key) or public_fallback_by_segment.get(segment)
        occurrence = occurrence_bundle[0] if occurrence_bundle else None
        file_name = occurrence_bundle[1] if occurrence_bundle else "-"
        ano_referencia = occurrence_bundle[3] if occurrence_bundle else None
        mes_referencia = occurrence_bundle[4] if occurrence_bundle else None
        extras = _parse_extras_json(occurrence.extras_json or "") if occurrence is not None else {}
        public_matricula_details.append(
            {
                "segment": segment,
                "label": base_segment_label(segment),
                "matricula": update.matricula,
                "service": update.servico_servidor or (occurrence.servico if occurrence is not None else ""),
                "status": update.situacao or (occurrence.situacao if occurrence is not None else ""),
                "entity": occurrence.entidade if occurrence is not None else "",
                "secretaria": occurrence.secretaria if occurrence is not None else "",
                "convenio": occurrence.convenio if occurrence is not None else "",
                "pis": occurrence.pis if occurrence is not None else "",
                "salary": occurrence.salario if occurrence is not None else None,
                "margin": update.margem_disponivel if update.margem_disponivel is not None else (occurrence.vl_margem if occurrence is not None else None),
                "margem_total": update.margem_total if update.margem_total is not None else (occurrence.margem_total if occurrence is not None else None),
                "margem_liquida": occurrence.margem_liquida if occurrence is not None else None,
                "margem_bruta": occurrence.margem_bruta if occurrence is not None else None,
                "margem_utilizada": occurrence.margem_utilizada if occurrence is not None else None,
                "tipo_vinculo": _first_extra_value(extras, "REGIME_CONTRATACAO", "TIPO DE VINCULO", "TIPO_VINCULO", "VINCULO", "REGIME DE CONTRATACAO"),
                "cargo_funcao": _first_extra_value(extras, "CARGO", "FUNCAO", "FUNÃ‡ÃƒO", "FUNÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢O", "CARGO/FUNCAO", "CARGO FUNCAO", "CBO TITULO", "CBO_TITULO") or (occurrence.cbo_titulo if occurrence is not None else ""),
                "phone": (occurrence.telefone1 or occurrence.telefone2 or occurrence.telefone3) if occurrence is not None else "",
                "email": occurrence.email if occurrence is not None else "",
                "city": occurrence.cidade if occurrence is not None else client.cidade_atual,
                "uf": occurrence.uf if occurrence is not None else client.uf_atual,
                "source": update.fonte or (occurrence.base_source if occurrence is not None else ""),
                "source_file": update.arquivo_origem or file_name,
                "reference": f"{mes_referencia:02d}/{ano_referencia}" if mes_referencia and ano_referencia else "-",
                "banks": _first_extra_value(extras, "PUBLIC_UPDATE_BANCOS"),
                "consignataria": _first_extra_value(extras, "PUBLIC_UPDATE_CONSIGNATARIA_PRINCIPAL") or update.consignataria,
                "contracts_count": _first_extra_value(extras, "PUBLIC_UPDATE_QTD_CONTRATOS"),
                "active_count": _first_extra_value(extras, "PUBLIC_UPDATE_QTD_ATIVOS"),
                "updated_at": update.criado_em or (occurrence.criado_em if occurrence is not None else None),
            }
        )

    for (segment, matricula), occurrence_bundle in public_occurrences_by_key.items():
        if (segment, matricula) in seen_public_keys:
            continue
        occurrence, file_name, _batch_id, ano_referencia, mes_referencia = occurrence_bundle
        extras = _parse_extras_json(occurrence.extras_json or "")
        public_matricula_details.append(
            {
                "segment": segment,
                "label": base_segment_label(segment),
                "matricula": matricula,
                "service": occurrence.servico or "",
                "status": occurrence.situacao or "",
                "entity": occurrence.entidade or "",
                "secretaria": occurrence.secretaria or "",
                "convenio": occurrence.convenio or "",
                "pis": occurrence.pis or "",
                "salary": occurrence.salario,
                "margin": occurrence.vl_margem,
                "margem_total": occurrence.margem_total,
                "margem_liquida": occurrence.margem_liquida,
                "margem_bruta": occurrence.margem_bruta,
                "margem_utilizada": occurrence.margem_utilizada,
                "tipo_vinculo": _first_extra_value(extras, "REGIME_CONTRATACAO", "TIPO DE VINCULO", "TIPO_VINCULO", "VINCULO", "REGIME DE CONTRATACAO"),
                "cargo_funcao": _first_extra_value(extras, "CARGO", "FUNCAO", "FUNÃ‡ÃƒO", "FUNÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢O", "CARGO/FUNCAO", "CARGO FUNCAO", "CBO TITULO", "CBO_TITULO") or occurrence.cbo_titulo,
                "phone": occurrence.telefone1 or occurrence.telefone2 or occurrence.telefone3 or "",
                "email": occurrence.email or "",
                "city": occurrence.cidade or client.cidade_atual,
                "uf": occurrence.uf or client.uf_atual,
                "source": occurrence.base_source or "",
                "source_file": file_name,
                "reference": f"{mes_referencia:02d}/{ano_referencia}" if mes_referencia and ano_referencia else "-",
                "banks": _first_extra_value(extras, "PUBLIC_UPDATE_BANCOS"),
                "consignataria": _first_extra_value(extras, "PUBLIC_UPDATE_CONSIGNATARIA_PRINCIPAL"),
                "contracts_count": _first_extra_value(extras, "PUBLIC_UPDATE_QTD_CONTRATOS"),
                "active_count": _first_extra_value(extras, "PUBLIC_UPDATE_QTD_ATIVOS"),
                "updated_at": occurrence.criado_em,
            }
        )

    public_matricula_details.sort(
        key=lambda item: (
            {"GOVERNO": 0, "PREFEITURA": 1}.get(str(item.get("segment", "")), 9),
            str(item.get("matricula", "")),
        )
    )
    log_audit(session, current_user.username, "client_detail_viewed", target_type="client", target_id=normalized_cpf, message="Detalhe do cliente consultado.")
    session.commit()
    return render_template(
        request,
        "client_detail.html",
        {
            "title": f"Cliente {client.nome_atual or client.cpf}",
            "client": client,
            "latest_occurrence": latest_occurrence,
            "occurrences": occurrences,
            "segment_details": [segment_details[key] for key in ["INSS", "CREFAZ", "GOVERNO", "PREFEITURA"] if key in segment_details],
            "public_matricula_details": public_matricula_details,
        },
    )


def save_client_filter(
    name: str = Form(...),
    query_string: str = Form(""),
    current_user: AppUser = Depends(require_permission("can_search")),
    session: Session = Depends(get_session),
) -> RedirectResponse:
    favorite_name = name.strip()
    if not favorite_name:
        return RedirectResponse(url="/clients?error=" + quote_plus("Informe um nome para o filtro."), status_code=303)
    session.add(SavedFilter(user_id=current_user.id, page_key="clients", name=favorite_name[:120], query_string=query_string.strip()))
    log_audit(session, current_user.username, "saved_filter_created", target_type="saved_filter", message="Filtro favorito salvo.", metadata={"nome": favorite_name})
    session.commit()
    return RedirectResponse(url="/clients?message=" + quote_plus("Filtro favorito salvo."), status_code=303)


def delete_client_filter(
    favorite_id: int,
    current_user: AppUser = Depends(require_permission("can_search")),
    session: Session = Depends(get_session),
) -> RedirectResponse:
    favorite = session.get(SavedFilter, favorite_id)
    if favorite is None or favorite.user_id != current_user.id:
        return RedirectResponse(url="/clients?error=" + quote_plus("Filtro favorito nao encontrado."), status_code=303)
    favorite_name = favorite.name
    session.delete(favorite)
    log_audit(session, current_user.username, "saved_filter_deleted", target_type="saved_filter", target_id=str(favorite_id), message="Filtro favorito removido.", metadata={"nome": favorite_name})
    session.commit()
    return RedirectResponse(url="/clients?message=" + quote_plus("Filtro favorito removido."), status_code=303)


def set_client_do_not_call(
    cpf: str,
    enabled: bool = Form(False),
    dnc_reason: str = Form(""),
    return_url: str = Form("/clients"),
    current_user: AppUser = Depends(require_permission("can_search")),
    session: Session = Depends(get_session),
) -> RedirectResponse:
    allowed_dnc_reasons = {
        "nao_e_o_cliente": "Nao e o cliente",
        "cliente_ja_fechou_conosco": "Cliente ja fechou conosco",
        "nao_quer_mais_receber_ligacao": "Nao quer mais receber ligacao",
    }
    normalized_cpf = normalize_cpf(cpf)
    normalized_reason = (dnc_reason or "").strip()
    if enabled and normalized_reason not in allowed_dnc_reasons:
        return RedirectResponse(url="/clients?error=" + quote_plus("Selecione um motivo para marcar NAO LIGAR."), status_code=303)
    client = session.get(Client, normalized_cpf)
    if client is None:
        rebuild_clients_for_cpfs(session, {normalized_cpf})
        session.commit()
        client = session.get(Client, normalized_cpf)
    if client is None:
        return RedirectResponse(url="/clients?error=" + quote_plus("Cliente nao encontrado para marcar bloqueio de ligacao."), status_code=303)
    client.do_not_call = bool(enabled)
    client.do_not_call_reason = normalized_reason if client.do_not_call else ""
    action = "do_not_call_enabled" if client.do_not_call else "do_not_call_disabled"
    if client.do_not_call:
        message = f"Cliente marcado como NAO LIGAR. Motivo: {allowed_dnc_reasons[client.do_not_call_reason]}."
    else:
        message = "Cliente removido da lista NAO LIGAR."
    log_audit(
        session,
        current_user.username,
        action,
        target_type="client",
        target_id=normalized_cpf,
        message=message,
        metadata={"cpf": normalized_cpf, "enabled": client.do_not_call, "reason_key": client.do_not_call_reason, "reason_label": allowed_dnc_reasons.get(client.do_not_call_reason, "")},
    )
    session.commit()
    target_url = safe_next_path(return_url, "/clients")
    separator = "&" if "?" in target_url else "?"
    return RedirectResponse(url=f"{target_url}{separator}message=" + quote_plus(message), status_code=303)


def users_page(
    request: Request,
    current_user: AppUser = Depends(require_permission("can_manage_users")),
    session: Session = Depends(get_session),
) -> HTMLResponse:
    users = session.execute(select(AppUser).order_by(AppUser.username.asc())).scalars().all()
    recent_audits = session.execute(select(AuditLog).order_by(AuditLog.created_at.desc()).limit(20)).scalars().all()
    return render_template(
        request,
        "users.html",
        {
            "title": "Usuarios",
            "users": users,
            "permission_labels": PERMISSION_LABELS,
            "recent_audits": recent_audits,
            "message": request.query_params.get("message", ""),
            "error": request.query_params.get("error", ""),
            "current_manager_id": current_user.id,
        },
    )


def create_user(
    username: str = Form(...),
    full_name: str = Form(""),
    password: str = Form(...),
    is_active: bool = Form(False),
    is_admin: bool = Form(False),
    can_view_dashboard: bool = Form(False),
    can_import: bool = Form(False),
    can_search: bool = Form(False),
    can_access_consignado_inss: bool = Form(False),
    can_export: bool = Form(False),
    can_view_history: bool = Form(False),
    can_delete_batches: bool = Form(False),
    can_manage_users: bool = Form(False),
    current_user: AppUser = Depends(require_permission("can_manage_users")),
    session: Session = Depends(get_session),
) -> RedirectResponse:
    normalized_username = username.strip().lower()
    if not normalized_username:
        return RedirectResponse(url="/users?error=" + quote_plus("Informe um usuario valido."), status_code=303)
    existing = session.scalar(select(AppUser).where(AppUser.username == normalized_username))
    if existing is not None:
        return RedirectResponse(url="/users?error=" + quote_plus("Ja existe um usuario com esse login."), status_code=303)
    password_hash, password_salt = hash_password(password)
    new_user = AppUser(
        username=normalized_username,
        full_name=full_name.strip(),
        password_hash=password_hash,
        password_salt=password_salt,
        is_active=is_active,
        is_admin=is_admin,
        can_view_dashboard=can_view_dashboard,
        can_import=can_import,
        can_search=can_search,
        can_access_consignado_inss=can_access_consignado_inss,
        can_export=can_export,
        can_view_history=can_view_history,
        can_delete_batches=can_delete_batches,
        can_manage_users=can_manage_users,
    )
    session.add(new_user)
    log_audit(session, current_user.username, "user_created", target_type="user", target_id=normalized_username, message="Usuario criado.")
    session.commit()
    return RedirectResponse(url="/users?message=" + quote_plus("Usuario criado com sucesso."), status_code=303)


def update_user(
    user_id: int,
    username: str = Form(...),
    full_name: str = Form(""),
    password: str = Form(""),
    is_active: bool = Form(False),
    is_admin: bool = Form(False),
    can_view_dashboard: bool = Form(False),
    can_import: bool = Form(False),
    can_search: bool = Form(False),
    can_access_consignado_inss: bool = Form(False),
    can_export: bool = Form(False),
    can_view_history: bool = Form(False),
    can_delete_batches: bool = Form(False),
    can_manage_users: bool = Form(False),
    current_user: AppUser = Depends(require_permission("can_manage_users")),
    session: Session = Depends(get_session),
) -> RedirectResponse:
    user = session.get(AppUser, user_id)
    if user is None:
        return RedirectResponse(url="/users?error=" + quote_plus("Usuario nao encontrado."), status_code=303)

    normalized_username = username.strip().lower()
    duplicate = session.scalar(select(AppUser).where(AppUser.username == normalized_username, AppUser.id != user_id))
    if duplicate is not None:
        return RedirectResponse(url="/users?error=" + quote_plus("Ja existe outro usuario com esse login."), status_code=303)

    user.username = normalized_username
    user.full_name = full_name.strip()
    user.is_active = is_active
    user.is_admin = is_admin
    user.can_view_dashboard = can_view_dashboard
    user.can_import = can_import
    user.can_search = can_search
    user.can_access_consignado_inss = can_access_consignado_inss
    user.can_export = can_export
    user.can_view_history = can_view_history
    user.can_delete_batches = can_delete_batches
    user.can_manage_users = can_manage_users
    if password.strip():
        user.password_hash, user.password_salt = hash_password(password.strip())
        user.last_password_reset_at = current_utc()
    if user.id == current_user.id and not user.is_active:
        return RedirectResponse(url="/users?error=" + quote_plus("Voce nao pode desativar seu proprio usuario."), status_code=303)
    log_audit(session, current_user.username, "user_updated", target_type="user", target_id=str(user.id), message="Usuario atualizado.", metadata={"username": normalized_username})
    session.commit()
    return RedirectResponse(url="/users?message=" + quote_plus("Usuario atualizado com sucesso."), status_code=303)


def reset_user_password(
    user_id: int,
    new_password: str = Form(...),
    current_user: AppUser = Depends(require_permission("can_manage_users")),
    session: Session = Depends(get_session),
) -> RedirectResponse:
    user = session.get(AppUser, user_id)
    if user is None:
        return RedirectResponse(url="/users?error=" + quote_plus("Usuario nao encontrado."), status_code=303)
    password = new_password.strip()
    if len(password) < 4:
        return RedirectResponse(url="/users?error=" + quote_plus("Informe uma senha valida para reset."), status_code=303)
    user.password_hash, user.password_salt = hash_password(password)
    user.last_password_reset_at = current_utc()
    user.locked_until = None
    log_audit(session, current_user.username, "user_password_reset", target_type="user", target_id=str(user.id), message="Senha redefinida.", metadata={"username": user.username})
    session.commit()
    return RedirectResponse(url="/users?message=" + quote_plus(f"Senha redefinida para {user.username}."), status_code=303)


def lock_user_temporarily(
    user_id: int,
    minutes: int = Form(60),
    current_user: AppUser = Depends(require_permission("can_manage_users")),
    session: Session = Depends(get_session),
) -> RedirectResponse:
    user = session.get(AppUser, user_id)
    if user is None:
        return RedirectResponse(url="/users?error=" + quote_plus("Usuario nao encontrado."), status_code=303)
    if user.id == current_user.id and minutes > 0:
        return RedirectResponse(url="/users?error=" + quote_plus("Voce nao pode bloquear seu proprio usuario."), status_code=303)
    if minutes <= 0:
        user.locked_until = None
        action = "user_unlocked"
        message = f"Bloqueio removido de {user.username}."
    else:
        from datetime import timedelta

        user.locked_until = current_utc() + timedelta(minutes=minutes)
        action = "user_locked"
        message = f"Usuario {user.username} bloqueado por {minutes} minuto(s)."
    log_audit(session, current_user.username, action, target_type="user", target_id=str(user.id), message=message, metadata={"username": user.username, "minutes": minutes})
    session.commit()
    return RedirectResponse(url="/users?message=" + quote_plus(message), status_code=303)


def healthcheck() -> JSONResponse:
    payload = {"status": "ok", "kind": "liveness"}
    return JSONResponse(payload, status_code=200)


def readiness_check(session: Session = Depends(get_session)) -> JSONResponse:
    checks: dict[str, object] = {
        "database": "ok",
        "directories": {},
        "queue": {},
    }
    errors: list[str] = []

    try:
        session.scalar(select(1))
    except Exception as exc:
        checks["database"] = "error"
        errors.append(f"database: {exc}")

    directory_map = {
        "uploads": settings.uploads_dir,
        "archive": settings.archive_dir,
        "exports": settings.exports_dir,
    }
    directory_checks: dict[str, str] = {}
    for label, path in directory_map.items():
        status = "ok" if path.exists() and path.is_dir() else "error"
        directory_checks[label] = status
        if status != "ok":
            errors.append(f"directory_{label}: path '{path}' nao existe ou nao e pasta")
    checks["directories"] = directory_checks

    queue_backend = getattr(settings, "queue_backend", "thread")
    queue_checks: dict[str, object] = {
        "backend": queue_backend,
        "status": "ok",
    }
    if queue_backend == "rq":
        try:
            from redis import Redis

            redis_url = getattr(settings, "queue_redis_url", "redis://localhost:6379/0")
            redis_conn = Redis.from_url(redis_url)
            redis_conn.ping()
            queue_checks["redis"] = "ok"
        except Exception as exc:
            queue_checks["status"] = "error"
            queue_checks["redis"] = "error"
            errors.append(f"queue_redis: {exc}")
    checks["queue"] = queue_checks

    if errors:
        payload = {"status": "degraded", "kind": "readiness", "checks": checks, "errors": errors}
        return JSONResponse(payload, status_code=503)

    payload = {"status": "ok", "kind": "readiness", "checks": checks}
    return JSONResponse(payload, status_code=200)


def c6_worker_health(current_user: AppUser = Depends(require_permission("can_access_consignado_inss"))) -> JSONResponse:
    _ = current_user
    payload = {
        "enabled": settings.c6_worker_enabled,
        "base_url": settings.c6_base_url,
        "has_credentials": bool(settings.c6_username and settings.c6_password),
    }
    return JSONResponse(payload, status_code=200)


def c6_generate_worker_offer(
    request: Request,
    body: dict[str, object] = Body(...),
    current_user: AppUser = Depends(require_permission("can_access_consignado_inss")),
    session: Session = Depends(get_session),
) -> JSONResponse:
    try:
        response = c6_worker_loan_client.generate_offer(body)
    except C6WorkerLoanError as exc:
        raise HTTPException(status_code=502, detail=str(exc)) from exc

    log_audit(
        session,
        current_user.username,
        "c6_worker_offer_generated",
        target_type="integration",
        message="Oferta do consignado trabalhador solicitada.",
        metadata={"path": str(request.url.path)},
    )
    session.commit()
    return JSONResponse({"status_code": response.status_code, "data": response.payload}, status_code=200)


def c6_simulate_worker_loan(
    request: Request,
    body: dict[str, object] = Body(...),
    version: str = "v2",
    current_user: AppUser = Depends(require_permission("can_access_consignado_inss")),
    session: Session = Depends(get_session),
) -> JSONResponse:
    try:
        response = c6_worker_loan_client.simulate_proposal(body, version=version)
    except C6WorkerLoanError as exc:
        raise HTTPException(status_code=502, detail=str(exc)) from exc

    log_audit(
        session,
        current_user.username,
        "c6_worker_simulation_requested",
        target_type="integration",
        message="Simulacao do consignado trabalhador solicitada.",
        metadata={"path": str(request.url.path), "version": version},
    )
    session.commit()
    return JSONResponse({"status_code": response.status_code, "data": response.payload}, status_code=200)


def c6_include_worker_loan(
    request: Request,
    body: dict[str, object] = Body(...),
    current_user: AppUser = Depends(require_permission("can_access_consignado_inss")),
    session: Session = Depends(get_session),
) -> JSONResponse:
    try:
        response = c6_worker_loan_client.include_proposal(body)
    except C6WorkerLoanError as exc:
        raise HTTPException(status_code=502, detail=str(exc)) from exc

    log_audit(
        session,
        current_user.username,
        "c6_worker_proposal_included",
        target_type="integration",
        message="Inclusao do consignado trabalhador solicitada.",
        metadata={"path": str(request.url.path)},
    )
    session.commit()
    return JSONResponse({"status_code": response.status_code, "data": response.payload}, status_code=200)


from .routes.auth import router as auth_router
from .routes.clients import router as clients_router
from .routes.imports import router as imports_router
from .routes.exports import router as exports_router
from .routes.history import router as history_router
from .routes.users import router as users_router
from .routes.system import router as system_router

app.include_router(auth_router)
app.include_router(clients_router)
app.include_router(imports_router)
app.include_router(exports_router)
app.include_router(history_router)
app.include_router(users_router)
app.include_router(system_router)







