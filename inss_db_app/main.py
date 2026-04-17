from __future__ import annotations

from contextlib import asynccontextmanager
from pathlib import Path
import threading
from urllib.parse import quote_plus

from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy import func, select
from sqlalchemy.orm import Session

from .config import settings
from .database import Base, SessionLocal, engine, ensure_runtime_indexes
from .importer import (
    FilterSet,
    create_import_batch,
    delete_batch,
    export_clients,
    format_money,
    import_files,
    parse_decimal,
    process_import_batch,
    query_clients,
    save_uploaded_files,
)
from .models import Client, ClientOccurrence, ExportHistory, ImportBatch, SourceFile
from .schemas import ImportRequest


BASE_DIR = Path(__file__).resolve().parent
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))


def format_reference(value: str) -> str:
    if not value or "-" not in value:
        return value
    year, month = value.split("-", 1)
    return f"{month}/{year}"


def format_phone(value: str) -> str:
    digits = "".join(char for char in (value or "") if char.isdigit())
    if len(digits) == 11:
        return f"({digits[:2]}) {digits[2:7]}-{digits[7:]}"
    if len(digits) == 10:
        return f"({digits[:2]}) {digits[2:6]}-{digits[6:]}"
    return value


templates.env.filters["brl"] = format_money
templates.env.filters["ref_br"] = format_reference
templates.env.filters["phone_br"] = format_phone

def get_session() -> Session:
    session = SessionLocal()
    try:
        yield session
    finally:
        session.close()


def run_import_job(batch_id: int, request_data: ImportRequest) -> None:
    session = SessionLocal()
    try:
        batch = session.get(ImportBatch, batch_id)
        if batch is None:
            return
        process_import_batch(session, batch, request_data)
    finally:
        session.close()


def create_schema() -> None:
    Base.metadata.create_all(bind=engine)
    ensure_runtime_indexes()


@asynccontextmanager
async def lifespan(_: FastAPI):
    create_schema()
    yield


app = FastAPI(title="Banco de Clientes INSS", lifespan=lifespan)
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")


@app.get("/", response_class=HTMLResponse)
def home(request: Request, session: Session = Depends(get_session)) -> HTMLResponse:
    total_batches = session.scalar(select(func.count()).select_from(ImportBatch)) or 0
    total_files = session.scalar(select(func.count()).select_from(SourceFile)) or 0
    total_occurrences = session.scalar(select(func.count()).select_from(ClientOccurrence)) or 0
    total_clients = session.scalar(select(func.count()).select_from(Client)) or 0
    total_with_phone = session.scalar(select(func.count()).select_from(Client).where(Client.tem_telefone.is_(True))) or 0
    latest_exports = session.execute(select(ExportHistory).order_by(ExportHistory.criado_em.desc()).limit(5)).scalars().all()
    top_states = session.execute(
        select(Client.uf_atual, func.count())
        .where(Client.uf_atual != "")
        .group_by(Client.uf_atual)
        .order_by(func.count().desc())
        .limit(5)
    ).all()
    latest_batches = session.execute(
        select(ImportBatch).order_by(ImportBatch.data_importacao.desc()).limit(4)
    ).scalars().all()

    phone_ratio = round((total_with_phone / total_clients) * 100, 1) if total_clients else 0
    no_phone_ratio = round(100 - phone_ratio, 1) if total_clients else 0
    state_total = sum(item[1] for item in top_states) or 1
    states_chart = [
        {"uf": uf or "N/D", "count": count, "pct": round((count / state_total) * 100, 1)}
        for uf, count in top_states
    ]
    return templates.TemplateResponse(
        request,
        "home.html",
        {
            "total_batches": total_batches,
            "total_files": total_files,
            "total_occurrences": total_occurrences,
            "total_clients": total_clients,
            "total_with_phone": total_with_phone,
            "phone_ratio": phone_ratio,
            "no_phone_ratio": no_phone_ratio,
            "states_chart": states_chart,
            "latest_batches": latest_batches,
            "latest_exports": latest_exports,
        },
    )


@app.get("/import", response_class=HTMLResponse)
def import_page(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(request, "import.html", {"summary": None, "error": ""})


@app.post("/import", response_class=HTMLResponse)
async def run_import(
    request: Request,
    year: int = Form(...),
    month: int = Form(...),
    user_name: str = Form("operador"),
    origin_folder: str = Form("upload_web"),
    allowed_species: list[str] = Form(default=[]),
    custom_species: str = Form(""),
    server_folder: str = Form(""),
    files: list[UploadFile] = File(default=[]),
    session: Session = Depends(get_session),
) -> HTMLResponse:
    selected_paths: list[Path] = []
    if server_folder.strip():
        folder = Path(server_folder.strip())
        if not folder.exists() or not folder.is_dir():
            return templates.TemplateResponse(request, "import.html", {"summary": None, "error": f"Pasta nao encontrada: {folder}"}, status_code=400)
        selected_paths = [path for path in sorted(folder.iterdir()) if path.suffix.lower() in {".csv", ".xlsx", ".xlsm", ".xls"}]
        origin_folder = str(folder)
    elif files:
        selected_paths = save_uploaded_files([(item.filename or "arquivo.csv", await item.read()) for item in files])
    else:
        return templates.TemplateResponse(request, "import.html", {"summary": None, "error": "Envie arquivos ou informe uma pasta do servidor."}, status_code=400)

    import_request = ImportRequest(
        year=year,
        month=month,
        user_name=user_name,
        origin_folder=origin_folder,
        files=selected_paths,
        allowed_species=[*allowed_species, custom_species],
    )
    batch = create_import_batch(session, import_request)
    session.commit()
    threading.Thread(target=run_import_job, args=(batch.id, import_request), daemon=True).start()
    return RedirectResponse(url=f"/history/{batch.id}", status_code=303)


@app.get("/clients", response_class=HTMLResponse)
def clients_page(
    request: Request,
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
    session: Session = Depends(get_session),
) -> HTMLResponse:
    filters = FilterSet(
        quick_mode="telefone" if quick_mode == "telefone" else "cpf",
        quick_value=quick_value,
        year=int(year) if year else None,
        month=int(month) if month else None,
        week=week,
        uf=uf,
        city=city,
        esp=esp,
        margin_min=parse_decimal(margin_min),
        margin_max=parse_decimal(margin_max),
        age_min=int(age_min) if age_min else None,
        age_max=int(age_max) if age_max else None,
        has_phone=True if has_phone == "sim" else False if has_phone == "nao" else None,
        source_file=source_file,
        cpf=cpf,
        phone=phone,
        ddd=ddd,
        name=name,
    )
    results = query_clients(session, filters)
    highlight = results[0] if results else None
    return templates.TemplateResponse(
        request,
        "clients.html",
        {
            "results": results,
            "highlight": highlight,
            "filters": filters,
            "result_count": len(results),
        },
    )


@app.get("/exports", response_class=HTMLResponse)
def exports_page(request: Request, session: Session = Depends(get_session)) -> HTMLResponse:
    exports = session.execute(select(ExportHistory).order_by(ExportHistory.criado_em.desc()).limit(50)).scalars().all()
    return templates.TemplateResponse(request, "exports.html", {"exports": exports})


@app.post("/exports")
def run_export(
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
    include_audit: bool = Form(False),
    file_format: str = Form("xlsx"),
    session: Session = Depends(get_session),
) -> RedirectResponse:
    filters = FilterSet(
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
    export_path = export_clients(session, filters, include_audit=include_audit, file_format=file_format)
    return RedirectResponse(url=f"/downloads/{export_path.name}", status_code=303)


@app.get("/downloads/{filename}")
def download_export(filename: str) -> FileResponse:
    path = settings.exports_dir / filename
    if not path.exists():
        raise HTTPException(status_code=404, detail="Arquivo nao encontrado.")
    return FileResponse(path=path, filename=filename, media_type="text/csv")


@app.get("/history", response_class=HTMLResponse)
def history_page(request: Request, message: str = "", session: Session = Depends(get_session)) -> HTMLResponse:
    batches = session.execute(select(ImportBatch).order_by(ImportBatch.data_importacao.desc()).limit(50)).scalars().all()
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
    return templates.TemplateResponse(request, "history.html", {"batches": batches, "file_counts": file_counts, "message": message})


@app.get("/history/{batch_id}", response_class=HTMLResponse)
def history_batch_detail(batch_id: int, request: Request, session: Session = Depends(get_session)) -> HTMLResponse:
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
    return templates.TemplateResponse(
        request,
        "history_detail.html",
        {
            "batch": batch,
            "file_rows": file_rows,
            "imported_rows": imported_rows,
            "total_rows": total_rows,
            "skipped_rows": skipped_rows,
        },
    )


@app.post("/history/{batch_id}/delete")
def delete_history_batch(batch_id: int, session: Session = Depends(get_session)) -> RedirectResponse:
    ok, message = delete_batch(session, batch_id)
    target = "/history?message=" + quote_plus(message if message else ("Campanha apagada com sucesso." if ok else "Falha ao apagar lote."))
    return RedirectResponse(url=target, status_code=303)
