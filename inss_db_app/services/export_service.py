from __future__ import annotations

from sqlalchemy.orm import Session

from .. import main as main_module
from ..importer import FilterSet
from ..models import ExportJob


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
    return main_module.enqueue_export_job(
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

