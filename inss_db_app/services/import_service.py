from __future__ import annotations

from .. import main as main_module
from ..schemas import ImportRequest


def run_import_job(batch_id: int, request_data: ImportRequest, request_id: str = "-") -> None:
    main_module.run_import_job(batch_id=batch_id, request_data=request_data, request_id=request_id)

