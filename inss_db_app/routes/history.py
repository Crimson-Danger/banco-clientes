from __future__ import annotations

from fastapi import APIRouter
from fastapi.responses import HTMLResponse

from .. import main as main_module


router = APIRouter()

router.add_api_route("/history", main_module.history_page, methods=["GET"], response_class=HTMLResponse)
router.add_api_route("/history/{batch_id}", main_module.history_batch_detail, methods=["GET"], response_class=HTMLResponse)
router.add_api_route("/history/{batch_id}/delete", main_module.delete_history_batch, methods=["POST"])
router.add_api_route("/history/{batch_id}/cancel", main_module.cancel_history_batch, methods=["POST"])
