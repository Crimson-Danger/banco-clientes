from __future__ import annotations

from fastapi import APIRouter
from fastapi.responses import HTMLResponse

from .. import main as main_module


router = APIRouter()

router.add_api_route("/import", main_module.import_page, methods=["GET"], response_class=HTMLResponse)
router.add_api_route("/import", main_module.run_import, methods=["POST"], response_class=HTMLResponse)
router.add_api_route("/import/public-update-import", main_module.run_public_update_import_from_import_page, methods=["POST"])
