from __future__ import annotations

from fastapi import APIRouter
from fastapi.responses import HTMLResponse

from .. import main as main_module


router = APIRouter()

router.add_api_route("/login", main_module.login_page, methods=["GET"], response_class=HTMLResponse)
router.add_api_route("/login", main_module.login_submit, methods=["POST"], response_class=HTMLResponse)
router.add_api_route("/logout", main_module.logout, methods=["GET"])
router.add_api_route("/c6", main_module.c6_entry, methods=["GET"])
router.add_api_route("/c6-inss", main_module.c6_inss_page, methods=["GET"], response_class=HTMLResponse)
router.add_api_route("/c6-inss", main_module.c6_inss_submit, methods=["POST"], response_class=HTMLResponse)
router.add_api_route("/c6-inss/include", main_module.c6_inss_include_page, methods=["GET"], response_class=HTMLResponse)
router.add_api_route("/c6-inss/include", main_module.c6_inss_include_submit, methods=["POST"], response_class=HTMLResponse)
