from __future__ import annotations

from fastapi import APIRouter
from fastapi.responses import HTMLResponse

from .. import main as main_module


router = APIRouter()

router.add_api_route("/exports", main_module.exports_page, methods=["GET"], response_class=HTMLResponse)
router.add_api_route("/exports", main_module.run_export, methods=["POST"])
router.add_api_route("/exports/cpf-enrichment", main_module.run_cpf_enrichment_export, methods=["POST"])
router.add_api_route("/exports/novavida-wrong-number", main_module.run_novavida_wrong_number_export, methods=["POST"])
router.add_api_route("/exports/updated-base", main_module.run_updated_base_export, methods=["POST"])
router.add_api_route("/exports/phone-enrichment-import", main_module.run_phone_enrichment_import, methods=["POST"])
router.add_api_route("/exports/public-update-import", main_module.run_public_update_import, methods=["POST"])
router.add_api_route("/downloads/{filename}", main_module.download_export, methods=["GET"])
