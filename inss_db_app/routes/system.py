from __future__ import annotations

from fastapi import APIRouter
from fastapi.responses import HTMLResponse

from .. import main as main_module


router = APIRouter()

router.add_api_route("/", main_module.home, methods=["GET"], response_class=HTMLResponse)
router.add_api_route("/health", main_module.healthcheck, methods=["GET"])
router.add_api_route("/ready", main_module.readiness_check, methods=["GET"])
router.add_api_route("/api/c6/inss-client/{cpf}/latest", main_module.c6_latest_inss_client_by_cpf, methods=["GET"])
router.add_api_route("/api/c6/worker-loan/health", main_module.c6_worker_health, methods=["GET"])
router.add_api_route("/api/c6/worker-loan/offer", main_module.c6_generate_worker_offer, methods=["POST"])
router.add_api_route("/api/c6/worker-loan/simulation", main_module.c6_simulate_worker_loan, methods=["POST"])
router.add_api_route("/api/c6/worker-loan/include", main_module.c6_include_worker_loan, methods=["POST"])
