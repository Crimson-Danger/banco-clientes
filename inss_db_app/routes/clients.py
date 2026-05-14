from __future__ import annotations

from fastapi import APIRouter
from fastapi.responses import HTMLResponse

from .. import main as main_module


router = APIRouter()

router.add_api_route("/clients", main_module.clients_page, methods=["GET"], response_class=HTMLResponse)
router.add_api_route("/clients/{cpf}", main_module.client_detail_page, methods=["GET"], response_class=HTMLResponse)
router.add_api_route("/clients/favorites", main_module.save_client_filter, methods=["POST"])
router.add_api_route("/clients/favorites/{favorite_id}/delete", main_module.delete_client_filter, methods=["POST"])
router.add_api_route("/clients/{cpf}/do-not-call", main_module.set_client_do_not_call, methods=["POST"])
