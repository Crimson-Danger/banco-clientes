from __future__ import annotations

from fastapi import APIRouter
from fastapi.responses import HTMLResponse

from .. import main as main_module


router = APIRouter()

router.add_api_route("/users", main_module.users_page, methods=["GET"], response_class=HTMLResponse)
router.add_api_route("/users", main_module.create_user, methods=["POST"])
router.add_api_route("/users/{user_id}", main_module.update_user, methods=["POST"])
router.add_api_route("/users/{user_id}/reset-password", main_module.reset_user_password, methods=["POST"])
router.add_api_route("/users/{user_id}/lock", main_module.lock_user_temporarily, methods=["POST"])
