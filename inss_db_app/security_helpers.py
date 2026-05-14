from __future__ import annotations

import hashlib
import hmac

from .config import settings
from .models import AppUser


PERMISSION_LABELS = {
    "can_view_dashboard": "Dashboard",
    "can_import": "Importacao",
    "can_search": "Consulta",
    "can_access_consignado_inss": "Consignado INSS",
    "can_export": "Gerar listas",
    "can_view_history": "Historico",
    "can_delete_batches": "Apagar campanhas",
    "can_manage_users": "Gerenciar usuarios",
}


def user_can(user: AppUser | None, permission: str) -> bool:
    if user is None or not user.is_active:
        return False
    if user.is_admin:
        return True
    return bool(getattr(user, permission, False))


def build_permissions(user: AppUser | None) -> dict[str, bool]:
    permissions = {key: user_can(user, key) for key in PERMISSION_LABELS}
    permissions["is_admin"] = bool(user.is_admin) if user else False
    return permissions


def create_auth_token(user_id: int) -> str:
    payload = str(user_id)
    signature = hmac.new(settings.session_secret_key.encode("utf-8"), payload.encode("utf-8"), hashlib.sha256).hexdigest()
    return f"{payload}:{signature}"


def parse_auth_token(token: str) -> int | None:
    if not token or ":" not in token:
        return None
    payload, signature = token.split(":", 1)
    expected = hmac.new(settings.session_secret_key.encode("utf-8"), payload.encode("utf-8"), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(signature, expected):
        return None
    try:
        return int(payload)
    except ValueError:
        return None
