from __future__ import annotations

from dataclasses import dataclass
from datetime import UTC, datetime, timedelta
from threading import Lock
from typing import Any

import httpx

from .config import settings


class C6WorkerLoanError(RuntimeError):
    pass


@dataclass(frozen=True)
class C6Response:
    status_code: int
    payload: dict[str, Any] | list[Any] | str


class C6WorkerLoanClient:
    def __init__(self) -> None:
        self._token: str | None = None
        self._token_expires_at: datetime | None = None
        self._lock = Lock()

    def _require_enabled(self) -> None:
        if not settings.c6_worker_enabled:
            raise C6WorkerLoanError("Integracao C6 desabilitada. Ative C6_WORKER_ENABLED=true.")
        if not settings.c6_username or not settings.c6_password:
            raise C6WorkerLoanError("Credenciais C6 ausentes. Configure C6_USERNAME e C6_PASSWORD.")

    def _token_is_valid(self) -> bool:
        if not self._token or not self._token_expires_at:
            return False
        return datetime.now(UTC) < self._token_expires_at

    def _refresh_token(self) -> str:
        self._require_enabled()
        with self._lock:
            if self._token_is_valid():
                return str(self._token)
            response = httpx.post(
                f"{settings.c6_base_url}/auth/token",
                headers={"Content-Type": "application/x-www-form-urlencoded"},
                data={"username": settings.c6_username, "password": settings.c6_password},
                timeout=settings.c6_timeout_seconds,
            )
            response.raise_for_status()
            payload = response.json()
            token = str(payload.get("access_token", "")).strip()
            expires = int(payload.get("expires_in_seconds", 0) or 0)
            if not token:
                raise C6WorkerLoanError("Token C6 vazio na autenticacao.")
            refresh_window = max(30, min(120, expires // 10))
            self._token = token
            self._token_expires_at = datetime.now(UTC) + timedelta(seconds=max(1, expires - refresh_window))
            return token

    def _auth_token(self) -> str:
        if self._token_is_valid():
            return str(self._token)
        return self._refresh_token()

    def _request(self, path: str, accept: str, payload: dict[str, Any]) -> C6Response:
        self._require_enabled()
        headers = {
            "Accept": accept,
            "Content-Type": "application/json",
            "Authorization": self._auth_token(),
        }
        try:
            response = httpx.post(
                f"{settings.c6_base_url}{path}",
                headers=headers,
                json=payload,
                timeout=settings.c6_timeout_seconds,
            )
        except httpx.HTTPError as exc:
            raise C6WorkerLoanError(f"Falha de comunicacao com C6: {exc}") from exc

        try:
            parsed: dict[str, Any] | list[Any] | str = response.json()
        except ValueError:
            parsed = response.text

        if response.status_code >= 400:
            raise C6WorkerLoanError(f"C6 retornou erro HTTP {response.status_code}: {parsed}")
        return C6Response(status_code=response.status_code, payload=parsed)

    def generate_offer(self, payload: dict[str, Any]) -> C6Response:
        return self._request(
            path="/marketplace/worker-payroll-loan-offers",
            accept="application/vnd.c6bank_generate_offer_v1+json",
            payload=payload,
        )

    def simulate_proposal(self, payload: dict[str, Any], version: str = "v2") -> C6Response:
        accept = "application/vnd.c6bank_simulate_proposal_v2+json" if version.lower() == "v2" else "application/vnd.c6bank_simulate_proposal_v1+json"
        return self._request(
            path="/marketplace/worker-payroll-loan-offers/simulation",
            accept=accept,
            payload=payload,
        )

    def include_proposal(self, payload: dict[str, Any]) -> C6Response:
        return self._request(
            path="/marketplace/worker-payroll-loan-offers/include",
            accept="application/vnd.c6bank_include_proposal_v1+json",
            payload=payload,
        )


c6_worker_loan_client = C6WorkerLoanClient()
