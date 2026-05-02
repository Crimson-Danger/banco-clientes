from __future__ import annotations

import base64
import json
import time
import urllib.parse
import urllib.request
from dataclasses import dataclass
from typing import Any

from .config import settings


class FactaError(RuntimeError):
    pass


@dataclass
class FactaClient:
    homolog_base: str
    offline_base: str
    username: str
    password: str
    timeout_seconds: int = 25

    def __post_init__(self) -> None:
        self._token_online: str | None = None
        self._token_online_exp: float = 0.0
        self._token_offline: str | None = None
        self._token_offline_exp: float = 0.0

    def _basic_header(self) -> str:
        raw = f"{self.username}:{self.password}".encode("utf-8")
        return "Basic " + base64.b64encode(raw).decode("ascii")

    def _request_json(self, method: str, url: str, headers: dict[str, str] | None = None, form: dict[str, str] | None = None) -> dict[str, Any]:
        body = None
        req_headers = headers.copy() if headers else {}
        if form is not None:
            body = urllib.parse.urlencode(form).encode("utf-8")
            req_headers["Content-Type"] = "application/x-www-form-urlencoded"
        req = urllib.request.Request(url=url, data=body, method=method.upper(), headers=req_headers)
        try:
            with urllib.request.urlopen(req, timeout=self.timeout_seconds) as resp:
                raw = resp.read().decode("utf-8", errors="replace")
        except Exception as exc:
            raise FactaError(f"Falha ao chamar FACTA: {exc}") from exc
        try:
            return json.loads(raw)
        except json.JSONDecodeError as exc:
            raise FactaError(f"Retorno FACTA invalido: {raw[:300]}") from exc

    def _get_token(self, offline: bool) -> str:
        now = time.time()
        if offline:
            if self._token_offline and now < self._token_offline_exp:
                return self._token_offline
            payload = self._request_json("GET", f"{self.offline_base}/gera-token", {"Authorization": self._basic_header()})
            token = str(payload.get("token") or "").strip()
            if not token:
                raise FactaError(str(payload.get("mensagem") or "Token offline nao retornado"))
            self._token_offline = token
            self._token_offline_exp = now + (55 * 60)
            return token

        if self._token_online and now < self._token_online_exp:
            return self._token_online
        payload = self._request_json("GET", f"{self.homolog_base}/gera-token", {"Authorization": self._basic_header()})
        token = str(payload.get("token") or "").strip()
        if not token:
            raise FactaError(str(payload.get("mensagem") or "Token online nao retornado"))
        self._token_online = token
        self._token_online_exp = now + (55 * 60)
        return token

    def _auth_header(self, offline: bool = False) -> dict[str, str]:
        return {"Authorization": f"Bearer {self._get_token(offline=offline)}"}

    def consulta_offline(self, cpf: str) -> dict[str, Any]:
        url = f"{self.offline_base}/clt/base-offline?{urllib.parse.urlencode({'cpf': cpf})}"
        return self._request_json("GET", url, headers=self._auth_header(offline=True))

    def solicita_autorizacao(self, cpf: str, nome: str, celular: str, tipo_envio: str = "WHATSAPP") -> dict[str, Any]:
        url = f"{self.homolog_base}/solicita-autorizacao-consulta"
        return self._request_json(
            "POST",
            url,
            headers=self._auth_header(),
            form={"averbador": "3", "nome": nome, "cpf": cpf, "celular": celular, "tipo_envio": tipo_envio},
        )

    def consulta_online(self, cpf: str) -> dict[str, Any]:
        url = f"{self.homolog_base}/consignado-trabalhador/autoriza-consulta?{urllib.parse.urlencode({'cpf': cpf})}"
        return self._request_json("GET", url, headers=self._auth_header())

    def consulta_ofertas(self, cpf: str) -> dict[str, Any]:
        url = f"{self.homolog_base}/consignado-trabalhador/consulta-ofertas?{urllib.parse.urlencode({'cpf': cpf})}"
        return self._request_json("GET", url, headers=self._auth_header())


def build_facta_client() -> FactaClient:
    if not settings.facta_username or not settings.facta_password:
        raise FactaError("Credenciais FACTA ausentes no ambiente (FACTA_USERNAME/FACTA_PASSWORD).")
    return FactaClient(
        homolog_base=settings.facta_homolog_base,
        offline_base=settings.facta_offline_base,
        username=settings.facta_username,
        password=settings.facta_password,
    )

