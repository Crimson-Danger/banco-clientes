from __future__ import annotations

from dataclasses import dataclass

import httpx

from .config import settings


@dataclass(frozen=True)
class CepAddress:
    cep: str = ""
    uf: str = ""
    cidade: str = ""
    bairro: str = ""
    logradouro: str = ""
    complemento: str = ""


@dataclass(frozen=True)
class DddInfo:
    ddd: str = ""
    uf: str = ""
    cidades: tuple[str, ...] = ()


def lookup_viacep_address(cep: str) -> CepAddress | None:
    if len(cep) != 8 or not cep.isdigit():
        return None

    url = f"{settings.viacep_base_url}/{cep}/json/"
    try:
        response = httpx.get(url, timeout=settings.viacep_timeout_seconds)
        response.raise_for_status()
        payload = response.json()
    except (httpx.HTTPError, ValueError):
        return None

    if not isinstance(payload, dict) or payload.get("erro"):
        return None

    return CepAddress(
        cep=str(payload.get("cep", "") or "").replace("-", "").strip(),
        uf=str(payload.get("uf", "") or "").strip(),
        cidade=str(payload.get("localidade", "") or "").strip(),
        bairro=str(payload.get("bairro", "") or "").strip(),
        logradouro=str(payload.get("logradouro", "") or "").strip(),
        complemento=str(payload.get("complemento", "") or "").strip(),
    )


def lookup_brasilapi_cep(cep: str) -> CepAddress | None:
    if len(cep) != 8 or not cep.isdigit():
        return None

    url = f"{settings.brasilapi_base_url}/cep/v1/{cep}"
    try:
        response = httpx.get(url, timeout=settings.brasilapi_timeout_seconds)
        response.raise_for_status()
        payload = response.json()
    except (httpx.HTTPError, ValueError):
        return None

    if not isinstance(payload, dict):
        return None

    return CepAddress(
        cep=str(payload.get("cep", "") or "").replace("-", "").strip(),
        uf=str(payload.get("state", "") or "").strip(),
        cidade=str(payload.get("city", "") or "").strip(),
        bairro=str(payload.get("neighborhood", "") or "").strip(),
        logradouro=str(payload.get("street", "") or "").strip(),
        complemento="",
    )


def lookup_cep_address(cep: str) -> CepAddress | None:
    return lookup_viacep_address(cep) or lookup_brasilapi_cep(cep)


def lookup_brasilapi_ddd(ddd: str) -> DddInfo | None:
    if len(ddd) != 2 or not ddd.isdigit():
        return None

    url = f"{settings.brasilapi_base_url}/ddd/v1/{ddd}"
    try:
        response = httpx.get(url, timeout=settings.brasilapi_timeout_seconds)
        response.raise_for_status()
        payload = response.json()
    except (httpx.HTTPError, ValueError):
        return None

    if not isinstance(payload, dict):
        return None

    raw_cities = payload.get("cities", [])
    cities = tuple(str(city).strip() for city in raw_cities if str(city).strip()) if isinstance(raw_cities, list) else ()
    return DddInfo(
        ddd=ddd,
        uf=str(payload.get("state", "") or "").strip(),
        cidades=cities,
    )
