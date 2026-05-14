from __future__ import annotations

from datetime import UTC, datetime

from .importer import FilterSet


def format_reference(value: str) -> str:
    if not value or "-" not in value:
        return value
    year, month = value.split("-", 1)
    return f"{month}/{year}"


def format_phone(value: str) -> str:
    digits = "".join(char for char in (value or "") if char.isdigit())
    if len(digits) == 11:
        return f"({digits[:2]}) {digits[2:7]}-{digits[7:]}"
    if len(digits) == 10:
        return f"({digits[:2]}) {digits[2:6]}-{digits[6:]}"
    return value


def format_dashboard_datetime(value) -> str:
    if value is None:
        return "-"
    try:
        return value.strftime("%d/%m/%Y %H:%M")
    except AttributeError:
        return str(value)


def current_utc():
    return datetime.now(UTC)


def parse_optional_int(value: str | None) -> int | None:
    text = str(value or "").strip()
    if not text:
        return None
    try:
        return int(text)
    except ValueError:
        return None


def is_future_timestamp(value: datetime | None) -> bool:
    if value is None:
        return False
    if value.tzinfo is None:
        value = value.replace(tzinfo=UTC)
    return value > current_utc()


def is_safe_internal_path(target: str) -> bool:
    value = (target or "").strip()
    return value.startswith("/") and not value.startswith("//")


def safe_next_path(target: str, default: str = "/") -> str:
    return target if is_safe_internal_path(target) else default


def build_filter_query_params(filters: FilterSet, base_segment: str, selected_matricula: str = "", selected_servico: str = "", page: int | None = None) -> dict[str, str]:
    payload = {
        "base_segment": base_segment,
        "base_source": filters.base_source,
        "quick_mode": filters.quick_mode,
        "quick_value": filters.quick_value,
        "year": str(filters.year or ""),
        "month": str(filters.month or ""),
        "week": filters.week,
        "uf": filters.uf,
        "city": filters.city,
        "esp": filters.esp,
        "margin_min": str(filters.margin_min or ""),
        "margin_max": str(filters.margin_max or ""),
        "age_min": str(filters.age_min or ""),
        "age_max": str(filters.age_max or ""),
        "has_phone": "sim" if filters.has_phone is True else "nao" if filters.has_phone is False else "",
        "source_file": filters.source_file,
        "cpf": filters.cpf,
        "phone": filters.phone,
        "ddd": filters.ddd,
        "name": filters.name,
        "selected_matricula": selected_matricula,
        "selected_servico": selected_servico,
    }
    if page is not None:
        payload["page"] = str(page)
    return payload


def build_pagination(total_items: int, page: int, page_size: int) -> dict[str, int | bool]:
    total_pages = max((total_items + page_size - 1) // page_size, 1)
    safe_page = min(max(page, 1), total_pages)
    return {
        "page": safe_page,
        "page_size": page_size,
        "total_items": total_items,
        "total_pages": total_pages,
        "offset": (safe_page - 1) * page_size,
        "has_prev": safe_page > 1,
        "has_next": safe_page < total_pages,
        "prev_page": max(safe_page - 1, 1),
        "next_page": min(safe_page + 1, total_pages),
    }
