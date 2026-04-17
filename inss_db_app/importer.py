from __future__ import annotations

import csv
import hashlib
import io
import json
import re
import shutil
import subprocess
import tempfile
import unicodedata
from dataclasses import dataclass
from datetime import UTC, date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook
from sqlalchemy import Select, and_, delete, func, or_, select
from sqlalchemy.orm import Session

from .config import settings
from .database import serialized_write
from .models import Client, ClientOccurrence, ExportHistory, ImportBatch, SourceFile
from .schemas import ImportRequest, ImportSummary


EXPORT_HEADERS = [
    "TELEFONE",
    "NOME",
    "CPF",
    "NU-NB",
    "VL-RMI",
    "VL MARGEM",
    "VL RMC",
    "DDB",
    "UF",
    "CIDADE",
    "ESP",
]
AUDIT_HEADERS = EXPORT_HEADERS + ["ARQUIVO_ORIGEM", "SEMANA", "ANO", "MES", "TELEFONE2", "TELEFONE3", "EMAIL"]
SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xlsm", ".xls"}
SQLITE_VARIABLE_CHUNK_SIZE = 900
HEADER_ALIASES = {
    "CPF": {
        "CPF",
        "NU CPF",
        "NUCPF",
        "NUMERO CPF",
        "DOCUMENTO CPF",
    },
    "NOME": {
        "NOME",
        "NOME BENEFICIARIO",
        "NOME DO BENEFICIARIO",
        "NOME X",
        "NOME_X",
    },
    "NU-NB": {
        "NU NB",
        "NU-NB",
        "NB",
        "NUMERO BENEFICIO",
        "NUMERO DO BENEFICIO",
        "NUMERO DE BENEFICIO",
        "BENEFICIO",
    },
    "ESP": {
        "ESP",
        "ESPECIE",
        "CODIGO ESPECIE",
        "COD ESPECIE",
    },
    "DDB": {
        "DDB",
        "DATA DE BENEFICIO",
        "DT BENEFICIO",
        "DT-BENEFICIO",
        "CONCESSAO",
        "DATA CONCESSAO",
        "DATA DE CONCESSAO",
    },
    "VL-RMI": {
        "VL RMI",
        "VL-RMI",
        "RMI",
        "SALARIO",
        "VALOR DO BENEFICIO",
        "VALOR BENEFICIO",
        "MARGEM TOTAL",
    },
    "VL MARGEM": {
        "VL MARGEM",
        "MARGEM",
        "MARGEM DISPONIVEL",
        "VALOR MARGEM",
    },
    "VL RMC": {
        "VL RMC",
        "RMC",
        "MARGEM CARTAO",
    },
    "UF": {"UF", "ESTADO"},
    "CIDADE": {"CIDADE", "MUNICIPIO"},
    "BAIRRO": {"BAIRRO"},
    "LOGRADOURO": {"LOGRADOURO", "ENDERECO", "ENDERECO LOGRADOURO"},
    "NUMERO": {"NUMERO", "NRO"},
    "COMPLEMENTO": {"COMPLEMENTO"},
    "CEP": {"CEP"},
    "TELEFONE1": {
        "TELEFONE1",
        "TELEFONE",
        "TEL",
        "CELULAR",
        "CEL",
        "LEMIT",
    },
    "TELEFONE2": {"TELEFONE2", "CELULAR2", "TEL2", "CEL2"},
    "TELEFONE3": {"TELEFONE3", "CELULAR3", "TEL3", "CEL3"},
    "DT-NASC": {
        "DT-NASC",
        "DT NASC",
        "DT NASCIMENTO",
        "DT_NASC",
        "DT_NASCIMENTO",
        "DATA NASCIMENTO",
        "DATA DE NASCIMENTO",
    },
    "NOME_MAE": {"NOME_MAE", "NOME MAE", "MAE", "NOME DA MAE"},
    "SEXO": {"SEXO", "GENERO"},
    "EMAIL": {"EMAIL", "E-MAIL", "MAIL"},
}
PHONE_HEADER_HINTS = ("TELEFONE", "TEL", "CELULAR", "CEL", "FONE", "LEMIT")


@dataclass
class FilterSet:
    year: int | None = None
    month: int | None = None
    week: str = ""
    ddb_start: str = ""
    ddb_end: str = ""
    uf: str = ""
    city: str = ""
    esp: str = ""
    margin_min: Decimal | None = None
    margin_max: Decimal | None = None
    age_min: int | None = None
    age_max: int | None = None
    has_phone: bool | None = None
    source_file: str = ""
    cpf: str = ""
    phone: str = ""
    ddd: str = ""
    name: str = ""
    quick_mode: str = "cpf"
    quick_value: str = ""


def norm_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip().replace("\ufeff", "")


def strip_accents(text: str) -> str:
    text = norm_text(text)
    return "".join(c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c))


def normalize_header_name(text: str) -> str:
    cleaned = strip_accents(text).upper()
    cleaned = re.sub(r"[^A-Z0-9]+", " ", cleaned)
    return re.sub(r"\s+", " ", cleaned).strip()


def strip_excel_trailing_decimal(text: str) -> str:
    cleaned = norm_text(text)
    if re.fullmatch(r"-?\d+\.0+", cleaned):
        return cleaned.split(".", 1)[0]
    return cleaned


def only_digits(text: str) -> str:
    return re.sub(r"\D", "", strip_excel_trailing_decimal(text))


def normalize_cpf(text: str) -> str:
    digits = only_digits(text)
    if not digits:
        return ""
    if len(digits) > 11:
        digits = digits[-11:]
    if set(digits) == {"0"}:
        return ""
    return digits.zfill(11)


def normalize_phone(text: str) -> str:
    digits = only_digits(text)
    if len(digits) < 10:
        return ""
    if len(digits) > 11:
        digits = digits[-11:]
    if len(set(digits)) == 1:
        return ""
    return digits


def normalize_ddd(text: str) -> str:
    digits = only_digits(text)
    if len(digits) < 2:
        return ""
    return digits[:2]


def parse_ddd_filters(value: str) -> list[str]:
    ddds: list[str] = []
    for item in re.split(r"[,\s;]+", norm_text(value)):
        ddd = normalize_ddd(item)
        if ddd and ddd not in ddds:
            ddds.append(ddd)
    return ddds


def normalize_species(text: str) -> str:
    digits = only_digits(text)
    if not digits:
        return ""
    return digits.lstrip("0") or "0"


def normalize_species_list(values: list[str]) -> list[str]:
    normalized: list[str] = []
    for value in values:
        for item in re.split(r"[,\s;]+", norm_text(value)):
            species = normalize_species(item)
            if species and species not in normalized:
                normalized.append(species)
    return normalized


def parse_species_filter(value: str) -> list[str]:
    return normalize_species_list([value]) if value else []


def normalize_upper(text: str) -> str:
    return norm_text(text).upper()


def parse_money(text: str) -> Decimal | None:
    cleaned = norm_text(text).replace("R$", "").replace(" ", "")
    if not cleaned:
        return None
    if "," in cleaned and "." in cleaned:
        cleaned = cleaned.replace(".", "").replace(",", ".")
    elif "," in cleaned:
        cleaned = cleaned.replace(".", "").replace(",", ".")
    elif "." in cleaned:
        parts = cleaned.split(".")
        if len(parts) > 2:
            cleaned = "".join(parts[:-1]) + "." + parts[-1]
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        return None


def calculate_margin_from_rmi(vl_rmi: Decimal | None, esp: str) -> Decimal | None:
    if vl_rmi is None:
        return None
    rate = Decimal("0.30") if esp in {"87", "88"} else Decimal("0.35")
    return (vl_rmi * rate).quantize(Decimal("0.01"))


def calculate_rmc_from_rmi(vl_rmi: Decimal | None) -> Decimal | None:
    if vl_rmi is None:
        return None
    return (vl_rmi * Decimal("0.05")).quantize(Decimal("0.01"))


def normalize_date_text(value: str) -> str:
    cleaned = norm_text(value)
    if not cleaned:
        return ""
    match = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", cleaned)
    if match:
        day, month, year = match.groups()
        return f"{int(day):02d}/{int(month):02d}/{year}"
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", cleaned):
        year, month, day = cleaned.split("-")
        return f"{day}/{month}/{year}"
    try:
        serial = float(cleaned.replace(",", "."))
    except ValueError:
        return cleaned
    if serial <= 0:
        return cleaned
    excel_epoch = datetime(1899, 12, 30)
    parsed = excel_epoch + timedelta(days=serial)
    return parsed.strftime("%d/%m/%Y")


def normalize_date_filter(value: str) -> str:
    normalized = normalize_date_text(value)
    match = re.fullmatch(r"(\d{2})/(\d{2})/(\d{4})", normalized)
    if not match:
        return ""
    day, month, year = match.groups()
    return f"{year}-{month}-{day}"


def ddb_date_expr():
    return (
        func.substr(ClientOccurrence.ddb, 7, 4)
        + "-"
        + func.substr(ClientOccurrence.ddb, 4, 2)
        + "-"
        + func.substr(ClientOccurrence.ddb, 1, 2)
    )


def birth_date_expr():
    return (
        func.substr(ClientOccurrence.dt_nasc, 7, 4)
        + "-"
        + func.substr(ClientOccurrence.dt_nasc, 4, 2)
        + "-"
        + func.substr(ClientOccurrence.dt_nasc, 1, 2)
    )


def parse_date_value(value: str) -> date | None:
    normalized = normalize_date_text(value)
    match = re.fullmatch(r"(\d{2})/(\d{2})/(\d{4})", normalized)
    if not match:
        return None
    day, month, year = match.groups()
    try:
        return date(int(year), int(month), int(day))
    except ValueError:
        return None


def add_years_safe(value: date, years: int) -> date:
    try:
        return value.replace(year=value.year + years)
    except ValueError:
        return value.replace(month=2, day=28, year=value.year + years)


def compute_age_from_birth(value: str, today: date | None = None) -> int | None:
    birth_date = parse_date_value(value)
    if birth_date is None:
        return None
    reference = today or date.today()
    age = reference.year - birth_date.year
    if (reference.month, reference.day) < (birth_date.month, birth_date.day):
        age -= 1
    return age


def filter_rows_by_ddb_range(
    rows: list[tuple[ClientOccurrence, SourceFile, ImportBatch]],
    ddb_start: str,
    ddb_end: str,
) -> list[tuple[ClientOccurrence, SourceFile, ImportBatch]]:
    if not ddb_start and not ddb_end:
        return rows
    start_date = parse_date_value(ddb_start)
    end_date = parse_date_value(ddb_end)
    if not start_date and not end_date:
        return rows

    filtered: list[tuple[ClientOccurrence, SourceFile, ImportBatch]] = []
    for occurrence, source_file, batch in rows:
        occurrence_date = parse_date_value(occurrence.ddb)
        if occurrence_date is None:
            continue
        if start_date and occurrence_date < start_date:
            continue
        if end_date and occurrence_date > end_date:
            continue
        filtered.append((occurrence, source_file, batch))
    return filtered


def format_money(value: Decimal | None) -> str:
    if value is None:
        return ""
    quantized = value.quantize(Decimal("0.01"))
    integer, decimals = f"{quantized:.2f}".split(".")
    groups = []
    while integer:
        groups.insert(0, integer[-3:])
        integer = integer[:-3]
    return f"R$ {'.'.join(groups)},{decimals}"


def decode_text_file(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1", errors="replace")


def build_header_alias_lookup() -> dict[str, str]:
    lookup: dict[str, str] = {}
    for target, aliases in HEADER_ALIASES.items():
        lookup[normalize_header_name(target)] = target
        for alias in aliases:
            lookup[normalize_header_name(alias)] = target
    return lookup


HEADER_ALIAS_LOOKUP = build_header_alias_lookup()


def canonicalize_row_keys(row: dict[str, str]) -> dict[str, str]:
    canonical: dict[str, str] = {}
    dynamic_phone_columns: list[tuple[str, str]] = []
    for raw_key, raw_value in row.items():
        key = norm_text(raw_key)
        value = norm_text(raw_value)
        if not key:
            continue
        normalized_key = normalize_header_name(key)
        canonical_key = HEADER_ALIAS_LOOKUP.get(normalized_key)
        if canonical_key:
            if value or canonical_key not in canonical:
                canonical[canonical_key] = value
            continue
        if any(hint in normalized_key for hint in PHONE_HEADER_HINTS):
            dynamic_phone_columns.append((normalized_key, value))
    for _normalized_key, value in dynamic_phone_columns:
        if not value:
            continue
        if not canonical.get("TELEFONE1"):
            canonical["TELEFONE1"] = value
        elif not canonical.get("TELEFONE2") and canonical["TELEFONE1"] != value:
            canonical["TELEFONE2"] = value
        elif not canonical.get("TELEFONE3") and canonical["TELEFONE1"] != value and canonical.get("TELEFONE2") != value:
            canonical["TELEFONE3"] = value
    return canonical


def get_first_value(row: dict[str, str], *keys: str) -> str:
    for key in keys:
        value = norm_text(row.get(key, ""))
        if value:
            return value
    return ""


def read_csv_rows(path: Path) -> list[dict[str, str]]:
    text = decode_text_file(path)
    sample = text[:4096]
    delimiter = ";" if sample.count(";") >= sample.count(",") else ","
    reader = csv.DictReader(io.StringIO(text), delimiter=delimiter)
    return [canonicalize_row_keys({norm_text(k): norm_text(v) for k, v in row.items() if k is not None}) for row in reader]


def export_xlsx_to_csv(source_path: Path, temp_dir: Path) -> Path:
    target_path = temp_dir / f"{source_path.stem}.csv"
    powershell_script = rf"""
$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {{
  $wb = $excel.Workbooks.Open('{source_path}')
  if (Test-Path '{target_path}') {{ Remove-Item -LiteralPath '{target_path}' -Force }}
  $wb.SaveAs('{target_path}', 62)
  $wb.Close($false)
}} finally {{
  $excel.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}}
"""
    subprocess.run(["powershell", "-NoProfile", "-Command", powershell_script], check=True, capture_output=True, text=True)
    return target_path


def iter_rows_from_file(path: Path, temp_dir: Path) -> list[dict[str, str]]:
    if path.suffix.lower() == ".csv":
        return read_csv_rows(path)
    if path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        return read_csv_rows(export_xlsx_to_csv(path, temp_dir))
    return []


def detect_phones(row: dict[str, str]) -> list[str]:
    phones: list[str] = []
    for key in ("CELULAR1", "CELULAR2", "CELULAR3", "TELEFONE1", "TELEFONE2", "TELEFONE3"):
        phone = normalize_phone(row.get(key, ""))
        if phone and phone not in phones:
            phones.append(phone)
    return phones[:3]


def week_label_from_ddb(value: str) -> str:
    parsed = parse_date_value(value)
    if parsed is None:
        return ""
    day = parsed.day
    if day <= 10:
        return "01 a 10"
    if day <= 17:
        return "11 a 17"
    if day <= 24:
        return "18 a 24"
    return "25 a 31"


def detect_week_label(rows: list[dict[str, str]]) -> str:
    labels: list[str] = []
    for row in rows:
        label = week_label_from_ddb(get_first_value(row, "DDB"))
        if label and label not in labels:
            labels.append(label)
    if not labels:
        return "SEMANA INDEFINIDA"
    if len(labels) == 1:
        return labels[0]
    return " / ".join(labels)


def compute_file_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while True:
            chunk = handle.read(1024 * 1024)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def compute_row_hash(payload: dict[str, str]) -> str:
    serialized = json.dumps(payload, ensure_ascii=True, sort_keys=True)
    return hashlib.sha256(serialized.encode("utf-8")).hexdigest()


def normalize_row(row: dict[str, str]) -> dict[str, str | Decimal | None]:
    phones = detect_phones(row)
    esp = normalize_species(get_first_value(row, "ESP", "ESPECIE"))
    vl_rmi = parse_money(get_first_value(row, "VL-RMI", "RMI", "SALARIO", "VALOR DO BENEFICIO", "MARGEM TOTAL"))
    vl_margem = parse_money(get_first_value(row, "VL MARGEM", "MARGEM", "MARGEM DISPONIVEL", "VALOR MARGEM"))
    vl_rmc = parse_money(get_first_value(row, "VL RMC", "RMC", "MARGEM CARTAO", " VL RMC "))
    if vl_margem is None:
        vl_margem = calculate_margin_from_rmi(vl_rmi, esp)
    if vl_rmc is None:
        vl_rmc = calculate_rmc_from_rmi(vl_rmi)
    return {
        "cpf": normalize_cpf(get_first_value(row, "CPF")),
        "nome": get_first_value(row, "NOME", "NOME BENEFICIARIO").upper(),
        "nu_nb": only_digits(get_first_value(row, "NU-NB", "NB")),
        "esp": esp,
        "ddb": normalize_date_text(get_first_value(row, "DDB")),
        "vl_rmi": vl_rmi,
        "vl_margem": vl_margem,
        "vl_rmc": vl_rmc,
        "uf": normalize_upper(get_first_value(row, "UF")),
        "cidade": normalize_upper(get_first_value(row, "CIDADE")),
        "bairro": normalize_upper(get_first_value(row, "BAIRRO")),
        "logradouro": normalize_upper(get_first_value(row, "LOGRADOURO")),
        "numero": get_first_value(row, "NUMERO"),
        "complemento": normalize_upper(get_first_value(row, "COMPLEMENTO")),
        "cep": only_digits(get_first_value(row, "CEP")),
        "telefone1": phones[0] if len(phones) > 0 else "",
        "telefone2": phones[1] if len(phones) > 1 else "",
        "telefone3": phones[2] if len(phones) > 2 else "",
        "dt_nasc": normalize_date_text(get_first_value(row, "DT-NASC", "DT_NASCIMENTO", "DT_NASC")),
        "nome_mae": normalize_upper(get_first_value(row, "NOME_MAE")),
        "sexo": normalize_upper(get_first_value(row, "SEXO")),
        "email": get_first_value(row, "EMAIL").lower(),
    }


def ensure_directories() -> None:
    settings.uploads_dir.mkdir(parents=True, exist_ok=True)
    settings.archive_dir.mkdir(parents=True, exist_ok=True)
    settings.exports_dir.mkdir(parents=True, exist_ok=True)


def save_uploaded_files(files: list[tuple[str, bytes]]) -> list[Path]:
    ensure_directories()
    stamp_dir = settings.uploads_dir / datetime.now(UTC).strftime("%Y%m%d_%H%M%S")
    stamp_dir.mkdir(parents=True, exist_ok=True)
    paths: list[Path] = []
    for name, content in files:
        path = stamp_dir / Path(name).name
        path.write_bytes(content)
        paths.append(path)
    return paths


def create_import_batch(session: Session, request: ImportRequest) -> ImportBatch:
    with serialized_write():
        batch = ImportBatch(
            mes_referencia=request.month,
            ano_referencia=request.year,
            origem_pasta=request.origin_folder,
            usuario=request.user_name or "operador",
            status="PROCESSANDO",
        )
        session.add(batch)
        session.flush()
        return batch


def latest_reference_label(year: int, month: int) -> str:
    return f"{year:04d}-{month:02d}"


def prefer_non_empty(current: str, candidate: str) -> str:
    return candidate if candidate else current


def chunked_values(values: set[str] | list[str], chunk_size: int = SQLITE_VARIABLE_CHUNK_SIZE) -> list[list[str]]:
    sequence = list(values)
    return [sequence[index:index + chunk_size] for index in range(0, len(sequence), chunk_size)]


def rebuild_clients_for_cpfs(session: Session, cpfs: set[str]) -> None:
    normalized_cpfs = {cpf for cpf in cpfs if cpf}
    if not normalized_cpfs:
        return

    for cpf_chunk in chunked_values(normalized_cpfs):
        session.execute(delete(Client).where(Client.cpf.in_(cpf_chunk)))

    grouped: dict[str, Client] = {}
    for cpf_chunk in chunked_values(normalized_cpfs):
        rows = session.execute(
            select(
                ClientOccurrence.cpf,
                ClientOccurrence.nome,
                ClientOccurrence.uf,
                ClientOccurrence.cidade,
                ClientOccurrence.telefone1,
                ClientOccurrence.telefone2,
                ClientOccurrence.telefone3,
                ClientOccurrence.vl_margem,
                ImportBatch.ano_referencia,
                ImportBatch.mes_referencia,
            )
            .join(SourceFile, ClientOccurrence.source_file_id == SourceFile.id)
            .join(ImportBatch, SourceFile.batch_id == ImportBatch.id)
            .where(ClientOccurrence.cpf.in_(cpf_chunk))
            .order_by(
                ClientOccurrence.cpf,
                ImportBatch.ano_referencia,
                ImportBatch.mes_referencia,
                SourceFile.id,
                ClientOccurrence.id,
            )
            .execution_options(yield_per=1000)
        )
        for row in rows:
            client = grouped.get(row.cpf)
            if client is None:
                client = Client(
                    cpf=row.cpf,
                    nome_atual=row.nome,
                    uf_atual=row.uf,
                    cidade_atual=row.cidade,
                    melhor_telefone=row.telefone1 or row.telefone2 or row.telefone3,
                    tem_telefone=bool(row.telefone1 or row.telefone2 or row.telefone3),
                    maior_margem=row.vl_margem,
                    ultima_referencia=latest_reference_label(row.ano_referencia, row.mes_referencia),
                    qtd_ocorrencias=0,
                )
                grouped[row.cpf] = client
            client.qtd_ocorrencias += 1
            client.nome_atual = prefer_non_empty(client.nome_atual, row.nome)
            client.uf_atual = prefer_non_empty(client.uf_atual, row.uf)
            client.cidade_atual = prefer_non_empty(client.cidade_atual, row.cidade)
            for candidate in (row.telefone1, row.telefone2, row.telefone3):
                if candidate:
                    client.melhor_telefone = candidate
                    client.tem_telefone = True
                    break
            if row.vl_margem is not None and (client.maior_margem is None or row.vl_margem > client.maior_margem):
                client.maior_margem = row.vl_margem
            client.ultima_referencia = latest_reference_label(row.ano_referencia, row.mes_referencia)
    if grouped:
        session.add_all(grouped.values())
        session.flush()


def delete_batch(session: Session, batch_id: int) -> tuple[bool, str]:
    with serialized_write():
        batch = session.get(ImportBatch, batch_id)
        if batch is None:
            return False, "Lote nao encontrado."
        if batch.status == "PROCESSANDO":
            return False, "Lote em processamento. Aguarde a conclusao para apagar."

        source_files = session.execute(
            select(SourceFile.id, SourceFile.caminho_arquivo).where(SourceFile.batch_id == batch_id)
        ).all()
        source_file_ids = [row.id for row in source_files]
        archived_files = [Path(row.caminho_arquivo) for row in source_files]
        affected_cpfs = set(
            session.scalars(
                select(ClientOccurrence.cpf)
                .where(ClientOccurrence.source_file_id.in_(source_file_ids))
                .distinct()
            ).all()
        ) if source_file_ids else set()

        if source_file_ids:
            session.execute(delete(ClientOccurrence).where(ClientOccurrence.source_file_id.in_(source_file_ids)))
            session.execute(delete(SourceFile).where(SourceFile.id.in_(source_file_ids)))
        session.execute(delete(ImportBatch).where(ImportBatch.id == batch_id))
        rebuild_clients_for_cpfs(session, affected_cpfs)
        session.commit()

    for path in archived_files:
        if path.exists():
            path.unlink(missing_ok=True)

    return True, "Campanha apagada com sucesso."


def process_import_batch(session: Session, batch: ImportBatch, request: ImportRequest) -> ImportSummary:
    ensure_directories()
    summary = ImportSummary(batch_id=batch.id)
    temp_dir = Path(tempfile.mkdtemp(prefix="inss_import_"))
    allowed_species = set(normalize_species_list(request.allowed_species))
    try:
        prepared_files: list[tuple[Path, str, list[dict[str, str]]]] = []
        for file_path in request.files:
            if file_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
                continue
            prepared_files.append((file_path, compute_file_hash(file_path), iter_rows_from_file(file_path, temp_dir)))

        affected_cpfs: set[str] = set()
        with serialized_write():
            for file_path, file_hash, rows in prepared_files:
                if session.scalar(select(func.count()).select_from(SourceFile).where(SourceFile.hash_arquivo == file_hash)):
                    summary.files_skipped += 1
                    summary.duplicate_files.append(file_path.name)
                    continue
                archive_path = settings.archive_dir / f"{datetime.now(UTC):%Y%m%d_%H%M%S}_{file_path.name}"
                shutil.copy2(file_path, archive_path)
                source_file = SourceFile(
                    batch_id=batch.id,
                    nome_arquivo=file_path.name,
                    caminho_arquivo=str(archive_path),
                    semana_label=detect_week_label(rows),
                    tipo_arquivo=file_path.suffix.lower().lstrip("."),
                    hash_arquivo=file_hash,
                    total_linhas=len(rows),
                    duplicado=False,
                )
                session.add(source_file)
                session.flush()
                seen_row_hashes: set[str] = set()
                for raw_row in rows:
                    normalized = normalize_row(raw_row)
                    if not normalized["cpf"]:
                        summary.invalid_rows += 1
                        continue
                    if allowed_species and str(normalized["esp"]) not in allowed_species:
                        summary.species_filtered_rows += 1
                        continue
                    row_hash = compute_row_hash(
                        {
                            "source_file": source_file.nome_arquivo,
                            "cpf": str(normalized["cpf"]),
                            "nu_nb": str(normalized["nu_nb"]),
                            "vl_margem": str(normalized["vl_margem"]),
                            "telefone1": str(normalized["telefone1"]),
                            "telefone2": str(normalized["telefone2"]),
                            "telefone3": str(normalized["telefone3"]),
                        }
                    )
                    if row_hash in seen_row_hashes:
                        summary.invalid_rows += 1
                        continue
                    seen_row_hashes.add(row_hash)
                    session.add(ClientOccurrence(source_file_id=source_file.id, row_hash=row_hash, **normalized))
                    affected_cpfs.add(str(normalized["cpf"]))
                    summary.rows_imported += 1
                summary.files_imported += 1
            batch.status = "CONCLUIDO"
            batch.resumo = (
                f"Arquivos importados: {summary.files_imported}; ignorados: {summary.files_skipped}; "
                f"linhas: {summary.rows_imported}; invalidas: {summary.invalid_rows}; "
                f"fora da especie: {summary.species_filtered_rows}"
            )
            rebuild_clients_for_cpfs(session, affected_cpfs)
            session.commit()
            return summary
    except Exception as exc:
        session.rollback()
        with serialized_write():
            failed_batch = session.get(ImportBatch, batch.id)
            if failed_batch is not None:
                failed_batch.status = "ERRO"
                failed_batch.resumo = f"Falha na importacao: {exc}"
                session.commit()
        raise
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def import_files(session: Session, request: ImportRequest) -> ImportSummary:
    batch = create_import_batch(session, request)
    return process_import_batch(session, batch, request)


def build_occurrence_query(filters: FilterSet) -> Select[tuple[ClientOccurrence, SourceFile, ImportBatch]]:
    stmt = select(ClientOccurrence, SourceFile, ImportBatch).join(SourceFile, ClientOccurrence.source_file_id == SourceFile.id).join(ImportBatch, SourceFile.batch_id == ImportBatch.id)
    conditions = []
    quick_digits = only_digits(filters.quick_value)
    ddb_start = normalize_date_filter(filters.ddb_start)
    ddb_end = normalize_date_filter(filters.ddb_end)
    has_ddb_range = bool(ddb_start or ddb_end)
    if quick_digits:
        if filters.quick_mode == "telefone":
            conditions.append(
                or_(
                    ClientOccurrence.telefone1.ilike(f"%{quick_digits}%"),
                    ClientOccurrence.telefone2.ilike(f"%{quick_digits}%"),
                    ClientOccurrence.telefone3.ilike(f"%{quick_digits}%"),
                )
            )
        else:
            normalized_cpf = normalize_cpf(quick_digits)
            conditions.append(
                or_(
                    ClientOccurrence.cpf == normalized_cpf,
                    ClientOccurrence.nu_nb == quick_digits,
                )
            )
    if filters.year and not has_ddb_range:
        conditions.append(ImportBatch.ano_referencia == filters.year)
    if filters.month and not has_ddb_range:
        conditions.append(ImportBatch.mes_referencia == filters.month)
    if filters.week:
        conditions.append(SourceFile.semana_label.ilike(f"%{filters.week}%"))
    if filters.uf:
        conditions.append(ClientOccurrence.uf == filters.uf.upper())
    if filters.city:
        conditions.append(ClientOccurrence.cidade.ilike(f"%{filters.city.upper()}%"))
    if filters.esp:
        species_filters = parse_species_filter(filters.esp)
        if species_filters:
            conditions.append(ClientOccurrence.esp.in_(species_filters))
    if filters.margin_min is not None:
        conditions.append(ClientOccurrence.vl_margem >= filters.margin_min)
    if filters.margin_max is not None:
        conditions.append(ClientOccurrence.vl_margem <= filters.margin_max)
    if filters.age_min is not None or filters.age_max is not None:
        today = date.today()
        if filters.age_min is not None:
            max_birth_date = add_years_safe(today, -filters.age_min)
            conditions.append(birth_date_expr() <= max_birth_date.isoformat())
        if filters.age_max is not None:
            min_birth_date = add_years_safe(today, -(filters.age_max + 1)) + timedelta(days=1)
            conditions.append(birth_date_expr() >= min_birth_date.isoformat())
    if ddb_start:
        conditions.append(ddb_date_expr() >= ddb_start)
    if ddb_end:
        conditions.append(ddb_date_expr() <= ddb_end)
    if filters.has_phone is True:
        conditions.append(or_(ClientOccurrence.telefone1 != "", ClientOccurrence.telefone2 != "", ClientOccurrence.telefone3 != ""))
    if filters.has_phone is False:
        conditions.append(and_(ClientOccurrence.telefone1 == "", ClientOccurrence.telefone2 == "", ClientOccurrence.telefone3 == ""))
    if filters.source_file:
        conditions.append(SourceFile.nome_arquivo.ilike(f"%{filters.source_file}%"))
    if filters.cpf:
        conditions.append(ClientOccurrence.cpf == normalize_cpf(filters.cpf))
    if filters.phone:
        phone_digits = only_digits(filters.phone)
        if phone_digits:
            conditions.append(
                or_(
                    ClientOccurrence.telefone1.ilike(f"%{phone_digits}%"),
                    ClientOccurrence.telefone2.ilike(f"%{phone_digits}%"),
                    ClientOccurrence.telefone3.ilike(f"%{phone_digits}%"),
                )
            )
    if filters.ddd:
        ddd_filters = parse_ddd_filters(filters.ddd)
        if ddd_filters:
            conditions.append(
                or_(*[
                    ClientOccurrence.telefone1.like(f"{ddd}%")
                    for ddd in ddd_filters
                ])
            )
    if filters.name:
        conditions.append(ClientOccurrence.nome.ilike(f"%{filters.name.upper()}%"))
    if conditions:
        stmt = stmt.where(and_(*conditions))
    return stmt.order_by(ImportBatch.ano_referencia.desc(), ImportBatch.mes_referencia.desc(), ClientOccurrence.nome.asc())


def parse_decimal(value: str) -> Decimal | None:
    cleaned = norm_text(value).replace(",", ".")
    if not cleaned:
        return None
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        return None


def query_clients(session: Session, filters: FilterSet, limit: int = 200) -> list[dict[str, object]]:
    results: list[dict[str, object]] = []
    rows = session.execute(build_occurrence_query(filters).limit(limit)).all()
    for occurrence, source_file, batch in rows:
        display_vl_margem = occurrence.vl_margem
        if display_vl_margem is None:
            display_vl_margem = calculate_margin_from_rmi(occurrence.vl_rmi, occurrence.esp)
        display_vl_rmc = occurrence.vl_rmc
        if display_vl_rmc is None:
            display_vl_rmc = calculate_rmc_from_rmi(occurrence.vl_rmi)
        results.append(
            {
                "cpf": occurrence.cpf,
                "nome": occurrence.nome,
                "nu_nb": occurrence.nu_nb,
                "ddb": occurrence.ddb,
                "dt_nasc": occurrence.dt_nasc,
                "idade": compute_age_from_birth(occurrence.dt_nasc),
                "esp": occurrence.esp,
                "cidade": occurrence.cidade,
                "uf": occurrence.uf,
                "vl_rmi": occurrence.vl_rmi,
                "vl_margem": display_vl_margem,
                "vl_rmc": display_vl_rmc,
                "telefone": occurrence.telefone1 or occurrence.telefone2 or occurrence.telefone3,
                "arquivo": source_file.nome_arquivo,
                "semana": source_file.semana_label,
                "referencia": latest_reference_label(batch.ano_referencia, batch.mes_referencia),
            }
        )
    return results


def _export_rows(rows: list[tuple[ClientOccurrence, SourceFile, ImportBatch]], include_audit: bool, filters: FilterSet) -> list[dict[str, object]]:
    output: list[dict[str, object]] = []
    ddd_filters = parse_ddd_filters(filters.ddd) if filters.ddd else []
    for occurrence, source_file, batch in rows:
        export_phone = occurrence.telefone1 or occurrence.telefone2 or occurrence.telefone3
        if ddd_filters and export_phone and not any(export_phone.startswith(ddd) for ddd in ddd_filters):
            continue
        export_vl_margem = occurrence.vl_margem
        if export_vl_margem is None:
            export_vl_margem = calculate_margin_from_rmi(occurrence.vl_rmi, occurrence.esp)
        export_vl_rmc = occurrence.vl_rmc
        if export_vl_rmc is None:
            export_vl_rmc = calculate_rmc_from_rmi(occurrence.vl_rmi)
        row = {
            "TELEFONE": export_phone,
            "NOME": occurrence.nome,
            "CPF": occurrence.cpf,
            "NU-NB": occurrence.nu_nb,
            "VL-RMI": format_money(occurrence.vl_rmi),
            "VL MARGEM": format_money(export_vl_margem),
            "VL RMC": format_money(export_vl_rmc),
            "DDB": occurrence.ddb,
            "UF": occurrence.uf,
            "CIDADE": occurrence.cidade,
            "ESP": occurrence.esp,
        }
        if include_audit:
            row.update(
                {
                    "ARQUIVO_ORIGEM": source_file.nome_arquivo,
                    "SEMANA": source_file.semana_label,
                    "ANO": batch.ano_referencia,
                    "MES": batch.mes_referencia,
                    "TELEFONE2": occurrence.telefone2,
                    "TELEFONE3": occurrence.telefone3,
                    "EMAIL": occurrence.email,
                }
            )
        output.append(row)
    return output


def export_clients(session: Session, filters: FilterSet, include_audit: bool, file_format: str = "csv") -> Path:
    ensure_directories()
    timestamp = datetime.now(UTC).strftime("%Y%m%d_%H%M%S")
    suffix = "auditoria" if include_audit else "operacional"
    rows = session.execute(build_occurrence_query(filters)).all()
    headers = AUDIT_HEADERS if include_audit else EXPORT_HEADERS
    export_rows = _export_rows(rows, include_audit, filters)
    if file_format == "xlsx":
        export_path = settings.exports_dir / f"base_{suffix}_{timestamp}.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Base"
        sheet.append(headers)
        for row in export_rows:
            sheet.append([row.get(header, "") for header in headers])
        for column in sheet.columns:
            sheet.column_dimensions[column[0].column_letter].width = min(max(len(str(column[0].value or "")) + 2, 14), 28)
        workbook.save(export_path)
    else:
        export_path = settings.exports_dir / f"base_{suffix}_{timestamp}.csv"
        with export_path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=headers, delimiter=";")
            writer.writeheader()
            writer.writerows(export_rows)
    with serialized_write():
        session.add(ExportHistory(filtros_json=json.dumps({**filters.__dict__, "file_format": file_format}, ensure_ascii=False, default=str), arquivo_saida=str(export_path), total_registros=len(rows)))
        session.commit()
    return export_path
