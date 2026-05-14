from __future__ import annotations

from datetime import date
from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace

from sqlalchemy import create_engine, select
from sqlalchemy.orm import Session, sessionmaker

from inss_db_app.config import settings
from inss_db_app.database import Base
from inss_db_app.importer import (
    FilterSet,
    canonicalize_row_keys,
    backfill_missing_addresses,
    compute_age_from_birth,
    delete_batch,
    detect_week_label,
    export_clients,
    export_cpfs_for_enrichment,
    export_updated_clients,
    export_updated_clients_from_latest_enrichment,
    import_phone_enrichments,
    import_public_campaign_updates,
    import_files,
    normalize_base_segment,
    normalize_date_text,
    normalize_row,
    query_clients,
    week_label_from_ddb,
)
from inss_db_app.models import Client, ClientOccurrence, EnrichmentImportItem, ImportBatch, PhoneEnrichment, SourceFile
from inss_db_app.schemas import ImportRequest


def make_session() -> Session:
    engine = create_engine("sqlite:///:memory:", future=True)
    Base.metadata.create_all(engine)
    return sessionmaker(bind=engine, autoflush=False, autocommit=False, future=True)()


def test_normalize_row() -> None:
    row = {
        "CPF": "123.456.789-00",
        "NOME BENEFICIARIO": "Maria Silva",
        "NU-NB": "12.34",
        "ESP": "087",
        "VL MARGEM": "R$ 455,40",
        "CELULAR1": "(11) 98888-7777",
        "UF": "sp",
        "CIDADE": "Sao Paulo",
    }
    normalized = normalize_row(row)
    assert normalized["cpf"] == "12345678900"
    assert normalized["nome"] == "MARIA SILVA"
    assert normalized["esp"] == "87"
    assert normalized["telefone1"] == "11988887777"
    assert normalized["vl_margem"] == Decimal("455.40")
    assert normalized["vl_rmc"] is None


def test_normalize_row_with_dot_decimal() -> None:
    row = {
        "CPF": "12345678900",
        "NOME": "Maria Silva",
        "VL MARGEM": "531.30",
        "VL-RMI": "1518.00",
        "VL RMC": "75.90",
    }
    normalized = normalize_row(row)
    assert normalized["vl_margem"] == Decimal("531.30")
    assert normalized["vl_rmi"] == Decimal("1518.00")
    assert normalized["vl_rmc"] == Decimal("75.90")


def test_import_enriches_missing_address_from_cep(tmp_path: Path, monkeypatch) -> None:
    csv_path = tmp_path / "cep_enrichment.csv"
    csv_path.write_text(
        "CPF;NOME;CEP;TELEFONE1\n"
        "12345678900;MARIA SILVA;01001000;11999999999\n",
        encoding="utf-8",
    )
    session = make_session()

    patched_settings = dict(vars(settings))
    patched_settings["enable_cep_enrichment"] = True
    monkeypatch.setattr("inss_db_app.importer.settings", SimpleNamespace(**patched_settings))

    class FakeAddress:
        cep = "01001000"
        uf = "SP"
        cidade = "Sao Paulo"
        bairro = "Se"
        logradouro = "Praca da Se"
        complemento = "lado impar"

    monkeypatch.setattr("inss_db_app.importer.lookup_cep_address", lambda cep: FakeAddress() if cep == "01001000" else None)

    summary = import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )

    assert summary.files_imported == 1
    occurrence = session.scalars(select(ClientOccurrence)).one()
    assert occurrence.cep == "01001000"
    assert occurrence.uf == "SP"
    assert occurrence.cidade == "SAO PAULO"
    assert occurrence.bairro == "SE"
    assert occurrence.logradouro == "PRACA DA SE"


def test_import_preserves_existing_address_when_cep_lookup_returns_data(tmp_path: Path, monkeypatch) -> None:
    csv_path = tmp_path / "cep_preserve.csv"
    csv_path.write_text(
        "CPF;NOME;CEP;UF;CIDADE;LOGRADOURO;TELEFONE1\n"
        "12345678900;MARIA SILVA;01001000;RJ;NITEROI;RUA JA PREENCHIDA;11999999999\n",
        encoding="utf-8",
    )
    session = make_session()

    patched_settings = dict(vars(settings))
    patched_settings["enable_cep_enrichment"] = True
    monkeypatch.setattr("inss_db_app.importer.settings", SimpleNamespace(**patched_settings))

    class FakeAddress:
        cep = "01001000"
        uf = "SP"
        cidade = "Sao Paulo"
        bairro = "Se"
        logradouro = "Praca da Se"
        complemento = ""

    monkeypatch.setattr("inss_db_app.importer.lookup_cep_address", lambda cep: FakeAddress() if cep == "01001000" else None)

    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )

    occurrence = session.scalars(select(ClientOccurrence)).one()
    assert occurrence.uf == "RJ"
    assert occurrence.cidade == "NITEROI"
    assert occurrence.logradouro == "RUA JA PREENCHIDA"
    assert occurrence.bairro == "SE"
    assert occurrence.complemento == ""


def test_import_infers_uf_from_ddd_when_missing(tmp_path: Path, monkeypatch) -> None:
    csv_path = tmp_path / "ddd_infer.csv"
    csv_path.write_text(
        "CPF;NOME;TELEFONE1\n"
        "12345678900;MARIA SILVA;21999999999\n",
        encoding="utf-8",
    )
    session = make_session()

    patched_settings = dict(vars(settings))
    patched_settings["enable_ddd_enrichment"] = True
    monkeypatch.setattr("inss_db_app.importer.settings", SimpleNamespace(**patched_settings))

    class FakeDddInfo:
        uf = "RJ"
        cidades = ("Rio de Janeiro", "Niteroi")

    monkeypatch.setattr("inss_db_app.importer.lookup_brasilapi_ddd", lambda ddd: FakeDddInfo() if ddd == "21" else None)

    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )

    occurrence = session.scalars(select(ClientOccurrence)).one()
    assert occurrence.uf == "RJ"
    assert '"UF_INFERIDA_POR_DDD": "RJ"' in occurrence.extras_json


def test_import_registers_regional_alert_when_ddd_disagrees_with_uf(tmp_path: Path, monkeypatch) -> None:
    csv_path = tmp_path / "ddd_alert.csv"
    csv_path.write_text(
        "CPF;NOME;UF;TELEFONE1\n"
        "12345678900;MARIA SILVA;SP;21999999999\n",
        encoding="utf-8",
    )
    session = make_session()

    patched_settings = dict(vars(settings))
    patched_settings["enable_ddd_enrichment"] = True
    monkeypatch.setattr("inss_db_app.importer.settings", SimpleNamespace(**patched_settings))

    class FakeDddInfo:
        uf = "RJ"
        cidades = ("Rio de Janeiro",)

    monkeypatch.setattr("inss_db_app.importer.lookup_brasilapi_ddd", lambda ddd: FakeDddInfo() if ddd == "21" else None)

    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )

    occurrence = session.scalars(select(ClientOccurrence)).one()
    assert occurrence.uf == "SP"
    assert '"ALERTA_REGIONAL": "DDD 21 indica RJ, mas UF informada e SP"' in occurrence.extras_json


def test_backfill_missing_addresses_updates_old_occurrences(monkeypatch) -> None:
    session = make_session()
    batch = ImportBatch(mes_referencia=8, ano_referencia=2025, origem_pasta="teste", usuario="teste", status="CONCLUIDO")
    session.add(batch)
    session.flush()
    source_file = SourceFile(
        batch_id=batch.id,
        nome_arquivo="teste.csv",
        caminho_arquivo="teste.csv",
        tipo_arquivo="csv",
        hash_arquivo="hash-teste",
        total_linhas=1,
    )
    session.add(source_file)
    session.flush()
    occurrence = ClientOccurrence(
        source_file_id=source_file.id,
        cpf="12345678900",
        cpf_original="12345678900",
        nome="MARIA SILVA",
        nome_higienizado="MARIA SILVA",
        cep="01001000",
        telefone1="11999999999",
        row_hash="abc123",
    )
    session.add(occurrence)
    session.commit()

    patched_settings = dict(vars(settings))
    patched_settings["enable_cep_enrichment"] = True
    patched_settings["enable_ddd_enrichment"] = True
    patched_settings["startup_address_backfill_batch_size"] = 100
    monkeypatch.setattr("inss_db_app.importer.settings", SimpleNamespace(**patched_settings))

    class FakeAddress:
        cep = "01001000"
        uf = "SP"
        cidade = "Sao Paulo"
        bairro = "Se"
        logradouro = "Praca da Se"
        complemento = ""

    monkeypatch.setattr("inss_db_app.importer.lookup_cep_address", lambda cep: FakeAddress() if cep == "01001000" else None)
    monkeypatch.setattr("inss_db_app.importer.lookup_brasilapi_ddd", lambda ddd: None)

    summary = backfill_missing_addresses(session)

    updated = session.get(ClientOccurrence, occurrence.id)
    assert summary.scanned_rows == 1
    assert summary.updated_rows == 1
    assert updated is not None
    assert updated.uf == "SP"
    assert updated.cidade == "SAO PAULO"
    assert updated.logradouro == "PRACA DA SE"


def test_backfill_missing_addresses_infers_uf_from_ddd_without_cep(monkeypatch) -> None:
    session = make_session()
    batch = ImportBatch(mes_referencia=8, ano_referencia=2025, origem_pasta="teste", usuario="teste", status="CONCLUIDO")
    session.add(batch)
    session.flush()
    source_file = SourceFile(
        batch_id=batch.id,
        nome_arquivo="teste.csv",
        caminho_arquivo="teste.csv",
        tipo_arquivo="csv",
        hash_arquivo="hash-teste-ddd",
        total_linhas=1,
    )
    session.add(source_file)
    session.flush()
    occurrence = ClientOccurrence(
        source_file_id=source_file.id,
        cpf="99988877766",
        cpf_original="99988877766",
        nome="ANA LIMA",
        nome_higienizado="ANA LIMA",
        telefone1="21999999999",
        row_hash="ddd-only",
    )
    session.add(occurrence)
    session.commit()

    patched_settings = dict(vars(settings))
    patched_settings["enable_cep_enrichment"] = True
    patched_settings["enable_ddd_enrichment"] = True
    patched_settings["startup_address_backfill_batch_size"] = 100
    monkeypatch.setattr("inss_db_app.importer.settings", SimpleNamespace(**patched_settings))

    class FakeDddInfo:
        uf = "RJ"
        cidades = ("Rio de Janeiro",)

    monkeypatch.setattr("inss_db_app.importer.lookup_cep_address", lambda cep: None)
    monkeypatch.setattr("inss_db_app.importer.lookup_brasilapi_ddd", lambda ddd: FakeDddInfo() if ddd == "21" else None)

    summary = backfill_missing_addresses(session)

    updated = session.get(ClientOccurrence, occurrence.id)
    assert summary.scanned_rows == 1
    assert summary.updated_rows == 1
    assert updated is not None
    assert updated.uf == "RJ"
    assert '"UF_INFERIDA_POR_DDD": "RJ"' in updated.extras_json


def test_normalize_date_text_accepts_datetime_strings() -> None:
    assert normalize_date_text("13/11/1963 00:00:00") == "13/11/1963"
    assert normalize_date_text("1963-11-13 00:00:00") == "13/11/1963"


def test_normalize_row_preserves_original_short_cpf() -> None:
    normalized = normalize_row({"CPF": "6034578", "NOME": "DIANA"})
    assert normalized["cpf"] == "00006034578"
    assert normalized["cpf_original"] == "6034578"


def test_normalize_row_calculates_margin_and_rmc_when_missing() -> None:
    row = {
        "CPF": "12345678900",
        "NOME": "Maria Silva",
        "ESP": "21",
        "VL-RMI": "1518,00",
    }
    normalized = normalize_row(row)
    assert normalized["vl_rmi"] == Decimal("1518.00")
    assert normalized["vl_margem"] == Decimal("531.30")
    assert normalized["vl_rmc"] == Decimal("75.90")


def test_normalize_row_calculates_special_margin_for_species_87_88() -> None:
    row = {
        "CPF": "12345678900",
        "NOME": "Maria Silva",
        "ESP": "87",
        "VL-RMI": "1518,00",
    }
    normalized = normalize_row(row)
    assert normalized["vl_margem"] == Decimal("455.40")
    assert normalized["vl_rmc"] == Decimal("75.90")


def test_normalize_row_ignores_excel_trailing_dot_zero_on_phone_and_benefit() -> None:
    row = {
        "CPF": "12345678900.0",
        "NOME": "Maria Silva",
        "NU-NB": "444555666.0",
        "TELEFONE1": "11999999999.0",
    }
    normalized = normalize_row(row)
    assert normalized["cpf"] == "12345678900"
    assert normalized["nu_nb"] == "444555666"
    assert normalized["telefone1"] == "11999999999"


def test_canonicalize_row_keys_accepts_aliases() -> None:
    row = canonicalize_row_keys(
        {
            "nu-cpf": "123.456.789-00",
            "Nome do beneficiario": "Maria Silva",
            "numero do beneficio": "123.456.789",
            "codigo especie": "087",
            "Concessão": "21/01/2026",
            "RMI": "1518,00",
            "Margem Cartao": "75,90",
            "CEL": "(11) 98888-7777",
            "dt_nascimento": "01/01/1960",
        }
    )
    normalized = normalize_row(row)
    assert normalized["cpf"] == "12345678900"
    assert normalized["nome"] == "MARIA SILVA"
    assert normalized["nu_nb"] == "123456789"
    assert normalized["esp"] == "87"
    assert normalized["ddb"] == "21/01/2026"
    assert normalized["vl_rmi"] == Decimal("1518.00")
    assert normalized["vl_margem"] == Decimal("455.40")
    assert normalized["vl_rmc"] == Decimal("75.90")
    assert normalized["telefone1"] == "11988887777"
    assert normalized["dt_nasc"] == "01/01/1960"


def test_normalize_date_text_accepts_single_digit_and_excel_serial() -> None:
    assert normalize_date_text("3/1/2026") == "03/01/2026"
    assert normalize_date_text("2026-01-31") == "31/01/2026"
    assert normalize_date_text("46023") == "01/01/2026"


def test_compute_age_from_birth() -> None:
    assert compute_age_from_birth("16/04/1966", today=date(2026, 4, 16)) == 60
    assert compute_age_from_birth("17/04/1966", today=date(2026, 4, 16)) == 59
    assert compute_age_from_birth("", today=date(2026, 4, 16)) is None


def test_week_label_from_ddb() -> None:
    assert week_label_from_ddb("03/06/2026") == "01 a 10"
    assert week_label_from_ddb("14/06/2026") == "11 a 17"
    assert week_label_from_ddb("22/06/2026") == "18 a 24"
    assert week_label_from_ddb("28/06/2026") == "25 a 31"
    assert week_label_from_ddb("") == ""


def test_detect_week_label_from_rows() -> None:
    assert detect_week_label([{"DDB": "03/06/2026"}, {"DDB": "08/06/2026"}]) == "01 a 10"
    assert detect_week_label([{"DDB": "03/06/2026"}, {"DDB": "14/06/2026"}]) == "01 a 10 / 11 a 17"


def test_import_and_query(tmp_path: Path) -> None:
    csv_path = tmp_path / "01-Janeiro_03a10_SemLemite.csv"
    csv_path.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;111;21;21/01/2026;R$ 100,00;R$ 50,00;R$ 10,00;SP;SAO PAULO;11999999999\n"
        "12345678900;MARIA SILVA;111;21;22/01/2026;R$ 100,00;R$ 70,00;R$ 10,00;SP;SAO PAULO;11888888888\n"
        "22233344455;JOAO SOUZA;87;87;23/01/2026;R$ 200,00;R$ 80,00;R$ 15,00;RJ;RIO DE JANEIRO;\n",
        encoding="utf-8",
    )
    session = make_session()
    summary = import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )
    assert summary.files_imported == 1
    assert summary.rows_imported == 3
    rows = query_clients(session, FilterSet(year=2025, month=8, uf="SP"))
    assert len(rows) == 2
    assert rows[0]["uf"] == "SP"


def test_import_and_query_with_header_aliases(tmp_path: Path) -> None:
    csv_path = tmp_path / "aliases.csv"
    csv_path.write_text(
        "NU-CPF;Nome do beneficiario;Numero do beneficio;Codigo Especie;Concessão;Valor do beneficio;Margem Cartao;CEL;DT_Nascimento;Cidade;UF\n"
        "12345678900;MARIA SILVA;111;21;21/01/2026;1518,00;75,90;11999999999;01/01/1960;SAO PAULO;SP\n",
        encoding="utf-8",
    )
    session = make_session()
    summary = import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )
    assert summary.files_imported == 1
    assert summary.rows_imported == 1
    rows = query_clients(session, FilterSet(year=2025, month=8, uf="SP"))
    assert len(rows) == 1
    assert rows[0]["cpf"] == "12345678900"
    assert rows[0]["telefone"] == "11999999999"
    assert rows[0]["semana"] == "18 a 24"


def test_query_clients_keeps_occurrence_history_ordered_by_reference(tmp_path: Path) -> None:
    jan_path = tmp_path / "jan.csv"
    jan_path.write_text(
        "CPF;NOME;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;SP;SAO PAULO;11999999999\n",
        encoding="utf-8",
    )
    feb_path = tmp_path / "feb.csv"
    feb_path.write_text(
        "CPF;NOME;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;RJ;RIO DE JANEIRO;21999999999\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(session, ImportRequest(year=2026, month=1, user_name="teste", origin_folder=str(tmp_path), files=[jan_path]))
    import_files(session, ImportRequest(year=2026, month=2, user_name="teste", origin_folder=str(tmp_path), files=[feb_path]))

    rows = query_clients(session, FilterSet(cpf="12345678900"))
    assert len(rows) == 2
    assert rows[0]["uf"] == "RJ"
    assert rows[0]["telefone"] == "21999999999"
    assert rows[1]["uf"] == "SP"
    assert rows[1]["telefone"] == "11999999999"


def test_query_clients_pagination_uses_distinct_cpfs(tmp_path: Path) -> None:
    csv_path = tmp_path / "pagination.csv"
    csv_path.write_text(
        "CPF;NOME;UF;CIDADE;TELEFONE1\n"
        "11111111111;CLIENTE 1;SP;SAO PAULO;11999999991\n"
        "22222222222;CLIENTE 2;RJ;RIO DE JANEIRO;21999999992\n"
        "33333333333;CLIENTE 3;MG;BELO HORIZONTE;31999999993\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(session, ImportRequest(year=2026, month=2, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]))

    first_page = query_clients(session, FilterSet(), limit=2, offset=0)
    second_page = query_clients(session, FilterSet(), limit=2, offset=2)

    assert len(first_page) == 2
    assert len(second_page) == 1
    assert {row["cpf"] for row in first_page + second_page} == {"11111111111", "22222222222", "33333333333"}


def test_query_clients_filters_by_phone_and_benefit(tmp_path: Path) -> None:
    csv_path = tmp_path / "phone_benefit.csv"
    csv_path.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;111222333;21;21/01/2026;R$ 100,00;R$ 50,00;R$ 10,00;SP;SAO PAULO;11999999999\n"
        "99988877766;ANA LIMA;444555666;87;22/01/2026;R$ 200,00;R$ 80,00;R$ 15,00;RJ;RIO DE JANEIRO;21988887777\n",
        encoding="utf-8",
    )
    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )

    phone_rows = query_clients(session, FilterSet(phone="9999"))
    assert len(phone_rows) == 1
    assert phone_rows[0]["cpf"] == "12345678900"

    benefit_rows = query_clients(session, FilterSet(quick_mode="cpf", quick_value="444555666"))
    assert len(benefit_rows) == 1
    assert benefit_rows[0]["cpf"] == "99988877766"


def test_query_clients_filters_by_ddd(tmp_path: Path) -> None:
    csv_path = tmp_path / "ddd_filter.csv"
    csv_path.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1;TELEFONE2\n"
        "12345678900;MARIA SILVA;111222333;21;21/01/2026;R$ 100,00;R$ 50,00;R$ 10,00;SP;SAO PAULO;11999999999;\n"
        "99988877766;ANA LIMA;444555666;87;22/01/2026;R$ 200,00;R$ 80,00;R$ 15,00;RJ;RIO DE JANEIRO;21988887777;\n"
        "11122233344;JOSE LIMA;777888999;32;23/01/2026;R$ 300,00;R$ 90,00;R$ 20,00;MG;BELO HORIZONTE;;31977776666\n",
        encoding="utf-8",
    )
    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )

    ddd_rows = query_clients(session, FilterSet(ddd="21"))
    assert len(ddd_rows) == 1
    assert ddd_rows[0]["cpf"] == "99988877766"

    second_phone_ddd_rows = query_clients(session, FilterSet(ddd="31"))
    assert len(second_phone_ddd_rows) == 1
    assert second_phone_ddd_rows[0]["cpf"] == "11122233344"


def test_query_clients_filters_by_multiple_ddds(tmp_path: Path) -> None:
    csv_path = tmp_path / "multiple_ddd_filter.csv"
    csv_path.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1;TELEFONE2\n"
        "12345678900;MARIA SILVA;111222333;21;21/01/2026;R$ 100,00;R$ 50,00;R$ 10,00;SP;SAO PAULO;11999999999;\n"
        "99988877766;ANA LIMA;444555666;87;22/01/2026;R$ 200,00;R$ 80,00;R$ 15,00;RJ;RIO DE JANEIRO;21988887777;\n"
        "11122233344;JOSE LIMA;777888999;32;23/01/2026;R$ 300,00;R$ 90,00;R$ 20,00;MG;BELO HORIZONTE;;31977776666\n",
        encoding="utf-8",
    )
    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )

    rows = query_clients(session, FilterSet(ddd="11,31"))
    assert len(rows) == 2
    assert {row["cpf"] for row in rows} == {"12345678900", "11122233344"}


def test_query_clients_recalculates_missing_rmc_for_legacy_rows(tmp_path: Path) -> None:
    csv_path = tmp_path / "legacy_query_rmc.csv"
    csv_path.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;111;21;21/01/2026;R$ 1.518,00;R$ 531,30;SP;SAO PAULO;11999999999\n",
        encoding="utf-8",
    )
    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )
    occurrence = session.query(ClientOccurrence).one()
    occurrence.vl_rmc = None
    session.commit()

    rows = query_clients(session, FilterSet(cpf="12345678900"))
    assert len(rows) == 1
    assert rows[0]["vl_rmc"] == Decimal("75.90")


def test_query_clients_filters_by_ddb_range(tmp_path: Path) -> None:
    csv_path = tmp_path / "ddb_range.csv"
    csv_path.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;111222333;21;05/01/2026;R$ 100,00;R$ 50,00;R$ 10,00;SP;SAO PAULO;11999999999\n"
        "99988877766;ANA LIMA;444555666;87;02/02/2026;R$ 200,00;R$ 80,00;R$ 15,00;RJ;RIO DE JANEIRO;21988887777\n",
        encoding="utf-8",
    )
    session = make_session()
    import_files(
        session,
        ImportRequest(year=2026, month=1, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )

    rows = query_clients(session, FilterSet(ddb_start="01/01/2026", ddb_end="31/01/2026"))
    assert len(rows) == 1
    assert rows[0]["cpf"] == "12345678900"


def test_query_clients_filters_by_ddb_range_with_unpadded_dates(tmp_path: Path) -> None:
    csv_path = tmp_path / "ddb_unpadded.csv"
    csv_path.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;111222333;21;3/1/2026;R$ 100,00;R$ 50,00;R$ 10,00;SP;SAO PAULO;11999999999\n"
        "99988877766;ANA LIMA;444555666;87;31/1/2026;R$ 200,00;R$ 80,00;R$ 15,00;RJ;RIO DE JANEIRO;21988887777\n",
        encoding="utf-8",
    )
    session = make_session()
    import_files(
        session,
        ImportRequest(year=2026, month=1, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )

    rows = query_clients(session, FilterSet(ddb_start="01/01/2026", ddb_end="31/01/2026"))
    assert len(rows) == 2


def test_query_clients_filters_by_age_range(tmp_path: Path) -> None:
    today = date.today()
    exact_sixty = today.replace(year=today.year - 60)
    still_fifty_nine = today.replace(year=today.year - 60) if (today.month, today.day) == (12, 31) else date(today.year - 60, today.month, min(today.day + 1, 28))
    csv_path = tmp_path / "age_range.csv"
    csv_path.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;DT-NASC;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1\n"
        f"12345678900;MARIA SILVA;111222333;21;05/01/2026;{exact_sixty.strftime('%d/%m/%Y')};R$ 100,00;R$ 50,00;R$ 10,00;SP;SAO PAULO;11999999999\n"
        f"99988877766;ANA LIMA;444555666;87;02/02/2026;{still_fifty_nine.strftime('%d/%m/%Y')};R$ 200,00;R$ 80,00;R$ 15,00;RJ;RIO DE JANEIRO;21988887777\n"
        "11122233344;JOSE LIMA;777888999;32;02/02/2026;01/01/1980;R$ 200,00;R$ 80,00;R$ 15,00;RJ;RIO DE JANEIRO;21988887778\n",
        encoding="utf-8",
    )
    session = make_session()
    import_files(
        session,
        ImportRequest(year=2026, month=1, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )

    rows = query_clients(session, FilterSet(age_min=60, age_max=60))
    assert len(rows) == 1
    assert rows[0]["cpf"] == "12345678900"
    assert rows[0]["idade"] == 60


def test_query_clients_filters_multiple_species(tmp_path: Path) -> None:
    csv_path = tmp_path / "multiple_species.csv"
    csv_path.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;111222333;21;05/01/2026;R$ 100,00;R$ 50,00;R$ 10,00;SP;SAO PAULO;11999999999\n"
        "99988877766;ANA LIMA;444555666;87;02/02/2026;R$ 200,00;R$ 80,00;R$ 15,00;RJ;RIO DE JANEIRO;21988887777\n"
        "11122233344;JOSE LIMA;777888999;32;02/02/2026;R$ 200,00;R$ 80,00;R$ 15,00;RJ;RIO DE JANEIRO;21988887778\n",
        encoding="utf-8",
    )
    session = make_session()
    import_files(
        session,
        ImportRequest(year=2026, month=1, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )

    rows = query_clients(session, FilterSet(esp="21,87"))
    assert len(rows) == 2


def test_query_clients_ddb_range_overrides_year_and_month_filters(tmp_path: Path) -> None:
    csv_path = tmp_path / "ddb_override.csv"
    csv_path.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;111222333;21;05/01/2026;R$ 100,00;R$ 50,00;R$ 10,00;SP;SAO PAULO;11999999999\n",
        encoding="utf-8",
    )
    session = make_session()
    import_files(
        session,
        ImportRequest(year=2026, month=1, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )

    rows = query_clients(
        session,
        FilterSet(year=2025, month=8, ddb_start="01/01/2026", ddb_end="31/01/2026"),
    )
    assert len(rows) == 1


def test_export_clients(tmp_path: Path) -> None:
    csv_path = tmp_path / "agosto.csv"
    csv_path.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;111;21;21/01/2026;R$ 100,00;R$ 50,00;R$ 10,00;SP;SAO PAULO;11999999999\n",
        encoding="utf-8",
    )
    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )
    export_path = export_clients(session, FilterSet(uf="SP"), include_audit=False)
    assert export_path.exists()
    content = export_path.read_text(encoding="utf-8-sig")
    assert "MARIA SILVA" in content


def test_export_clients_falls_back_to_consolidated_uf_and_city(tmp_path: Path) -> None:
    csv_path = tmp_path / "export_consolidated_fallback.csv"
    csv_path.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;TELEFONE1\n"
        "12345678900;MARIA SILVA;111;21;21/01/2026;R$ 100,00;R$ 50,00;93991859770\n",
        encoding="utf-8",
    )
    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )
    client = session.get(Client, "12345678900")
    assert client is not None
    client.uf_atual = "PA"
    client.cidade_atual = "SANTAREM"
    session.commit()

    export_path = export_clients(session, FilterSet(), include_audit=False)
    content = export_path.read_text(encoding="utf-8-sig")
    assert "PA" in content
    assert "SANTAREM" in content


def test_export_clients_uses_first_phone_with_ddd_filter(tmp_path: Path) -> None:
    csv_path = tmp_path / "export_ddd.csv"
    csv_path.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1;TELEFONE2\n"
        "12345678900;MARIA SILVA;111;21;21/01/2026;R$ 100,00;R$ 50,00;R$ 10,00;SP;SAO PAULO;11999999999;21988887777\n",
        encoding="utf-8",
    )
    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )
    export_path = export_clients(session, FilterSet(ddd="21"), include_audit=False)
    content = export_path.read_text(encoding="utf-8-sig")
    assert "MARIA SILVA" not in content


def test_export_clients_recalculates_missing_rmc_for_legacy_rows(tmp_path: Path) -> None:
    csv_path = tmp_path / "legacy_rmc.csv"
    csv_path.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;111;21;21/01/2026;R$ 1.518,00;R$ 531,30;SP;SAO PAULO;11999999999\n",
        encoding="utf-8",
    )
    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )
    occurrence = session.query(ClientOccurrence).one()
    occurrence.vl_rmc = None
    session.commit()

    export_path = export_clients(session, FilterSet(uf="SP"), include_audit=False)
    content = export_path.read_text(encoding="utf-8-sig")
    assert "R$ 75,90" in content


def test_export_clients_xlsx(tmp_path: Path) -> None:
    csv_path = tmp_path / "agosto.xlsx.csv"
    csv_path.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;111;21;27/01/2026;R$ 1.518,00;R$ 455,40;R$ 75,90;SP;SAO PAULO;11999999999\n",
        encoding="utf-8",
    )
    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )
    export_path = export_clients(session, FilterSet(uf="SP"), include_audit=False, file_format="xlsx")
    assert export_path.exists()
    assert export_path.suffix == ".xlsx"


def test_export_updated_clients_keeps_benefit_fields_from_latest_inss_occurrence(tmp_path: Path) -> None:
    old_csv = tmp_path / "benefit_old.csv"
    old_csv.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;111;31;10/01/2026;R$ 1000,00;R$ 50,00;R$ 10,00;SP;SAO PAULO;11999999999\n",
        encoding="utf-8",
    )
    new_csv = tmp_path / "benefit_new.csv"
    new_csv.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;222;32;15/02/2026;R$ 2000,00;R$ 90,00;R$ 20,00;RJ;RIO DE JANEIRO;\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(session, ImportRequest(year=2026, month=1, user_name="teste", origin_folder=str(tmp_path), files=[old_csv]))
    import_files(session, ImportRequest(year=2026, month=2, user_name="teste", origin_folder=str(tmp_path), files=[new_csv]))

    export_path = export_updated_clients(session, FilterSet(), file_format="csv")
    content = export_path.read_text(encoding="utf-8-sig")

    assert "12345678900" in content
    assert "222" in content
    assert "32" in content
    assert "15/02/2026" in content
    assert "R$ 90,00" in content
    assert "R$ 20,00" in content
    assert "11999999999" in content
    assert "RIO DE JANEIRO" in content


def test_export_updated_clients_from_latest_enrichment_uses_last_return(tmp_path: Path) -> None:
    source_csv = tmp_path / "source_updated.csv"
    source_csv.write_text(
        "CPF;NOME;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;SP;SAO PAULO;11999999999\n",
        encoding="utf-8",
    )
    return_csv = tmp_path / "return_updated.csv"
    return_csv.write_text(
        "CPF;NOME;NOME_MAE;SEXO;NASC;RENDA;LOGRADOURO;NUMERO;COMPLEMENTO;BAIRRO;CIDADE;UF;CEP;CEL1;FLGWHATSCEL1;EMAIL1\n"
        "12345678900;MARIA SILVA;JOANA SILVA;F;1980-05-19 00:00:00;3200,50;RUA A;45;AP 2;CENTRO;CAMPINAS;SP;13000000;11911112222;SIM;maria@exemplo.com\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(session, ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[source_csv]))
    import_phone_enrichments(session, [return_csv], source_name="NOVAVIDA", imported_by="tester")

    export_path = export_updated_clients_from_latest_enrichment(session, base_segment="INSS", file_format="csv")
    content = export_path.read_text(encoding="utf-8-sig")
    assert "11911112222" in content
    assert "SIM" in content
    assert "maria@exemplo.com" in content
    assert "CAMPINAS" in content
    assert "RUA A, 45, AP 2, CENTRO" in content


def test_export_updated_clients_from_latest_enrichment_includes_cpfs_without_new_phone_log(tmp_path: Path) -> None:
    source_csv = tmp_path / "source_latest.csv"
    source_csv.write_text(
        "CPF;NOME;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;SP;SAO PAULO;11999999999\n"
        "22233344455;JOAO LIMA;SP;SANTOS;11988887777\n",
        encoding="utf-8",
    )
    return_csv = tmp_path / "return_latest.csv"
    return_csv.write_text(
        "CPF;NOME;CIDADE;UF;EMAIL1\n"
        "12345678900;MARIA SILVA;CAMPINAS;SP;maria@exemplo.com\n"
        "22233344455;JOAO LIMA;GUARUJA;SP;joao@exemplo.com\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(session, ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[source_csv]))
    import_phone_enrichments(session, [return_csv], source_name="NOVAVIDA", imported_by="tester")

    export_path = export_updated_clients_from_latest_enrichment(session, base_segment="INSS", file_format="csv")
    content = export_path.read_text(encoding="utf-8-sig")
    assert "12345678900" in content
    assert "22233344455" in content
    assert "maria@exemplo.com" in content
    assert "joao@exemplo.com" in content


def test_delete_batch(tmp_path: Path) -> None:
    csv_path = tmp_path / "delete.csv"
    csv_path.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;111;21;21/01/2026;R$ 100,00;R$ 50,00;R$ 10,00;SP;SAO PAULO;11999999999\n",
        encoding="utf-8",
    )
    session = make_session()
    summary = import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )
    ok, _ = delete_batch(session, summary.batch_id)
    assert ok is True
    rows = query_clients(session, FilterSet(year=2025, month=8))
    assert rows == []


def test_delete_batch_rebuilds_only_affected_cpfs(tmp_path: Path) -> None:
    csv_batch_1 = tmp_path / "batch1.csv"
    csv_batch_1.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;111;21;21/01/2026;R$ 100,00;R$ 50,00;R$ 10,00;SP;SAO PAULO;11999999999\n"
        "55566677788;ANA LIMA;222;87;22/01/2026;R$ 200,00;R$ 80,00;R$ 15,00;MG;BELO HORIZONTE;31999999999\n",
        encoding="utf-8",
    )
    csv_batch_2 = tmp_path / "batch2.csv"
    csv_batch_2.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;333;21;23/02/2026;R$ 100,00;R$ 90,00;R$ 10,00;RJ;RIO DE JANEIRO;21999999999\n",
        encoding="utf-8",
    )

    session = make_session()
    first_summary = import_files(
        session,
        ImportRequest(year=2025, month=1, user_name="teste", origin_folder=str(tmp_path), files=[csv_batch_1]),
    )
    import_files(
        session,
        ImportRequest(year=2025, month=2, user_name="teste", origin_folder=str(tmp_path), files=[csv_batch_2]),
    )

    ok, _ = delete_batch(session, first_summary.batch_id)
    assert ok is True

    rows = query_clients(session, FilterSet())
    assert len(rows) == 1
    assert rows[0]["cpf"] == "12345678900"
    assert rows[0]["uf"] == "RJ"


def test_import_files_with_allowed_species(tmp_path: Path) -> None:
    csv_path = tmp_path / "species.csv"
    csv_path.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;111;21;21/01/2026;R$ 100,00;R$ 50,00;R$ 10,00;SP;SAO PAULO;11999999999\n"
        "22233344455;JOAO SOUZA;222;87;23/01/2026;R$ 200,00;R$ 80,00;R$ 15,00;RJ;RIO DE JANEIRO;21999999999\n",
        encoding="utf-8",
    )
    session = make_session()
    summary = import_files(
        session,
        ImportRequest(
            year=2025,
            month=8,
            user_name="teste",
            origin_folder=str(tmp_path),
            files=[csv_path],
            allowed_species=["21"],
        ),
    )
    assert summary.rows_imported == 1
    assert summary.species_filtered_rows == 1
    rows = query_clients(session, FilterSet(year=2025, month=8))
    assert len(rows) == 1
    assert rows[0]["esp"] == "21"


def test_import_skips_duplicate_rows_in_same_file(tmp_path: Path) -> None:
    csv_path = tmp_path / "duplicadas.csv"
    csv_path.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;111;21;21/01/2026;R$ 100,00;R$ 50,00;R$ 10,00;SP;SAO PAULO;11999999999\n"
        "12345678900;MARIA SILVA;111;21;21/01/2026;R$ 100,00;R$ 50,00;R$ 10,00;SP;SAO PAULO;11999999999\n",
        encoding="utf-8",
    )
    session = make_session()
    summary = import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]),
    )
    assert summary.rows_imported == 1
    assert summary.invalid_rows == 1
    rows = query_clients(session, FilterSet(year=2025, month=8))
    assert len(rows) == 1


def test_export_cpfs_for_enrichment_deduplicates_cpfs(tmp_path: Path) -> None:
    csv_path = tmp_path / "cpf_export.csv"
    csv_path.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;111;21;21/01/2026;100;50;10;SP;SAO PAULO;11999999999\n"
        "12345678900;MARIA SILVA;222;21;22/01/2026;100;50;10;SP;SAO PAULO;11888888888\n"
        "98765432100;JOSE LIMA;333;21;23/01/2026;100;50;10;RJ;RIO;21999999999\n",
        encoding="utf-8",
    )
    session = make_session()
    import_files(session, ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path]))

    export_path = export_cpfs_for_enrichment(session, FilterSet(), file_format="csv")
    content = export_path.read_text(encoding="utf-8-sig")
    assert content.count("12345678900") == 1
    assert content.count("98765432100") == 1


def test_import_phone_enrichments_updates_occurrences_and_logs_history(tmp_path: Path) -> None:
    source_csv = tmp_path / "origem.csv"
    source_csv.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1;TELEFONE2\n"
        "12345678900;MARIA SILVA;111;21;21/01/2026;100;50;10;SP;SAO PAULO;11999999999;1133334444\n",
        encoding="utf-8",
    )
    return_csv = tmp_path / "retorno.csv"
    return_csv.write_text(
        "CPF;LEMIT;TELEFONE2\n"
        "12345678900;11911112222;11922223333\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(session, ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[source_csv]))

    summary = import_phone_enrichments(session, [return_csv], source_name="LEMIT", imported_by="tester")
    assert summary.clients_matched == 1
    assert summary.clients_updated == 1
    assert summary.phones_added == 2

    occurrence = session.execute(select(ClientOccurrence).where(ClientOccurrence.cpf == "12345678900")).scalar_one()
    assert occurrence.telefone1 == "11911112222"
    assert occurrence.telefone2 == "11922223333"
    assert occurrence.telefone3 == "11999999999"

    client = session.get(Client, "12345678900")
    assert client is not None
    assert client.melhor_telefone == "11911112222"
    assert client.tem_telefone is True

    history = session.execute(select(PhoneEnrichment).order_by(PhoneEnrichment.telefone)).scalars().all()
    assert [item.telefone for item in history] == ["11911112222", "11922223333"]
    imported_items = session.execute(select(EnrichmentImportItem).where(EnrichmentImportItem.arquivo_origem == "retorno.csv")).scalars().all()
    assert len(imported_items) == 1


def test_import_phone_enrichments_updates_other_available_fields(tmp_path: Path) -> None:
    source_csv = tmp_path / "origem_enriquecimento.csv"
    source_csv.write_text(
        "CPF;NOME;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;SP;SAO PAULO;11999999999\n",
        encoding="utf-8",
    )
    return_csv = tmp_path / "retorno_novavida.csv"
    return_csv.write_text(
        "CPF;NOME;NOME_MAE;SEXO;NASC;RENDA;TIPO;TITULO;LOGRADOURO;NUMERO;COMPLEMENTO;BAIRRO;CIDADE;UF;CEP;CEL1;FLGWHATSCEL1;EMAIL1;EMAIL2\n"
        "12345678900;MARIA SILVA;JOANA SILVA;F;1980-05-19 00:00:00;3200,50;EFETIVO;ANALISTA;RUA A;45;AP 2;CENTRO;CAMPINAS;SP;13000000;11911112222;SIM;maria@exemplo.com;maria2@exemplo.com\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(session, ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[source_csv]))

    summary = import_phone_enrichments(session, [return_csv], source_name="NOVAVIDA", imported_by="tester")
    assert summary.clients_matched == 1
    assert summary.clients_updated == 1
    assert summary.data_fields_updated == 1

    occurrence = session.execute(select(ClientOccurrence).where(ClientOccurrence.cpf == "12345678900")).scalar_one()
    assert occurrence.nome_mae == "JOANA SILVA"
    assert occurrence.sexo == "F"
    assert occurrence.dt_nasc == "19/05/1980"
    assert occurrence.logradouro == "RUA A"
    assert occurrence.numero == "45"
    assert occurrence.complemento == "AP 2"
    assert occurrence.bairro == "CENTRO"
    assert occurrence.cidade == "CAMPINAS"
    assert occurrence.cep == "13000000"
    assert occurrence.email == "maria@exemplo.com"
    assert occurrence.telefone1 == "11911112222"
    assert "FLGWHATSCEL1" in occurrence.extras_json

    rows = query_clients(session, FilterSet(base_segment="INSS", cpf="12345678900"))
    assert len(rows) == 1
    assert rows[0]["nome_mae"] == "JOANA SILVA"
    assert rows[0]["email"] == "maria@exemplo.com"
    assert rows[0]["cidade"] == "CAMPINAS"
    assert rows[0]["whatsapp_telefone1"] == "SIM"


def test_normalize_row_supports_government_fields_and_extras() -> None:
    row = canonicalize_row_keys(
        {
            "CPF": "12345678900",
            "NOME HIGIENIZADO": "MARIA SILVA",
            "MATRÍCULA": "ABC123",
            "PIS": "123.45678.90-1",
            "ENTIDADE": "PREFEITURA",
            "SECRETARIA": "SAUDE",
            "Convenio": "GOV BA",
            "Celular Atual": "(71) 99111-2222",
            "FIXO2": "(71) 3333-4444",
            "NASC": "01/02/1980",
            "CARGO EXTRA": "ANALISTA",
        }
    )
    normalized = normalize_row(row)
    assert normalized["nome"] == "MARIA SILVA"
    assert normalized["matricula"] == "ABC123"
    assert normalized["pis"] == "12345678901"
    assert normalized["telefone1"] == "71991112222"
    assert normalized["telefone2"] == "7133334444"
    assert "CARGO EXTRA" in str(normalized["extras_json"])


def test_normalize_base_segment_accepts_aliases() -> None:
    assert normalize_base_segment("gov") == "GOVERNO"
    assert normalize_base_segment("pref") == "PREFEITURA"
    assert normalize_base_segment("inss") == "INSS"
    assert normalize_base_segment("qualquer") == "INSS"


def test_import_files_government_stays_separated_from_inss_queries(tmp_path: Path) -> None:
    inss_csv = tmp_path / "inss.csv"
    inss_csv.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1\n"
        "11122233344;CLIENTE INSS;111;21;21/01/2026;100;50;10;BA;SALVADOR;71999999999\n",
        encoding="utf-8",
    )
    gov_csv = tmp_path / "gov.csv"
    gov_csv.write_text(
        "CPF;NOME HIGIENIZADO;MATRÍCULA;PIS;ENTIDADE;Celular Atual;NASC;UF;CIDADE\n"
        "11122233344;CLIENTE GOV;MAT-1;12345678901;PREFEITURA;71988887777;01/02/1980;BA;JUAZEIRO\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[inss_csv], base_segment="INSS"),
    )
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[gov_csv], base_segment="GOV", base_source="BAHIA"),
    )

    inss_rows = query_clients(session, FilterSet())
    gov_rows = query_clients(session, FilterSet(base_segment="GOV"))
    assert len(inss_rows) == 1
    assert len(gov_rows) == 1
    assert inss_rows[0]["telefone"] == "71999999999"
    assert gov_rows[0]["telefone"] == "71988887777"

    client = session.get(Client, "11122233344")
    assert client is not None
    assert client.tem_inss is True
    assert client.tem_governo is True
    assert client.matricula_atual == "MAT-1"
    assert client.pis_atual == "12345678901"


def test_import_files_prefeitura_stays_separated_from_government_and_inss(tmp_path: Path) -> None:
    gov_csv = tmp_path / "gov.csv"
    gov_csv.write_text(
        "CPF;NOME HIGIENIZADO;MATRICULA;PIS;ENTIDADE;Celular Atual;NASC;UF;CIDADE\n"
        "33322211100;CLIENTE GOV;MAT-GOV;12345678901;GOVERNO ESTADUAL;71988887777;01/02/1980;BA;SALVADOR\n",
        encoding="utf-8",
    )
    prefeitura_csv = tmp_path / "prefeitura.csv"
    prefeitura_csv.write_text(
        "CPF;NOME DO SERVIDOR;MATRICULA;ENTIDADE;SECRETARIA;TELEFONE_01;NASC;UF;CIDADE\n"
        "33322211100;CLIENTE PREF;MAT-PREF;PREFEITURA DE JUAZEIRO;SAUDE;74991112222;01/02/1980;BA;JUAZEIRO\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[gov_csv], base_segment="GOVERNO", base_source="BAHIA"),
    )
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[prefeitura_csv], base_segment="PREFEITURA", base_source="JUAZEIRO"),
    )

    gov_rows = query_clients(session, FilterSet(base_segment="GOVERNO"))
    prefeitura_rows = query_clients(session, FilterSet(base_segment="PREFEITURA"))
    assert len(gov_rows) == 1
    assert len(prefeitura_rows) == 1
    assert gov_rows[0]["telefone"] == "71988887777"
    assert prefeitura_rows[0]["telefone"] == "74991112222"
    assert gov_rows[0]["matricula"] == "MAT-GOV"
    assert prefeitura_rows[0]["matricula"] == "MAT-PREF"

    client = session.get(Client, "33322211100")
    assert client is not None
    assert client.tem_inss is False
    assert client.tem_governo is True


def test_query_clients_government_does_not_recalculate_margin_or_rmc(tmp_path: Path) -> None:
    gov_csv = tmp_path / "gov_margin.csv"
    gov_csv.write_text(
        "CPF;NOME HIGIENIZADO;MATRICULA;SALARIO;Margem Liquida;Margem Bruta;Margem Utilizada;UF;CIDADE;Celular Atual;NASC\n"
        "55544433322;CLIENTE GOV;MAT-55;2500,00;300,00;500,00;200,00;CE;FORTALEZA;85991112222;01/02/1980\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[gov_csv], base_segment="GOVERNO", base_source="CEARA"),
    )

    occurrence = session.execute(select(ClientOccurrence).where(ClientOccurrence.cpf == "55544433322")).scalar_one()
    assert occurrence.vl_margem is None
    assert occurrence.vl_rmc is None

    rows = query_clients(session, FilterSet(base_segment="GOVERNO", cpf="55544433322"))
    assert len(rows) == 1
    assert rows[0]["vl_margem"] is None
    assert rows[0]["vl_rmc"] is None
    assert rows[0]["margem_liquida"] == Decimal("300.00")
    assert rows[0]["margem_bruta"] == Decimal("500.00")
    assert rows[0]["margem_utilizada"] == Decimal("200.00")


def test_query_clients_government_quick_search_accepts_matricula(tmp_path: Path) -> None:
    gov_csv = tmp_path / "gov_matricula.csv"
    gov_csv.write_text(
        "CPF;NOME HIGIENIZADO;MATRICULA;UF;CIDADE;Celular Atual;NASC\n"
        "77766655544;CLIENTE GOV;MAT12345;MA;SAO LUIS;98991112222;01/02/1980\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[gov_csv], base_segment="GOVERNO", base_source="MARANHAO"),
    )

    rows = query_clients(session, FilterSet(base_segment="GOVERNO", quick_mode="cpf", quick_value="12345"))
    assert len(rows) == 1
    assert rows[0]["cpf"] == "77766655544"


def test_query_clients_accepts_short_original_cpf_search(tmp_path: Path) -> None:
    gov_csv = tmp_path / "gov_short_cpf.csv"
    gov_csv.write_text(
        "CPF;NOME;UF;CIDADE;TELEFONE_01;NASC\n"
        "6034578;DIANA GLEISS OLIVEIRA GUIMARAES;BA;LENCOIS;75988358243;1980-05-19 00:00:00\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[gov_csv], base_segment="GOVERNO", base_source="BAHIA"),
    )

    rows = query_clients(session, FilterSet(base_segment="GOVERNO", quick_mode="cpf", quick_value="6034578"))
    assert len(rows) == 1
    assert rows[0]["cpf"] == "00006034578"


def test_query_clients_public_exposes_personal_and_functional_fields(tmp_path: Path) -> None:
    gov_csv = tmp_path / "gov_fields.csv"
    gov_csv.write_text(
        "CPF;NOME;SEXO;NASC;NOME DA MAE;EMAIL;ENDERECO;NRO;BAIRRO;CIDADE;UF;PIS;ENTIDADE;MATRICULA;REGIME DE CONTRATACAO;CARGO;DATA ADMISSAO;SITUACAO;MARGEM\n"
        "12345678900;MARIA TESTE;F;1980-05-19 00:00:00;JOANA TESTE;maria@exemplo.com;RUA A;45;CENTRO;SALVADOR;BA;12345678901;PREFEITURA;MAT-01;EFETIVO;ANALISTA;2020-01-10 00:00:00;ATIVO;300,00\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[gov_csv], base_segment="GOVERNO", base_source="BAHIA"),
    )

    rows = query_clients(session, FilterSet(base_segment="GOVERNO", cpf="12345678900"))
    assert len(rows) == 1
    assert rows[0]["sexo"] == "F"
    assert rows[0]["nome_mae"] == "JOANA TESTE"
    assert rows[0]["email"] == "maria@exemplo.com"
    assert rows[0]["endereco"] == "RUA A, 45, CENTRO"
    assert rows[0]["tipo_vinculo"] == "EFETIVO"
    assert rows[0]["cargo_funcao"] == "ANALISTA"
    assert rows[0]["data_admissao"] == "10/01/2020"


def test_import_files_prefeitura_salvador_titles_are_mapped(tmp_path: Path) -> None:
    csv_path = tmp_path / "prefeitura_salvador.csv"
    csv_path.write_text(
        "CPF;NOME_COMPLETO;ENDEREÇO;NUMERO;BAIRRO;CIDADE;UF;CEP;CATEGORIA;DT_ADMISSAO;AGEN_BCO_COD;AGEN_COD;CTA_PAG\n"
        "12345678900;MARIA TESTE;RUA A;45;CENTRO;SALVADOR;BA;40000000;EFETIVO;2020-01-10 00:00:00;001;1234;998877\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path], base_segment="PREFEITURA", base_source="SALVADOR"),
    )

    rows = query_clients(session, FilterSet(base_segment="PREFEITURA", cpf="12345678900"))
    assert len(rows) == 1
    assert rows[0]["nome"] == "MARIA TESTE"
    assert rows[0]["logradouro"] == "RUA A"
    assert rows[0]["numero"] == "45"
    assert rows[0]["bairro"] == "CENTRO"
    assert rows[0]["cidade"] == "SALVADOR"
    assert rows[0]["uf"] == "BA"
    assert rows[0]["cep"] == "40000000"
    assert rows[0]["tipo_vinculo"] == "EFETIVO"
    assert rows[0]["data_admissao"] == "10/01/2020"

    occurrence = session.scalars(select(ClientOccurrence).where(ClientOccurrence.cpf == "12345678900")).one()
    assert '"AGEN_BCO_COD": "001"' in occurrence.extras_json
    assert '"AGEN_COD": "1234"' in occurrence.extras_json
    assert '"CTA_PAG": "998877"' in occurrence.extras_json


def test_query_clients_public_fills_all_missing_fields_from_same_cpf_group(tmp_path: Path) -> None:
    cadastro_csv = tmp_path / "gov_cadastro_full.csv"
    cadastro_csv.write_text(
        "CPF;NOME;SEXO;NASC;NOME DA MAE;EMAIL;ENDERECO;NRO;COMPLEMENTO;BAIRRO;CIDADE;UF;CEP;TELEFONE_01\n"
        "12345678900;MARIA TESTE;F;1980-05-19 00:00:00;JOANA TESTE;maria@exemplo.com;RUA A;45;AP 2;CENTRO;SALVADOR;BA;40000000;71999990000\n",
        encoding="utf-8",
    )
    operacional_csv = tmp_path / "gov_operacional_full.csv"
    operacional_csv.write_text(
        "CPF;ENTIDADE;MATRICULA;REGIME DE CONTRATACAO;CARGO;DATA ADMISSAO;SITUACAO;CONVENIO;SECRETARIA;MARGEM\n"
        "12345678900;PREFEITURA;MAT-01;EFETIVO;ANALISTA;2020-01-10 00:00:00;ATIVO;BANCO X;SAUDE;300,00\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[cadastro_csv], base_segment="GOVERNO", base_source="BAHIA"),
    )
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[operacional_csv], base_segment="GOVERNO", base_source="BAHIA_ATUALIZADA"),
    )

    rows = query_clients(session, FilterSet(base_segment="GOVERNO", cpf="12345678900"))
    assert len(rows) == 1
    assert rows[0]["telefone"] == "71999990000"
    assert rows[0]["email"] == "maria@exemplo.com"
    assert rows[0]["nome_mae"] == "JOANA TESTE"
    assert rows[0]["endereco"] == "RUA A, 45, AP 2, CENTRO"
    assert rows[0]["cep"] == "40000000"
    assert rows[0]["matricula"] == "MAT-01"
    assert rows[0]["entidade"] == "PREFEITURA"
    assert rows[0]["tipo_vinculo"] == "EFETIVO"
    assert rows[0]["cargo_funcao"] == "ANALISTA"
    assert rows[0]["data_admissao"] == "10/01/2020"
    assert rows[0]["convenio"] == "BANCO X"
    assert rows[0]["secretaria"] == "SAUDE"


def test_query_clients_gov_sp_titles_are_mapped(tmp_path: Path) -> None:
    gov_csv = tmp_path / "gov_sp.csv"
    gov_csv.write_text(
        "CPF;NOME;SEXO;NASC;IDADE;CNPJ;RAZAO_SOCIAL;GRUPO;CARGO;ORGÃO;REMUNERAÇÃO DO MÊS;ENDERECO;BAIRRO;CIDADE;CEP;UF;TELEFONE_01;TELEFONE_02\n"
        "12345678900;MARIA TESTE;F;1980-05-19 00:00:00;40;12345678000199;ORG TESTE;EFETIVO;ANALISTA;SECRETARIA ESTADUAL;2500,00;RUA A;CENTRO;SAO PAULO;01000000;SP;11999990000;1133334444\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[gov_csv], base_segment="GOVERNO", base_source="SP"),
    )

    rows = query_clients(session, FilterSet(base_segment="GOVERNO", cpf="12345678900"))
    assert len(rows) == 1
    assert rows[0]["entidade"] == "SECRETARIA ESTADUAL"
    assert rows[0]["tipo_vinculo"] == "EFETIVO"
    assert rows[0]["cargo_funcao"] == "ANALISTA"
    assert rows[0]["salario"] == Decimal("2500.00")
    assert rows[0]["telefone"] == "11999990000"


def test_query_clients_nova_vida_fields_are_mapped(tmp_path: Path) -> None:
    nova_vida_csv = tmp_path / "nova_vida.csv"
    nova_vida_csv.write_text(
        "CPF;NOME;NOME_MAE;SEXO;NASC;RENDA;TIPO;TITULO;LOGRADOURO;NUMERO;COMPLEMENTO;BAIRRO;CIDADE;UF;CEP;AREARISCO;CEL1;FLGWHATSCEL1;PROCONCEL1;CEL2;FLGWHATSCEL2;EMAIL1;EMAIL2;EMAIL3\n"
        "12345678900;MARIA TESTE;JOANA TESTE;F;1980-05-19 00:00:00;3200,50;EFETIVO;ANALISTA ADMINISTRATIVO;RUA A;45;AP 2;CENTRO;SALVADOR;BA;40000000;NAO;71999990000;SIM;NAO;71988887777;NAO;maria1@exemplo.com;maria2@exemplo.com;maria3@exemplo.com\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[nova_vida_csv], base_segment="GOVERNO", base_source="NOVA VIDA"),
    )

    rows = query_clients(session, FilterSet(base_segment="GOVERNO", cpf="12345678900"))
    assert len(rows) == 1
    assert rows[0]["nome_mae"] == "JOANA TESTE"
    assert rows[0]["salario"] == Decimal("3200.50")
    assert rows[0]["tipo_vinculo"] == "EFETIVO"
    assert rows[0]["cargo_funcao"] == "ANALISTA ADMINISTRATIVO"
    assert rows[0]["telefone"] == "71999990000"
    assert rows[0]["telefone2"] == "71988887777"
    assert rows[0]["whatsapp_telefone1"] == "SIM"
    assert rows[0]["whatsapp_telefone2"] == "NAO"
    assert rows[0]["email"] == "maria1@exemplo.com"
    assert rows[0]["email2"] == "maria2@exemplo.com"
    assert rows[0]["email3"] == "maria3@exemplo.com"


def test_import_files_prefeitura_simple_service_sheet_keeps_financial_fields(tmp_path: Path) -> None:
    csv_path = tmp_path / "prefeitura_servico_simples.csv"
    csv_path.write_text(
        "cpf;matric;servico;name;situacao;entidade;cbo_titulo;margem;margem total\n"
        "01222536714;00003455;Prefeitura;CLAUDIA;Ativo - AT 0002 - SEC ESTATUTA;PREFEITURA;DUQUE CAXIAS;677,99;677,99\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path], base_segment="PREFEITURA", base_source="OPERACIONAL"),
    )

    rows = query_clients(session, FilterSet(base_segment="PREFEITURA", cpf="01222536714"))
    assert len(rows) == 1
    assert rows[0]["matricula"] == "00003455"
    assert rows[0]["servico"] == "PREFEITURA"
    assert rows[0]["situacao"] == "ATIVO - AT 0002 - SEC ESTATUTA"
    assert rows[0]["cbo_titulo"] == "DUQUE CAXIAS"
    assert rows[0]["vl_margem"] == Decimal("677.99")
    assert rows[0]["margem_total"] == Decimal("677.99")


def test_import_files_government_multiple_services_same_matricula_creates_multiple_rows(tmp_path: Path) -> None:
    csv_path = tmp_path / "gov_multiservico.csv"
    csv_path.write_text(
        "CPF;Servidor;Matrícula;Serviço;margem;Margem Total (R$)\n"
        "23820977520;LUCIANO SILVA RIOS;7171000000000000;BENEFÍCIO - CREDESTA - COMPRA;7,95;603\n"
        "23820977520;LUCIANO SILVA RIOS;7171000000000000;BENEFÍCIO - CREDCESTA SAQUE;496,45;603\n"
        "23820977520;LUCIANO SILVA RIOS;7171000000000000;EMPRÉSTIMO - 1;36,18;1326,59\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[csv_path], base_segment="GOVERNO", base_source="CEARA"),
    )

    rows = query_clients(session, FilterSet(base_segment="GOVERNO", cpf="23820977520"))
    assert len(rows) == 3
    assert {row["servico"] for row in rows} == {
        "BENEFÍCIO - CREDESTA - COMPRA",
        "BENEFÍCIO - CREDCESTA SAQUE",
        "EMPRÉSTIMO - 1",
    }
    assert {row["margem_total"] for row in rows} == {Decimal("603.00"), Decimal("1326.59")}
    assert {row["vl_margem"] for row in rows} == {Decimal("7.95"), Decimal("496.45"), Decimal("36.18")}

    client = session.get(Client, "23820977520")
    assert client is not None
    assert client.tem_governo is True
    assert client.qtd_ocorrencias == 3


def test_query_clients_public_consolidates_operational_row_with_cadastral_data(tmp_path: Path) -> None:
    cadastro_csv = tmp_path / "pref_cadastro.csv"
    cadastro_csv.write_text(
        "CPF;NOME HIGIENIZADO;NASC;TELEFONE_01;UF;CIDADE\n"
        "00641012748;MARIA HELENA RIBEIRO;01/02/1980;21999887766;RJ;DUQUE DE CAXIAS\n",
        encoding="utf-8",
    )
    operacional_csv = tmp_path / "pref_operacional.csv"
    operacional_csv.write_text(
        "CPF;MATRICULA;SERVICO;SITUACAO;ENTIDADE;CBO_TITULO;MARGEM;MARGEM TOTAL\n"
        "00641012748;00000223781;PREFEITURA DUQUE CAXIAS;ATIVO - ATIVO;0035 - SEC MUN SAUDE SMS;ESTATUTARIO CTRIENIO AUTOMATICO;229,93;229,93\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[cadastro_csv], base_segment="PREFEITURA", base_source="DUQUE DE CAXIAS"),
    )
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[operacional_csv], base_segment="PREFEITURA", base_source="DUQUE DE CAXIAS - ATUALIZACAO"),
    )

    rows = query_clients(session, FilterSet(base_segment="PREFEITURA", cpf="00641012748"))
    assert len(rows) == 1
    assert rows[0]["matricula"] == "00000223781"
    assert rows[0]["servico"] == "PREFEITURA DUQUE CAXIAS"
    assert rows[0]["telefone"] == "21999887766"
    assert rows[0]["dt_nasc"] == "01/02/1980"
    assert rows[0]["idade"] == compute_age_from_birth("01/02/1980")


def test_public_update_import_materializes_missing_prefeitura_matriculas(tmp_path: Path) -> None:
    cadastro_csv = tmp_path / "pref_cadastro.csv"
    cadastro_csv.write_text(
        "CPF;NOME HIGIENIZADO;MATRICULA;TELEFONE_01;UF;CIDADE\n"
        "06838941708;ESTELA MARES DA SILVA;00000169284;21988111995;RJ;DUQUE DE CAXIAS\n",
        encoding="utf-8",
    )
    update_csv = tmp_path / "pref_update.csv"
    update_csv.write_text(
        "CPF;MATRICULA;SERVICO_SERVIDOR;SITUACAO;MARGEM_DISPONIVEL;MARGEM_TOTAL\n"
        "06838941708;00000169284;PREFEITURA DUQUE CAXIAS;ATIVO - ATIVO;602,61;602,61\n"
        "06838941708;00000462187;PREFEITURA DUQUE CAXIAS;ATIVO - ATIVO;263,39;263,39\n",
        encoding="utf-8",
    )

    session = make_session()
    import_files(
        session,
        ImportRequest(year=2025, month=8, user_name="teste", origin_folder=str(tmp_path), files=[cadastro_csv], base_segment="PREFEITURA", base_source="DUQUE DE CAXIAS"),
    )

    summary = import_public_campaign_updates(
        session,
        [update_csv],
        base_segment="PREFEITURA",
        source_name="DUQUE DE CAXIAS",
        imported_by="teste",
    )

    assert summary.clients_updated == 1
    rows = query_clients(session, FilterSet(base_segment="PREFEITURA", cpf="06838941708"))
    assert len(rows) == 2
    assert {row["matricula"] for row in rows} == {"00000169284", "00000462187"}
    by_matricula = {row["matricula"]: row for row in rows}
    assert by_matricula["00000169284"]["vl_margem"] == Decimal("602.61")
    assert by_matricula["00000462187"]["vl_margem"] == Decimal("263.39")
