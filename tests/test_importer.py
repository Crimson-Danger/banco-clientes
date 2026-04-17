from __future__ import annotations

from datetime import date
from decimal import Decimal
from pathlib import Path

from sqlalchemy import create_engine
from sqlalchemy.orm import Session, sessionmaker

from inss_db_app.database import Base
from inss_db_app.importer import (
    FilterSet,
    canonicalize_row_keys,
    compute_age_from_birth,
    delete_batch,
    detect_week_label,
    export_clients,
    import_files,
    normalize_date_text,
    normalize_row,
    query_clients,
    week_label_from_ddb,
)
from inss_db_app.models import ClientOccurrence
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
    csv_path = tmp_path / "age_range.csv"
    csv_path.write_text(
        "CPF;NOME;NU-NB;ESP;DDB;DT-NASC;VL-RMI;VL MARGEM;VL RMC;UF;CIDADE;TELEFONE1\n"
        "12345678900;MARIA SILVA;111222333;21;05/01/2026;16/04/1966;R$ 100,00;R$ 50,00;R$ 10,00;SP;SAO PAULO;11999999999\n"
        "99988877766;ANA LIMA;444555666;87;02/02/2026;17/04/1966;R$ 200,00;R$ 80,00;R$ 15,00;RJ;RIO DE JANEIRO;21988887777\n"
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
