from __future__ import annotations

from types import SimpleNamespace

from sqlalchemy import or_, select

from inss_db_app.config import settings
from inss_db_app.database import SessionLocal, serialized_write
from inss_db_app.importer import (
    apply_enrichment_to_occurrence,
    enrich_normalized_row_with_cep,
    enrich_normalized_row_with_ddd,
    rebuild_clients_for_cpfs,
)
from inss_db_app.models import ClientOccurrence


def occurrence_to_normalized(occurrence: ClientOccurrence) -> dict[str, object]:
    return {
        "nome": occurrence.nome or "",
        "nome_higienizado": occurrence.nome_higienizado or "",
        "cpf_original": occurrence.cpf_original or "",
        "matricula": occurrence.matricula or "",
        "servico": occurrence.servico or "",
        "situacao": occurrence.situacao or "",
        "nu_nb": occurrence.nu_nb or "",
        "pis": occurrence.pis or "",
        "entidade": occurrence.entidade or "",
        "secretaria": occurrence.secretaria or "",
        "convenio": occurrence.convenio or "",
        "cbo_titulo": occurrence.cbo_titulo or "",
        "esp": occurrence.esp or "",
        "ddb": occurrence.ddb or "",
        "uf": occurrence.uf or "",
        "cidade": occurrence.cidade or "",
        "bairro": occurrence.bairro or "",
        "logradouro": occurrence.logradouro or "",
        "numero": occurrence.numero or "",
        "complemento": occurrence.complemento or "",
        "cep": occurrence.cep or "",
        "dt_nasc": occurrence.dt_nasc or "",
        "nome_mae": occurrence.nome_mae or "",
        "sexo": occurrence.sexo or "",
        "email": occurrence.email or "",
        "salario": occurrence.salario,
        "vl_rmi": occurrence.vl_rmi,
        "vl_margem": occurrence.vl_margem,
        "vl_rmc": occurrence.vl_rmc,
        "margem_total": occurrence.margem_total,
        "margem_liquida": occurrence.margem_liquida,
        "margem_bruta": occurrence.margem_bruta,
        "margem_utilizada": occurrence.margem_utilizada,
        "extras_json": occurrence.extras_json or "",
    }


def main() -> None:
    session = SessionLocal()
    patched_settings = dict(vars(settings))
    patched_settings["enable_cep_enrichment"] = True
    patched_settings["enable_ddd_enrichment"] = True

    import inss_db_app.importer as importer_module

    importer_module.settings = SimpleNamespace(**patched_settings)

    cep_cache: dict[str, object | None] = {}
    ddd_cache: dict[str, object | None] = {}
    affected_cpfs: set[str] = set()
    rows_updated = 0

    try:
        with serialized_write():
            candidates = session.scalars(
                select(ClientOccurrence).where(
                    or_(
                        ClientOccurrence.cep != "",
                        ClientOccurrence.telefone1 != "",
                        ClientOccurrence.telefone2 != "",
                        ClientOccurrence.telefone3 != "",
                    )
                )
            ).all()

            for occurrence in candidates:
                normalized = occurrence_to_normalized(occurrence)
                normalized = enrich_normalized_row_with_cep(normalized, cep_cache)
                normalized = enrich_normalized_row_with_ddd(normalized, ddd_cache)
                phones = [occurrence.telefone1 or "", occurrence.telefone2 or "", occurrence.telefone3 or ""]
                _updated_phone, updated_data = apply_enrichment_to_occurrence(occurrence, normalized, phones)
                if updated_data:
                    rows_updated += 1
                    affected_cpfs.add(occurrence.cpf)

            if affected_cpfs:
                rebuild_clients_for_cpfs(session, affected_cpfs)
            session.commit()
    finally:
        session.close()

    print(f"rows_updated={rows_updated}")
    print(f"cpfs_rebuilt={len(affected_cpfs)}")


if __name__ == "__main__":
    main()
