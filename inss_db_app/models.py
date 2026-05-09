from __future__ import annotations

from datetime import UTC, datetime
from decimal import Decimal

from sqlalchemy import Boolean, DateTime, ForeignKey, Integer, Numeric, String, Text, UniqueConstraint
from sqlalchemy.orm import Mapped, mapped_column, relationship

from .database import Base


class ImportBatch(Base):
    __tablename__ = "import_batches"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    mes_referencia: Mapped[int] = mapped_column(Integer, nullable=False, index=True)
    ano_referencia: Mapped[int] = mapped_column(Integer, nullable=False, index=True)
    origem_pasta: Mapped[str] = mapped_column(String(500), nullable=False)
    base_segment: Mapped[str] = mapped_column(String(12), nullable=False, default="INSS", index=True)
    base_source: Mapped[str] = mapped_column(String(120), nullable=False, default="")
    usuario: Mapped[str] = mapped_column(String(120), nullable=False, default="operador")
    status: Mapped[str] = mapped_column(String(40), nullable=False, default="PENDENTE")
    data_importacao: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(UTC), nullable=False)
    resumo: Mapped[str] = mapped_column(Text, nullable=False, default="")

    source_files: Mapped[list["SourceFile"]] = relationship(back_populates="batch", cascade="all, delete-orphan")


class SourceFile(Base):
    __tablename__ = "source_files"
    __table_args__ = (UniqueConstraint("hash_arquivo", name="uq_source_files_hash"),)

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    batch_id: Mapped[int] = mapped_column(ForeignKey("import_batches.id"), nullable=False, index=True)
    nome_arquivo: Mapped[str] = mapped_column(String(255), nullable=False, index=True)
    caminho_arquivo: Mapped[str] = mapped_column(String(500), nullable=False)
    base_segment: Mapped[str] = mapped_column(String(12), nullable=False, default="INSS", index=True)
    base_source: Mapped[str] = mapped_column(String(120), nullable=False, default="")
    semana_label: Mapped[str] = mapped_column(String(120), nullable=False, default="")
    tipo_arquivo: Mapped[str] = mapped_column(String(20), nullable=False)
    hash_arquivo: Mapped[str] = mapped_column(String(64), nullable=False)
    total_linhas: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    duplicado: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    criado_em: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(UTC), nullable=False)

    batch: Mapped["ImportBatch"] = relationship(back_populates="source_files")
    occurrences: Mapped[list["ClientOccurrence"]] = relationship(back_populates="source_file", cascade="all, delete-orphan")


class ClientOccurrence(Base):
    __tablename__ = "client_occurrences"
    __table_args__ = (UniqueConstraint("source_file_id", "row_hash", name="uq_occurrence_row_hash"),)

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    source_file_id: Mapped[int] = mapped_column(ForeignKey("source_files.id"), nullable=False, index=True)
    cpf: Mapped[str] = mapped_column(String(11), nullable=False, index=True)
    cpf_original: Mapped[str] = mapped_column(String(20), nullable=False, default="", index=True)
    base_segment: Mapped[str] = mapped_column(String(12), nullable=False, default="INSS", index=True)
    base_source: Mapped[str] = mapped_column(String(120), nullable=False, default="", index=True)
    nome: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    nome_higienizado: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    matricula: Mapped[str] = mapped_column(String(80), nullable=False, default="", index=True)
    servico: Mapped[str] = mapped_column(String(255), nullable=False, default="", index=True)
    situacao: Mapped[str] = mapped_column(String(120), nullable=False, default="")
    nu_nb: Mapped[str] = mapped_column(String(30), nullable=False, default="")
    pis: Mapped[str] = mapped_column(String(20), nullable=False, default="", index=True)
    entidade: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    secretaria: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    convenio: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    cbo_titulo: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    salario: Mapped[Decimal | None] = mapped_column(Numeric(14, 2))
    esp: Mapped[str] = mapped_column(String(10), nullable=False, default="")
    ddb: Mapped[str] = mapped_column(String(40), nullable=False, default="")
    vl_rmi: Mapped[Decimal | None] = mapped_column(Numeric(14, 2))
    vl_margem: Mapped[Decimal | None] = mapped_column(Numeric(14, 2), index=True)
    vl_rmc: Mapped[Decimal | None] = mapped_column(Numeric(14, 2))
    margem_total: Mapped[Decimal | None] = mapped_column(Numeric(14, 2))
    margem_liquida: Mapped[Decimal | None] = mapped_column(Numeric(14, 2))
    margem_bruta: Mapped[Decimal | None] = mapped_column(Numeric(14, 2))
    margem_utilizada: Mapped[Decimal | None] = mapped_column(Numeric(14, 2))
    uf: Mapped[str] = mapped_column(String(2), nullable=False, default="", index=True)
    cidade: Mapped[str] = mapped_column(String(120), nullable=False, default="", index=True)
    bairro: Mapped[str] = mapped_column(String(120), nullable=False, default="")
    logradouro: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    numero: Mapped[str] = mapped_column(String(40), nullable=False, default="")
    complemento: Mapped[str] = mapped_column(String(120), nullable=False, default="")
    cep: Mapped[str] = mapped_column(String(20), nullable=False, default="")
    telefone1: Mapped[str] = mapped_column(String(20), nullable=False, default="")
    telefone2: Mapped[str] = mapped_column(String(20), nullable=False, default="")
    telefone3: Mapped[str] = mapped_column(String(20), nullable=False, default="")
    dt_nasc: Mapped[str] = mapped_column(String(40), nullable=False, default="")
    nome_mae: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    sexo: Mapped[str] = mapped_column(String(20), nullable=False, default="")
    email: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    extras_json: Mapped[str] = mapped_column(Text, nullable=False, default="")
    row_hash: Mapped[str] = mapped_column(String(64), nullable=False)
    criado_em: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(UTC), nullable=False)

    source_file: Mapped["SourceFile"] = relationship(back_populates="occurrences")


class Client(Base):
    __tablename__ = "clients"

    cpf: Mapped[str] = mapped_column(String(11), primary_key=True)
    nome_atual: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    uf_atual: Mapped[str] = mapped_column(String(2), nullable=False, default="", index=True)
    cidade_atual: Mapped[str] = mapped_column(String(120), nullable=False, default="", index=True)
    melhor_telefone: Mapped[str] = mapped_column(String(20), nullable=False, default="")
    tem_telefone: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False, index=True)
    do_not_call: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False, index=True)
    tem_inss: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False, index=True)
    tem_governo: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False, index=True)
    pis_atual: Mapped[str] = mapped_column(String(20), nullable=False, default="")
    matricula_atual: Mapped[str] = mapped_column(String(80), nullable=False, default="")
    entidade_atual: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    maior_margem: Mapped[Decimal | None] = mapped_column(Numeric(14, 2), index=True)
    ultima_referencia: Mapped[str] = mapped_column(String(20), nullable=False, default="", index=True)
    qtd_ocorrencias: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    atualizado_em: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(UTC), nullable=False)


class ExportHistory(Base):
    __tablename__ = "exports"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    filtros_json: Mapped[str] = mapped_column(Text, nullable=False)
    arquivo_saida: Mapped[str] = mapped_column(String(500), nullable=False)
    total_registros: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    criado_em: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(UTC), nullable=False)


class SavedFilter(Base):
    __tablename__ = "saved_filters"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("app_users.id"), nullable=False, index=True)
    page_key: Mapped[str] = mapped_column(String(40), nullable=False, default="clients", index=True)
    name: Mapped[str] = mapped_column(String(120), nullable=False)
    query_string: Mapped[str] = mapped_column(Text, nullable=False, default="")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(UTC), nullable=False)


class AuditLog(Base):
    __tablename__ = "audit_logs"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    actor_username: Mapped[str] = mapped_column(String(120), nullable=False, default="sistema", index=True)
    action: Mapped[str] = mapped_column(String(80), nullable=False, index=True)
    target_type: Mapped[str] = mapped_column(String(60), nullable=False, default="")
    target_id: Mapped[str] = mapped_column(String(120), nullable=False, default="")
    message: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    metadata_json: Mapped[str] = mapped_column(Text, nullable=False, default="")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(UTC), nullable=False, index=True)


class ExportJob(Base):
    __tablename__ = "export_jobs"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    requested_by: Mapped[str] = mapped_column(String(120), nullable=False, default="operador", index=True)
    job_type: Mapped[str] = mapped_column(String(40), nullable=False, default="campaign")
    status: Mapped[str] = mapped_column(String(40), nullable=False, default="PENDENTE", index=True)
    filters_json: Mapped[str] = mapped_column(Text, nullable=False, default="")
    include_audit: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    file_format: Mapped[str] = mapped_column(String(10), nullable=False, default="xlsx")
    output_file: Mapped[str] = mapped_column(String(500), nullable=False, default="")
    total_records: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    error_message: Mapped[str] = mapped_column(Text, nullable=False, default="")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(UTC), nullable=False, index=True)
    started_at: Mapped[datetime | None] = mapped_column(DateTime)
    finished_at: Mapped[datetime | None] = mapped_column(DateTime)


class ImportJob(Base):
    __tablename__ = "import_jobs"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    batch_id: Mapped[int] = mapped_column(ForeignKey("import_batches.id"), nullable=False, index=True)
    requested_by: Mapped[str] = mapped_column(String(120), nullable=False, default="operador", index=True)
    status: Mapped[str] = mapped_column(String(40), nullable=False, default="PENDENTE", index=True)
    request_json: Mapped[str] = mapped_column(Text, nullable=False, default="{}")
    error_message: Mapped[str] = mapped_column(Text, nullable=False, default="")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(UTC), nullable=False, index=True)
    started_at: Mapped[datetime | None] = mapped_column(DateTime)
    finished_at: Mapped[datetime | None] = mapped_column(DateTime)


class PhoneEnrichment(Base):
    __tablename__ = "phone_enrichments"
    __table_args__ = (UniqueConstraint("cpf", "telefone", "fonte", name="uq_phone_enrichment_unique"),)

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    cpf: Mapped[str] = mapped_column(String(11), nullable=False, index=True)
    telefone: Mapped[str] = mapped_column(String(20), nullable=False, default="")
    fonte: Mapped[str] = mapped_column(String(40), nullable=False, default="")
    arquivo_origem: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    importado_por: Mapped[str] = mapped_column(String(120), nullable=False, default="operador")
    criado_em: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(UTC), nullable=False)


class EnrichmentImportItem(Base):
    __tablename__ = "enrichment_import_items"
    __table_args__ = (UniqueConstraint("cpf", "arquivo_origem", "fonte", name="uq_enrichment_import_item"),)

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    cpf: Mapped[str] = mapped_column(String(11), nullable=False, index=True)
    fonte: Mapped[str] = mapped_column(String(40), nullable=False, default="")
    arquivo_origem: Mapped[str] = mapped_column(String(255), nullable=False, default="", index=True)
    importado_por: Mapped[str] = mapped_column(String(120), nullable=False, default="operador")
    criado_em: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(UTC), nullable=False)


class PublicCampaignUpdate(Base):
    __tablename__ = "public_campaign_updates"
    __table_args__ = (UniqueConstraint("cpf", "matricula", "ade", "fonte", "arquivo_origem", name="uq_public_campaign_update"),)

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    base_segment: Mapped[str] = mapped_column(String(12), nullable=False, default="GOVERNO", index=True)
    cpf: Mapped[str] = mapped_column(String(11), nullable=False, index=True)
    matricula: Mapped[str] = mapped_column(String(80), nullable=False, default="", index=True)
    tipo_servico_servidor: Mapped[str] = mapped_column(String(120), nullable=False, default="")
    servico_servidor: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    margem_disponivel: Mapped[Decimal | None] = mapped_column(Numeric(14, 2))
    margem_total: Mapped[Decimal | None] = mapped_column(Numeric(14, 2))
    consignataria: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    situacao: Mapped[str] = mapped_column(String(120), nullable=False, default="")
    ade: Mapped[str] = mapped_column(String(80), nullable=False, default="")
    servico_consignataria: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    prestacoes: Mapped[int | None] = mapped_column(Integer)
    pagas: Mapped[int | None] = mapped_column(Integer)
    valor: Mapped[Decimal | None] = mapped_column(Numeric(14, 2))
    deferimento: Mapped[str] = mapped_column(String(40), nullable=False, default="")
    quitacao: Mapped[str] = mapped_column(String(40), nullable=False, default="")
    ultimo_desconto: Mapped[str] = mapped_column(String(40), nullable=False, default="")
    ultima_parcela: Mapped[str] = mapped_column(String(40), nullable=False, default="")
    fonte: Mapped[str] = mapped_column(String(120), nullable=False, default="")
    arquivo_origem: Mapped[str] = mapped_column(String(255), nullable=False, default="", index=True)
    importado_por: Mapped[str] = mapped_column(String(120), nullable=False, default="operador")
    criado_em: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(UTC), nullable=False)


class AppUser(Base):
    __tablename__ = "app_users"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    username: Mapped[str] = mapped_column(String(80), nullable=False, unique=True, index=True)
    full_name: Mapped[str] = mapped_column(String(160), nullable=False, default="")
    password_hash: Mapped[str] = mapped_column(String(128), nullable=False)
    password_salt: Mapped[str] = mapped_column(String(64), nullable=False)
    is_active: Mapped[bool] = mapped_column(Boolean, nullable=False, default=True)
    is_admin: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    can_view_dashboard: Mapped[bool] = mapped_column(Boolean, nullable=False, default=True)
    can_import: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    can_search: Mapped[bool] = mapped_column(Boolean, nullable=False, default=True)
    can_export: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    can_view_history: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    can_delete_batches: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    can_manage_users: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    last_login_at: Mapped[datetime | None] = mapped_column(DateTime)
    last_password_reset_at: Mapped[datetime | None] = mapped_column(DateTime)
    locked_until: Mapped[datetime | None] = mapped_column(DateTime)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(UTC), nullable=False)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(UTC), nullable=False)
