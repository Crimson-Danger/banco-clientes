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
    nome: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    nu_nb: Mapped[str] = mapped_column(String(30), nullable=False, default="")
    esp: Mapped[str] = mapped_column(String(10), nullable=False, default="")
    ddb: Mapped[str] = mapped_column(String(40), nullable=False, default="")
    vl_rmi: Mapped[Decimal | None] = mapped_column(Numeric(14, 2))
    vl_margem: Mapped[Decimal | None] = mapped_column(Numeric(14, 2), index=True)
    vl_rmc: Mapped[Decimal | None] = mapped_column(Numeric(14, 2))
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
