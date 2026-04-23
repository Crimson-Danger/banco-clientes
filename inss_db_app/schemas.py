from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class ImportSummary:
    batch_id: int
    files_imported: int = 0
    files_skipped: int = 0
    rows_imported: int = 0
    duplicate_files: list[str] = field(default_factory=list)
    invalid_files: list[str] = field(default_factory=list)
    invalid_rows: int = 0
    duplicate_rows: int = 0
    species_filtered_rows: int = 0
    cpfs_new: int = 0
    cpfs_updated: int = 0
    rows_without_phone: int = 0
    rows_without_city: int = 0
    validation_errors: list[str] = field(default_factory=list)


@dataclass
class ImportRequest:
    year: int
    month: int
    user_name: str
    origin_folder: str
    files: list[Path]
    base_segment: str = "INSS"
    base_source: str = ""
    allowed_species: list[str] = field(default_factory=list)
