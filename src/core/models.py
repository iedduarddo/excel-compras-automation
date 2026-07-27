"""Modelos simples compartilhados entre os módulos."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class SheetLayout:
    """Localização de uma tabela lógica dentro de uma aba."""

    title: str
    header_row: int
    columns: dict[str, int]


@dataclass(frozen=True)
class WorkbookLayout:
    """Abas e colunas identificadas no arquivo."""

    base: SheetLayout
    policies: SheetLayout
    responses: SheetLayout


@dataclass(frozen=True)
class Policy:
    """Regra de custo e antecedência para um tipo de serviço."""

    service_type: str
    limit_value: float
    min_lead_days: int
    worksheet_row: int


@dataclass
class TravelResult:
    """Resultado calculado para uma solicitação."""

    worksheet_row: int
    request_id: str
    request_date: date | datetime
    travel_date: date | datetime
    service_type: str
    supplier: str
    cost_center: str
    booking_status: str
    criticality: str
    card_status: str
    total_value: float
    lead_days: int
    policy_limit: float
    min_lead_days: int
    policy_status: str
    limit_difference: float
    score: float = 0.0
    priority: str = "Normal"
    reasons: list[str] = field(default_factory=list)
    recommended_actions: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class RunPaths:
    """Caminhos produzidos para uma execução."""

    input_file: Path
    backup_file: Path
    output_file: Path
    log_file: Path


@dataclass(frozen=True)
class RunResult:
    """Resumo devolvido ao terminal no final."""

    output_file: Path
    backup_file: Path
    log_file: Path
    native_pivot_created: bool
    detected_sheets: dict[str, str]
    detected_columns: dict[str, dict[str, int]]
    checks: dict[str, Any]
