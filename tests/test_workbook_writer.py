"""Testes do escritor usando somente workbooks OpenPyXL em memória."""

from __future__ import annotations

from datetime import date
from pathlib import Path
from unittest.mock import Mock

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

import src.excel.workbook_writer as writer
from src.core.models import Policy, SheetLayout, TravelResult, WorkbookLayout
from src.excel.workbook_writer import (
    _absolute_range,
    _cell,
    _find_pivot_title_row,
    _find_priority_table,
    _prepare_pivot_area,
    _quoted_sheet,
    apply_conditional_formatting,
    create_fallback_summary_and_chart,
    create_support_sheet,
    ensure_derived_columns,
    finalize_workbook,
    load_source_workbook,
    write_base_formulas,
    write_responses,
)
from src.settings import load_aliases, load_rules

BASE_COLUMNS = {
    "request_id": 1,
    "request_date": 2,
    "travel_date": 3,
    "cost_center": 4,
    "service_type": 5,
    "supplier": 6,
    "quantity": 7,
    "unit_value": 8,
    "fees": 9,
    "booking_status": 10,
    "criticality": 11,
    "card_status": 12,
}
POLICY_COLUMNS = {
    "service_type": 1,
    "limit_value": 2,
    "min_lead_days": 3,
}


def make_travel(
    request_id: str,
    *,
    row: int,
    cost_center: str,
    service_type: str,
    supplier: str,
    total: float,
    score: float,
    priority: str,
) -> TravelResult:
    return TravelResult(
        worksheet_row=row,
        request_id=request_id,
        request_date=date(2026, 1, 1),
        travel_date=date(2026, 1, 10),
        service_type=service_type,
        supplier=supplier,
        cost_center=cost_center,
        booking_status="Confirmada",
        criticality="Normal",
        card_status="Conferido",
        total_value=total,
        lead_days=9,
        policy_limit=1000.0,
        min_lead_days=5,
        policy_status="OK",
        limit_difference=total - 1000,
        score=score,
        priority=priority,
        reasons=["custo elevado", "prazo curto"],
        recommended_actions=["renegociar tarifa", "obter aprovação"],
    )


def make_writer_fixture() -> tuple[
    Workbook,
    WorkbookLayout,
    list[TravelResult],
    dict[str, Policy],
    dict[str, object],
    dict[str, object],
]:
    workbook = Workbook()
    base = workbook.active
    base.title = "Base Viagens"
    for field, column in BASE_COLUMNS.items():
        base.cell(1, column, field)
    base.append(
        [
            "VIA-001",
            date(2026, 1, 1),
            date(2026, 1, 10),
            "CC 100",
            "Aéreo",
            "Fornecedor Á",
            1,
            1200,
            50,
            "Confirmada",
            "Normal",
            "Conferido",
        ]
    )
    base.append(
        [
            "VIA-002",
            date(2026, 1, 2),
            date(2026, 1, 12),
            "CC 200",
            "Hotel",
            "Fornecedor B",
            2,
            400,
            20,
            "Pendente",
            "Executivo",
            "Pendente",
        ]
    )
    for row in range(2, 4):
        base.cell(row, 12).fill = PatternFill("solid", fgColor="FFABCDEF")

    policies_sheet = workbook.create_sheet("Políticas O'Brien")
    for field, column in POLICY_COLUMNS.items():
        policies_sheet.cell(1, column, field)
    policies_sheet.append(["Aéreo", 1000, 7])
    policies_sheet.append(["Hotel", 900, 5])

    responses = workbook.create_sheet("Respostas")
    aliases = load_aliases()
    for row, options in enumerate(aliases["indicator_labels"].values(), start=2):
        responses.cell(row, 1, options[0])
    priority_row = 15
    for column, header in enumerate(
        [
            "ID Solicitação",
            "Motivo da prioridade",
            "Ação recomendada",
            "Ordem (1 a 5)",
        ],
        start=1,
    ):
        responses.cell(priority_row, column, header)
    responses.cell(25, 1, "Tabela dinâmica e gráfico")

    layout = WorkbookLayout(
        base=SheetLayout("Base Viagens", 1, BASE_COLUMNS),
        policies=SheetLayout("Políticas O'Brien", 1, POLICY_COLUMNS),
        responses=SheetLayout(
            "Respostas",
            1,
            {"indicator": 1, "answer": 2, "formula_used": 3},
        ),
    )
    travels = [
        make_travel(
            "VIA-001",
            row=2,
            cost_center="CC 100",
            service_type="Aéreo",
            supplier="Fornecedor Á",
            total=1250,
            score=90,
            priority="Crítica",
        ),
        make_travel(
            "VIA-002",
            row=3,
            cost_center="CC 200",
            service_type="Hotel",
            supplier="Fornecedor B",
            total=820,
            score=45,
            priority="Alta",
        ),
    ]
    policies = {
        "aereo": Policy("Aéreo", 1000, 7, 2),
        "hotel": Policy("Hotel", 900, 5, 3),
    }
    return workbook, layout, travels, policies, aliases, load_rules()


@pytest.mark.parametrize(
    ("filename", "keep_vba"),
    [("entrada.xlsx", False), ("ENTRADA.XLSM", True)],
)
def test_load_source_workbook_preserves_macros_only_when_needed(
    filename: str,
    keep_vba: bool,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    expected = object()
    mocked_load = Mock(return_value=expected)
    monkeypatch.setattr(writer, "load_workbook", mocked_load)

    result = load_source_workbook(Path(filename))

    assert result is expected
    mocked_load.assert_called_once_with(
        Path(filename),
        data_only=False,
        keep_vba=keep_vba,
    )


def test_complete_writer_pipeline_serializes_auditable_workbook(
    tmp_path: Path,
) -> None:
    workbook, layout, travels, policies, aliases, rules = make_writer_fixture()

    updated_layout = ensure_derived_columns(workbook, layout, last_row=3)
    repeated_layout = ensure_derived_columns(workbook, updated_layout, last_row=3)
    support_refs = create_support_sheet(
        workbook,
        repeated_layout,
        travels,
        rules,
        last_row=3,
    )
    write_base_formulas(
        workbook,
        repeated_layout,
        policies,
        support_refs,
        last_row=3,
    )
    pivot_start_row = write_responses(
        workbook,
        repeated_layout,
        aliases,
        travels,
        support_refs,
        last_row=3,
        top_quantity=5,
    )
    apply_conditional_formatting(workbook, repeated_layout, last_row=3)
    create_fallback_summary_and_chart(
        workbook,
        repeated_layout,
        travels,
        pivot_start_row,
        last_row=3,
    )
    finalize_workbook(workbook)

    output_file = tmp_path / "resultado.xlsx"
    workbook.save(output_file)
    workbook.close()
    reopened = load_workbook(output_file, data_only=False)
    try:
        base = reopened["Base Viagens"]
        responses = reopened["Respostas"]
        support = reopened["Apoio_Automacao"]
        columns = repeated_layout.base.columns

        assert repeated_layout == updated_layout
        assert layout.base.columns == BASE_COLUMNS
        assert len(base.tables) == 1
        assert next(iter(base.tables.values())).ref == "A1:S3"
        assert columns["total_value"] == 13
        assert columns["priority_score"] == 19
        assert base.column_dimensions["S"].hidden is True
        assert base["M2"].value.startswith("=IFERROR(")
        assert "INDEX(" in base["O2"].value
        assert base["S2"].value.startswith("=ROUND(")
        assert base["R1"].comment is not None
        assert len(base.conditional_formatting._cf_rules) == 2
        assert base.freeze_panes == "A2"

        assert support.sheet_state == "hidden"
        assert support["K3"].value == "Centro de Custo"
        assert support["M4"].value == 1250
        assert support_refs["critical_threshold"].startswith("'Apoio_Automacao'!$B$")

        assert responses["B2"].value.startswith("=SUM(")
        assert responses["C2"].value == "SOMA"
        assert responses["A16"].value == "VIA-001"
        assert responses["D16"].value == 1
        assert responses["D16"].comment is not None
        assert responses.cell(pivot_start_row, 1).value == "Centro de Custo"
        assert responses.cell(pivot_start_row, 2).value == "Aéreo"
        assert len(responses._charts) == 1
        assert responses._charts[0].title is not None

        assert reopened.calculation.calcMode == "auto"
        assert reopened.calculation.fullCalcOnLoad is True
        assert reopened.calculation.forceFullCalc is True
    finally:
        reopened.close()


def test_support_sheet_is_replaced_and_sorted_deterministically() -> None:
    workbook, layout, travels, _, _, rules = make_writer_fixture()
    layout = ensure_derived_columns(workbook, layout, last_row=3)
    old_support = workbook.create_sheet("Apoio_Automacao")
    old_support["A1"] = "conteúdo antigo"
    travels.append(
        make_travel(
            "VIA-003",
            row=4,
            cost_center="cc 100",
            service_type="Hotel",
            supplier="fornecedor á",
            total=100,
            score=10,
            priority="Normal",
        )
    )

    references = create_support_sheet(workbook, layout, travels, rules, last_row=3)

    support = workbook["Apoio_Automacao"]
    assert support["A1"].value != "conteúdo antigo"
    cost_centers = [support["E4"].value, support["E5"].value, support["E6"].value]
    assert set(cost_centers) == {"CC 100", "cc 100", "CC 200"}
    assert cost_centers[-1] == "CC 200"
    assert support["F4"].value.startswith("=SUMIF(")
    assert references["cost_center_labels"].endswith("$E$6")
    assert references["supplier_labels"].endswith("$H$6")


def test_writer_search_and_reference_helpers_cover_fallbacks() -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "O'Brien"
    worksheet["A1"] = "conteúdo"
    worksheet.merge_cells("A5:C5")

    priority_row, columns = _find_priority_table(worksheet)
    pivot_title_row = _find_pivot_title_row(worksheet)
    _prepare_pivot_area(
        worksheet,
        title_row=3,
        pivot_start_row=5,
    )

    assert priority_row == 7
    assert columns == {"request_id": 1, "reason": 2, "action": 3, "order": 4}
    assert pivot_title_row == 7
    assert "A5:C5" not in {str(item) for item in worksheet.merged_cells.ranges}
    assert _quoted_sheet("O'Brien") == "'O''Brien'"
    assert _absolute_range("O'Brien", 28, 2, 10) == "'O''Brien'!$AB$2:$AB$10"
    assert _cell(28, 10) == "AB10"
