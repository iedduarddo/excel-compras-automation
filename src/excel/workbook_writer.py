"""Escrita de fórmulas, indicadores, análises e formatação."""

from __future__ import annotations

from copy import copy
from dataclasses import replace
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet

from src.business.priorities import rank_immediate_requests
from src.core.exceptions import DetectionError
from src.core.models import Policy, SheetLayout, TravelResult, WorkbookLayout
from src.excel.detection import find_indicator_rows
from src.services.text import normalize_text, text_similarity


DERIVED_HEADERS = {
    "total_value": "Valor Total",
    "lead_days": "Dias Antecedência",
    "policy_limit": "Limite Política",
    "policy_status": "Status Política",
    "limit_difference": "Diferença p/ Limite",
    "priority": "Prioridade",
    "priority_score": "Score Prioridade",
}

SUPPORT_SHEET = "Apoio_Automacao"

BLUE = "FF2F75B5"
DARK_BLUE = "FF17365D"
ORANGE = "FFC55A11"
LIGHT_YELLOW = "FFFFF2CC"
LIGHT_BLUE = "FFD9EAF7"
LIGHT_RED = "FFF4CCCC"
LIGHT_ORANGE = "FFFCE4D6"
DARK_RED = "FF9C0006"
WHITE = "FFFFFFFF"
GRAY = "FFD9E1F2"
THIN_GRAY = Side(style="thin", color="FFD9E1F2")


def load_source_workbook(path: Path) -> Workbook:
    """Abre .xlsx ou .xlsm preservando macros quando necessário."""

    return load_workbook(
        path,
        data_only=False,
        keep_vba=path.suffix.casefold() == ".xlsm",
    )


def ensure_derived_columns(
    workbook: Workbook,
    layout: WorkbookLayout,
    last_row: int,
) -> WorkbookLayout:
    """Localiza ou adiciona as colunas calculadas e a coluna auxiliar."""

    worksheet = workbook[layout.base.title]
    columns = dict(layout.base.columns)
    next_column = max(worksheet.max_column, *columns.values()) + 1

    for canonical, header in DERIVED_HEADERS.items():
        if canonical in columns:
            continue
        column = next_column
        next_column += 1
        columns[canonical] = column
        header_cell = worksheet.cell(layout.base.header_row, column, header)
        _style_new_header(header_cell, auxiliary=canonical == "priority_score")

        source_column = max(1, column - 1)
        for row in range(layout.base.header_row + 1, last_row + 1):
            source = worksheet.cell(row, source_column)
            target = worksheet.cell(row, column)
            target._style = copy(source._style)
            target.fill = PatternFill(
                "solid",
                fgColor=GRAY if canonical == "priority_score" else LIGHT_YELLOW,
            )

    worksheet.column_dimensions[
        get_column_letter(columns["priority_score"])
    ].hidden = True
    worksheet.column_dimensions[
        get_column_letter(columns["total_value"])
    ].width = 15
    worksheet.column_dimensions[
        get_column_letter(columns["lead_days"])
    ].width = 13
    worksheet.column_dimensions[
        get_column_letter(columns["policy_limit"])
    ].width = 15
    worksheet.column_dimensions[
        get_column_letter(columns["policy_status"])
    ].width = 15
    worksheet.column_dimensions[
        get_column_letter(columns["limit_difference"])
    ].width = 18
    worksheet.column_dimensions[
        get_column_letter(columns["priority"])
    ].width = 13

    _extend_or_create_table(
        worksheet=worksheet,
        header_row=layout.base.header_row,
        last_row=last_row,
        last_column=max(columns.values()),
    )

    updated_base = replace(layout.base, columns=columns)
    return replace(layout, base=updated_base)


def create_support_sheet(
    workbook: Workbook,
    layout: WorkbookLayout,
    travels: list[TravelResult],
    rules: dict[str, object],
    last_row: int,
) -> dict[str, str]:
    """Cria premissas visíveis às fórmulas e resumos auxiliares."""

    if SUPPORT_SHEET in workbook.sheetnames:
        del workbook[SUPPORT_SHEET]
    worksheet = workbook.create_sheet(SUPPORT_SHEET)
    worksheet.sheet_view.showGridLines = False
    worksheet.merge_cells("A1:I1")
    worksheet["A1"] = "APOIO DA AUTOMAÇÃO — REGRAS E RESUMOS AUDITÁVEIS"
    worksheet["A1"].fill = PatternFill("solid", fgColor=DARK_BLUE)
    worksheet["A1"].font = Font(color=WHITE, bold=True, size=14)
    worksheet["A1"].alignment = Alignment(horizontal="center")

    worksheet["A3"] = "Regra"
    worksheet["B3"] = "Valor"
    worksheet["C3"] = "Aplicação"
    _style_header_range(worksheet, "A3:C3")

    weights = rules["priority_weights"]
    thresholds = rules["priority_thresholds"]
    rule_rows = [
        ("card_divergent", "Cartão divergente", weights["card_divergent"]),
        ("card_pending", "Cartão pendente", weights["card_pending"]),
        ("card_other_issue", "Outro status de cartão", weights["card_other_issue"]),
        (
            "criticality_emergency",
            "Criticidade emergencial",
            weights["criticality_emergency"],
        ),
        (
            "criticality_executive",
            "Criticidade executiva",
            weights["criticality_executive"],
        ),
        ("outside_policy", "Fora da política", weights["outside_policy"]),
        (
            "policy_not_found",
            "Política não localizada",
            weights["policy_not_found"],
        ),
        ("booking_pending", "Reserva pendente", weights["booking_pending"]),
        (
            "booking_reschedule",
            "Reserva em remarcação",
            weights["booking_reschedule"],
        ),
        (
            "cost_over_limit_max",
            "Máximo por excesso de custo",
            weights["cost_over_limit_max"],
        ),
        (
            "lead_time_shortfall_max",
            "Máximo por falta de antecedência",
            weights["lead_time_shortfall_max"],
        ),
        (
            "total_value_max",
            "Máximo por valor total",
            weights["total_value_max"],
        ),
        ("critical_threshold", "Limite — prioridade crítica", thresholds["critical"]),
        ("high_threshold", "Limite — prioridade alta", thresholds["high"]),
    ]
    rule_cells: dict[str, str] = {}
    for row, (key, label, value) in enumerate(rule_rows, start=4):
        worksheet.cell(row, 1, label)
        worksheet.cell(row, 2, value)
        worksheet.cell(row, 3, "Componente do score de prioridade")
        rule_cells[key] = f"'{SUPPORT_SHEET}'!$B${row}"

    base_sheet = _quoted_sheet(layout.base.title)
    base_columns = layout.base.columns
    first_data_row = layout.base.header_row + 1
    total_range = _absolute_range(
        layout.base.title,
        base_columns["total_value"],
        first_data_row,
        last_row,
    )
    cost_center_range = _absolute_range(
        layout.base.title,
        base_columns["cost_center"],
        first_data_row,
        last_row,
    )
    supplier_range = _absolute_range(
        layout.base.title,
        base_columns["supplier"],
        first_data_row,
        last_row,
    )

    worksheet["E3"] = "Centro de Custo"
    worksheet["F3"] = "Valor Total"
    _style_header_range(worksheet, "E3:F3")
    cost_centers = sorted(
        {item.cost_center for item in travels if item.cost_center},
        key=normalize_text,
    )
    for row, cost_center in enumerate(cost_centers, start=4):
        worksheet.cell(row, 5, cost_center)
        worksheet.cell(
            row,
            6,
            f"=SUMIF({cost_center_range},E{row},{total_range})",
        )
        worksheet.cell(row, 6).number_format = 'R$ #,##0.00'

    worksheet["H3"] = "Fornecedor"
    worksheet["I3"] = "Valor Total"
    _style_header_range(worksheet, "H3:I3")
    suppliers = sorted(
        {item.supplier for item in travels if item.supplier},
        key=normalize_text,
    )
    for row, supplier in enumerate(suppliers, start=4):
        worksheet.cell(row, 8, supplier)
        worksheet.cell(
            row,
            9,
            f"=SUMIF({supplier_range},H{row},{total_range})",
        )
        worksheet.cell(row, 9).number_format = 'R$ #,##0.00'

    for column, width in {"A": 30, "B": 12, "C": 35, "E": 20, "F": 16, "H": 24, "I": 16}.items():
        worksheet.column_dimensions[column].width = width
    for row in worksheet.iter_rows(min_row=4, max_row=max(17, 3 + len(suppliers))):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    # Fonte estática e auditável para a Tabela Dinâmica nativa. Os valores
    # vêm do cálculo Python já validado, enquanto as fórmulas permanecem na
    # Base_Viagens para atender ao teste e permitir auditoria.
    worksheet["K3"] = "Centro de Custo"
    worksheet["L3"] = "Tipo de Serviço"
    worksheet["M3"] = "Valor Total"
    _style_header_range(worksheet, "K3:M3")
    for row, travel in enumerate(travels, start=4):
        worksheet.cell(row, 11, travel.cost_center)
        worksheet.cell(row, 12, travel.service_type)
        worksheet.cell(row, 13, travel.total_value)
        worksheet.cell(row, 13).number_format = 'R$ #,##0.00'
    worksheet.column_dimensions["K"].width = 20
    worksheet.column_dimensions["L"].width = 24
    worksheet.column_dimensions["M"].width = 16

    worksheet.sheet_state = "hidden"
    return {
        **rule_cells,
        "cost_center_labels": (
            f"'{SUPPORT_SHEET}'!$E$4:$E${3 + len(cost_centers)}"
        ),
        "cost_center_values": (
            f"'{SUPPORT_SHEET}'!$F$4:$F${3 + len(cost_centers)}"
        ),
        "supplier_labels": f"'{SUPPORT_SHEET}'!$H$4:$H${3 + len(suppliers)}",
        "supplier_values": f"'{SUPPORT_SHEET}'!$I$4:$I${3 + len(suppliers)}",
        "base_sheet": base_sheet,
    }


def write_base_formulas(
    workbook: Workbook,
    layout: WorkbookLayout,
    policies: dict[str, Policy],
    support_refs: dict[str, str],
    last_row: int,
) -> None:
    """Preenche todas as colunas calculadas com fórmulas auditáveis."""

    worksheet = workbook[layout.base.title]
    columns = layout.base.columns
    first_row = layout.base.header_row + 1
    policy_rows = sorted(policy.worksheet_row for policy in policies.values())
    policy_first = min(policy_rows)
    policy_last = max(policy_rows)
    policy_sheet = layout.policies.title
    policy_columns = layout.policies.columns

    service_policy_range = _absolute_range(
        policy_sheet,
        policy_columns["service_type"],
        policy_first,
        policy_last,
    )
    limit_policy_range = _absolute_range(
        policy_sheet,
        policy_columns["limit_value"],
        policy_first,
        policy_last,
    )
    days_policy_range = _absolute_range(
        policy_sheet,
        policy_columns["min_lead_days"],
        policy_first,
        policy_last,
    )

    for row in range(first_row, last_row + 1):
        reference = {
            key: _cell(columns[key], row)
            for key in (
                "request_date",
                "travel_date",
                "service_type",
                "quantity",
                "unit_value",
                "fees",
                "booking_status",
                "criticality",
                "card_status",
                "total_value",
                "lead_days",
                "policy_limit",
                "policy_status",
                "limit_difference",
                "priority",
                "priority_score",
            )
        }
        min_days_lookup = (
            f"IFERROR(INDEX({days_policy_range},"
            f"MATCH({reference['service_type']},{service_policy_range},0)),0)"
        )

        worksheet[reference["total_value"]] = (
            f"=IFERROR({reference['quantity']}*{reference['unit_value']}"
            f"+{reference['fees']},0)"
        )
        worksheet[reference["lead_days"]] = (
            f"=IFERROR({reference['travel_date']}-{reference['request_date']},0)"
        )
        worksheet[reference["policy_limit"]] = (
            f"=IFERROR(INDEX({limit_policy_range},"
            f"MATCH({reference['service_type']},{service_policy_range},0)),0)"
        )
        worksheet[reference["policy_status"]] = (
            f'=IF({reference["policy_limit"]}=0,"Revisar",'
            f'IF(OR({reference["total_value"]}>{reference["policy_limit"]},'
            f'{reference["lead_days"]}<{min_days_lookup}),"Fora","OK"))'
        )
        worksheet[reference["limit_difference"]] = (
            f"=IFERROR({reference['total_value']}-{reference['policy_limit']},0)"
        )

        card_score = (
            f'IF({reference["card_status"]}="Divergente",'
            f'{support_refs["card_divergent"]},'
            f'IF({reference["card_status"]}="Pendente",'
            f'{support_refs["card_pending"]},'
            f'IF(AND({reference["card_status"]}<>"",'
            f'{reference["card_status"]}<>"Conferido"),'
            f'{support_refs["card_other_issue"]},0)))'
        )
        criticality_score = (
            f'IF({reference["criticality"]}="Emergencial",'
            f'{support_refs["criticality_emergency"]},'
            f'IF({reference["criticality"]}="Executivo",'
            f'{support_refs["criticality_executive"]},0))'
        )
        policy_score = (
            f'IF({reference["policy_status"]}="Revisar",'
            f'{support_refs["policy_not_found"]},'
            f'IF({reference["policy_status"]}="Fora",'
            f'{support_refs["outside_policy"]},0))'
        )
        booking_score = (
            f'IF({reference["booking_status"]}="Pendente",'
            f'{support_refs["booking_pending"]},'
            f'IF({reference["booking_status"]}="Remarcação",'
            f'{support_refs["booking_reschedule"]},0))'
        )
        cost_score = (
            f"IF(AND({reference['policy_limit']}>0,"
            f"{reference['limit_difference']}>0),"
            f"MIN({support_refs['cost_over_limit_max']},"
            f"({reference['limit_difference']}/{reference['policy_limit']})"
            f"*{support_refs['cost_over_limit_max']}),0)"
        )
        lead_score = (
            f"IFERROR(IF({reference['lead_days']}<{min_days_lookup},"
            f"MIN({support_refs['lead_time_shortfall_max']},"
            f"(({min_days_lookup}-{reference['lead_days']})/{min_days_lookup})"
            f"*{support_refs['lead_time_shortfall_max']}),0),0)"
        )
        total_range = (
            f"${get_column_letter(columns['total_value'])}${first_row}:"
            f"${get_column_letter(columns['total_value'])}${last_row}"
        )
        value_score = (
            f"IFERROR(({reference['total_value']}/MAX({total_range}))"
            f"*{support_refs['total_value_max']},0)"
        )
        worksheet[reference["priority_score"]] = (
            f"=ROUND({card_score}+{criticality_score}+{policy_score}"
            f"+{booking_score}+{cost_score}+{lead_score}+{value_score},2)"
        )
        worksheet[reference["priority"]] = (
            f'=IF({reference["priority_score"]}>={support_refs["critical_threshold"]},'
            f'"Crítica",IF({reference["priority_score"]}>='
            f'{support_refs["high_threshold"]},"Alta","Normal"))'
        )

    currency_columns = ("total_value", "policy_limit", "limit_difference")
    for canonical in currency_columns:
        letter = get_column_letter(columns[canonical])
        for row in range(first_row, last_row + 1):
            worksheet[f"{letter}{row}"].number_format = 'R$ #,##0.00'
    score_letter = get_column_letter(columns["priority_score"])
    for row in range(first_row, last_row + 1):
        worksheet[f"{score_letter}{row}"].number_format = "0.00"

    worksheet.cell(
        layout.base.header_row,
        columns["priority"],
    ).comment = Comment(
        "Classificação derivada do Score Prioridade. O score considera custo, "
        "antecedência, cartão corporativo, criticidade, reserva e aderência à política.",
        "Excel Compras Automation",
    )


def write_responses(
    workbook: Workbook,
    layout: WorkbookLayout,
    aliases: dict[str, object],
    travels: list[TravelResult],
    support_refs: dict[str, str],
    last_row: int,
    top_quantity: int,
) -> int:
    """Preenche indicadores, as cinco prioridades e reserva a área do pivô."""

    worksheet = workbook[layout.responses.title]
    indicator_rows = find_indicator_rows(
        worksheet,
        aliases["indicator_labels"],
    )
    answer_column = layout.responses.columns["answer"]
    formula_column = layout.responses.columns.get("formula_used")
    base_columns = layout.base.columns
    first_row = layout.base.header_row + 1

    ranges = {
        key: _absolute_range(layout.base.title, column, first_row, last_row)
        for key, column in base_columns.items()
    }
    formulas = {
        "total_travel_value": f"=SUM({ranges['total_value']})",
        "outside_policy_value": (
            f'=SUMIFS({ranges["total_value"]},{ranges["policy_status"]},"Fora")'
        ),
        "outside_policy_count": (
            f'=COUNTIF({ranges["policy_status"]},"Fora")'
        ),
        "outside_policy_percent": (
            f'=IFERROR(COUNTIF({ranges["policy_status"]},"Fora")/'
            f"COUNTA({ranges['request_id']}),0)"
        ),
        "card_issue_value": (
            f'=SUMIFS({ranges["total_value"]},{ranges["card_status"]},"Pendente")'
            f'+SUMIFS({ranges["total_value"]},{ranges["card_status"]},"Divergente")'
        ),
        "emergency_count": (
            f'=COUNTIF({ranges["criticality"]},"Emergencial")'
        ),
        "top_cost_center": (
            f"=INDEX({support_refs['cost_center_labels']},"
            f"MATCH(MAX({support_refs['cost_center_values']}),"
            f"{support_refs['cost_center_values']},0))"
        ),
        "top_supplier": (
            f"=INDEX({support_refs['supplier_labels']},"
            f"MATCH(MAX({support_refs['supplier_values']}),"
            f"{support_refs['supplier_values']},0))"
        ),
        "potential_savings": (
            f'=SUMIFS({ranges["limit_difference"]},'
            f'{ranges["policy_status"]},"Fora",'
            f'{ranges["limit_difference"]},">0")'
        ),
        "average_ticket": f"=AVERAGE({ranges['total_value']})",
    }
    descriptions = {
        "total_travel_value": "SOMA",
        "outside_policy_value": "SOMASES",
        "outside_policy_count": "CONT.SE",
        "outside_policy_percent": "CONT.SE / CONT.VALORES",
        "card_issue_value": "SOMASES + SOMASES",
        "emergency_count": "CONT.SE",
        "top_cost_center": "ÍNDICE / CORRESP / MÁXIMO",
        "top_supplier": "ÍNDICE / CORRESP / MÁXIMO",
        "potential_savings": "SOMASES com diferença positiva",
        "average_ticket": "MÉDIA",
    }
    currency_indicators = {
        "total_travel_value",
        "outside_policy_value",
        "card_issue_value",
        "potential_savings",
        "average_ticket",
    }

    for key, (row, _) in indicator_rows.items():
        answer = worksheet.cell(row, answer_column)
        answer.value = formulas[key]
        if key in currency_indicators:
            answer.number_format = 'R$ #,##0.00'
        elif key == "outside_policy_percent":
            answer.number_format = "0.0%"
        elif key in {"outside_policy_count", "emergency_count"}:
            answer.number_format = "0"
        if formula_column:
            worksheet.cell(row, formula_column, descriptions[key])

    priority_header_row, priority_columns = _find_priority_table(worksheet)
    _write_top_requests(
        worksheet=worksheet,
        header_row=priority_header_row,
        columns=priority_columns,
        ranked=rank_immediate_requests(travels, top_quantity),
        quantity=top_quantity,
    )

    pivot_title_row = _find_pivot_title_row(worksheet)
    pivot_start_row = pivot_title_row + 2
    _prepare_pivot_area(worksheet, pivot_title_row, pivot_start_row)
    _format_response_sheet(worksheet, priority_header_row)
    return pivot_start_row


def create_fallback_summary_and_chart(
    workbook: Workbook,
    layout: WorkbookLayout,
    travels: list[TravelResult],
    pivot_start_row: int,
    last_row: int,
) -> None:
    """Cria um resumo formula-driven caso o Excel Desktop não esteja disponível."""

    worksheet = workbook[layout.responses.title]
    columns = layout.base.columns
    first_row = layout.base.header_row + 1
    cost_centers = sorted(
        {item.cost_center for item in travels if item.cost_center},
        key=normalize_text,
    )
    services = sorted(
        {item.service_type for item in travels if item.service_type},
        key=normalize_text,
    )
    total_range = _absolute_range(
        layout.base.title,
        columns["total_value"],
        first_row,
        last_row,
    )
    cost_center_range = _absolute_range(
        layout.base.title,
        columns["cost_center"],
        first_row,
        last_row,
    )
    service_range = _absolute_range(
        layout.base.title,
        columns["service_type"],
        first_row,
        last_row,
    )

    worksheet.cell(pivot_start_row, 1, "Centro de Custo")
    for index, service in enumerate(services, start=2):
        worksheet.cell(pivot_start_row, index, service)
    total_column = 2 + len(services)
    worksheet.cell(pivot_start_row, total_column, "Total Geral")
    _style_header_range(
        worksheet,
        f"A{pivot_start_row}:"
        f"{get_column_letter(total_column)}{pivot_start_row}",
    )

    for row_offset, cost_center in enumerate(cost_centers, start=1):
        row = pivot_start_row + row_offset
        worksheet.cell(row, 1, cost_center)
        for column_offset, _ in enumerate(services, start=2):
            service_header = _cell(column_offset, pivot_start_row)
            worksheet.cell(
                row,
                column_offset,
                f"=SUMIFS({total_range},{cost_center_range},$A{row},"
                f"{service_range},{service_header})",
            )
            worksheet.cell(row, column_offset).number_format = 'R$ #,##0.00'
        worksheet.cell(
            row,
            total_column,
            f"=SUM(B{row}:{get_column_letter(total_column - 1)}{row})",
        )
        worksheet.cell(row, total_column).number_format = 'R$ #,##0.00'

    total_row = pivot_start_row + len(cost_centers) + 1
    worksheet.cell(total_row, 1, "Total Geral")
    for column in range(2, total_column + 1):
        letter = get_column_letter(column)
        worksheet.cell(
            total_row,
            column,
            f"=SUM({letter}{pivot_start_row + 1}:{letter}{total_row - 1})",
        )
        worksheet.cell(total_row, column).number_format = 'R$ #,##0.00'
    _style_total_range(
        worksheet,
        f"A{total_row}:{get_column_letter(total_column)}{total_row}",
    )

    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.grouping = "clustered"
    chart.title = "Valor Total por Centro de Custo e Tipo de Serviço"
    chart.y_axis.title = "Valor Total (R$)"
    chart.x_axis.title = "Centro de Custo"
    chart.height = 9
    chart.width = 18
    data = Reference(
        worksheet,
        min_col=2,
        max_col=total_column - 1,
        min_row=pivot_start_row,
        max_row=pivot_start_row + len(cost_centers),
    )
    categories = Reference(
        worksheet,
        min_col=1,
        min_row=pivot_start_row + 1,
        max_row=pivot_start_row + len(cost_centers),
    )
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    chart.legend.position = "b"
    worksheet.add_chart(chart, f"{get_column_letter(total_column + 2)}{pivot_start_row}")


def apply_conditional_formatting(
    workbook: Workbook,
    layout: WorkbookLayout,
    last_row: int,
) -> None:
    """Destaca linhas fora da política e solicitações críticas."""

    worksheet = workbook[layout.base.title]
    first_row = layout.base.header_row + 1
    last_visible_column = layout.base.columns["priority"]
    visible_range = (
        f"A{first_row}:{get_column_letter(last_visible_column)}{last_row}"
    )
    policy_letter = get_column_letter(layout.base.columns["policy_status"])
    priority_letter = get_column_letter(layout.base.columns["priority"])

    critical_rule = FormulaRule(
        formula=[f'${priority_letter}{first_row}="Crítica"'],
        stopIfTrue=True,
        fill=PatternFill("solid", fgColor=LIGHT_RED),
        font=Font(color=DARK_RED, bold=True),
    )
    outside_rule = FormulaRule(
        formula=[f'${policy_letter}{first_row}="Fora"'],
        fill=PatternFill("solid", fgColor=LIGHT_ORANGE),
        font=Font(color="FF9C5700", bold=True),
    )
    worksheet.conditional_formatting.add(visible_range, critical_rule)
    outside_range = (
        f"{policy_letter}{first_row}:"
        f"{get_column_letter(layout.base.columns['limit_difference'])}{last_row}"
    )
    worksheet.conditional_formatting.add(outside_range, outside_rule)
    worksheet.freeze_panes = _cell(1, first_row)


def finalize_workbook(workbook: Workbook) -> None:
    """Solicita recálculo completo na próxima abertura."""

    workbook.calculation.calcMode = "auto"
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True


def _style_new_header(cell, auxiliary: bool) -> None:
    cell.fill = PatternFill("solid", fgColor=GRAY if auxiliary else ORANGE)
    cell.font = Font(color=WHITE if not auxiliary else DARK_BLUE, bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _extend_or_create_table(
    worksheet: Worksheet,
    header_row: int,
    last_row: int,
    last_column: int,
) -> None:
    matching_tables = [
        table
        for table in worksheet.tables.values()
        if table.ref.split(":")[0].rstrip("0123456789")
    ]
    if matching_tables:
        table = matching_tables[0]
        _, _, old_last_column, _ = range_boundaries(table.ref)
        for column in range(old_last_column + 1, last_column + 1):
            table.tableColumns.append(
                TableColumn(
                    id=len(table.tableColumns) + 1,
                    name=str(worksheet.cell(header_row, column).value),
                )
            )
        start = table.ref.split(":")[0]
        end = f"{get_column_letter(last_column)}{last_row}"
        table.ref = f"{start}:{end}"
        return

    table = Table(
        displayName="BaseViagensTable",
        ref=(
            f"A{header_row}:"
            f"{get_column_letter(last_column)}{last_row}"
        ),
    )
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    worksheet.add_table(table)


def _quoted_sheet(title: str) -> str:
    return "'" + title.replace("'", "''") + "'"


def _absolute_range(title: str, column: int, first_row: int, last_row: int) -> str:
    letter = get_column_letter(column)
    return f"{_quoted_sheet(title)}!${letter}${first_row}:${letter}${last_row}"


def _cell(column: int, row: int) -> str:
    return f"{get_column_letter(column)}{row}"


def _style_header_range(worksheet: Worksheet, cell_range: str) -> None:
    for row in worksheet[cell_range]:
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=BLUE)
            cell.font = Font(color=WHITE, bold=True)
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )
            cell.border = Border(bottom=THIN_GRAY)


def _style_total_range(worksheet: Worksheet, cell_range: str) -> None:
    for row in worksheet[cell_range]:
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
            cell.font = Font(color=DARK_BLUE, bold=True)
            cell.border = Border(top=THIN_GRAY)


def _find_priority_table(worksheet: Worksheet) -> tuple[int, dict[str, int]]:
    aliases = {
        "request_id": ["ID Solicitação", "ID da Solicitação"],
        "reason": ["Motivo da prioridade", "Justificativa"],
        "action": ["Ação recomendada", "Ação"],
        "order": ["Ordem (1 a 5)", "Ordem", "Ranking"],
    }
    for row in range(1, min(worksheet.max_row, 80) + 1):
        mapping: dict[str, int] = {}
        for column in range(1, min(worksheet.max_column, 12) + 1):
            value = worksheet.cell(row, column).value
            for canonical, options in aliases.items():
                if value is not None and max(
                    text_similarity(value, option) for option in options
                ) >= 0.86:
                    mapping[canonical] = column
        if len(mapping) >= 3 and "request_id" in mapping:
            defaults = {
                "request_id": mapping["request_id"],
                "reason": mapping.get("reason", mapping["request_id"] + 1),
                "action": mapping.get("action", mapping["request_id"] + 2),
                "order": mapping.get("order", mapping["request_id"] + 3),
            }
            return row, defaults

    row = worksheet.max_row + 2
    return row, {"request_id": 1, "reason": 2, "action": 3, "order": 4}


def _write_top_requests(
    worksheet: Worksheet,
    header_row: int,
    columns: dict[str, int],
    ranked: list[TravelResult],
    quantity: int,
) -> None:
    headers = {
        "request_id": "ID Solicitação",
        "reason": "Motivo da prioridade",
        "action": "Ação recomendada",
        "order": "Ordem (1 a 5)",
    }
    for canonical, title in headers.items():
        worksheet.cell(header_row, columns[canonical], title)
    first_column = min(columns.values())
    last_column = max(columns.values())
    _style_header_range(
        worksheet,
        f"{_cell(first_column, header_row)}:{_cell(last_column, header_row)}",
    )

    for offset in range(1, quantity + 1):
        row = header_row + offset
        for column in range(first_column, last_column + 1):
            cell = worksheet.cell(row, column)
            cell.value = None
            cell.fill = PatternFill("solid", fgColor=LIGHT_YELLOW)
            cell.border = Border(bottom=THIN_GRAY)
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    for order, travel in enumerate(ranked, start=1):
        row = header_row + order
        reasons = "; ".join(travel.reasons[:4]) or "maior score combinado de risco"
        actions = "; ".join(travel.recommended_actions[:3]) or "revisar a solicitação"
        worksheet.cell(row, columns["request_id"], travel.request_id)
        worksheet.cell(row, columns["reason"], reasons[:1].upper() + reasons[1:] + ".")
        worksheet.cell(row, columns["action"], actions[:1].upper() + actions[1:] + ".")
        order_cell = worksheet.cell(row, columns["order"], order)
        order_cell.comment = Comment(
            f"Score de prioridade calculado: {travel.score:.2f}",
            "Excel Compras Automation",
        )
        # Altura fixa generosa porque o Excel não oferece AutoFit confiável
        # para células preenchidas por openpyxl antes da abertura do arquivo.
        worksheet.row_dimensions[row].height = 115


def _find_pivot_title_row(worksheet: Worksheet) -> int:
    aliases = [
        "Tabela dinâmica e gráfico",
        "Tabela dinamica e grafico",
        "Resumo por centro de custo",
    ]
    for row in range(1, min(worksheet.max_row, 100) + 1):
        for column in range(1, min(worksheet.max_column, 12) + 1):
            value = worksheet.cell(row, column).value
            if value is not None and max(
                text_similarity(value, alias) for alias in aliases
            ) >= 0.86:
                return row
    return worksheet.max_row + 2


def _prepare_pivot_area(
    worksheet: Worksheet,
    title_row: int,
    pivot_start_row: int,
) -> None:
    for merged in list(worksheet.merged_cells.ranges):
        if merged.min_row > title_row:
            worksheet.unmerge_cells(str(merged))

    worksheet.merge_cells(
        start_row=title_row,
        start_column=1,
        end_row=title_row,
        end_column=8,
    )
    worksheet.cell(title_row, 1, "Tabela dinâmica e gráfico")
    for cell in worksheet[title_row][0:8]:
        cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
        cell.font = Font(color=DARK_BLUE, bold=True)

    for row in range(title_row + 1, min(worksheet.max_row + 20, title_row + 45)):
        for column in range(1, 16):
            worksheet.cell(row, column).value = None
    worksheet.cell(
        title_row + 1,
        1,
        "Valor Total por Centro de Custo e Tipo de Serviço",
    )
    worksheet.merge_cells(
        start_row=title_row + 1,
        start_column=1,
        end_row=title_row + 1,
        end_column=6,
    )
    worksheet.cell(title_row + 1, 1).font = Font(italic=True, color="FF7F6000")
    worksheet.cell(title_row + 1, 1).alignment = Alignment(
        vertical="center",
        wrap_text=True,
    )
    worksheet.row_dimensions[pivot_start_row - 1].height = 30


def _format_response_sheet(worksheet: Worksheet, priority_header_row: int) -> None:
    worksheet.sheet_view.showGridLines = False
    worksheet.column_dimensions["A"].width = 42
    worksheet.column_dimensions["B"].width = 33
    worksheet.column_dimensions["C"].width = 45
    worksheet.column_dimensions["D"].width = 16
    for row in range(1, max(worksheet.max_row, priority_header_row + 5) + 1):
        for column in range(1, 5):
            worksheet.cell(row, column).alignment = Alignment(
                vertical="center",
                wrap_text=True,
            )
