"""Fachada compatível para escrita, fórmulas, análises e formatação."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook

from src.excel._writer_common import (
    BLUE,
    DARK_BLUE,
    DARK_RED,
    GRAY,
    LIGHT_BLUE,
    LIGHT_ORANGE,
    LIGHT_RED,
    LIGHT_YELLOW,
    ORANGE,
    THIN_GRAY,
    WHITE,
    _absolute_range,
    _cell,
    _quoted_sheet,
    _style_header_range,
    _style_total_range,
)
from src.excel.writer_base import (
    DERIVED_HEADERS,
    _extend_or_create_table,
    _style_new_header,
    apply_conditional_formatting,
    ensure_derived_columns,
    write_base_formulas,
)
from src.excel.writer_responses import (
    _find_pivot_title_row,
    _find_priority_table,
    _format_response_sheet,
    _prepare_pivot_area,
    _write_top_requests,
    create_fallback_summary_and_chart,
    write_responses,
)
from src.excel.writer_support import SUPPORT_SHEET, create_support_sheet

_COMPATIBILITY_EXPORTS = (
    BLUE,
    DARK_BLUE,
    DARK_RED,
    DERIVED_HEADERS,
    GRAY,
    LIGHT_BLUE,
    LIGHT_ORANGE,
    LIGHT_RED,
    LIGHT_YELLOW,
    ORANGE,
    SUPPORT_SHEET,
    THIN_GRAY,
    WHITE,
    _absolute_range,
    _cell,
    _extend_or_create_table,
    _find_pivot_title_row,
    _find_priority_table,
    _format_response_sheet,
    _prepare_pivot_area,
    _quoted_sheet,
    _style_header_range,
    _style_new_header,
    _style_total_range,
    _write_top_requests,
)

__all__ = [
    "load_source_workbook",
    "ensure_derived_columns",
    "create_support_sheet",
    "write_base_formulas",
    "write_responses",
    "create_fallback_summary_and_chart",
    "apply_conditional_formatting",
    "finalize_workbook",
]


def load_source_workbook(path: Path) -> Workbook:
    """Abre .xlsx ou .xlsm preservando macros quando necessário."""

    return load_workbook(
        path,
        data_only=False,
        keep_vba=path.suffix.casefold() == ".xlsm",
    )


def finalize_workbook(workbook: Workbook) -> None:
    """Solicita recálculo completo na próxima abertura."""

    workbook.calculation.calcMode = "auto"
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
