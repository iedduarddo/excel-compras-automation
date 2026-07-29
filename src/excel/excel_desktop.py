"""Integração opcional com o Excel Desktop para PivotTable nativa."""

from __future__ import annotations

import logging
import os
import platform
import re
import tempfile
import time
from collections.abc import Callable
from contextlib import suppress
from pathlib import Path
from typing import Any, TypeAlias
from zipfile import ZipFile

from src.core.exceptions import ExcelDesktopCleanupError, ExcelDesktopError
from src.core.models import WorkbookLayout

XL_DATABASE = 1
XL_ROW_FIELD = 1
XL_COLUMN_FIELD = 2
XL_SUM = -4157
XL_COLUMN_CLUSTERED = 51
XL_LEGEND_BOTTOM = -4107
XL_CALCULATION_AUTOMATIC = -4105
XL_CALCULATION_MANUAL = -4135
XL_A1 = 1
XL_TABULAR_ROW = 1
XL_UP = -4162
PIVOT_SOURCE_SHEET = "Apoio_Automacao"
_NATIVE_EXCEL_SETTINGS = (
    "Visible",
    "DisplayAlerts",
    "ScreenUpdating",
    "AskToUpdateLinks",
    "EnableEvents",
    "Calculation",
    "CalculateBeforeSave",
)
_RECALCULATION_EXCEL_SETTINGS = (
    "Visible",
    "DisplayAlerts",
    "ScreenUpdating",
    "Calculation",
)
_SAFE_EXCEL_SETTINGS = {
    "Visible": False,
    "DisplayAlerts": True,
    "ScreenUpdating": True,
    "AskToUpdateLinks": True,
    "EnableEvents": True,
    "Calculation": XL_CALCULATION_AUTOMATIC,
    "CalculateBeforeSave": True,
}
_CleanupFailure: TypeAlias = tuple[str, Exception, bool]


def pywin32_is_available() -> bool:
    """Informa se a dependência do Excel Desktop pode ser importada."""

    if platform.system() != "Windows":
        return False
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return True


def create_native_pivot_and_recalculate(
    output_file: Path,
    layout: WorkbookLayout,
    pivot_start_row: int,
    last_row: int,
    logger: logging.Logger,
) -> bool:
    """Cria PivotTable nativa sem depender de um recálculo global bloqueante.

    A fonte da Pivot contém valores estáticos previamente validados em Python.
    As fórmulas permanecem na Base_Viagens e são configuradas para recálculo
    automático quando o arquivo for aberto no Excel.

    Retorna ``False`` porque esta etapa não exige caches recalculados para que a
    Pivot seja criada e validada.
    """

    if not pywin32_is_available():
        raise ExcelDesktopError(
            "pywin32 não está disponível. Instale as dependências do requirements.txt."
        )

    import pythoncom
    import win32com.client

    excel = None
    workbook = None
    saved = False
    com_initialized = False
    operation_error: Exception | None = None
    cleanup_failures: list[_CleanupFailure] = []
    original_settings: dict[str, object] = {}
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        logger.info("Abrindo o Excel Desktop em segundo plano.")
        excel = win32com.client.DispatchEx("Excel.Application")
        original_settings = _capture_excel_settings(
            excel,
            _NATIVE_EXCEL_SETTINGS,
            logger,
        )
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False
        try:
            excel.Calculation = XL_CALCULATION_MANUAL
        except Exception:
            logger.debug(
                "O Excel aplicará o modo manual após abrir a pasta de trabalho.",
                exc_info=True,
            )

        workbook = excel.Workbooks.Open(
            str(output_file.resolve()),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
        )
        excel.Calculation = XL_CALCULATION_MANUAL
        excel.CalculateBeforeSave = False
        response_sheet = workbook.Worksheets(layout.responses.title)
        source_sheet = workbook.Worksheets(PIVOT_SOURCE_SHEET)

        for index in range(response_sheet.ChartObjects().Count, 0, -1):
            response_sheet.ChartObjects(index).Delete()

        source_last_row = (
            source_sheet.Cells(
                source_sheet.Rows.Count,
                11,
            )
            .End(XL_UP)
            .Row
        )
        if source_last_row < 4:
            raise ExcelDesktopError(
                "A fonte estática da Tabela Dinâmica não possui dados."
            )

        source_address = f"'{PIVOT_SOURCE_SHEET}'!R3C11:R{source_last_row}C13"
        destination = response_sheet.Cells(pivot_start_row, 1)

        logger.info("Criando a Tabela Dinâmica nativa.")
        pivot_cache = workbook.PivotCaches().Create(
            SourceType=XL_DATABASE,
            SourceData=source_address,
        )
        pivot_table = pivot_cache.CreatePivotTable(
            TableDestination=destination,
            TableName="PivotValorTotalCentroCusto",
        )

        center_header = str(source_sheet.Cells(3, 11).Value)
        service_header = str(source_sheet.Cells(3, 12).Value)
        total_header = str(source_sheet.Cells(3, 13).Value)

        center_field = pivot_table.PivotFields(center_header)
        center_field.Orientation = XL_ROW_FIELD
        center_field.Position = 1

        service_field = pivot_table.PivotFields(service_header)
        service_field.Orientation = XL_COLUMN_FIELD
        service_field.Position = 1

        value_field = pivot_table.AddDataField(
            pivot_table.PivotFields(total_header),
            "Soma de Valor Total",
            XL_SUM,
        )
        value_field.NumberFormat = "R$ #,##0.00"
        pivot_table.RowAxisLayout(XL_TABULAR_ROW)
        pivot_table.TableStyle2 = "PivotStyleMedium2"
        # Remove os totais gerais da área usada pelo gráfico. A linha
        # "Grand Total" não deve virar uma categoria, pois sua escala achata
        # visualmente as barras dos centros de custo.
        pivot_table.RowGrand = False
        pivot_table.ColumnGrand = False
        try:
            # A criação da Pivot pode executar AutoFit e estreitar as colunas
            # usadas pelas justificativas. Preservamos a formatação definida
            # no arquivo para manter todo o texto gerencial legível.
            pivot_table.HasAutoFormat = False
            pivot_table.PreserveFormatting = True
        except Exception:
            logger.debug(
                "O Excel não aceitou as opções de preservação da Pivot.",
                exc_info=True,
            )

        for column, width in {
            "A:A": 42,
            "B:B": 33,
            "C:C": 45,
            "D:D": 16,
        }.items():
            response_sheet.Columns(column).ColumnWidth = width

        logger.info("Criando o gráfico vinculado à Tabela Dinâmica.")
        anchor = response_sheet.Cells(pivot_start_row, 8)
        chart_object = response_sheet.ChartObjects().Add(
            anchor.Left,
            anchor.Top,
            760,
            380,
        )
        chart_object.Name = "GraficoValorTotalCentroCusto"
        chart = chart_object.Chart
        chart.SetSourceData(pivot_table.TableRange1)
        chart.ChartType = XL_COLUMN_CLUSTERED
        chart.HasTitle = True
        chart.ChartTitle.Text = "Valor Total por Centro de Custo e Tipo de Serviço"
        chart.HasLegend = True
        chart.Legend.Position = XL_LEGEND_BOTTOM
        try:
            chart.Axes(2).TickLabels.NumberFormat = "R$ #,##0"
        except Exception:
            logger.debug(
                "O Excel não aceitou o formato explícito do eixo; "
                "o gráfico continuará usando o formato padrão.",
                exc_info=True,
            )

        workbook.Save()
        saved = True
    except Exception as error:  # noqa: BLE001
        # A automação COM pode falhar com exceções específicas de diferentes
        # versões do Excel. A decisão de propagação ocorre após o cleanup.
        operation_error = error
    finally:
        cleanup_failures = _cleanup_excel_session(
            workbook=workbook,
            excel=excel,
            pythoncom=pythoncom,
            com_initialized=com_initialized,
            close_save_changes=False,
            original_settings=original_settings,
            logger=logger,
        )

    fatal_cleanup_failures = [failure for failure in cleanup_failures if failure[2]]

    if operation_error is not None:
        error_type = (
            ExcelDesktopCleanupError if fatal_cleanup_failures else ExcelDesktopError
        )
        public_error = error_type(
            "O Excel Desktop não conseguiu criar a Tabela Dinâmica nativa. "
            f"Detalhe técnico: {operation_error}"
        )
        _add_cleanup_notes(public_error, cleanup_failures)
        raise public_error from operation_error

    if fatal_cleanup_failures:
        stage, cleanup_error, _ = fatal_cleanup_failures[0]
        public_error = ExcelDesktopCleanupError(
            "O Excel Desktop criou a Tabela Dinâmica, mas não conseguiu "
            f"finalizar todos os recursos. Etapa: {stage}. "
            f"Detalhe técnico: {cleanup_error}"
        )
        _add_cleanup_notes(public_error, cleanup_failures)
        raise public_error from cleanup_error

    if saved:
        _set_calculation_on_open(output_file)
    logger.info("Excel Desktop concluiu a Tabela Dinâmica nativa.")
    return False


def recalculate_only(
    output_file: Path,
    logger: logging.Logger,
) -> None:
    """Recalcula e salva um arquivo que usa o resumo de compatibilidade."""

    if not pywin32_is_available():
        return

    import pythoncom
    import win32com.client

    excel = None
    workbook = None
    saved = False
    com_initialized = False
    operation_error: Exception | None = None
    cleanup_failures: list[_CleanupFailure] = []
    original_settings: dict[str, object] = {}
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        excel = win32com.client.DispatchEx("Excel.Application")
        original_settings = _capture_excel_settings(
            excel,
            _RECALCULATION_EXCEL_SETTINGS,
            logger,
        )
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        workbook = excel.Workbooks.Open(
            str(output_file.resolve()),
            UpdateLinks=0,
            ReadOnly=False,
        )
        excel.Calculation = XL_CALCULATION_AUTOMATIC
        excel.CalculateFullRebuild()
        _wait_for_calculation(excel, timeout_seconds=30)
        workbook.Save()
        saved = True
    except Exception as error:  # noqa: BLE001
        # Mantém o erro original para propagá-lo somente depois de liberar
        # todos os recursos COM adquiridos.
        operation_error = error
    finally:
        cleanup_failures = _cleanup_excel_session(
            workbook=workbook,
            excel=excel,
            pythoncom=pythoncom,
            com_initialized=com_initialized,
            close_save_changes=saved,
            original_settings=original_settings,
            logger=logger,
        )

    fatal_cleanup_failures = [failure for failure in cleanup_failures if failure[2]]

    if operation_error is not None and fatal_cleanup_failures:
        stage, cleanup_error, _ = fatal_cleanup_failures[0]
        public_error = ExcelDesktopCleanupError(
            "O Excel Desktop não concluiu o recálculo e também não conseguiu "
            f"finalizar todos os recursos. Etapa: {stage}. "
            f"Erro original: {operation_error}. "
            f"Detalhe da limpeza: {cleanup_error}"
        )
        _add_cleanup_notes(public_error, cleanup_failures)
        raise public_error from operation_error

    if operation_error is not None:
        _add_cleanup_notes(operation_error, cleanup_failures)
        raise operation_error

    if fatal_cleanup_failures:
        stage, cleanup_error, _ = fatal_cleanup_failures[0]
        public_error = ExcelDesktopCleanupError(
            "O Excel Desktop recalculou o arquivo, mas não conseguiu "
            f"finalizar todos os recursos. Etapa: {stage}. "
            f"Detalhe técnico: {cleanup_error}"
        )
        _add_cleanup_notes(public_error, cleanup_failures)
        raise public_error from cleanup_error

    logger.info("Fórmulas recalculadas pelo Excel Desktop.")


def _add_cleanup_notes(
    error: BaseException,
    failures: list[_CleanupFailure],
) -> None:
    """Acrescenta contexto das falhas secundárias sem substituir a principal."""

    for stage, cleanup_error, _ in failures:
        error.add_note(f"Falha adicional durante {stage}: {cleanup_error}")


def _capture_excel_settings(
    excel: object,
    property_names: tuple[str, ...],
    logger: logging.Logger,
) -> dict[str, object]:
    """Registra valores simples para restaurá-los se o encerramento falhar."""

    settings: dict[str, object] = {}
    for property_name in property_names:
        try:
            value = getattr(excel, property_name)
        except Exception:
            logger.debug(
                "Não foi possível ler Excel.%s antes da automação.",
                property_name,
                exc_info=True,
            )
            value = _SAFE_EXCEL_SETTINGS[property_name]

        if value is not None and not isinstance(value, (bool, int, float, str)):
            value = _SAFE_EXCEL_SETTINGS[property_name]

        settings[property_name] = value

    return settings


def _cleanup_excel_session(
    *,
    workbook: Any | None,
    excel: Any | None,
    pythoncom: Any,
    com_initialized: bool,
    close_save_changes: bool,
    original_settings: dict[str, object],
    logger: logging.Logger,
) -> list[_CleanupFailure]:
    """Libera cada recurso COM sem interromper as etapas seguintes."""

    failures: list[_CleanupFailure] = []

    if workbook is not None:
        _attempt_cleanup(
            stage="fechar a pasta de trabalho",
            action=lambda: workbook.Close(SaveChanges=close_save_changes),
            critical_without_primary=True,
            failures=failures,
            logger=logger,
        )

    quit_succeeded = True
    if excel is not None:
        quit_succeeded = _attempt_cleanup(
            stage="encerrar o Excel Desktop",
            action=lambda: excel.Quit(),
            critical_without_primary=True,
            failures=failures,
            logger=logger,
        )

    if excel is not None and not quit_succeeded:
        for property_name, value in reversed(original_settings.items()):
            _attempt_cleanup(
                stage=f"restaurar Excel.{property_name}",
                action=lambda name=property_name, setting=value: setattr(
                    excel,
                    name,
                    setting,
                ),
                critical_without_primary=False,
                failures=failures,
                logger=logger,
            )

    if com_initialized:
        _attempt_cleanup(
            stage="liberar a inicialização COM",
            action=lambda: pythoncom.CoUninitialize(),
            critical_without_primary=False,
            failures=failures,
            logger=logger,
        )

    return failures


def _attempt_cleanup(
    *,
    stage: str,
    action: Callable[[], object],
    critical_without_primary: bool,
    failures: list[_CleanupFailure],
    logger: logging.Logger,
) -> bool:
    """Executa uma etapa de limpeza e registra a falha sem interromper o fluxo."""

    try:
        action()
    except Exception as error:  # noqa: BLE001
        failures.append((stage, error, critical_without_primary))
        with suppress(Exception):
            logger.warning(
                "Falha durante a limpeza do Excel Desktop (%s): %s",
                stage,
                error,
                exc_info=True,
            )
        return False
    return True


def _wait_for_calculation(excel, timeout_seconds: int) -> None:
    deadline = time.monotonic() + timeout_seconds
    while excel.CalculationState != 0 and time.monotonic() < deadline:
        time.sleep(0.2)
    if excel.CalculationState != 0:
        raise ExcelDesktopError(
            f"O recálculo do Excel excedeu {timeout_seconds} segundos."
        )


def _set_calculation_on_open(output_file: Path) -> None:
    """Mantém a Pivot e marca o arquivo para recálculo automático ao abrir."""

    calc_pattern = re.compile(
        rb"<calcPr\b[^>]*/>|<calcPr\b[^>]*>.*?</calcPr>",
        flags=re.DOTALL,
    )
    calc_element = (
        b'<calcPr calcId="191029" calcMode="auto" '
        b'fullCalcOnLoad="1" forceFullCalc="1"/>'
    )

    with tempfile.NamedTemporaryFile(
        dir=output_file.parent,
        prefix=f".{output_file.stem}_",
        suffix=output_file.suffix,
        delete=False,
    ) as temporary_handle:
        temporary_path = Path(temporary_handle.name)

    try:
        with (
            ZipFile(output_file, "r") as source_archive,
            ZipFile(temporary_path, "w") as target_archive,
        ):
            for entry in source_archive.infolist():
                data = source_archive.read(entry.filename)
                if entry.filename == "xl/workbook.xml":
                    if calc_pattern.search(data):
                        data = calc_pattern.sub(calc_element, data, count=1)
                    else:
                        data = data.replace(
                            b"</workbook>",
                            calc_element + b"</workbook>",
                            1,
                        )
                target_archive.writestr(entry, data)
        os.replace(temporary_path, output_file)
    finally:
        temporary_path.unlink(missing_ok=True)
