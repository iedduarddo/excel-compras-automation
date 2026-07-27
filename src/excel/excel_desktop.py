"""Integração opcional com o Excel Desktop para PivotTable nativa."""

from __future__ import annotations

import logging
import os
import platform
import re
import tempfile
import time
from pathlib import Path
from zipfile import ZipFile

from src.core.exceptions import ExcelDesktopError
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
    pythoncom.CoInitialize()
    try:
        logger.info("Abrindo o Excel Desktop em segundo plano.")
        excel = win32com.client.DispatchEx("Excel.Application")
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
        logger.info("Excel Desktop concluiu a Tabela Dinâmica nativa.")
    except Exception as error:
        raise ExcelDesktopError(
            "O Excel Desktop não conseguiu criar a Tabela Dinâmica nativa. "
            f"Detalhe técnico: {error}"
        ) from error
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            try:
                excel.EnableEvents = True
                excel.Calculation = XL_CALCULATION_AUTOMATIC
            except (pythoncom.com_error, AttributeError):
                logger.debug(
                    "Não foi possível restaurar todas as opções do Excel.",
                    exc_info=True,
                )
            excel.Quit()
        pythoncom.CoUninitialize()

    if saved:
        _set_calculation_on_open(output_file)
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
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
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
        logger.info("Fórmulas recalculadas pelo Excel Desktop.")
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=saved)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


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
