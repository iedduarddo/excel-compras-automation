"""Testes seguros da integração opcional, sem iniciar o Excel Desktop real."""

from __future__ import annotations

import builtins
import logging
import sys
from pathlib import Path
from types import ModuleType, SimpleNamespace
from unittest.mock import Mock
from zipfile import ZipFile

import pytest

import src.excel.excel_desktop as desktop
from src.core.exceptions import ExcelDesktopCleanupError, ExcelDesktopError
from src.core.models import SheetLayout, WorkbookLayout


def install_fake_com_modules(
    monkeypatch: pytest.MonkeyPatch,
    excel: Mock,
) -> tuple[ModuleType, ModuleType]:
    pythoncom = ModuleType("pythoncom")
    pythoncom.CoInitialize = Mock()
    pythoncom.CoUninitialize = Mock()
    pythoncom.com_error = RuntimeError

    client = ModuleType("win32com.client")
    client.DispatchEx = Mock(return_value=excel)
    win32com = ModuleType("win32com")
    win32com.client = client

    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    return pythoncom, client


class FakeComError(Exception):
    """Erro COM controlado, sem depender do pywin32 real."""


class RecordingExcel:
    """Objeto Excel mínimo que registra atribuições relevantes."""

    _tracked_properties = {
        "Visible",
        "DisplayAlerts",
        "ScreenUpdating",
        "AskToUpdateLinks",
        "EnableEvents",
        "Calculation",
        "CalculateBeforeSave",
    }

    def __init__(
        self,
        events: list[tuple[object, ...]],
        workbooks: object,
        *,
        fail_initial_manual: bool = False,
        fail_restore: bool = False,
        fail_read_property: str | None = None,
        fail_quit_lookup: bool = False,
    ) -> None:
        object.__setattr__(self, "_events", events)
        object.__setattr__(self, "_recording", False)
        object.__setattr__(self, "_manual_assignments", 0)
        object.__setattr__(self, "_fail_initial_manual", fail_initial_manual)
        object.__setattr__(self, "_fail_restore", fail_restore)
        object.__setattr__(self, "_fail_read_property", fail_read_property)
        object.__setattr__(self, "_read_failure_consumed", False)
        object.__setattr__(self, "_fail_quit_lookup", fail_quit_lookup)
        object.__setattr__(self, "Workbooks", workbooks)
        object.__setattr__(self, "Visible", True)
        object.__setattr__(self, "DisplayAlerts", True)
        object.__setattr__(self, "ScreenUpdating", True)
        object.__setattr__(self, "AskToUpdateLinks", True)
        object.__setattr__(self, "EnableEvents", True)
        object.__setattr__(
            self,
            "Calculation",
            desktop.XL_CALCULATION_AUTOMATIC,
        )
        object.__setattr__(self, "CalculateBeforeSave", True)
        object.__setattr__(self, "Quit", event_mock(events, "excel.quit"))
        object.__setattr__(
            self,
            "CalculateFullRebuild",
            event_mock(events, "excel.calculate_full_rebuild"),
        )
        object.__setattr__(self, "_recording", True)

    def __getattribute__(self, name: str) -> object:
        if name == "Quit" and object.__getattribute__(self, "_fail_quit_lookup"):
            raise FakeComError("lookup de Quit indisponível")

        fail_read_property = object.__getattribute__(self, "_fail_read_property")
        read_failure_consumed = object.__getattribute__(
            self,
            "_read_failure_consumed",
        )
        if name == fail_read_property and not read_failure_consumed:
            object.__setattr__(self, "_read_failure_consumed", True)
            raise FakeComError(f"leitura de {name} indisponível")

        return object.__getattribute__(self, name)

    def __setattr__(self, name: str, value: object) -> None:
        if self._recording and name in self._tracked_properties:
            self._events.append(("excel.set", name, value))
            if name == "Calculation" and value == desktop.XL_CALCULATION_MANUAL:
                manual_assignments = self._manual_assignments
                object.__setattr__(
                    self,
                    "_manual_assignments",
                    manual_assignments + 1,
                )
                if self._fail_initial_manual and manual_assignments == 0:
                    raise RuntimeError("modo manual inicial indisponível")
            if name == "EnableEvents" and value is True and self._fail_restore:
                raise FakeComError("restauração indisponível")
        object.__setattr__(self, name, value)


class RecordingPivot:
    """Pivot mínima com uma falha opcional de formatação não essencial."""

    def __init__(
        self,
        events: list[tuple[object, ...]],
        fields: dict[str, object],
        *,
        fail_optional_formatting: bool,
    ) -> None:
        object.__setattr__(self, "_events", events)
        object.__setattr__(
            self,
            "_fail_optional_formatting",
            fail_optional_formatting,
        )
        object.__setattr__(self, "_recording", False)
        object.__setattr__(
            self,
            "PivotFields",
            event_mock(
                events,
                "pivot.fields",
                side_effect=lambda name: fields[name],
            ),
        )
        object.__setattr__(
            self,
            "AddDataField",
            event_mock(
                events,
                "pivot.add_data_field",
                return_value=fields["data"],
            ),
        )
        object.__setattr__(
            self,
            "RowAxisLayout",
            event_mock(events, "pivot.row_axis_layout"),
        )
        object.__setattr__(self, "TableRange1", object())
        object.__setattr__(self, "_recording", True)

    def __setattr__(self, name: str, value: object) -> None:
        if (
            self._recording
            and name == "HasAutoFormat"
            and self._fail_optional_formatting
        ):
            self._events.append(("pivot.set", name, value))
            raise RuntimeError("formatação opcional indisponível")
        object.__setattr__(self, name, value)


def event_mock(
    events: list[tuple[object, ...]],
    name: str,
    *,
    return_value: object = None,
    side_effect=None,
) -> Mock:
    """Cria um Mock estrito o bastante para registrar ordem e argumentos."""

    def invoke(*args: object, **kwargs: object) -> object:
        events.append((name, args, kwargs))
        if isinstance(side_effect, BaseException):
            raise side_effect
        if side_effect is not None:
            return side_effect(*args, **kwargs)
        return return_value

    return Mock(side_effect=invoke)


def make_layout() -> WorkbookLayout:
    return WorkbookLayout(
        base=SheetLayout("Base", 1, {}),
        policies=SheetLayout("Políticas", 1, {}),
        responses=SheetLayout("Respostas", 1, {}),
    )


def make_native_graph(
    *,
    source_last_row: int = 6,
    open_error: Exception | None = None,
    save_error: Exception | None = None,
    fail_initial_manual: bool = False,
    fail_optional_pivot_formatting: bool = False,
    fail_axes_formatting: bool = False,
    fail_restore: bool = False,
    fail_read_property: str | None = None,
    fail_quit_lookup: bool = False,
) -> SimpleNamespace:
    """Monta o grafo COM mínimo usado pelo fluxo de Pivot nativa."""

    events: list[tuple[object, ...]] = []
    old_chart_1 = SimpleNamespace(Delete=event_mock(events, "old_chart_1.delete"))
    old_chart_2 = SimpleNamespace(Delete=event_mock(events, "old_chart_2.delete"))

    chart_title = SimpleNamespace(Text=None)
    legend = SimpleNamespace(Position=None)
    tick_labels = SimpleNamespace(NumberFormat=None)
    axis = SimpleNamespace(TickLabels=tick_labels)
    axes_error = RuntimeError("eixo indisponível") if fail_axes_formatting else None
    chart = SimpleNamespace(
        SetSourceData=event_mock(events, "chart.set_source_data"),
        Axes=event_mock(
            events,
            "chart.axes",
            return_value=axis,
            side_effect=axes_error,
        ),
        ChartTitle=chart_title,
        Legend=legend,
    )
    chart_object = SimpleNamespace(Name=None, Chart=chart)
    chart_collection = SimpleNamespace(
        Count=2,
        Add=event_mock(
            events,
            "chart_objects.add",
            return_value=chart_object,
        ),
    )

    def chart_objects(index: int | None = None) -> object:
        if index is None:
            return chart_collection
        return {1: old_chart_1, 2: old_chart_2}[index]

    destination = object()
    anchor = SimpleNamespace(Left=100, Top=200)

    def response_cells(row: int, column: int) -> object:
        return {
            (20, 1): destination,
            (20, 8): anchor,
        }[(row, column)]

    response_columns: dict[str, SimpleNamespace] = {}

    def columns(name: str) -> SimpleNamespace:
        response_columns.setdefault(name, SimpleNamespace(ColumnWidth=None))
        return response_columns[name]

    response_sheet = SimpleNamespace(
        ChartObjects=event_mock(
            events,
            "response.chart_objects",
            side_effect=chart_objects,
        ),
        Cells=event_mock(
            events,
            "response.cells",
            side_effect=response_cells,
        ),
        Columns=event_mock(
            events,
            "response.columns",
            side_effect=columns,
        ),
    )

    end_cell = SimpleNamespace(
        End=event_mock(
            events,
            "source.end",
            return_value=SimpleNamespace(Row=source_last_row),
        )
    )
    headers = {
        (3, 11): "Centro de Custo",
        (3, 12): "Tipo de Serviço",
        (3, 13): "Valor Total",
    }
    rows = SimpleNamespace(Count=1_048_576)

    def source_cells(row: int, column: int) -> object:
        if (row, column) == (rows.Count, 11):
            return end_cell
        return SimpleNamespace(Value=headers[(row, column)])

    source_sheet = SimpleNamespace(
        Rows=rows,
        Cells=event_mock(
            events,
            "source.cells",
            side_effect=source_cells,
        ),
    )

    fields = {
        "Centro de Custo": SimpleNamespace(Orientation=None, Position=None),
        "Tipo de Serviço": SimpleNamespace(Orientation=None, Position=None),
        "Valor Total": SimpleNamespace(),
        "data": SimpleNamespace(NumberFormat=None),
    }
    pivot = RecordingPivot(
        events,
        fields,
        fail_optional_formatting=fail_optional_pivot_formatting,
    )
    pivot_cache = SimpleNamespace(
        CreatePivotTable=event_mock(
            events,
            "pivot_cache.create_pivot_table",
            return_value=pivot,
        )
    )
    pivot_caches = SimpleNamespace(
        Create=event_mock(
            events,
            "pivot_caches.create",
            return_value=pivot_cache,
        )
    )

    sheets = {
        "Respostas": response_sheet,
        desktop.PIVOT_SOURCE_SHEET: source_sheet,
    }
    save_side_effect = save_error
    workbook = SimpleNamespace(
        Worksheets=event_mock(
            events,
            "workbook.worksheets",
            side_effect=lambda name: sheets[name],
        ),
        PivotCaches=event_mock(
            events,
            "workbook.pivot_caches",
            return_value=pivot_caches,
        ),
        Save=event_mock(
            events,
            "workbook.save",
            side_effect=save_side_effect,
        ),
        Close=event_mock(events, "workbook.close"),
    )
    workbooks = SimpleNamespace(
        Open=event_mock(
            events,
            "workbooks.open",
            return_value=workbook,
            side_effect=open_error,
        )
    )
    excel = RecordingExcel(
        events,
        workbooks,
        fail_initial_manual=fail_initial_manual,
        fail_restore=fail_restore,
        fail_read_property=fail_read_property,
        fail_quit_lookup=fail_quit_lookup,
    )
    return SimpleNamespace(
        events=events,
        excel=excel,
        workbooks=workbooks,
        workbook=workbook,
        response_sheet=response_sheet,
        response_columns=response_columns,
        source_sheet=source_sheet,
        pivot_caches=pivot_caches,
        pivot_cache=pivot_cache,
        pivot=pivot,
        fields=fields,
        chart=chart,
        chart_object=chart_object,
        axis=axis,
        destination=destination,
        anchor=anchor,
        old_chart_1=old_chart_1,
        old_chart_2=old_chart_2,
    )


def install_recording_com(
    monkeypatch: pytest.MonkeyPatch,
    graph: SimpleNamespace,
    *,
    dispatch_error: Exception | None = None,
) -> tuple[ModuleType, ModuleType]:
    pythoncom, client = install_fake_com_modules(monkeypatch, graph.excel)
    pythoncom.com_error = FakeComError
    pythoncom.CoInitialize.side_effect = lambda: graph.events.append(
        ("pythoncom.initialize", (), {})
    )
    pythoncom.CoUninitialize.side_effect = lambda: graph.events.append(
        ("pythoncom.uninitialize", (), {})
    )

    def dispatch(application: str) -> object:
        graph.events.append(("client.dispatch", (application,), {}))
        if dispatch_error is not None:
            raise dispatch_error
        return graph.excel

    client.DispatchEx.side_effect = dispatch
    return pythoncom, client


def event_names(graph: SimpleNamespace) -> list[str]:
    return [str(item[0]) for item in graph.events]


def test_pywin32_availability_short_circuits_outside_windows(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(desktop.platform, "system", lambda: "Linux")

    assert desktop.pywin32_is_available() is False


def test_pywin32_availability_handles_missing_and_present_import(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(desktop.platform, "system", lambda: "Windows")
    original_import = builtins.__import__

    def missing_import(name: str, *args: object, **kwargs: object) -> object:
        if name == "win32com.client":
            raise ImportError("ausente")
        return original_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", missing_import)
    assert desktop.pywin32_is_available() is False

    client = ModuleType("win32com.client")
    win32com = ModuleType("win32com")
    win32com.client = client
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    monkeypatch.setattr(builtins, "__import__", original_import)

    assert desktop.pywin32_is_available() is True


def test_wait_for_calculation_finishes_without_real_sleep(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    states = iter([1, 0, 0])

    class FakeExcel:
        @property
        def CalculationState(self) -> int:
            return next(states)

    sleep = Mock()
    monkeypatch.setattr(desktop.time, "monotonic", Mock(side_effect=[0.0, 0.1]))
    monkeypatch.setattr(desktop.time, "sleep", sleep)

    desktop._wait_for_calculation(FakeExcel(), timeout_seconds=1)

    sleep.assert_called_once_with(0.2)


def test_wait_for_calculation_reports_timeout(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    excel = Mock()
    excel.CalculationState = 1
    monkeypatch.setattr(desktop.time, "monotonic", Mock(side_effect=[0.0, 2.0]))

    with pytest.raises(ExcelDesktopError, match="excedeu 1 segundos"):
        desktop._wait_for_calculation(excel, timeout_seconds=1)


@pytest.mark.parametrize(
    "workbook_xml",
    [
        b'<workbook><calcPr calcId="1"/></workbook>',
        b'<workbook><calcPr calcId="1"></calcPr></workbook>',
        b"<workbook><sheets/></workbook>",
    ],
)
def test_set_calculation_on_open_updates_ooxml_safely(
    workbook_xml: bytes,
    tmp_path: Path,
) -> None:
    output_file = tmp_path / "resultado.xlsx"
    with ZipFile(output_file, "w") as archive:
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("docProps/app.xml", b"<Properties/>")

    desktop._set_calculation_on_open(output_file)

    with ZipFile(output_file) as archive:
        updated = archive.read("xl/workbook.xml")
        preserved = archive.read("docProps/app.xml")
    assert updated.count(b"<calcPr") == 1
    assert b'calcMode="auto"' in updated
    assert b'fullCalcOnLoad="1"' in updated
    assert preserved == b"<Properties/>"
    assert not list(tmp_path.glob(".resultado_*.xlsx"))


def test_set_calculation_on_open_cleans_temporary_file_on_replace_error(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    output_file = tmp_path / "resultado.xlsx"
    with ZipFile(output_file, "w") as archive:
        archive.writestr("xl/workbook.xml", b"<workbook></workbook>")
    original = output_file.read_bytes()
    monkeypatch.setattr(
        desktop.os,
        "replace",
        Mock(side_effect=OSError("arquivo bloqueado")),
    )

    with pytest.raises(OSError, match="arquivo bloqueado"):
        desktop._set_calculation_on_open(output_file)

    assert output_file.read_bytes() == original
    assert not list(tmp_path.glob(".resultado_*.xlsx"))


def test_recalculate_only_returns_when_integration_is_unavailable(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: False)

    desktop.recalculate_only(tmp_path / "resultado.xlsx", Mock(spec=logging.Logger))


def test_recalculate_only_uses_com_lifecycle_and_saves(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    workbook = Mock()
    workbooks = Mock()
    workbooks.Open.return_value = workbook
    excel = Mock()
    excel.Workbooks = workbooks
    pythoncom, client = install_fake_com_modules(monkeypatch, excel)
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: True)
    wait = Mock()
    monkeypatch.setattr(desktop, "_wait_for_calculation", wait)
    output_file = tmp_path / "resultado.xlsx"
    logger = Mock(spec=logging.Logger)

    desktop.recalculate_only(output_file, logger)

    pythoncom.CoInitialize.assert_called_once_with()
    client.DispatchEx.assert_called_once_with("Excel.Application")
    workbooks.Open.assert_called_once_with(
        str(output_file.resolve()),
        UpdateLinks=0,
        ReadOnly=False,
    )
    assert excel.Visible is False
    assert excel.Calculation == desktop.XL_CALCULATION_AUTOMATIC
    excel.CalculateFullRebuild.assert_called_once_with()
    wait.assert_called_once_with(excel, timeout_seconds=30)
    workbook.Save.assert_called_once_with()
    workbook.Close.assert_called_once_with(SaveChanges=True)
    excel.Quit.assert_called_once_with()
    pythoncom.CoUninitialize.assert_called_once_with()
    logger.info.assert_called_once()


def test_recalculate_only_closes_without_saving_after_failure(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    workbook = Mock()
    workbooks = Mock()
    workbooks.Open.return_value = workbook
    excel = Mock()
    excel.Workbooks = workbooks
    excel.CalculateFullRebuild.side_effect = RuntimeError("falha de recálculo")
    pythoncom, _ = install_fake_com_modules(monkeypatch, excel)
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: True)

    with pytest.raises(RuntimeError, match="falha de recálculo"):
        desktop.recalculate_only(
            tmp_path / "resultado.xlsx",
            Mock(spec=logging.Logger),
        )

    workbook.Save.assert_not_called()
    workbook.Close.assert_called_once_with(SaveChanges=False)
    excel.Quit.assert_called_once_with()
    pythoncom.CoUninitialize.assert_called_once_with()


def test_native_pivot_requires_pywin32(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: False)

    with pytest.raises(ExcelDesktopError, match="pywin32 não está disponível"):
        desktop.create_native_pivot_and_recalculate(
            tmp_path / "resultado.xlsx",
            Mock(),
            pivot_start_row=20,
            last_row=40,
            logger=Mock(spec=logging.Logger),
        )


def test_native_pivot_characterizes_successful_com_sequence(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    graph = make_native_graph(source_last_row=6)
    pythoncom, client = install_recording_com(monkeypatch, graph)
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: True)
    output_file = tmp_path / "resultado.xlsx"
    postprocess = Mock(
        side_effect=lambda path: graph.events.append(
            ("postprocess.calculation_on_open", (path,), {})
        )
    )
    monkeypatch.setattr(desktop, "_set_calculation_on_open", postprocess)
    logger = Mock(spec=logging.Logger)

    result = desktop.create_native_pivot_and_recalculate(
        output_file,
        make_layout(),
        pivot_start_row=20,
        last_row=40,
        logger=logger,
    )

    assert result is False
    pythoncom.CoInitialize.assert_called_once_with()
    client.DispatchEx.assert_called_once_with("Excel.Application")
    graph.workbooks.Open.assert_called_once_with(
        str(output_file.resolve()),
        UpdateLinks=0,
        ReadOnly=False,
        IgnoreReadOnlyRecommended=True,
    )
    assert graph.workbook.Worksheets.call_args_list[0].args == ("Respostas",)
    assert graph.workbook.Worksheets.call_args_list[1].args == (
        desktop.PIVOT_SOURCE_SHEET,
    )
    assert ("source.end", (desktop.XL_UP,), {}) in graph.events
    graph.pivot_caches.Create.assert_called_once_with(
        SourceType=desktop.XL_DATABASE,
        SourceData="'Apoio_Automacao'!R3C11:R6C13",
    )
    graph.pivot_cache.CreatePivotTable.assert_called_once_with(
        TableDestination=graph.destination,
        TableName="PivotValorTotalCentroCusto",
    )
    assert graph.fields["Centro de Custo"].Orientation == desktop.XL_ROW_FIELD
    assert graph.fields["Centro de Custo"].Position == 1
    assert graph.fields["Tipo de Serviço"].Orientation == desktop.XL_COLUMN_FIELD
    assert graph.fields["Tipo de Serviço"].Position == 1
    graph.pivot.AddDataField.assert_called_once_with(
        graph.fields["Valor Total"],
        "Soma de Valor Total",
        desktop.XL_SUM,
    )
    assert graph.fields["data"].NumberFormat == "R$ #,##0.00"
    graph.pivot.RowAxisLayout.assert_called_once_with(desktop.XL_TABULAR_ROW)
    assert graph.pivot.TableStyle2 == "PivotStyleMedium2"
    assert graph.pivot.RowGrand is False
    assert graph.pivot.ColumnGrand is False
    assert graph.pivot.HasAutoFormat is False
    assert graph.pivot.PreserveFormatting is True
    assert {
        column: cell.ColumnWidth for column, cell in graph.response_columns.items()
    } == {
        "A:A": 42,
        "B:B": 33,
        "C:C": 45,
        "D:D": 16,
    }
    graph.response_sheet.ChartObjects.assert_any_call(2)
    graph.response_sheet.ChartObjects.assert_any_call(1)
    assert event_names(graph).index("old_chart_2.delete") < event_names(graph).index(
        "old_chart_1.delete"
    )
    graph.response_sheet.ChartObjects().Add.assert_called_once_with(
        100,
        200,
        760,
        380,
    )
    assert graph.chart_object.Name == "GraficoValorTotalCentroCusto"
    graph.chart.SetSourceData.assert_called_once_with(graph.pivot.TableRange1)
    assert graph.chart.ChartType == desktop.XL_COLUMN_CLUSTERED
    assert graph.chart.HasTitle is True
    assert (
        graph.chart.ChartTitle.Text
        == "Valor Total por Centro de Custo e Tipo de Serviço"
    )
    assert graph.chart.HasLegend is True
    assert graph.chart.Legend.Position == desktop.XL_LEGEND_BOTTOM
    assert graph.axis.TickLabels.NumberFormat == "R$ #,##0"
    graph.workbook.Save.assert_called_once_with()
    graph.workbook.Close.assert_called_once_with(SaveChanges=False)
    assert graph.excel.EnableEvents is False
    assert graph.excel.Calculation == desktop.XL_CALCULATION_MANUAL
    graph.excel.Quit.assert_called_once_with()
    pythoncom.CoUninitialize.assert_called_once_with()
    postprocess.assert_called_once_with(output_file)

    names = event_names(graph)
    assert names.index("workbook.save") < names.index("workbook.close")
    assert names.index("workbook.close") < names.index("excel.quit")
    assert names.index("excel.quit") < names.index("pythoncom.uninitialize")
    assert names.index("pythoncom.uninitialize") < names.index(
        "postprocess.calculation_on_open"
    )


def test_native_pivot_wraps_empty_static_source_and_cleans_up(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    graph = make_native_graph(source_last_row=3)
    pythoncom, _ = install_recording_com(monkeypatch, graph)
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: True)
    postprocess = Mock()
    monkeypatch.setattr(desktop, "_set_calculation_on_open", postprocess)

    with pytest.raises(
        ExcelDesktopError,
        match="não possui dados",
    ) as captured:
        desktop.create_native_pivot_and_recalculate(
            tmp_path / "resultado.xlsx",
            make_layout(),
            pivot_start_row=20,
            last_row=40,
            logger=Mock(spec=logging.Logger),
        )

    assert isinstance(captured.value.__cause__, ExcelDesktopError)
    graph.old_chart_2.Delete.assert_called_once_with()
    graph.old_chart_1.Delete.assert_called_once_with()
    graph.workbook.PivotCaches.assert_not_called()
    graph.workbook.Save.assert_not_called()
    graph.workbook.Close.assert_called_once_with(SaveChanges=False)
    graph.excel.Quit.assert_called_once_with()
    pythoncom.CoUninitialize.assert_called_once_with()
    postprocess.assert_not_called()


def test_native_pivot_wraps_dispatch_failure_and_only_uninitializes(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    graph = make_native_graph()
    pythoncom, _ = install_recording_com(
        monkeypatch,
        graph,
        dispatch_error=RuntimeError("falha dispatch"),
    )
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: True)
    postprocess = Mock()
    monkeypatch.setattr(desktop, "_set_calculation_on_open", postprocess)

    with pytest.raises(ExcelDesktopError, match="falha dispatch") as captured:
        desktop.create_native_pivot_and_recalculate(
            tmp_path / "resultado.xlsx",
            make_layout(),
            pivot_start_row=20,
            last_row=40,
            logger=Mock(spec=logging.Logger),
        )

    assert isinstance(captured.value.__cause__, RuntimeError)
    graph.workbooks.Open.assert_not_called()
    graph.workbook.Close.assert_not_called()
    graph.excel.Quit.assert_not_called()
    pythoncom.CoUninitialize.assert_called_once_with()
    postprocess.assert_not_called()


def test_native_pivot_wraps_open_failure_and_quits_excel(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    graph = make_native_graph(open_error=RuntimeError("falha open"))
    pythoncom, _ = install_recording_com(monkeypatch, graph)
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: True)
    postprocess = Mock()
    monkeypatch.setattr(desktop, "_set_calculation_on_open", postprocess)

    with pytest.raises(ExcelDesktopError, match="falha open") as captured:
        desktop.create_native_pivot_and_recalculate(
            tmp_path / "resultado.xlsx",
            make_layout(),
            pivot_start_row=20,
            last_row=40,
            logger=Mock(spec=logging.Logger),
        )

    assert isinstance(captured.value.__cause__, RuntimeError)
    graph.workbook.Close.assert_not_called()
    graph.excel.Quit.assert_called_once_with()
    pythoncom.CoUninitialize.assert_called_once_with()
    postprocess.assert_not_called()


def test_native_pivot_save_failure_does_not_postprocess(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    graph = make_native_graph(save_error=RuntimeError("falha save"))
    pythoncom, _ = install_recording_com(monkeypatch, graph)
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: True)
    postprocess = Mock()
    monkeypatch.setattr(desktop, "_set_calculation_on_open", postprocess)

    with pytest.raises(ExcelDesktopError, match="falha save") as captured:
        desktop.create_native_pivot_and_recalculate(
            tmp_path / "resultado.xlsx",
            make_layout(),
            pivot_start_row=20,
            last_row=40,
            logger=Mock(spec=logging.Logger),
        )

    assert isinstance(captured.value.__cause__, RuntimeError)
    graph.workbook.Close.assert_called_once_with(SaveChanges=False)
    graph.excel.Quit.assert_called_once_with()
    pythoncom.CoUninitialize.assert_called_once_with()
    postprocess.assert_not_called()


def test_native_pivot_tolerates_optional_formatting_failures(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    graph = make_native_graph(
        fail_initial_manual=True,
        fail_optional_pivot_formatting=True,
        fail_axes_formatting=True,
    )
    pythoncom, _ = install_recording_com(monkeypatch, graph)
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: True)
    postprocess = Mock()
    monkeypatch.setattr(desktop, "_set_calculation_on_open", postprocess)
    logger = Mock(spec=logging.Logger)

    result = desktop.create_native_pivot_and_recalculate(
        tmp_path / "resultado.xlsx",
        make_layout(),
        pivot_start_row=20,
        last_row=40,
        logger=logger,
    )

    assert result is False
    assert logger.debug.call_count == 3
    assert all(
        call.kwargs == {"exc_info": True} for call in logger.debug.call_args_list
    )
    graph.workbook.Save.assert_called_once_with()
    graph.workbook.Close.assert_called_once_with(SaveChanges=False)
    graph.excel.Quit.assert_called_once_with()
    pythoncom.CoUninitialize.assert_called_once_with()
    postprocess.assert_called_once()


def test_native_cleanup_continues_after_workbook_close_failure(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    graph = make_native_graph()
    graph.workbook.Close.side_effect = RuntimeError("falha ao fechar workbook")
    pythoncom, _ = install_recording_com(monkeypatch, graph)
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: True)
    postprocess = Mock()
    monkeypatch.setattr(desktop, "_set_calculation_on_open", postprocess)
    logger = Mock(spec=logging.Logger)

    with pytest.raises(
        ExcelDesktopCleanupError,
        match="fechar a pasta de trabalho",
    ):
        desktop.create_native_pivot_and_recalculate(
            tmp_path / "resultado.xlsx",
            make_layout(),
            pivot_start_row=20,
            last_row=40,
            logger=logger,
        )

    graph.workbook.Close.assert_called_once_with(SaveChanges=False)
    assert graph.excel.EnableEvents is False
    assert graph.excel.Calculation == desktop.XL_CALCULATION_MANUAL
    assert graph.excel.DisplayAlerts is False
    assert graph.excel.ScreenUpdating is False
    assert graph.excel.AskToUpdateLinks is False
    graph.excel.Quit.assert_called_once_with()
    pythoncom.CoUninitialize.assert_called_once_with()
    postprocess.assert_not_called()
    logger.warning.assert_called_once()


def test_native_cleanup_restores_the_original_excel_settings(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    graph = make_native_graph()
    original_settings = {
        "Visible": True,
        "DisplayAlerts": False,
        "ScreenUpdating": True,
        "AskToUpdateLinks": False,
        "EnableEvents": False,
        "Calculation": 777,
        "CalculateBeforeSave": True,
    }
    for property_name, value in original_settings.items():
        object.__setattr__(graph.excel, property_name, value)

    graph.excel.Quit.side_effect = RuntimeError("falha quit")
    install_recording_com(monkeypatch, graph)
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: True)
    postprocess = Mock()
    monkeypatch.setattr(desktop, "_set_calculation_on_open", postprocess)

    with pytest.raises(
        ExcelDesktopCleanupError,
        match="encerrar o Excel Desktop",
    ):
        desktop.create_native_pivot_and_recalculate(
            tmp_path / "resultado.xlsx",
            make_layout(),
            pivot_start_row=20,
            last_row=40,
            logger=Mock(spec=logging.Logger),
        )

    assert {
        property_name: getattr(graph.excel, property_name)
        for property_name in original_settings
    } == original_settings
    postprocess.assert_not_called()


def test_native_cleanup_catches_quit_attribute_lookup_failure(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    graph = make_native_graph(fail_quit_lookup=True)
    pythoncom, _ = install_recording_com(monkeypatch, graph)
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: True)
    postprocess = Mock()
    monkeypatch.setattr(desktop, "_set_calculation_on_open", postprocess)
    logger = Mock(spec=logging.Logger)

    with pytest.raises(
        ExcelDesktopCleanupError,
        match="encerrar o Excel Desktop",
    ):
        desktop.create_native_pivot_and_recalculate(
            tmp_path / "resultado.xlsx",
            make_layout(),
            pivot_start_row=20,
            last_row=40,
            logger=logger,
        )

    pythoncom.CoUninitialize.assert_called_once_with()
    postprocess.assert_not_called()
    assert any(
        "lookup de Quit indisponível" in str(call.args)
        for call in logger.warning.call_args_list
    )


def test_native_cleanup_uses_safe_restore_after_setting_read_failure(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    graph = make_native_graph(fail_read_property="DisplayAlerts")
    graph.excel.Quit.side_effect = RuntimeError("falha quit")
    pythoncom, _ = install_recording_com(monkeypatch, graph)
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: True)
    logger = Mock(spec=logging.Logger)

    with pytest.raises(
        ExcelDesktopCleanupError,
        match="encerrar o Excel Desktop",
    ):
        desktop.create_native_pivot_and_recalculate(
            tmp_path / "resultado.xlsx",
            make_layout(),
            pivot_start_row=20,
            last_row=40,
            logger=logger,
        )

    assert graph.excel.DisplayAlerts is True
    logger.debug.assert_any_call(
        "Não foi possível ler Excel.%s antes da automação.",
        "DisplayAlerts",
        exc_info=True,
    )
    pythoncom.CoUninitialize.assert_called_once_with()


def test_native_cleanup_preserves_primary_error_and_attempts_every_step(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    graph = make_native_graph(source_last_row=3)
    graph.workbook.Close.side_effect = RuntimeError("falha close")
    graph.excel.Quit.side_effect = RuntimeError("falha quit")
    pythoncom, _ = install_recording_com(monkeypatch, graph)
    pythoncom.CoUninitialize.side_effect = RuntimeError("falha uninitialize")
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: True)
    postprocess = Mock()
    monkeypatch.setattr(desktop, "_set_calculation_on_open", postprocess)
    logger = Mock(spec=logging.Logger)
    logger.warning.side_effect = RuntimeError("falha no logging")

    with pytest.raises(
        ExcelDesktopCleanupError,
        match="não possui dados",
    ) as captured:
        desktop.create_native_pivot_and_recalculate(
            tmp_path / "resultado.xlsx",
            make_layout(),
            pivot_start_row=20,
            last_row=40,
            logger=logger,
        )

    assert isinstance(captured.value.__cause__, ExcelDesktopError)
    assert "não possui dados" in str(captured.value.__cause__)
    assert any("falha close" in note for note in captured.value.__notes__)
    assert any("falha quit" in note for note in captured.value.__notes__)
    assert any("falha uninitialize" in note for note in captured.value.__notes__)
    graph.workbook.Close.assert_called_once_with(SaveChanges=False)
    graph.excel.Quit.assert_called_once_with()
    pythoncom.CoUninitialize.assert_called_once_with()
    postprocess.assert_not_called()
    assert logger.warning.call_count == 3


def test_native_cleanup_tolerates_independent_restore_and_uninitialize_failures(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    graph = make_native_graph(fail_restore=True)
    graph.excel.Quit.side_effect = FakeComError("falha quit")
    pythoncom, _ = install_recording_com(monkeypatch, graph)
    pythoncom.CoUninitialize.side_effect = FakeComError("falha uninitialize")
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: True)
    postprocess = Mock()
    monkeypatch.setattr(desktop, "_set_calculation_on_open", postprocess)
    logger = Mock(spec=logging.Logger)

    with pytest.raises(
        ExcelDesktopCleanupError,
        match="encerrar o Excel Desktop",
    ):
        desktop.create_native_pivot_and_recalculate(
            tmp_path / "resultado.xlsx",
            make_layout(),
            pivot_start_row=20,
            last_row=40,
            logger=logger,
        )

    assert graph.excel.EnableEvents is False
    assert graph.excel.Calculation == desktop.XL_CALCULATION_AUTOMATIC
    assert graph.excel.DisplayAlerts is True
    assert graph.excel.ScreenUpdating is True
    assert graph.excel.AskToUpdateLinks is True
    graph.excel.Quit.assert_called_once_with()
    pythoncom.CoUninitialize.assert_called_once_with()
    postprocess.assert_not_called()
    assert logger.warning.call_count == 3


def test_native_coinitialize_failure_does_not_uninitialize_com(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    graph = make_native_graph()
    pythoncom, client = install_recording_com(monkeypatch, graph)
    pythoncom.CoInitialize.side_effect = RuntimeError("falha CoInitialize")
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: True)

    with pytest.raises(ExcelDesktopError, match="falha CoInitialize"):
        desktop.create_native_pivot_and_recalculate(
            tmp_path / "resultado.xlsx",
            make_layout(),
            pivot_start_row=20,
            last_row=40,
            logger=Mock(spec=logging.Logger),
        )

    client.DispatchEx.assert_not_called()
    pythoncom.CoUninitialize.assert_not_called()


def test_recalculate_only_timeout_closes_without_saving(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    workbook = Mock()
    workbooks = Mock()
    workbooks.Open.return_value = workbook
    excel = Mock()
    excel.Workbooks = workbooks
    pythoncom, _ = install_fake_com_modules(monkeypatch, excel)
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: True)
    monkeypatch.setattr(
        desktop,
        "_wait_for_calculation",
        Mock(side_effect=ExcelDesktopError("tempo excedido")),
    )

    with pytest.raises(ExcelDesktopError, match="tempo excedido"):
        desktop.recalculate_only(
            tmp_path / "resultado.xlsx",
            Mock(spec=logging.Logger),
        )

    excel.CalculateFullRebuild.assert_called_once_with()
    workbook.Save.assert_not_called()
    workbook.Close.assert_called_once_with(SaveChanges=False)
    excel.Quit.assert_called_once_with()
    pythoncom.CoUninitialize.assert_called_once_with()


def test_recalculate_cleanup_failure_does_not_skip_quit_or_uninitialize(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    workbook = Mock()
    workbook.Close.side_effect = RuntimeError("falha ao fechar workbook")
    workbooks = Mock()
    workbooks.Open.return_value = workbook
    excel = Mock()
    excel.Workbooks = workbooks
    pythoncom, _ = install_fake_com_modules(monkeypatch, excel)
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: True)
    monkeypatch.setattr(desktop, "_wait_for_calculation", Mock())
    logger = Mock(spec=logging.Logger)

    with pytest.raises(
        ExcelDesktopCleanupError,
        match="fechar a pasta de trabalho",
    ):
        desktop.recalculate_only(
            tmp_path / "resultado.xlsx",
            logger,
        )

    workbook.Save.assert_called_once_with()
    workbook.Close.assert_called_once_with(SaveChanges=True)
    excel.Quit.assert_called_once_with()
    pythoncom.CoUninitialize.assert_called_once_with()
    logger.warning.assert_called_once()


def test_recalculate_cleanup_preserves_timeout_as_primary_error(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    workbook = Mock()
    workbook.Close.side_effect = RuntimeError("falha close")
    workbooks = Mock()
    workbooks.Open.return_value = workbook
    excel = Mock()
    excel.Workbooks = workbooks
    excel.Quit.side_effect = RuntimeError("falha quit")
    pythoncom, _ = install_fake_com_modules(monkeypatch, excel)
    pythoncom.CoUninitialize.side_effect = RuntimeError("falha uninitialize")
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: True)
    timeout_error = ExcelDesktopError("tempo excedido")
    monkeypatch.setattr(
        desktop,
        "_wait_for_calculation",
        Mock(side_effect=timeout_error),
    )
    logger = Mock(spec=logging.Logger)

    with pytest.raises(
        ExcelDesktopCleanupError,
        match="tempo excedido",
    ) as captured:
        desktop.recalculate_only(
            tmp_path / "resultado.xlsx",
            logger,
        )

    assert captured.value.__cause__ is timeout_error
    assert any("falha close" in note for note in captured.value.__notes__)
    assert any("falha quit" in note for note in captured.value.__notes__)
    assert any("falha uninitialize" in note for note in captured.value.__notes__)
    workbook.Save.assert_not_called()
    workbook.Close.assert_called_once_with(SaveChanges=False)
    excel.Quit.assert_called_once_with()
    pythoncom.CoUninitialize.assert_called_once_with()
    assert logger.warning.call_count == 3


def test_recalculate_only_open_failure_quits_and_uninitializes(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    workbooks = Mock()
    workbooks.Open.side_effect = RuntimeError("falha ao abrir")
    excel = Mock()
    excel.Workbooks = workbooks
    pythoncom, _ = install_fake_com_modules(monkeypatch, excel)
    monkeypatch.setattr(desktop, "pywin32_is_available", lambda: True)

    with pytest.raises(RuntimeError, match="falha ao abrir"):
        desktop.recalculate_only(
            tmp_path / "resultado.xlsx",
            Mock(spec=logging.Logger),
        )

    excel.Quit.assert_called_once_with()
    pythoncom.CoUninitialize.assert_called_once_with()
