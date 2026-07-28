"""Testes seguros da integração opcional, sem iniciar o Excel Desktop real."""

from __future__ import annotations

import builtins
import logging
import sys
from pathlib import Path
from types import ModuleType
from unittest.mock import Mock
from zipfile import ZipFile

import pytest

import src.excel.excel_desktop as desktop
from src.core.exceptions import ExcelDesktopError


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
