"""Testes do diagnóstico read-only do ambiente."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import Mock

import pytest

import src.services.diagnostics as diagnostics
import src.services.files as files_service
from src.core.exceptions import AutomationError, DetectionError
from src.core.models import SheetLayout, WorkbookLayout
from src.services.diagnostics import (
    DiagnosticItem,
    DiagnosticReport,
    DiagnosticStatus,
)


def valid_aliases() -> dict[str, object]:
    return {key: {} for key in diagnostics.REQUIRED_ALIAS_KEYS}


def valid_rules() -> dict[str, object]:
    return {key: {} for key in diagnostics.REQUIRED_RULE_KEYS}


def detected_layout() -> WorkbookLayout:
    return WorkbookLayout(
        base=SheetLayout("Base", 1, {"request_id": 1}),
        policies=SheetLayout("Políticas", 1, {"service_type": 1}),
        responses=SheetLayout("Respostas", 1, {"indicator": 1, "answer": 2}),
    )


@pytest.mark.parametrize(
    ("version", "expected"),
    [
        ((3, 10), DiagnosticStatus.ERROR),
        ((3, 11), DiagnosticStatus.OK),
        ((3, 14), DiagnosticStatus.OK),
        ((3, 15), DiagnosticStatus.ERROR),
    ],
)
def test_python_version_matches_supported_range(
    version: tuple[int, int],
    expected: DiagnosticStatus,
) -> None:
    assert diagnostics.check_python_version(version).status is expected


def test_report_allows_warning_and_formats_summary() -> None:
    ready_report = DiagnosticReport(
        (
            DiagnosticItem("Python", DiagnosticStatus.OK, "suportado"),
            DiagnosticItem("Excel Desktop", DiagnosticStatus.WARNING, "fallback"),
        )
    )
    failing_report = DiagnosticReport(
        (DiagnosticItem("Aliases", DiagnosticStatus.ERROR, "ausente"),)
    )

    ready_output = diagnostics.format_diagnostic_report(ready_report)
    failing_output = diagnostics.format_diagnostic_report(failing_report)

    assert ready_report.ready is True
    assert ready_report.exit_code == 0
    assert "[OK] Python: suportado" in ready_output
    assert "[AVISO] Excel Desktop: fallback" in ready_output
    assert "AMBIENTE PRONTO" in ready_output
    assert failing_report.ready is False
    assert failing_report.exit_code == 1
    assert "[ERRO] Aliases: ausente" in failing_output
    assert "AMBIENTE REQUER ATENÇÃO" in failing_output


def test_diagnostics_is_read_only_and_closes_workbook(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    input_file = tmp_path / "entrada.xlsx"
    input_file.write_bytes(b"conteudo original")
    original_content = input_file.read_bytes()
    workbook = Mock()
    detector = Mock()
    detector.detect.return_value = detected_layout()
    prepare_paths = Mock(side_effect=AssertionError("não deve criar backup"))

    monkeypatch.setattr(diagnostics, "load_aliases", valid_aliases)
    monkeypatch.setattr(diagnostics, "load_rules", valid_rules)
    monkeypatch.setattr(diagnostics, "resolve_input_file", lambda _: input_file)
    monkeypatch.setattr(
        diagnostics, "load_source_workbook", Mock(return_value=workbook)
    )
    monkeypatch.setattr(
        diagnostics,
        "WorkbookDetector",
        Mock(return_value=detector),
    )
    monkeypatch.setattr(
        diagnostics,
        "_check_native_excel",
        lambda: DiagnosticItem(
            "Excel Desktop",
            DiagnosticStatus.WARNING,
            "fallback disponível",
        ),
    )
    monkeypatch.setattr(files_service, "prepare_run_paths", prepare_paths)

    report = diagnostics.run_diagnostics(input_file)

    assert report.exit_code == 0
    assert input_file.read_bytes() == original_content
    workbook.save.assert_not_called()
    workbook.close.assert_called_once_with()
    prepare_paths.assert_not_called()
    assert any(
        item.name == "Estrutura da planilha" and item.status is DiagnosticStatus.OK
        for item in report.items
    )


def test_detection_error_is_reported_and_workbook_is_closed(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    workbook = Mock()
    detector = Mock()
    detector.detect.side_effect = DetectionError("abas inválidas")

    monkeypatch.setattr(diagnostics, "load_aliases", valid_aliases)
    monkeypatch.setattr(diagnostics, "load_rules", valid_rules)
    monkeypatch.setattr(
        diagnostics,
        "resolve_input_file",
        lambda _: tmp_path / "entrada.xlsx",
    )
    monkeypatch.setattr(
        diagnostics, "load_source_workbook", Mock(return_value=workbook)
    )
    monkeypatch.setattr(
        diagnostics,
        "WorkbookDetector",
        Mock(return_value=detector),
    )
    monkeypatch.setattr(
        diagnostics,
        "_check_native_excel",
        lambda: DiagnosticItem("Excel Desktop", DiagnosticStatus.WARNING, "fallback"),
    )

    report = diagnostics.run_diagnostics(None)

    workbook.close.assert_called_once_with()
    assert report.exit_code == 1
    assert any(
        item.status is DiagnosticStatus.ERROR and "abas inválidas" in item.message
        for item in report.items
    )


def test_missing_configuration_keys_are_required() -> None:
    item, value = diagnostics._check_configuration(
        name="Aliases",
        loader=lambda: {"sheets": {}},
        required_keys=diagnostics.REQUIRED_ALIAS_KEYS,
    )

    assert item.status is DiagnosticStatus.ERROR
    assert "base_columns" in item.message
    assert value is None


def test_input_error_returns_failure_without_opening_workbook(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    load_workbook = Mock()
    monkeypatch.setattr(diagnostics, "load_aliases", valid_aliases)
    monkeypatch.setattr(diagnostics, "load_rules", valid_rules)
    monkeypatch.setattr(
        diagnostics,
        "resolve_input_file",
        Mock(side_effect=AutomationError("entrada ausente")),
    )
    monkeypatch.setattr(diagnostics, "load_source_workbook", load_workbook)
    monkeypatch.setattr(
        diagnostics,
        "_check_native_excel",
        lambda: DiagnosticItem("Excel Desktop", DiagnosticStatus.WARNING, "fallback"),
    )

    report = diagnostics.run_diagnostics(None)

    assert report.exit_code == 1
    load_workbook.assert_not_called()
    assert any("entrada ausente" in item.message for item in report.items)


def test_native_excel_unavailable_is_only_a_warning(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(diagnostics, "pywin32_is_available", lambda: False)

    item = diagnostics._check_native_excel()

    assert item.status is DiagnosticStatus.WARNING
    assert "fallback" in item.message


def test_native_excel_available_is_ok(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(diagnostics, "pywin32_is_available", lambda: True)

    item = diagnostics._check_native_excel()

    assert item.status is DiagnosticStatus.OK
    assert "será validado na execução" in item.message
