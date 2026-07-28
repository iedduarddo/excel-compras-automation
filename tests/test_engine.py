"""Testes do orquestrador sem acesso real a arquivos ou ao Excel Desktop."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from unittest.mock import Mock

import pytest

import src.core.engine as engine_module
from src.core.exceptions import ExcelDesktopError
from src.core.models import (
    Policy,
    RunPaths,
    SheetLayout,
    TravelResult,
    WorkbookLayout,
)


@dataclass
class EngineHarness:
    paths: RunPaths
    layout: WorkbookLayout
    travels: list[TravelResult]
    policies: dict[str, Policy]
    primary_workbook: Mock
    fallback_workbook: Mock | None
    logger: Mock
    load_workbook: Mock
    native_create: Mock
    recalculate: Mock
    fallback_create: Mock
    validate: Mock
    pywin_available: Mock


def make_layout() -> WorkbookLayout:
    return WorkbookLayout(
        base=SheetLayout(
            "Base",
            1,
            {
                "request_id": 1,
                "total_value": 2,
                "priority": 3,
            },
        ),
        policies=SheetLayout(
            "Policies",
            1,
            {
                "service_type": 1,
                "limit_value": 2,
                "min_lead_days": 3,
            },
        ),
        responses=SheetLayout("Responses", 1, {"indicator": 1, "answer": 2}),
    )


def make_travel() -> TravelResult:
    return TravelResult(
        worksheet_row=2,
        request_id="VIA-001",
        request_date=date(2026, 1, 1),
        travel_date=date(2026, 1, 10),
        service_type="Aéreo",
        supplier="Fornecedor",
        cost_center="CC1001",
        booking_status="Confirmada",
        criticality="Normal",
        card_status="Conferido",
        total_value=500.0,
        lead_days=9,
        policy_limit=1000.0,
        min_lead_days=5,
        policy_status="OK",
        limit_difference=-500.0,
    )


def install_harness(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    *,
    native_available: bool,
    needs_fallback: bool,
) -> EngineHarness:
    paths = RunPaths(
        input_file=tmp_path / "entrada.xlsx",
        backup_file=tmp_path / "backup.xlsx",
        output_file=tmp_path / "resultado.xlsx",
        log_file=tmp_path / "execucao.log",
    )
    layout = make_layout()
    travels = [make_travel()]
    policies = {
        "aereo": Policy(
            service_type="Aéreo",
            limit_value=1000.0,
            min_lead_days=5,
            worksheet_row=2,
        )
    }
    primary = Mock(name="primary_workbook")
    fallback = Mock(name="fallback_workbook") if needs_fallback else None
    load_workbook = Mock(
        side_effect=[primary, fallback] if fallback is not None else [primary]
    )
    logger = Mock(spec=logging.Logger)
    detector = Mock()
    detector.detect.return_value = layout
    native_create = Mock(return_value=True)
    recalculate = Mock()
    fallback_create = Mock()
    validate = Mock(
        return_value={
            "travel_rows": 1,
            "formula_errors": 0,
            "chart_parts": 1,
        }
    )
    pywin_available = Mock(return_value=native_available)

    monkeypatch.setattr(
        engine_module, "resolve_input_file", Mock(return_value=paths.input_file)
    )
    monkeypatch.setattr(
        engine_module,
        "validate_candidate_name",
        Mock(return_value="Carlos Eduardo"),
    )
    monkeypatch.setattr(engine_module, "prepare_run_paths", Mock(return_value=paths))
    monkeypatch.setattr(engine_module, "configure_logging", Mock(return_value=logger))
    monkeypatch.setattr(
        engine_module, "load_aliases", Mock(return_value={"aliases": {}})
    )
    monkeypatch.setattr(
        engine_module,
        "load_rules",
        Mock(return_value={"top_requests": 5}),
    )
    monkeypatch.setattr(
        engine_module,
        "WorkbookDetector",
        Mock(return_value=detector),
    )
    monkeypatch.setattr(engine_module, "load_source_workbook", load_workbook)
    monkeypatch.setattr(engine_module, "read_policies", Mock(return_value=policies))
    monkeypatch.setattr(engine_module, "find_last_data_row", Mock(return_value=2))
    monkeypatch.setattr(engine_module, "analyze_travels", Mock(return_value=travels))
    monkeypatch.setattr(
        engine_module, "apply_priority_scores", Mock(return_value=travels)
    )
    monkeypatch.setattr(
        engine_module,
        "ensure_derived_columns",
        Mock(return_value=layout),
    )
    monkeypatch.setattr(
        engine_module,
        "create_support_sheet",
        Mock(return_value={"support": "Apoio"}),
    )
    monkeypatch.setattr(engine_module, "write_base_formulas", Mock())
    monkeypatch.setattr(engine_module, "write_responses", Mock(return_value=20))
    monkeypatch.setattr(engine_module, "apply_conditional_formatting", Mock())
    monkeypatch.setattr(engine_module, "finalize_workbook", Mock())
    monkeypatch.setattr(engine_module, "pywin32_is_available", pywin_available)
    monkeypatch.setattr(
        engine_module,
        "create_native_pivot_and_recalculate",
        native_create,
    )
    monkeypatch.setattr(engine_module, "recalculate_only", recalculate)
    monkeypatch.setattr(
        engine_module,
        "create_fallback_summary_and_chart",
        fallback_create,
    )
    monkeypatch.setattr(engine_module, "validate_output", validate)

    return EngineHarness(
        paths=paths,
        layout=layout,
        travels=travels,
        policies=policies,
        primary_workbook=primary,
        fallback_workbook=fallback,
        logger=logger,
        load_workbook=load_workbook,
        native_create=native_create,
        recalculate=recalculate,
        fallback_create=fallback_create,
        validate=validate,
        pywin_available=pywin_available,
    )


def test_engine_creates_fallback_when_excel_is_unavailable(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    harness = install_harness(
        monkeypatch,
        tmp_path,
        native_available=False,
        needs_fallback=True,
    )

    result = engine_module.AutomationEngine().run(
        input_value=None,
        candidate_name="Carlos Eduardo",
        use_native_pivot=True,
    )

    assert result.output_file == harness.paths.output_file
    assert result.native_pivot_created is False
    assert result.detected_sheets == {
        "base": "Base",
        "policies": "Policies",
        "responses": "Responses",
    }
    harness.pywin_available.assert_called_once_with()
    harness.native_create.assert_not_called()
    harness.fallback_create.assert_called_once_with(
        harness.fallback_workbook,
        harness.layout,
        harness.travels,
        20,
        2,
    )
    harness.primary_workbook.save.assert_called_once_with(harness.paths.output_file)
    harness.primary_workbook.close.assert_called_once_with()
    assert harness.fallback_workbook is not None
    harness.fallback_workbook.save.assert_called_once_with(harness.paths.output_file)
    harness.fallback_workbook.close.assert_called_once_with()
    harness.validate.assert_called_once()
    validation_kwargs = harness.validate.call_args.kwargs
    assert validation_kwargs["native_pivot_expected"] is False
    assert validation_kwargs["cached_values_expected"] is False


def test_engine_uses_native_pivot_without_fallback(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    harness = install_harness(
        monkeypatch,
        tmp_path,
        native_available=True,
        needs_fallback=False,
    )

    result = engine_module.AutomationEngine().run(
        input_value=harness.paths.input_file,
        candidate_name="Carlos Eduardo",
        use_native_pivot=True,
        verbose=True,
    )

    assert result.native_pivot_created is True
    harness.native_create.assert_called_once_with(
        harness.paths.output_file,
        harness.layout,
        20,
        2,
        harness.logger,
    )
    harness.fallback_create.assert_not_called()
    harness.recalculate.assert_not_called()
    validation_kwargs = harness.validate.call_args.kwargs
    assert validation_kwargs["native_pivot_expected"] is True
    assert validation_kwargs["cached_values_expected"] is True


def test_engine_falls_back_after_excel_desktop_error(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    harness = install_harness(
        monkeypatch,
        tmp_path,
        native_available=True,
        needs_fallback=True,
    )
    harness.native_create.side_effect = ExcelDesktopError("Falha controlada")

    result = engine_module.AutomationEngine().run(
        input_value=None,
        candidate_name="Carlos Eduardo",
    )

    assert result.native_pivot_created is False
    harness.fallback_create.assert_called_once()
    harness.recalculate.assert_not_called()
    warning_messages = [call.args[0] for call in harness.logger.warning.call_args_list]
    assert "%s" in warning_messages
    assert any("compatibilidade" in message for message in warning_messages)


def test_engine_closes_primary_workbook_when_writer_fails(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    harness = install_harness(
        monkeypatch,
        tmp_path,
        native_available=False,
        needs_fallback=False,
    )
    engine_module.ensure_derived_columns.side_effect = RuntimeError("writer falhou")

    with pytest.raises(RuntimeError, match="writer falhou"):
        engine_module.AutomationEngine().run(
            input_value=None,
            candidate_name="Carlos Eduardo",
        )

    harness.primary_workbook.close.assert_called_once_with()
    harness.primary_workbook.save.assert_not_called()


def test_engine_closes_fallback_workbook_when_summary_fails(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    harness = install_harness(
        monkeypatch,
        tmp_path,
        native_available=False,
        needs_fallback=True,
    )
    harness.fallback_create.side_effect = RuntimeError("resumo falhou")

    with pytest.raises(RuntimeError, match="resumo falhou"):
        engine_module.AutomationEngine().run(
            input_value=None,
            candidate_name="Carlos Eduardo",
        )

    harness.primary_workbook.close.assert_called_once_with()
    assert harness.fallback_workbook is not None
    harness.fallback_workbook.close.assert_called_once_with()
    harness.fallback_workbook.save.assert_not_called()
