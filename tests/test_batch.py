"""Testes do processamento em lote."""

from __future__ import annotations

from pathlib import Path

import pytest

import src.core.batch as batch_module
from src.core.exceptions import AutomationError, ExcelDesktopCleanupError
from src.core.models import RunResult


def make_result(input_file: Path) -> RunResult:
    return RunResult(
        output_file=input_file.with_name(f"saida_{input_file.name}"),
        backup_file=input_file.with_name(f"backup_{input_file.name}"),
        log_file=input_file.with_suffix(".log"),
        native_pivot_created=False,
        detected_sheets={},
        detected_columns={},
        checks={"travel_rows": 1, "formula_errors": 0},
    )


def test_batch_processes_all_inputs_and_preserves_order(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    inputs = (tmp_path / "a.xlsx", tmp_path / "b.xlsm")
    calls: list[dict[str, object]] = []

    class FakeEngine:
        def run(self, **kwargs: object) -> RunResult:
            calls.append(kwargs)
            return make_result(kwargs["input_value"])  # type: ignore[arg-type]

    monkeypatch.setattr(batch_module, "list_input_files", lambda: inputs)

    result = batch_module.BatchAutomation(FakeEngine()).run(
        candidate_name="Carlos Eduardo",
        use_native_pivot=False,
        verbose=True,
    )

    assert result.succeeded == 2
    assert result.failed == 0
    assert [item.input_file for item in result.items] == list(inputs)
    assert [call["output_label"] for call in calls] == ["a", "b"]
    assert all(call["candidate_name"] == "Carlos Eduardo" for call in calls)
    assert all(call["use_native_pivot"] is False for call in calls)
    assert all(call["verbose"] is True for call in calls)


def test_batch_continues_after_expected_individual_failure(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    inputs = tuple(tmp_path / name for name in ("a.xlsx", "b.xlsx", "c.xlsx"))

    class FakeEngine:
        def run(self, **kwargs: object) -> RunResult:
            input_file = kwargs["input_value"]
            if input_file == inputs[1]:
                raise AutomationError("estrutura inválida")
            return make_result(input_file)  # type: ignore[arg-type]

    monkeypatch.setattr(batch_module, "list_input_files", lambda: inputs)

    result = batch_module.BatchAutomation(FakeEngine()).run(
        candidate_name="Carlos Eduardo"
    )

    assert result.succeeded == 2
    assert result.failed == 1
    assert result.items[1].result is None
    assert result.items[1].error == "estrutura inválida"
    assert result.items[2].succeeded is True


def test_batch_aborts_after_critical_excel_cleanup_failure(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    inputs = tuple(tmp_path / name for name in ("a.xlsx", "b.xlsx", "c.xlsx"))
    processed: list[Path] = []

    class FakeEngine:
        def run(self, **kwargs: object) -> RunResult:
            input_file = kwargs["input_value"]
            processed.append(input_file)  # type: ignore[arg-type]
            if input_file == inputs[1]:
                raise ExcelDesktopCleanupError("arquivo bloqueado")
            return make_result(input_file)  # type: ignore[arg-type]

    monkeypatch.setattr(batch_module, "list_input_files", lambda: inputs)

    with pytest.raises(ExcelDesktopCleanupError, match="arquivo bloqueado"):
        batch_module.BatchAutomation(FakeEngine()).run(candidate_name="Carlos Eduardo")

    assert processed == list(inputs[:2])


def test_batch_validates_candidate_before_listing_inputs(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def forbidden_listing() -> tuple[Path, ...]:
        raise AssertionError("não deve listar")

    monkeypatch.setattr(batch_module, "list_input_files", forbidden_listing)

    with pytest.raises(AutomationError, match="nome completo"):
        batch_module.BatchAutomation().run(candidate_name="Carlos")


def test_batch_forwards_adapter_to_every_input(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    inputs = (tmp_path / "a.xlsx", tmp_path / "b.xlsx")
    calls: list[dict[str, object]] = []

    class FakeEngine:
        def run(self, **kwargs: object) -> RunResult:
            calls.append(kwargs)
            return make_result(kwargs["input_value"])  # type: ignore[arg-type]

    monkeypatch.setattr(batch_module, "list_input_files", lambda: inputs)

    batch_module.BatchAutomation(FakeEngine()).run(
        candidate_name="Carlos Eduardo",
        adapter=Path("config/cliente.json"),
    )

    assert [call["adapter"] for call in calls] == [
        Path("config/cliente.json"),
        Path("config/cliente.json"),
    ]
