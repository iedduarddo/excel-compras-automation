"""Testes da interface de linha de comando."""

from __future__ import annotations

from pathlib import Path

import pytest

import src.main as main_module
from src.core.exceptions import AutomationError
from src.core.models import RunResult


def make_result(tmp_path: Path, *, native: bool = False) -> RunResult:
    return RunResult(
        output_file=tmp_path / "resultado.xlsx",
        backup_file=tmp_path / "backup.xlsx",
        log_file=tmp_path / "execucao.log",
        native_pivot_created=native,
        detected_sheets={},
        detected_columns={},
        checks={"travel_rows": 40, "formula_errors": 0},
    )


def test_build_parser_maps_all_command_line_options() -> None:
    args = main_module.build_parser().parse_args(
        [
            "--input",
            "entrada.xlsx",
            "--nome",
            "Carlos Eduardo",
            "--sem-pivot-nativo",
            "--verbose",
        ]
    )

    assert args.input == Path("entrada.xlsx")
    assert args.candidate_name == "Carlos Eduardo"
    assert args.sem_pivot_nativo is True
    assert args.verbose is True


def test_main_runs_engine_and_prints_native_result(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
    tmp_path: Path,
) -> None:
    received: dict[str, object] = {}

    class FakeEngine:
        def run(self, **kwargs: object) -> RunResult:
            received.update(kwargs)
            return make_result(tmp_path, native=True)

    monkeypatch.setattr(main_module, "AutomationEngine", FakeEngine)

    exit_code = main_module.main(
        [
            "--input",
            "entrada.xlsx",
            "--candidate-name",
            "Carlos Eduardo",
            "--verbose",
        ]
    )

    output = capsys.readouterr().out
    assert exit_code == 0
    assert received == {
        "input_value": Path("entrada.xlsx"),
        "candidate_name": "Carlos Eduardo",
        "use_native_pivot": True,
        "verbose": True,
    }
    assert "AUTOMAÇÃO CONCLUÍDA COM SUCESSO" in output
    assert "nativa do Excel" in output
    assert "40 solicitações" in output


def test_main_prompts_for_name_and_reports_fallback(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
    tmp_path: Path,
) -> None:
    received: dict[str, object] = {}

    class FakeEngine:
        def run(self, **kwargs: object) -> RunResult:
            received.update(kwargs)
            return make_result(tmp_path)

    monkeypatch.setattr(main_module, "AutomationEngine", FakeEngine)
    monkeypatch.setattr("builtins.input", lambda _: "  Carlos Eduardo  ")

    exit_code = main_module.main(["--sem-pivot-nativo"])

    output = capsys.readouterr().out
    assert exit_code == 0
    assert received["candidate_name"] == "Carlos Eduardo"
    assert received["use_native_pivot"] is False
    assert "resumo compatível com fórmulas" in output


def test_main_translates_expected_error_to_exit_code(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    class FailingEngine:
        def run(self, **_: object) -> RunResult:
            raise AutomationError("Planilha inválida")

    monkeypatch.setattr(main_module, "AutomationEngine", FailingEngine)

    exit_code = main_module.main(["--nome", "Carlos Eduardo"])

    captured = capsys.readouterr()
    assert exit_code == 1
    assert "automação não foi concluída" in captured.err
    assert "Planilha inválida" in captured.err


def test_main_handles_keyboard_interrupt_during_prompt(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    def interrupt(_: str) -> str:
        raise KeyboardInterrupt

    monkeypatch.setattr("builtins.input", interrupt)

    exit_code = main_module.main([])

    assert exit_code == 130
    assert "cancelada pelo usuário" in capsys.readouterr().err
