"""Testes da interface de linha de comando."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import Mock

import pytest

import src.main as main_module
from src import __version__
from src.core.batch import BatchItemResult, BatchResult
from src.core.exceptions import AutomationError
from src.core.models import RunResult
from src.services.diagnostics import (
    DiagnosticItem,
    DiagnosticReport,
    DiagnosticStatus,
)


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
            "--adaptador",
            "config/cliente.json",
            "--sem-pivot-nativo",
            "--verbose",
            "--diagnostic",
        ]
    )

    assert args.input == Path("entrada.xlsx")
    assert args.candidate_name == "Carlos Eduardo"
    assert args.adapter == Path("config/cliente.json")
    assert args.sem_pivot_nativo is True
    assert args.verbose is True
    assert args.diagnostic is True

    batch_args = main_module.build_parser().parse_args(["--lote"])
    assert batch_args.batch is True
    assert batch_args.input is None

    assistant_args = main_module.build_parser().parse_args(
        ["--assistente", "--comando", "ajuda", "--pasta-assistente", "central"]
    )
    assert assistant_args.assistant is True
    assert assistant_args.command == ["ajuda"]
    assert assistant_args.assistant_root == Path("central")


def test_parser_rejects_batch_with_explicit_input() -> None:
    with pytest.raises(SystemExit) as error:
        main_module.build_parser().parse_args(["--lote", "--input", "entrada.xlsx"])

    assert error.value.code == 2


def test_main_prepares_assistant_workspace(
    capsys: pytest.CaptureFixture[str],
    tmp_path: Path,
) -> None:
    root = tmp_path / "central"

    exit_code = main_module.main(
        [
            "--assistente",
            "--preparar-pastas",
            "--pasta-assistente",
            str(root),
        ]
    )

    assert exit_code == 0
    assert (root / "entrada").is_dir()
    assert (root / "comandos" / "pendentes").is_dir()
    assert "Pastas do assistente preparadas" in capsys.readouterr().out


def test_main_executes_assistant_help_command(
    capsys: pytest.CaptureFixture[str],
    tmp_path: Path,
) -> None:
    exit_code = main_module.main(
        [
            "--assistente",
            "--comando",
            "ajuda",
            "--pasta-assistente",
            str(tmp_path / "central"),
        ]
    )

    assert exit_code == 0
    assert "Comandos disponíveis" in capsys.readouterr().out


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

    exit_code = main_module.main(
        ["--sem-pivot-nativo", "--adaptador", "config/cliente.json"]
    )

    output = capsys.readouterr().out
    assert exit_code == 0
    assert received["candidate_name"] == "Carlos Eduardo"
    assert received["use_native_pivot"] is False
    assert received["adapter"] == Path("config/cliente.json")
    assert "resumo compatível com fórmulas" in output


def test_main_runs_batch_and_returns_one_when_one_input_fails(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
    tmp_path: Path,
) -> None:
    successful_input = tmp_path / "ok.xlsx"
    failed_input = tmp_path / "falha.xlsx"
    batch_result = BatchResult(
        (
            BatchItemResult(successful_input, result=make_result(tmp_path)),
            BatchItemResult(failed_input, error="estrutura inválida"),
        )
    )
    received: dict[str, object] = {}

    class FakeBatch:
        def run(self, **kwargs: object) -> BatchResult:
            received.update(kwargs)
            return batch_result

    monkeypatch.setattr(main_module, "BatchAutomation", FakeBatch)

    exit_code = main_module.main(
        ["--lote", "--nome", "Carlos Eduardo", "--sem-pivot-nativo"]
    )

    output = capsys.readouterr().out
    assert exit_code == 1
    assert received == {
        "candidate_name": "Carlos Eduardo",
        "use_native_pivot": False,
        "verbose": False,
    }
    assert "PROCESSAMENTO EM LOTE CONCLUÍDO" in output
    assert "Sucessos      : 1" in output
    assert "Falhas        : 1" in output
    assert "[OK] ok.xlsx" in output
    assert "[FALHA] falha.xlsx" in output
    assert "estrutura inválida" in output


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


@pytest.mark.parametrize("option", ["--diagnostico", "--diagnostic"])
def test_main_runs_diagnostic_without_prompt_or_engine(
    option: str,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    received: dict[str, object] = {}
    prompt = Mock(side_effect=AssertionError("não deve solicitar nome"))

    class ForbiddenEngine:
        def __init__(self) -> None:
            raise AssertionError("não deve iniciar o engine")

    report = DiagnosticReport(
        (
            DiagnosticItem("Python", DiagnosticStatus.OK, "suportado"),
            DiagnosticItem("Excel Desktop", DiagnosticStatus.WARNING, "fallback"),
        )
    )

    def diagnose(input_value: object) -> DiagnosticReport:
        received["input"] = input_value
        return report

    monkeypatch.setattr("builtins.input", prompt)
    monkeypatch.setattr(main_module, "AutomationEngine", ForbiddenEngine)
    monkeypatch.setattr(main_module, "run_diagnostics", diagnose)

    exit_code = main_module.main([option, "--input", "entrada.xlsx"])

    assert exit_code == 0
    assert received == {"input": Path("entrada.xlsx")}
    prompt.assert_not_called()
    output = capsys.readouterr().out
    assert "[OK] Python" in output
    assert "[AVISO] Excel Desktop" in output
    assert "AMBIENTE PRONTO" in output


def test_main_returns_one_when_diagnostic_has_error(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    report = DiagnosticReport(
        (DiagnosticItem("Entrada", DiagnosticStatus.ERROR, "ausente"),)
    )
    monkeypatch.setattr(main_module, "run_diagnostics", lambda _: report)

    exit_code = main_module.main(["--diagnostico"])

    assert exit_code == 1
    assert "AMBIENTE REQUER ATENÇÃO" in capsys.readouterr().out


def test_version_uses_package_metadata_without_prompt_or_engine(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    prompt = Mock(side_effect=AssertionError("não deve solicitar nome"))

    class ForbiddenEngine:
        def __init__(self) -> None:
            raise AssertionError("não deve iniciar o engine")

    monkeypatch.setattr("builtins.input", prompt)
    monkeypatch.setattr(main_module, "AutomationEngine", ForbiddenEngine)

    with pytest.raises(SystemExit) as error:
        main_module.main(["--version"])

    assert error.value.code == 0
    assert __version__ in capsys.readouterr().out
    prompt.assert_not_called()


def test_run_script_exposes_read_only_diagnostic_and_version_modes() -> None:
    script = (Path(__file__).parents[1] / "run.ps1").read_text(encoding="utf-8")

    assert '[Alias("Diagnostic")]' in script
    assert "[switch]$Diagnostico" in script
    assert "[switch]$Version" in script
    assert '$ApplicationArguments += "--diagnostico"' in script
    assert '$ApplicationArguments += "--version"' in script
    assert "if ($Diagnostico -and $Version)" in script
    assert "if ($Version)" in script
    assert "if ($Diagnostico)" in script
    assert '$NomeCompleto = Read-Host "Digite seu nome completo"' in script
    assert "exit $ExitCode" in script
    assert "[switch]$Lote" in script
    assert "[string]$Adaptador" in script
    assert '@("--adaptador", $Adaptador)' in script
    assert "$QuantidadeAcoesAssistente = 0" in script
    assert "$AcoesAssistente.Count" not in script
    assert '$ApplicationArguments += "--lote"' in script
    assert "Nao combine -Lote com -Arquivo" in script
