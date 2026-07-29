"""Contratos dos launchers Windows, sem instalar pacotes nem abrir o Excel."""

from __future__ import annotations

import os
import re
import shutil
import subprocess  # noqa: S404
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SETUP_SCRIPT = PROJECT_ROOT / "setup.ps1"
RUN_SCRIPT = PROJECT_ROOT / "run.ps1"
LAUNCHER = PROJECT_ROOT / "iniciar.cmd"


def read_script(path: Path) -> str:
    assert path.is_file(), f"Script obrigatório ausente: {path.name}"
    return path.read_text(encoding="utf-8")


def find_powershell() -> str | None:
    return shutil.which("powershell") or shutil.which("pwsh")


@pytest.mark.parametrize("script_path", [SETUP_SCRIPT, RUN_SCRIPT])
def test_powershell_scripts_have_valid_syntax(script_path: Path) -> None:
    """Apenas analisa a AST; o conteúdo do script não é executado."""

    executable = find_powershell()
    if executable is None:
        pytest.skip("PowerShell não está disponível neste sistema.")

    command = """
$tokens = $null
$errors = $null
[System.Management.Automation.Language.Parser]::ParseFile(
    $env:CODEX_SCRIPT_TO_PARSE,
    [ref]$tokens,
    [ref]$errors
) | Out-Null
if ($errors.Count -gt 0) {
    [Console]::Error.WriteLine(
        ($errors | ForEach-Object { $_.Message }) -join [Environment]::NewLine
    )
    exit 1
}
"""
    environment = os.environ.copy()
    environment["CODEX_SCRIPT_TO_PARSE"] = str(script_path)

    completed = subprocess.run(  # noqa: S603
        [
            executable,
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            command,
        ],
        check=False,
        capture_output=True,
        text=True,
        env=environment,
    )

    assert completed.returncode == 0, completed.stderr or completed.stdout


def test_setup_accepts_only_supported_python_versions() -> None:
    script = read_script(SETUP_SCRIPT)
    lower_bound = re.search(
        r"3[.,]11|Minor\s*-ge\s*11",
        script,
        flags=re.IGNORECASE,
    )
    upper_bound = re.search(
        r"3[.,](?:14|15)|Minor\s*-(?:le\s*14|lt\s*15)",
        script,
        flags=re.IGNORECASE,
    )

    assert lower_bound, "setup.ps1 deve rejeitar Python anterior ao 3.11."
    assert upper_bound, "setup.ps1 deve rejeitar Python posterior ao 3.14."


def test_setup_python_probe_survives_windows_powershell_quoting() -> None:
    executable = find_powershell()
    if executable is None:
        pytest.skip("PowerShell não está disponível neste sistema.")

    script = read_script(SETUP_SCRIPT)
    probe_match = re.search(
        r'\$SupportedPythonProbe\s*=\s*"(?P<probe>[^"\r\n]+)"',
        script,
    )
    assert probe_match, "O probe deve evitar aspas duplas internas no PowerShell 5.1."

    environment = os.environ.copy()
    environment["EXCEL_COMPRAS_TEST_PYTHON"] = sys.executable
    environment["EXCEL_COMPRAS_TEST_PROBE"] = probe_match.group("probe")
    completed = subprocess.run(  # noqa: S603
        [
            executable,
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            (
                "& $env:EXCEL_COMPRAS_TEST_PYTHON -I -c "
                "$env:EXCEL_COMPRAS_TEST_PROBE; exit $LASTEXITCODE"
            ),
        ],
        check=False,
        capture_output=True,
        text=True,
        env=environment,
    )

    assert completed.returncode == 0, completed.stderr or completed.stdout
    assert f"{sys.version_info.major}.{sys.version_info.minor}." in completed.stdout


def test_setup_stops_after_external_command_failure() -> None:
    script = read_script(SETUP_SCRIPT)

    assert "$LASTEXITCODE" in script
    direct_guard = re.search(
        r"\$LASTEXITCODE\s*-(?:ne|gt)\s*0",
        script,
        flags=re.IGNORECASE,
    )
    captured_exit = re.search(
        r"\$(?P<name>\w+)\s*=\s*\$LASTEXITCODE",
        script,
        flags=re.IGNORECASE,
    )
    captured_guard = bool(
        captured_exit
        and re.search(
            rf"\${captured_exit.group('name')}\s*-(?:ne|gt)\s*0"
            r"[\s\S]{0,300}(?:throw|exit)",
            script,
            flags=re.IGNORECASE,
        )
    )

    assert direct_guard or captured_guard


def test_setup_creates_operational_directories_idempotently() -> None:
    script = read_script(SETUP_SCRIPT)

    for directory in ("input", "output", "backup", "logs"):
        assert re.search(
            rf"""["'][^"'\r\n]*[\\/]*{directory}[\\/]*["']""",
            script,
            flags=re.IGNORECASE,
        ), f"setup.ps1 deve preparar a pasta {directory}."

    idempotent_directory_creation = re.search(
        r"New-Item[\s\S]{0,200}-Force|"
        r"\[System\.IO\.Directory\]::CreateDirectory|"
        r"\[IO\.Directory\]::CreateDirectory",
        script,
        flags=re.IGNORECASE,
    )
    assert idempotent_directory_creation
    assert re.search(
        r"Test-Path[\s\S]{0,200}(?:\.venv|\$Venv(?:Python|Directory))",
        script,
        flags=re.IGNORECASE,
    )


def test_run_uses_automatic_input_discovery_by_default() -> None:
    script = read_script(RUN_SCRIPT)

    assert re.search(
        r"\[string\]\s*\$Arquivo\s*=\s*(?:\"\"|'')",
        script,
        flags=re.IGNORECASE,
    ), "-Arquivo deve ser vazio por padrão."
    assert re.search(
        r"if\s*\((?=[^{}]*\$Arquivo)[^{}]*\)\s*\{"
        r"(?=[^{}]*--input)[^{}]*\}",
        script,
        flags=re.IGNORECASE | re.DOTALL,
    ), "--input só deve ser enviado dentro de uma condição que valide -Arquivo."


def test_run_preserves_public_switches_and_exit_code() -> None:
    script = read_script(RUN_SCRIPT)

    for parameter in ("Diagnostico", "Version", "SemPivotNativo", "Verbose"):
        assert re.search(
            rf"\[switch\]\s*\${parameter}\b",
            script,
            flags=re.IGNORECASE,
        )

    assert re.search(r'\[Alias\(\s*"Diagnostic"\s*\)\]', script)
    for argument in (
        "--diagnostico",
        "--version",
        "--sem-pivot-nativo",
        "--verbose",
    ):
        assert argument in script

    direct_exit = re.search(
        r"exit\s+\$LASTEXITCODE\b",
        script,
        flags=re.IGNORECASE,
    )
    captured_exit = re.search(
        r"\$(?P<name>\w+)\s*=\s*\$LASTEXITCODE",
        script,
        flags=re.IGNORECASE,
    )
    captured_status_is_returned = bool(
        captured_exit
        and re.search(
            rf"exit\s+\${captured_exit.group('name')}\b",
            script,
            flags=re.IGNORECASE,
        )
    )
    assert direct_exit or captured_status_is_returned


def test_launcher_bootstraps_then_runs_and_propagates_status() -> None:
    script = read_script(LAUNCHER)

    assert re.search(
        r"if\s+not\s+exist[\s\S]{0,400}(?:%VENV_PYTHON%|"
        r"\.venv[\\/]+Scripts[\\/]+python\.exe)",
        script,
        flags=re.IGNORECASE,
    )
    assert re.search(
        r"if\s+not\s+exist[\s\S]{0,600}setup\.ps1",
        script,
        flags=re.IGNORECASE,
    )

    setup_position = script.casefold().index("setup.ps1")
    run_position = script.casefold().index("run.ps1")
    assert setup_position < run_position
    assert re.search(
        r"setup\.ps1[\s\S]{0,500}(?:if\s+errorlevel|\|\||"
        r"(?:%|!)?\w*SETUP\w*(?:%|!))["
        r"\s\S]{0,300}exit\s+/b",
        script,
        flags=re.IGNORECASE,
    ), "Uma falha do setup deve impedir a execução do run.ps1."

    assert re.search(
        r"run\.ps1[^\r\n]*%\*",
        script,
        flags=re.IGNORECASE,
    ), "iniciar.cmd deve repassar todos os argumentos ao run.ps1."
    assert re.search(
        r"exit\s+/b\s+(?:%ERRORLEVEL%|!ERRORLEVEL!|%\w+%|!\w+!)",
        script,
        flags=re.IGNORECASE,
    ), "iniciar.cmd deve devolver o código de saída do run.ps1."

    for required_token in (
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
    ):
        assert required_token.casefold() in script.casefold()


def test_launcher_pauses_only_for_argumentless_interactive_use() -> None:
    script = read_script(LAUNCHER)

    assert re.search(
        r'if\s+"%~1"\s*==\s*""\s+set\s+"[^"]*PAUSE[^"]*=1"',
        script,
        flags=re.IGNORECASE,
    )
    assert re.search(
        r'if\s+"%\w*PAUSE\w*%"\s*==\s*"1"\s+pause',
        script,
        flags=re.IGNORECASE,
    )
    assert not re.search(
        r"(?m)^\s*pause\s*$",
        script,
        flags=re.IGNORECASE,
    ), "A pausa não pode bloquear chamadas automatizadas com argumentos."
