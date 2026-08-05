"""Testes do interpretador conservador de comandos escritos."""

from pathlib import Path

import pytest

from src.assistant.commands import AssistantIntent, parse_command
from src.core.exceptions import AutomationError


def test_parse_process_command_extracts_options() -> None:
    command = parse_command(
        'Por favor, processe arquivo="compras agosto.xlsx" '
        'nome="Maria Aparecida" adaptador="cliente.json" sem Excel'
    )

    assert command.intent is AssistantIntent.PROCESS
    assert command.target == "compras agosto.xlsx"
    assert command.candidate_name == "Maria Aparecida"
    assert command.adapter == Path("cliente.json")
    assert command.use_native_pivot is False


def test_parse_recovers_values_split_by_windows_command_line() -> None:
    command = parse_command(
        "processar arquivo=compras agosto.xlsx nome=Maria Aparecida sem excel"
    )

    assert command.target == "compras agosto.xlsx"
    assert command.candidate_name == "Maria Aparecida"
    assert command.use_native_pivot is False


def test_parse_all_files_overrides_incidental_workbook_name() -> None:
    command = parse_command("diagnosticar todas, inclusive antiga.xlsx")

    assert command.intent is AssistantIntent.DIAGNOSE
    assert command.all_files is True


@pytest.mark.parametrize(
    ("text", "intent"),
    [
        ("ajuda", AssistantIntent.HELP),
        ("reconheça todas as planilhas", AssistantIntent.RECOGNIZE),
        ("valide arquivo=entrada.xlsx", AssistantIntent.DIAGNOSE),
        ("execute lote", AssistantIntent.PROCESS),
    ],
)
def test_parse_supported_intents(text: str, intent: AssistantIntent) -> None:
    assert parse_command(text).intent is intent


@pytest.mark.parametrize("text", ["", "apague tudo", "resuma e envie por e-mail"])
def test_parse_rejects_empty_or_unsupported_commands(text: str) -> None:
    with pytest.raises(AutomationError):
        parse_command(text)
