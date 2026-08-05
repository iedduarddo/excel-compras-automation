"""Testes do interpretador conservador de comandos escritos."""

from pathlib import Path

import pytest

from src.assistant.commands import AssistantIntent, parse_command
from src.assistant.universal import UniversalAction
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


def test_parse_varied_universal_request_builds_safe_plan() -> None:
    command = parse_command(
        'Por favor, limpe, organize e crie um relatório de arquivo="clientes.xlsx" '
        'remover duplicados ordenar por="Cidade" decrescente'
    )

    assert command.intent is AssistantIntent.PLAN
    assert command.target == "clientes.xlsx"
    assert command.actions == (
        UniversalAction.CLEAN,
        UniversalAction.ORGANIZE,
        UniversalAction.REPORT,
    )
    assert command.options == {
        "remove_duplicates": True,
        "descending": True,
        "sort_column": "Cidade",
    }


@pytest.mark.parametrize(
    ("text", "intent"),
    [
        ('confirmar plano="a1b2c3d4e5f6"', AssistantIntent.CONFIRM),
        ("cancele o plano a1b2c3d4e5f6", AssistantIntent.CANCEL),
    ],
)
def test_parse_plan_decision_requires_explicit_identifier(
    text: str,
    intent: AssistantIntent,
) -> None:
    command = parse_command(text)

    assert command.intent is intent
    assert command.plan_id == "a1b2c3d4e5f6"
