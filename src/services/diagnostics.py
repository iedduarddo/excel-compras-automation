"""Diagnóstico read-only do ambiente e da planilha de entrada."""

from __future__ import annotations

import sys
from collections.abc import Callable
from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path
from typing import Any

from src.excel.detection import WorkbookDetector
from src.excel.excel_desktop import pywin32_is_available
from src.excel.workbook_writer import load_source_workbook
from src.services.files import resolve_input_file
from src.settings import load_aliases, load_rules

REQUIRED_ALIAS_KEYS = {
    "sheets",
    "base_columns",
    "policy_columns",
    "response_columns",
    "indicator_labels",
}
REQUIRED_RULE_KEYS = {
    "priority_weights",
    "priority_thresholds",
    "top_requests",
}


class DiagnosticStatus(StrEnum):
    """Estados exibidos pelo diagnóstico."""

    OK = "OK"
    WARNING = "AVISO"
    ERROR = "ERRO"


@dataclass(frozen=True, slots=True)
class DiagnosticItem:
    """Resultado de uma verificação individual."""

    name: str
    status: DiagnosticStatus
    message: str


@dataclass(frozen=True, slots=True)
class DiagnosticReport:
    """Coleção de verificações e resultado final do ambiente."""

    items: tuple[DiagnosticItem, ...]

    @property
    def ready(self) -> bool:
        """Avisos não impedem a execução pelo modo fallback."""

        return all(item.status is not DiagnosticStatus.ERROR for item in self.items)

    @property
    def exit_code(self) -> int:
        return 0 if self.ready else 1


def run_diagnostics(input_value: str | Path | None) -> DiagnosticReport:
    """Executa verificações sem criar backup, saída ou log."""

    items: list[DiagnosticItem] = [check_python_version()]

    aliases_item, aliases = _check_configuration(
        name="Aliases",
        loader=load_aliases,
        required_keys=REQUIRED_ALIAS_KEYS,
    )
    items.append(aliases_item)

    rules_item, _ = _check_configuration(
        name="Regras",
        loader=load_rules,
        required_keys=REQUIRED_RULE_KEYS,
    )
    items.append(rules_item)

    input_file: Path | None = None
    try:
        input_file = resolve_input_file(input_value)
        items.append(
            DiagnosticItem(
                "Planilha de entrada",
                DiagnosticStatus.OK,
                f"Arquivo localizado: {input_file}",
            )
        )
    except Exception as error:  # noqa: BLE001
        items.append(
            DiagnosticItem(
                "Planilha de entrada",
                DiagnosticStatus.ERROR,
                str(error),
            )
        )

    if input_file is not None:
        if aliases is None:
            items.append(
                DiagnosticItem(
                    "Estrutura da planilha",
                    DiagnosticStatus.ERROR,
                    "Não foi possível detectar as abas porque os aliases são inválidos.",
                )
            )
        else:
            items.extend(_check_workbook_structure(input_file, aliases))

    items.append(_check_native_excel())
    return DiagnosticReport(tuple(items))


def check_python_version(
    version_info: tuple[int, int] | None = None,
) -> DiagnosticItem:
    """Confere a faixa declarada no ``pyproject.toml``."""

    major, minor = version_info or sys.version_info[:2]
    version = f"{major}.{minor}"
    if (3, 11) <= (major, minor) < (3, 15):
        return DiagnosticItem(
            "Python",
            DiagnosticStatus.OK,
            f"Versão {version} suportada.",
        )
    return DiagnosticItem(
        "Python",
        DiagnosticStatus.ERROR,
        f"Versão {version} não suportada; use Python >=3.11 e <3.15.",
    )


def format_diagnostic_report(report: DiagnosticReport) -> str:
    """Produz saída curta e estável para o terminal."""

    lines = [
        "",
        "=" * 72,
        "DIAGNÓSTICO DO AMBIENTE",
        "=" * 72,
    ]
    lines.extend(
        f"[{item.status.value}] {item.name}: {item.message}" for item in report.items
    )
    lines.extend(
        [
            "-" * 72,
            "AMBIENTE PRONTO" if report.ready else "AMBIENTE REQUER ATENÇÃO",
            "=" * 72,
        ]
    )
    return "\n".join(lines)


def _check_configuration(
    *,
    name: str,
    loader: Callable[[], dict[str, Any]],
    required_keys: set[str],
) -> tuple[DiagnosticItem, dict[str, Any] | None]:
    try:
        value = loader()
    except Exception as error:  # noqa: BLE001
        return (
            DiagnosticItem(name, DiagnosticStatus.ERROR, str(error)),
            None,
        )

    if not isinstance(value, dict):
        return (
            DiagnosticItem(
                name,
                DiagnosticStatus.ERROR,
                "A configuração deve conter um objeto JSON no nível principal.",
            ),
            None,
        )

    missing = sorted(required_keys.difference(value))
    if missing:
        return (
            DiagnosticItem(
                name,
                DiagnosticStatus.ERROR,
                "Chaves obrigatórias ausentes: " + ", ".join(missing),
            ),
            None,
        )

    return (
        DiagnosticItem(
            name,
            DiagnosticStatus.OK,
            "Arquivo carregado e chaves obrigatórias encontradas.",
        ),
        value,
    )


def _check_workbook_structure(
    input_file: Path,
    aliases: dict[str, Any],
) -> list[DiagnosticItem]:
    workbook = None
    items: list[DiagnosticItem] = []
    try:
        workbook = load_source_workbook(input_file)
        layout = WorkbookDetector(aliases).detect(workbook)
        items.append(
            DiagnosticItem(
                "Estrutura da planilha",
                DiagnosticStatus.OK,
                (
                    f"Abas detectadas: base={layout.base.title}; "
                    f"policies={layout.policies.title}; "
                    f"responses={layout.responses.title}."
                ),
            )
        )
    except Exception as error:  # noqa: BLE001
        items.append(
            DiagnosticItem(
                "Estrutura da planilha",
                DiagnosticStatus.ERROR,
                str(error),
            )
        )
    finally:
        if workbook is not None:
            try:
                workbook.close()
            except Exception as error:  # noqa: BLE001
                items.append(
                    DiagnosticItem(
                        "Fechamento da planilha",
                        DiagnosticStatus.ERROR,
                        f"Não foi possível fechar o arquivo: {error}",
                    )
                )
    return items


def _check_native_excel() -> DiagnosticItem:
    if pywin32_is_available():
        return DiagnosticItem(
            "Excel Desktop",
            DiagnosticStatus.OK,
            "Integração pywin32 disponível; o Excel será validado na execução.",
        )
    return DiagnosticItem(
        "Excel Desktop",
        DiagnosticStatus.WARNING,
        "Integração nativa indisponível; o fallback por fórmulas está disponível.",
    )
