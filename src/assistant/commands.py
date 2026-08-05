"""Interpretação conservadora dos comandos escritos do assistente."""

from __future__ import annotations

import re
from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path

from src.core.exceptions import AutomationError
from src.services.text import normalize_text


class AssistantIntent(StrEnum):
    """Ações determinísticas disponíveis no primeiro MVP."""

    HELP = "ajuda"
    RECOGNIZE = "reconhecer"
    DIAGNOSE = "diagnosticar"
    PROCESS = "processar"


@dataclass(frozen=True, slots=True)
class AssistantCommand:
    """Comando normalizado antes de acessar qualquer planilha."""

    intent: AssistantIntent
    raw: str
    target: str | None = None
    candidate_name: str | None = None
    adapter: Path | None = None
    use_native_pivot: bool = True

    @property
    def all_files(self) -> bool:
        return self.target is None


_INTENT_WORDS = {
    AssistantIntent.HELP: {"ajuda", "comandos", "help"},
    AssistantIntent.RECOGNIZE: {
        "analisar",
        "analise",
        "catalogar",
        "reconhecer",
        "reconheca",
    },
    AssistantIntent.DIAGNOSE: {
        "diagnosticar",
        "diagnostico",
        "diagnostique",
        "validar",
        "valide",
    },
    AssistantIntent.PROCESS: {
        "automatizar",
        "automatize",
        "executar",
        "execute",
        "processar",
        "processe",
    },
}


def parse_command(value: str) -> AssistantCommand:
    """Converte uma frase conhecida em uma ação sem inferir operações livres."""

    raw = " ".join(value.strip().split())
    if not raw:
        raise AutomationError("O comando está vazio.")

    normalized = normalize_text(raw)
    words = set(normalized.split())
    matches = [
        intent
        for intent, intent_words in _INTENT_WORDS.items()
        if words.intersection(intent_words)
    ]
    if len(matches) != 1:
        raise AutomationError(
            "Comando não reconhecido. Use ajuda, reconhecer, diagnosticar ou "
            "processar. Operações livres ainda não são executadas automaticamente."
        )

    intent = matches[0]
    target = _extract_parameter(raw, "arquivo")
    if target is None:
        target = _extract_workbook_name(raw)
    if {"todas", "todos", "lote"}.intersection(words):
        target = None

    candidate_name = _extract_parameter(raw, "nome")
    adapter_value = _extract_parameter(raw, "adaptador")
    adapter = Path(adapter_value) if adapter_value else None
    use_native_pivot = not (
        "fallback" in words
        or "sem excel" in normalized
        or "sem pivot" in normalized
        or "sem tabela dinamica" in normalized
    )

    return AssistantCommand(
        intent=intent,
        raw=raw,
        target=target,
        candidate_name=candidate_name,
        adapter=adapter,
        use_native_pivot=use_native_pivot,
    )


def _extract_parameter(raw: str, name: str) -> str | None:
    quoted = re.search(
        rf"\b{re.escape(name)}\s*(?:=|:)?\s*[\"']([^\"']+)[\"']",
        raw,
        flags=re.IGNORECASE,
    )
    if quoted:
        return quoted.group(1).strip()

    unquoted = re.search(
        rf"\b{re.escape(name)}\s*(?:=|:)\s*(.+?)"
        rf"(?=\s+(?:nome|arquivo|adaptador)\s*(?:=|:)"
        rf"|\s+sem\s+(?:excel|pivot|tabela)\b|$)",
        raw,
        flags=re.IGNORECASE,
    )
    if unquoted:
        return unquoted.group(1).strip().strip("\"'")
    return None


def _extract_workbook_name(raw: str) -> str | None:
    quoted = re.search(r"[\"']([^\"']+\.(?:xlsx|xlsm))[\"']", raw, re.IGNORECASE)
    if quoted:
        return quoted.group(1).strip()
    plain = re.search(r"\b([^\s\"']+\.(?:xlsx|xlsm))\b", raw, re.IGNORECASE)
    return plain.group(1).strip() if plain else None
