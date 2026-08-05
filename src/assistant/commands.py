"""Interpretação conservadora dos comandos escritos e falados."""

from __future__ import annotations

import re
from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path

from src.assistant.universal import UniversalAction
from src.core.exceptions import AutomationError
from src.services.text import normalize_text


class AssistantIntent(StrEnum):
    """Intenções determinísticas disponíveis no assistente."""

    HELP = "ajuda"
    RECOGNIZE = "reconhecer"
    DIAGNOSE = "diagnosticar"
    PROCESS = "processar"
    PLAN = "planejar"
    CONFIRM = "confirmar"
    CANCEL = "cancelar"


@dataclass(frozen=True, slots=True)
class AssistantCommand:
    """Comando normalizado antes de acessar qualquer planilha."""

    intent: AssistantIntent
    raw: str
    target: str | None = None
    candidate_name: str | None = None
    adapter: Path | None = None
    use_native_pivot: bool = True
    actions: tuple[UniversalAction, ...] = ()
    plan_id: str | None = None
    options: dict[str, object] | None = None

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
    """Converte uma frase em uma ação permitida, sem executar código livre."""

    raw = " ".join(value.strip().split())
    if not raw:
        raise AutomationError("O comando está vazio.")

    normalized = normalize_text(raw)
    words = set(normalized.split())
    forbidden = {
        "apagar",
        "apague",
        "deletar",
        "delete",
        "enviar",
        "envie",
        "excluir",
        "exclua",
        "mandar",
        "mande",
        "sobrescrever",
    }
    if words.intersection(forbidden) or "e mail" in normalized or "email" in words:
        raise AutomationError(
            "O pedido combina a planilha com uma ação não permitida. O assistente "
            "não apaga originais, sobrescreve arquivos nem envia dados."
        )
    if "plano" in words and {
        "confirmar",
        "confirme",
        "aplicar",
        "aplique",
    }.intersection(words):
        return AssistantCommand(
            intent=AssistantIntent.CONFIRM,
            raw=raw,
            plan_id=_extract_plan_id(raw),
        )
    if "plano" in words and {"cancelar", "cancele", "descartar"}.intersection(words):
        return AssistantCommand(
            intent=AssistantIntent.CANCEL,
            raw=raw,
            plan_id=_extract_plan_id(raw),
        )

    actions = _parse_universal_actions(normalized, words)
    if actions:
        target = _extract_parameter(raw, "arquivo")
        if target is None:
            target = _extract_workbook_name(raw)
        if {"todas", "todos", "lote"}.intersection(words):
            target = None
        options: dict[str, object] = {
            "remove_duplicates": (
                "remover duplicados" in normalized
                or "remova duplicados" in normalized
                or "sem duplicados" in normalized
            ),
            "descending": bool({"decrescente", "descendente"}.intersection(words)),
        }
        column = _extract_parameter(raw, "coluna")
        if column:
            options["column"] = column
        sort_column = _extract_parameter(raw, "ordenar por")
        if sort_column:
            options["sort_column"] = sort_column
        sheet = _extract_parameter(raw, "aba")
        if sheet:
            options["sheet"] = sheet
        return AssistantCommand(
            intent=AssistantIntent.PLAN,
            raw=raw,
            target=target,
            actions=actions,
            options=options,
        )

    matches = [
        intent
        for intent, intent_words in _INTENT_WORDS.items()
        if words.intersection(intent_words)
    ]
    if len(matches) != 1:
        raise AutomationError(
            "Comando não reconhecido. Use ajuda, reconhecer, diagnosticar, processar "
            "ou peça para limpar, organizar, calcular, resumir, criar relatório ou "
            "gerar adaptador."
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
        rf"(?=\s+(?:nome|arquivo|adaptador|aba|coluna|ordenar\s+por|plano)"
        rf"\s*(?:=|:)"
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


def _extract_plan_id(raw: str) -> str:
    value = _extract_parameter(raw, "plano")
    if value is None:
        match = re.search(r"\b[0-9a-f]{12}\b", raw, re.IGNORECASE)
        value = match.group(0) if match else None
    if value is None or not re.fullmatch(r"[0-9a-f]{12}", value, re.IGNORECASE):
        raise AutomationError(
            'Informe o identificador exibido na prévia: confirmar plano="abc123...".'
        )
    return value.casefold()


def _parse_universal_actions(
    normalized: str,
    words: set[str],
) -> tuple[UniversalAction, ...]:
    actions: list[UniversalAction] = []
    action_words = {
        UniversalAction.CLEAN: {
            "limpar",
            "limpe",
            "higienizar",
            "higienize",
            "sanear",
        },
        UniversalAction.ORGANIZE: {
            "organizar",
            "organize",
            "ordenar",
            "ordene",
            "classificar",
        },
        UniversalAction.CALCULATE: {
            "calcular",
            "calcule",
            "calculo",
            "calculos",
            "soma",
            "somar",
            "some",
            "total",
            "totalizar",
            "media",
        },
        UniversalAction.SUMMARIZE: {
            "resumir",
            "resuma",
            "resumo",
            "sintetizar",
        },
        UniversalAction.REPORT: {
            "relatorio",
            "relatorios",
            "painel",
        },
    }
    for action, aliases in action_words.items():
        if words.intersection(aliases):
            actions.append(action)
    if "adaptador" in words and {
        "gerar",
        "gere",
        "criar",
        "crie",
        "sugerir",
        "sugira",
        "mapear",
        "mapeie",
    }.intersection(words):
        actions.append(UniversalAction.GENERATE_ADAPTER)
    if "remover duplicados" in normalized and UniversalAction.CLEAN not in actions:
        actions.append(UniversalAction.CLEAN)
    return tuple(actions)
