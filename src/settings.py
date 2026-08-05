"""Carregamento das configurações externas do projeto."""

from __future__ import annotations

import json
import sys
from copy import deepcopy
from pathlib import Path
from typing import Any

from src.core.exceptions import AutomationError


def _resolve_project_root(
    *,
    frozen: bool | None = None,
    executable: str | Path | None = None,
    module_file: str | Path | None = None,
) -> Path:
    """Localiza a raiz persistente no código-fonte e no pacote portátil."""

    is_frozen = getattr(sys, "frozen", False) if frozen is None else frozen
    if is_frozen:
        executable_path = Path(executable or sys.executable)
        return executable_path.resolve().parent

    source_path = Path(module_file or __file__)
    return source_path.resolve().parents[1]


PROJECT_ROOT = _resolve_project_root()
CONFIG_DIR = PROJECT_ROOT / "config"
INPUT_DIR = PROJECT_ROOT / "input"
OUTPUT_DIR = PROJECT_ROOT / "output"
BACKUP_DIR = PROJECT_ROOT / "backup"
LOG_DIR = PROJECT_ROOT / "logs"

ALIAS_GROUPS = (
    "sheets",
    "base_columns",
    "policy_columns",
    "response_columns",
    "indicator_labels",
)


def load_json(path: Path) -> dict[str, Any]:
    """Lê um JSON e transforma erros técnicos em uma mensagem útil."""

    try:
        with path.open("r", encoding="utf-8") as file:
            return json.load(file)
    except FileNotFoundError as error:
        raise AutomationError(
            f"Arquivo de configuração não encontrado: {path}"
        ) from error
    except json.JSONDecodeError as error:
        raise AutomationError(
            f"O arquivo de configuração possui JSON inválido: {path}\n"
            f"Linha {error.lineno}, coluna {error.colno}: {error.msg}"
        ) from error


def load_aliases(adapter: str | Path | None = None) -> dict[str, Any]:
    """Carrega os aliases padrão e aplica um adaptador opcional."""

    aliases = load_json(CONFIG_DIR / "aliases.json")
    if adapter is None:
        return aliases

    adapter_path = Path(adapter).expanduser()
    if not adapter_path.is_absolute():
        adapter_path = PROJECT_ROOT / adapter_path
    profile = load_json(adapter_path.resolve())
    return merge_aliases(aliases, profile)


def merge_aliases(
    aliases: dict[str, Any],
    profile: dict[str, Any],
) -> dict[str, Any]:
    """Acrescenta aliases de um perfil sem alterar a configuração base."""

    additions = profile.get("aliases")
    if not isinstance(additions, dict):
        raise AutomationError(
            "O adaptador deve conter um objeto JSON chamado 'aliases'."
        )

    merged = deepcopy(aliases)
    for group, fields in additions.items():
        if group not in ALIAS_GROUPS or group not in merged:
            raise AutomationError(
                f"Grupo de aliases desconhecido no adaptador: {group}."
            )
        if not isinstance(fields, dict):
            raise AutomationError(
                f"O grupo '{group}' do adaptador deve ser um objeto JSON."
            )

        base_fields = merged[group]
        if not isinstance(base_fields, dict):
            raise AutomationError(f"O grupo '{group}' da configuração base é inválido.")

        for canonical, values in fields.items():
            if canonical not in base_fields:
                raise AutomationError(
                    f"Campo canônico desconhecido em '{group}': {canonical}."
                )
            if (
                not isinstance(values, list)
                or not values
                or any(
                    not isinstance(value, str) or not value.strip() for value in values
                )
            ):
                raise AutomationError(
                    f"Os aliases de '{group}.{canonical}' devem formar uma "
                    "lista não vazia de textos."
                )

            current = list(base_fields[canonical])
            normalized = {value.strip().casefold() for value in current}
            for value in values:
                cleaned = value.strip()
                key = cleaned.casefold()
                if key not in normalized:
                    current.append(cleaned)
                    normalized.add(key)
            base_fields[canonical] = current

    return merged


def load_rules() -> dict[str, Any]:
    """Carrega pesos e limites do motor de prioridade."""

    return load_json(CONFIG_DIR / "rules.json")
