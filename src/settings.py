"""Carregamento das configurações externas do projeto."""

from __future__ import annotations

import json
import sys
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


def load_aliases() -> dict[str, Any]:
    """Carrega nomes alternativos de abas, colunas e indicadores."""

    return load_json(CONFIG_DIR / "aliases.json")


def load_rules() -> dict[str, Any]:
    """Carrega pesos e limites do motor de prioridade."""

    return load_json(CONFIG_DIR / "rules.json")
