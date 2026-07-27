"""Carregamento das configurações externas do projeto."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from src.core.exceptions import AutomationError


PROJECT_ROOT = Path(__file__).resolve().parents[1]
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