"""Operações seguras de entrada, backup e saída."""

from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

from src.core.exceptions import AutomationError
from src.core.models import RunPaths
from src.services.text import sanitize_filename
from src.settings import BACKUP_DIR, INPUT_DIR, LOG_DIR, OUTPUT_DIR


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}


def ensure_project_directories() -> None:
    """Garante que as pastas operacionais existam."""

    for directory in (INPUT_DIR, OUTPUT_DIR, BACKUP_DIR, LOG_DIR):
        directory.mkdir(parents=True, exist_ok=True)


def resolve_input_file(value: str | Path | None) -> Path:
    """Localiza uma entrada explícita ou a única planilha da pasta input."""

    if value:
        path = Path(value).expanduser().resolve()

        if not path.exists():
            raise AutomationError(
                f"Planilha de entrada não encontrada: {path}"
            )

        if path.suffix.casefold() not in SUPPORTED_EXTENSIONS:
            raise AutomationError(
                "O arquivo de entrada deve ser .xlsx ou .xlsm. "
                f"Recebido: {path.name}"
            )

        if path.name.startswith("~$"):
            raise AutomationError(
                "Foi selecionado um arquivo temporário do Excel. "
                "Feche a planilha e escolha o arquivo sem o prefixo '~$'."
            )

        return path

    candidates = sorted(
        (
            path
            for path in INPUT_DIR.iterdir()
            if path.is_file()
            and path.suffix.casefold() in SUPPORTED_EXTENSIONS
            and not path.name.startswith("~$")
        ),
        key=lambda path: path.name.casefold(),
    )

    if not candidates:
        raise AutomationError(
            f"Nenhuma planilha foi encontrada em {INPUT_DIR}. "
            "Copie o arquivo do teste para a pasta input."
        )

    if len(candidates) > 1:
        names = "\n- ".join(path.name for path in candidates)

        raise AutomationError(
            "Há mais de uma planilha na pasta input. "
            "Informe o arquivo explicitamente para evitar ambiguidade:\n"
            f"- {names}"
        )

    return candidates[0].resolve()


def validate_candidate_name(value: str) -> str:
    """Exige nome e sobrenome, conforme a orientação do teste."""

    cleaned = " ".join(value.split())

    if len(cleaned.split()) < 2:
        raise AutomationError(
            "Informe seu nome completo, com pelo menos nome e sobrenome."
        )

    return cleaned


def prepare_run_paths(
    input_file: Path,
    candidate_name: str,
) -> RunPaths:
    """Define os nomes dos arquivos e cria o backup."""

    ensure_project_directories()

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_candidate = sanitize_filename(candidate_name)

    backup_file = (
        BACKUP_DIR
        / f"{input_file.stem}_backup_{timestamp}{input_file.suffix}"
    )

    output_file = (
        OUTPUT_DIR
        / (
            f"{safe_candidate}_"
            f"Teste_Excel_Analista_Compras"
            f"{input_file.suffix}"
        )
    )

    if output_file.exists():
        output_file = (
            OUTPUT_DIR
            / (
                f"{safe_candidate}_"
                f"Teste_Excel_Analista_Compras_"
                f"{timestamp}"
                f"{input_file.suffix}"
            )
        )

    log_file = LOG_DIR / f"execucao_{timestamp}.log"

    shutil.copy2(input_file, backup_file)

    return RunPaths(
        input_file=input_file,
        backup_file=backup_file,
        output_file=output_file,
        log_file=log_file,
    )