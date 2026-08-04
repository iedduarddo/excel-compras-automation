"""Operações seguras de entrada, backup e saída."""

from __future__ import annotations

import shutil
from datetime import UTC, datetime
from pathlib import Path

from src.core.exceptions import AutomationError
from src.core.models import RunPaths
from src.services.text import sanitize_filename
from src.settings import BACKUP_DIR, INPUT_DIR, LOG_DIR, OUTPUT_DIR

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}


def ensure_project_directories() -> None:
    """Garante que as pastas operacionais existem."""

    for directory in (INPUT_DIR, OUTPUT_DIR, BACKUP_DIR, LOG_DIR):
        directory.mkdir(parents=True, exist_ok=True)


def list_input_files() -> tuple[Path, ...]:
    """Lista, em ordem estável, todas as planilhas válidas da pasta input."""

    ensure_project_directories()
    candidates = tuple(
        sorted(
            (
                path.resolve()
                for path in INPUT_DIR.iterdir()
                if path.is_file()
                and path.suffix.casefold() in SUPPORTED_EXTENSIONS
                and not path.name.startswith("~$")
            ),
            key=lambda path: path.name.casefold(),
        )
    )
    if not candidates:
        raise AutomationError(
            f"Nenhuma planilha foi encontrada em {INPUT_DIR}. "
            "Copie ao menos um arquivo para a pasta input."
        )
    return candidates


def resolve_input_file(value: str | Path | None) -> Path:
    """Localiza uma entrada explícita ou a única planilha da pasta input."""

    if value:
        path = Path(value).expanduser().resolve()
        if not path.exists():
            raise AutomationError(f"Planilha de entrada não encontrada: {path}")
        if path.suffix.casefold() not in SUPPORTED_EXTENSIONS:
            raise AutomationError(
                f"O arquivo de entrada deve ser .xlsx ou .xlsm. Recebido: {path.name}"
            )
        if path.name.startswith("~$"):
            raise AutomationError(
                "Foi selecionado um arquivo temporário do Excel. "
                "Feche a planilha e escolha o arquivo sem o prefixo '~$'."
            )
        return path

    candidates = list_input_files()
    if len(candidates) > 1:
        names = "\n- ".join(path.name for path in candidates)
        raise AutomationError(
            "Há mais de uma planilha na pasta input. Informe --input para evitar "
            f"ambiguidade:\n- {names}"
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
    *,
    output_label: str | None = None,
) -> RunPaths:
    """Define nomes únicos e cria o backup antes do processamento."""

    ensure_project_directories()
    timestamp = datetime.now(UTC).astimezone().strftime("%Y%m%d_%H%M%S")
    safe_candidate = sanitize_filename(candidate_name)
    safe_label = sanitize_filename(output_label) if output_label else ""

    backup_file = _next_available_path(
        BACKUP_DIR,
        f"{input_file.stem}_backup_{timestamp}",
        input_file.suffix,
    )
    output_stem = f"{safe_candidate}_Teste_Excel_Analista_Compras"
    if safe_label:
        output_stem += f"_{safe_label}"
    output_file = _next_available_path(
        OUTPUT_DIR,
        output_stem,
        input_file.suffix,
        collision_suffix=timestamp,
    )
    log_file = _next_available_path(LOG_DIR, f"execucao_{timestamp}", ".log")
    shutil.copy2(input_file, backup_file)

    return RunPaths(
        input_file=input_file,
        backup_file=backup_file,
        output_file=output_file,
        log_file=log_file,
    )


def _next_available_path(
    directory: Path,
    stem: str,
    suffix: str,
    *,
    collision_suffix: str | None = None,
) -> Path:
    """Retorna um caminho livre sem sobrescrever artefatos anteriores."""

    candidate = directory / f"{stem}{suffix}"
    if not candidate.exists():
        return candidate

    alternate_stem = f"{stem}_{collision_suffix}" if collision_suffix else stem
    candidate = directory / f"{alternate_stem}{suffix}"
    counter = 2
    while candidate.exists():
        candidate = directory / f"{alternate_stem}_{counter}{suffix}"
        counter += 1
    return candidate
