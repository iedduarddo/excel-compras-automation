"""Testes dos serviços de arquivos, configurações e logging."""

from __future__ import annotations

import logging
from pathlib import Path

import pytest

from src.core.exceptions import AutomationError
from src.services import files
from src.services.logging_setup import configure_logging
from src.settings import load_aliases, load_json, load_rules


def configure_temporary_directories(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> dict[str, Path]:
    """Redireciona todas as pastas operacionais para uma área temporária."""

    directories = {
        "INPUT_DIR": tmp_path / "input",
        "OUTPUT_DIR": tmp_path / "output",
        "BACKUP_DIR": tmp_path / "backup",
        "LOG_DIR": tmp_path / "logs",
    }
    for name, path in directories.items():
        monkeypatch.setattr(files, name, path)
    return directories


def close_logger_handlers(logger: logging.Logger) -> None:
    """Fecha os arquivos de log abertos durante um teste."""

    for handler in logger.handlers:
        handler.close()
    logger.handlers.clear()


def test_ensure_project_directories_creates_operational_folders(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    directories = configure_temporary_directories(monkeypatch, tmp_path)

    files.ensure_project_directories()

    assert all(path.is_dir() for path in directories.values())


def test_resolve_input_file_accepts_supported_explicit_file(tmp_path: Path) -> None:
    workbook = tmp_path / "Viagens.XLSX"
    workbook.write_bytes(b"conteudo")

    result = files.resolve_input_file(workbook)

    assert result == workbook.resolve()


@pytest.mark.parametrize(
    ("filename", "expected_message"),
    [
        ("inexistente.xlsx", "não encontrada"),
        ("dados.csv", "deve ser .xlsx ou .xlsm"),
        ("~$temporario.xlsx", "arquivo temporário"),
    ],
)
def test_resolve_input_file_rejects_invalid_explicit_input(
    filename: str,
    expected_message: str,
    tmp_path: Path,
) -> None:
    path = tmp_path / filename
    if filename != "inexistente.xlsx":
        path.write_bytes(b"conteudo")

    with pytest.raises(AutomationError, match=expected_message):
        files.resolve_input_file(path)


def test_resolve_input_file_finds_only_valid_candidate(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    directories = configure_temporary_directories(monkeypatch, tmp_path)
    input_dir = directories["INPUT_DIR"]
    input_dir.mkdir()
    expected = input_dir / "teste.xlsm"
    expected.write_bytes(b"planilha")
    (input_dir / "~$teste.xlsx").write_bytes(b"temporario")
    (input_dir / "observacoes.txt").write_text("texto", encoding="utf-8")

    result = files.resolve_input_file(None)

    assert result == expected.resolve()


def test_resolve_input_file_reports_empty_and_ambiguous_directories(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    directories = configure_temporary_directories(monkeypatch, tmp_path)
    input_dir = directories["INPUT_DIR"]
    input_dir.mkdir()

    with pytest.raises(AutomationError, match="Nenhuma planilha"):
        files.resolve_input_file(None)

    (input_dir / "b.xlsx").write_bytes(b"b")
    (input_dir / "a.xlsx").write_bytes(b"a")

    with pytest.raises(AutomationError) as error:
        files.resolve_input_file(None)

    message = str(error.value)
    assert "mais de uma planilha" in message
    assert message.index("a.xlsx") < message.index("b.xlsx")


@pytest.mark.parametrize("value", ["Alex", "  ", "\nCarlos\t"])
def test_validate_candidate_name_requires_name_and_surname(value: str) -> None:
    with pytest.raises(AutomationError, match="nome completo"):
        files.validate_candidate_name(value)


def test_validate_candidate_name_normalizes_whitespace() -> None:
    assert files.validate_candidate_name("  Carlos   Eduardo  ") == "Carlos Eduardo"


def test_prepare_run_paths_creates_backup_and_avoids_output_overwrite(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    directories = configure_temporary_directories(monkeypatch, tmp_path)
    files.ensure_project_directories()
    input_file = directories["INPUT_DIR"] / "teste.xlsm"
    input_file.write_bytes(b"planilha original")
    default_output = (
        directories["OUTPUT_DIR"] / "Carlos_Eduardo_Teste_Excel_Analista_Compras.xlsm"
    )
    default_output.write_bytes(b"resultado anterior")

    result = files.prepare_run_paths(input_file, "Carlos Eduardo")

    assert result.input_file == input_file
    assert result.backup_file.read_bytes() == b"planilha original"
    assert result.backup_file.parent == directories["BACKUP_DIR"]
    assert result.output_file.parent == directories["OUTPUT_DIR"]
    assert result.output_file != default_output
    assert result.output_file.name.startswith(
        "Carlos_Eduardo_Teste_Excel_Analista_Compras_"
    )
    assert result.output_file.suffix == ".xlsm"
    assert result.log_file.parent == directories["LOG_DIR"]
    assert result.log_file.name.startswith("execucao_")


def test_load_json_reads_valid_configuration(tmp_path: Path) -> None:
    config_file = tmp_path / "config.json"
    config_file.write_text('{"limite": 10}', encoding="utf-8")

    assert load_json(config_file) == {"limite": 10}


def test_load_json_translates_missing_and_invalid_json_errors(tmp_path: Path) -> None:
    missing_file = tmp_path / "ausente.json"
    with pytest.raises(AutomationError, match="não encontrado"):
        load_json(missing_file)

    invalid_file = tmp_path / "invalido.json"
    invalid_file.write_text('{\n  "limite":\n}', encoding="utf-8")

    with pytest.raises(AutomationError) as error:
        load_json(invalid_file)

    message = str(error.value)
    assert "JSON inválido" in message
    assert "Linha 3, coluna 1" in message


def test_configuration_helpers_use_config_directory(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    aliases = {"sheets": {"base": ["Base"]}}
    rules = {"top_requests": 5}
    (tmp_path / "aliases.json").write_text(
        '{"sheets": {"base": ["Base"]}}',
        encoding="utf-8",
    )
    (tmp_path / "rules.json").write_text(
        '{"top_requests": 5}',
        encoding="utf-8",
    )
    monkeypatch.setattr("src.settings.CONFIG_DIR", tmp_path)

    assert load_aliases() == aliases
    assert load_rules() == rules


@pytest.mark.parametrize(
    ("verbose", "console_level"),
    [
        (False, logging.INFO),
        (True, logging.DEBUG),
    ],
)
def test_configure_logging_writes_file_and_sets_console_level(
    verbose: bool,
    console_level: int,
    tmp_path: Path,
) -> None:
    log_file = tmp_path / f"execucao_{verbose}.log"
    logger = configure_logging(log_file, verbose=verbose)

    try:
        file_handlers = [
            handler
            for handler in logger.handlers
            if isinstance(handler, logging.FileHandler)
        ]
        console_handlers = [
            handler
            for handler in logger.handlers
            if type(handler) is logging.StreamHandler
        ]

        assert logger.level == logging.DEBUG
        assert len(file_handlers) == 1
        assert file_handlers[0].level == logging.DEBUG
        assert len(console_handlers) == 1
        assert console_handlers[0].level == console_level

        logger.debug("detalhe de diagnóstico")
        file_handlers[0].flush()

        content = log_file.read_text(encoding="utf-8")
        assert "DEBUG" in content
        assert "detalhe de diagnóstico" in content
    finally:
        close_logger_handlers(logger)


def test_configure_logging_replaces_existing_handlers(tmp_path: Path) -> None:
    logger = configure_logging(tmp_path / "primeiro.log")
    first_handlers = tuple(logger.handlers)

    try:
        configured_again = configure_logging(tmp_path / "segundo.log")

        assert configured_again is logger
        assert len(configured_again.handlers) == 2
        assert not any(
            handler in configured_again.handlers for handler in first_handlers
        )
        closed_file_handlers = [
            handler
            for handler in first_handlers
            if isinstance(handler, logging.FileHandler)
        ]
        assert all(handler.stream is None for handler in closed_file_handlers)
    finally:
        close_logger_handlers(logger)
