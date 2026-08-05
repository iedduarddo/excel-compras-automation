"""Testes da estrutura persistente do assistente."""

import json

import pytest

from src.assistant.workspace import AssistantWorkspace
from src.core.exceptions import AutomationError
from src.settings import PROJECT_ROOT


def test_assistant_runtime_directory_is_ignored_by_git() -> None:
    gitignore = (PROJECT_ROOT / ".gitignore").read_text(encoding="utf-8")

    assert "assistente_planilhas/" in gitignore.splitlines()


def test_workspace_creates_expected_structure_without_overwriting_config(
    tmp_path,
) -> None:
    workspace = AssistantWorkspace(tmp_path / "assistente")
    workspace.ensure()
    original = workspace.config_file.read_text(encoding="utf-8")
    workspace.ensure()

    assert workspace.config_file.read_text(encoding="utf-8") == original
    for directory in (
        workspace.input_dir,
        workspace.output_dir,
        workspace.backup_dir,
        workspace.logs_dir,
        workspace.review_dir,
        workspace.adapters_dir,
        workspace.pending_commands_dir,
        workspace.completed_commands_dir,
        workspace.failed_commands_dir,
    ):
        assert directory.is_dir()


def test_workspace_lists_only_supported_non_temporary_inputs(tmp_path) -> None:
    workspace = AssistantWorkspace(tmp_path / "assistente")
    workspace.ensure()
    for filename in ("b.xlsm", "A.xlsx", "~$aberta.xlsx", "notas.txt"):
        (workspace.input_dir / filename).write_bytes(b"fixture")

    assert [path.name for path in workspace.list_input_files()] == ["A.xlsx", "b.xlsm"]


def test_workspace_loads_valid_config_and_rejects_invalid_interval(tmp_path) -> None:
    workspace = AssistantWorkspace(tmp_path / "assistente")
    workspace.ensure()
    workspace.config_file.write_text(
        json.dumps(
            {
                "candidate_name": "Maria Aparecida",
                "use_native_pivot": False,
                "poll_interval_seconds": 1.5,
            }
        ),
        encoding="utf-8",
    )

    config = workspace.load_config()

    assert config.candidate_name == "Maria Aparecida"
    assert config.use_native_pivot is False
    assert config.poll_interval_seconds == 1.5

    workspace.config_file.write_text(
        '{"poll_interval_seconds": 0}',
        encoding="utf-8",
    )
    with pytest.raises(AutomationError, match="entre 0.5 e 60"):
        workspace.load_config()
