"""Testes do reconhecimento e da execução orientada por pastas."""

from __future__ import annotations

import json
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from threading import Barrier

from openpyxl import Workbook

from scripts.build_portable import create_smoke_workbook
from src.assistant.service import FolderAssistant
from src.assistant.workspace import AssistantWorkspace
from src.core.models import RunResult
from src.settings import CONFIG_DIR


def make_supported_input(
    workspace: AssistantWorkspace, name: str = "entrada.xlsx"
) -> Path:
    workspace.ensure()
    input_file = workspace.input_dir / name
    create_smoke_workbook(input_file, CONFIG_DIR / "aliases.json")
    return input_file


def test_assistant_recognizes_supported_workbook(tmp_path) -> None:
    workspace = AssistantWorkspace(tmp_path / "assistente")
    input_file = make_supported_input(workspace)

    recognition = FolderAssistant(workspace).recognize(input_file)

    assert recognition.recognized is True
    assert recognition.adapter is None
    assert recognition.detected_sheets == {
        "base": "Base_Viagens",
        "policies": "Políticas_Fornecedores",
        "responses": "Respostas",
    }


def test_unknown_workbook_is_sent_to_review_without_modification(tmp_path) -> None:
    workspace = AssistantWorkspace(tmp_path / "assistente")
    workspace.ensure()
    input_file = workspace.input_dir / "desconhecida.xlsx"
    workbook = Workbook()
    workbook.active.append(["Produto", "Quantidade"])
    workbook.active.append(["Caneta", 2])
    workbook.save(input_file)
    original = input_file.read_bytes()

    result = FolderAssistant(workspace).execute("reconhecer todas")

    assert result.succeeded is False
    assert result.items[0].status == "revisao"
    assert input_file.read_bytes() == original
    report = workspace.review_dir / "revisar_desconhecida.json"
    payload = json.loads(report.read_text(encoding="utf-8"))
    assert payload["recognized"] is False
    assert payload["observed_headers"]["Sheet"] == ["Produto", "Quantidade"]


def test_process_uses_workspace_artifact_directories_and_config(tmp_path) -> None:
    workspace = AssistantWorkspace(tmp_path / "assistente")
    input_file = make_supported_input(workspace)
    config = json.loads(workspace.config_file.read_text(encoding="utf-8"))
    config["candidate_name"] = "Maria Aparecida"
    workspace.config_file.write_text(json.dumps(config), encoding="utf-8")
    received: dict[str, object] = {}

    class FakeEngine:
        def run(self, **kwargs: object) -> RunResult:
            received.update(kwargs)
            return RunResult(
                output_file=workspace.output_dir / "resultado.xlsx",
                backup_file=workspace.backup_dir / "backup.xlsx",
                log_file=workspace.logs_dir / "execucao.log",
                native_pivot_created=False,
                detected_sheets={},
                detected_columns={},
                checks={},
            )

    result = FolderAssistant(workspace, FakeEngine()).execute(
        "processar todas sem excel"
    )

    assert result.succeeded is True
    assert result.items[0].input_file == input_file.resolve()
    assert received["candidate_name"] == "Maria Aparecida"
    assert received["use_native_pivot"] is False
    assert received["output_dir"] == workspace.output_dir
    assert received["backup_dir"] == workspace.backup_dir
    assert received["log_dir"] == workspace.logs_dir


def test_pending_command_is_archived_with_result(tmp_path) -> None:
    workspace = AssistantWorkspace(tmp_path / "assistente")
    workspace.ensure()
    command_file = workspace.pending_commands_dir / "001-ajuda.txt"
    command_file.write_text("ajuda", encoding="utf-8")

    results = FolderAssistant(workspace).run_pending_once()

    assert len(results) == 1
    assert not command_file.exists()
    assert (workspace.completed_commands_dir / command_file.name).is_file()
    assert (workspace.completed_commands_dir / "001-ajuda_resultado.json").is_file()


def test_concurrent_consumers_claim_pending_command_once(tmp_path, monkeypatch) -> None:
    workspace = AssistantWorkspace(tmp_path / "assistente")
    workspace.ensure()
    command_file = workspace.pending_commands_dir / "001-ajuda.txt"
    command_file.write_text("ajuda", encoding="utf-8")
    barrier = Barrier(2)
    original_glob = Path.glob

    def synchronized_glob(path, pattern, *args, **kwargs):
        matches = list(original_glob(path, pattern, *args, **kwargs))
        if path == workspace.pending_commands_dir and pattern == "*.txt":
            barrier.wait(timeout=5)
        return iter(matches)

    monkeypatch.setattr(Path, "glob", synchronized_glob)

    with ThreadPoolExecutor(max_workers=2) as executor:
        batches = tuple(
            executor.map(
                lambda _: FolderAssistant(workspace).run_pending_once(),
                range(2),
            )
        )

    results = tuple(result for batch in batches for result in batch)
    assert len(results) == 1
    assert results[0].succeeded is True
    assert not command_file.exists()
    assert (workspace.completed_commands_dir / command_file.name).is_file()
    assert len(tuple(workspace.completed_commands_dir.glob("*.txt"))) == 1
    assert not tuple(workspace.pending_commands_dir.glob("*.processing"))
    assert not tuple(workspace.pending_commands_dir.glob("*.lock"))
