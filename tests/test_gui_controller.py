"""Testes do controlador desktop sem abrir janelas."""

from hashlib import sha256
from pathlib import Path

import pytest
from openpyxl import Workbook

from src.assistant.commands import parse_command
from src.assistant.service import AssistantItemResult, AssistantResult
from src.core.exceptions import AutomationError
from src.gui.controller import DesktopController, default_desktop_root


def test_controller_imports_supported_files_without_overwriting(tmp_path: Path) -> None:
    source = tmp_path / "origem" / "compras.xlsx"
    source.parent.mkdir()
    source.write_bytes(b"planilha")
    controller = DesktopController(tmp_path / "central")

    first = controller.import_files((source,))
    second = controller.import_files((source,))

    assert [path.name for path in (*first, *second)] == [
        "compras.xlsx",
        "compras_2.xlsx",
    ]
    assert [path.name for path in controller.list_inputs()] == [
        "compras.xlsx",
        "compras_2.xlsx",
    ]
    assert source.read_bytes() == b"planilha"


def test_controller_rejects_unsupported_import_and_direct_processing(
    tmp_path: Path,
) -> None:
    controller = DesktopController(tmp_path / "central")
    unsupported = tmp_path / "dados.csv"
    unsupported.write_text("a,b", encoding="utf-8")

    with pytest.raises(AutomationError, match="Formato não suportado"):
        controller.import_files((unsupported,))
    with pytest.raises(AutomationError, match="recebe uma prévia"):
        controller.execute('processar arquivo="compras.xlsx" nome="Maria Silva"')
    with pytest.raises(AutomationError, match="Use os botões"):
        controller.execute('confirmar plano="abc123def456"')

    temporary = tmp_path / "~$aberta.xlsx"
    temporary.write_bytes(b"temporaria")
    with pytest.raises(AutomationError, match="temporário"):
        controller.import_files((temporary,))


def test_controller_reads_only_preview_from_safe_plans_directory(
    tmp_path: Path,
) -> None:
    controller = DesktopController(tmp_path / "central")
    preview = controller.workspace.plans_dir / "abc123def456.md"
    preview.write_text("# Prévia segura", encoding="utf-8")
    command = parse_command('limpar arquivo="compras.xlsx"')
    result = AssistantResult(
        command,
        (
            AssistantItemResult(
                tmp_path / "compras.xlsx",
                "confirmacao",
                "Revise.",
                preview_file=preview,
                plan_id="abc123def456",
            ),
        ),
    )

    plan_ids, content = controller.read_previews(result)

    assert plan_ids == ("abc123def456",)
    assert content == "# Prévia segura"

    outside = tmp_path / "fora.md"
    outside.write_text("fora", encoding="utf-8")
    unsafe = AssistantResult(
        command,
        (
            AssistantItemResult(
                tmp_path / "compras.xlsx",
                "confirmacao",
                "Revise.",
                preview_file=outside,
                plan_id="abc123def456",
            ),
        ),
    )
    with pytest.raises(AutomationError, match="fora da pasta segura"):
        controller.read_previews(unsafe)


def test_controller_persists_desktop_settings(tmp_path: Path) -> None:
    controller = DesktopController(tmp_path / "central")

    saved = controller.save_config(
        candidate_name="  Carlos Eduardo  ",
        use_native_pivot=False,
        poll_interval_seconds=3,
    )

    assert saved.candidate_name == "Carlos Eduardo"
    assert saved.use_native_pivot is False
    assert saved.poll_interval_seconds == 3


def test_controller_handles_multiple_visible_previews(tmp_path: Path) -> None:
    controller = DesktopController(tmp_path / "central")
    command = parse_command("limpar todas")
    items = []
    for plan_id in ("abc123def456", "def456abc123"):
        preview = controller.workspace.plans_dir / f"{plan_id}.md"
        preview.write_text(f"# Plano {plan_id}", encoding="utf-8")
        items.append(
            AssistantItemResult(
                tmp_path / f"{plan_id}.xlsx",
                "confirmacao",
                "Revise.",
                preview_file=preview,
                plan_id=plan_id,
            )
        )

    plan_ids, content = controller.read_previews(AssistantResult(command, tuple(items)))

    assert plan_ids == ("abc123def456", "def456abc123")
    assert "# Plano abc123def456" in content
    assert "# Plano def456abc123" in content


def test_controller_executes_only_after_real_plan_confirmation(tmp_path: Path) -> None:
    controller = DesktopController(tmp_path / "central")
    source = controller.workspace.input_dir / "clientes.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Cliente", "Valor"])
    sheet.append([" Ana ", 10])
    sheet.append(["Ana", 10])
    workbook.save(source)
    workbook.close()
    original_hash = sha256(source.read_bytes()).hexdigest()

    preview_result = controller.execute('limpar e resumir arquivo="clientes.xlsx"')
    plan_ids, _ = controller.read_previews(preview_result)

    assert len(plan_ids) == 1
    assert not tuple(controller.workspace.output_dir.iterdir())
    assert sha256(source.read_bytes()).hexdigest() == original_hash

    confirmed = controller.confirm_plan(plan_ids[0])

    assert confirmed.succeeded
    assert confirmed.items[0].output_file is not None
    assert confirmed.items[0].output_file.is_file()
    assert tuple(controller.workspace.backup_dir.iterdir())
    assert sha256(source.read_bytes()).hexdigest() == original_hash


def test_desktop_root_uses_user_data_when_frozen(tmp_path: Path) -> None:
    windows = default_desktop_root(
        platform_name="win32",
        frozen=True,
        environment={"LOCALAPPDATA": str(tmp_path / "Local")},
        home=tmp_path,
    )
    macos = default_desktop_root(
        platform_name="darwin",
        frozen=True,
        environment={},
        home=tmp_path,
    )

    assert windows == tmp_path / "Local" / "ExcelComprasAutomation"
    assert macos == (
        tmp_path / "Library" / "Application Support" / "ExcelComprasAutomation"
    )
