"""Controlador testável da interface gráfica, sem dependência do Tk."""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from uuid import uuid4

from src.assistant.commands import AssistantIntent, parse_command
from src.assistant.service import (
    AssistantResult,
    FolderAssistant,
    format_assistant_result,
)
from src.assistant.voice import VoiceRecognition, recognize_voice
from src.assistant.workspace import (
    DEFAULT_ASSISTANT_ROOT,
    AssistantConfig,
    AssistantWorkspace,
)
from src.core.exceptions import AutomationError
from src.services.files import SUPPORTED_EXTENSIONS


@dataclass(frozen=True, slots=True)
class DesktopCapabilities:
    """Recursos que variam entre Windows, macOS e outros ambientes."""

    platform: str
    voice_available: bool
    native_excel_available: bool


def detect_capabilities() -> DesktopCapabilities:
    platform = sys.platform
    return DesktopCapabilities(
        platform=platform,
        voice_available=platform == "win32",
        native_excel_available=platform == "win32",
    )


def default_desktop_root(
    *,
    platform_name: str | None = None,
    frozen: bool | None = None,
    environment: Mapping[str, str] | None = None,
    home: Path | None = None,
) -> Path:
    """Mantém dados da GUI em uma pasta gravável quando empacotada."""

    current_platform = platform_name or sys.platform
    is_frozen = getattr(sys, "frozen", False) if frozen is None else frozen
    if not is_frozen:
        return DEFAULT_ASSISTANT_ROOT
    user_home = (home or Path.home()).expanduser().resolve()
    values = os.environ if environment is None else environment
    if current_platform == "win32":
        base = Path(values.get("LOCALAPPDATA", user_home / "AppData" / "Local"))
        return base / "ExcelComprasAutomation"
    if current_platform == "darwin":
        return user_home / "Library" / "Application Support" / "ExcelComprasAutomation"
    base = Path(values.get("XDG_DATA_HOME", user_home / ".local" / "share"))
    return base / "ExcelComprasAutomation"


class DesktopController:
    """Opera o assistente em nome da GUI e preserva seus limites de segurança."""

    def __init__(
        self,
        root: Path | None = None,
        *,
        assistant: FolderAssistant | None = None,
    ) -> None:
        self.workspace = AssistantWorkspace(root or default_desktop_root())
        self.assistant = assistant or FolderAssistant(self.workspace)
        self.capabilities = detect_capabilities()
        self.assistant.initialize()

    def load_config(self) -> AssistantConfig:
        return self.workspace.load_config()

    def save_config(
        self,
        *,
        candidate_name: str,
        use_native_pivot: bool,
        poll_interval_seconds: float,
    ) -> AssistantConfig:
        config = AssistantConfig(
            candidate_name=candidate_name,
            use_native_pivot=use_native_pivot,
            poll_interval_seconds=poll_interval_seconds,
        )
        self.workspace.save_config(config)
        return self.workspace.load_config()

    def list_inputs(self) -> tuple[Path, ...]:
        return self.workspace.list_input_files()

    def import_files(self, sources: tuple[Path, ...]) -> tuple[Path, ...]:
        """Copia entradas para a central sem sobrescrever arquivos existentes."""

        self.workspace.ensure()
        validated: list[Path] = []
        for source in sources:
            source = source.expanduser().resolve()
            if not source.is_file():
                raise AutomationError(f"Planilha não encontrada: {source}")
            if source.name.startswith("~$"):
                raise AutomationError(
                    f"Arquivo temporário do Excel não pode ser importado: {source.name}"
                )
            if source.suffix.casefold() not in SUPPORTED_EXTENSIONS:
                raise AutomationError(
                    f"Formato não suportado: {source.name}. Use .xlsx ou .xlsm."
                )
            validated.append(source)

        imported: list[Path] = []
        try:
            for source in validated:
                imported.append(
                    _copy_without_overwrite(source, self.workspace.input_dir)
                )
        except Exception:
            for destination in imported:
                destination.unlink(missing_ok=True)
            raise
        return tuple(imported)

    def execute(self, command: str) -> AssistantResult:
        """Executa apenas intenções compatíveis com a revisão visual da GUI."""

        parsed = parse_command(command)
        if parsed.intent is AssistantIntent.PROCESS:
            raise AutomationError(
                "Na interface gráfica, pedidos que alteram planilhas devem usar "
                "limpar, organizar, calcular, resumir ou criar relatório. Assim "
                "você recebe uma prévia antes de confirmar."
            )
        if parsed.intent in {AssistantIntent.CONFIRM, AssistantIntent.CANCEL}:
            raise AutomationError(
                "Use os botões Confirmar plano ou Cancelar plano da interface. "
                "Comandos digitados ou falados não podem decidir essa etapa."
            )
        return self.assistant.execute(parsed)

    def confirm_plan(self, plan_id: str) -> AssistantResult:
        return self.assistant.execute(parse_command(f'confirmar plano="{plan_id}"'))

    def cancel_plan(self, plan_id: str) -> AssistantResult:
        return self.assistant.execute(parse_command(f'cancelar plano="{plan_id}"'))

    def read_previews(self, result: AssistantResult) -> tuple[tuple[str, ...], str]:
        plans_root = self.workspace.plans_dir.resolve()
        plan_ids: list[str] = []
        contents: list[str] = []
        for item in result.items:
            if item.preview_file is None or item.plan_id is None:
                continue
            preview = item.preview_file.resolve()
            if preview.parent != plans_root or not preview.is_file():
                raise AutomationError("A prévia retornada está fora da pasta segura.")
            plan_ids.append(item.plan_id)
            contents.append(preview.read_text(encoding="utf-8"))
        if plan_ids:
            return tuple(plan_ids), "\n\n".join(contents)
        return (), format_assistant_result(result)

    def recognize_voice(self) -> VoiceRecognition:
        if not self.capabilities.voice_available:
            raise AutomationError(
                "Use a digitação por voz do sistema no campo Pedido. "
                "O capturador integrado está disponível no Windows."
            )
        return recognize_voice()

    def open_directory(self, directory: Path) -> None:
        directory.mkdir(parents=True, exist_ok=True)
        if sys.platform == "win32":
            os.startfile(directory)  # type: ignore[attr-defined]  # noqa: S606
        elif sys.platform == "darwin":
            subprocess.Popen(("open", str(directory)))  # noqa: S603,S607
        else:
            subprocess.Popen(("xdg-open", str(directory)))  # noqa: S603,S607


def _copy_without_overwrite(source: Path, directory: Path) -> Path:
    temporary = directory / f".import-{uuid4().hex}.tmp"
    destination: Path | None = None
    try:
        shutil.copy2(source, temporary)
        counter = 1
        while destination is None:
            suffix = "" if counter == 1 else f"_{counter}"
            candidate = directory / f"{source.stem}{suffix}{source.suffix}"
            try:
                candidate.touch(exist_ok=False)
            except FileExistsError:
                counter += 1
                continue
            destination = candidate
        temporary.replace(destination)
        return destination.resolve()
    except Exception:
        if destination is not None:
            destination.unlink(missing_ok=True)
        raise
    finally:
        temporary.unlink(missing_ok=True)
