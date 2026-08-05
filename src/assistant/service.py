"""Reconhecimento, diagnóstico e execução orientados por pasta."""

from __future__ import annotations

import os
import shutil
import time
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any
from uuid import uuid4

from src.assistant.commands import AssistantCommand, AssistantIntent, parse_command
from src.assistant.workspace import AssistantWorkspace
from src.core.engine import AutomationEngine
from src.core.exceptions import AutomationError, ExcelDesktopCleanupError
from src.excel.detection import WorkbookDetector
from src.excel.workbook_writer import load_source_workbook
from src.services.diagnostics import format_diagnostic_report, run_diagnostics
from src.services.files import validate_candidate_name
from src.settings import load_aliases


@dataclass(frozen=True, slots=True)
class RecognitionResult:
    """Compatibilidade encontrada para uma planilha da caixa de entrada."""

    input_file: Path
    recognized: bool
    adapter: Path | None = None
    detected_sheets: dict[str, str] | None = None
    detected_columns: dict[str, dict[str, int]] | None = None
    observed_headers: dict[str, list[str]] | None = None
    errors: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class AssistantItemResult:
    input_file: Path
    status: str
    message: str
    output_file: Path | None = None

    @property
    def succeeded(self) -> bool:
        return self.status == "ok"


@dataclass(frozen=True, slots=True)
class AssistantResult:
    command: AssistantCommand
    items: tuple[AssistantItemResult, ...] = ()
    message: str = ""

    @property
    def succeeded(self) -> bool:
        return all(item.succeeded for item in self.items)


class FolderAssistant:
    """Executa somente comandos conhecidos sobre entradas explicitamente listadas."""

    def __init__(
        self,
        workspace: AssistantWorkspace | None = None,
        engine: AutomationEngine | None = None,
    ) -> None:
        self.workspace = workspace or AssistantWorkspace()
        self.engine = engine or AutomationEngine()

    def initialize(self) -> AssistantWorkspace:
        self.workspace.ensure()
        return self.workspace

    def recognize(
        self,
        input_file: Path,
        *,
        requested_adapter: Path | None = None,
    ) -> RecognitionResult:
        """Tenta o padrão e os adaptadores salvos antes de solicitar revisão."""

        self.workspace.ensure()
        adapter_candidates: list[Path | None]
        if requested_adapter is not None:
            adapter_candidates = [self.workspace.resolve_adapter(requested_adapter)]
        else:
            adapter_candidates = [None, *self.workspace.list_adapters()]

        workbook = load_source_workbook(input_file)
        errors: list[str] = []
        try:
            observed_headers = _observe_headers(workbook)
            for adapter in adapter_candidates:
                label = adapter.name if adapter else "mapeamento padrão"
                try:
                    layout = WorkbookDetector(load_aliases(adapter)).detect(workbook)
                except Exception as error:  # noqa: BLE001
                    errors.append(f"{label}: {error}")
                    continue
                return RecognitionResult(
                    input_file=input_file,
                    recognized=True,
                    adapter=adapter,
                    detected_sheets={
                        "base": layout.base.title,
                        "policies": layout.policies.title,
                        "responses": layout.responses.title,
                    },
                    detected_columns={
                        "base": layout.base.columns,
                        "policies": layout.policies.columns,
                        "responses": layout.responses.columns,
                    },
                    observed_headers=observed_headers,
                    errors=tuple(errors),
                )
        finally:
            workbook.close()

        return RecognitionResult(
            input_file=input_file,
            recognized=False,
            observed_headers=observed_headers,
            errors=tuple(errors),
        )

    def execute(self, value: str | AssistantCommand) -> AssistantResult:
        command = parse_command(value) if isinstance(value, str) else value
        self.workspace.ensure()
        if command.intent is AssistantIntent.HELP:
            return AssistantResult(command, message=_help_text())

        inputs = self._select_inputs(command.target)
        if command.intent is AssistantIntent.RECOGNIZE:
            return AssistantResult(
                command,
                tuple(self._recognize_item(path, command.adapter) for path in inputs),
            )
        if command.intent is AssistantIntent.DIAGNOSE:
            return AssistantResult(
                command,
                tuple(self._diagnose_item(path, command.adapter) for path in inputs),
            )
        return AssistantResult(
            command,
            tuple(self._process_items(inputs, command)),
        )

    def run_pending_once(self) -> tuple[AssistantResult, ...]:
        """Consome uma vez os arquivos .txt da fila e os arquiva pelo resultado."""

        self.workspace.ensure()
        results: list[AssistantResult] = []
        pending = sorted(
            self.workspace.pending_commands_dir.glob("*.txt"),
            key=lambda path: path.name.casefold(),
        )
        for command_file in pending:
            claim = _claim_command_file(command_file)
            if claim is None:
                continue
            claimed_file, lock_file = claim

            try:
                destination_dir = self.workspace.completed_commands_dir
                try:
                    result = self.execute(claimed_file.read_text(encoding="utf-8"))
                    if not result.succeeded:
                        destination_dir = self.workspace.failed_commands_dir
                except Exception as error:  # noqa: BLE001
                    destination_dir = self.workspace.failed_commands_dir
                    raw = claimed_file.read_text(encoding="utf-8", errors="replace")
                    fallback = AssistantCommand(AssistantIntent.HELP, raw)
                    result = AssistantResult(
                        fallback,
                        items=(
                            AssistantItemResult(
                                command_file,
                                "erro",
                                str(error),
                            ),
                        ),
                    )

                destination = _next_available(destination_dir / command_file.name)
                shutil.move(str(claimed_file), destination)
                self.workspace.write_json_report(
                    destination_dir,
                    destination.stem + "_resultado",
                    _serialize_result(result),
                )
                results.append(result)
            finally:
                lock_file.unlink(missing_ok=True)
        return tuple(results)

    def watch(self) -> None:
        """Monitora a fila até Ctrl+C; a voz futuramente alimentará esta mesma fila."""

        config = self.workspace.load_config()
        seen: set[Path] = set()
        while True:
            for input_file in self.workspace.list_input_files():
                if input_file not in seen:
                    self._recognize_item(input_file, None)
                    seen.add(input_file)
            self.run_pending_once()
            time.sleep(config.poll_interval_seconds)

    def _select_inputs(self, target: str | None) -> tuple[Path, ...]:
        inputs = self.workspace.list_input_files()
        if not inputs:
            raise AutomationError(
                f"Nenhuma planilha encontrada em {self.workspace.input_dir}."
            )
        if target is None:
            return inputs

        normalized = target.casefold()
        matches = tuple(
            path
            for path in inputs
            if path.name.casefold() == normalized or path.stem.casefold() == normalized
        )
        if len(matches) != 1:
            raise AutomationError(
                f"A planilha '{target}' não foi encontrada de forma única na entrada."
            )
        return matches

    def _recognize_item(
        self,
        input_file: Path,
        requested_adapter: Path | None,
    ) -> AssistantItemResult:
        recognition = self.recognize(
            input_file,
            requested_adapter=requested_adapter,
        )
        payload = _serialize_recognition(recognition)
        if recognition.recognized:
            self.workspace.write_json_report(
                self.workspace.logs_dir,
                f"reconhecimento_{input_file.stem}",
                payload,
            )
            adapter_name = recognition.adapter.name if recognition.adapter else "padrão"
            return AssistantItemResult(
                input_file,
                "ok",
                f"Estrutura reconhecida com o adaptador {adapter_name}.",
            )

        self.workspace.write_json_report(
            self.workspace.review_dir,
            f"revisar_{input_file.stem}",
            payload,
        )
        return AssistantItemResult(
            input_file,
            "revisao",
            "Estrutura desconhecida; relatório criado na pasta revisao.",
        )

    def _diagnose_item(
        self,
        input_file: Path,
        requested_adapter: Path | None,
    ) -> AssistantItemResult:
        recognition = self.recognize(
            input_file,
            requested_adapter=requested_adapter,
        )
        if not recognition.recognized:
            return self._recognize_item(input_file, requested_adapter)

        report = run_diagnostics(input_file, adapter=recognition.adapter)
        formatted = format_diagnostic_report(report)
        log_file = self.workspace.logs_dir / f"diagnostico_{input_file.stem}.txt"
        log_file.write_text(formatted + "\n", encoding="utf-8")
        return AssistantItemResult(
            input_file,
            "ok" if report.ready else "erro",
            "Diagnóstico aprovado." if report.ready else "Diagnóstico requer atenção.",
        )

    def _process_items(
        self,
        inputs: tuple[Path, ...],
        command: AssistantCommand,
    ) -> list[AssistantItemResult]:
        config = self.workspace.load_config()
        candidate_name = command.candidate_name or config.candidate_name
        if not candidate_name:
            raise AutomationError(
                'Informe nome="NOME SOBRENOME" no comando ou configure '
                "candidate_name em assistente_planilhas/config.json."
            )
        candidate_name = validate_candidate_name(candidate_name)
        use_native = command.use_native_pivot and config.use_native_pivot

        items: list[AssistantItemResult] = []
        for input_file in inputs:
            recognition = self.recognize(
                input_file,
                requested_adapter=command.adapter,
            )
            if not recognition.recognized:
                items.append(self._recognize_item(input_file, command.adapter))
                continue
            try:
                result = self.engine.run(
                    input_value=input_file,
                    candidate_name=candidate_name,
                    use_native_pivot=use_native,
                    output_label=input_file.stem,
                    adapter=recognition.adapter,
                    output_dir=self.workspace.output_dir,
                    backup_dir=self.workspace.backup_dir,
                    log_dir=self.workspace.logs_dir,
                )
            except ExcelDesktopCleanupError:
                raise
            except AutomationError as error:
                items.append(AssistantItemResult(input_file, "erro", str(error)))
            else:
                items.append(
                    AssistantItemResult(
                        input_file,
                        "ok",
                        "Processamento concluído.",
                        output_file=result.output_file,
                    )
                )
        return items


def format_assistant_result(result: AssistantResult) -> str:
    lines = ["", "=" * 72, "ASSISTENTE DE PLANILHAS", "=" * 72]
    if result.message:
        lines.append(result.message)
    for item in result.items:
        lines.append(f"[{item.status.upper()}] {item.input_file.name}")
        lines.append(f"        {item.message}")
        if item.output_file is not None:
            lines.append(f"        Saída: {item.output_file}")
    lines.append("=" * 72)
    return "\n".join(lines)


def _observe_headers(workbook: Any) -> dict[str, list[str]]:
    observed: dict[str, list[str]] = {}
    for worksheet in workbook.worksheets:
        best: list[str] = []
        for row in worksheet.iter_rows(
            min_row=1,
            max_row=min(worksheet.max_row, 20),
            max_col=min(worksheet.max_column, 40),
            values_only=True,
        ):
            values = [str(value).strip() for value in row if value not in (None, "")]
            if len(values) > len(best):
                best = values
        observed[worksheet.title] = best
    return observed


def _serialize_recognition(result: RecognitionResult) -> dict[str, Any]:
    payload = asdict(result)
    payload["input_file"] = str(result.input_file)
    payload["adapter"] = str(result.adapter) if result.adapter else None
    return payload


def _serialize_result(result: AssistantResult) -> dict[str, Any]:
    return {
        "command": {
            "intent": result.command.intent.value,
            "raw": result.command.raw,
            "target": result.command.target,
        },
        "message": result.message,
        "succeeded": result.succeeded,
        "items": [
            {
                "input_file": str(item.input_file),
                "status": item.status,
                "message": item.message,
                "output_file": str(item.output_file) if item.output_file else None,
            }
            for item in result.items
        ],
    }


def _next_available(path: Path) -> Path:
    if not path.exists():
        return path
    counter = 2
    while True:
        candidate = path.with_name(f"{path.stem}_{counter}{path.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def _claim_command_file(command_file: Path) -> tuple[Path, Path] | None:
    lock_file = command_file.with_name(f".{command_file.name}.lock")
    try:
        descriptor = os.open(
            lock_file,
            os.O_CREAT | os.O_EXCL | os.O_WRONLY,
        )
    except FileExistsError:
        return None
    os.close(descriptor)

    claimed_file = command_file.with_name(
        f".{command_file.name}.{uuid4().hex}.processing"
    )
    try:
        command_file.replace(claimed_file)
    except FileNotFoundError:
        lock_file.unlink(missing_ok=True)
        return None
    return claimed_file, lock_file


def _help_text() -> str:
    return "\n".join(
        (
            "Comandos disponíveis:",
            "- reconhecer todas",
            "- diagnosticar todas",
            '- diagnosticar arquivo="planilha.xlsx"',
            '- processar todas nome="NOME SOBRENOME"',
            '- processar arquivo="planilha.xlsx" nome="NOME SOBRENOME" sem excel',
            '- processar todas nome="NOME SOBRENOME" adaptador="cliente.json"',
        )
    )
