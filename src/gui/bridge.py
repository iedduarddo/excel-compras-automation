"""Ponte testavel entre Qt Quick e o controlador seguro da automacao."""

from __future__ import annotations

from collections.abc import Callable, Sequence
from concurrent.futures import Future, ThreadPoolExecutor
from pathlib import Path
from typing import Any, TypeVar
from uuid import uuid4

from PySide6.QtCore import Property, QObject, QTimer, QUrl, Signal, Slot

from src.assistant.service import AssistantResult, format_assistant_result
from src.gui.controller import DesktopController

_T = TypeVar("_T")


class DesktopBridge(QObject):
    """Estado serializado da GUI; IDs de planos nunca sao expostos ao QML."""

    inputsChanged = Signal()
    busyChanged = Signal()
    statusChanged = Signal()
    previewChanged = Signal()
    planStateChanged = Signal()
    settingsChanged = Signal()
    monitoringChanged = Signal()
    voiceTranscriptReady = Signal(str)
    errorRaised = Signal(str)
    closeBlocked = Signal()
    _completionReady = Signal(object, object)

    def __init__(self, controller: DesktopController) -> None:
        super().__init__()
        self.controller = controller
        self._inputs: list[dict[str, str]] = []
        self._busy = False
        self._status_text = "Pronto"
        self._status_tone = "normal"
        self._preview = (
            "Adicione planilhas, descreva o resultado e revise a previa antes "
            "de confirmar."
        )
        self._plan_ids: tuple[str, ...] = ()
        self._plan_revision = ""
        self._confirmation_tokens: dict[str, tuple[str, tuple[str, ...]]] = {}
        self._candidate_name = ""
        self._use_native_pivot = False
        self._poll_interval_seconds = 2.0
        self._monitoring = False
        self._executor = ThreadPoolExecutor(
            max_workers=1, thread_name_prefix="excel-qml"
        )
        self._monitor = QTimer(self)
        self._monitor.timeout.connect(self._monitor_tick)
        self._completionReady.connect(self._finish_operation)

    @Property(list, notify=inputsChanged)
    def inputs(self) -> list[dict[str, str]]:
        return list(self._inputs)

    @Property(bool, notify=busyChanged)
    def busy(self) -> bool:
        return self._busy

    @Property(str, notify=statusChanged)
    def statusText(self) -> str:  # noqa: N802 - nome consumido pelo QML
        return self._status_text

    @Property(str, notify=statusChanged)
    def statusTone(self) -> str:  # noqa: N802 - nome consumido pelo QML
        return self._status_tone

    @Property(str, notify=previewChanged)
    def previewMarkdown(self) -> str:  # noqa: N802 - nome consumido pelo QML
        return self._preview

    @Property(int, notify=planStateChanged)
    def planCount(self) -> int:  # noqa: N802 - nome consumido pelo QML
        return len(self._plan_ids)

    @Property(str, notify=planStateChanged)
    def planRevision(self) -> str:  # noqa: N802 - nome consumido pelo QML
        return self._plan_revision

    @Property(bool, notify=planStateChanged)
    def canConfirm(self) -> bool:  # noqa: N802 - nome consumido pelo QML
        return bool(self._plan_ids) and not self._busy

    @Property(str, constant=True)
    def platformName(self) -> str:  # noqa: N802 - nome consumido pelo QML
        return self.controller.capabilities.platform

    @Property(bool, constant=True)
    def voiceAvailable(self) -> bool:  # noqa: N802 - nome consumido pelo QML
        return self.controller.capabilities.voice_available

    @Property(bool, constant=True)
    def nativeExcelAvailable(self) -> bool:  # noqa: N802
        return self.controller.capabilities.native_excel_available

    @Property(str, notify=settingsChanged)
    def candidateName(self) -> str:  # noqa: N802 - nome consumido pelo QML
        return self._candidate_name

    @Property(bool, notify=settingsChanged)
    def useNativePivot(self) -> bool:  # noqa: N802 - nome consumido pelo QML
        return self._use_native_pivot

    @Property(float, notify=settingsChanged)
    def pollIntervalSeconds(self) -> float:  # noqa: N802
        return self._poll_interval_seconds

    @Property(bool, notify=monitoringChanged)
    def monitoring(self) -> bool:
        return self._monitoring

    @Slot()
    def initialize(self) -> None:
        try:
            config = self.controller.load_config()
            self._candidate_name = config.candidate_name
            self._use_native_pivot = config.use_native_pivot
            self._poll_interval_seconds = config.poll_interval_seconds
            self.settingsChanged.emit()
            self._refresh_inputs(announce=False)
        except Exception as error:  # noqa: BLE001 - fronteira visual
            self._raise_error(str(error))

    @Slot(list)
    def importFiles(self, values: list[Any]) -> None:  # noqa: N802
        try:
            paths = tuple(_local_path(value) for value in values)
        except (TypeError, ValueError) as error:
            self._raise_error(str(error))
            return
        self._submit(
            lambda: self.controller.import_files(paths),
            self._after_import,
            "Copiando planilhas para a area segura...",
        )

    @Slot()
    def refreshInputs(self) -> None:  # noqa: N802
        try:
            self._refresh_inputs(announce=True)
        except Exception as error:  # noqa: BLE001 - fronteira visual
            self._raise_error(str(error))

    @Slot(str)
    def executeRequest(self, text: str) -> None:  # noqa: N802
        command = text.strip()
        if not command:
            self._raise_error("Digite ou fale um pedido.")
            return
        self._submit(
            lambda: self.controller.execute(command),
            self._show_result,
            "Interpretando o pedido e preparando a previa...",
        )

    @Slot()
    def captureVoice(self) -> None:  # noqa: N802
        self._submit(
            self.controller.recognize_voice,
            self._after_voice,
            "Aguardando comando por voz...",
        )

    @Slot(str)
    def openWorkspace(self, kind: str) -> None:  # noqa: N802
        directories = {
            "input": self.controller.workspace.input_dir,
            "output": self.controller.workspace.output_dir,
            "backup": self.controller.workspace.backup_dir,
            "logs": self.controller.workspace.logs_dir,
        }
        directory = directories.get(kind)
        if directory is None:
            self._raise_error("Pasta nao autorizada pela interface.")
            return
        try:
            self.controller.open_directory(directory)
        except Exception as error:  # noqa: BLE001 - fronteira visual
            self._raise_error(str(error))

    @Slot(str, bool, float)
    def saveSettings(  # noqa: N802
        self, candidate_name: str, native_pivot: bool, interval: float
    ) -> None:
        self._submit(
            lambda: self.controller.save_config(
                candidate_name=candidate_name,
                use_native_pivot=native_pivot,
                poll_interval_seconds=interval,
            ),
            self._after_settings,
            "Salvando configuracoes...",
        )

    @Slot()
    def toggleMonitor(self) -> None:  # noqa: N802
        if self._monitoring:
            self._monitor.stop()
            self._monitoring = False
            self.monitoringChanged.emit()
            self._set_status("Monitor parado")
            return
        self._monitor.setInterval(max(int(self._poll_interval_seconds * 1000), 500))
        self._monitor.start()
        self._monitoring = True
        self.monitoringChanged.emit()
        self._set_status("Monitor ativo: acompanhando a pasta de entrada", "success")
        self._refresh_inputs(announce=False)

    @Slot(result=str)
    def prepareConfirmation(self) -> str:  # noqa: N802
        if not self.canConfirm:
            return ""
        token = uuid4().hex
        self._confirmation_tokens.clear()
        self._confirmation_tokens[token] = (self._plan_revision, self._plan_ids)
        return token

    @Slot(str)
    def confirmPlans(self, token: str) -> None:  # noqa: N802
        snapshot = self._confirmation_tokens.pop(token, None)
        self._confirmation_tokens.clear()
        if snapshot is None:
            self._raise_error("A confirmacao expirou. Revise a previa atual.")
            return
        revision, plan_ids = snapshot
        if revision != self._plan_revision or plan_ids != self._plan_ids or self._busy:
            self._raise_error("A previa mudou. Revise novamente antes de confirmar.")
            return
        self._invalidate_plans()
        self._submit(
            lambda: tuple(
                self.controller.confirm_plan(plan_id) for plan_id in plan_ids
            ),
            self._show_plan_results,
            "Aplicando os planos confirmados...",
        )

    @Slot(str)
    def cancelPlans(self, revision: str) -> None:  # noqa: N802
        if revision != self._plan_revision or not self._plan_ids or self._busy:
            self._raise_error("A previa mudou. Atualize a tela antes de cancelar.")
            return
        plan_ids = self._plan_ids
        self._submit(
            lambda: tuple(self.controller.cancel_plan(plan_id) for plan_id in plan_ids),
            self._show_plan_results,
            "Cancelando os planos...",
        )

    @Slot(result=bool)
    def requestClose(self) -> bool:  # noqa: N802
        if self._busy:
            self.closeBlocked.emit()
            return False
        self._monitor.stop()
        self._executor.shutdown(wait=False, cancel_futures=True)
        return True

    def _refresh_inputs(self, *, announce: bool) -> None:
        paths = self.controller.list_inputs()
        self._inputs = [{"name": path.name, "path": str(path)} for path in paths]
        self.inputsChanged.emit()
        if announce:
            count = len(paths)
            self._set_status(
                "Adicione uma planilha para comecar"
                if count == 0
                else f"Lista atualizada: {count} planilha(s)"
            )

    def _monitor_tick(self) -> None:
        if not self._busy:
            try:
                self._refresh_inputs(announce=False)
            except Exception as error:  # noqa: BLE001 - fronteira visual
                self._raise_error(str(error))

    def _after_import(self, value: object) -> None:
        imported = tuple(value)  # type: ignore[arg-type]
        self._refresh_inputs(announce=False)
        self._set_status(f"{len(imported)} planilha(s) adicionada(s)", "success")

    def _after_voice(self, value: object) -> None:
        text = str(value.text)  # type: ignore[attr-defined]
        self.voiceTranscriptReady.emit(text)
        self._set_status("Revise a transcricao e clique em Executar.", "warning")

    def _after_settings(self, value: object) -> None:
        self._candidate_name = str(value.candidate_name)  # type: ignore[attr-defined]
        self._use_native_pivot = bool(  # type: ignore[attr-defined]
            value.use_native_pivot
        )
        self._poll_interval_seconds = float(  # type: ignore[attr-defined]
            value.poll_interval_seconds
        )
        self.settingsChanged.emit()
        if self._monitoring:
            self._monitor.setInterval(int(self._poll_interval_seconds * 1000))
        self._set_status("Configuracoes salvas", "success")

    def _show_result(self, value: object) -> None:
        if not isinstance(value, AssistantResult):
            raise TypeError("Resultado inesperado do assistente.")
        plan_ids, content = self.controller.read_previews(value)
        for plan_id in plan_ids:
            content = content.replace(plan_id, "[identificador protegido]")
        self._confirmation_tokens.clear()
        self._plan_ids = plan_ids
        self._plan_revision = uuid4().hex if plan_ids else ""
        self._preview = content
        self.previewChanged.emit()
        self.planStateChanged.emit()
        self._refresh_inputs(announce=False)
        self._set_status(
            "Previa pronta: revise e confirme para gerar as copias"
            if plan_ids
            else "Operacao concluida",
            "warning" if plan_ids else "success",
        )

    def _show_plan_results(self, value: object) -> None:
        results: Sequence[AssistantResult] = tuple(value)  # type: ignore[arg-type]
        self._invalidate_plans()
        self._preview = "\n".join(format_assistant_result(item) for item in results)
        self.previewChanged.emit()
        self._refresh_inputs(announce=False)
        self._set_status(
            "Planos concluidos. Os arquivos originais foram preservados.", "success"
        )

    def _invalidate_plans(self) -> None:
        self._confirmation_tokens.clear()
        self._plan_ids = ()
        self._plan_revision = ""
        self.planStateChanged.emit()

    def _submit(
        self,
        operation: Callable[[], _T],
        callback: Callable[[_T], None] | None,
        status: str,
    ) -> None:
        if self._busy:
            self._set_status("Aguarde a operacao atual terminar", "warning")
            return
        self._set_busy(True)
        self._set_status(status)
        future: Future[object] = self._executor.submit(operation)
        future.add_done_callback(
            lambda completed: self._completionReady.emit(completed, callback)
        )

    @Slot(object, object)
    def _finish_operation(
        self, future: Future[object], callback: Callable[[Any], None] | None
    ) -> None:
        try:
            result = future.result()
            if callback is not None:
                callback(result)
        except Exception as error:  # noqa: BLE001 - fronteira visual
            self._raise_error(str(error))
        finally:
            self._set_busy(False)

    def _set_busy(self, active: bool) -> None:
        if self._busy == active:
            return
        self._busy = active
        self.busyChanged.emit()
        self.planStateChanged.emit()

    def _set_status(self, message: str, tone: str = "normal") -> None:
        self._status_text = message
        self._status_tone = tone
        self.statusChanged.emit()

    def _raise_error(self, message: str) -> None:
        self._set_status(message, "error")
        self.errorRaised.emit(message)


def _local_path(value: object) -> Path:
    url = value if isinstance(value, QUrl) else QUrl(str(value))
    local = url.toLocalFile() if url.isLocalFile() else ""
    if not local:
        raise ValueError("Somente arquivos locais podem ser importados.")
    return Path(local)
