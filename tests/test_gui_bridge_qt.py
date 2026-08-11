"""Limites de seguranca da ponte Qt/QML."""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

pytest.importorskip("PySide6")

from PySide6.QtGui import QGuiApplication

from src.gui.bridge import DesktopBridge


class FakeController:
    def __init__(self, root: Path) -> None:
        self.capabilities = SimpleNamespace(
            platform="win32", voice_available=True, native_excel_available=True
        )
        self.workspace = SimpleNamespace(
            input_dir=root / "entrada",
            output_dir=root / "saida",
            backup_dir=root / "backup",
            logs_dir=root / "logs",
        )
        self.executed: list[str] = []
        self.confirmed: list[str] = []
        self.cancelled: list[str] = []

    def load_config(self):
        return SimpleNamespace(
            candidate_name="Pessoa Teste",
            use_native_pivot=False,
            poll_interval_seconds=2.0,
        )

    def list_inputs(self) -> tuple[Path, ...]:
        return ()

    def execute(self, command: str):
        self.executed.append(command)
        return None

    def confirm_plan(self, plan_id: str):
        self.confirmed.append(plan_id)
        return plan_id

    def cancel_plan(self, plan_id: str):
        self.cancelled.append(plan_id)
        return plan_id


@pytest.fixture(scope="module", autouse=True)
def qt_gui_application():
    application = QGuiApplication.instance() or QGuiApplication([])
    yield application


def test_stale_confirmation_token_cannot_apply_new_preview(tmp_path) -> None:
    controller = FakeController(tmp_path)
    bridge = DesktopBridge(controller)  # type: ignore[arg-type]
    errors: list[str] = []
    bridge.errorRaised.connect(errors.append)
    bridge._plan_ids = ("plano-a",)
    bridge._plan_revision = "revisao-a"
    token = bridge.prepareConfirmation()

    bridge._plan_ids = ("plano-b",)
    bridge._plan_revision = "revisao-b"
    bridge.confirmPlans(token)

    assert not controller.confirmed
    assert errors == ["A previa mudou. Revise novamente antes de confirmar."]
    bridge.requestClose()


def test_confirmation_token_is_one_use(tmp_path, monkeypatch) -> None:
    controller = FakeController(tmp_path)
    bridge = DesktopBridge(controller)  # type: ignore[arg-type]
    bridge._plan_ids = ("plano-a", "plano-b")
    bridge._plan_revision = "revisao"
    token = bridge.prepareConfirmation()

    monkeypatch.setattr(
        bridge,
        "_submit",
        lambda operation, _callback, _status: operation(),
    )
    bridge.confirmPlans(token)
    bridge.confirmPlans(token)

    assert controller.confirmed == ["plano-a", "plano-b"]
    bridge.requestClose()


def test_failed_plan_batch_cannot_be_confirmed_again(tmp_path, monkeypatch) -> None:
    controller = FakeController(tmp_path)
    bridge = DesktopBridge(controller)  # type: ignore[arg-type]
    bridge._plan_ids = ("plano-a", "plano-b")
    bridge._plan_revision = "revisao"
    token = bridge.prepareConfirmation()
    submitted: list[object] = []
    monkeypatch.setattr(
        bridge,
        "_submit",
        lambda operation, _callback, _status: submitted.append(operation),
    )

    bridge.confirmPlans(token)

    assert len(submitted) == 1
    assert not bridge.canConfirm
    bridge.requestClose()


def test_voice_transcript_never_executes_automatically(tmp_path) -> None:
    controller = FakeController(tmp_path)
    bridge = DesktopBridge(controller)  # type: ignore[arg-type]
    transcripts: list[str] = []
    bridge.voiceTranscriptReady.connect(transcripts.append)

    bridge._after_voice(SimpleNamespace(text="limpar todas"))

    assert transcripts == ["limpar todas"]
    assert not controller.executed
    bridge.requestClose()


def test_busy_state_rejects_second_operation(tmp_path) -> None:
    controller = FakeController(tmp_path)
    bridge = DesktopBridge(controller)  # type: ignore[arg-type]
    bridge._set_busy(True)

    bridge.executeRequest("diagnosticar todas")

    assert not controller.executed
    assert "Aguarde" in bridge.statusText
    bridge._set_busy(False)
    bridge.requestClose()
