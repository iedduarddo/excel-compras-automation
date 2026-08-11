"""Limites de seguranca da ponte Qt/QML."""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

pytest.importorskip("PySide6")

from PySide6.QtCore import QCoreApplication, QEvent, QPointF, QSize, QUrl
from PySide6.QtGui import QGuiApplication
from PySide6.QtQml import QQmlApplicationEngine
from PySide6.QtQuick import QQuickItem
from PySide6.QtQuickControls2 import QQuickStyle

from src import __version__
from src.gui.bridge import DesktopBridge
from src.gui.qt_app import qml_source_path


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


def test_qml_typography_and_header_respond_to_window_size(tmp_path, request) -> None:
    controller = FakeController(tmp_path)
    bridge = DesktopBridge(controller)  # type: ignore[arg-type]
    QQuickStyle.setStyle("Basic")
    engine = QQmlApplicationEngine()

    def cleanup() -> None:
        engine.deleteLater()
        QCoreApplication.sendPostedEvents(engine, QEvent.Type.DeferredDelete)
        QGuiApplication.processEvents()
        bridge.requestClose()

    request.addfinalizer(cleanup)
    engine.rootContext().setContextProperty("bridge", bridge)
    engine.rootContext().setContextProperty("appVersion", __version__)
    engine.rootContext().setContextProperty("initialDarkMode", False)
    engine.load(QUrl.fromLocalFile(str(qml_source_path())))

    roots = engine.rootObjects()
    assert len(roots) == 1
    window = roots[0]
    title = window.findChild(QQuickItem, "headerTitle")
    theme_button = window.findChild(QQuickItem, "themeToggleButton")
    header = window.findChild(QQuickItem, "headerCard")
    status = window.findChild(QQuickItem, "statusPill")
    assert title is not None
    assert theme_button is not None
    assert header is not None
    assert status is not None

    samples: list[tuple[float, int, int]] = []
    for width, height in ((940, 660), (1280, 820), (1600, 1000)):
        window.resize(QSize(width, height))
        QGuiApplication.processEvents()
        QGuiApplication.processEvents()
        assert window.property("width") == width
        assert window.property("height") == height
        assert not bool(title.property("truncated"))
        samples.append(
            (
                float(window.property("typographyScale")),
                title.property("font").pixelSize(),
                theme_button.property("font").pixelSize(),
            )
        )

        for item in (theme_button, status):
            origin = item.mapToItem(header, QPointF(0, 0))
            assert origin.x() >= -0.5
            assert origin.y() >= -0.5
            assert origin.x() + item.property("width") <= header.property("width") + 0.5
            assert (
                origin.y() + item.property("height") <= header.property("height") + 0.5
            )

    assert samples[0][0] == pytest.approx(0.92)
    assert samples[1][0] == pytest.approx(1.0)
    assert samples[2][0] == pytest.approx(1.18)
    assert samples[0][1] < samples[1][1] < samples[2][1]
    assert samples[0][2] < samples[1][2] < samples[2][2]
    assert samples[0][1] > samples[0][2]
