"""Contratos de interação da interface gráfica."""

from types import SimpleNamespace

from src.assistant.commands import parse_command
from src.assistant.service import AssistantResult
from src.gui.app import (
    DesktopApp,
    calculate_window_geometry,
    platform_description,
)


def test_window_geometry_is_centered_and_bounded() -> None:
    assert calculate_window_geometry(1920, 1080) == "1180x780+370+150"
    assert calculate_window_geometry(800, 600) == "760x560+20+20"


def test_platform_description_is_readable() -> None:
    assert platform_description("win32") == "Windows"
    assert platform_description("darwin") == "macOS"
    assert platform_description("linux") == "Desktop"


def test_preview_controls_only_the_visible_plans() -> None:
    app = DesktopApp.__new__(DesktopApp)
    result = AssistantResult(parse_command("reconhecer todas"), ())
    toggles: list[bool] = []
    previews: list[str] = []
    statuses: list[tuple[str, str]] = []
    refreshes: list[bool] = []
    app.controller = SimpleNamespace(
        read_previews=lambda _result: (("abc123def456",), "# Prévia segura")
    )
    app.current_plan_ids = ()
    app._toggle_plan_buttons = toggles.append
    app._set_preview = previews.append
    app._set_status = lambda message, tone="normal": statuses.append((message, tone))
    app.refresh_inputs = lambda *, announce=True: refreshes.append(announce)

    app._show_result(result)

    assert app.current_plan_ids == ("abc123def456",)
    assert toggles == [True]
    assert previews == ["# Prévia segura"]
    assert refreshes == [False]
    assert statuses == [
        ("Prévia pronta: revise e confirme para gerar as cópias", "warning")
    ]


def test_confirmation_requires_yes_and_snapshots_visible_ids(monkeypatch) -> None:
    from tkinter import messagebox

    app = DesktopApp.__new__(DesktopApp)
    app.current_plan_ids = ("abc123def456", "def456abc123")
    app.root = object()
    confirmed: list[str] = []
    submitted: list[tuple[object, object, str]] = []
    app.controller = SimpleNamespace(
        confirm_plan=lambda plan_id: confirmed.append(plan_id) or plan_id
    )
    app._show_plan_results = lambda _value: None
    app._submit = lambda operation, callback, status: submitted.append(
        (operation, callback, status)
    )

    monkeypatch.setattr(messagebox, "askyesno", lambda *args, **kwargs: False)
    app.confirm_plan()
    assert not submitted
    assert not confirmed

    monkeypatch.setattr(messagebox, "askyesno", lambda *args, **kwargs: True)
    app.confirm_plan()
    assert len(submitted) == 1
    operation = submitted[0][0]
    app.current_plan_ids = ("outro-plano",)

    assert operation() == ("abc123def456", "def456abc123")
    assert confirmed == ["abc123def456", "def456abc123"]
