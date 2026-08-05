"""Testes da entrada pela Digitacao por Voz do Windows."""

from __future__ import annotations

import base64
import json
import subprocess

import pytest

import src.assistant.voice as voice_module
from src.assistant.voice import VoiceRecognition, recognize_voice
from src.core.exceptions import AutomationError


def encoded_result(
    text: str = "reconhecer todas",
    *,
    culture: str = "pt-BR",
) -> str:
    payload = json.dumps(
        {"text": text, "culture": culture},
        ensure_ascii=False,
    ).encode("utf-8")
    return base64.b64encode(payload).decode("ascii") + "\n"


def test_recognize_voice_opens_windows_voice_typing_for_confirmation(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    received: dict[str, object] = {}

    def run(command: tuple[str, ...], **kwargs: object) -> subprocess.CompletedProcess:
        received["command"] = command
        received["kwargs"] = kwargs
        return subprocess.CompletedProcess(command, 0, encoded_result(), "")

    monkeypatch.setattr(voice_module.subprocess, "run", run)

    result = recognize_voice(timeout_seconds=90)

    command = received["command"]
    assert isinstance(command, tuple)
    assert command[:6] == (
        "powershell.exe",
        "-NoProfile",
        "-NonInteractive",
        "-Sta",
        "-ExecutionPolicy",
        "Bypass",
    )
    script = base64.b64decode(command[-1]).decode("utf-16-le")
    assert "System.Windows.Forms" in script
    assert "VoiceTypingKeyboard" in script
    assert "keybd_event" in script
    assert 'Culture.Name -ieq "pt-BR"' in script
    assert "Usar comando" in script
    assert received["kwargs"] == {
        "capture_output": True,
        "text": True,
        "encoding": "ascii",
        "errors": "replace",
        "timeout": 90,
        "check": False,
        "creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0),
    }
    assert result == VoiceRecognition("reconhecer todas", None, "pt-BR")


@pytest.mark.parametrize(
    ("technical", "expected"),
    [
        ("VOICE_CULTURE_NOT_INSTALLED", "pt-BR nao esta instalado"),
        ("VOICE_CANCELLED", "cancelado"),
        ("VOICE_TYPING_UNAVAILABLE", "Win\\+H"),
        ("unexpected", "falhou"),
    ],
)
def test_recognize_voice_translates_windows_errors(
    technical: str,
    expected: str,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        voice_module.subprocess,
        "run",
        lambda *_, **__: subprocess.CompletedProcess([], 2, "", technical),
    )

    with pytest.raises(AutomationError, match=expected):
        recognize_voice()


def test_recognize_voice_rejects_malformed_payload(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        voice_module.subprocess,
        "run",
        lambda *_, **__: subprocess.CompletedProcess([], 0, "not-base64", ""),
    )

    with pytest.raises(AutomationError, match="transcricao de voz invalida"):
        recognize_voice()


@pytest.mark.parametrize(
    "options",
    [
        {"timeout_seconds": 9},
        {"timeout_seconds": 601},
        {"culture": "pt BR"},
    ],
)
def test_recognize_voice_validates_options(options: dict[str, object]) -> None:
    with pytest.raises(ValueError):
        recognize_voice(**options)  # type: ignore[arg-type]


def test_recognize_voice_reports_missing_powershell(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def fail(*_: object, **__: object) -> None:
        raise FileNotFoundError

    monkeypatch.setattr(voice_module.subprocess, "run", fail)

    with pytest.raises(AutomationError, match="PowerShell nao foi encontrado"):
        recognize_voice()


def test_recognize_voice_reports_confirmation_timeout(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def fail(*_: object, **__: object) -> None:
        raise subprocess.TimeoutExpired("powershell.exe", 180)

    monkeypatch.setattr(voice_module.subprocess, "run", fail)

    with pytest.raises(AutomationError, match="excedeu o tempo limite"):
        recognize_voice()


def test_recognize_voice_requires_windows(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(voice_module.os, "name", "posix")

    with pytest.raises(AutomationError, match="somente no Windows"):
        recognize_voice()
