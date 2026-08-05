"""Entrada de voz assistida pela Digitacao por Voz do Windows."""

from __future__ import annotations

import base64
import binascii
import json
import os
import re
import subprocess
from dataclasses import dataclass
from typing import Any

from src.core.exceptions import AutomationError

DEFAULT_CULTURE = "pt-BR"
DEFAULT_TIMEOUT_SECONDS = 180
_CULTURE_PATTERN = re.compile(r"^[A-Za-z]{2,3}(?:-[A-Za-z0-9]{2,8})*$")


@dataclass(frozen=True, slots=True)
class VoiceRecognition:
    """Texto revisado pelo usuario e idioma selecionado no Windows."""

    text: str
    confidence: float | None
    culture: str


def recognize_voice(
    *,
    timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS,
    culture: str = DEFAULT_CULTURE,
) -> VoiceRecognition:
    """Abre a Digitacao por Voz e devolve somente o texto confirmado."""

    _validate_options(timeout_seconds, culture)
    if os.name != "nt":
        raise AutomationError("A entrada por voz esta disponivel somente no Windows.")

    script = _build_powershell_script(culture)
    encoded_script = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
    command = (
        "powershell.exe",
        "-NoProfile",
        "-NonInteractive",
        "-Sta",
        "-ExecutionPolicy",
        "Bypass",
        "-EncodedCommand",
        encoded_script,
    )

    try:
        completed = subprocess.run(  # noqa: S603 - comando fixo do Windows
            command,
            capture_output=True,
            text=True,
            encoding="ascii",
            errors="replace",
            timeout=timeout_seconds,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except FileNotFoundError as error:
        raise AutomationError(
            "O Windows PowerShell nao foi encontrado para capturar a voz."
        ) from error
    except subprocess.TimeoutExpired as error:
        raise AutomationError(
            "A confirmacao do comando por voz excedeu o tempo limite."
        ) from error

    if completed.returncode != 0:
        raise AutomationError(_format_voice_error(completed.stderr, culture))

    return _decode_recognition(completed.stdout)


def _validate_options(timeout_seconds: int, culture: str) -> None:
    if not 10 <= timeout_seconds <= 600:
        raise ValueError("timeout_seconds deve estar entre 10 e 600.")
    if not _CULTURE_PATTERN.fullmatch(culture):
        raise ValueError("culture deve usar um identificador como pt-BR.")


def _decode_recognition(value: str) -> VoiceRecognition:
    encoded_payload = value.strip().splitlines()[-1] if value.strip() else ""
    try:
        payload_bytes = base64.b64decode(encoded_payload, validate=True)
        payload: dict[str, Any] = json.loads(payload_bytes.decode("utf-8"))
        text = " ".join(str(payload["text"]).strip().split())
        culture = str(payload["culture"])
    except (binascii.Error, ValueError, KeyError, TypeError) as error:
        raise AutomationError(
            "O Windows retornou uma transcricao de voz invalida."
        ) from error

    if not text:
        raise AutomationError("Nenhuma fala foi confirmada.")
    return VoiceRecognition(text=text, confidence=None, culture=culture)


def _format_voice_error(stderr: str, culture: str) -> str:
    technical = " ".join(stderr.strip().split())
    if "VOICE_CULTURE_NOT_INSTALLED" in technical:
        return (
            f"O idioma de entrada {culture} nao esta instalado no Windows. "
            "Adicione o teclado desse idioma nas Configuracoes."
        )
    if "VOICE_CANCELLED" in technical:
        return "O comando por voz foi cancelado antes da confirmacao."
    if "VOICE_TYPING_UNAVAILABLE" in technical:
        return (
            "A Digitacao por Voz do Windows nao pode ser aberta. "
            "Confirme o acesso ao microfone e tente Win+H."
        )
    return "A entrada por voz do Windows falhou antes da confirmacao."


def _build_powershell_script(culture: str) -> str:
    return _POWERSHELL_SCRIPT.replace("__CULTURE__", culture)


_POWERSHELL_SCRIPT = r"""
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
$form = $null

try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class VoiceTypingKeyboard
{
    [DllImport("user32.dll")]
    private static extern void keybd_event(
        byte virtualKey,
        byte scanCode,
        uint flags,
        UIntPtr extraInfo
    );

    public static void Open()
    {
        const byte leftWindows = 0x5B;
        const byte h = 0x48;
        const uint keyUp = 0x0002;
        keybd_event(leftWindows, 0, 0, UIntPtr.Zero);
        keybd_event(h, 0, 0, UIntPtr.Zero);
        keybd_event(h, 0, keyUp, UIntPtr.Zero);
        keybd_event(leftWindows, 0, keyUp, UIntPtr.Zero);
    }
}
"@

    $inputLanguage = @(
        [System.Windows.Forms.InputLanguage]::InstalledInputLanguages |
            Where-Object { $_.Culture.Name -ieq "__CULTURE__" }
    ) | Select-Object -First 1
    if ($null -eq $inputLanguage) {
        throw "VOICE_CULTURE_NOT_INSTALLED"
    }
    [System.Windows.Forms.InputLanguage]::CurrentInputLanguage = $inputLanguage

    [System.Windows.Forms.Application]::EnableVisualStyles()
    $form = [System.Windows.Forms.Form]::new()
    $form.Text = "Comando por voz - Excel Compras Automation"
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true
    $form.ClientSize = [System.Drawing.Size]::new(560, 245)
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $instructions = [System.Windows.Forms.Label]::new()
    $instructions.Location = [System.Drawing.Point]::new(18, 16)
    $instructions.Size = [System.Drawing.Size]::new(524, 58)
    $instructions.Text = (
        "Fale quando a Digitacao por Voz abrir. " +
        "Revise o texto abaixo e clique em Usar comando."
    )
    $form.Controls.Add($instructions)

    $textBox = [System.Windows.Forms.TextBox]::new()
    $textBox.Location = [System.Drawing.Point]::new(18, 78)
    $textBox.Size = [System.Drawing.Size]::new(524, 92)
    $textBox.Multiline = $true
    $textBox.Font = [System.Drawing.Font]::new("Segoe UI", 12)
    $textBox.ScrollBars = "Vertical"
    $form.Controls.Add($textBox)

    $confirm = [System.Windows.Forms.Button]::new()
    $confirm.Text = "Usar comando"
    $confirm.Location = [System.Drawing.Point]::new(326, 188)
    $confirm.Size = [System.Drawing.Size]::new(104, 32)
    $form.Controls.Add($confirm)
    $form.AcceptButton = $confirm

    $cancel = [System.Windows.Forms.Button]::new()
    $cancel.Text = "Cancelar"
    $cancel.Location = [System.Drawing.Point]::new(438, 188)
    $cancel.Size = [System.Drawing.Size]::new(104, 32)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancel)
    $form.CancelButton = $cancel

    $confirm.Add_Click({
        if ([string]::IsNullOrWhiteSpace($textBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Fale ou digite um comando antes de confirmar.",
                "Comando vazio",
                "OK",
                "Information"
            ) | Out-Null
            return
        }
        $form.Tag = $textBox.Text.Trim()
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    $timer = [System.Windows.Forms.Timer]::new()
    $timer.Interval = 700
    $timer.Add_Tick({
        $timer.Stop()
        $form.Activate()
        $textBox.Focus()
        try {
            [VoiceTypingKeyboard]::Open()
        }
        catch {
            $form.Tag = "VOICE_TYPING_UNAVAILABLE"
            $form.DialogResult = [System.Windows.Forms.DialogResult]::Abort
            $form.Close()
        }
    })
    $form.Add_Shown({ $timer.Start() })

    try {
        $dialogResult = $form.ShowDialog()
    }
    finally {
        $timer.Dispose()
    }

    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Abort) {
        throw "VOICE_TYPING_UNAVAILABLE"
    }
    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        throw "VOICE_CANCELLED"
    }

    $payload = [ordered]@{
        text = [string]$form.Tag
        culture = [string]$inputLanguage.Culture.Name
    } | ConvertTo-Json -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
    [Convert]::ToBase64String($bytes)
}
catch {
    [Console]::Error.WriteLine([string]$_.Exception.Message)
    exit 2
}
finally {
    if ($null -ne $form) {
        $form.Dispose()
    }
}
"""
