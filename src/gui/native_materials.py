"""Integra materiais nativos sem tornar a interface dependente deles."""

from __future__ import annotations

import ctypes
import sys
from dataclasses import dataclass
from typing import Protocol

DWMWA_USE_IMMERSIVE_DARK_MODE = 20
DWMWA_WINDOW_CORNER_PREFERENCE = 33
DWMWA_SYSTEMBACKDROP_TYPE = 38
DWMWCP_ROUND = 2
DWMSBT_MAINWINDOW = 2
MINIMUM_MICA_BUILD = 22621


class NativeWindow(Protocol):
    """Superficie minima oferecida por ``QQuickWindow``."""

    def winId(self) -> int: ...

    def setProperty(self, name: str, value: object) -> bool: ...


@dataclass(frozen=True, slots=True)
class MaterialResult:
    """Resultado observavel da tentativa de ativar um material do sistema."""

    effective: str
    native: bool
    transparent: bool
    reason: str = ""


def material_for_platform(
    platform_name: str,
    *,
    windows_build: int = 0,
    macos_major: int = 0,
) -> MaterialResult:
    """Seleciona o contrato visual sem chamar APIs de plataforma."""

    if platform_name == "win32" and windows_build >= MINIMUM_MICA_BUILD:
        return MaterialResult("mica", True, True)
    if platform_name == "darwin":
        # O QML usa uma superficie translucida compativel. Liquid Glass nativo
        # exige NSGlassEffectView/AppKit e nao e simulado por este modulo.
        label = "macos-material" if macos_major else "macos-material-fallback"
        return MaterialResult(
            label,
            False,
            False,
            "Liquid Glass nativo requer uma casca AppKit no macOS 26 ou posterior.",
        )
    return MaterialResult(
        "solid",
        False,
        False,
        "O sistema usa o fundo solido acessivel da interface.",
    )


def current_windows_build() -> int:
    """Retorna o build do Windows sem falhar em outras plataformas."""

    if sys.platform != "win32":
        return 0
    return int(sys.getwindowsversion().build)


def apply_native_material(
    window: NativeWindow,
    *,
    platform_name: str | None = None,
    windows_build: int | None = None,
    dark_mode: bool = False,
) -> MaterialResult:
    """Aplica Mica e cantos nativos quando a API documentada esta disponivel.

    A falha nunca impede a abertura da aplicacao: a janela volta ao material
    solido, inclusive em alto contraste, Windows antigo ou sessao remota.
    """

    platform_value = platform_name or sys.platform
    build = current_windows_build() if windows_build is None else windows_build
    selected = material_for_platform(platform_value, windows_build=build)
    if selected.effective != "mica":
        _publish_result(window, selected)
        return selected

    try:
        hwnd = int(window.winId())
        dwmapi = ctypes.WinDLL("dwmapi", use_last_error=True)
        setter = dwmapi.DwmSetWindowAttribute
        setter.argtypes = (
            ctypes.c_void_p,
            ctypes.c_uint,
            ctypes.c_void_p,
            ctypes.c_uint,
        )
        setter.restype = ctypes.c_long
        corner = ctypes.c_int(DWMWCP_ROUND)
        backdrop = ctypes.c_int(DWMSBT_MAINWINDOW)
        dark = ctypes.c_int(int(dark_mode))
        calls = (
            (DWMWA_WINDOW_CORNER_PREFERENCE, corner),
            (DWMWA_SYSTEMBACKDROP_TYPE, backdrop),
            (DWMWA_USE_IMMERSIVE_DARK_MODE, dark),
        )
        for attribute, value in calls:
            result = setter(
                hwnd,
                attribute,
                ctypes.byref(value),
                ctypes.sizeof(value),
            )
            if result != 0:
                raise OSError(f"DwmSetWindowAttribute retornou HRESULT {result}.")
    except (AttributeError, OSError, TypeError, ValueError) as error:
        fallback = MaterialResult("solid", False, False, str(error))
        _publish_result(window, fallback)
        return fallback

    _publish_result(window, selected)
    return selected


def _publish_result(window: NativeWindow, result: MaterialResult) -> None:
    window.setProperty("nativeMaterialActive", result.native)
    window.setProperty("nativeMaterialName", result.effective)
    window.setProperty("nativeMaterialReason", result.reason)
