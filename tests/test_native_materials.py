"""Contratos dos materiais nativos e seus fallbacks."""

from types import SimpleNamespace

from src.gui.native_materials import (
    MINIMUM_MICA_BUILD,
    apply_native_material,
    material_for_platform,
)


class FakeWindow:
    def __init__(self) -> None:
        self.properties: dict[str, object] = {}

    def winId(self) -> int:  # noqa: N802 - contrato do Qt
        return 123

    def setProperty(self, name: str, value: object) -> bool:  # noqa: N802
        self.properties[name] = value
        return True


def test_material_selection_is_explicit_about_native_support() -> None:
    assert (
        material_for_platform("win32", windows_build=MINIMUM_MICA_BUILD).effective
        == "mica"
    )
    assert not material_for_platform(
        "win32", windows_build=MINIMUM_MICA_BUILD - 1
    ).native

    macos = material_for_platform("darwin", macos_major=26)
    assert macos.effective == "macos-material"
    assert not macos.native
    assert "AppKit" in macos.reason


def test_mica_uses_documented_dwm_attributes(monkeypatch) -> None:
    calls: list[tuple[int, int]] = []

    class Setter:
        argtypes = None
        restype = None

        def __call__(self, _hwnd, attribute, _value, size) -> int:
            calls.append((attribute, size))
            return 0

    setter = Setter()
    monkeypatch.setattr(
        "src.gui.native_materials.ctypes.WinDLL",
        lambda *_args, **_kwargs: SimpleNamespace(DwmSetWindowAttribute=setter),
        raising=False,
    )
    window = FakeWindow()

    result = apply_native_material(
        window,
        platform_name="win32",
        windows_build=MINIMUM_MICA_BUILD,
        dark_mode=True,
    )

    assert result.native
    assert [attribute for attribute, _ in calls] == [33, 38, 20]
    assert window.properties["nativeMaterialActive"] is True
    assert window.properties["nativeMaterialName"] == "mica"


def test_native_api_failure_falls_back_without_breaking_window(monkeypatch) -> None:
    def fail(*_args, **_kwargs):
        raise OSError("DWM indisponivel")

    monkeypatch.setattr("src.gui.native_materials.ctypes.WinDLL", fail, raising=False)
    window = FakeWindow()

    result = apply_native_material(
        window,
        platform_name="win32",
        windows_build=MINIMUM_MICA_BUILD,
    )

    assert result.effective == "solid"
    assert not result.native
    assert window.properties["nativeMaterialActive"] is False
