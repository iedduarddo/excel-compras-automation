"""Contratos do empacotamento gráfico multiplataforma."""

from pathlib import Path

import pytest

import desktop
from scripts import build_desktop

PROJECT_ROOT = Path(__file__).resolve().parents[1]


@pytest.mark.parametrize(
    ("system", "machine", "expected"),
    [
        ("Windows", "AMD64", "windows-x64"),
        ("Darwin", "x86_64", "macos-intel"),
        ("Darwin", "arm64", "macos-apple-silicon"),
    ],
)
def test_detect_target_maps_supported_native_builds(
    system: str,
    machine: str,
    expected: str,
) -> None:
    assert build_desktop.detect_target(system, machine) == expected


def test_detect_target_rejects_unsupported_platform() -> None:
    with pytest.raises(RuntimeError, match="não suportada"):
        build_desktop.detect_target("Linux", "x86_64")


def test_desktop_package_name_contains_version_architecture_and_kind() -> None:
    assert build_desktop.desktop_package_name("1.12.0", "macos-intel") == (
        "ExcelComprasAutomation-v1.12.0-macos-intel-desktop"
    )


def test_desktop_spec_is_windowed_and_platform_aware() -> None:
    spec = (
        PROJECT_ROOT / "packaging" / "ExcelComprasAutomationDesktop.spec"
    ).read_text(encoding="utf-8")

    assert 'project_root / "desktop.py"' in spec
    assert "console=False" in spec
    assert 'sys.platform == "win32"' in spec
    assert "win32com.client" in spec
    assert 'sys.platform == "darwin"' in spec
    assert "BUNDLE(" in spec
    assert '"config"' in spec


def test_desktop_workflow_builds_three_native_artifacts() -> None:
    workflow = (
        PROJECT_ROOT / ".github" / "workflows" / "package-desktop.yml"
    ).read_text(encoding="utf-8")

    for token in (
        "windows-2025",
        "macos-15-intel",
        "macos-15",
        "windows-x64",
        "macos-intel",
        "macos-apple-silicon",
        "scripts/build_desktop.py",
        "import tkinter",
    ):
        assert token in workflow


def test_release_workflow_promotes_all_desktop_artifacts() -> None:
    workflow = (
        PROJECT_ROOT / ".github" / "workflows" / "release-windows.yml"
    ).read_text(encoding="utf-8")

    assert "build-desktop:" in workflow
    assert "- build-desktop" in workflow
    assert "pattern: release-desktop-*-${{ github.ref_name }}" in workflow
    for target in (
        "windows-x64-desktop",
        "macos-intel-desktop",
        "macos-apple-silicon-desktop",
    ):
        assert target in workflow


def test_desktop_smoke_entrypoint_uses_no_window() -> None:
    assert desktop.main(["--smoke-test"]) == 0
    assert desktop.main(["--unknown"]) == 2
