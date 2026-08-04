"""Contratos do empacotador portátil sem executar o PyInstaller."""

from __future__ import annotations

import json
import os
from pathlib import Path
from zipfile import ZipFile

import pytest

from scripts import build_portable as portable
from src.services.diagnostics import run_diagnostics

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def write_version_fixture(root: Path, project: str, package: str) -> None:
    (root / "src").mkdir()
    (root / "pyproject.toml").write_text(
        f'[project]\nname = "fixture"\nversion = "{project}"\n',
        encoding="utf-8",
    )
    (root / "src" / "__init__.py").write_text(
        f'__version__ = "{package}"\n',
        encoding="utf-8",
    )


def make_valid_package(root: Path) -> Path:
    package = root / portable.portable_package_name("1.7.0")
    package.mkdir()
    (package / "_internal").mkdir()
    (package / "_internal" / "pythoncom314.dll").write_bytes(b"pythoncom")
    (package / "_internal" / "pywintypes314.dll").write_bytes(b"pywintypes")
    (package / "config").mkdir()
    (package / "config" / "aliases.json").write_text("{}", encoding="utf-8")
    (package / "config" / "rules.json").write_text("{}", encoding="utf-8")
    (package / "docs").mkdir()
    (package / "docs" / "SOLUCAO_DE_PROBLEMAS.md").write_text(
        "ajuda",
        encoding="utf-8",
    )
    for directory in portable.OPERATING_DIRECTORIES:
        (package / directory).mkdir()
    for filename in (
        "ExcelComprasAutomation.exe",
        "iniciar.cmd",
        "LEIA-ME.txt",
        "run.ps1",
        "README_PORTATIL.md",
        "CHANGELOG.md",
        "VERSAO.txt",
    ):
        (package / filename).write_text(filename, encoding="utf-8")
    return package


def test_read_project_version_requires_synchronized_versions(
    tmp_path: Path,
) -> None:
    write_version_fixture(tmp_path, "1.7.0", "1.7.0")

    assert portable.read_project_version(tmp_path) == "1.7.0"


def test_read_project_version_rejects_mismatch(tmp_path: Path) -> None:
    write_version_fixture(tmp_path, "1.7.0", "1.6.0")

    with pytest.raises(ValueError, match="não estão sincronizadas"):
        portable.read_project_version(tmp_path)


@pytest.mark.parametrize("version", ["1.7", "v1.7.0", "1.7.0-rc1"])
def test_portable_package_name_rejects_non_release_version(version: str) -> None:
    with pytest.raises(ValueError, match="Versão inválida"):
        portable.portable_package_name(version)


def test_version_resource_contains_windows_and_product_versions() -> None:
    resource = portable.render_version_resource("1.7.0")

    assert "filevers=(1, 7, 0, 0)" in resource
    assert "prodvers=(1, 7, 0, 0)" in resource
    assert "ExcelComprasAutomation.exe" in resource
    assert "ProductVersion', '1.7.0'" in resource


def test_commit_requires_full_git_sha() -> None:
    assert portable.normalize_commit("A" * 40) == "a" * 40

    with pytest.raises(ValueError, match="SHA Git"):
        portable.normalize_commit("abc123")


def test_validate_portable_package_accepts_explicit_allowlist(
    tmp_path: Path,
) -> None:
    package = make_valid_package(tmp_path)

    portable.validate_portable_package(package)


def test_validate_portable_package_rejects_spreadsheet(tmp_path: Path) -> None:
    package = make_valid_package(tmp_path)
    (package / "_internal" / "dados.xlsx").write_bytes(b"not a workbook")

    with pytest.raises(ValueError, match="Arquivo proibido"):
        portable.validate_portable_package(package)


def test_validate_portable_package_rejects_unexpected_top_level(
    tmp_path: Path,
) -> None:
    package = make_valid_package(tmp_path)
    (package / "setup.ps1").write_text("unexpected", encoding="utf-8")

    with pytest.raises(ValueError, match="primeiro nível inválido"):
        portable.validate_portable_package(package)


def test_smoke_workbook_passes_read_only_diagnostics(tmp_path: Path) -> None:
    input_file = tmp_path / "entrada_sintetica.xlsx"
    portable.create_smoke_workbook(
        input_file,
        PROJECT_ROOT / "config" / "aliases.json",
    )

    report = run_diagnostics(input_file)

    assert report.ready is True
    assert input_file.is_file()
    assert not list(tmp_path.glob("*.log"))


def test_create_portable_zip_is_stable_and_keeps_empty_directories(
    tmp_path: Path,
) -> None:
    package = make_valid_package(tmp_path)
    first = tmp_path / "first.zip"
    second = tmp_path / "second.zip"

    portable.create_portable_zip(package, first)
    portable.create_portable_zip(package, second)

    assert first.read_bytes() == second.read_bytes()
    with ZipFile(first) as archive:
        names = set(archive.namelist())
    root = f"{package.name}/"
    assert root in names
    for directory in portable.OPERATING_DIRECTORIES:
        assert f"{root}{directory}/" in names


def test_write_sha256_file_uses_portable_filename(tmp_path: Path) -> None:
    archive = tmp_path / "pacote.zip"
    archive.write_bytes(b"conteudo estavel")

    checksum = portable.write_sha256_file(archive)
    digest, filename = checksum.read_text(encoding="ascii").strip().split("  ")

    assert len(digest) == 64
    assert filename == archive.name


@pytest.mark.skipif(os.name != "nt", reason="Contrato específico do cmd.exe.")
def test_windows_batch_command_preserves_paths_and_arguments_with_spaces(
    tmp_path: Path,
) -> None:
    batch_directory = tmp_path / "pacote portátil com espaços"
    batch_directory.mkdir()
    batch_file = batch_directory / "iniciar teste.cmd"
    batch_file.write_text(
        "@echo off\r\n"
        'if not "%~1"=="-NomeCompleto" exit /b 7\r\n'
        'if not "%~2"=="Teste Pacote Portatil" exit /b 8\r\n'
        'if not "%~3"=="-SemPivotNativo" exit /b 9\r\n'
        "echo LAUNCHER_OK\r\n"
        "exit /b 0\r\n",
        encoding="utf-8",
        newline="",
    )
    command_prompt = Path(os.environ.get("COMSPEC", r"C:\Windows\System32\cmd.exe"))

    result = portable._run_command(
        portable._windows_batch_command(
            command_prompt,
            batch_file,
            "-NomeCompleto",
            "Teste Pacote Portatil",
            "-SemPivotNativo",
        ),
        cwd=tmp_path,
        timeout=30,
        capture_output=True,
    )

    assert "LAUNCHER_OK" in result.stdout


def test_pyinstaller_spec_freezes_console_onedir_with_pywin32() -> None:
    spec = (PROJECT_ROOT / "packaging" / "ExcelComprasAutomation.spec").read_text(
        encoding="utf-8"
    )

    assert 'name="ExcelComprasAutomation"' in spec
    assert 'contents_directory="_internal"' in spec
    assert "console=True" in spec
    assert "exclude_binaries=True" in spec
    assert "upx=False" in spec
    for module in ("pythoncom", "pywintypes", "win32com.client"):
        assert json.dumps(module) in spec


@pytest.mark.parametrize(
    "workflow_name",
    ["package-windows.yml", "release-windows.yml"],
)
def test_portable_workflows_use_canonical_windows_builder(
    workflow_name: str,
) -> None:
    workflow = (PROJECT_ROOT / ".github" / "workflows" / workflow_name).read_text(
        encoding="utf-8"
    )

    assert "windows-2025" in workflow
    assert 'python-version: "3.14.6"' in workflow
    assert "python scripts/build_portable.py" in workflow
    assert "dist/*.zip.sha256" in workflow
    assert "actions/upload-artifact@v7" in workflow


def test_release_workflow_only_publishes_a_draft_from_annotated_tag() -> None:
    workflow = (
        PROJECT_ROOT / ".github" / "workflows" / "release-windows.yml"
    ).read_text(encoding="utf-8")

    assert 'tags:\n      - "v*.*.*"' in workflow
    assert 'if ($tagType -ne "tag")' in workflow
    assert 'git rev-parse "$tag^{}"' in workflow
    assert "origin/main" in workflow
    assert "--draft" in workflow
    assert "--verify-tag" in workflow
    assert "--generate-notes" in workflow
    assert "--clobber" not in workflow


def test_release_workflow_finds_and_validates_draft_by_database_id() -> None:
    workflow = (
        PROJECT_ROOT / ".github" / "workflows" / "release-windows.yml"
    ).read_text(encoding="utf-8")

    assert "id: draft_metadata" in workflow
    assert "gh release list" in workflow
    assert "databaseId,tagName,isDraft,isPrerelease" in workflow
    assert '"release_id=$releaseId" >> $env:GITHUB_OUTPUT' in workflow
    assert "releases/$releaseId" in workflow
    assert "releases/tags/$tag" not in workflow
    assert workflow.count("contents: write") == 1
    assert workflow.count("contents: read") == 2
    assert "Ã" not in workflow
