"""Constrói e valida os pacotes gráficos nativos para Windows e macOS."""

from __future__ import annotations

import argparse
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path

if __package__:
    from scripts.build_portable import (
        create_portable_zip,
        normalize_commit,
        read_project_version,
        write_sha256_file,
    )
else:
    from build_portable import (  # type: ignore[import-not-found]
        create_portable_zip,
        normalize_commit,
        read_project_version,
        write_sha256_file,
    )

APP_NAME = "ExcelComprasAutomation"
MAC_APP_NAME = "Excel Compras Automation.app"
SUPPORTED_TARGETS = {
    ("Windows", "amd64"): "windows-x64",
    ("Windows", "x86_64"): "windows-x64",
    ("Darwin", "x86_64"): "macos-intel",
    ("Darwin", "arm64"): "macos-apple-silicon",
}


def detect_target(
    system_name: str | None = None,
    machine_name: str | None = None,
) -> str:
    key = (
        system_name or platform.system(),
        (machine_name or platform.machine()).casefold(),
    )
    try:
        return SUPPORTED_TARGETS[key]
    except KeyError as error:
        raise RuntimeError(
            f"Plataforma gráfica não suportada: sistema={key[0]}; arquitetura={key[1]}."
        ) from error


def desktop_package_name(version: str, target: str) -> str:
    return f"{APP_NAME}-v{version}-{target}-desktop"


def _run(command: list[str], *, cwd: Path, timeout: int = 300) -> None:
    completed = subprocess.run(  # noqa: S603
        command,
        cwd=cwd,
        check=False,
        timeout=timeout,
    )
    if completed.returncode != 0:
        raise RuntimeError(
            f"O comando terminou com código {completed.returncode}: {' '.join(command)}"
        )


def _remove_child(path: Path, *, parent: Path) -> None:
    resolved = path.resolve()
    expected_parent = parent.resolve()
    if resolved == expected_parent or resolved.parent != expected_parent:
        raise ValueError(f"Caminho inseguro para limpeza: {resolved}")
    if resolved.is_dir():
        shutil.rmtree(resolved)
    elif resolved.exists():
        resolved.unlink()


def build_application(project_root: Path, scratch: Path, target: str) -> Path:
    distribution = scratch / "pyinstaller-dist"
    work = scratch / "pyinstaller-work"
    spec = project_root / "packaging" / f"{APP_NAME}Desktop.spec"
    _run(
        [
            sys.executable,
            "-m",
            "PyInstaller",
            "--noconfirm",
            "--clean",
            "--distpath",
            str(distribution),
            "--workpath",
            str(work),
            str(spec),
        ],
        cwd=project_root,
        timeout=900,
    )
    artifact = (
        distribution / MAC_APP_NAME
        if target.startswith("macos-")
        else distribution / APP_NAME
    )
    if not artifact.exists():
        raise RuntimeError(f"O PyInstaller não criou o aplicativo esperado: {artifact}")
    return artifact


def _application_executable(artifact: Path, target: str) -> Path:
    if target.startswith("macos-"):
        return artifact / "Contents" / "MacOS" / APP_NAME
    return artifact / f"{APP_NAME}.exe"


def smoke_test_application(artifact: Path, target: str) -> None:
    executable = _application_executable(artifact, target)
    if not executable.is_file():
        raise RuntimeError(f"Executável gráfico ausente: {executable}")
    _run([str(executable), "--smoke-test"], cwd=artifact.parent, timeout=120)


def assemble_package(
    *,
    project_root: Path,
    artifact: Path,
    package: Path,
    version: str,
    target: str,
    commit: str | None,
) -> None:
    package.mkdir(parents=True)
    if target.startswith("macos-"):
        shutil.copytree(artifact, package / artifact.name, symlinks=True)
    else:
        for item in artifact.iterdir():
            destination = package / item.name
            if item.is_dir():
                shutil.copytree(item, destination)
            else:
                shutil.copy2(item, destination)

    shutil.copytree(project_root / "config", package / "config")
    for source_name, destination_name in (
        ("README_DESKTOP.md", "README_DESKTOP.md"),
        ("LEIA-ME-DESKTOP.txt", "LEIA-ME.txt"),
    ):
        content = (project_root / "packaging" / source_name).read_text(encoding="utf-8")
        (package / destination_name).write_text(
            content.replace("vX.Y.Z", f"v{version}"),
            encoding="utf-8",
            newline="\n",
        )

    commit_value = normalize_commit(commit) or "não informado"
    (package / "VERSAO.txt").write_text(
        (
            f"Excel Compras Automation {version}\n"
            f"Plataforma: {target}\n"
            f"Commit: {commit_value}\n"
        ),
        encoding="utf-8",
        newline="\n",
    )


def validate_package(package: Path, target: str) -> None:
    common = {"config", "LEIA-ME.txt", "README_DESKTOP.md", "VERSAO.txt"}
    expected = (
        common | {MAC_APP_NAME}
        if target.startswith("macos-")
        else common | {"_internal", f"{APP_NAME}.exe"}
    )
    actual = {item.name for item in package.iterdir()}
    if actual != expected:
        raise ValueError(
            f"Conteúdo do pacote gráfico inválido: "
            f"ausentes={sorted(expected - actual)}; extras={sorted(actual - expected)}."
        )
    if {item.name for item in (package / "config").iterdir()} != {
        "aliases.json",
        "rules.json",
    }:
        raise ValueError("A configuração externa do pacote está incompleta.")
    smoke_test_application(
        package / MAC_APP_NAME if target.startswith("macos-") else package,
        target,
    )


def create_archive(package: Path, archive: Path, target: str) -> None:
    if target.startswith("macos-"):
        _run(
            [
                "/usr/bin/ditto",
                "-c",
                "-k",
                "--sequesterRsrc",
                "--keepParent",
                str(package),
                str(archive),
            ],
            cwd=package.parent,
        )
    else:
        create_portable_zip(package, archive)


def _parse_arguments(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--expected-version", default=None)
    parser.add_argument("--commit", default=None)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    arguments = _parse_arguments(argv)
    project_root = Path(__file__).resolve().parents[1]
    target = detect_target()
    version = read_project_version(project_root)
    if arguments.expected_version and version != arguments.expected_version:
        raise ValueError(
            f"Versão inesperada: projeto={version}; esperada={arguments.expected_version}."
        )

    build_root = project_root / "build"
    dist_root = project_root / "dist"
    build_root.mkdir(exist_ok=True)
    dist_root.mkdir(exist_ok=True)
    scratch = build_root / f"desktop-{target}"
    name = desktop_package_name(version, target)
    package = dist_root / name
    archive = dist_root / f"{name}.zip"
    checksum = archive.with_name(f"{archive.name}.sha256")
    for path, parent in (
        (scratch, build_root),
        (package, dist_root),
        (archive, dist_root),
        (checksum, dist_root),
    ):
        _remove_child(path, parent=parent)
    scratch.mkdir()

    artifact = build_application(project_root, scratch, target)
    smoke_test_application(artifact, target)
    assemble_package(
        project_root=project_root,
        artifact=artifact,
        package=package,
        version=version,
        target=target,
        commit=arguments.commit or os.environ.get("GITHUB_SHA"),
    )
    validate_package(package, target)
    create_archive(package, archive, target)
    checksum = write_sha256_file(archive)
    print(f"Aplicativo: {archive}")
    print(f"SHA-256: {checksum}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
