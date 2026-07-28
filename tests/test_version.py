"""Testes dos metadados de versão do projeto."""

import tomllib
from pathlib import Path

from src import __version__


def test_package_version_matches_pyproject() -> None:
    """A versão importável deve ser igual à versão declarada do projeto."""
    project_root = Path(__file__).resolve().parents[1]
    pyproject = tomllib.loads(
        (project_root / "pyproject.toml").read_text(encoding="utf-8")
    )

    assert pyproject["project"]["version"] == __version__
