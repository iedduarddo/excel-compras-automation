"""Entrada do aplicativo grafico empacotado."""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

from src.gui.app import run_desktop_app
from src.gui.controller import DesktopController
from src.settings import load_aliases, load_rules


def main(argv: list[str] | None = None) -> int:
    arguments = list(sys.argv[1:] if argv is None else argv)
    if arguments in (["--smoke-test"], ["--smoke-test-qt"]):
        require_qt = arguments == ["--smoke-test-qt"]
        load_aliases()
        load_rules()
        with tempfile.TemporaryDirectory(prefix="Excel Compras GUI ") as temporary:
            root = Path(temporary) / "assistente_planilhas"
            controller = DesktopController(root)
            controller.load_config()
            try:
                from src.gui.qt_app import smoke_test_qml

                return 0 if smoke_test_qml(root) else 1
            except ModuleNotFoundError as error:
                if not _missing_pyside(error):
                    raise
                if require_qt:
                    return 1
                import tkinter as tk

                interpreter = tk.Tcl()
                return 0 if interpreter.eval("info patchlevel") else 1
    if arguments:
        return 2
    return run_desktop_app()


def _missing_pyside(error: ModuleNotFoundError) -> bool:
    return bool(error.name and error.name.split(".", maxsplit=1)[0] == "PySide6")


if __name__ == "__main__":
    raise SystemExit(main())
