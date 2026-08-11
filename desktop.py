"""Entrada do aplicativo gráfico empacotado."""

from __future__ import annotations

import sys
import tempfile
import tkinter as tk
from pathlib import Path

from src.gui.app import run_desktop_app
from src.gui.controller import DesktopController
from src.settings import load_aliases, load_rules


def main(argv: list[str] | None = None) -> int:
    arguments = list(sys.argv[1:] if argv is None else argv)
    if arguments == ["--smoke-test"]:
        interpreter = tk.Tcl()
        if not interpreter.eval("info patchlevel"):
            return 1
        load_aliases()
        load_rules()
        with tempfile.TemporaryDirectory(prefix="Excel Compras GUI ") as temporary:
            controller = DesktopController(Path(temporary) / "assistente_planilhas")
            controller.load_config()
        return 0
    if arguments:
        return 2
    return run_desktop_app()


if __name__ == "__main__":
    raise SystemExit(main())
