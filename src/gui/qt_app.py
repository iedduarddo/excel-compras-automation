"""Inicializador Qt Quick da interface Fluent/Liquid-Glass-inspired."""

from __future__ import annotations

import os
import sys
from pathlib import Path

from src import __version__
from src.gui.controller import DesktopController


def qml_source_path() -> Path:
    """Resolve o QML tanto no codigo-fonte quanto no pacote PyInstaller."""

    bundle_root = Path(getattr(sys, "_MEIPASS", Path(__file__).parents[2]))
    return bundle_root / "src" / "gui" / "qml" / "Main.qml"


def run_qt_desktop_app(root: Path | None = None) -> int:
    """Abre a interface Qt; erros de carga nao sao ocultados pelo fallback Tk."""

    from PySide6.QtCore import QSettings, QUrl
    from PySide6.QtGui import QGuiApplication
    from PySide6.QtQml import QQmlApplicationEngine
    from PySide6.QtQuickControls2 import QQuickStyle

    from src.gui.bridge import DesktopBridge
    from src.gui.native_materials import apply_native_material

    QQuickStyle.setStyle("Basic")
    app = QGuiApplication.instance() or QGuiApplication(sys.argv[:1])
    app.setApplicationName("Excel Compras Automation")
    app.setApplicationVersion(__version__)
    app.setOrganizationName("iedduarddo")

    controller = DesktopController(root)
    bridge = DesktopBridge(controller)
    settings = QSettings()
    initial_dark_mode = settings.value(
        "appearance/darkMode", _prefers_dark_mode(), type=bool
    )
    engine = QQmlApplicationEngine()
    engine.rootContext().setContextProperty("bridge", bridge)
    engine.rootContext().setContextProperty("appVersion", __version__)
    engine.rootContext().setContextProperty("initialDarkMode", initial_dark_mode)
    source = qml_source_path()
    if not source.is_file():
        print(f"Interface QML nao encontrada: {source}", file=sys.stderr)
        return 1
    engine.load(QUrl.fromLocalFile(str(source)))
    roots = engine.rootObjects()
    if not roots:
        print("A interface QML nao criou uma janela.", file=sys.stderr)
        return 1

    window = roots[0]

    def apply_current_theme() -> None:
        dark_mode = bool(window.property("darkMode"))
        settings.setValue("appearance/darkMode", dark_mode)
        apply_native_material(window, dark_mode=dark_mode)

    window.darkModeChanged.connect(apply_current_theme)
    window.show()
    app.processEvents()
    apply_current_theme()
    app.aboutToQuit.connect(bridge.requestClose)
    bridge.initialize()
    return int(app.exec())


def smoke_test_qml(root: Path) -> bool:
    """Carrega a arvore visual sem abrir uma janela, usado no pacote/CI."""

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    os.environ.setdefault("QSG_RHI_BACKEND", "software")
    from PySide6.QtCore import QUrl
    from PySide6.QtGui import QGuiApplication
    from PySide6.QtQml import QQmlApplicationEngine
    from PySide6.QtQuickControls2 import QQuickStyle

    from src.gui.bridge import DesktopBridge

    QQuickStyle.setStyle("Basic")
    app = QGuiApplication.instance() or QGuiApplication(["excel-compras-smoke"])
    bridge = DesktopBridge(DesktopController(root))
    engine = QQmlApplicationEngine()
    qml_warnings: list[object] = []
    engine.warnings.connect(lambda values: qml_warnings.extend(values))
    engine.rootContext().setContextProperty("bridge", bridge)
    engine.rootContext().setContextProperty("appVersion", __version__)
    engine.rootContext().setContextProperty("initialDarkMode", False)
    engine.load(QUrl.fromLocalFile(str(qml_source_path())))
    app.processEvents()
    roots = engine.rootObjects()
    for warning in qml_warnings:
        print(f"Aviso QML: {warning}", file=sys.stderr)
    valid = (
        not qml_warnings
        and len(roots) == 1
        and roots[0].property("objectName") == "mainWindow"
        and not roots[0].isVisible()
    )
    bridge.requestClose()
    return valid


def _prefers_dark_mode() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
        ) as key:
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        return int(value) == 0
    except (OSError, ValueError):
        return False
