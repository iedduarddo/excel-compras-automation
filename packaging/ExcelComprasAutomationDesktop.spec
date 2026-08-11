# -*- mode: python ; coding: utf-8 -*-

import sys
from pathlib import Path


project_root = Path(SPECPATH).parent
hidden_imports = []
if sys.platform == "win32":
    hidden_imports = ["pythoncom", "pywintypes", "win32com", "win32com.client"]

analysis = Analysis(
    [str(project_root / "desktop.py")],
    pathex=[str(project_root)],
    binaries=[],
    datas=[(str(project_root / "config"), "config")],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
python_archive = PYZ(analysis.pure)

executable = EXE(
    python_archive,
    analysis.scripts,
    [],
    exclude_binaries=True,
    name="ExcelComprasAutomation",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    contents_directory="_internal",
)
collection = COLLECT(
    executable,
    analysis.binaries,
    analysis.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="ExcelComprasAutomation",
)

if sys.platform == "darwin":
    app = BUNDLE(
        collection,
        name="Excel Compras Automation.app",
        icon=None,
        bundle_identifier="com.iedduarddo.excel-compras-automation",
        info_plist={"NSHighResolutionCapable": "True"},
    )
