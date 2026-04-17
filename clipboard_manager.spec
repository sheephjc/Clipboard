# -*- mode: python ; coding: utf-8 -*-
import os
from pathlib import Path

project_root = Path(os.getcwd()).resolve()
assets_dir = project_root / "assets"
asset_files = [(str(path), "assets") for path in assets_dir.iterdir() if path.is_file()]


a = Analysis(
    ["clipboard_manager.py"],
    pathex=[str(project_root)],
    binaries=[],
    datas=asset_files,
    hiddenimports=["win32timezone"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="Clipboard",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    icon=str(assets_dir / "clipboard_icon.ico"),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    name="Clipboard 1.0.0",
)
