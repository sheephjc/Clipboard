# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

project_root = Path.cwd().resolve()
assets_dir = project_root / "assets"
asset_files = [(str(path), "assets") for path in assets_dir.iterdir() if path.is_file()]


a = Analysis(
    ["clipboard_manager.py"],
    pathex=[str(project_root)],
    binaries=[],
    datas=asset_files,
    hiddenimports=[
        "platforms.macos.services",
        "AppKit",
        "Foundation",
        "rumps",
    ],
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
    upx=False,
    console=False,
)

app = BUNDLE(
    exe,
    name="Clipboard.app",
    bundle_identifier="com.hjc.clipboard",
    info_plist={
        "CFBundleName": "Clipboard",
        "CFBundleDisplayName": "Clipboard",
        "NSHighResolutionCapable": "True",
    },
)

