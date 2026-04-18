# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs, collect_submodules

project_root = Path.cwd().resolve()
assets_dir = project_root / "assets"
asset_files = [(str(path), "assets") for path in assets_dir.iterdir() if path.is_file()]

try:
    ocr_datas = collect_data_files("rapidocr_onnxruntime")
    onnxruntime_datas = collect_data_files("onnxruntime")
    ocr_binaries = collect_dynamic_libs("onnxruntime")
    ocr_hiddenimports = collect_submodules("rapidocr_onnxruntime")
except Exception:
    ocr_datas = []
    onnxruntime_datas = []
    ocr_binaries = []
    ocr_hiddenimports = []


a = Analysis(
    ["clipboard_manager.py"],
    pathex=[str(project_root)],
    binaries=ocr_binaries,
    datas=asset_files + ocr_datas + onnxruntime_datas,
    hiddenimports=[
        "platforms.macos.services",
        "AppKit",
        "Foundation",
        "rumps",
        "onnxruntime",
        "rapidocr_onnxruntime",
        *ocr_hiddenimports,
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
