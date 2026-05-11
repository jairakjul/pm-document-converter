# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_submodules


project_root = Path(SPECPATH).parent
src_dir = project_root / "src"

datas = []
datas += collect_data_files("customtkinter")

hiddenimports = []
hiddenimports += collect_submodules("customtkinter")


a = Analysis(
    [str(src_dir / "app_gui.py")],
    pathex=[str(src_dir)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={
        "matplotlib": {
            "backends": ["Agg"],
        },
    },
    runtime_hooks=[],
    excludes=[
        "IPython",
        "PyQt5",
        "PyQt6",
        "PySide2",
        "PySide6",
        "matplotlib.backends.backend_gtk3",
        "matplotlib.backends.backend_gtk3agg",
        "matplotlib.backends.backend_gtk4",
        "matplotlib.backends.backend_gtk4agg",
        "matplotlib.backends.backend_macosx",
        "matplotlib.backends.backend_qt",
        "matplotlib.backends.backend_qt5",
        "matplotlib.backends.backend_qt5agg",
        "matplotlib.backends.backend_qtagg",
        "matplotlib.backends.backend_tkagg",
        "matplotlib.backends.backend_wx",
        "matplotlib.backends.backend_wxagg",
        "matplotlib.tests",
        "notebook",
        "pandas",
        "pytest",
        "scipy",
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="PMDocumentConverter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="PMDocumentConverter",
)

app = BUNDLE(
    coll,
    name="PMDocumentConverter.app",
    icon=None,
    bundle_identifier="com.mfec.pm-document-converter",
    info_plist={
        "CFBundleName": "PM Document Converter",
        "CFBundleDisplayName": "PM Document Converter",
        "CFBundleShortVersionString": "0.1.0",
        "CFBundleVersion": "0.1.0",
        "NSHighResolutionCapable": True,
    },
)
