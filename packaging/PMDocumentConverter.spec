# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_submodules


project_root = Path(SPECPATH).parent
src_dir = project_root / "src"

datas = []
datas += collect_data_files("customtkinter")
translations_dir = project_root / "translations"
if translations_dir.exists():
    datas += [(str(translations_dir), "translations")]

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
        "PIL.ImageQt",
        "PyQt5",
        "PyQt6",
        "PySide2",
        "PySide6",
        "Pythonwin",
        "lxml.html",
        "lxml.isoschematron",
        "lxml.objectify",
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
        "numpy.f2py",
        "numpy.fft",
        "numpy.random",
        "numpy.testing",
        "numpy.tests",
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
