# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
import sys

from PyInstaller.utils.hooks import collect_dynamic_libs, copy_metadata

project_root = Path(SPECPATH).resolve().parents[0]
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

icon_path = project_root / "packaging" / "assets" / "capacity_optimizer.ico"
icon_arg = str(icon_path) if icon_path.exists() else None
datas = []
for package_name in ("ortools", "pandas", "openpyxl", "PySide6", "cryptography"):
    try:
        datas += copy_metadata(package_name)
    except Exception:
        pass

hiddenimports = [
    "ortools.sat.python.cp_model",
    "ortools.sat.python.cp_model_helper",
    "PySide6.QtCore",
    "PySide6.QtGui",
    "PySide6.QtWidgets",
    "cryptography.hazmat.primitives.asymmetric.ed25519",
]

binaries = []
for package_name in ("ortools",):
    binaries += collect_dynamic_libs(package_name)

a = Analysis(
    [str(project_root / "FortuneBJOptimizerLauncher.pyw")],
    pathex=[str(project_root)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "matplotlib",
        "scipy",
        "sklearn",
        "pyarrow",
        "pytest",
        "IPython",
        "notebook",
        "jinja2",
        "PIL",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="FortuneBJOptimizer",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    icon=icon_arg,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="FortuneBJOptimizer",
)
