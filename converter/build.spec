# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：单文件、无控制台、带图标与版本信息。"""
import os
from pathlib import Path

# PyInstaller 在 spec 命名空间里注入 SPECPATH（spec 所在目录）
ROOT = Path(SPECPATH)  # noqa: F821

# pdf2docx / pymupdf 里有动态导入，显式列出来
hidden = [
    "pdf2docx",
    "pdf2docx.converter",
    "pdf2docx.page",
    "pdf2docx.common",
    "pdf2docx.table",
    "pymupdf",
    "fitz",
    "docx",
    "openpyxl",
    "pptx",
    "PIL",
    "PIL.Image",
    "win32com.client",
    "pythoncom",
    "pywintypes",
    "tkinterdnd2",
    "xml.etree.ElementTree",
    "lxml",
    "lxml.etree",
    "numpy",
    "cv2",
    "onnxruntime",
]

# tkinterdnd2 自带 tcl 扩展，必须一起打包
datas = []
try:
    import tkinterdnd2 as _dnd

    pkg_dir = Path(_dnd.__file__).parent
    for sub in ("tkdnd",):
        src = pkg_dir / sub
        if src.is_dir():
            datas.append((str(src), f"tkinterdnd2/{sub}"))
except Exception:
    pass

a = Analysis(  # noqa: F821
    ["main.py"],
    pathex=[str(ROOT)],
    binaries=[],
    datas=datas,
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "matplotlib",
        "scipy",
        "pandas",
        "jupyter",
        "notebook",
        "IPython",
        "pytest",
        "setuptools",
        "pip",
        "wheel",
        "PyQt5",
        "PySide2",
        "PyQt6",
        "PySide6",
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)  # noqa: F821

exe = EXE(  # noqa: F821
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="格式转换器",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(ROOT / "assets" / "icon.ico"),
    version=str(ROOT / "build" / "version_info.txt"),
)
