# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包：公文工具箱 Qt 版（Windows 单文件 exe）。

在 `D:\\公文处理` 下执行：
    python -m PyInstaller --noconfirm --clean --distpath gongwen-qt/dist --workpath .tmp/pyi-qt gongwen-qt/build_win.spec
产物：`gongwen-qt/dist/公文工具箱.exe`

要点：
- 界面是 PySide6，内核仍是 `gwtyper.core`（纯 Python），两个包都要进去；
- rules/ 与 assets/ 是运行期资源，运行时从 sys._MEIPASS 读（dest 必须叫 rules / assets）；
- rapidocr / onnxruntime / cv2 等是运行期才 import 的，整目录收最稳；
- 排除 tkinter 与一堆用不上的重库压体积。
"""
import importlib.util as _ilu
import os
from pathlib import Path

# 一律用相对路径（cwd = D:\公文处理）：绝对中文路径经 argv 传给 PyInstaller 会被
# 按 ANSI 解码成乱码，报 script 'D:\????\main.py' not found。
SCRIPT = "main.py"   # 相对 spec 所在目录（gongwen-qt）
_ICON_DIR = Path(SPECPATH).resolve().parent
ICON = str(_ICON_DIR / "gwtyper" / "assets" / "icon.ico")

hidden = [
    "docx", "lxml", "lxml.etree", "yaml", "fitz", "pymupdf",
    "PIL", "PIL.Image", "openpyxl", "pptx",
    "win32com.client", "pythoncom", "pywintypes",
    "fontTools", "fontTools.subset", "fontTools.ttLib",
    "rapidocr_onnxruntime", "onnxruntime",
    "PySide6.QtCore", "PySide6.QtGui", "PySide6.QtWidgets",
    "PySide6.QtPdf", "PySide6.QtPdfWidgets",
    "gwtyper.core.sysenv", "gwtyper.core.rules", "gwtyper.core.model",
    "gwtyper.core.parse_docx", "gwtyper.core.classify", "gwtyper.core.check",
    "gwtyper.core.text_check", "gwtyper.core.doc_check", "gwtyper.core.apply",
    "gwtyper.core.report", "gwtyper.core.fonts", "gwtyper.core.fontsizes",
    "gwtyper.core.office", "gwtyper.core.appicon", "gwtyper.core.ocr",
    "gwtyper.core.ocr_text", "gwtyper.core.fontfind", "gwtyper.core.pdfwrite",
    "gwtyper.core.convert", "gwtyper.core.formats", "gwtyper.core.pdf_engine",
    "gwtyper.core.doc_source", "gwtyper.core.wordjump", "gwtyper.core.imgpress",
    "gwtyper.core.dupfind", "gwtyper.core.notice", "gwtyper.core.archive",
    "gwtyper.core.listdiff", "gwtyper.core.watermark", "gwtyper.core.mask",
    "gwtyper.core.merge_tables", "gwtyper.core.printcheck", "gwtyper.core.templates",
    "gwtyper.core.scanim", "gwtyper.core.villages",
    "gwtyper.ui._tools_extra",
]

datas = [
    ("../gwtyper/rules", "rules"),
    ("../gwtyper/assets", "assets"),
    ("assets", "assets"),
    ("VERSION", "."),
]

# 纯 Python 依赖只在函数内部 import，静态分析走不到；整目录收进来最稳
for _pkg in ("fontTools", "openpyxl", "pptx", "rapidocr_onnxruntime",
             "onnxruntime", "cv2", "numpy", "fitz", "pymupdf",
             "docx", "pyclipper", "shapely"):
    _spec = _ilu.find_spec(_pkg)
    if _spec and _spec.submodule_search_locations:
        datas.append((str(_spec.submodule_search_locations[0]), _pkg))
        print(f"[build_win.spec] 整目录收 {_pkg}")
    else:
        print(f"[build_win.spec] 警告：本机没装 {_pkg}")

a = Analysis(
    [SCRIPT],
    pathex=[".", ".."],
    binaries=[],
    datas=datas,
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "tkinterdnd2", "matplotlib", "scipy",
        "pandas", "IPython", "notebook", "pytest", "setuptools",
        "PyQt5", "PyQt6", "PySide2", "wx",
    ],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="公文工具箱",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    icon=ICON if os.path.exists(ICON) else None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="公文工具箱",
)

