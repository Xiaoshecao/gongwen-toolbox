# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：Windows 单文件 exe。

在 `D:\\公文处理` 下执行：
    python -m PyInstaller --noconfirm --clean gwtyper\\build_win.spec
产物：`gwtyper\\dist\\公文排版工具.exe`

要点：
- `rules/` 必须打进去（内置规则集是只读资源，运行时从 sys._MEIPASS 读）；
- tkinterdnd2 自带 tkdnd 的 tcl 扩展，漏了拖拽会失效；
- pywin32 走 hiddenimports（COM 是运行期动态导入的）；
- 排除 pandas/matplotlib 这类用不上的重库，压体积。
"""
import os
from pathlib import Path

ROOT = Path(SPECPATH).resolve().parent   # noqa: F821  SPECPATH 是 spec 所在目录 → 仓库根 D:\公文处理
PKG = ROOT / "gwtyper"
# 图标从 gwtyper/assets 取（以前借 converter/assets，那个目录删掉就静默无图标）
_ASSETS = PKG / "assets"
ICON = _ASSETS / "icon.ico"
_ICON_PNG = _ASSETS / "icon.png"

hidden = [
    "docx",
    "lxml",
    "lxml.etree",
    "yaml",
    "pymupdf",
    "fitz",
    "PIL",
    "PIL.Image",
    "PIL.ImageTk",
    "tkinter",
    "tkinter.ttk",
    "tkinterdnd2",
    "win32com.client",
    "pythoncom",
    "pywintypes",
    # 纯 Python PDF 排版（fontTools 是运行期才 import 的）
    "fontTools",
    "fontTools.subset",
    "fontTools.ttLib",
    "openpyxl",
    "pptx",
    "gwtyper.core.sysenv",
    "gwtyper.core.office",
    "gwtyper.core.appicon",
    "gwtyper.core.ocr",
    "rapidocr_onnxruntime",
    "onnxruntime",
    "gwtyper.core.fontfind",
    "gwtyper.core.pdfwrite",
    "gwtyper.core.convert",
    "gwtyper.core.formats",
    "gwtyper.core.pdf_engine",
]

datas = [(str(PKG / "rules"), "rules")]

# 纯 Python 排版依赖的包只在函数内部 import，静态分析走不到；
# 写成 hiddenimports 若打钩失败只打一行 ERROR 就跳过，结果是「能开、一转 PDF 就崩」。
# 直接按目录整包收进来最稳。
import importlib.util as _ilu

for _pkg in ("fontTools", "openpyxl", "pptx", "rapidocr_onnxruntime"):
    _spec = _ilu.find_spec(_pkg)
    if _spec and _spec.submodule_search_locations:
        datas.append((str(_spec.submodule_search_locations[0]), _pkg))
        print(f"[build_win.spec] 整目录收 {_pkg}")
    else:
        print(f"[build_win.spec] 警告：本机没装 {_pkg}")
if _ASSETS.is_dir():
    datas.append((str(_ASSETS), "assets"))

try:
    import tkinterdnd2 as _dnd

    _tkdnd = Path(_dnd.__file__).parent / "tkdnd"
    if _tkdnd.is_dir():
        datas.append((str(_tkdnd), "tkinterdnd2/tkdnd"))
except Exception:
    pass

a = Analysis(  # noqa: F821
    [str(PKG / "packaging" / "entry.py")],
    pathex=[str(ROOT)],
    binaries=[],
    datas=datas,
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 注意：numpy / cv2 / onnxruntime 是 RapidOCR（双层 PDF）的硬依赖，
    #      以前为了压体积排掉了，结果 OCR 一跑就崩。不能排。
    "matplotlib", "scipy", "pandas",
        "jupyter", "notebook", "IPython", "pytest", "setuptools", "pip", "wheel",
        "PyQt5", "PyQt6", "PySide2", "PySide6", "pdf2docx",
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
    name="公文工具箱",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(ICON) if ICON.exists() else None,
)
