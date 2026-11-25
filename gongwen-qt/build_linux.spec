# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包：公文工具箱 Qt 版（Linux onedir，自带运行时）。

在 `D:\\公文处理` 下执行（Linux/WSL 里路径为 /mnt/d/公文处理）：
    python3 -m PyInstaller --noconfirm --clean \
        --distpath gongwen-qt/dist-linux --workpath /tmp/pyi-qt-linux \
        gongwen-qt/build_linux.spec
产物：gongwen-qt/dist-linux/gwtyper-qt/（onedir），再打成 AppImage / tar.gz。

要点：
- PySide6 的 Qt 库由 PyInstaller 的 hook 自动收；
- 中文字体不打包（用系统字体，缺字体时界面会提示）；需要时可用 GWTYPER_FONTS 旁挂；
- rapidocr/onnxruntime/cv2 整目录收。
"""
import importlib.util as _ilu
import os
import sysconfig
from pathlib import Path

SCRIPT = "main.py"

hidden = [
    "docx", "lxml", "lxml.etree", "yaml", "fitz", "pymupdf",
    "PIL", "PIL.Image", "openpyxl", "pptx",
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

for _pkg in ("fontTools", "openpyxl", "pptx", "rapidocr_onnxruntime",
             "onnxruntime", "cv2", "numpy", "fitz", "pymupdf",
             "docx", "pyclipper", "shapely"):
    _spec = _ilu.find_spec(_pkg)
    if _spec and _spec.submodule_search_locations:
        datas.append((str(_spec.submodule_search_locations[0]), _pkg))
        print(f"[build_linux.spec] 整目录收 {_pkg}")
    else:
        print(f"[build_linux.spec] 警告：没装 {_pkg}")

# 显式收集 Qt 依赖的系统库（AppImage 里要在没有 Qt 的机器上也能跑）
import glob
import os as _os

_MULTI = sysconfig.get_config_var("MULTIARCH") or ""
_LIBDIRS = [f"/usr/lib/{_MULTI}", f"/lib/{_MULTI}", "/usr/lib64", "/lib64",
            "/usr/lib", "/lib", "/usr/local/lib"]
_NEEDED = """
libxcb.so.1 libxcb-cursor.so.0 libxcb-icccm.so.4 libxcb-image.so.0
libxcb-keysyms.so.1 libxcb-randr.so.0 libxcb-render-util.so.0
libxcb-render.so.0 libxcb-shape.so.0 libxcb-shm.so.0 libxcb-sync.so.1
libxcb-xfixes.so.0 libxcb-xinerama.so.0 libxcb-xkb.so.1 libxkbcommon.so.0
libxkbcommon-x11.so.0 libX11.so.6 libX11-xcb.so.1 libXau.so.6 libXdmcp.so.6
libXext.so.6 libXrender.so.1 libXi.so.6 libXfixes.so.3 libXcursor.so.1
libXrandr.so.2 libSM.so.6 libICE.so.6 libfontconfig.so.1 libfreetype.so.6
libexpat.so.1 libz.so.1 libbz2.so.1.0 liblzma.so.5 libpng16.so.16
libglib-2.0.so.0 libgcc_s.so.1 libstdc++.so.6 libGL.so.1 libEGL.so.1
libGLX.so.0 libGLdispatch.so.0 libdbus-1.so.3 libb2.so.1 libmd4c.so.0
libgomp.so.1 libharfbuzz.so.0 libgraphite2.so.3 libglib-2.0.so.0
""".split()


def _find_so(name):
    for d in _LIBDIRS:
        p = _os.path.join(d, name)
        if _os.path.exists(p):
            return _os.path.realpath(p)
    return None


binaries = []
_missing = []
for _n in _NEEDED:
    _p = _find_so(_n)
    if _p:
        binaries.append((_p, "."))
    else:
        _missing.append(_n)
# PySide6 / shiboken6 的共享库与 Qt 库：PyInstaller 的 hook 在 aarch64 + --target 安装时收不全，这里显式补
for _mod in ("PySide6", "shiboken6"):
    _spec = _ilu.find_spec(_mod)
    if _spec and _spec.submodule_search_locations:
        _d = Path(_spec.submodule_search_locations[0])
        for _so in list(_d.glob("*.so*")) + list(_d.glob("Qt/lib/*.so*")):
            _dest = "PySide6/Qt/lib" if _so.parent.name == "lib" else "."
            binaries.append((str(_so.resolve()), _dest))
print(f"[build_linux.spec] 显式收集 {len(binaries)} 个 .so；缺 {_missing}")

a = Analysis(
    [SCRIPT],
    pathex=[".", ".."],
    binaries=binaries,
    datas=datas,
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinterdnd2", "matplotlib", "scipy", "pandas",
              "IPython", "notebook", "pytest", "PyQt5", "PyQt6", "PySide2", "wx"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="gwtyper-qt",
    debug=False,
    strip=False,
    upx=False,
    console=False,
    icon="assets/icon.png" if _os.path.exists("assets/icon.png") else None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="gwtyper-qt",
)
