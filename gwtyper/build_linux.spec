# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 规格文件 —— Linux 自带运行时版（onedir）。

与 Windows 的 build_win.spec 是两套：Windows 只要一个 exe（系统自带 python 不必打），
Linux 的干净机器上往往没有 python3-tk / python-docx / PyMuPDF，所以必须把
解释器和依赖一起塞进 AppImage，否则换台机器就跑不起来。

构建（在 Linux/WSL 里，于 D:\\公文处理 目录执行）：
    python3 -m PyInstaller --noconfirm --clean \
        --distpath gwtyper/build/linux-dist --workpath gwtyper/build/linux-work \
        gwtyper/build_linux.spec
产物：gwtyper/build/linux-dist/公文排版工具/
"""

import os
import sysconfig
from pathlib import Path

# SPECPATH 是 spec 文件所在目录（gwtyper/）
ROOT = Path(SPECPATH).resolve().parent        # -> .../gwtyper
PKG = ROOT / "gwtyper"                        # -> .../gwtyper/gwtyper
TOP = ROOT.parent                             # -> 工作区根目录

datas = [(str(PKG / "rules"), "rules")]

# ---- 把纯 Python 排版依赖的包整目录收进来 -------------------------------
# 这几包（fontTools / openpyxl / pptx）只在函数内部 import，PyInstaller 的
# 静态分析走不到；写成 hiddenimports 时如果打钩失败只会打一行 ERROR 就跳过，
# 最后得到一个「能开、但一转 PDF 就崩」的残包。直接按目录收最稳。
import importlib.util as _ilu

# pymupdf 2.x 的核心在 pymupdf/ 目录（fitz 只是兼容壳）；
# 交叉打包时 x86 侧可能只有旧版 fitz，PyInstaller 就漏了 pymupdf。
for _pkg in ("fontTools", "openpyxl", "pptx", "rapidocr_onnxruntime",
             "pymupdf", "fitz",
             # OCR / 双层 PDF 的硬依赖：模型文件、.so 都在包里，
             # PyInstaller 静态分析看不到 onnxruntime 的 capi 目录，必须整目录收
             "onnxruntime", "cv2", "numpy", "pyclipper", "shapely"):
    _spec = _ilu.find_spec(_pkg)
    if _spec and _spec.submodule_search_locations:
        _root = Path(list(_spec.submodule_search_locations)[0]).parent
        datas.append((str(_spec.submodule_search_locations[0]), _pkg))
        print(f"[build_linux.spec] 整目录收 {_pkg} ← {_spec.submodule_search_locations[0]}")
    else:
        print(f"[build_linux.spec] 警告：本机没装 {_pkg}，纯 Python PDF 会缺功能")

# ---- 显式收集第三方 .so ------------------------------------------------
# 为什么不交给 PyInstaller 自动做：它靠 ldd 解析依赖，而在 qemu+proot 里跑
# ldd 对任何文件都报 "you do not have read permission"，于是 tkinter 的
# _tkinter.so 能进包、它依赖的 libtk8.6/libtcl8.6/libX11 却一个都没进，
# 结果 GUI 一启动就崩。这里按名字清单显式找文件，跟 ldd 无关，两个架构通吃。
_NEEDED_SO = """
libtk8.6.so libtcl8.6.so
libX11.so.6 libXau.so.6 libXdmcp.so.6 libXext.so.6 libXft.so.2
libXrender.so.1 libXss.so.1 libbsd.so.0 libmd.so.0
libfontconfig.so.1 libfreetype.so.6 libexpat.so.1 libpng16.so.16
libglib-2.0.so.0 libharfbuzz.so.0 libgraphite2.so.3 libbrotlicommon.so.1
libbrotlidec.so.1 libpcre2-8.so.0 libbz2.so.1.0 liblzma.so.5 libzstd.so.1
libxml2.so.2 libxslt.so.1 libexslt.so.0 libyaml-0.so.2
libjpeg.so.8 libtiff.so.6 libLerc.so.4 libdeflate.so.0 libjbig.so.0
libjbig2dec.so.0 libwebp.so.7 libwebpdemux.so.2 libwebpmux.so.3
libsharpyuv.so.0 libopenjp2.so.7 liblcms2.so.2 libimagequant.so.0
libgumbo.so.2 libmujs.so.3 libicuuc.so.74 libicudata.so.74
libffi.so.8 libreadline.so.8 libtinfo.so.6 libgomp.so.1
libssl.so.3 libcrypto.so.3 libgcrypt.so.20 libgpg-error.so.0
libgcc_s.so.1 libstdc++.so.6
""".split()

# 也可以直接从外部给一份「依赖库绝对路径清单」（用 scripts/elfdeps.py 生成），
# 这样不同发行版的库版本号（libtiff.so.5 / .so.6、libicuuc.so.63 / .so.74）都不用手改。
_explicit = {}
_so_list_file = os.environ.get("GWTYPER_SO_LIST", "")
if _so_list_file and os.path.exists(_so_list_file):
    with open(_so_list_file, encoding="utf-8") as _f:
        for _line in _f:
            _line = _line.strip()
            if _line and os.path.exists(_line):
                _explicit[os.path.basename(os.path.realpath(_line))] = os.path.realpath(_line)
    _NEEDED_SO = sorted(set(list(_NEEDED_SO) + list(_explicit)))
    print("[build_linux.spec] 外部依赖清单 %s：%d 个库" % (_so_list_file, len(_explicit)))

_multiarch = sysconfig.get_config_var("MULTIARCH") or ""
_libdirs = [
    f"/usr/lib/{_multiarch}", f"/lib/{_multiarch}",
    "/usr/lib64", "/lib64", "/usr/lib", "/lib", "/usr/local/lib",
]


def _find_so(name: str):
    for d in _libdirs:
        if not d:
            continue
        p = os.path.join(d, name)
        if os.path.exists(p):
            return os.path.realpath(p)   # 解软链，否则 PyInstaller 只能拷到个链接
    return None


binaries = []
_missing = []
for _n in _NEEDED_SO:
    _p = _explicit.get(_n) or _find_so(_n)
    if _p:
        binaries.append((_p, "."))
    else:
        _missing.append(_n)
if _missing:
    print("[build_linux.spec] 本机缺这些库，跳过：", ", ".join(_missing))
print(f"[build_linux.spec] 显式收集 {len(binaries)} 个 .so")

# 图标：用 gwtyper 自己的 assets（含 ico 和 png），
# 运行期 appicon 从 resource_dir()/assets 里找
_icon_dir = PKG / "assets"
if _icon_dir.is_dir():
    datas.append((str(_icon_dir), "assets"))

# tkinterdnd2 是可选依赖：装了就打进去，没装就算了（程序里 _HAS_DND 会退化）
try:
    import tkinterdnd2
    _tkdnd = Path(tkinterdnd2.__file__).parent / "tkdnd"
    if _tkdnd.exists():
        datas.append((str(_tkdnd), "tkinterdnd2/tkdnd"))
except Exception:
    pass

hiddenimports = [
    "docx", "lxml", "lxml.etree", "yaml",
    "fitz", "pymupdf", "PIL", "PIL.ImageTk",
    "tkinter", "tkinter.ttk", "tkinter.font", "tkinter.filedialog",
    "tkinter.messagebox", "tkinter.colorchooser",
    "tkinterdnd2",
    # 纯 Python PDF 排版：fontTools 是运行期才 import 的，不写进来会被漏掉
    "fontTools", "fontTools.subset", "fontTools.ttLib",
    "openpyxl", "pptx",
    "gwtyper.core.sysenv", "gwtyper.core.office", "gwtyper.core.pdfmod",
    "gwtyper.core.fonts", "gwtyper.core.rules",
    "gwtyper.core.appicon",
    "gwtyper.core.ocr",
    "rapidocr_onnxruntime",
    "onnxruntime",
    "gwtyper.core.fontfind", "gwtyper.core.pdfwrite",
    "gwtyper.core.lo_bundle",
    "gwtyper.core.convert", "gwtyper.core.formats", "gwtyper.core.pdf_engine",
]

excludes = [
    # 注意：numpy / cv2 / onnxruntime 是 RapidOCR（双层 PDF）的硬依赖，
    #      以前为了压体积排掉了，结果 OCR 一跑就崩。不能排。
    "matplotlib", "scipy", "pandas",
    "pdf2docx", "IPython", "notebook", "pytest", "setuptools",
    "PyQt5", "PySide2", "PySide6", "wx",
]

a = Analysis(
    [str(PKG / "packaging" / "entry.py")],
    pathex=[str(TOP)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="gwtyper",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
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
    upx=False,
    upx_exclude=[],
    name="gwtyper",
)
