# -*- coding: utf-8 -*-
"""PyMuPDF 的导入兼容层。

不同发行版给的模块名不一样：pip 装的 PyMuPDF（>=1.24）能用 `import pymupdf`，
但 Debian/Ubuntu 的 `python3-fitz`（1.23.x）只有 `import fitz`，
更老的一批包则相反。这里统一成 `pdfmod.load()`，两个名字都试。
"""
from __future__ import annotations

_cache: object | None = None


class PdfMissing(RuntimeError):
    """没装 PyMuPDF。"""


def load():
    """返回可用的 PyMuPDF 模块；都没有就抛 PdfMissing。"""
    global _cache
    if _cache is not None:
        return _cache
    try:
        import pymupdf  # type: ignore
        _cache = pymupdf
        return _cache
    except ImportError:
        pass
    try:
        import fitz  # type: ignore
        _cache = fitz
        return _cache
    except ImportError as exc:
        raise PdfMissing("未安装 PyMuPDF") from exc


def available() -> bool:
    try:
        load()
        return True
    except PdfMissing:
        return False
