# -*- coding: utf-8 -*-
"""把 .doc / .wps / .rtf 悄悄换成 .docx，让后面的流水线只认一种格式。

转换走 `office.py`（Windows 的 WPS/Office COM，或任意平台的 LibreOffice），
结果缓存在平台对应的 cache 目录，按「路径+大小+修改时间」做键，
源文件不变就一直用缓存，不会每次重转。
"""
from __future__ import annotations

import hashlib
from pathlib import Path

from . import office, sysenv

DOCX = ".docx"
LEGACY = {".doc", ".wps", ".rtf", ".dot", ".dotx", ".docm"}
SUPPORTED = LEGACY | {DOCX}


class DocSourceError(RuntimeError):
    """读不了这个文件（没装办公软件、文件坏了、加密等）。"""


def _cache_dir() -> Path:
    return sysenv.cache_dir("docx-cache")


def is_supported(path: str | Path) -> bool:
    return Path(path).suffix.lower() in SUPPORTED


def _cache_path(src: Path) -> Path:
    try:
        st = src.stat()
        stamp = f"{st.st_size}-{int(st.st_mtime)}"
    except OSError:
        stamp = "0-0"
    key = hashlib.sha1(f"{src.resolve()}|{stamp}".encode("utf-8")).hexdigest()[:12]
    return _cache_dir() / f"{src.stem}-{key}.docx"


def to_docx(src: str | Path, force: bool = False) -> Path:
    """返回可直接给 python-docx 的 .docx 路径；本来就是 docx 就原样返回。"""
    src = Path(src)
    if not src.exists():
        raise DocSourceError(f"找不到文件：{src}")
    if src.suffix.lower() == DOCX:
        return src

    dst = _cache_path(src)
    if dst.exists() and not force:
        return dst

    if not office.engine_name():
        raise DocSourceError(
            "打不开 .doc/.wps 这类老格式：本机没有可用的转换引擎。\n"
            "装一个 WPS 或 LibreOffice 就能自动转换。"
        )

    dst.parent.mkdir(parents=True, exist_ok=True)
    try:
        office.word_to_docx(src, dst)
    except Exception as exc:  # noqa: BLE001
        raise DocSourceError(
            f"转换失败（{exc}）。\n"
            "多半是本机没装 WPS、Office 或 LibreOffice，或者这个文件有密码保护。\n"
            "可以先用办公软件打开它，另存为 .docx 再试。"
        ) from exc

    if not dst.exists() or dst.stat().st_size == 0:
        raise DocSourceError("转换没有产出文件，可能文件损坏或受密码保护")

    # 转换器有时会「成功」产出一个空壳 docx（能打开但没有页面设置），得真验一次
    try:
        from docx import Document

        _probe = Document(str(dst))
        _sec = _probe.sections[0]
        if _sec.page_width is None or _sec.left_margin is None:
            raise ValueError("转换结果缺少页面设置")
    except Exception as exc:  # noqa: BLE001
        try:
            dst.unlink()
        except OSError:
            pass
        raise DocSourceError(
            f"这个文件转出来是空的，多半已经损坏：{Path(src).name}\n"
            "可以用办公软件打开看看，能打开就另存为 .docx 再试。"
        ) from exc
    return dst


def display_name(path: str | Path) -> str:
    return Path(path).name


def clear_cache() -> int:
    """清掉全部缓存，返回删除的文件数。"""
    n = 0
    base = _cache_dir()
    if not base.exists():
        return n
    for f in base.glob("*.docx"):
        try:
            f.unlink()
            n += 1
        except OSError:
            continue
    return n


def prune(max_age_days: int = 30, keep: int = 200) -> int:
    """删掉过期缓存，免得一直堆。返回删除数。"""
    base = _cache_dir()
    if not base.exists():
        return 0
    import time

    now = time.time()
    files = sorted(base.glob("*.docx"), key=lambda p: p.stat().st_mtime, reverse=True)
    n = 0
    for i, f in enumerate(files):
        try:
            too_old = (now - f.stat().st_mtime) > max_age_days * 86400
            if i >= keep or too_old:
                f.unlink()
                n += 1
        except OSError:
            continue
    return n
