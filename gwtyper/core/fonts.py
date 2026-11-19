# -*- coding: utf-8 -*-
"""本机字体检测：规则里写了但系统没装的字体要提前告警，别让它静默替换。

Windows 用 GDI 枚举；Linux / macOS 用 fontconfig 的 fc-list。
取不到就返回空集（视为「未知」），宁可不提示也不误报。
"""
from __future__ import annotations

import subprocess
import sys
from functools import lru_cache
from pathlib import Path

LF_FACESIZE = 32


def _win_families() -> set[str]:
    import ctypes
    from ctypes import wintypes

    class LOGFONTW(ctypes.Structure):
        _fields_ = [
            ("lfHeight", wintypes.LONG), ("lfWidth", wintypes.LONG),
            ("lfEscapement", wintypes.LONG), ("lfOrientation", wintypes.LONG),
            ("lfWeight", wintypes.LONG), ("lfItalic", wintypes.BYTE),
            ("lfUnderline", wintypes.BYTE), ("lfStrikeOut", wintypes.BYTE),
            ("lfCharSet", wintypes.BYTE), ("lfOutPrecision", wintypes.BYTE),
            ("lfClipPrecision", wintypes.BYTE), ("lfQuality", wintypes.BYTE),
            ("lfPitchAndFamily", wintypes.BYTE), ("lfFaceName", wintypes.WCHAR * LF_FACESIZE),
        ]

    class ENUMLOGFONTEXW(ctypes.Structure):
        _fields_ = [("elfLogFont", LOGFONTW), ("elfFullName", wintypes.WCHAR * 64),
                    ("elfStyle", wintypes.WCHAR * 32), ("elfScript", wintypes.WCHAR * 32)]

    families: set[str] = set()
    gdi32, user32 = ctypes.WinDLL("gdi32"), ctypes.WinDLL("user32")
    hdc = user32.GetDC(None)
    if not hdc:
        return families
    try:
        proc_type = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.POINTER(ENUMLOGFONTEXW),
                                       ctypes.POINTER(ctypes.c_void_p), wintypes.DWORD,
                                       ctypes.c_ssize_t)

        def cb(lpelfe, _pntm, _ftype, _lparam):
            families.add(lpelfe.contents.elfLogFont.lfFaceName)
            return 1

        lf = LOGFONTW()
        for cs in (134, 1):        # 先 GB2312 字符集，再补全
            lf.lfCharSet = cs
            gdi32.EnumFontFamiliesExW(hdc, ctypes.byref(lf), proc_type(cb), 0, 0)
    finally:
        user32.ReleaseDC(None, hdc)
    return families


def _fc_families() -> set[str]:
    """Linux / macOS：fc-list 的 family 列，一行可能多个别名，逗号分开。"""
    families: set[str] = set()
    try:
        out = subprocess.run(["fc-list", ":", "family"], capture_output=True, timeout=20)
    except Exception:  # noqa: BLE001
        return families
    for line in out.stdout.decode("utf-8", "replace").splitlines():
        for name in line.split(","):
            name = name.strip()
            if name:
                families.add(name)
    return families


def _mac_families() -> set[str]:
    """没有 fc-list 时退到 system_profiler。"""
    families = _fc_families()
    if families:
        return families
    try:
        out = subprocess.run(["system_profiler", "SPFontsDataType", "-json"],
                             capture_output=True, timeout=30)
    except Exception:  # noqa: BLE001
        return families
    import json

    try:
        data = json.loads(out.stdout.decode("utf-8", "replace") or "{}")
    except Exception:  # noqa: BLE001
        return families
    for item in data.get("SPFontsDataType", []):
        for t in item.get("_items", []) or item.get("typefaces", []):
            name = (t.get("_name") or t.get("family") or "").strip()
            if name:
                families.add(name)
    return families


def _fc_resolves(font: str) -> bool:
    """用 fc-match 问一句：这个名字最终会落到哪个字体文件上。

    光靠 fc-list 的 family 列表不够——用户可能用 fontconfig 别名把
    「仿宋_GB2312」映射到了 FangSong，这时 fc-list 里没有中文名，
    但排版（LibreOffice/WPS 都走 fontconfig）实际是能拿到正确字体的。
    """
    import subprocess

    try:
        out = subprocess.run(["fc-match", "-f", "%{family}|%{file}", font],
                             capture_output=True, timeout=10)
    except Exception:  # noqa: BLE001
        return False
    text = out.stdout.decode("utf-8", "replace").strip()
    if not text or "|" not in text:
        return False
    fam, _, path = text.partition("|")
    path = path.strip()
    if not path or not Path(path).exists():
        return False
    # fc-match 对完全没装的名字会退回 DejaVu/Liberation 之类，那种不算命中
    fams = {x.strip() for x in fam.split(",") if x.strip()}
    if font in fams:
        return True
    low = " ".join(fams).lower()
    if any(x in low for x in ("dejavu", "liberation", "noto sans mono", "freeserif")):
        return False
    return True


@lru_cache(maxsize=1)
def installed() -> set[str]:
    """枚举本机所有字体家族名。失败时返回空集（视为未知，不做告警）。"""
    try:
        if sys.platform.startswith("win"):
            return _win_families()
        if sys.platform == "darwin":
            return _mac_families()
        return _fc_families()
    except Exception:  # noqa: BLE001
        return set()


def missing(fonts) -> list[str]:
    inst = installed()
    if not inst:
        return []
    miss = []
    for f in sorted({x for x in fonts if x}):
        if f in inst:
            continue
        # 不在 family 列表里，再问一次 fontconfig 的最终解析结果
        if not sys.platform.startswith("win") and _fc_resolves(f):
            continue
        miss.append(f)
    return miss


def check_fonts(spec_fonts) -> dict:
    inst = installed()
    miss = missing(spec_fonts)
    return {"checked": bool(inst), "total": len(inst), "missing": miss}


def font_dir_candidates() -> list[Path]:
    """想装字体时该放哪儿——README 里给人看用。"""
    if sys.platform.startswith("win"):
        return [Path(r"C:\Windows\Fonts")]
    if sys.platform == "darwin":
        return [Path.home() / "Library" / "Fonts", Path("/Library/Fonts")]
    return [Path.home() / ".local" / "share" / "fonts", Path("/usr/share/fonts")]
