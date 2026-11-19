# -*- coding: utf-8 -*-
"""应用图标：窗口标题栏、任务栏、打包产物统一用这一份。

图标放在 `gwtyper/assets/`（icon.ico 给 Windows，icon.png 给其他平台）。
以前是从旧项目 `converter/assets/` 借的——那个目录一旦被删，
打包就静默变成无图标 exe，标题栏和任务栏也是空白。
"""
from __future__ import annotations

import sys
import tkinter as tk
from pathlib import Path

from . import sysenv


def icon_paths() -> list[Path]:
    """按优先级列出可用的图标文件。"""
    base = sysenv.resource_dir() / "assets"
    out: list[Path] = []
    # Windows 任务栏/标题栏优先 .ico（多尺寸，缩放清晰）
    if sys.platform.startswith("win"):
        out.append(base / "icon.ico")
    out.append(base / "icon.png")
    out.append(base / "icon.ico")
    # 兼容旧布局（源码运行时 resources 在上一级）
    out.append(base.parent / "converter" / "assets" / "icon.ico")
    out.append(base.parent / "converter" / "assets" / "icon.png")
    seen: set[Path] = set()
    uniq: list[Path] = []
    for p in out:
        if p not in seen and p.exists():
            seen.add(p)
            uniq.append(p)
    return uniq


def apply_to(root: tk.Misc, title: str = "") -> bool:
    """把图标挂到窗口上。成功返回 True。

    Windows：`iconbitmap(.ico)` 同时影响标题栏和任务栏；
    其它平台：`iconphoto` 用 PNG。两者都试一遍，哪个成功都算成功。
    """
    paths = icon_paths()
    if not paths:
        return False
    ok = False
    ico = next((p for p in paths if p.suffix.lower() == ".ico"), None)
    png = next((p for p in paths if p.suffix.lower() == ".png"), None)

    # Windows 上只用 iconbitmap(.ico)：iconphoto 在这套 Tk 上对标题栏不生效，
    # 而且它后调用会把 iconbitmap 的结果顶掉（实测 WM_GETICON 取回来是 0）。
    if ico and sys.platform.startswith("win"):
        try:
            root.iconbitmap(str(ico))
            ok = True
        except Exception:  # noqa: BLE001
            try:
                root.iconbitmap(default=str(ico))
                ok = True
            except Exception:  # noqa: BLE001
                pass
        if ok:
            return True

    # 其它平台（以及 Windows 上 ico 失败时）用 iconphoto
    if png:
        try:
            img = tk.PhotoImage(file=str(png))
            root.iconphoto(True, img)
            # 必须留引用，否则 PhotoImage 被回收、图标变空白
            setattr(root, "_icon_photo", img)
            ok = True
        except Exception:  # noqa: BLE001
            pass
    if ico and not sys.platform.startswith("win"):
        try:
            root.iconbitmap(default=str(ico))
            ok = True
        except Exception:  # noqa: BLE001
            pass
    return ok
