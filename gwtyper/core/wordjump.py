# -*- coding: utf-8 -*-
"""在 Word / WPS 里定位到指定段落并选中。

体检结果双击时用：比在界面里跳行更实用 —— 直接看到原文那一段。
"""
from __future__ import annotations

import os
from pathlib import Path


class JumpError(RuntimeError):
    pass


def _norm(p: str) -> str:
    try:
        return os.path.normcase(os.path.abspath(p))
    except Exception:  # noqa: BLE001
        return p


def jump(src: str, para_index: int) -> None:
    """打开 src（已打开就复用），选中第 para_index+1 段。para_index 从 0 起。"""
    src = str(Path(src))
    if not Path(src).exists():
        raise JumpError(f"找不到文件：{src}")
    try:
        import pythoncom
        import win32com.client as win32
    except Exception as exc:  # noqa: BLE001
        raise JumpError("这台电脑上没有 Word/WPS 的自动化接口（pywin32 不可用）") from exc

    pythoncom.CoInitialize()
    app = None
    try:
        app = _get_app(win32)
        doc = _find_open(app, src) or app.Documents.Open(
            src, ReadOnly=True, AddToRecentFiles=False)
        try:
            app.Visible = True
        except Exception:  # noqa: BLE001
            pass
        try:
            app.Activate()
        except Exception:  # noqa: BLE001
            pass
        n = max(1, int(para_index) + 1)
        total = doc.Paragraphs.Count
        if n > total:
            n = total
        rng = doc.Paragraphs(n).Range
        rng.Select()
        try:
            app.ActiveWindow.ScrollIntoView(rng, True)
        except Exception:  # noqa: BLE001
            pass
    except JumpError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise JumpError(f"在 Word 里定位失败：{exc}") from exc
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass


def _get_app(win32):
    last = None
    for progid in ("Word.Application", "KWPS.Application"):
        try:
            return win32.GetActiveObject(progid)
        except Exception as exc:  # noqa: BLE001
            last = exc
        try:
            return win32.DispatchEx(progid)
        except Exception as exc:  # noqa: BLE001
            last = exc
    raise JumpError(f"没找到 Word/WPS：{last}")


def _find_open(app, src: str):
    target = _norm(src)
    try:
        for d in app.Documents:
            try:
                if _norm(d.FullName) == target:
                    return d
            except Exception:  # noqa: BLE001
                continue
    except Exception:  # noqa: BLE001
        pass
    return None


def available() -> bool:
    try:
        import win32com.client  # noqa: F401
        import pythoncom  # noqa: F401
    except Exception:  # noqa: BLE001
        return False
    try:
        from . import office

        return office.com_available()
    except Exception:  # noqa: BLE001
        return False
