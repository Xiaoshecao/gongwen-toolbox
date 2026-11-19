# -*- coding: utf-8 -*-
"""Office 转换引擎：优先调用本机 WPS / Microsoft Office 的 COM 接口。

COM 必须在自带 CoInitialize 的线程里使用，所以这里不直接在主线程调用，
而是由任务线程统一处理（见 `run_com`）。
"""
from __future__ import annotations

import os
import time
from pathlib import Path

# WPS 与 Office 的 SaveAs 常量一致
WD_PDF = 17
XL_PDF = 57
PP_PDF = 32
WD_DOCX = 16
XL_XLSX = 51
PP_PPTX = 24

_APP_PROGIDS = {
    "word": ["Word.Application", "KWPS.Application"],
    "excel": ["Excel.Application", "KET.Application"],
    "ppt": ["PowerPoint.Application", "KWPP.Application"],
}


class ComError(RuntimeError):
    pass


def ensure_com() -> None:
    """在非主线程里用 COM 必须先初始化，否则会报「CoInitialize 未调用」。"""
    try:
        import pythoncom

        pythoncom.CoInitialize()
    except Exception:
        pass


def _client():
    import win32com.client  # noqa: WPS433  延迟导入，避免无 pywin32 时启动失败

    return win32com.client


def _dispatch(kind: str):
    client = _client()
    last = None
    for progid in _APP_PROGIDS[kind]:
        try:
            app = client.DispatchEx(progid)
            return app, progid
        except Exception as exc:  # 没装这个办公软件就换下一个
            last = exc
    raise ComError(f"没有找到可用的 {kind} 办公软件：{last}")


def _quit(app) -> None:
    """WPS 的 Quit 经常抛 RPC 已断开，属于正常现象，吞掉即可。"""
    for attr in ("Quit", "quit"):
        try:
            getattr(app, attr)()
            return
        except Exception:
            continue


def word_to(path: str, out: str, pdf: bool = True) -> None:
    app, _ = _dispatch("word")
    doc = None
    try:
        try:
            app.Visible = False
        except Exception:
            pass
        try:
            app.DisplayAlerts = 0
        except Exception:
            pass
        doc = app.Documents.Open(
            os.path.abspath(path), ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False
        )
        doc.SaveAs(os.path.abspath(out), FileFormat=WD_PDF if pdf else WD_DOCX)
    finally:
        if doc is not None:
            try:
                doc.Close(0)
            except Exception:
                pass
        _quit(app)


def excel_to(path: str, out: str, pdf: bool = True) -> None:
    app, _ = _dispatch("excel")
    book = None
    try:
        try:
            app.Visible = False
        except Exception:
            pass
        try:
            app.DisplayAlerts = False
        except Exception:
            pass
        book = app.Workbooks.Open(os.path.abspath(path), ReadOnly=True, UpdateLinks=0)
        if pdf:
            # WPS 的 Excel 不支持 SaveAs(FileFormat=57)，必须走导出接口
            book.ExportAsFixedFormat(0, os.path.abspath(out))
        else:
            book.SaveAs(os.path.abspath(out), FileFormat=XL_XLSX)
    finally:
        if book is not None:
            try:
                book.Close(False)
            except Exception:
                pass
        _quit(app)


def ppt_to(path: str, out: str, pdf: bool = True) -> None:
    app, _ = _dispatch("ppt")
    pres = None
    try:
        try:
            app.Visible = True  # WPS 演示在不可见时会拒绝打开文件
        except Exception:
            pass
        pres = app.Presentations.Open(os.path.abspath(path), ReadOnly=True, WithWindow=False)
        pres.SaveAs(os.path.abspath(out), PP_PDF if pdf else PP_PPTX)
    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        _quit(app)


def to_pdf(path: str, out: str) -> None:
    from .formats import suffix

    s = suffix(path)
    if s in {".doc", ".docx", ".docm", ".wps", ".rtf", ".dot", ".dotx"}:
        word_to(path, out, pdf=True)
    elif s in {".xls", ".xlsx", ".xlsm", ".et"}:
        excel_to(path, out, pdf=True)
    elif s in {".ppt", ".pptx", ".pps", ".ppsx", ".dps"}:
        ppt_to(path, out, pdf=True)
    else:
        raise ComError(f"不支持的 Office 格式：{s}")


def upgrade(path: str, out: str) -> None:
    """doc->docx / xls->xlsx / ppt->pptx，保留原有排版。"""
    from .formats import suffix

    s = suffix(path)
    if s in {".doc", ".docm", ".wps", ".rtf", ".dot", ".dotx"}:
        word_to(path, out, pdf=False)
    elif s in {".xls", ".et"}:
        excel_to(path, out, pdf=False)
    elif s in {".ppt", ".pps", ".dps"}:
        ppt_to(path, out, pdf=False)
    else:
        raise ComError(f"该格式无需升级：{s}")


def available() -> dict[str, str]:
    """探测哪些办公软件可用，返回 {类别: 程序标识}。"""
    found: dict[str, str] = {}
    for kind in _APP_PROGIDS:
        try:
            app, progid = _dispatch(kind)
            found[kind] = progid
            _quit(app)
        except Exception:
            continue
    return found


def _server_path(progid: str) -> str:
    """从注册表读出 COM 服务器实际指向的 exe，用来区分 WPS 和 Office。"""
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf"{progid}\CLSID") as k:
            clsid = winreg.QueryValueEx(k, "")[0]
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf"CLSID\{clsid}\LocalServer32") as k:
            return winreg.QueryValueEx(k, "")[0]
    except Exception:
        return ""


def vendor() -> str:
    """返回 'WPS' / 'Microsoft Office' / ''（都没有）。"""
    for kind in _APP_PROGIDS:
        for progid in _APP_PROGIDS[kind]:
            path = _server_path(progid).lower()
            if not path:
                continue
            if "kingsoft" in path or "wps.exe" in path or "\\et.exe" in path or "\\wpp.exe" in path:
                return "WPS"
            if "microsoft" in path or "winword" in path or "excel.exe" in path or "powerpnt" in path:
                return "Microsoft Office"
    return ""


def warm_up() -> None:
    """预热：首次启动 COM 比较慢，提前把三个组件叫醒。"""
    for kind in _APP_PROGIDS:
        try:
            app, _ = _dispatch(kind)
            _quit(app)
            time.sleep(0.05)
        except Exception:
            continue
