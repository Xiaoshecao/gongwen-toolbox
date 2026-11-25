# -*- coding: utf-8 -*-
"""平台差异都收在这里：目录、打开文件、常量。其余模块不许再判断 sys.platform。

Windows：数据与缓存在 %LOCALAPPDATA%\\公文排版工具（沿用旧路径，老缓存继续有效）。
Linux/macOS：遵循 XDG，数据在 ~/.local/share/公文排版工具，缓存在 ~/.cache/公文排版工具。
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

APP_TITLE = "公文工具箱"
# 目录名故意还用旧名：用户已经在 %LOCALAPPDATA%\公文排版工具\rules 里存了自己的
# 格式标准（如 发城.yaml），改成新名会让这些设置"消失"。改名只改窗口标题。
APP_NAME = "公文排版工具"
IS_WIN = sys.platform.startswith("win")
IS_MAC = sys.platform == "darwin"
IS_LINUX = sys.platform.startswith("linux")


def frozen() -> bool:
    """是不是 PyInstaller 打出来的单文件包。"""
    return getattr(sys, "frozen", False)


def resource_dir() -> Path:
    """只读资源（rules/、assets/）所在目录。打包后在 _MEIPASS 里。"""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return Path(base)
    return Path(__file__).resolve().parent.parent


def data_dir() -> Path:
    """可写的数据目录：用户自建规则集、日志。"""
    if IS_WIN:
        base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA") or str(Path.home())
    elif IS_MAC:
        base = str(Path.home() / "Library" / "Application Support")
    else:
        base = os.environ.get("XDG_DATA_HOME") or str(Path.home() / ".local" / "share")
    p = Path(base) / APP_NAME
    p.mkdir(parents=True, exist_ok=True)
    return p


def cache_dir(name: str) -> Path:
    """缓存目录（可随时删）。Windows 下与 data_dir 同源，保持旧缓存可用。"""
    if IS_WIN:
        p = data_dir() / name
    elif IS_MAC:
        p = Path.home() / "Library" / "Caches" / APP_NAME / name
    else:
        base = os.environ.get("XDG_CACHE_HOME") or str(Path.home() / ".cache")
        p = Path(base) / APP_NAME / name
    p.mkdir(parents=True, exist_ok=True)
    return p


def user_rules_dir() -> Path:
    p = data_dir() / "rules"
    p.mkdir(parents=True, exist_ok=True)
    return p


def open_with_default(path: str | Path) -> None:
    """用系统默认程序打开文件。"""
    path = str(path)
    try:
        if IS_WIN:
            os.startfile(path)  # type: ignore[attr-defined]
        elif IS_MAC:
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:  # noqa: BLE001
        pass


def reveal(path: str | Path) -> None:
    """在文件管理器里定位到该文件（Windows 会选中它）。"""
    path = Path(path)
    try:
        if IS_WIN:
            if path.exists() and path.is_file():
                subprocess.Popen(["explorer", "/select,", str(path)])
            else:
                os.startfile(str(path.parent))  # type: ignore[attr-defined]
        elif IS_MAC:
            subprocess.Popen(["open", "-R", str(path)] if path.is_file() else ["open", str(path.parent)])
        else:
            # 没有统一的「在文件夹中选中」，退化为打开所在目录
            subprocess.Popen(["xdg-open", str(path.parent if path.is_file() else path)])
    except Exception:  # noqa: BLE001
        pass


def prefs_path() -> Path:
    return data_dir() / "prefs.json"


def load_prefs() -> dict:
    """读用户偏好（窗口大小、布局密度之类）。

    坏了就当空的，绝不让设置文件把程序卡在启动阶段。
    """
    import json

    try:
        f = prefs_path()
        if f.is_file():
            data = json.loads(f.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
    except Exception:  # noqa: BLE001
        pass
    return {}


def save_prefs(prefs: dict) -> None:
    import json

    try:
        prefs_path().write_text(
            json.dumps(prefs, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass


def desktop_dir() -> Path:
    """真实的「桌面」目录。

    不能只写 Path.home()/"Desktop"：政府/单位机器常把桌面重定向到别的盘
    （本机实测是 D:\\SystemFiles\\Desktop，而 home 下还有一份空的），
    按 home 猜会让默认输出目录指向一个用户根本没在用的地方。
    Windows 上问 Shell；其他平台按惯例回退。
    """
    if IS_WIN:
        try:
            import ctypes
            from ctypes import wintypes

            CSIDL_DESKTOPDIRECTORY = 0x0010
            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
            r = ctypes.windll.shell32.SHGetFolderPathW(
                None, CSIDL_DESKTOPDIRECTORY, None, 0, buf)
            if r == 0 and buf.value:
                p = Path(buf.value)
                if p.is_dir():
                    return p
        except Exception:  # noqa: BLE001
            pass
    else:
        try:
            import subprocess

            r = subprocess.run(["xdg-user-dir", "DESKTOP"],
                               capture_output=True, timeout=5)
            got = r.stdout.decode("utf-8", "replace").strip()
            if got and Path(got).is_dir():
                return Path(got)
        except Exception:  # noqa: BLE001
            pass
    for cand in (Path.home() / "Desktop", Path.home() / "桌面"):
        if cand.is_dir():
            return cand
    return Path.home()


def open_folder(folder: str | Path) -> None:
    folder = Path(folder)
    try:
        if IS_WIN:
            os.startfile(str(folder))  # type: ignore[attr-defined]
        elif IS_MAC:
            subprocess.Popen(["open", str(folder)])
        else:
            subprocess.Popen(["xdg-open", str(folder)])
    except Exception:  # noqa: BLE001
        pass


def log_path(name: str = "error.log") -> Path:
    return data_dir() / name


def platform_note() -> str:
    """状态栏用的短说明。"""
    if IS_WIN:
        return "Windows"
    if IS_MAC:
        return "macOS"
    return "Linux"


def documents_dir() -> Path:
    """「文档」目录（Windows 问 Shell，其他平台按惯例）。"""
    if IS_WIN:
        try:
            import ctypes
            from ctypes import wintypes

            CSIDL_PERSONAL = 0x0005
            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
            if ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, 0, buf) == 0 and buf.value:
                p = Path(buf.value)
                if p.is_dir():
                    return p
        except Exception:  # noqa: BLE001
            pass
    return Path.home()


def writable(p) -> bool:
    """真写一下试试——Windows 上「有权限」和「能写」经常不是一回事。"""
    p = Path(p)
    try:
        p.mkdir(parents=True, exist_ok=True)
        t = p / ".gw_write_test"
        t.write_text("x", encoding="utf-8")
        t.unlink()
        return True
    except OSError:
        return False


def export_dir() -> Path:
    """默认输出目录：优先桌面，桌面不能写就退「文档」→ 程序数据目录下的「导出」。

    Windows 上如果程序被装进 Program Files、或者桌面被安全软件锁住，
    直接往桌面写会弹「需要管理员权限」；这里提前挑一个确定能写的目录。
    """
    for cand in (desktop_dir(), documents_dir(), data_dir() / "导出"):
        try:
            if writable(cand):
                return Path(cand)
        except OSError:
            continue
    return data_dir()
