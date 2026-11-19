# -*- coding: utf-8 -*-
"""文档转换：docx → PDF、老格式 → docx。三个引擎，按平台自动挑。

- **Windows COM**（WPS / Microsoft Office）：只在本机装了才走，排版还原度最高。
- **WPS for Linux**（`wps --convert-to`）：政府单位多用 WPS，字体与版式跟桌面里看到的一致。
- **LibreOffice**（Linux / macOS / Windows 装了都行）：`soffice --headless --convert-to`，
  跨平台一致，也是 AppImage 里最稳的兜底通道。

对外只暴露 `word_to_pdf` / `word_to_docx` / `engine_name`，调用方不判断平台。
"""
from __future__ import annotations

import logging
import os
import shutil
import subprocess
import sys
import tempfile
import threading
from pathlib import Path
from typing import Callable

from . import sysenv

log = logging.getLogger(__name__)


class OfficeError(RuntimeError):
    """没有可用的转换引擎，或者转换失败。"""


# ---------------------------------------------------------------- LibreOffice

_SOFFICE_CANDIDATES = [
    "soffice",
    "libreoffice",
    "/usr/bin/soffice",
    "/usr/lib/libreoffice/program/soffice",
    "/opt/libreoffice/program/soffice",
    "/snap/bin/libreoffice",
    "/var/lib/flatpak/exports/bin/org.libreoffice.LibreOffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "soffice.bin",
]

_soffice_cache: str | None = None
_soffice_probed = False


def soffice() -> str | None:
    """找到 soffice 可执行文件就返回路径，找不到返回 None。"""
    global _soffice_cache, _soffice_probed
    if _soffice_probed:
        return _soffice_cache
    _soffice_probed = True

    env = os.environ.get("GWTYPER_SOFFICE")
    if env and Path(env).exists():
        _soffice_cache = env
        return _soffice_cache

    for cand in _SOFFICE_CANDIDATES:
        got = shutil.which(cand) if not Path(cand).is_absolute() else (cand if Path(cand).exists() else None)
        if got:
            _soffice_cache = got
            return _soffice_cache
    return None


def _lo_convert(src: Path, dst_dir: Path, target: str, timeout: int = 120) -> Path:
    """跑一次 soffice 转换，返回产物路径。target 形如 pdf / docx。"""
    exe = soffice()
    if not exe:
        raise OfficeError("没找到 LibreOffice")
    dst_dir.mkdir(parents=True, exist_ok=True)

    # 单独一个 profile 目录，避免和用户正在开的 LibreOffice 抢锁
    profile = sysenv.cache_dir("lo-profile")
    cmd = [
        exe, "--headless", "--norestore", "--invisible", "--nodefault", "--nolockcheck",
        f"-env:UserInstallation={profile.as_uri()}",
        "--convert-to", target, "--outdir", str(dst_dir), str(src),
    ]
    try:
        proc = subprocess.run(cmd, capture_output=True, timeout=timeout)
    except subprocess.TimeoutExpired as exc:
        raise OfficeError(f"LibreOffice 转换超时（{timeout} 秒）") from exc
    except OSError as exc:
        raise OfficeError(f"LibreOffice 起不来：{exc}") from exc

    want = dst_dir / f"{src.stem}.{target.split(':')[0]}"
    if want.exists() and want.stat().st_size > 0:
        return want

    msg = (proc.stderr or proc.stdout or b"").decode("utf-8", "replace").strip()
    raise OfficeError(_explain_lo_failure(src, msg))


def _explain_lo_failure(src: Path, msg: str) -> str:
    """把 LibreOffice 那几句外语翻成人能照着做的话。

    `source file could not be loaded` 最常见的原因是 Linux 上只装了
    libreoffice-writer —— 它能开 doc/docx，但完全不认识 xlsx/pptx，
    于是表格和演示都报这一句，看着像文件坏了，其实是组件没装。
    """
    low = msg.lower()
    s = src.suffix.lower()
    if "source file could not be loaded" in low:
        if s in {".xls", ".xlsx", ".xlsm", ".csv", ".et"}:
            return ("LibreOffice 打不开表格文件——多半是没装表格组件。\n"
                    "Debian/Ubuntu：sudo apt install libreoffice-calc\n"
                    "Fedora：sudo dnf install libreoffice-calc")
        if s in {".ppt", ".pptx", ".pps", ".ppsx", ".dps"}:
            return ("LibreOffice 打不开演示文件——多半是没装演示组件。\n"
                    "Debian/Ubuntu：sudo apt install libreoffice-impress\n"
                    "Fedora：sudo dnf install libreoffice-impress")
        return ("LibreOffice 打不开这份文件：可能已损坏、加了密码，"
                "或者 Linux 上缺对应组件（表格要 libreoffice-calc，"
                f"演示要 libreoffice-impress）。\n原始信息：{msg[:200]}")
    return f"LibreOffice 没产出文件：{msg[:300] or '未知原因'}"


# ---------------------------------------------------------------- WPS for Linux

# WPS Office for Linux 的命令行入口。装了 WPS 的机器上这几个都在 PATH 或固定目录里。
_WPS_LINUX_CANDIDATES = [
    "wps",
    "/usr/bin/wps",
    "/opt/kingsoft/wps-office/office6/wps",
    "/opt/apps/cn.wps.wps-office/files/office6/wps",
    "/usr/lib/office6/wps",
    "/snap/bin/wps-office.wps",
]

_wps_cache: str | None = None
_wps_probed = False
# 有 wps 可执行文件但转不了。实测 WPS for Linux 12.x 根本不解析命令行参数
# （给什么都直接进 GUI 或立刻返回），所以第一次失败就不再白试。
# 但只在本次进程内生效，下次启动还会再试一次——万一用户换了版本或改了配置。
_wps_broken = False


def wps_can_convert() -> bool:
    """WPS 到底能不能命令行转换。不能用就直接说清楚，别让用户以为在用 WPS。"""
    return wps_linux() is not None and not _wps_broken


def wps_linux() -> str | None:
    """找到 Linux 版 WPS 的可执行文件就返回路径。"""
    global _wps_cache, _wps_probed
    if _wps_probed:
        return _wps_cache
    _wps_probed = True

    env = os.environ.get("GWTYPER_WPS")
    if env and Path(env).exists():
        _wps_cache = env
        return _wps_cache
    if sysenv.IS_WIN:
        return None

    for cand in _WPS_LINUX_CANDIDATES:
        got = shutil.which(cand) if not Path(cand).is_absolute() else (
            cand if Path(cand).exists() else None)
        if got:
            _wps_cache = got
            return _wps_cache
    return None


def _wps_convert(src: Path, dst_dir: Path, target: str, timeout: int = 120) -> Path:
    """用 WPS for Linux 转换。WPS 的命令行转换能力各版本不一，失败就交给上层换引擎。"""
    global _wps_broken
    exe = wps_linux()
    if not exe:
        raise OfficeError("没找到 WPS")
    dst_dir.mkdir(parents=True, exist_ok=True)
    cmd = [exe, "--headless", "--convert-to", target, "--outdir", str(dst_dir), str(src)]
    try:
        proc = subprocess.run(cmd, capture_output=True, timeout=timeout)
    except subprocess.TimeoutExpired as exc:
        raise OfficeError(f"WPS 转换超时（{timeout} 秒）") from exc
    except OSError as exc:
        _wps_broken = True
        raise OfficeError(f"WPS 起不来：{exc}") from exc

    want = dst_dir / f"{src.stem}.{target.split(':')[0]}"
    if want.exists() and want.stat().st_size > 0:
        return want
    msg = (proc.stderr or proc.stdout or b"").decode("utf-8", "replace").strip()
    # 这个版本的 WPS 没有命令行转换：记下来，别再每次都白试一遍
    _wps_broken = True
    raise OfficeError(f"WPS 没产出文件（这个版本的 WPS 可能不支持命令行转换）：{msg[:200] or '未知原因'}")


# ---------------------------------------------------------------- Windows COM

_WD_PDF = 17
_XL_PDF = 57
_PP_PDF = 32
_WD_DOCX = 16
_XL_XLSX = 51
_PP_PPTX = 24

# 三类 Office 各有一组 ProgID：WPS 的在前，Microsoft Office 的在后
_PROGIDS = {
    "word": ["Word.Application", "KWPS.Application"],
    "excel": ["Excel.Application", "KET.Application"],
    "ppt": ["PowerPoint.Application", "KWPP.Application"],
}

_com_broken = False


def kind_of(src: str | Path) -> str:
    """按扩展名判断这是 word / excel / ppt。认不出就按 word 处理。

    这一步很关键：Excel 用 `Word.Application.Documents.Open` 打开会直接报
    「文档打开失败」，WPS 给的错误码还特别难懂，所以必须先分对类。
    """
    s = Path(src).suffix.lower()
    if s in {".xls", ".xlsx", ".xlsm", ".csv", ".et"}:
        return "excel"
    if s in {".ppt", ".pptx", ".pps", ".ppsx", ".dps"}:
        return "ppt"
    return "word"


def _com_word_to(src: Path, dst: Path, pdf: bool) -> None:
    _com_to(src, dst, "word", pdf)


# ---------------------------------------------------------------- 实例复用
# 每份文件都 DispatchEx 一个新实例的话，批量转 50 份就是 50 个 wps.exe，
# 启动成本全摊在用户身上（WPS 冷启动 1~2 秒）。这里按线程缓存实例：
# COM 的 STA 规定对象只能在自己创建的线程里用，所以缓存必须挂在
# threading.local 上，不能用全局字典。
_tls = threading.local()
_ALL_APPS: list = []            # 所有线程创建过的实例，退出/清理时兜底关掉
_ALL_APPS_LOCK = threading.Lock()


def _cached_apps() -> dict:
    d = getattr(_tls, "apps", None)
    if d is None:
        d = {}
        _tls.apps = d
    return d


def _drop_app(kind: str) -> None:
    """把某个实例从缓存里摘掉并尝试退出。"""
    d = _cached_apps()
    app = d.pop(kind, None)
    if app is None:
        return
    try:
        app.Quit()
    except Exception:  # noqa: BLE001
        pass
    with _ALL_APPS_LOCK:
        _ALL_APPS[:] = [x for x in _ALL_APPS if x[1] is not app]


def _get_app(kind: str):
    """拿一个能用的实例：先试缓存（顺手探活），不行再新建。"""
    import win32com.client  # noqa: WPS433

    d = _cached_apps()
    app = d.get(kind)
    if app is not None:
        try:
            _ = app.Version          # 探活：进程被杀掉这里会抛
            if kind == "ppt":
                app.Visible = True
            return app
        except Exception:  # noqa: BLE001
            d.pop(kind, None)
    last = None
    for progid in _PROGIDS[kind]:
        try:
            app = win32com.client.DispatchEx(progid)
            break
        except Exception as exc:  # noqa: BLE001
            last = exc
    else:
        raise OfficeError(f"没有可用的 {kind} 办公软件：{last}")
    try:
        # WPS 演示在不可见时会拒绝打开文件，只能开窗口
        app.Visible = kind == "ppt"
    except Exception:  # noqa: BLE001
        pass
    try:
        app.DisplayAlerts = 0 if kind == "word" else False
    except Exception:  # noqa: BLE001
        pass
    d[kind] = app
    with _ALL_APPS_LOCK:
        _ALL_APPS.append((kind, app))
    return app


def release_apps() -> None:
    """退出本进程缓存的所有办公软件实例（转换跑完或退出程序时调）。"""
    for kind in list(_cached_apps().keys()):
        _drop_app(kind)


def kill_zombies() -> list[tuple[int, str]]:
    """清理没有窗口的 wps/et/wpp 残留进程。

    WPS 的 Quit() 经常不真退出，批量转换后会攒下一堆无窗口实例。判据很保守：
    只杀「没有主窗口标题」的，用户自己开着的文档（有标题）绝不碰。
    """
    if not sysenv.IS_WIN:
        return []
    import csv as _csv
    import io as _io
    import subprocess as _sp

    # 1) 收集所有有可见标题的窗口 PID
    titled: set[int] = set()
    try:
        import ctypes
        from ctypes import wintypes

        user32 = ctypes.windll.user32
        CB = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

        def _cb(hwnd, _lparam):
            if user32.IsWindowVisible(hwnd):
                n = user32.GetWindowTextLengthW(hwnd)
                if n > 0:
                    pid = wintypes.DWORD()
                    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                    titled.add(int(pid.value))
            return True

        user32.EnumWindows(CB(_cb), 0)
    except Exception as exc:  # noqa: BLE001
        log.warning("枚举窗口失败，跳过残留清理：%s", exc)
        return []

    # 2) 找目标进程
    names = {"wps.exe", "et.exe", "wpp.exe", "wpspdf.exe", "wpsoffice.exe"}
    try:
        out = _sp.run(["tasklist", "/fo", "csv", "/nh"], capture_output=True,
                      timeout=20, creationflags=getattr(_sp, "CREATE_NO_WINDOW", 0))
        text = out.stdout.decode("gbk", "replace")
    except Exception as exc:  # noqa: BLE001
        log.warning("tasklist 失败：%s", exc)
        return []
    killed: list[tuple[int, str]] = []
    for row in _csv.reader(_io.StringIO(text)):
        if len(row) < 2:
            continue
        name, pid_s = row[0].strip().lower(), row[1].strip()
        if name not in names or not pid_s.isdigit():
            continue
        pid = int(pid_s)
        if pid in titled or pid == os.getpid():
            continue
        try:
            _sp.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True,
                    timeout=10, creationflags=getattr(_sp, "CREATE_NO_WINDOW", 0))
            killed.append((pid, name))
        except Exception as exc:  # noqa: BLE001
            log.warning("结束 %s(%s) 失败：%s", name, pid, exc)
    if killed:
        log.info("清理无窗口办公进程：%s", killed)
    return killed


def _com_to(src: Path, dst: Path, kind: str, pdf: bool) -> None:
    """WPS 或 Office 的另存/导出。失败了抛异常，由上层换引擎。

    实例按线程复用（见 _get_app），成功后只关文档不退出程序——
    批量转换时这样才快；残留实例由 kill_zombies() 或 release_apps() 收尾。
    """
    import pythoncom  # noqa: WPS433

    pythoncom.CoInitialize()
    app = None
    doc = None
    ok = False
    try:
        app = _get_app(kind)
        if kind == "word":
            doc = app.Documents.Open(
                os.path.abspath(src), ReadOnly=True, AddToRecentFiles=False,
                ConfirmConversions=False,
            )
            doc.SaveAs(os.path.abspath(dst), FileFormat=_WD_PDF if pdf else _WD_DOCX)
        elif kind == "excel":
            doc = app.Workbooks.Open(os.path.abspath(src), ReadOnly=True, UpdateLinks=0)
            if pdf:
                # WPS 的表格不支持 SaveAs(FileFormat=57)，必须走导出接口
                doc.ExportAsFixedFormat(0, os.path.abspath(dst))
            else:
                doc.SaveAs(os.path.abspath(dst), FileFormat=_XL_XLSX)
        else:
            doc = app.Presentations.Open(os.path.abspath(src), ReadOnly=True,
                                         WithWindow=False)
            doc.SaveAs(os.path.abspath(dst), FileFormat=_PP_PDF if pdf else _PP_PPTX)
        ok = True
    finally:
        if doc is not None:
            try:
                doc.Close(0 if kind == "word" else False)
            except Exception:  # noqa: BLE001
                pass
        if not ok and app is not None:
            # 失败过的实例状态不明（可能弹了对话框），扔掉别再用
            _drop_app(kind)
        try:
            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass


_com_probed: bool | None = None      # 真正探测过一次的结果，缓存在模块里


def _com_progid_registered(progid: str) -> bool:
    """注册表里有没有这个 ProgID 的 CLSID —— 纯查询，不启动任何程序。"""
    try:
        import winreg  # noqa: WPS433

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + r"\CLSID"):
            return True
    except Exception:  # noqa: BLE001
        return False


def com_available() -> bool:
    """本机装没装能走 COM 的办公软件：只查注册表，绝不真的启动一次。

    这里以前是 DispatchEx 一次 Word.Application 做探测。代价看运气：
    机器上已经堆了十几个 WPS 窗口时，这一步能把界面卡到出不来
    （真实发生过：主窗口的「转换引擎」标签是在 UI 线程上同步调这个的），
    而且每次探测本身就会多留一个后台 wps.exe。
    注册表查询是微秒级、零副作用的；真到转换时打不开，_com_to 会抛得很
    清楚，上层自动回落 LibreOffice 并在界面上说明。
    """
    global _com_broken, _com_probed
    if not sysenv.IS_WIN or _com_broken:
        return False
    if _com_probed is not None:
        return _com_probed
    try:
        import win32com.client  # noqa: F401,WPS433
    except Exception:  # noqa: BLE001
        _com_broken = True
        _com_probed = False
        return False

    _com_probed = any(_com_progid_registered(p)
                      for group in _PROGIDS.values() for p in group)
    if not _com_probed:
        _com_broken = True
    return bool(_com_probed)


# ---------------------------------------------------------------- 对外接口

def engine_name() -> str:
    """当前首选哪个引擎：'com' / 'bundle' / 'libreoffice' / ''。

    优先级：
      1. Windows 的 WPS/Word COM —— 还原度最好，Windows 上 WPS 也没毛病
      2. **包内自带 LibreOffice** —— Linux 上首选。装在包里，目标机不用另装
      3. 系统 LibreOffice —— 兜底

    注意这里**不再考虑 Linux 版 WPS**：实测 12.1.2 根本不支持命令行转换
    （给参数直接返回、不产出文件），而且它会在后台常驻一个 wpscloudsvr
    进程，每次都弹。留着只会添乱，所以 Linux 侧直接跳过。
    （`GWTYPER_TRY_WPS=1` 可以强制重新考虑，调试用。）
    """
    if com_available():
        return "com"
    try:
        from . import lo_bundle

        if lo_bundle.status()["usable"]:
            return "bundle"
    except Exception:  # noqa: BLE001
        pass
    if os.environ.get("GWTYPER_TRY_WPS") == "1" and wps_can_convert():
        return "wps"
    if soffice():
        return "libreoffice"
    return ""


def engine_hint() -> str:
    """给界面/报错用的一句话说明。"""
    eng = engine_name()
    if eng == "com":
        return "WPS / Microsoft Office"
    if eng == "wps":
        return "WPS Office"
    if eng == "bundle":
        return "内置 LibreOffice"
    if eng == "libreoffice":
        # 装了 WPS 但它不支持命令行转换时说清楚，免得用户以为程序没认出 WPS
        if wps_linux():
            return "LibreOffice（已装 WPS，但这版不支持命令行转换）"
        return "LibreOffice"
    return "无（装 WPS 或 LibreOffice 即可）"


def _run(src: Path, dst: Path, pdf: bool,
         progress: Callable[[str], None] | None = None,
         kind: str | None = None) -> None:
    """按引擎优先级依次尝试，全失败才抛错。

    目标文件先写临时位置、成功才 replace 过去。以前是一上来就 unlink(dst)，
    结果所有引擎都失败时，用户原来那份文件已经被删掉且不会恢复
    （UI 里因为 out_name 总生成新名字所以没踩到，但这是导出给外部用的公共
    入口，直接传已有路径就会丢文件）。
    """
    src, dst = Path(src), Path(dst)
    if not src.exists():
        raise OfficeError(f"找不到文件：{src}")
    kind = kind or kind_of(src)
    dst.parent.mkdir(parents=True, exist_ok=True)

    # 临时目录必须和目标文件**同一个盘**：os.replace 不能跨驱动器
    # （实测 WinError 17 系统无法将文件移到不同的磁盘驱动器）。
    # mkdtemp 默认落在 C: 的系统临时目录，目标常在 D:，所以这里指定 dir。
    tmpdir = Path(tempfile.mkdtemp(prefix=".gwconv_", dir=str(dst.parent)))
    work = tmpdir / f"out{dst.suffix}"
    try:
        _run_to(src, work, pdf, progress, kind, dst)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


_COM_TIMEOUT = float(os.environ.get("GWTYPER_COM_TIMEOUT", "120"))


def _com_to_watched(src: Path, work: Path, kind: str, pdf: bool,
                    timeout: float | None = None) -> None:
    """在子线程里跑 COM 转换，超时就不等了。

    真实故障：机器上已经开了十几个 WPS 窗口时，`DispatchEx` / `Open`
    会无限期阻塞（实测挂过十几分钟，界面一直转圈、按钮全灰）。
    这里最多等 timeout 秒，超时抛 OfficeError 让上层回落到内置 LibreOffice，
    至少界面还是活的、用户能拿到文件。
    """
    import threading  # noqa: WPS433

    tmo = float(timeout if timeout is not None else _COM_TIMEOUT)
    box: dict = {}

    def worker() -> None:
        try:
            _com_to(src, work, kind, pdf)
            box["ok"] = True
        except BaseException as exc:  # noqa: BLE001
            box["err"] = exc

    t = threading.Thread(target=worker, daemon=True)
    t.start()
    t.join(tmo)
    if t.is_alive():
        raise OfficeError(
            f"WPS/Office 超过 {int(tmo)} 秒没响应——多半是开了太多 WPS 窗口，"
            "或者有个看不见的对话框在等确认。关掉多余的 WPS 再试一次。")
    if "err" in box:
        raise box["err"]


def _run_to(src: Path, work: Path, pdf: bool,
            progress: Callable[[str], None] | None,
            kind: str, final_dst: Path) -> None:
    errors: list[str] = []
    # 顺序和 engine_name() 一致。wps 只在显式开启时才试（GWTYPER_TRY_WPS=1）
    _try_wps = os.environ.get("GWTYPER_TRY_WPS") == "1"
    for eng in ("com", "bundle", "wps", "libreoffice"):
        if eng == "com" and not com_available():
            continue
        if eng == "wps" and (not _try_wps or not wps_linux() or _wps_broken):
            continue
        if eng == "libreoffice" and not soffice():
            continue
        if eng == "bundle":
            try:
                from . import lo_bundle

                if lo_bundle.bundled() is None:
                    continue
            except Exception:  # noqa: BLE001
                continue
        if progress:
            progress({"com": "WPS/Office 转换中…", "wps": "WPS 转换中…",
                      "bundle": "内置 LibreOffice 转换中…",
                      "libreoffice": "LibreOffice 转换中…"}[eng])
        try:
            if eng == "com":
                # COM 能分清 word/excel/ppt；命令行引擎自己看扩展名
                _com_to_watched(src, work, kind, pdf)
            elif eng == "wps":
                with tempfile.TemporaryDirectory() as td:
                    got = _wps_convert(src, Path(td), "pdf" if pdf else "docx")
                    shutil.move(str(got), str(work))
            elif eng == "bundle":
                from . import lo_bundle

                with tempfile.TemporaryDirectory() as td:
                    got = lo_bundle.convert(src, Path(td),
                                            "pdf" if pdf else "docx")
                    shutil.move(str(got), str(work))
            else:
                with tempfile.TemporaryDirectory() as td:
                    got = _lo_convert(src, Path(td), "pdf" if pdf else "docx")
                    shutil.move(str(got), str(work))
            if work.exists() and work.stat().st_size > 0:
                # 只有真出了文件才覆盖用户那份，失败时原文件还在
                os.replace(str(work), str(final_dst))
                return
            errors.append(f"{eng}：没有产出文件")
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{eng}：{exc}")

    detail = "；".join(errors) or "本机没装 LibreOffice，也没有 WPS 或 Office"
    raise OfficeError(detail)


def word_to_pdf(src: str | Path, dst: str | Path,
                progress: Callable[[str], None] | None = None) -> None:
    _run(Path(src), Path(dst), True, progress)


def word_to_docx(src: str | Path, dst: str | Path,
                 progress: Callable[[str], None] | None = None) -> None:
    _run(Path(src), Path(dst), False, progress)


def word_to_pdf_bookmarked(src, dst, progress=None) -> None:
    """导 PDF 并让 Word 按标题生成书签（一/二/三级）。

    WPS 对 CreateBookmarks 支持不稳定，失败就退回普通导出——
    宁可没有书签，也要有文件。
    """
    src, dst = Path(src), Path(dst)
    dst.parent.mkdir(parents=True, exist_ok=True)
    if com_available():
        try:
            import pythoncom  # noqa: WPS433

            pythoncom.CoInitialize()
            app = _get_app("word")
            doc = app.Documents.Open(str(src), ReadOnly=True)
            try:
                # 17 = wdExportFormatPDF，1 = wdExportCreateHeadingBookmarks
                doc.ExportAsFixedFormat(OutputFileName=str(dst), ExportFormat=17,
                                        OpenAfterExport=False, CreateBookmarks=1)
            finally:
                doc.Close(0)
            if dst.exists() and dst.stat().st_size > 0:
                return
        except Exception:  # noqa: BLE001
            log.debug("带书签导出失败，退回普通导出", exc_info=True)
    word_to_pdf(src, dst, progress)


# 名字带什么无所谓，实现都是同一套；下面两个是按语义取的别名，
# 免得在转换页里把 xlsx→PDF 写成 word_to_pdf 看着别扭。
def to_pdf(src: str | Path, dst: str | Path,
           progress: Callable[[str], None] | None = None) -> None:
    """任意 Office 格式（Word/Excel/PPT/老格式）转 PDF。"""
    _run(Path(src), Path(dst), True, progress)


def upgrade(src: str | Path, dst: str | Path,
            progress: Callable[[str], None] | None = None) -> None:
    """老格式升级：doc→docx / xls→xlsx / ppt→pptx，保留原排版。"""
    _run(Path(src), Path(dst), False, progress)
