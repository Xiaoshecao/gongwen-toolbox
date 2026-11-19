# -*- coding: utf-8 -*-
"""内嵌 LibreOffice 的接入层（Linux）。

包里自带一份 LibreOffice 时，麻烦在于它的可执行文件写死了
`/usr/lib/libreoffice`，放到别处起不来。两条路：

**A. unshare（首选，免 root）**
    `unshare -m --map-root-user` 建私有 mount namespace，把包内那份
    bind 到 /usr/lib/libreoffice，只对本进程有效、退出即消失、不动系统。

**B. root 提权（A 不可用时）**
    有些发行版禁用了非特权 user namespace。这时用 pkexec / sudo 请求一次
    root，做**全局** bind 挂载。做了就一直在，所以：
      - 挂之前先记标记文件，判断是否已经挂过（幂等，不重复挂）
      - 挂载点内容与包内不一致时（比如换了新版本）重新挂
      - 设置一个退出钩子，正常退出时尝试卸载（失败也无妨，下次会复用）

调用方只需要 `convert()`，内部自动选 A 或 B。
"""
from __future__ import annotations

import os
import shlex
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from . import sysenv

_BUNDLE_SUBDIR = "lib/libreoffice"
_MOUNT_POINT = Path("/usr/lib/libreoffice")
_MARK = "gwtyper-bundled-lo.marker"

_state: dict[str, object] = {
    "probed": False,
    "bundled": None,
    "unshare": None,
    "mounted_by_us": False,
}


def _q(s: str | Path) -> str:
    return shlex.quote(str(s))


# ---------------------------------------------------------------- 定位
def _bundle_dir() -> Path | None:
    """包内自带的 LibreOffice 目录（含 program/soffice）。"""
    roots: list[Path] = []
    # AppRun 会显式传进来，最可靠
    env = os.environ.get("GWTYPER_LO_BUNDLE")
    if env:
        roots.append(Path(env))
    appdir = os.environ.get("APPDIR")
    if appdir:
        roots.append(Path(appdir) / "usr" / _BUNDLE_SUBDIR)
    res = sysenv.resource_dir()
    roots += [
        res / "libreoffice",
        res.parent / "libreoffice",
        res / "usr" / _BUNDLE_SUBDIR,
        res.parent / _BUNDLE_SUBDIR,
    ]
    for r in roots:
        if (r / "program" / "soffice").exists() or (r / "program" / "soffice.bin").exists():
            return r.resolve()
    return None


def bundled() -> Path | None:
    if not _state["probed"]:
        _state["probed"] = True
        _state["bundled"] = _bundle_dir()
    return _state["bundled"]  # type: ignore[return-value]


# ---------------------------------------------------------------- 能力探测
def unshare_ok() -> bool:
    if _state["unshare"] is not None:
        return bool(_state["unshare"])
    ok = False
    if sys.platform.startswith("linux"):
        try:
            r = subprocess.run(["unshare", "-m", "--map-root-user", "true"],
                               capture_output=True, timeout=10)
            ok = r.returncode == 0
        except Exception:  # noqa: BLE001
            ok = False
    _state["unshare"] = ok
    return ok


def _root_helper() -> list[str] | None:
    """能用的提权命令。pkexec 有图形提示框，比 sudo 更适合桌面程序。"""
    for cand in ("pkexec", "sudo"):
        if shutil.which(cand):
            if cand == "sudo":
                # sudo 需要免密才行，否则会卡在终端里等密码
                try:
                    r = subprocess.run(["sudo", "-n", "true"],
                                       capture_output=True, timeout=8)
                    if r.returncode != 0:
                        continue
                except Exception:  # noqa: BLE001
                    continue
            return [cand]
    return None


def _is_root() -> bool:
    """当前是不是 root。Windows 上没有 geteuid，直接当非 root。"""
    try:
        return os.geteuid() == 0
    except AttributeError:
        return False


def can_elevate() -> bool:
    return _is_root() or _root_helper() is not None


# ---------------------------------------------------------------- 全局挂载
def _is_mounted() -> bool:
    """/usr/lib/libreoffice 现在是不是指向我们包内那份。"""
    b = bundled()
    if not b or not _MOUNT_POINT.exists():
        return False
    try:
        marker = _MOUNT_POINT / _MARK
        if marker.exists():
            return marker.read_text(encoding="utf-8").strip() == str(b)
    except OSError:
        pass
    return False


def _mount_globally() -> bool:
    """用 root 做一次全局 bind。已挂且指向同一份时直接返回 True。"""
    b = bundled()
    if b is None:
        return False
    if _is_mounted():
        return True

    helper = None if _is_root() else _root_helper()
    if not _is_root() and helper is None:
        return False

    script = (
        f"mkdir -p {_q(_MOUNT_POINT)} && "
        f"mountpoint -q {_q(_MOUNT_POINT)} && umount {_q(_MOUNT_POINT)} || true; "
        f"mount --bind {_q(b)} {_q(_MOUNT_POINT)} && "
        f"printf '%s' {_q(b)} > {_q(_MOUNT_POINT / _MARK)}"
    )
    cmd = (helper or []) + ["sh", "-c", script]
    try:
        r = subprocess.run(cmd, capture_output=True, timeout=120)
    except Exception:  # noqa: BLE001
        return False
    if r.returncode != 0:
        _state["mount_error"] = (r.stderr or r.stdout or b"").decode(
            "utf-8", "replace")[:200]
        return False
    _state["mounted_by_us"] = True
    return True


# ---------------------------------------------------------------- 对外
def status() -> dict:
    """给界面/诊断用。"""
    b = bundled()
    return {
        "bundled": str(b) if b else "",
        "unshare": unshare_ok(),
        "can_elevate": can_elevate(),
        "mounted": _is_mounted(),
        "usable": bool(b) and (unshare_ok() or can_elevate() or _is_mounted()),
    }


def prepare() -> str:
    """确保包内 LibreOffice 可用，返回用到的方式：'unshare' / 'root' / ''。"""
    if bundled() is None:
        return ""
    if unshare_ok():
        return "unshare"
    if _mount_globally():
        return "root"
    return ""


def convert(src: Path, dst_dir: Path, target: str = "pdf",
            timeout: int = 240) -> Path:
    """用包内 LibreOffice 转一个文件。失败抛 RuntimeError。"""
    b = bundled()
    if b is None:
        raise RuntimeError("这个包里没有自带 LibreOffice")

    dst_dir.mkdir(parents=True, exist_ok=True)
    profile = sysenv.cache_dir("lo-bundled-profile")
    profile.mkdir(parents=True, exist_ok=True)
    want = dst_dir / f"{src.stem}.{target}"

    conv = (
        f" --headless --norestore --invisible --nodefault --nolockcheck"
        f" -env:UserInstallation={profile.as_uri()}"
        f" --convert-to {target} --outdir {_q(dst_dir)} {_q(src)}"
    )

    # A. unshare —— 免 root，最干净
    if unshare_ok():
        # 挂载点可能不存在（有的发行版 /usr/lib 下没有 libreoffice 目录），
        # 先 mkdir -p；挂载失败时把错误留住，别用 2>/dev/null 吞掉。
        inner = (
            f"mkdir -p {_q(_MOUNT_POINT)}; "
            f"mount --bind {_q(b)} {_q(_MOUNT_POINT)} || "
            f"{{ echo 'MOUNT_FAILED' >&2; exit 9; }}; "
            f"exec /usr/lib/libreoffice/program/soffice{conv}"
        )
        cmd = ["unshare", "-m", "--map-root-user", "bash", "-c", inner]
        if _try(cmd, want, timeout, label="unshare"):
            return want

    # B. root 全局挂载
    if _mount_globally():
        cmd = ["sh", "-c", f"{_q(_MOUNT_POINT / 'program' / 'soffice')}{conv}"]
        if _try(cmd, want, timeout, label="root-mount"):
            return want

    err = last_error()
    raise RuntimeError(
        "包内 LibreOffice 起不来。\n"
        + (f"详细信息：{err}\n" if err else "")
        + "可以改用「纯 Python」引擎，或在本机装一个 LibreOffice。")


def _try(cmd: list[str], want: Path, timeout: int, label: str) -> bool:
    try:
        r = subprocess.run(cmd, capture_output=True, timeout=timeout)
    except Exception as exc:  # noqa: BLE001
        _state[f"err_{label}"] = f"{type(exc).__name__}: {exc}"[:300]
        _log(f"{label}: 启动失败 {_state[f'err_{label}']}")
        return False
    if want.exists() and want.stat().st_size > 0:
        return True
    out = (r.stderr or r.stdout or b"").decode("utf-8", "replace").strip()
    _state[f"err_{label}"] = f"rc={r.returncode} {out[:400]}"
    _log(f"{label}: rc={r.returncode} 没产物。输出：{out[:400]}")
    _log(f"{label}: 期望产物 {want}  存在={want.exists()}")
    return False


def _log(msg: str) -> None:
    """诊断日志写到缓存目录，界面看不到时也能事后查。"""
    try:
        d = sysenv.cache_dir("logs")
        d.mkdir(parents=True, exist_ok=True)
        with open(d / "lo-bundle.log", "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:  # noqa: BLE001
        pass
    if os.environ.get("GWTYPER_DEBUG"):
        print(f"[lo_bundle] {msg}", file=sys.stderr)


def last_error() -> str:
    for k, v in _state.items():
        if k.startswith("err_"):
            return f"{k}: {v}"
    return ""
