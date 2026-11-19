# -*- coding: utf-8 -*-
"""字体查找与子集化：纯 Python 出一条 PDF 时，字体是唯一的硬依赖。

两个问题要一起解决：
1. **找得到** —— 公文里写的是「仿宋_GB2312」「方正小标宋简体」，这些名字
   在不同机器上对应的文件各不相同。要按名字一层层退到实际文件。
2. **打得小** —— 中文字体一整份 10 MB，全量嵌进 PDF 就是 48 MB（实测）。
   只嵌入用到的字，同样的文档能压到 141 KB。

子集化用 fontTools；找不到 fontTools 就退回全量嵌入，功能不受影响、只是文件大。
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

# 公文字体名 → 常见的实际文件名（按优先级）。
# 前几个是 Windows 自带，后面是 Linux 上装 Noto/文泉驿时的替代。
FONT_FILES: dict[str, list[str]] = {
    "方正小标宋简体": ["FZXBSJW.TTF", "方正小标宋简体.ttf", "simsun.ttc"],
    "方正小标宋_GBK": ["FZXBSJW_GBK.TTF", "方正小标宋_GBK.ttf", "simsun.ttc"],
    "方正小标宋": ["FZXBSJW.TTF", "simsun.ttc"],
    "仿宋_GB2312": ["simfang.ttf", "FangSong_GB2312.ttf", "fangsong.ttf"],
    "仿宋": ["simfang.ttf", "fangsong.ttf"],
    "楷体_GB2312": ["simkai.ttf", "KaiTi_GB2312.ttf", "kaiti.ttf"],
    "楷体": ["simkai.ttf", "kaiti.ttf"],
    "黑体": ["simhei.ttf", "SimHei.ttf"],
    "宋体": ["simsun.ttc", "SimSun.ttf"],
    "微软雅黑": ["msyh.ttc", "msyh.ttf"],
    "Times New Roman": ["times.ttf"],
    "Arial": ["arial.ttf"],
}

# 找不到原始字体时，用这些兜底（顺序即优先级）
FALLBACK = ["simfang.ttf", "simsun.ttc", "simhei.ttf", "msyh.ttc"]


def _font_dirs() -> list[Path]:
    """这台机器上可能放字体的目录。"""
    out: list[Path] = []
    if sys.platform.startswith("win"):
        out.append(Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts")
        local = os.environ.get("LOCALAPPDATA")
        if local:
            out.append(Path(local) / "Microsoft" / "Windows" / "Fonts")
    elif sys.platform == "darwin":
        out += [Path.home() / "Library" / "Fonts", Path("/Library/Fonts"),
                Path("/System/Library/Fonts")]
    else:
        out += [Path.home() / ".local" / "share" / "fonts",
                Path.home() / ".fonts",
                Path("/usr/share/fonts"), Path("/usr/local/share/fonts")]

    # 程序自带的那份（AppImage / exe 里 /usr/share/fonts/truetype/gwtyper）
    try:
        from . import sysenv

        res = sysenv.resource_dir()
        for rel in ("assets/fonts", "../usr/share/fonts/truetype/gwtyper",
                    "fonts"):
            p = (res / rel).resolve()
            if p.is_dir():
                out.append(p)
    except Exception:  # noqa: BLE001
        pass
    return [d for d in out if d.is_dir()]


def _fc_match(family: str) -> Path | None:
    """问 fontconfig 要文件路径——Linux 上这是最靠谱的一条。"""
    if sys.platform.startswith("win"):
        return None
    try:
        import subprocess

        r = subprocess.run(["fc-match", "-f", "%{file}", family],
                           capture_output=True, timeout=8)
    except Exception:  # noqa: BLE001
        return None
    p = r.stdout.decode("utf-8", "replace").strip()
    if p and Path(p).exists():
        return Path(p)
    return None


def resolve(name: str) -> Path | None:
    """把一个字体名解析成实际文件。找不到返回 None。"""
    if not name:
        return None
    name = name.strip()

    # 1) 本身就是路径
    p = Path(name)
    if p.suffix.lower() in {".ttf", ".ttc", ".otf"} and p.exists():
        return p

    # 2) 按已知文件名表找
    wanted = FONT_FILES.get(name, [])
    dirs = _font_dirs()
    for fn in wanted:
        low = fn.lower()
        for d in dirs:
            direct = d / fn
            if direct.exists():
                return direct
            try:
                for got in d.rglob("*"):
                    if got.is_file() and got.name.lower() == low:
                        return got
            except OSError:
                continue

    # 3) 问 fontconfig（Linux 上装的是别名也能命中）
    got = _fc_match(name)
    if got:
        return got

    # 4) Windows 注册表能给出准确文件名
    if sys.platform.startswith("win"):
        try:
            import winreg

            key = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key) as k:
                for i in range(winreg.QueryInfoKey(k)[1]):
                    vname, vdata, _ = winreg.EnumValue(k, i)
                    if name in vname:
                        f = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts" / vdata
                        if f.exists():
                            return f
        except Exception:  # noqa: BLE001
            pass
    return None


def resolve_or_fallback(name: str) -> Path | None:
    """解析字体；实在找不到就拿一份通用中文宋/仿宋顶上。"""
    got = resolve(name)
    if got:
        return got
    for fn in FALLBACK:
        for d in _font_dirs():
            p = d / fn
            if p.exists():
                return p
            try:
                for g in d.rglob(fn):
                    if g.is_file():
                        return g
            except OSError:
                continue
    # 最后问 fontconfig 要一个能显示中文的
    for fam in ("sans-serif:lang=zh", "Noto Sans CJK SC", "WenQuanYi Zen Hei"):
        got = _fc_match(fam)
        if got:
            return got
    return None


def subset(src: Path, text: str, out: Path) -> Path | None:
    """把 src 裁成只剩 text 里用到的字，写到 out。失败返回 None。

    .ttc 是字体集合，要先取出第 0 号再当普通字体处理，否则 fontTools 报
    TTLibFileIsCollectionError。
    """
    try:
        from fontTools import subset as ftsubset
    except ImportError:
        return None

    try:
        use_src = src
        if src.suffix.lower() == ".ttc":
            from fontTools.ttLib import TTFont

            coll = TTFont(str(src), fontNumber=0, lazy=False)
            tmp = out.with_suffix(".base.ttf")
            coll.save(str(tmp))
            coll.close()
            use_src = tmp

        opts = ftsubset.Options()
        opts.layout_features = ["*"]
        opts.notdef_outline = True
        opts.recalc_bounds = True
        opts.drop_tables += ["meta"]

        font = ftsubset.load_font(str(use_src), opts)
        try:
            sub = ftsubset.Subsetter(options=opts)
            sub.populate(text=text)
            sub.subset(font)
            font.save(str(out))
        finally:
            font.close()
        if use_src != src:
            try:
                use_src.unlink()
            except OSError:
                pass
        return out if out.exists() and out.stat().st_size > 0 else None
    except Exception:  # noqa: BLE001
        return None
