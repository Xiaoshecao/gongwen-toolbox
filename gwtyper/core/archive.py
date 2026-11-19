# -*- coding: utf-8 -*-
"""按台账把收到的材料归档到「单位名」子文件夹。"""
from __future__ import annotations

import shutil
from pathlib import Path

from . import villages


def archive(files, names: list[str], out_dir, move: bool = False,
            progress=None) -> dict:
    """把 files 按名单归到 out_dir/<单位>/。

    返回 {"done": {单位: [文件]}, "unknown": [没认出来的文件], "out": 目录}
    """
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    done: dict[str, list[str]] = {}
    unknown: list[str] = []
    files = [Path(f) for f in files]
    total = len(files)
    for i, f in enumerate(files, 1):
        hit = None
        for n in names:
            if any(a and a in villages._key(f.stem) for a in villages._aliases(n)):
                hit = n
                break
        if not hit:
            unknown.append(str(f))
        else:
            dst_dir = out / hit
            dst_dir.mkdir(parents=True, exist_ok=True)
            dst = dst_dir / f.name
            k = 1
            while dst.exists():
                dst = dst_dir / f"{f.stem}({k}){f.suffix}"
                k += 1
            try:
                if move:
                    shutil.move(str(f), str(dst))
                else:
                    shutil.copy2(str(f), str(dst))
                done.setdefault(hit, []).append(str(dst))
            except OSError as exc:
                unknown.append(f"{f}（复制失败：{exc}）")
        if progress:
            try:
                progress(i, total)
            except Exception:  # noqa: BLE001
                pass
    return {"done": done, "unknown": unknown, "out": str(out)}
