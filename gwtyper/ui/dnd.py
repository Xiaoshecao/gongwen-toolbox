# -*- coding: utf-8 -*-
"""统一的拖拽落点处理：三处（排版页 / 格式转换 / OCR）共用一套。

踩过的坑：
- tkdnd 给的 `event.data` 是 Tcl 列表，含空格或中文的路径会被花括号包住；
- 但有的拖拽源直接给裸路径（没有花括号），这时 `tk.splitlist` 会按空格
  把一条路径切碎 → 所以先把整串当成一条路径试，命中文件系统就用它。
"""
from __future__ import annotations

import logging
from pathlib import Path

log = logging.getLogger(__name__)


def paths_from_event(widget, event, accept=None, recursive=True) -> list[Path]:
    """把一次拖拽事件解析成文件路径列表。

    accept(path) -> bool 可选，用来筛后缀；文件夹会展开成里面的文件。
    """
    raw = str(getattr(event, "data", "") or "").strip()
    if not raw:
        return []

    whole = Path(raw.strip("{}"))
    if whole.exists():
        items = [raw.strip("{}")]
    else:
        try:
            items = [str(x) for x in widget.tk.splitlist(raw)]
        except Exception:  # noqa: BLE001
            items = raw.split()

    out: list[Path] = []
    for it in items:
        p = Path(str(it).strip().strip("{}"))
        try:
            if p.is_dir():
                files = p.rglob("*") if recursive else p.iterdir()
                out.extend(sorted(x for x in files
                                  if x.is_file() and (accept is None or accept(x))))
            elif p.is_file() and (accept is None or accept(p)):
                out.append(p)
        except OSError as exc:
            log.warning("拖进来的路径读不了：%s（%s）", p, exc)
    return out


def bind_targets(widgets, handler) -> bool:
    """给若干控件注册文件拖放。返回是否至少成功注册一个。"""
    try:
        from tkinterdnd2 import DND_FILES
    except Exception:  # noqa: BLE001
        return False
    ok = False
    for w in widgets:
        if w is None:
            continue
        try:
            w.drop_target_register(DND_FILES)
            w.dnd_bind("<<Drop>>", handler)
            ok = True
        except Exception:  # noqa: BLE001
            continue
    return ok


def suffix_filter(exts):
    exts = {e.lower() for e in exts}
    return lambda p: p.suffix.lower() in exts
