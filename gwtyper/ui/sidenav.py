# -*- coding: utf-8 -*-
"""左侧竖排菜单 + 右侧内容区。

用户要求：功能不要横着排在顶上，改成左边竖着一条菜单，并按类别分组。
`NavShim` 故意做成"长得像 ttk.Notebook"（有 add/select/tabs），
这样 app.py 里原来的写法几乎不用改。
"""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from . import theme


class SideNav(ttk.Frame):
    def __init__(self, master, on_select=None, width: int = 190):
        super().__init__(master, style="Canvas.TFrame")
        self._on_select = on_select
        self._items: list[dict] = []
        self._rows: dict[int, dict] = {}
        self._current = -1
        self._last_group = None
        self.configure(width=width)
        self._canvas = tk.Canvas(self, background=theme.CANVAS, highlightthickness=0, bd=0,
                                 width=width)
        self._sb = ttk.Scrollbar(self, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=self._sb.set)
        self._canvas.pack(side="left", fill="both", expand=True)
        self._sb.pack(side="right", fill="y")
        self._box = tk.Frame(self._canvas, background=theme.CANVAS)
        self._win = self._canvas.create_window((0, 0), window=self._box, anchor="nw")
        self._box.bind("<Configure>",
                       lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>",
                          lambda e: self._canvas.itemconfigure(self._win, width=e.width))
        self._canvas.bind("<Enter>", lambda e: self._canvas.bind_all("<MouseWheel>", self._wheel))
        self._canvas.bind("<Leave>", lambda e: self._canvas.unbind_all("<MouseWheel>"))

    def _wheel(self, event):
        try:
            self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:  # noqa: BLE001
            pass

    # ---------------------------------------------------------------- add
    def add(self, frame, text: str, group: str | None = None) -> None:
        idx = len(self._items)
        self._items.append({"text": text, "group": group, "frame": frame})

        if group and group != self._last_group:
            self._last_group = group
            lbl = tk.Label(self._box, text=group, bg=theme.CANVAS, fg=theme.TEXT3,
                           font=(theme.FAMILY, 9), anchor="w", padx=12, pady=0)
            lbl.pack(fill="x", pady=(14, 4))

        row = tk.Frame(self._box, background=theme.CANVAS, cursor="hand2")
        row.pack(fill="x", padx=8, pady=1)
        bar = tk.Frame(row, background=theme.CANVAS, width=3)
        bar.pack(side="left", fill="y")
        label = tk.Label(row, text=text, bg=theme.CANVAS, fg=theme.TEXT2,
                         font=(theme.FAMILY, 10), anchor="w", padx=10, pady=7)
        label.pack(side="left", fill="x", expand=True)

        for w in (row, label):
            w.bind("<Button-1>", lambda _e, i=idx: self.select(i))
        row.bind("<Enter>", lambda _e, i=idx: self._hover(i, True))
        row.bind("<Leave>", lambda _e, i=idx: self._hover(i, False))
        self._rows[idx] = {"row": row, "bar": bar, "label": label}

        if idx == 0:
            self.select(0)

    # ---------------------------------------------------------------- 样式
    def _paint(self, idx: int, selected: bool, hover: bool = False) -> None:
        r = self._rows.get(idx)
        if not r:
            return
        if selected:
            bg, fg, bar = theme.PRIMARY_SOFT, theme.PRIMARY, theme.PRIMARY
        elif hover:
            bg, fg, bar = theme.HOVER, theme.TEXT, theme.CANVAS
        else:
            bg, fg, bar = theme.CANVAS, theme.TEXT2, theme.CANVAS
        for w in (r["row"], r["label"]):
            w.configure(background=bg)
        r["label"].configure(foreground=fg)
        r["bar"].configure(background=bar)

    def _hover(self, idx: int, on: bool) -> None:
        if idx != self._current:
            self._paint(idx, False, hover=on)

    # ---------------------------------------------------------------- select
    def select(self, which) -> None:
        if isinstance(which, int):
            idx = which
        else:
            idx = next((i for i, it in enumerate(self._items)
                        if it["frame"] is which or it["text"] == which), -1)
        if idx < 0 or idx >= len(self._items):
            return
        old = self._current
        self._current = idx
        if old >= 0:
            self._paint(old, False)
        self._paint(idx, True)
        frame = self._items[idx]["frame"]
        try:
            frame.tkraise()
        except Exception:  # noqa: BLE001
            pass
        if self._on_select:
            self._on_select(self._items[idx]["text"])

    def tabs(self) -> list[str]:
        return [it["text"] for it in self._items]

    def current_text(self) -> str:
        if 0 <= self._current < len(self._items):
            return self._items[self._current]["text"]
        return ""


class NavShim(tk.Frame):
    """既是右侧内容容器，又提供 Notebook 风格的 add/select/tabs。"""

    def __init__(self, master, nav: SideNav, **kw):
        super().__init__(master, background=theme.CANVAS, **kw)
        self._nav = nav
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self._frames: list = []

    def add(self, frame, text: str = "", group: str | None = None):
        frame.grid(row=0, column=0, sticky="nsew")
        self._frames.append(frame)
        self._nav.add(frame, text, group)
        return frame

    def select(self, frame) -> None:
        self._nav.select(frame)

    def tabs(self) -> list[str]:
        return self._nav.tabs()
