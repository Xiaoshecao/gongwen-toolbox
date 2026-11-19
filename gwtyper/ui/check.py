# -*- coding: utf-8 -*-
"""勾选项。

ttk 的 Checkbutton 在 clam 主题下会把选中态画成一个「✗」（indicator 由
element 直接描叉），看起来像报错。这里不用 ttk 的 indicator，改用字符
☑ / ☐ 自绘，字形在任何中文字体里都有，且和文字行高对齐。
"""
from __future__ import annotations

import tkinter as tk
from typing import Callable

from . import theme

ON = "\u2611"    # ☑
OFF = "\u2610"   # ☐


class CheckRow(tk.Frame):
    """一行可点的勾选：☑ 文字。点整行都能切换。"""

    def __init__(self, master, text: str, variable: tk.BooleanVar,
                 command: Callable[[], None] | None = None,
                 font=None, bg: str = theme.SURFACE,
                 fg: str = theme.TEXT, padx=0, pady=0):
        super().__init__(master, background=bg, bd=0, highlightthickness=0,
                         padx=padx, pady=pady, cursor="hand2")
        self.var = variable
        self.command = command
        self._bg = bg
        body = (font or ("Microsoft YaHei UI", 11))
        self.box = tk.Label(self, text="", font=body, bg=bg, fg=theme.PRIMARY,
                            bd=0, padx=0, pady=0)
        self.box.pack(side="left")
        self.lbl = tk.Label(self, text=text, font=body, bg=bg, fg=fg,
                            bd=0, padx=0, pady=0)
        self.lbl.pack(side="left", padx=(6, 0))

        for w in (self, self.box, self.lbl):
            w.bind("<Button-1>", self._toggle)
            w.bind("<Enter>", self._hover_on)
            w.bind("<Leave>", self._hover_off)
        self.bind("<space>", self._toggle)
        self.configure(takefocus=1)
        self.bind("<FocusIn>", lambda _e: self.configure(
            highlightthickness=1, highlightbackground=theme.FOCUS,
            highlightcolor=theme.FOCUS))
        self.bind("<FocusOut>", lambda _e: self.configure(highlightthickness=0))
        self._sync()

    # ------------------------------------------------------------------
    def _sync(self) -> None:
        self.box.configure(text=ON if self.var.get() else OFF,
                           fg=theme.PRIMARY if self.var.get() else theme.BORDER_STRONG)

    def _hover_on(self, _e=None) -> None:
        for w in (self, self.box, self.lbl):
            w.configure(background=theme.HOVER)

    def _hover_off(self, _e=None) -> None:
        for w in (self, self.box, self.lbl):
            w.configure(background=self._bg)

    def set(self, value: bool) -> None:
        self.var.set(value)
        self._sync()

    def get(self) -> bool:
        return bool(self.var.get())

    def set_enabled(self, on: bool) -> None:
        """置灰但不改值：切回来时用户原来的勾选还在。"""
        fg = theme.TEXT if on else theme.TEXT3
        self.lbl.configure(fg=fg)
        self.box.configure(fg=(theme.PRIMARY if self.var.get() else theme.BORDER_STRONG)
                           if on else theme.TEXT3)
        self.configure(cursor="hand2" if on else "")
        self._enabled = on

    def _toggle(self, _e=None) -> None:
        if not getattr(self, "_enabled", True):
            return "break"
        self.var.set(not self.var.get())
        self._sync()
        if self.command:
            self.command()
        return "break"
