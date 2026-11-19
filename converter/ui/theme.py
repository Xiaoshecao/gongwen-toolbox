# -*- coding: utf-8 -*-
"""界面主题：颜色、字体、控件样式统一在这里，方便整体调色。"""
from __future__ import annotations

import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

BG = "#F4F6FA"
CARD = "#FFFFFF"
INK = "#1F2733"
MUTED = "#7A8698"
LINE = "#E2E7F0"
BRAND = "#2563EB"
BRAND_DARK = "#1D4ED8"
BRAND_SOFT = "#E8F0FE"
OK = "#0E9F6E"
WARN = "#D97706"
ERR = "#DC2626"


def pick_family(root: tk.Misc) -> str:
    """优先用系统里的中文字体，避免出现方块。"""
    families = set(tkfont.families(root))
    for name in ("Microsoft YaHei UI", "微软雅黑", "Microsoft YaHei", "PingFang SC", "SimHei"):
        if name in families:
            return name
    return "TkDefaultFont"


def apply(root: tk.Tk) -> dict[str, tkfont.Font]:
    fam = pick_family(root)
    root.configure(bg=BG)
    fonts = {
        "h1": tkfont.Font(family=fam, size=16, weight="bold"),
        "h2": tkfont.Font(family=fam, size=11, weight="bold"),
        "body": tkfont.Font(family=fam, size=10),
        "small": tkfont.Font(family=fam, size=9),
        "mono": tkfont.Font(family="Consolas" if "Consolas" in tkfont.families(root) else fam, size=9),
        "btn": tkfont.Font(family=fam, size=10, weight="bold"),
    }

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    style.configure(".", background=BG, foreground=INK, font=fonts["body"])
    style.configure("TFrame", background=BG)
    style.configure("Card.TFrame", background=CARD, relief="flat")
    style.configure("TLabel", background=BG, foreground=INK, font=fonts["body"])
    style.configure("Card.TLabel", background=CARD, foreground=INK, font=fonts["body"])
    style.configure("Muted.TLabel", background=CARD, foreground=MUTED, font=fonts["small"])
    style.configure("H1.TLabel", background=BG, foreground=INK, font=fonts["h1"])
    style.configure("H2.TLabel", background=CARD, foreground=INK, font=fonts["h2"])
    style.configure("Status.TLabel", background=BG, foreground=MUTED, font=fonts["small"])

    style.configure(
        "TButton",
        background="#FFFFFF",
        foreground=INK,
        font=fonts["body"],
        borderwidth=1,
        focusthickness=0,
        padding=(12, 7),
    )
    style.map("TButton", background=[("active", "#EEF2F9")], relief=[("pressed", "flat")])
    style.configure(
        "Accent.TButton",
        background=BRAND,
        foreground="#FFFFFF",
        font=fonts["btn"],
        borderwidth=0,
        padding=(16, 9),
    )
    style.map("Accent.TButton", background=[("active", BRAND_DARK), ("disabled", "#9DB6EE")])
    style.configure(
        "Ghost.TButton",
        background="#FFFFFF",
        foreground=INK,
        font=fonts["body"],
        borderwidth=1,
        padding=(10, 6),
    )

    style.configure(
        "TEntry",
        fieldbackground="#FFFFFF",
        bordercolor=LINE,
        lightcolor=LINE,
        darkcolor=LINE,
        padding=6,
    )
    style.configure("TCombobox", fieldbackground="#FFFFFF", background="#FFFFFF", padding=5)
    style.map("TCombobox", fieldbackground=[("readonly", "#FFFFFF")])
    style.configure("TCheckbutton", background=CARD, foreground=INK, font=fonts["body"])
    style.map("TCheckbutton", background=[("active", CARD)])
    style.configure("TRadiobutton", background=CARD, foreground=INK, font=fonts["body"])

    style.configure(
        "Treeview",
        background="#FFFFFF",
        fieldbackground="#FFFFFF",
        foreground=INK,
        rowheight=26,
        borderwidth=0,
        font=fonts["body"],
    )
    style.configure(
        "Treeview.Heading",
        background="#F7F9FC",
        foreground=MUTED,
        font=fonts["small"],
        relief="flat",
        padding=(6, 6),
    )
    style.map("Treeview.Heading", background=[("active", "#EEF2F9")])
    style.map("Treeview", background=[("selected", BRAND_SOFT)], foreground=[("selected", INK)])

    style.configure(
        "Brand.Horizontal.TProgressbar",
        troughcolor="#E9EEF7",
        background=BRAND,
        borderwidth=0,
        thickness=8,
    )
    style.configure("TNotebook", background=BG, borderwidth=0)
    style.configure("TNotebook.Tab", padding=(16, 8), font=fonts["body"], background="#E9EEF6")
    style.map(
        "TNotebook.Tab",
        background=[("selected", "#FFFFFF")],
        foreground=[("selected", BRAND)],
    )
    style.configure("TSeparator", background=LINE)
    return fonts
