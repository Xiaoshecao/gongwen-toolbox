# -*- coding: utf-8 -*-
"""界面主题：一套给 ttk 定制的视觉规范。

ttk 没有圆角、阴影、行距，层次只能靠三件事建立：
底色差（≥5% 明度）+ 1px 描边 + 内边距节奏。下面的数值都为此服务。
"""
from __future__ import annotations

import sys
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

IS_WIN = sys.platform.startswith("win")

# ---------------------------------------------------------------- 令牌
# 表面
CANVAS = "#EDEFF3"      # 窗口底色、栏间空隙
SURFACE = "#FFFFFF"     # 卡片、表格、纸张
SUNKEN = "#F6F7F9"      # 表头带、只读区
INSET = "#E3E7ED"       # 纸张预览凹槽
HOVER = "#F3F5F9"       # 悬停
ACTIVE = "#E9EDF4"      # 按下
ROW_EVEN = "#FAFBFC"    # 斑马纹

# 描边
BORDER_SUBTLE = "#E4E8EE"
BORDER = "#D3D9E2"
BORDER_STRONG = "#B9C2CF"

# 文字三级（对白底 16:1 / 5.8:1 / 3.1:1）
TEXT = "#1B2230"
TEXT2 = "#5A6675"
TEXT3 = "#8A94A3"
TEXT_INV = "#FFFFFF"

# 主色
PRIMARY = "#2563EB"
PRIMARY_HOVER = "#1D4ED8"
PRIMARY_PRESS = "#1E40AF"
PRIMARY_SOFT = "#E8EFFC"
FOCUS = "#7EA6F2"
SEL_BG = "#DCE7FB"
SEL_FG = "#12305F"
SEL_BG_OFF = "#E8ECF2"
SEL_FG_OFF = "#3A4553"

# 语义（填充 / 浅底 / 文字）
OK = "#0E9F6E"
OK_TX = "#0A7A54"
OK_SOFT = "#E6F6F0"
WARN = "#D97706"
WARN_TX = "#9A5B08"
WARN_SOFT = "#FDF3E4"
ERR = "#DC2626"
ERR_TX = "#B01C1C"
ERR_SOFT = "#FDECEC"

# 旧名兼容（其它模块仍可能引用）
BG = CANVAS
CARD = SURFACE
INK = TEXT
MUTED = TEXT2
LINE = BORDER
BRAND = PRIMARY
BRAND_DARK = PRIMARY_HOVER
BRAND_SOFT = PRIMARY_SOFT

SCALE = 1.0


FAMILY = "Microsoft YaHei UI"


def pick_family(root: tk.Misc) -> str:
    """挑一个本机有的中文字体，Windows / Linux / macOS 都能看。"""
    families = set(tkfont.families(root))
    for name in ("Microsoft YaHei UI", "微软雅黑", "Microsoft YaHei",
                 "PingFang SC", "Hiragino Sans GB",
                 "Noto Sans CJK SC", "Noto Sans SC", "Source Han Sans SC",
                 "WenQuanYi Micro Hei", "WenQuanYi Zen Hei",
                 "Droid Sans Fallback", "SimHei", "DejaVu Sans"):
        if name in families:
            return name
    return "TkDefaultFont"


def enable_dpi_awareness() -> None:
    """Windows 高分屏下先声明 DPI 感知，否则整个界面被系统拉伸糊掉。"""
    if not IS_WIN:
        return
    try:
        from ctypes import windll

        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:  # noqa: BLE001
        pass


_ROW_H_COMPACT = 20
_ROW_H_NORMAL = 26


def apply(root: tk.Tk, density: str = "normal") -> dict[str, tkfont.Font]:
    """density: normal / compact。compact 会把表格行高压低、间距收紧，
    给 1366×768 这类小屏多腾出内容空间。"""
    global SCALE
    global FAMILY
    fam = pick_family(root)
    FAMILY = fam
    try:
        fp = float(root.winfo_fpixels("1i"))
    except Exception:  # noqa: BLE001
        fp = 96.0
    if fp > 0:
        root.tk.call("tk", "scaling", fp / 72.0)
    SCALE = (fp / 96.0) if fp else 1.0

    f = {
        "h1": tkfont.Font(family=fam, size=16, weight="bold"),
        "h2": tkfont.Font(family=fam, size=11, weight="bold"),
        "h3": tkfont.Font(family=fam, size=10, weight="bold"),
        "body": tkfont.Font(family=fam, size=10),
        "cap": tkfont.Font(family=fam, size=9),
        "mic": tkfont.Font(family=fam, size=8),
        "btn": tkfont.Font(family=fam, size=10),
        "num": tkfont.Font(family="Segoe UI" if IS_WIN else fam, size=10),
    }

    st = ttk.Style(root)
    try:
        st.theme_use("clam")     # 唯一在 Win/Linux 表现一致的底座
    except tk.TclError:
        pass

    st.configure(".", background=CANVAS, foreground=TEXT, font=f["body"],
                 borderwidth=0, focuscolor=FOCUS)

    # ---- 容器 ----
    st.configure("TFrame", background=CANVAS)
    st.configure("Canvas.TFrame", background=CANVAS)
    st.configure("Card.TFrame", background=SURFACE, relief="flat", borderwidth=0)

    # ---- 文字 ----
    st.configure("H1.TLabel", font=f["h1"], foreground=TEXT, background=CANVAS)
    st.configure("H2.TLabel", font=f["h2"], foreground=TEXT, background=SURFACE)
    st.configure("H3.TLabel", font=f["h3"], foreground="#2B3542", background=SURFACE)
    st.configure("Body.TLabel", font=f["body"], foreground=TEXT, background=SURFACE)
    st.configure("Cap.TLabel", font=f["cap"], foreground=TEXT2, background=SURFACE)
    st.configure("Mic.TLabel", font=f["mic"], foreground=TEXT3, background=SURFACE)
    st.configure("Muted.TLabel", font=f["cap"], foreground=TEXT2, background=SURFACE)
    st.configure("Status.TLabel", font=f["cap"], foreground=TEXT2, background=CANVAS)
    st.configure("Num.TLabel", font=f["num"], foreground=TEXT, background=SURFACE)
    st.configure("Err.TLabel", font=f["cap"], foreground=ERR_TX, background=SURFACE)
    st.configure("Warn.TLabel", font=f["cap"], foreground=WARN_TX, background=SURFACE)
    st.configure("Ok.TLabel", font=f["cap"], foreground=OK_TX, background=SURFACE)
    st.configure("Brand.TLabel", font=f["body"], foreground=PRIMARY, background=SURFACE)

    # ---- 按钮：三档（实心 / 白底描边 / 无底） ----
    # relief 必须是 solid：clam 主题的 Button.border 元素在 relief="flat"
    # 时压根不画边（实测描边色像素 0 个），白底按钮贴在白卡上就只剩一行字。
    def button(name, fill, fill_hv, fill_pr, fill_dis, edge, edge_hv, edge_dis,
               fg, fg_dis, pad, font, bw=1, relief="solid"):
        st.configure(name, background=fill, foreground=fg, bordercolor=edge,
                     lightcolor=fill, darkcolor=edge, borderwidth=bw,
                     relief=relief, padding=pad, font=font, anchor="center")
        st.map(name,
               background=[("disabled", fill_dis), ("pressed", fill_pr), ("active", fill_hv)],
               lightcolor=[("disabled", fill_dis), ("pressed", fill_pr), ("active", fill_hv)],
               bordercolor=[("disabled", edge_dis), ("pressed", fill_pr), ("active", edge_hv)],
               darkcolor=[("disabled", edge_dis), ("pressed", fill_pr), ("active", edge_hv)],
               foreground=[("disabled", fg_dis)])

    pad_y = int(round(7 * SCALE))
    button("Primary.TButton", PRIMARY, PRIMARY_HOVER, PRIMARY_PRESS, "#C7D2E4",
           PRIMARY_HOVER, PRIMARY_HOVER, "#C7D2E4",
           TEXT_INV, "#F2F5FA", (int(16 * SCALE), pad_y), f["btn"])
    button("Accent.TButton", PRIMARY, PRIMARY_HOVER, PRIMARY_PRESS, "#C7D2E4",
           PRIMARY_HOVER, PRIMARY_HOVER, "#C7D2E4",
           TEXT_INV, "#F2F5FA", (int(16 * SCALE), pad_y), f["btn"])
    button("Secondary.TButton", SURFACE, HOVER, ACTIVE, SUNKEN,
           BORDER, BORDER_STRONG, BORDER_SUBTLE,
           "#2B3542", "#A8B1BF", (int(14 * SCALE), pad_y), f["body"])
    # Ghost 是「无底无框」，故意保留 flat —— 它本来就该融进背景
    button("Ghost.TButton", SURFACE, "#EDF0F5", "#E2E7EE", SURFACE,
           SURFACE, "#EDF0F5", SURFACE,
           "#4A5566", "#A8B1BF", (int(10 * SCALE), int(6 * SCALE)), f["body"],
           relief="flat")
    button("TButton", SURFACE, HOVER, ACTIVE, SUNKEN,
           BORDER, BORDER_STRONG, BORDER_SUBTLE,
           "#2B3542", "#A8B1BF", (int(14 * SCALE), pad_y), f["body"])


    # ---- 输入 ----
    st.configure("TEntry", fieldbackground=SURFACE, foreground=TEXT,
                 bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER,
                 insertcolor=TEXT, padding=(8, 5), borderwidth=1)
    st.map("TEntry", bordercolor=[("focus", FOCUS)],
           lightcolor=[("focus", FOCUS)], darkcolor=[("focus", FOCUS)])

    st.configure("TCombobox", fieldbackground=SURFACE, background=SURFACE,
                 foreground=TEXT, bordercolor=BORDER, lightcolor=BORDER,
                 darkcolor=BORDER, arrowcolor=TEXT2, padding=(8, 5),
                 selectbackground=SURFACE, selectforeground=TEXT)
    st.map("TCombobox", fieldbackground=[("readonly", SURFACE)],
           foreground=[("readonly", TEXT)],
           bordercolor=[("focus", FOCUS)],
           arrowcolor=[("active", TEXT)])
    for opt, val in (("*TCombobox*Listbox.background", SURFACE),
                     ("*TCombobox*Listbox.foreground", TEXT),
                     ("*TCombobox*Listbox.selectBackground", SEL_BG),
                     ("*TCombobox*Listbox.selectForeground", SEL_FG),
                     ("*TCombobox*Listbox.font", f["body"])):
        root.option_add(opt, val)

    # ---- 勾选 / 单选 ----
    for nm in ("TCheckbutton", "Card.TCheckbutton", "TRadiobutton", "Card.TRadiobutton"):
        st.configure(nm, background=SURFACE, foreground=TEXT, font=f["body"],
                     indicatorcolor=SURFACE, indicatorbackground=SURFACE,
                     indicatorsize=int(round(13 * SCALE)),
                     indicatormargin=(0, 0, int(round(8 * SCALE)), 0),
                     bordercolor=BORDER_STRONG, focuscolor=FOCUS, padding=(2, 4))
        st.map(nm, background=[("active", SURFACE)],
               foreground=[("active", TEXT), ("disabled", TEXT3)],
               bordercolor=[("active", PRIMARY), ("selected", PRIMARY)],
               indicatorcolor=[("selected", PRIMARY), ("active", PRIMARY_HOVER),
                               ("disabled", INSET)])
    for nm in ("Canvas.TCheckbutton", "Canvas.TRadiobutton"):
        st.configure(nm, background=CANVAS, foreground=TEXT, font=f["body"],
                     indicatorcolor=SURFACE, indicatorbackground=SURFACE,
                     bordercolor=BORDER, focuscolor=FOCUS, padding=(2, 3))
        st.map(nm, background=[("active", CANVAS)],
               indicatorcolor=[("selected", PRIMARY)])

    # ---- 表格 ----
    st.configure("Treeview", background=SURFACE, fieldbackground=SURFACE,
                 foreground=TEXT,
                 # 行高别跟着 DPI 线性放大：150% 缩放下 30×1.5=45px，
                 # 一张表 9 行就吃掉 450px，1366×768 的笔记本直接放不下。
                 # 正常密度封顶 26px；紧凑密度再压到 20px，多腾一屏出来。
                 rowheight=(_ROW_H_COMPACT if density == "compact"
                            else min(26, int(round(24 * SCALE)))),
                 borderwidth=0, relief="flat", font=f["body"])
    st.map("Treeview",
           background=[("selected", SEL_BG)],
           foreground=[("selected", SEL_FG)])
    st.configure("Treeview.Heading", background=SUNKEN, foreground=TEXT2,
                 font=f["cap"], relief="flat", borderwidth=0,
                 padding=(int(10 * SCALE), int(7 * SCALE)))
    st.map("Treeview.Heading",
           background=[("active", HOVER)],
           foreground=[("active", TEXT)])

    # ---- 滚动条 ----
    for orient in ("Vertical", "Horizontal"):
        st.configure(f"{orient}.TScrollbar", background=SUNKEN, troughcolor=CANVAS,
                     bordercolor=CANVAS, arrowcolor=TEXT2, relief="flat",
                     borderwidth=0, arrowsize=int(12 * SCALE))
        st.map(f"{orient}.TScrollbar",
               background=[("active", BORDER), ("pressed", BORDER_STRONG)])

    # ---- 进度条 ----
    st.configure("TProgressbar", background=PRIMARY, troughcolor=CANVAS,
                 bordercolor=CANVAS, lightcolor=PRIMARY, darkcolor=PRIMARY,
                 thickness=int(4 * SCALE), borderwidth=0)

    # ---- 页签（规则集编辑器用） ----
    st.configure("TNotebook", background=CANVAS, borderwidth=0, tabmargins=(0, 0, 0, 0))
    st.configure("TNotebook.Tab", font=f["body"],
                 padding=(int(18 * SCALE), int(9 * SCALE)),
                 background=CANVAS, foreground=TEXT2, borderwidth=0)
    st.map("TNotebook.Tab",
           # 选中的页签：白底 + 品牌蓝字 + 稍微加高，一眼能看出在哪一页
           background=[("selected", SURFACE), ("active", HOVER)],
           foreground=[("selected", PRIMARY), ("active", TEXT)],
           expand=[("selected", (0, 2, 0, 0))])

    # ---- 分隔线 ----
    st.configure("TSeparator", background=BORDER_SUBTLE)
    return f


def stripe(tree: ttk.Treeview) -> None:
    """斑马纹。Tk 8.6 里 tag 的 background 会盖掉选中色，所以选中行要去掉斑马标签。"""
    tree.tag_configure("even", background=ROW_EVEN)

    def apply_stripes(_e=None) -> None:
        sel = set(tree.selection())
        for i, iid in enumerate(tree.get_children("")):
            tags = [t for t in tree.item(iid, "tags") if t != "even"]
            if i % 2 and iid not in sel:
                tags.append("even")
            tree.item(iid, tags=tags)

    tree.bind("<<TreeviewSelect>>", apply_stripes, add="+")
    tree._stripe = apply_stripes          # 刷新完表格后可以再调一次
    apply_stripes()   # 建表时就上色，不用等用户点一下
