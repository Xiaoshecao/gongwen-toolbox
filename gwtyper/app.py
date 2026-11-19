# -*- coding: utf-8 -*-
"""公文排版工具 界面。

两个动作分开：检查（只读）/ 排版（写副本）。中间一栏是段落角色，可手工改。
"""
from __future__ import annotations

import logging
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .core import office, sysenv
from .core import rules as rulesmod
from .core.check import check, check_features, summary
from .core.classify import classify
from .core.model import ROLE_CN, Doc, Role
from .core.parse_docx import parse

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    _HAS_DND = True
except Exception:  # 拖拽组件可选
    _HAS_DND = False

from .ui import theme
from .ui import pump
from .ui import dnd
from .ui.check import CheckRow
from .ui.convert_pane import ConvertPane
from .ui.sidenav import NavShim, SideNav
from .ui.quick_panes import DupPane, ImageCompressPane
from .ui.pdf_tools_pane import PdfToolsPane
from .ui.template_pane import TemplatePane
from .ui.batch_panes import (ListDiffPane, MaskPane, MergeTablePane,
                             PrintCheckPane, WatermarkPane)
from .ui.ocr_pane import OcrPane
from .ui.tools_pane import ToolsPane
from .ui.checklist_pane import ChecklistPane

APP_NAME = "公文工具箱"
APP_VER = "3.4"
ROLE_CHOICES = [r.value for r in Role if r != Role.BLANK]


def _center_dialog(win: tk.Toplevel) -> None:
    """按内容实际大小定尺寸并居中到父窗口。"""
    win.update_idletasks()
    w = max(360, win.winfo_reqwidth())
    h = max(200, win.winfo_reqheight())
    try:
        px, py = win.master.winfo_rootx(), win.master.winfo_rooty()
        pw, ph = win.master.winfo_width(), win.master.winfo_height()
        x = px + max(0, (pw - w) // 2)
        y = py + max(0, (ph - h) // 3)
    except Exception:  # noqa: BLE001
        x = y = 200
    win.geometry(f"{w}x{h}+{x}+{y}")


def _parallel(items, fn, workers: int | None = None):
    """并行执行 fn(item)，结果由 fn 自己收。数目少时直接串行，省线程开销。"""
    items = list(items)
    if not items:
        return
    if len(items) == 1:
        fn(items[0])
        return
    import os
    from concurrent.futures import ThreadPoolExecutor

    n = workers or min(6, (os.cpu_count() or 2), len(items))
    with ThreadPoolExecutor(max_workers=max(1, n)) as ex:
        list(ex.map(fn, items))

# 规则集的展示名：把 gb9704 这类代号翻成人话
PROFILE_CN = {
    "gb9704": "国家标准 GB/T 9704 公文格式",
    "report": "述职/总结报告",
    "news": "新闻稿/通讯",
}
PROFILE_TIP = {
    "gb9704": "党政机关公文格式，红头文件、通知、请示都用这套",
    "report": "述职述廉、工作总结一类，页边距更宽、无页码",
    "news": "宣传稿、通讯，标题两行、正文夹黑体小标题",
}
CONF_TIERS = ((0.75, "确定"), (0.55, "可能"), (0.0, "拿不准"))


def conf_text(c: float) -> str:
    for floor, name in CONF_TIERS:
        if c >= floor:
            return name
    return "拿不准"


def profile_label(key: str) -> str:
    return PROFILE_CN.get(key, key)


def open_path(path: str) -> None:
    sysenv.open_with_default(path)


def open_folder(path: str) -> None:
    sysenv.open_folder(path)


def _fix_text(f) -> str:
    """把「现在」和「应为」拼成一列。

    窄屏上分成两列时后半截（该改成什么）会被切掉，
    而「该改成什么」恰恰是用户唯一关心的。
    两值相同时只写一次，免得出现「仿宋 → 仿宋」这种看着像没检查的行。
    """
    now = (f.actual or "").strip()
    want = (f.expected or "").strip()
    if not want:
        return now
    if not now or now == want:
        return want
    return f"{now} → {want}"


def _pages_note(shown: list | int, total: int) -> str:
    """预览页数提示。只渲了前几页时说清楚，别让人以为文件只有这么点。"""
    got = len(shown) if isinstance(shown, list) else int(shown)
    if got < total:
        return f"{total} 页（预览前 {got} 页）"
    return f"{total} 页"


def work_area() -> tuple[int, int, int, int]:
    """可用桌面区域（已扣任务栏）。取不到就退回整屏。"""
    try:
        import ctypes
        from ctypes import wintypes

        rect = wintypes.RECT()
        if ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0):
            w, h = rect.right - rect.left, rect.bottom - rect.top
            if w > 400 and h > 300:
                return rect.left, rect.top, w, h
    except Exception:  # noqa: BLE001
        pass
    return 0, 0, 0, 0


class Base(tk.Tk if not _HAS_DND else TkinterDnD.Tk):  # type: ignore[misc]
    pass


class App(Base):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"{APP_NAME} {APP_VER}")
        # 任何界面回调里抛出的异常都要看得见：Tk 默认只往 stderr 打，
        # 而窗口版 exe 根本没有 stderr，结果就是「点了按钮毫无反应」。
        self.report_callback_exception = self._on_callback_error
        # 布局密度：政府电脑现在多是 1920×1080，默认就用常规密度；
        # 老机器 1366×768 可以在「更多 ▾」里切成紧凑。
        # 标题栏和任务栏的图标（统一用 gwtyper/assets 那份）
        try:
            from .core import appicon

            self._has_icon = appicon.apply_to(self, APP_NAME)
        except Exception:  # noqa: BLE001
            self._has_icon = False
        self._prefs = sysenv.load_prefs()
        self.ui_density = self._prefs.get("ui_density", "normal")
        self._density_var = tk.StringVar(value=self.ui_density)
        self.fonts = theme.apply(self, density=self.ui_density)

        self.files: list[str] = []
        self.busy = False
        self.msgq: "queue.Queue[tuple]" = queue.Queue()
        self.doc: Doc | None = None
        self.cur: str | None = None
        self.overrides: dict[int, str] = {}
        self.last_out: Path | None = None
        self._batch_owner: dict[str, str] = {}
        self._layout_rows: list[str] = []      # 版面问题行 iid
        self._last_findings: list = []        # 最近一次检查结果，状态栏要用
        self._placeholder: dict[str, ttk.Label] = {}

        # 上次用的格式标准（存的是 yaml 的 stem）；文件没了就退回国标
        _pkey = str(self._prefs.get("profile") or "gb9704")
        try:
            if _pkey not in rulesmod.available():
                _pkey = "gb9704"
        except Exception:  # noqa: BLE001
            _pkey = "gb9704"
        self.profile = tk.StringVar(value=_pkey)
        self.profile_show = tk.StringVar(value="")   # 下拉里显示的名字
        self.status = tk.StringVar(value="把公文拖进左边，或点「添加文件…」")
        self.font_hint = tk.StringVar(value="")
        self.do_layout = tk.BooleanVar(value=False)
        self.do_text = tk.BooleanVar(value=True)
        self.prev_mode = tk.StringVar(value="rule")   # src=当前原件  rule=排版效果
        self.prev_text = tk.StringVar(value="")
        self.sum_text = tk.StringVar(value="")
        self.role_count = tk.StringVar(value="")
        self._prev_tok = 0
        self._prev_job = None

        self._build()
        # 有上次的窗口大小位置就照原样开，没有才按内容需求算
        if not self._restore_geometry():
            self._fit_to_screen()
        self.minsize(980, 620)
        self._setup_dnd()
        self._update_actions()
        # 图标要等窗口真正建好再挂一次：__init__ 早期窗口还没映射，
        # 那时候调 iconbitmap 有的环境不生效（实测 WM_GETICON 取回来是 0）。
        self.after(60, self._reapply_icon)
        self.after(80, self._poll)
        self.after(1500, self._prune_cache)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _sync_body_scrollbars(self) -> None:
        """让内容宽度跟着画布走，并按需显示滚动条。

        宽度上**强制** body 等于画布宽：四栏各自有内容需求宽度（表格列宽之和），
        如果放任它们撑开，body 就会比窗口还宽、出现横向滚动条——四栏一起横滚
        体验很差。强制等于画布宽之后，表格内部的 stretch 列会被压缩到可用宽度，
        这才是用户想看到的（宁可列窄一点，也不要整屏横着滚）。
        高度方向相反：内容多高就是多高，装不下再出纵向滚动条。
        """
        try:
            self.update_idletasks()
            cw = self.body_canvas.winfo_width()
            ch = self.body_canvas.winfo_height()
            # 宽度策略：画布够宽就把 body 铺满；不够就保留内容自然宽度并出横向
            # 滚动条。以前无条件锁死到画布宽，结果四栏需要 1780px 而画布只有
            # 1197px 时，栏被硬压、右侧内容直接被裁掉（预览栏只剩一条边）。
            need_w = self._body_content_width()
            if cw > 10:
                if need_w and need_w > cw:
                    self.body_canvas.itemconfigure(
                        self._body_win, width=need_w)
                else:
                    self.body_canvas.itemconfigure(self._body_win, width=cw - 1)
            if cw > 10 and ch > 10:
                # 高度也要撑满：只锁宽度的话，body 会停在内容需求高度
                # （实测 655px），最大化后下面留一大片空白，四栏挤在上半屏。
                # 但内容比画布高时不能压——那样会把内容切掉，
                # 所以取 max(画布高, 内容需求高)。
                need = self._body_content_height()
                self.body_canvas.itemconfigure(
                    self._body_win, height=max(ch - 1, need))
                self.update_idletasks()
            bbox = self.body_canvas.bbox("all")
            if not bbox:
                return
            need_h = bbox[3] - bbox[1]
            if need_h > ch + 2:
                self._body_vsb.grid()
            else:
                self._body_vsb.grid_remove()
                self.body_canvas.yview_moveto(0)
            # 宽度超出时才显示横向滚动条（保证所有栏都能看到）
            if bbox[2] - bbox[0] > cw + 2:
                self._body_hsb.grid()
            else:
                self._body_hsb.grid_remove()
                self.body_canvas.xview_moveto(0)
        except Exception:  # noqa: BLE001
            pass

    def _body_content_width(self) -> int:
        """四栏内容真正需要的宽度（含栏间距）。"""
        try:
            if not hasattr(self, "_body_frame"):
                return 0
            self._body_frame.update_idletasks()
            total = 0
            for col in self._body_frame.winfo_children():
                total += col.winfo_reqwidth()
                gi = col.grid_info()
                pad = gi.get("padx", 0)
                if isinstance(pad, tuple):
                    total += sum(int(x) for x in pad)
                elif pad:
                    total += int(pad) * 2
            return max(total, 1)
        except Exception:  # noqa: BLE001
            return 0

    def _body_content_height(self) -> int:
        """四栏内容真正需要的高度（不撑满时的高度）。"""
        try:
            if not hasattr(self, "_body_frame"):
                return 0
            self._body_frame.update_idletasks()
            return max(self._body_frame.winfo_reqheight(), 1)
        except Exception:  # noqa: BLE001
            return 0

    def _bind_body_wheel(self, _e=None) -> None:
        def on_wheel(event):
            # Shift+滚轮 = 横滚，跟浏览器一致
            if event.state & 0x1:
                self.body_canvas.xview_scroll(
                    -1 if event.delta > 0 else 1, "units")
            else:
                self.body_canvas.yview_scroll(
                    -1 if event.delta > 0 else 1, "units")
        self._wheel_cb = on_wheel
        self.body_canvas.bind_all("<MouseWheel>", on_wheel)

    def _unbind_body_wheel(self, _e=None) -> None:
        try:
            self.body_canvas.unbind_all("<MouseWheel>")
        except Exception:  # noqa: BLE001
            pass

    def _fit_to_screen(self) -> None:
        """按界面实际需要的尺寸开窗，不超出可用桌面。

        窗口的 reqwidth 会受 Canvas 初始值影响（它只报自己的 width 选项，
        不报里面四栏的内容宽度）。所以这里显式把「四栏真实需求」算进来，
        否则默认窗口会开成 1229px 宽，而四栏要 1688px——右边的预览栏会被裁掉。
        """
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        l, t, w, h = work_area()
        if not w or not h:
            l, t, w, h = 0, 0, sw, sh

        # 内容真实需求：头部/页脚等固定部分 + 四栏所需宽度
        base_w = self.winfo_reqwidth() - self.body_canvas.winfo_reqwidth()
        body_need_w = self._body_content_width() or 1180
        body_need_h = self._body_content_height() or 660
        canvas_chrome = self.body_canvas.winfo_reqwidth() - 1180   # 滚动条等
        need_w = max(self.winfo_reqwidth(), base_w + body_need_w + max(0, canvas_chrome))
        need_h = self.winfo_reqheight()
        if body_need_h > 660:
            need_h += body_need_h - 660

        win_w = min(max(need_w, 900), int(w * 0.96), sw)
        win_h = min(max(need_h, 560), int(h * 0.94), sh)
        x = l + max(0, (w - win_w) // 2)
        y = t + max(0, (h - win_h) // 3)
        self.geometry(f"{win_w}x{win_h}+{x}+{y}")

    def _reapply_icon(self) -> None:
        try:
            from .core import appicon

            self._has_icon = appicon.apply_to(self, APP_NAME)
        except Exception:  # noqa: BLE001
            pass

    def _icon_for(self, win: tk.Misc) -> None:
        """子窗口也带上同一个图标，免得对话框顶着 Tk 的默认羽毛。"""
        try:
            from .core import appicon

            appicon.apply_to(win, APP_NAME)
        except Exception:  # noqa: BLE001
            pass

    def _on_callback_error(self, exc_type, exc, tb) -> None:
        """界面回调里抛出的异常，改写成「写日志 + 弹窗」。

        Tk 默认走 report_callback_exception 只往 stderr 打一行，
        而 exe 是窗口程序、根本没有 stderr——于是用户看到的就是
        「点了按钮毫无反应」。OCR 页签那次就是栽在这上面（_mode_info
        少返回一个值，run_ocr 一解包就 ValueError，被静默吞掉）。
        """
        import logging
        import traceback

        detail = "".join(traceback.format_exception(exc_type, exc, tb))
        try:
            logging.error("界面回调异常\n%s", detail)
        except Exception:  # noqa: BLE001
            pass
        try:
            from tkinter import messagebox

            log = ""
            try:
                from .core import sysenv

                log = str(sysenv.data_dir() / "error.log")
            except Exception:  # noqa: BLE001
                pass
            tip = f"这一步出错了：\n\n{exc}\n\n"
            if log:
                tip += f"详细信息已写进：\n{log}"
            messagebox.showerror(f"{APP_NAME} {APP_VER}", tip)
        except Exception:  # noqa: BLE001
            pass

    def _prune_cache(self) -> None:
        """顺手清理老格式转换缓存，免得越堆越多。"""
        def job():
            try:
                from .core.doc_source import prune

                prune()
            except Exception:  # noqa: BLE001
                pass

        threading.Thread(target=job, daemon=True).start()

    # ---------------------------------------------------------------- 小工具
    def _card(self, parent, column: int, minsize: int = 0, weight: int = 1,
              padx=(0, 12)):
        """一张卡片：白底 + 1px 描边（ttk 没阴影，靠描边和底色差分层）。"""
        outer = tk.Frame(parent, background=theme.BORDER_SUBTLE, bd=0,
                         highlightthickness=0)
        outer.grid(row=0, column=column, sticky="nsew", padx=padx, pady=(0, 0))
        inner = ttk.Frame(outer, style="Card.TFrame", padding=12)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        inner.columnconfigure(0, weight=1)
        if minsize:
            parent.columnconfigure(column, minsize=minsize, weight=weight)
        else:
            parent.columnconfigure(column, weight=weight)
        return inner

    def _placeholder_for(self, host: tk.Widget, key: str, text: str) -> None:
        lbl = ttk.Label(host, text=text, style="Mic.TLabel", justify="center",
                        background=theme.SURFACE)
        lbl.place(relx=0.5, rely=0.42, anchor="center")
        self._placeholder[key] = lbl

    def _toggle_placeholder(self, key: str, show: bool) -> None:
        lbl = self._placeholder.get(key)
        if not lbl:
            return
        if show:
            lbl.place(relx=0.5, rely=0.42, anchor="center")
        else:
            lbl.place_forget()

    # ---------------------------------------------------------------- 布局
    def _on_nav(self, text: str) -> None:
        self._nav_page = text

    def _build(self) -> None:
        head = ttk.Frame(self, padding=(14, 12, 14, 8))
        head.pack(fill="x")
        ttk.Label(head, text=APP_NAME, style="H1.TLabel").pack(side="left")
        ttk.Label(head, text=f"v{APP_VER}", style="Mic.TLabel").pack(
            side="left", padx=(8, 0), pady=(6, 0))
        ttk.Label(head, text="公文排版 · 格式转换 · OCR 识别 · 办公小工具 · 材料上交", style="Status.TLabel").pack(
            side="left", padx=(12, 0), pady=(5, 0))
        ttk.Label(head, text=f"{sysenv.platform_note()} · 转换引擎：{office.engine_hint()}",
                  style="Status.TLabel").pack(side="right", pady=(5, 0))

        body_nav = tk.Frame(self, background=theme.CANVAS)
        body_nav.pack(fill="both", expand=True, padx=14, pady=(0, 0))
        self.nav = SideNav(body_nav, on_select=self._on_nav, width=196)
        self.nav.pack(side="left", fill="y", padx=(0, 12))
        self.nb = NavShim(body_nav, nav=self.nav)
        self.nb.pack(side="left", fill="both", expand=True)

        self.tab_fmt = ttk.Frame(self.nb, padding=(0, 8, 0, 0),
                                 style="Canvas.TFrame")
        self.nb.add(self.tab_fmt, text="公文排版", group="公文处理")

        self.tpl_pane = TemplatePane(self.nb)
        self.nb.add(self.tpl_pane, text="公文模板库", group="公文处理")

        # 四栏内容在 1366×768 上装不下（实测需要 1948×1211），
        # 以前窗口被压小之后多出来的部分直接消失、也没有滚动条，
        # 文件列表只剩 69px、检查结果的「应为」列看不见。
        # 这里套一层可滚动的 Canvas：屏幕够大时完全看不出区别，
        # 屏幕小时至少能滚到想看的地方。
        scroll_host = ttk.Frame(self.tab_fmt, style="Canvas.TFrame")
        scroll_host.pack(fill="both", expand=True)
        scroll_host.rowconfigure(0, weight=1)
        scroll_host.columnconfigure(0, weight=1)

        # Canvas 会向父级报告自己 width/height 选项的值（默认 378x265），
        # 不会报告里面 window 的大小。不给它设初始尺寸，_fit_to_screen 就会
        # 以为"内容只要 566x396"，按最小尺寸开一个很小的窗口——
        # 正是「最大化前窗口很小、内容显示不全」的原因。
        self.body_canvas = tk.Canvas(scroll_host, background=theme.CANVAS,
                                     highlightthickness=0, bd=0,
                                     width=1180, height=660)
        self.body_canvas.grid(row=0, column=0, sticky="nsew")
        vsb = ttk.Scrollbar(scroll_host, orient="vertical",
                            command=self.body_canvas.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb = ttk.Scrollbar(scroll_host, orient="horizontal",
                            command=self.body_canvas.xview)
        hsb.grid(row=1, column=0, sticky="ew")
        self.body_canvas.configure(yscrollcommand=vsb.set,
                                   xscrollcommand=hsb.set)
        self._body_vsb, self._body_hsb = vsb, hsb

        body = ttk.Frame(self.body_canvas)
        self._body_frame = body
        self._body_win = self.body_canvas.create_window(
            (0, 0), window=body, anchor="nw")
        body.rowconfigure(0, weight=1)

        def _on_body_configure(_e=None):
            # 内容尺寸变了就更新滚动范围；同时让滚动条在不需要时自动隐藏
            self.body_canvas.configure(scrollregion=self.body_canvas.bbox("all"))
            self._sync_body_scrollbars()

        body.bind("<Configure>", _on_body_configure)
        self.body_canvas.bind("<Configure>", lambda _e: self._sync_body_scrollbars())

        # 鼠标在四栏区域里时可滚（Windows 上 <MouseWheel> 只在有焦点的控件上触发，
        # 这里改成进入区域就 bind_all，离开再解绑）
        self.body_canvas.bind("<Enter>", self._bind_body_wheel)
        self.body_canvas.bind("<Leave>", self._unbind_body_wheel)

        self._build_files(body)
        self._build_roles(body)
        self._build_findings(body)
        self._build_preview(body)

        # 四栏内容自然宽度合计约 1650，1366 屏幕放不下时会出横向滚动条。
        # 关键是让 body 的宽度**跟着画布走**（见下面的 _sync_body_width），
        # 这样表格会被压缩到可用宽度内，而不是把 body 撑开。

        self.conv = ConvertPane(self.nb, default_dir=str(sysenv.desktop_dir()),
                                prefs=self._prefs)
        self.nb.add(self.conv, text="格式转换", group="文档处理")

        # OCR 单独一页：它和「格式转换」解决的是不同问题——
        # 那页管格式互转，这页专门把「图上的字」认出来。
        self.ocr_pane = OcrPane(self.nb, default_dir=str(sysenv.desktop_dir()),
                                   prefs=self._prefs)
        self.nb.add(self.ocr_pane, text="OCR 识别", group="文档处理")

        # 办公小工具：批量重命名、批量字数统计
        self.rename_pane = ToolsPane(self.nb, which="rename")
        self.nb.add(self.rename_pane, text="批量重命名", group="办公小工具")
        self.stat_pane = ToolsPane(self.nb, which="stat")
        self.nb.add(self.stat_pane, text="字数统计", group="办公小工具")
        self.scan_pane = ToolsPane(self.nb, which="deskew")
        self.nb.add(self.scan_pane, text="扫描件纠偏", group="办公小工具")
        self.form_pane = ToolsPane(self.nb, which="form")
        self.nb.add(self.form_pane, text="签到表·交接单", group="办公小工具")

        # 材料上交：按村名单核对谁交了、谁没交
        self.img_pane = ImageCompressPane(self.nb)
        self.nb.add(self.img_pane, text="图片压缩", group="办公小工具")
        self.dup_pane = DupPane(self.nb)
        self.nb.add(self.dup_pane, text="文件查重", group="办公小工具")

        self.pdf_pane = PdfToolsPane(self.nb)
        self.nb.add(self.pdf_pane, text="PDF 工具箱", group="PDF 工具")

        self.wm_pane = WatermarkPane(self.nb)
        self.nb.add(self.wm_pane, text="批量加水印", group="批量处理")
        self.mask_pane = MaskPane(self.nb)
        self.nb.add(self.mask_pane, text="敏感信息脱敏", group="批量处理")
        self.sum_pane = MergeTablePane(self.nb)
        self.nb.add(self.sum_pane, text="数据汇总", group="批量处理")
        self.pc_pane = PrintCheckPane(self.nb)
        self.nb.add(self.pc_pane, text="打印体检", group="批量处理")

        self.checklist_pane = ChecklistPane(self.nb)
        self.nb.add(self.checklist_pane, text="材料上交", group="材料管理")
        self.diff_pane = ListDiffPane(self.nb)
        self.nb.add(self.diff_pane, text="名单比对", group="材料管理")

        foot = ttk.Frame(self, padding=(14, 4, 14, 8))
        foot.pack(fill="x")
        self.pb = ttk.Progressbar(foot, mode="determinate")
        self.pb.pack(fill="x", pady=(0, 5))
        ttk.Label(foot, textvariable=self.status, style="Status.TLabel").pack(side="left")
        ttk.Label(foot, textvariable=self.font_hint, style="Err.TLabel").pack(side="right")

    def _build_files(self, body) -> None:
        col = ttk.Frame(body, style="Canvas.TFrame")
        col.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        # 左栏要有明确的比例和下限：之前它没设 columnconfigure，
        # 宽度纯由内容决定（498px），「添加文件…」这类按钮被压得很窄。
        body.columnconfigure(0, minsize=300, weight=0)
        col.columnconfigure(0, weight=1)
        col.rowconfigure(0, weight=1)
        col.rowconfigure(1, weight=0)

        card = ttk.Frame(col, style="Card.TFrame", padding=12)
        card.grid(row=0, column=0, sticky="nsew")
        card.columnconfigure(0, weight=1)
        card.rowconfigure(2, weight=1)

        top = ttk.Frame(card, style="Card.TFrame")
        top.grid(row=0, column=0, sticky="ew")
        ttk.Label(top, text="公文", style="H2.TLabel").pack(side="left")

        btns = ttk.Frame(card, style="Card.TFrame")
        btns.grid(row=1, column=0, sticky="ew", pady=(8, 8))
        ttk.Button(btns, text="添加文件…", style="Secondary.TButton",
                   command=self.add_files).pack(side="left")
        ttk.Button(btns, text="添加文件夹…", style="Ghost.TButton",
                   command=self.add_folder).pack(side="left", padx=(4, 0))

        wrap = ttk.Frame(card, style="Card.TFrame")
        wrap.grid(row=2, column=0, sticky="nsew")
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(wrap, columns=("name", "hint"), show="headings",
                                 selectmode="browse", height=6)
        self.tree.heading("name", text="文件")
        self.tree.heading("hint", text="")
        self.tree.column("name", width=180, anchor="w", stretch=True)
        self.tree.column("hint", width=56, anchor="e", stretch=False)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.bind("<<TreeviewSelect>>", self._on_pick)
        theme.stripe(self.tree)
        self._placeholder_for(wrap, "files", "把公文拖到这里\n或点上面的「添加文件…」")

        ttk.Label(card, text="格式标准", style="H3.TLabel").grid(
            row=3, column=0, sticky="w", pady=(10, 3))
        # 下拉里显示中文名（「国家标准 GB/T 9704 公文格式」），
        # 不直接暴露 gb9704 这种 yaml 文件名——排版确认弹窗里早就是中文了，
        # 主界面没同步，用户看到「gb9704」会不知道是什么。
        self.cb = ttk.Combobox(
            card, textvariable=self.profile_show,
            values=[profile_label(k) for k in rulesmod.available()],
            state="readonly")
        self.cb.grid(row=4, column=0, sticky="ew")
        self.cb.bind("<<ComboboxSelected>>", self._on_profile_pick)
        self.cb.bind("<<ComboboxSelected>>", lambda _e: self._sync_profile_tip(), add="+")

        self.prof_tip = ttk.Label(card, text="", style="Mic.TLabel", wraplength=230,
                                  justify="left")
        self.prof_tip.grid(row=5, column=0, sticky="w", pady=(3, 0))

        chk = ttk.Frame(card, style="Card.TFrame")
        chk.grid(row=6, column=0, sticky="w", pady=(6, 0))
        CheckRow(chk, "校对文字（标点、错词、序号）", self.do_text,
                 command=self.recheck, font=self.fonts["body"]).grid(row=0, column=0, sticky="w")
        # 勾上就立刻重算（以前没有 command，勾了没反应，用户以为已经查过了）
        CheckRow(chk, "检查版面（转 PDF，较慢）", self.do_layout,
                 command=self.recheck,
                 font=self.fonts["body"]).grid(row=1, column=0, sticky="w", pady=(3, 0))

        # 排版结果条：平时不占位
        self.res_bar = ttk.Frame(card, style="Card.TFrame")
        self.res_bar.grid(row=7, column=0, sticky="ew", pady=(8, 0))
        self.res_bar.columnconfigure(0, weight=1)
        self.res_lbl = ttk.Label(self.res_bar, text="", style="Ok.TLabel")
        self.res_lbl.grid(row=0, column=0, sticky="w")
        res_btns = ttk.Frame(self.res_bar, style="Card.TFrame")
        res_btns.grid(row=1, column=0, sticky="w", pady=(2, 0))
        ttk.Button(res_btns, text="打开", style="Ghost.TButton", width=6,
                   command=self._open_last).pack(side="left")
        # 报材料基本都要 PDF，排完就地出一份，不必再切到「格式转换」页
        ttk.Button(res_btns, text="转成 PDF", style="Ghost.TButton", width=9,
                   command=self.export_last_pdf).pack(side="left", padx=(6, 0))
        ttk.Button(res_btns, text="所在文件夹", style="Ghost.TButton", width=10,
                   command=self.open_dir).pack(side="left", padx=(6, 0))
        self.res_bar.grid_remove()

        # 动作区：只留两个主要按钮，其余收进「更多 ▾」。
        # 以前这里纵向堆了 4 排按钮（约 246px 高），把左栏顶到 856px，
        # 1366×768 的笔记本上整页装不下、上下都得滚。
        act = ttk.Frame(col, style="Card.TFrame", padding=12)
        act.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        act.columnconfigure(0, weight=1)
        act.columnconfigure(1, weight=1)
        self.btn_fmt = ttk.Button(act, text="按规则排版…", style="Primary.TButton",
                                  command=self.run_format)
        self.btn_fmt.grid(row=0, column=0, sticky="ew")
        self.btn_check = ttk.Button(act, text="检查（只读）", style="Secondary.TButton",
                                    command=self.run_check)
        self.btn_check.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        row = ttk.Frame(act, style="Card.TFrame")
        row.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        row.columnconfigure(0, weight=1)
        self.btn_all = ttk.Button(row, text="检查全部", style="Ghost.TButton",
                                  command=self.run_batch)
        self.btn_all.grid(row=0, column=0, sticky="ew")
        self.btn_exp = ttk.Button(row, text="导出报告", style="Ghost.TButton",
                                  command=self.export_report)
        self.btn_exp.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        more = ttk.Button(row, text="更多 ▾", style="Ghost.TButton",
                          width=8, command=self._more_menu)
        more.grid(row=0, column=2, sticky="e", padx=(6, 0))

        self._refresh_profile_combo()
        self._sync_profile_tip()

    def _more_menu(self) -> None:
        m = tk.Menu(self, tearoff=0)
        m.add_command(label="一键处理（体检 + 排版 + 导 PDF + 报告）…",
                      command=self.run_pipeline)
        m.add_separator()
        m.add_command(label="按规则排版全部文件…", command=self.run_batch_format)
        m.add_command(label="把当前文件转成 PDF…", command=lambda: self.export_pdf(self.cur))
        m.add_command(label="导出检查结果（CSV）…", command=self.export_findings_csv)
        m.add_command(label="对比排版前后差异…", command=self.show_diff)
        m.add_separator()
        m.add_command(label="备份设置…", command=self.backup_prefs)
        m.add_command(label="恢复设置…", command=self.restore_prefs)
        m.add_separator()
        m.add_command(label="格式标准设置…", command=self.edit_profile)
        m.add_command(label="从参照件学一套标准…", command=self.learn_from_ref)
        m.add_separator()
        # 布局密度：1920×1080 用常规，1366×768 换紧凑能一屏放下
        dens = tk.Menu(m, tearoff=0)
        self._density_var.set(self.ui_density)
        dens.add_radiobutton(label="常规（适合 1920×1080 及以上）", value="normal",
                             variable=self._density_var,
                             command=lambda: self.set_density("normal"))
        dens.add_radiobutton(label="紧凑（适合 1366×768 老机器）", value="compact",
                             variable=self._density_var,
                             command=lambda: self.set_density("compact"))
        m.add_cascade(label=f"界面密度：{self._density_cn()}", menu=dens)
        m.add_separator()
        m.add_command(label="打开公文所在文件夹", command=self.open_dir)
        m.add_command(label="打开我的规则目录", command=self.open_rules_dir)
        m.add_command(label="清理后台 WPS 残留进程", command=self.clean_zombies)
        m.add_separator()
        m.add_command(label="清空列表", command=self.clear_files)
        try:
            x = self.winfo_pointerx()
            y = self.winfo_pointery()
            m.tk_popup(x, y)
        finally:
            m.grab_release()

    def _density_cn(self) -> str:
        return "紧凑" if self.ui_density == "compact" else "常规"

    def set_density(self, mode: str) -> None:
        """切换界面密度。表格行高和 tree 高度变更后要重建窗口才生效。"""
        if mode == self.ui_density:
            return
        self.ui_density = mode
        self._prefs["ui_density"] = mode
        sysenv.save_prefs(self._prefs)
        messagebox.showinfo(
            APP_NAME,
            f"已切换到「{self._density_cn()}」布局。\n\n"
            "重新打开程序后生效（表格行高会变，需要重建窗口）。")

    def _profile_key(self) -> str:
        """当前生效的规则集 key（yaml 的 stem）。

        注意有两个变量：
          - self.profile      存 key（gb9704），给 rules.load 用
          - self.profile_show 存中文名，只做下拉显示
        **不能只读 self.profile**：用户改的是下拉（profile_show），
        读 profile 永远拿到旧值，切换就失效了（这个 bug 真出现过）。
        所以这里优先看 profile_show，解析不出再退回 profile。
        """
        for raw in (self.profile_show.get().strip(), self.profile.get().strip()):
            if not raw:
                continue
            # 已经是 key
            try:
                if raw in rulesmod.available():
                    return raw
            except Exception:  # noqa: BLE001
                pass
            # 是中文名
            for k in PROFILE_CN:
                if raw == profile_label(k):
                    return k
            for k in PROFILE_CN:
                if raw == PROFILE_CN[k]:
                    return k
            # 是路径
            if raw.lower().endswith((".yaml", ".yml")):
                return raw
        return Path(self.profile.get() or "gb9704").stem

    def _on_profile_pick(self, _e=None) -> None:
        """用户在下拉里选了格式标准：换回 key 再重算。"""
        key = self._profile_key()
        self.profile.set(key)
        self.profile_show.set(profile_label(key))
        self._sync_profile_tip()
        self.recheck()

    def _refresh_profile_combo(self) -> None:
        """候选列表和当前显示都跟 rules 目录保持一致。

        这里的显示值要按 **self.profile（内部 key）** 来算，不能再用
        _profile_key()——那个函数为了支持「用户改了下拉」而优先读
        profile_show，代码里直接设置 profile 时会把刚设的 key 又冲掉，
        形成循环依赖。
        """
        try:
            keys = rulesmod.available()
            self.cb.configure(values=[profile_label(k) for k in keys])
            key = Path(self.profile.get() or "gb9704").stem
            if key not in keys:
                key = keys[0] if keys else "gb9704"
                self.profile.set(key)
            self.profile_show.set(profile_label(key))
        except Exception:  # noqa: BLE001
            pass

    def _sync_profile_tip(self) -> None:
        key = self._profile_key()
        tip = PROFILE_TIP.get(key, "这套标准由你自己生成，随时可以在「格式标准设置」里调")
        self.prof_tip.configure(text=tip)

    def _build_roles(self, body) -> None:
        card = self._card(body, 1, minsize=250, weight=4)
        card.rowconfigure(1, weight=1)

        top = ttk.Frame(card, style="Card.TFrame")
        top.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(top, text="段落角色", style="H2.TLabel").pack(side="left")
        ttk.Label(top, textvariable=self.role_count, style="Mic.TLabel").pack(
            side="left", padx=(8, 0), pady=(2, 0))

        wrap = ttk.Frame(card, style="Card.TFrame")
        wrap.grid(row=1, column=0, sticky="nsew")
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        cols = ("no", "role", "conf", "text")
        self.roles = ttk.Treeview(wrap, columns=cols, show="headings",
                                  selectmode="browse")
        for c, t, w, a in (("no", "#", 30, "center"), ("role", "角色", 88, "w"),
                           ("conf", "确定度", 84, "center"),
                           ("text", "内容", 110, "w")):
            self.roles.heading(c, text=t)
            self.roles.column(c, width=w, anchor=a, stretch=(c == "text"))
        self.roles.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.roles.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.roles.configure(yscrollcommand=sb.set)
        self.roles.tag_configure("low", foreground=theme.WARN_TX)
        self.roles.tag_configure("manual", foreground=theme.PRIMARY)
        theme.stripe(self.roles)
        self.roles.bind("<Double-1>", self._edit_role)
        self.roles.bind("<Button-3>", self._role_menu)
        self._placeholder_for(wrap, "roles", "导入文件后，这里列出每一段\n被认成了什么")

        ttk.Label(card, text="双击改角色 · 右键可撤销 · 确定度低的行排版会跳过",
                  style="Mic.TLabel", wraplength=240).grid(
            row=2, column=0, sticky="w", pady=(6, 0))

    def _build_findings(self, body) -> None:
        card = self._card(body, 2, minsize=290, weight=4)
        card.rowconfigure(1, weight=1)

        top = ttk.Frame(card, style="Card.TFrame")
        top.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(top, text="检查结果", style="H2.TLabel").pack(side="left")
        ttk.Label(top, textvariable=self.sum_text, style="Mic.TLabel").pack(
            side="left", padx=(8, 0), pady=(2, 0))

        wrap = ttk.Frame(card, style="Card.TFrame")
        wrap.grid(row=1, column=0, sticky="nsew")
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        # 列压到 400 以内：1366 屏幕上四栏加起来才装得下。
        # 「现在→应为」合成一列，因为用户真正要看的就是「该改成什么」，
        # 分成两列在窄屏上会被切掉后半截（"30磅"看得见、"32磅固定值"看不见）。
        cols = ("where", "field", "fix")
        self.finds = ttk.Treeview(wrap, columns=cols, show="tree headings",
                                  selectmode="browse")
        self.finds.column("#0", width=22, stretch=False, anchor="center")
        self.finds.heading("#0", text="")
        for c, t, w, a in (("where", "位置", 66, "w"), ("field", "检查项", 82, "w"),
                           ("fix", "现在 → 应为", 160, "w")):
            self.finds.heading(c, text=t)
            self.finds.column(c, width=w, anchor=a, stretch=(c == "fix"))
        self.finds.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.finds.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.finds.configure(yscrollcommand=sb.set)
        self.finds.tag_configure("error", foreground=theme.ERR_TX)
        self.finds.tag_configure("warn", foreground=theme.WARN_TX)
        self.finds.tag_configure("info", foreground=theme.TEXT2)
        self.finds.tag_configure("file", background=theme.SUNKEN, foreground=theme.TEXT)
        theme.stripe(self.finds)
        self.finds.bind("<Double-1>", self._jump_to_para)
        self._placeholder_for(wrap, "finds", "还没有检查结果\n选一份公文点「检查」")

        ttk.Label(card, text="双击某条可跳到对应段落", style="Mic.TLabel",
                  wraplength=260).grid(row=2, column=0, sticky="w", pady=(6, 0))

    def _build_preview(self, body) -> None:
        from .ui.preview_pane import PreviewPane

        card = self._card(body, 3, minsize=260, weight=4, padx=(0, 0))
        card.rowconfigure(1, weight=1)

        bar = ttk.Frame(card, style="Card.TFrame")
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(bar, text="版面", style="H2.TLabel").pack(side="left", padx=(0, 10))
        ttk.Radiobutton(bar, text="排版效果", value="rule", variable=self.prev_mode,
                        style="Card.TRadiobutton", command=self.refresh_preview).pack(side="left")
        ttk.Radiobutton(bar, text="当前原件", value="src", variable=self.prev_mode,
                        style="Card.TRadiobutton", command=self.refresh_preview).pack(
            side="left", padx=(8, 0))
        ttk.Button(bar, text="重绘", style="Ghost.TButton",
                   command=lambda: self.refresh_preview(force=True)).pack(side="right")

        holder = ttk.Frame(card, style="Card.TFrame")
        holder.grid(row=1, column=0, sticky="nsew")
        holder.columnconfigure(0, weight=1)
        holder.rowconfigure(0, weight=1)
        self.preview = PreviewPane(holder, width_hint=200)
        self.preview.grid(row=0, column=0, sticky="nsew")
        self.preview.hide_head()
        # 预览栏的占位文字由 PreviewPane 自己画在画布上；
        # 这里再放一个 Label 会和它叠在一起（两个"导入文件后"重影）。

        foot = ttk.Frame(card, style="Card.TFrame")
        foot.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        ttk.Label(foot, textvariable=self.prev_text, style="Mic.TLabel").pack(side="left")

    # ---------------------------------------------------------------- 拖拽
    def _setup_dnd(self) -> None:
        if not _HAS_DND:
            self.status.set("拖拽组件没装上（tkinterdnd2），请用「添加文件…」")
            return
        targets = [self.tree, self, self._placeholder.get("files")]
        if not dnd.bind_targets(targets, self._on_drop):
            self.status.set("拖拽注册失败，请用「添加文件…」")

    def _on_drop(self, event) -> None:  # pragma: no cover
        from .core.doc_source import is_supported

        paths = dnd.paths_from_event(self, event, accept=is_supported)
        if paths:
            self._add_paths([str(p) for p in paths])

    # ---------------------------------------------------------------- 状态
    def _update_actions(self) -> None:
        has = self.doc is not None and not self.busy
        state = "normal" if has else "disabled"
        for b in (self.btn_check, self.btn_fmt):
            b.configure(state=state)
        self.btn_all.configure(state="normal" if self.files and not self.busy else "disabled")
        self.btn_exp.configure(state="normal" if self.files and not self.busy else "disabled")
        self._toggle_placeholder("files", not self.files)
        self._toggle_placeholder("roles", self.doc is None)
        self._toggle_placeholder("finds", not self.finds.get_children())
        self._toggle_placeholder("preview", not self.preview._pages)

    # ---------------------------------------------------------------- 文件
    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="选择公文",
            filetypes=[("Word 文档", "*.docx *.doc *.wps *.rtf"), ("全部", "*.*")])
        self._add_paths(list(paths))

    def add_folder(self) -> None:
        d = filedialog.askdirectory(title="选择文件夹")
        if d:
            self._add_paths([d])

    def _add_paths(self, paths: list[str]) -> None:
        from .core.doc_source import is_supported

        added = 0
        rejected: list[str] = []      # 格式不收
        missing: list[str] = []       # 文件不存在
        for p in paths:
            p = Path(p)
            if p.is_dir():
                found = sorted(x for x in p.rglob("*")
                               if x.is_file() and is_supported(x))
                if not found:
                    rejected.append(p.name)
            else:
                found = [p]
            for f in found:
                s = str(f)
                if f.name.startswith("~$"):
                    continue
                if not is_supported(f):
                    rejected.append(f.name)
                    continue
                if not f.exists():
                    missing.append(f.name)
                    continue
                if s not in self.files:
                    self.files.append(s)
                    self.tree.insert("", "end", iid=s, values=(f.name, ""))
                    added += 1
        self._update_actions()
        if self.files and not self.tree.selection():
            self.tree.selection_set(self.files[0])
        if added:
            self.status.set(f"已添加 {added} 份，共 {len(self.files)} 份")
        # 拖错格式以前是完全静默的——拖一次没反应、拖两次还没有，
        # 用户就会认定「这软件坏了」。必须说清楚为什么不收。
        if not added and (rejected or missing):
            if missing:
                self.status.set(f"找不到文件：{missing[0]}"
                                + ("…" if len(missing) > 1 else ""))
            else:
                self.status.set(
                    f"这个格式不支持：{rejected[0]}"
                    f"（只收 Word/WPS/RTF 文档）"
                    + ("…" if len(rejected) > 1 else ""))

    def clear_files(self) -> None:
        self.files.clear()
        self.tree.delete(*self.tree.get_children())
        self.roles.delete(*self.roles.get_children())
        self.finds.delete(*self.finds.get_children())
        self.doc = None
        self.cur = None
        self.overrides.clear()
        self._batch_owner.clear()
        self.preview.clear("")
        self.prev_text.set("")
        self.sum_text.set("")
        self.role_count.set("")
        self.res_bar.grid_remove()
        self.status.set("把公文拖进左边，或点「添加文件…」")
        self._update_actions()

    def _on_pick(self, _e=None) -> None:
        sel = self.tree.selection()
        if sel and sel[0] != self.cur:
            self.load_doc(sel[0])

    def load_doc(self, path: str) -> None:
        from .core.doc_source import DocSourceError

        if self.busy:
            return
        try:
            self.doc = parse(path)
        except DocSourceError as e:
            messagebox.showerror(APP_NAME, str(e))
            return
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(APP_NAME, f"打不开这个文件：\n{e}")
            return
        # 换了文件才清手工指定；同一个文件重新载入要保留
        if self.cur and Path(self.cur) != Path(path):
            self.overrides.clear()
        self.cur = path
        self.recheck()
        if self._prev_job:
            self.after_cancel(self._prev_job)
            self._prev_job = None
        self.preview.clear("正在生成预览，请稍候…")
        self.after(30, lambda: self.refresh_preview(force=True))
        if Path(path).suffix.lower() != ".docx":
            self.status.set(f"{Path(path).name} 已转换后读取（原件不动）")
        self._update_actions()

    def _profile_rules(self) -> dict:
        return rulesmod.load(self._profile_key())

    def recheck(self) -> None:
        if not self.doc:
            return
        rules = self._profile_rules()
        auto, review = classify(self.doc.paras, 0.45, self.overrides)
        for p in auto + review:
            self.doc.paras[p.index].role = p.role
            self.doc.paras[p.index].confidence = p.confidence
        findings = check(self.doc, rules, text=self.do_text.get())
        findings.extend(check_features(self.doc, rules))
        self._last_findings = findings      # run_check 要用它在状态栏报数量
        self._fill_roles(review)
        self._fill_findings(findings)
        self._check_fonts(rules)
        self._update_actions()
        self.schedule_preview()

    # ---------------------------------------------------------------- 预览
    def schedule_preview(self, delay: int = 900) -> None:
        """改规则/改角色后稍等一下再重渲染，免得连点按钮时反复转 PDF。"""
        if self._prev_job:
            try:
                self.after_cancel(self._prev_job)
            except Exception:  # noqa: BLE001
                pass
        self._prev_job = self.after(delay, self.refresh_preview)

    def refresh_preview(self, force: bool = False) -> None:
        if self.busy or not self.doc:
            return
        self._prev_job = None
        self._prev_tok += 1
        tok = self._prev_tok
        mode = self.prev_mode.get()
        rules = self._profile_rules() if mode == "rule" else None
        overrides = dict(self.overrides) if mode == "rule" else None
        path = self.cur
        width = max(300, self.preview.canvas.winfo_width() - 28)

        self.prev_text.set("渲染中…")
        if not self.preview._pages:
            self.preview.clear("正在生成预览，请稍候…")
            self._toggle_placeholder("preview", False)

        def job():
            from .core import preview as prevmod

            try:
                src = prevmod.make_docx(path, rules, overrides)
                pdf = prevmod.pdf_of(src, force=force or mode == "rule")
                pages, n = prevmod.render(pdf, max_width=width)
                self.msgq.put(("prev", tok, pages, n, mode, None))
            except Exception as exc:  # noqa: BLE001
                self.msgq.put(("prev", tok, [], 0, mode, str(exc)))

        threading.Thread(target=job, daemon=True).start()

    def _check_fonts(self, rules: dict) -> None:
        from .core.fonts import check_fonts

        want = {s.get("font") for s in rules.get("roles", {}).values() if s.get("font")}
        info = check_fonts(want)
        if not info["checked"]:
            self.font_hint.set("")
        elif info["missing"]:
            self.font_hint.set("缺字体：" + "、".join(info["missing"]) + "（会用别的字替代，版面会走形）")
        else:
            self.font_hint.set(f"字体齐备（{info['total']} 个可用）")

    # ---------------------------------------------------------------- 填表
    def _fill_roles(self, review) -> None:
        self.roles.delete(*self.roles.get_children())
        low = {p.index for p in review}
        self.role_count.set(f"{len(self.doc.paras)} 段 · 需要你确认 {len(low)} 段")
        for p in self.doc.paras:
            if not p.text.strip():
                continue
            tags = []
            if p.index in self.overrides:
                tags.append("manual")
            elif p.index in low:
                tags.append("low")
            self.roles.insert("", "end", iid=str(p.index), tags=tags,
                              values=(p.index + 1, ROLE_CN[p.role.value],
                                      conf_text(p.confidence),
                                      p.text[:60].replace("\n", " ")))
        self._toggle_placeholder("roles", False)

    def _fill_batch(self, out, failed) -> None:
        """批量结果：右侧按文件分组列问题，左侧文件行显示计数。"""
        self.finds.delete(*self.finds.get_children())
        self._batch_owner.clear()
        tot_e = tot_w = 0
        for path, doc, findings, review in out:
            e = sum(1 for f in findings if f.level == "error")
            w = sum(1 for f in findings if f.level == "warn")
            tot_e += e
            tot_w += w
            if self.tree.exists(path):
                hint = f"必改 {e}" if e else (f"建议 {w}" if w else "通过")
                self.tree.item(path, values=(Path(path).name, hint))
            head = self.finds.insert("", "end", open=True,
                                     values=(Path(path).name, f"{len(findings)} 项",
                                             "",
                                             f"必改 {e}／建议 {w}／待确认 {len(review)}"))
            self.finds.item(head, tags=("file",))
            self._batch_owner[head] = path
            for f in sorted(findings, key=lambda x: (x.level != "error", x.index or 0)):
                loc = f"第{f.index + 1}段" if f.index is not None else "整篇"
                iid = self.finds.insert(head, "end", tags=(f.level,),
                                        values=(loc, f.field, _fix_text(f)))
                self._batch_owner[iid] = path
            for r in review:
                iid = self.finds.insert(head, "end", tags=("info",),
                                        values=(f"第{r.index + 1}段", "需要确认",
                                                r.text[:40], ROLE_CN[r.role.value]))
                self._batch_owner[iid] = path

        for path, err in failed:
            self.finds.insert("", "end", tags=("error",),
                              values=(Path(path).name, "读取失败", err, ""))
        self.sum_text.set(f"{len(out)} 份｜必改 {tot_e}／建议 {tot_w}")
        self.status.set(f"检查完成：{len(out)} 份，必改 {tot_e}／建议 {tot_w}"
                        + (f"，{len(failed)} 份读取失败" if failed else ""))
        self._toggle_placeholder("finds", not self.finds.get_children())

    def _fill_findings(self, findings) -> None:
        self.finds.delete(*self.finds.get_children())
        self._batch_owner.clear()
        self._layout_rows.clear()
        s = summary(findings)
        from .core import report as _rm
        _pts = _rm.score(findings)[0]
        self.sum_text.set(f"必改 {s['error']}／建议 {s['warn']}"
                          f"　健康分 {_pts}（{_rm.score_cn(_pts)}）")
        order = {"error": 0, "warn": 1, "info": 2}
        for f in sorted(findings, key=lambda x: (order.get(x.level, 3),
                                                 x.index if x.index is not None else -1)):
            loc = f"第{f.index + 1}段" if f.index is not None else "整篇"
            mark = "●" if f.level == "error" else ("▲" if f.level == "warn" else "·")
            self.finds.insert("", "end", text=mark, tags=(f.level,),
                              values=(loc, f.field, _fix_text(f)))
        if findings:
            self._toggle_placeholder("finds", False)
            return
        # 零问题是好事，但以前这里会把占位文字换回「还没有检查结果」，
        # 用户刚点过检查、界面却说没检查，分不清是「通过」还是「没跑」。
        # 插一行明确的通过标记，让人一眼看到结果。
        self.finds.insert("", "end", text="✓", tags=("ok",),
                          values=("整篇", "未发现问题", "符合当前格式标准"))
        self.finds.tag_configure("ok", foreground=theme.OK_TX)
        self._toggle_placeholder("finds", False)

    # ---------------------------------------------------------------- 交互
    def _edit_role(self, event) -> None:
        iid = self.roles.identify_row(event.y)
        if not iid or not self.doc:
            return
        self._role_dialog(int(iid))

    def _role_menu(self, event) -> None:
        iid = self.roles.identify_row(event.y)
        if not iid:
            return
        idx = int(iid)
        self.roles.selection_set(iid)
        m = tk.Menu(self, tearoff=0)
        m.add_command(label="改角色…", command=lambda: self._role_dialog(idx))
        if idx in self.overrides:
            m.add_command(label="撤销这个指定",
                          command=lambda: self._undo_role(idx))
        m.add_separator()
        m.add_command(label="撤销全部指定", command=self._undo_all_roles)
        try:
            m.tk_popup(event.x_root, event.y_root)
        finally:
            m.grab_release()

    def _undo_role(self, idx: int) -> None:
        self.overrides.pop(idx, None)
        self.recheck()

    def _undo_all_roles(self) -> None:
        if not self.overrides:
            return
        n = len(self.overrides)
        self.overrides.clear()
        self.recheck()
        self.status.set(f"已撤销 {n} 段的手工指定")

    def _role_dialog(self, idx: int) -> None:
        para = self.doc.paras[idx]
        win = tk.Toplevel(self)
        win.title(f"第 {idx + 1} 段")
        self._icon_for(win)
        win.transient(self)
        win.resizable(False, False)
        win.configure(background=theme.SURFACE)
        pad = ttk.Frame(win, padding=16, style="Card.TFrame")
        pad.pack(fill="both", expand=True)
        ttk.Label(pad, text=para.text[:90] + ("…" if len(para.text) > 90 else ""),
                  wraplength=400, style="Body.TLabel").pack(anchor="w", pady=(0, 12))
        var = tk.StringVar(value=ROLE_CN.get(para.role.value, para.role.value))
        cb = ttk.Combobox(pad, textvariable=var, state="readonly", width=22,
                          values=[ROLE_CN[r] for r in ROLE_CHOICES])
        cb.pack(anchor="w")
        # 判断依据只保留人话，去掉「(+2.0)」这种打分公式——
        # 基层文书看不懂，只会觉得界面在报错
        ev = [e.split("(")[0].strip() for e in (para.evidence or [])][:3]
        ttk.Label(pad, text="依据：" + ("、".join(ev) or "看不清，请手工选"),
                  style="Mic.TLabel", wraplength=400).pack(anchor="w", pady=(8, 0))

        def ok(_e=None):
            # 下拉里是中文，存回 overrides 要转成英文 key
            picked = var.get()
            key = next((k for k, v in ROLE_CN.items() if v == picked), picked)
            self.overrides[idx] = key
            win.destroy()
            self.recheck()

        row = ttk.Frame(pad, style="Card.TFrame")
        row.pack(fill="x", pady=(14, 0))
        ttk.Button(row, text="确定", style="Primary.TButton", command=ok).pack(side="right")
        ttk.Button(row, text="取消", style="Ghost.TButton",
                   command=win.destroy).pack(side="right", padx=(0, 8))
        cb.focus_set()
        _center_dialog(win)
        win.grab_set()

    def _jump_to_para(self, event) -> None:
        iid = self.finds.identify_row(event.y)
        if not iid:
            return
        vals = self.finds.item(iid, "values")
        if len(vals) < 1 or not str(vals[0]).startswith("第"):
            return
        try:
            n = int(str(vals[0]).lstrip("第").rstrip("段"))
        except ValueError:
            return
        owner = self._batch_owner.get(iid)
        if owner and owner != self.cur:
            # 批量结果里点的是别的文件：先切过去并载入，再跳段
            if not self.tree.exists(owner):
                return
            self.tree.selection_set(owner)
            self.tree.see(owner)
            self.load_doc(owner)
            self.after(260, lambda: self._goto_para(n))
            return
        self._goto_para(n)

    def _goto_para(self, n: int) -> None:
        """界面里跳到第 n 段，并在 Word / WPS 里定位。"""
        target = str(n - 1)
        if self.roles.exists(target):
            self.roles.selection_set(target)
            self.roles.see(target)
        # 同时在 Word / WPS 里定位（打不开就只留在界面跳转）
        self._jump_in_word(n)

    def _jump_in_word(self, para_no: int, quiet: bool = True) -> None:
        """在 Word/WPS 里打开当前文档并选中第 para_no 段（1 起）。"""
        from .core import wordjump

        if not self.doc:
            return
        src = self.doc.source or self.doc.path
        if Path(src).suffix.lower() not in (".docx", ".doc"):
            if not quiet:
                messagebox.showinfo(APP_NAME, "只有 Word 文档能在 Word 里定位。")
            return
        if not wordjump.available():
            if not quiet:
                messagebox.showinfo(APP_NAME, "这台电脑上没找到 Word / WPS。")
            return

        def job():
            try:
                wordjump.jump(src, para_no - 1)
            except Exception as exc:  # noqa: BLE001
                if not quiet:
                    self.msgq.put(("error", f"在 Word 里定位失败：{exc}"))

        threading.Thread(target=job, daemon=True).start()

    def open_dir(self) -> None:
        """打开文件夹：优先选中的公文，其次排版结果，再次列表第一份。"""
        path = ""
        sel = self.tree.selection()
        if sel:
            path = sel[0]
        elif self.last_out:
            path = str(self.last_out)
        elif self.files:
            path = self.files[0]
        elif self.cur:
            path = self.cur

        if not path:
            if messagebox.askyesno(APP_NAME, "还没有公文。要现在挑一个文件夹看看吗？"):
                self.add_folder()
            return

        folder = Path(path).parent
        if not folder.exists():
            messagebox.showwarning(APP_NAME, f"这个文件夹不在了：\n{folder}")
            return
        try:
            open_folder(str(folder))
        except Exception as exc:  # noqa: BLE001
            messagebox.showwarning(APP_NAME, f"打不开文件夹：\n{folder}\n\n{exc}")

    def open_rules_dir(self) -> None:
        d = rulesmod.USER_DIR
        d.mkdir(parents=True, exist_ok=True)
        open_folder(str(d))

    def load_profile(self) -> None:
        p = filedialog.askopenfilename(title="选择格式标准",
                                       filetypes=[("规则集", "*.yaml"), ("全部", "*.*")])
        if not p:
            return
        self.profile.set(p)
        self._refresh_profile_combo()
        self._sync_profile_tip()
        self.recheck()

    # ---------------------------------------------------------------- 执行
    def _busy(self, on: bool, text: str = "") -> None:
        self.busy = on
        self.pb.configure(mode="indeterminate" if on else "determinate")
        if on:
            self.pb.start(12)
            self.status.set(text or "处理中…")
        else:
            self.pb.stop()
            self.pb["value"] = 0
        self._update_actions()

    def run_check(self) -> None:
        if not self.doc or self.busy:
            return
        if not self.do_layout.get():
            self.recheck()
            s = summary(self._last_findings) if getattr(self, "_last_findings", None) else None
            tail = f"：必改 {s['error']}／建议 {s['warn']}" if s else ""
            from .core import report as _rm
            pts = _rm.score(self._last_findings or [])[0]
            self.status.set(f"{Path(self.cur).name} 检查完成{tail}"
                            f"　健康分 {pts}（{_rm.score_cn(pts)}）"
                            f"　（勾上「检查版面」还能查越界和孤行）")
            return
        self._busy(True, "正在转 PDF 检查版面…")
        path, page, n_tables = self.doc.path, dict(
            top_cm=self.doc.page.top_cm, bottom_cm=self.doc.page.bottom_cm,
            left_cm=self.doc.page.left_cm, right_cm=self.doc.page.right_cm), self.doc.n_tables

        def job():
            from .core import layout_check
            try:
                rep = layout_check.check(path, page, n_tables)
                self.msgq.put(("layout", rep))
            except Exception as e:  # noqa: BLE001
                self.msgq.put(("error", f"版面检查失败：{e}"))

        threading.Thread(target=job, daemon=True).start()

    def _collect(self, path: str, rules: dict | None = None, text: bool | None = None,
                 overrides: dict | None = None):
        """解析 + 识别 + 体检，返回 (doc, findings, review)。不碰界面。

        规则与开关必须在主线程先取好传进来：tkinter 变量不能在后台线程读
        （会报 ``main thread is not in main loop``）。
        """
        if rules is None:
            rules = self._profile_rules()
        if text is None:
            text = self.do_text.get()
        if overrides is None:
            overrides = self.overrides if path == self.cur else None
        doc = parse(path)
        _auto, review = classify(doc.paras, 0.45, overrides)
        findings = check(doc, rules, text=text)
        findings.extend(check_features(doc, rules))
        return doc, findings, review

    def run_batch(self) -> None:
        """把左侧所有文件跑一遍，结果替换右侧清单（多线程并行）。"""
        if not self.files or self.busy:
            return
        files = list(self.files)
        self._busy(True, f"正在检查 {len(files)} 份…")
        rules = self._profile_rules()
        text = self.do_text.get()

        def job():
            out, failed = [], []
            lock = threading.Lock()
            done = [0]

            def one(f):
                try:
                    doc, findings, review = self._collect(f, rules, text)
                    with lock:
                        out.append((f, doc, findings, review))
                except Exception as e:  # noqa: BLE001
                    with lock:
                        failed.append((f, str(e)))
                with lock:
                    done[0] += 1
                    self.msgq.put(("batch_fmt_progress", done[0], len(files)))

            _parallel(files, one)
            out.sort(key=lambda t: files.index(t[0]))
            self.msgq.put(("batch", out, failed))

        threading.Thread(target=job, daemon=True).start()

    def export_report(self) -> None:
        from .core import report as reportmod

        if not self.files:
            messagebox.showinfo(APP_NAME, "先添加文件")
            return
        types = [("Word 文档", "*.docx"), ("PDF", "*.pdf"), ("纯文本", "*.txt")]
        if len(self.files) > 1:
            dest = filedialog.asksaveasfilename(
                title="保存汇总报告", defaultextension=".docx",
                initialdir=str(sysenv.desktop_dir()),
                initialfile="公文体检汇总.docx", filetypes=types)
        else:
            dest = filedialog.asksaveasfilename(
                title="保存体检报告", defaultextension=".docx",
                initialdir=str(Path(self.files[0]).parent),
                initialfile=f"{Path(self.files[0]).stem}-体检报告.docx",
                filetypes=types)
        if not dest:
            return
        self._busy(True, "正在生成报告…")
        files = list(self.files)
        rules = self._profile_rules()
        text_on = self.do_text.get()

        def job():
            entries, failed, lock = [], [], threading.Lock()

            def one(f):
                try:
                    doc, findings, review = self._collect(f, rules, text_on)
                    with lock:
                        entries.append((f, doc, findings, review))
                except Exception as e:  # noqa: BLE001
                    with lock:
                        failed.append(f"{Path(f).name}: {e}")

            _parallel(files, one)
            entries.sort(key=lambda t: files.index(t[0]))
            try:
                pname = rules.get("name", "")
                if len(entries) > 1:
                    reportmod.save(dest, None, None, None, pname,
                                   entries=[(d, fs, rv) for _f, d, fs, rv in entries])
                else:
                    _f, d, fs, rv = entries[0]
                    reportmod.save(dest, d, fs, rv, pname)
                msg = f"报告已保存：{Path(dest).name}"
                if failed:
                    msg += f"（{len(failed)} 份读取失败）"
                self.msgq.put(("done", msg))
            except Exception as e:  # noqa: BLE001
                self.msgq.put(("error", f"生成报告失败：{e}"))

        threading.Thread(target=job, daemon=True).start()

    def learn_from_ref(self) -> None:
        """选一份已排好的公文，反推一套格式标准存进用户规则目录。"""
        from .core import learn as learnmod

        p = filedialog.askopenfilename(title="选一份已排好的参照公文",
                                       filetypes=[("Word 文档", "*.docx *.doc *.wps"),
                                                  ("全部", "*.*")])
        if not p:
            return
        try:
            prof = learnmod.learn(p)
            out = rulesmod.save_target(Path(p).stem)
            saved = learnmod.save(prof, out)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(APP_NAME, f"生成格式标准失败：\n{e}")
            return
        rulesmod._load_path.cache_clear()
        self.cb.configure(values=rulesmod.available())
        self.profile.set(saved.stem)
        self._refresh_profile_combo()
        self._sync_profile_tip()
        self.recheck()
        messagebox.showinfo(APP_NAME,
                            f"已生成：{saved.name}\n\n"
                            f"识别到 {len(prof['roles'])} 类段落，已设为当前标准。\n"
                            f"建议打开 yaml 核对数值后再用。")

    def edit_profile(self) -> None:
        from .ui.rules_editor import RulesEditor

        try:
            rules = self._profile_rules()
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(APP_NAME, f"格式标准打不开：\n{e}")
            return

        def saved(path: Path) -> None:
            rulesmod._load_path.cache_clear()
            self.cb.configure(values=rulesmod.available())
            self.profile.set(path.stem)
            self._refresh_profile_combo()
            self._sync_profile_tip()
            self.recheck()
            self.status.set(f"格式标准已保存：{path.name}")

        RulesEditor(self, rules, saved)

    # ---------------------------------------------------------------- 排版
    def _target_of(self, src: str) -> Path:
        stem = Path(src).stem
        for tail in ("-排版后", "-排版", "-体检报告"):
            if stem.endswith(tail):
                stem = stem[: -len(tail)]
        return Path(src).with_name(f"{stem}-排版后.docx")

    def _ask_format(self, target: Path, review_n: int) -> dict | None:
        """排版前先说清楚：用哪套标准、写到哪、跳过几段。

        「写到」这一行可以改：默认放原文件旁边，但允许换目录或改名字。
        用户经常要把排好的稿子集中放一处（比如统一交到某个文件夹），
        不让改就得先排完再一个个搬。
        """
        win = tk.Toplevel(self)
        win.title("按规则排版")
        self._icon_for(win)
        win.transient(self)
        win.resizable(False, False)
        win.configure(background=theme.SURFACE)
        pad = ttk.Frame(win, padding=18, style="Card.TFrame")
        pad.pack(fill="both", expand=True)

        ttk.Label(pad, text="按这套标准重排一份副本", style="H2.TLabel").pack(anchor="w")
        info = ttk.Frame(pad, style="Card.TFrame")
        info.pack(fill="x", pady=(12, 0))
        info.columnconfigure(1, weight=1)
        rows = [("格式标准", profile_label(self.profile.get())),
                ("原文件", Path(self.cur).name)]
        for i, (k, v) in enumerate(rows):
            ttk.Label(info, text=k, style="Mic.TLabel").grid(row=i, column=0, sticky="w", pady=3)
            ttk.Label(info, text=v, style="Body.TLabel", wraplength=330).grid(
                row=i, column=1, sticky="w", padx=(12, 0), pady=3)

        # ---- 输出位置：可以改 ----
        ttk.Label(info, text="写到", style="Mic.TLabel").grid(
            row=2, column=0, sticky="nw", pady=3)
        out_row = ttk.Frame(info, style="Card.TFrame")
        out_row.grid(row=2, column=1, sticky="ew", padx=(12, 0), pady=3)
        out_row.columnconfigure(0, weight=1)
        target_var = tk.StringVar(value=str(target))
        ttk.Entry(out_row, textvariable=target_var, width=34).grid(
            row=0, column=0, sticky="ew")
        ttk.Button(out_row, text="浏览…", style="Ghost.TButton",
                   command=lambda: self._pick_out_path(target_var)).grid(
            row=0, column=1, padx=(6, 0))

        # 这几行随输入实时更新（覆盖提示、目录是否有效）
        notes_lbl = ttk.Label(pad, text="", style="Warn.TLabel",
                              wraplength=400, justify="left")
        notes_lbl.pack(anchor="w", pady=(10, 0))

        def refresh_notes(*_a) -> None:
            p = Path(target_var.get().strip() or str(target))
            notes = []
            if review_n:
                notes.append(f"· {review_n} 段确定度不足，会跳过不改"
                             f"（可先在中间栏确认）")
            if not p.parent.is_dir():
                notes.append(f"· 这个文件夹不存在：{p.parent}")
            elif p.exists():
                notes.append(f"· {p.name} 已存在，将被覆盖")
            notes_lbl.configure(
                text="\n".join(notes),
                style="Warn.TLabel" if notes else "Mic.TLabel")

        target_var.trace_add("write", refresh_notes)
        refresh_notes()

        ttk.Label(pad, text="原件不动，改动只写进新文件。", style="Mic.TLabel").pack(
            anchor="w", pady=(8, 0))

        open_var = tk.BooleanVar(value=True)
        CheckRow(pad, "完成后打开这份文件", open_var, font=self.fonts["body"]).pack(
            anchor="w", pady=(10, 0))

        res = {"go": False, "open": True, "target": target}

        def go():
            p = Path(target_var.get().strip() or str(target))
            if p.suffix.lower() != ".docx":
                p = p.with_suffix(".docx")
            if not p.parent.is_dir():
                messagebox.showwarning("按规则排版",
                                       f"这个文件夹不存在：\n{p.parent}", parent=win)
                return
            if p.resolve() == Path(self.cur).resolve():
                messagebox.showwarning(
                    "按规则排版",
                    "输出路径不能和原文件相同——那样会覆盖原件。\n换个名字或换个文件夹。",
                    parent=win)
                return
            res.update(go=True, open=open_var.get(), target=p)
            win.destroy()

        row = ttk.Frame(pad, style="Card.TFrame")
        row.pack(fill="x", pady=(14, 0))
        ttk.Button(row, text="开始排版", style="Primary.TButton", command=go).pack(side="right")
        ttk.Button(row, text="取消", style="Ghost.TButton",
                   command=win.destroy).pack(side="right", padx=(0, 8))
        _center_dialog(win)
        win.grab_set()
        self.wait_window(win)
        return res if res["go"] else None

    def _pick_out_path(self, var) -> None:
        """选一个保存位置：默认还是原文件旁边的名字。"""
        cur = Path(var.get().strip())
        chosen = filedialog.asksaveasfilename(
            title="排好的文件存到哪",
            defaultextension=".docx",
            initialdir=str(cur.parent if cur.parent.is_dir() else sysenv.desktop_dir()),
            initialfile=cur.name,
            filetypes=[("Word 文档", "*.docx"), ("全部文件", "*.*")],
        )
        if chosen:
            var.set(chosen)

    def run_format(self) -> None:
        if not self.doc or self.busy:
            return
        from .core.apply import apply

        src = self.cur
        target = self._target_of(src)
        _auto, review = classify(self.doc.paras, 0.45, self.overrides)
        plan = self._ask_format(target, len(review))
        if not plan:
            return
        target = plan["target"]      # 用户在对话框里可能改过输出位置

        self._busy(True, "正在排版…")
        rules = self._profile_rules()
        ov = dict(self.overrides)
        do_layout = self.do_layout.get()
        page = dict(top_cm=self.doc.page.top_cm, bottom_cm=self.doc.page.bottom_cm,
                    left_cm=self.doc.page.left_cm, right_cm=self.doc.page.right_cm)
        n_tables = self.doc.n_tables

        def job():
            try:
                path, _doc, skipped = apply(src, rules, target, 0.45, ov)
                self.msgq.put(("format", str(path), len(skipped), plan["open"]))
                if do_layout:
                    from .core import layout_check
                    rep = layout_check.check(str(path), page, n_tables)
                    self.msgq.put(("layout", rep))
            except Exception as e:  # noqa: BLE001
                self.msgq.put(("error", f"排版失败：{e}"))

        threading.Thread(target=job, daemon=True).start()

    def _show_result(self, path: Path) -> None:
        self.last_out = path
        self.res_lbl.configure(text=f"已生成 {path.name}")
        self.res_bar.grid()

    def _open_last(self) -> None:
        if self.last_out and self.last_out.exists():
            open_path(str(self.last_out))
        else:
            self.open_dir()

    # ------------------------------------------------------------ 导出 PDF
    def export_last_pdf(self) -> None:
        """把刚排好的那份直接转成 PDF，省得再切到「格式转换」页。"""
        if not self.last_out:
            messagebox.showinfo(APP_NAME, "先排一份，再点这里转 PDF。")
            return
        self.export_pdf(str(self.last_out))

    def export_findings_csv(self) -> None:
        """把检查结果导成 CSV（Excel 直接打开，便于分工改稿）。"""
        from .core import report as reportmod
        from .core.model import ROLE_CN

        if not self.files:
            messagebox.showinfo(APP_NAME, "先添加文件")
            return
        dest = filedialog.asksaveasfilename(
            title="导出检查结果", defaultextension=".csv",
            initialdir=str(sysenv.desktop_dir()),
            initialfile="公文检查结果.csv",
            filetypes=[("CSV 表格", "*.csv")])
        if not dest:
            return
        self._busy(True, "正在汇总检查结果…")
        files = list(self.files)
        rules = self._profile_rules()
        text_on = self.do_text.get()

        def job():
            import csv

            entries, failed, lock = [], [], threading.Lock()

            def one(f):
                try:
                    doc, findings, review = self._collect(f, rules, text_on)
                    with lock:
                        entries.append((f, doc, findings, review))
                except Exception as e:  # noqa: BLE001
                    with lock:
                        failed.append(f"{Path(f).name}: {e}")

            _parallel(files, one)
            entries.sort(key=lambda t: files.index(t[0]))
            try:
                with open(dest, "w", newline="", encoding="utf-8-sig") as fh:
                    w = csv.writer(fh)
                    w.writerow(["文件", "健康分", "等级", "位置", "角色",
                                "检查项", "现在", "应为"])
                    for _f, doc, fs, _rv in entries:
                        pts = reportmod.score(fs)[0]
                        if not fs:
                            w.writerow([doc.name, pts, "", "整篇", "",
                                        "未发现问题", "", ""])
                        for x in fs:
                            w.writerow([
                                doc.name, pts,
                                reportmod.LEVEL_CN.get(x.level, x.level),
                                f"第{x.index + 1}段" if x.index is not None else "整篇",
                                ROLE_CN.get(x.role, x.role),
                                x.field, x.actual, x.expected])
                msg = f"检查结果已导出：{Path(dest).name}"
                if failed:
                    msg += f"（{len(failed)} 份读取失败）"
                self.msgq.put(("done", msg))
            except Exception as e:  # noqa: BLE001
                self.msgq.put(("error", f"导出 CSV 失败：{e}"))

        threading.Thread(target=job, daemon=True).start()

    def show_diff(self) -> None:
        """对比「当前文件」与「排版后副本」的段落差异。"""
        if not self.cur:
            messagebox.showinfo(APP_NAME, "先选一份公文")
            return
        src = Path(self.cur)
        cands = []
        if self.last_out:
            cands.append(Path(self.last_out))
        cands.append(src.with_name(f"{src.stem}-排版后.docx"))
        cands.append(src.with_name(f"{src.stem}-排版.docx"))
        target = next((p for p in cands if p.exists()), None)
        if target is None:
            messagebox.showinfo(
                APP_NAME,
                "还没排版过这份文件。\n先点「按规则排版…」生成副本，回来就能对比。")
            return
        try:
            old = parse(str(src))
            new = parse(str(target))
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(APP_NAME, f"读不了文件：\n{e}")
            return

        a = [(p.text.strip(), p.role) for p in old.paras if p.text.strip()]
        b = [(p.text.strip(), p.role) for p in new.paras if p.text.strip()]
        rows = []
        for i in range(max(len(a), len(b))):
            x = a[i] if i < len(a) else None
            y = b[i] if i < len(b) else None
            if x and y and x[0] == y[0] and x[1] == y[1]:
                continue
            rows.append((i + 1, x, y))

        win = tk.Toplevel(self)
        win.title(f"排版前后差异 — {src.name}")
        win.configure(background=theme.CANVAS)
        win.transient(self)
        head = ttk.Label(win, text=f"对比：当前文件 ↔ {target.name}", style="H3.TLabel")
        head.pack(anchor="w", padx=12, pady=(10, 4))
        ttk.Label(win, text=f"共 {len(rows)} 处不同（文字或角色有变化的段落）",
                  style="Muted.TLabel").pack(anchor="w", padx=12)
        wrap = ttk.Frame(win, style="Card.TFrame")
        wrap.pack(fill="both", expand=True, padx=12, pady=(6, 10))
        cols = ("no", "old", "new")
        tv = ttk.Treeview(wrap, columns=cols, show="headings", height=16)
        for c, t, w in (("no", "#", 44), ("old", "当前文件", 320), ("new", "排版后", 320)):
            tv.heading(c, text=t)
            tv.column(c, width=w, anchor="w" if c != "no" else "center")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        tv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        for no, x, y in rows:
            o = "（无）" if x is None else f"[{ROLE_CN.get(x[1].value, x[1].value)}] {x[0]}"
            n = "（无）" if y is None else f"[{ROLE_CN.get(y[1].value, y[1].value)}] {y[0]}"
            tv.insert("", "end", values=(no, o, n))
        _center_dialog(win)

    def backup_prefs(self) -> None:
        """把设置和自建的格式标准打成一个 zip。"""
        import zipfile

        dest = filedialog.asksaveasfilename(
            title="备份设置", defaultextension=".zip",
            initialdir=str(sysenv.desktop_dir()),
            initialfile="公文工具箱设置备份.zip",
            filetypes=[("压缩包", "*.zip")])
        if not dest:
            return
        try:
            n = 0
            with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z:
                pp = sysenv.prefs_path()
                if pp.exists():
                    z.write(pp, "prefs.json")
                    n += 1
                rd = rulesmod.USER_DIR
                if rd.exists():
                    for f in sorted(rd.glob("*.y*ml")):
                        z.write(f, f"rules/{f.name}")
                        n += 1
            messagebox.showinfo(APP_NAME, f"已备份 {n} 个文件到：\n{dest}")
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(APP_NAME, f"备份失败：\n{e}")

    def restore_prefs(self) -> None:
        """从备份 zip 恢复设置和格式标准。"""
        import zipfile

        src = filedialog.askopenfilename(
            title="恢复设置", filetypes=[("压缩包", "*.zip")])
        if not src:
            return
        if not messagebox.askyesno(
                APP_NAME,
                "会用备份里的格式标准和设置覆盖当前的。\n"
                "同名格式标准会被替换，其它文件不动。继续？"):
            return
        try:
            n = 0
            with zipfile.ZipFile(src) as z:
                for name in z.namelist():
                    if name == "prefs.json":
                        sysenv.prefs_path().write_bytes(z.read(name))
                        n += 1
                    elif (name.startswith("rules/")
                          and name.lower().endswith((".yaml", ".yml"))):
                        rulesmod.USER_DIR.mkdir(parents=True, exist_ok=True)
                        (rulesmod.USER_DIR / Path(name).name).write_bytes(z.read(name))
                        n += 1
            rulesmod._load_path.cache_clear()
            self._prefs = sysenv.load_prefs()
            self._refresh_profile_combo()
            messagebox.showinfo(APP_NAME, f"已恢复 {n} 个文件。\n格式标准列表已刷新。")
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(APP_NAME, f"恢复失败：\n{e}")

    def clean_zombies(self) -> None:
        """手动清理没有窗口的 WPS/ET 残留进程。"""
        try:
            killed = office.kill_zombies()
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(APP_NAME, f"清理失败：\n{e}")
            return
        if killed:
            names = "、".join(f"{n}({p})" for p, n in killed[:12])
            messagebox.showinfo(APP_NAME, f"已结束 {len(killed)} 个残留进程：\n{names}")
        else:
            messagebox.showinfo(APP_NAME, "没有发现需要清理的残留进程。")

    def export_pdf(self, src: str | None) -> None:
        if not src or self.busy:
            if not src:
                messagebox.showinfo(APP_NAME, "先在左边选一份文件。")
            return
        src_p = Path(src)
        if not src_p.exists():
            messagebox.showinfo(APP_NAME, f"文件不在了：{src_p}")
            return
        dst = src_p.with_suffix(".pdf")
        if dst.exists():
            if not messagebox.askyesno(APP_NAME, f"{dst.name} 已经存在，覆盖它吗？"):
                return
        self._busy(True, "正在转 PDF…")

        def job():
            try:
                from .core import office

                office.word_to_pdf(str(src_p), str(dst))
                self.msgq.put(("pdf_done", str(dst)))
            except Exception as e:  # noqa: BLE001
                self.msgq.put(("error", f"转 PDF 失败：{e}"))
            finally:
                # 单发任务：收尾把实例关掉，别留后台进程
                try:
                    from .core import office as _of

                    _of.release_apps()
                    _of.kill_zombies()
                except Exception:  # noqa: BLE001
                    pass

        threading.Thread(target=job, daemon=True).start()

    # ------------------------------------------------------------ 批量排版
    def run_batch_format(self) -> None:
        """左侧所有文件按当前格式标准各排一份。"""
        if not self.files or self.busy:
            if not self.files:
                messagebox.showinfo(APP_NAME, "左边还没有文件。")
            return
        from .core.apply import apply

        files = list(self.files)
        rules = self._profile_rules()
        name = rules.get("name", "") or self._profile_key()
        plan = self._ask_batch_format(len(files), name)
        if not plan:
            return
        same_dir, out_dir = plan["same_dir"], plan["out_dir"]

        self._busy(True, f"正在排版 1/{len(files)}…")
        n = len(files)

        def job():
            ok, failed = [], []
            lock = threading.Lock()
            done = [0]
            if not same_dir:
                try:
                    Path(out_dir).mkdir(parents=True, exist_ok=True)
                except OSError as e:
                    self.msgq.put(("error", f"建不了输出文件夹：{e}"))
                    return

            def one(f):
                try:
                    src = Path(f)
                    dst = (src.with_name(f"{src.stem}-排版后.docx") if same_dir
                           else Path(out_dir) / f"{src.stem}-排版后.docx")
                    path, _doc, _skipped = apply(str(src), rules, str(dst), 0.45)
                    with lock:
                        ok.append(str(path))
                except Exception as e:  # noqa: BLE001
                    with lock:
                        failed.append(f"{Path(f).name}：{e}")
                with lock:
                    done[0] += 1
                    self.msgq.put(("batch_fmt_progress", done[0], n))

            _parallel(files, one)
            self.msgq.put(("batch_fmt_done", ok, failed, "" if same_dir else out_dir))

        threading.Thread(target=job, daemon=True).start()

    # ------------------------------------------------------------ 一键流水线
    def run_pipeline(self) -> None:
        """体检 → 按规则排版 → 导出 PDF → 汇总报告，一条龙。"""
        if not self.files or self.busy:
            if not self.files:
                messagebox.showinfo(APP_NAME, "左边还没有文件。")
            return
        from .core import report as reportmod
        from .core.apply import apply

        files = list(self.files)
        rules = self._profile_rules()
        name = rules.get("name", "") or self._profile_key()
        plan = self._ask_pipeline(len(files), name)
        if not plan:
            return
        out_dir, want_pdf, want_report = plan["out_dir"], plan["pdf"], plan["report"]
        text_on = self.do_text.get()

        self._busy(True, f"正在处理 1/{len(files)}…")
        n = len(files)

        def job():
            ok, failed, lock = [], [], threading.Lock()
            done = [0]
            try:
                Path(out_dir).mkdir(parents=True, exist_ok=True)
            except OSError as e:
                self.msgq.put(("error", f"建不了输出文件夹：{e}"))
                return

            def one(f):
                try:
                    src = Path(f)
                    dst = Path(out_dir) / f"{src.stem}-排版后.docx"
                    doc, findings, review = self._collect(f, rules, text_on)
                    apply(str(src), rules, str(dst), 0.45)
                    with lock:
                        ok.append((str(src), str(dst), doc, findings, review))
                except Exception as e:  # noqa: BLE001
                    with lock:
                        failed.append(f"{Path(f).name}：{e}")
                with lock:
                    done[0] += 1
                    self.msgq.put(("batch_fmt_progress", done[0], n))

            _parallel(files, one)
            ok.sort(key=lambda t: files.index(t[0]))

            pdf_n = 0
            if want_pdf and ok:
                from .core import office

                for i, (_src, dst, _d, _fs, _rv) in enumerate(ok, 1):
                    try:
                        office.word_to_pdf(dst, str(Path(dst).with_suffix(".pdf")))
                        pdf_n += 1
                    except Exception as e:  # noqa: BLE001
                        failed.append(f"{Path(dst).name} 转 PDF：{e}")
                    self.msgq.put(("pipeline_pdf", i, len(ok)))
                try:
                    office.release_apps()
                    office.kill_zombies()
                except Exception:  # noqa: BLE001
                    pass

            rpt = ""
            if want_report and ok:
                try:
                    rpt = str(Path(out_dir) / "体检汇总.docx")
                    reportmod.save(rpt, None, None, None, name,
                                   entries=[(d, fs, rv) for _s, _dst, d, fs, rv in ok])
                except Exception as e:  # noqa: BLE001
                    failed.append(f"生成报告：{e}")
                    rpt = ""

            self.msgq.put(("pipeline_done", len(ok), failed, out_dir, pdf_n, rpt))

        threading.Thread(target=job, daemon=True).start()

    def _ask_pipeline(self, count: int, profile_name: str) -> dict | None:
        win = tk.Toplevel(self)
        win.title("一键处理")
        win.configure(background=theme.CANVAS)
        win.resizable(False, False)
        win.transient(self)
        pad = ttk.Frame(win, style="Card.TFrame", padding=18)
        pad.pack(fill="both", expand=True)

        ttk.Label(pad, text=f"把左边 {count} 份文件走一遍：体检 → 排版 → 导 PDF → 报告",
                  style="H2.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(pad, text=f"格式标准：{profile_name}", style="Muted.TLabel").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(4, 10))

        ttk.Label(pad, text="结果放到：", style="Card.TLabel").grid(
            row=2, column=0, sticky="w")
        out_var = tk.StringVar(value=str(sysenv.desktop_dir() / "公文处理结果"))
        ttk.Entry(pad, textvariable=out_var, width=38).grid(
            row=3, column=0, sticky="ew", pady=(2, 0))

        def pick():
            d = filedialog.askdirectory(title="结果放到哪", initialdir=out_var.get())
            if d:
                out_var.set(d)

        ttk.Button(pad, text="浏览…", style="Ghost.TButton", command=pick).grid(
            row=3, column=1, sticky="w", padx=(6, 0))

        pdf_var = tk.BooleanVar(value=True)
        rpt_var = tk.BooleanVar(value=True)
        CheckRow(pad, "同时导出 PDF", variable=pdf_var, bg=theme.SURFACE).grid(
            row=4, column=0, columnspan=2, sticky="w", pady=(10, 0))
        CheckRow(pad, "生成一份体检汇总报告（Word）", variable=rpt_var,
                 bg=theme.SURFACE).grid(row=5, column=0, columnspan=2, sticky="w",
                                        pady=(4, 0))
        ttk.Label(pad, text="每份文件输出「原名-排版后.docx」，原件一律不动。",
                  style="Mic.TLabel").grid(row=6, column=0, columnspan=2,
                                           sticky="w", pady=(8, 0))

        res: dict = {}

        def go():
            res.update(out_dir=out_var.get().strip(), pdf=pdf_var.get(),
                       report=rpt_var.get())
            win.destroy()

        bar = ttk.Frame(pad, style="Card.TFrame")
        bar.grid(row=7, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(bar, text="取消", style="Ghost.TButton", command=win.destroy).pack(
            side="right")
        ttk.Button(bar, text="开始", style="Primary.TButton", command=go).pack(
            side="right", padx=(0, 8))

        _center_dialog(win)
        win.grab_set()
        win.wait_window()
        return res or None

    def _ask_batch_format(self, count: int, profile_name: str) -> dict | None:
        """批量排版的确认框：排到哪、排成什么。"""
        win = tk.Toplevel(self)
        win.title("排版全部文件")
        win.configure(background=theme.CANVAS)
        win.resizable(False, False)
        win.transient(self)
        pad = ttk.Frame(win, style="Card.TFrame", padding=18)
        pad.pack(fill="both", expand=True)

        ttk.Label(pad, text=f"把左边 {count} 份文件各排一份",
                  style="H2.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(pad, text=f"格式标准：{profile_name}",
                  style="Muted.TLabel").grid(row=1, column=0, columnspan=2,
                                             sticky="w", pady=(4, 10))

        where = tk.StringVar(value="same")
        ttk.Radiobutton(pad, text="每份文件旁边（原名-排版后.docx）", value="same",
                        variable=where, style="Card.TRadiobutton").grid(
            row=2, column=0, columnspan=2, sticky="w")
        ttk.Radiobutton(pad, text="统一放到一个文件夹：", value="dir",
                        variable=where, style="Card.TRadiobutton").grid(
            row=3, column=0, sticky="w", pady=(4, 0))

        out_var = tk.StringVar(value=str(sysenv.desktop_dir() / "排版结果"))
        ent = ttk.Entry(pad, textvariable=out_var, width=38)
        ent.grid(row=4, column=0, sticky="ew", pady=(2, 0))

        def pick():
            d = filedialog.askdirectory(title="排好的文件放到哪", initialdir=out_var.get())
            if d:
                out_var.set(d)

        ttk.Button(pad, text="浏览…", style="Ghost.TButton",
                   command=pick).grid(row=4, column=1, sticky="w", padx=(6, 0))
        ttk.Label(pad, text="已存在的同名文件会被覆盖。原件一律不动。",
                  style="Mic.TLabel").grid(row=5, column=0, columnspan=2,
                                           sticky="w", pady=(8, 0))

        res: dict = {}

        def go():
            same = where.get() == "same"
            target = out_var.get().strip()
            if not same and not target:
                messagebox.showwarning(APP_NAME, "选个文件夹，或者改成「每份文件旁边」。",
                                       parent=win)
                return
            res.update(go=True, same_dir=same, out_dir=target)
            win.destroy()

        row = ttk.Frame(pad, style="Card.TFrame")
        row.grid(row=6, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(row, text="开始排版", style="Primary.TButton",
                   command=go).pack(side="right")
        ttk.Button(row, text="取消", style="Ghost.TButton",
                   command=win.destroy).pack(side="right", padx=(0, 8))
        _center_dialog(win)
        win.grab_set()
        self.wait_window(win)
        return res if res.get("go") else None

    # ------------------------------------------------------- 关闭与记忆设置
    def _restore_geometry(self) -> bool:
        """按上次的窗口大小位置开窗；越界就夹回可用桌面。"""
        raw = str(self._prefs.get("geometry") or "").strip()
        if not raw:
            return False
        try:
            size, _, pos = raw.partition("+")
            w, h = (int(x) for x in size.split("x"))
            x, y = (int(v) for v in pos.split("+")) if pos else (0, 0)
            _l, _t, aw, ah = work_area()
            if not aw or not ah:
                aw, ah = self.winfo_screenwidth(), self.winfo_screenheight()
            if w < 700 or h < 500 or w > aw or h > ah:
                return False
            # 上次的窗口比四栏的最低需求还窄时，别照抄——否则打开就是断栏
            self.update_idletasks()
            # 四栏并排是这个软件的基本盘，窗口再窄就没法看了。
            # 取「内容需求」和「桌面宽度的 62%」里大的那个作为下限。
            need_w = max(self.winfo_reqwidth(), int(aw * 0.62))
            if w < need_w <= int(aw * 0.98):
                w = need_w
            x = max(0, min(x, aw - 200))
            y = max(0, min(y, ah - 200))
            self.geometry(f"{w}x{h}+{x}+{y}")
            if str(self._prefs.get("window_state") or "") == "zoomed":
                self.state("zoomed")
            return True
        except Exception:  # noqa: BLE001
            return False

    def _save_prefs(self) -> None:
        p = self._prefs
        p["ui_density"] = self.ui_density
        try:
            p["profile"] = self._profile_key()
        except Exception:  # noqa: BLE001
            pass
        try:
            st = self.state()
            p["window_state"] = st
            if st in ("normal", "zoomed"):
                p["geometry"] = self.geometry()
        except Exception:  # noqa: BLE001
            pass
        for key, getter in (
            ("conv_target", lambda: self.conv._target_key()),
            ("conv_out", lambda: self.conv.out_dir.get()),
            ("ocr_mode", lambda: self.ocr_pane.mode.get()),
            ("ocr_out", lambda: self.ocr_pane.out_dir.get()),
            ("ocr_edit", lambda: bool(self.ocr_pane.editable.get())),
            ("nav_page", lambda: self.nav.current_text()),
        ):
            try:
                p[key] = getter()
            except Exception:  # noqa: BLE001
                pass
        sysenv.save_prefs(p)

    def _on_close(self) -> None:
        try:
            self._save_prefs()
        finally:
            # 退出前把自动化用过的办公软件实例关掉，别在后台攒进程
            try:
                office.release_apps()
            except Exception:  # noqa: BLE001
                pass
            self.destroy()

    # ---------------------------------------------------------------- 轮询
    def _poll(self) -> None:
        # 逐条处理：一条坏消息不会把消息泵打死，处理完再注册下一次。
        pump.drain(self.msgq, self._handle_msg)
        self.after(120, self._poll)

    def _handle_msg(self, msg: tuple) -> None:
        kind = msg[0]
        if kind == "format":
            _, path, n, want_open = msg
            self._busy(False)
            skip = f"，跳过 {n} 段" if n else ""
            self.status.set(f"已排版：{Path(path).name}{skip}　原件没动")
            self._show_result(Path(path))
            if want_open:
                open_path(path)
        elif kind == "layout":
            rep = msg[1]
            self._busy(False)
            # 重复检查时先清掉上一次的版面问题，否则会堆出重复行
            if self._layout_rows:
                for iid in self._layout_rows:
                    try:
                        self.finds.delete(iid)
                    except Exception:  # noqa: BLE001
                        pass
                self._layout_rows.clear()
            if rep.skipped:
                self.status.set(self.finds and (self.status.get() + f"　（{rep.skipped}）"))
            elif not rep.issues:
                self.status.set(self.status.get() + f"　版面正常，共 {rep.pages} 页")
            else:
                self.status.set(self.status.get() + f"　版面问题 {len(rep.issues)} 处")
                for i in rep.issues:
                    mark = "●" if i.level == "error" else "▲"
                    iid = self.finds.insert("", "end", text=mark, tags=(i.level,),
                                            values=(f"第{i.page}页", i.kind,
                                                    i.detail))
                    self._layout_rows.append(iid)
                self._toggle_placeholder("finds", False)
        elif kind == "prev":
            _, tok, pages, n, mode, err = msg
            if tok != self._prev_tok:
                return
            if err:
                self.prev_text.set("预览失败")
                self.preview.clear(err)
            else:
                self.prev_text.set(_pages_note(pages, n))
                self.preview.set_pages(pages, f"{n} 页")
                self._toggle_placeholder("preview", False)
        elif kind == "batch":
            _, out, failed = msg
            self._busy(False)
            self._fill_batch(out, failed)
        elif kind == "done":
            self._busy(False)
            self.status.set(msg[1])
        elif kind == "pdf_done":
            self._busy(False)
            path = Path(msg[1])
            size = ""
            try:
                size = f"（{path.stat().st_size / 1024:.0f} KB）"
            except Exception:  # noqa: BLE001
                pass
            self.status.set(f"已转成 PDF：{path.name}{size}")
            messagebox.showinfo(APP_NAME, f"PDF 已生成：\n{path}")
        elif kind == "batch_fmt_progress":
            _, i, n = msg
            self.status.set(f"正在排版 {i}/{n}…")
        elif kind == "batch_fmt_done":
            _, ok, failed, out_dir = msg
            self._busy(False)
            msg_txt = f"排版完成 {len(ok)} 份"
            if failed:
                msg_txt += f"，失败 {len(failed)} 份"
            self.status.set(msg_txt)
            lines = [msg_txt]
            if out_dir:
                lines.append(f"输出到：{out_dir}")
            else:
                lines.append("输出到每份原件旁边")
            if failed:
                lines.append("\n失败明细：")
                lines.extend(f"· {x}" for x in failed[:12])
                if len(failed) > 12:
                    lines.append(f"… 还有 {len(failed) - 12} 份")
            messagebox.showinfo(APP_NAME, "\n".join(lines))
            if out_dir and Path(out_dir).exists():
                sysenv.open_folder(out_dir)
        elif kind == "pipeline_pdf":
            _, i, n = msg
            self.status.set(f"正在导出 PDF {i}/{n}…")
        elif kind == "pipeline_done":
            _, ok_n, failed, out_dir, pdf_n, rpt = msg
            self._busy(False)
            head = f"一键处理完成 {ok_n} 份"
            if pdf_n:
                head += f"，导出 PDF {pdf_n} 份"
            if failed:
                head += f"，{len(failed)} 项失败"
            self.status.set(head)
            lines = [head, f"输出到：{out_dir}"]
            if rpt:
                lines.append(f"体检汇总：{Path(rpt).name}")
            if failed:
                lines.append("")
                lines.append("失败明细：")
                lines.extend(f"· {x}" for x in failed[:10])
                if len(failed) > 10:
                    lines.append(f"… 还有 {len(failed) - 10} 项")
            messagebox.showinfo(APP_NAME, "\n".join(lines))
            if Path(out_dir).exists():
                sysenv.open_folder(out_dir)
        elif kind == "error":
            self._busy(False)
            self.status.set(msg[1])
            messagebox.showerror(APP_NAME, msg[1])


def main() -> None:
    App().mainloop()


if __name__ == "__main__":
    main()
