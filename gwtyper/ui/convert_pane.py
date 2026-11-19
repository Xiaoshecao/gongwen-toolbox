# -*- coding: utf-8 -*-
"""格式转换页：公文之外的那些杂活——Word/Excel/PPT/PDF/图片 互转。

和排版页共用同一套转换引擎（`core.office`），所以 Windows 走 WPS/Office，
Linux 上有 WPS 就走 WPS、没有就走 LibreOffice。
"""
from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from ..core import convert as cv
from ..core import formats, office, pdf_engine, preview as prevmod
from ..core import sysenv
from . import theme
from .check import CheckRow
from .preview_pane import PreviewPane

APP_FILETYPES = [
    ("Word / Excel / PPT / PDF / 图片 / 文本",
     "*.doc *.docx *.docm *.rtf *.wps *.xls *.xlsx *.xlsm *.csv *.et "
     "*.ppt *.pptx *.pps *.ppsx *.dps *.pdf *.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp *.txt"),
    ("全部文件", "*.*"),
]

IMG_FILETYPES = [
    ("图片", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp *.gif"),
    ("全部文件", "*.*"),
]

PDF_FILETYPES = [("PDF", "*.pdf"), ("全部文件", "*.*")]


def human_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024.0
    return f"{n:.1f} GB"


class ConvertPane(ttk.Frame):
    """一个页签：左选文件、右设目标、下面一栏 PDF 小工具。"""

    def __init__(self, master, default_dir: str = "", prefs: dict | None = None):
        super().__init__(master)
        self.files: list[str] = []
        self.busy = False
        self.msgq: "queue.Queue[tuple]" = queue.Queue()
        self.last_dir: str = default_dir or str(sysenv.desktop_dir())
        self._prefs = prefs if prefs is not None else {}

        _tkey = str(self._prefs.get("conv_target") or "PDF")
        if _tkey not in formats.TARGETS:
            _tkey = "PDF"
        _out = str(self._prefs.get("conv_out") or "").strip()
        self.target = tk.StringVar(value=_tkey)
        self.out_dir = tk.StringVar(value=_out or str(Path(self.last_dir) / "转换结果"))
        self.dpi = tk.IntVar(value=150)
        self.split_every = tk.IntVar(value=1)
        self.orientation = tk.StringVar(value="auto")
        self.margin_mm = tk.DoubleVar(value=0.0)
        self.rotate_content = tk.BooleanVar(value=False)
        self.pdf_engine = tk.StringVar(value="auto")
        self.status = tk.StringVar(value="选几个文件，挑个目标格式，点「开始转换」。")
        self.count = tk.StringVar(value="0 个文件")
        self.engine = tk.StringVar(value="")
        # 预览用的状态：tok 递增，过期的渲染结果直接丢掉
        self._prev_tok = 0
        self._prev_job = None
        self._prev_iid = ""

        self._build()
        # 恢复上次选的目标格式与输出目录后，把提示和 PDF 选项的可用状态同步一遍
        self._sync_target()
        self.after(120, self._poll)
        threading.Thread(target=self._detect_engine, daemon=True).start()

    # ------------------------------------------------------------ 布局
    def _build(self) -> None:
        # 三栏的 minsize 之和要能塞进 1366 的笔记本：
        # 以前是 340+400+300 + 两侧 padding，实际需要 1749 宽，
        # 在 1366 上第三栏被压到几乎看不见、"自动"单选框宽度只剩 1px。
        self.columnconfigure(0, weight=3, minsize=260)
        self.columnconfigure(1, weight=4, minsize=330)
        self.columnconfigure(2, weight=3, minsize=240)
        self.rowconfigure(0, weight=1)

        self._build_files()
        self._build_options()
        self._build_preview()

    def _build_files(self) -> None:
        card = ttk.Frame(self, style="Card.TFrame", padding=12)
        card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        card.columnconfigure(0, weight=1)
        card.rowconfigure(2, weight=1)

        head = ttk.Frame(card, style="Card.TFrame")
        head.grid(row=0, column=0, sticky="ew")
        ttk.Label(head, text="要转换的文件", style="H2.TLabel").pack(side="left")
        ttk.Label(head, textvariable=self.count, style="Mic.TLabel").pack(
            side="left", padx=(8, 0), pady=(3, 0))

        # 三个按钮并排要 534px，会把左栏撑得放不进 1366 屏幕。
        # 改成两行：主操作 + 次操作分开，最窄也能用。
        bar = ttk.Frame(card, style="Card.TFrame")
        bar.grid(row=1, column=0, sticky="ew", pady=(8, 8))
        row1 = ttk.Frame(bar, style="Card.TFrame")
        row1.pack(anchor="w")
        ttk.Button(row1, text="添加文件…", style="Secondary.TButton",
                   command=self.add_files).pack(side="left")
        ttk.Button(row1, text="文件夹…", style="Secondary.TButton",
                   command=self.add_folder).pack(side="left", padx=(6, 0))
        row2 = ttk.Frame(bar, style="Card.TFrame")
        row2.pack(anchor="w", pady=(4, 0))
        ttk.Button(row2, text="清空列表", style="Ghost.TButton",
                   command=self.clear).pack(side="left")

        box = ttk.Frame(card, style="Card.TFrame")
        box.grid(row=2, column=0, sticky="nsew")
        box.columnconfigure(0, weight=1)
        box.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(box, columns=("name", "kind", "size"),
                                 show="headings", selectmode="extended")
        self.tree.heading("name", text="文件")
        self.tree.heading("kind", text="类型")
        self.tree.heading("size", text="大小")
        self.tree.column("name", width=180, anchor="w")
        self.tree.column("kind", width=52, anchor="center")
        self.tree.column("size", width=64, anchor="e")
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.bind("<Delete>", lambda e: self.remove_selected())
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.schedule_preview(220))

        # 注意顺序：占位标签要先建，_setup_dnd 里会按拖拽可用性改它的文案
        self.ph = ttk.Label(box, text="把文件拖到这里，\n或点「添加文件…」",
                            style="Mic.TLabel", justify="center",
                            wraplength=170)
        self.ph.place(relx=0.5, rely=0.45, anchor="center")
        self._setup_dnd()

        tip = ttk.Frame(card, style="Card.TFrame")
        tip.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(tip, text="输出不会覆盖原文件，重名自动加 (1)(2)。",
                  style="Mic.TLabel", wraplength=250).pack(side="left")

    def _build_preview(self) -> None:
        """右栏：选中文件长什么样。转换前先看一眼，免得转完才发现不对。"""
        holder = tk.Frame(self, background=theme.BORDER_SUBTLE)
        holder.grid(row=0, column=2, sticky="nsew", padx=(12, 0))
        card = ttk.Frame(holder, style="Card.TFrame", padding=12)
        card.pack(fill="both", expand=True, padx=1, pady=1)
        card.columnconfigure(0, weight=1)
        card.rowconfigure(1, weight=1)

        head = ttk.Frame(card, style="Card.TFrame")
        head.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(head, text="预览", style="H2.TLabel").pack(side="left")
        self.prev_note = tk.StringVar(value="")
        ttk.Label(head, textvariable=self.prev_note, style="Mic.TLabel").pack(
            side="left", padx=(8, 0), pady=(3, 0))
        ttk.Button(head, text="重绘", style="Ghost.TButton",
                   command=lambda: self.refresh_preview(force=True)).pack(side="right")

        self.preview = PreviewPane(card, width_hint=200)
        self.preview.grid(row=1, column=0, sticky="nsew")
        self.preview.hide_head()
        self.preview.clear("选一个文件，这里显示它现在的样子")

    def _build_options(self) -> None:
        # 选项区内容较多（转成什么/保存到/纸面/引擎/按钮/进度条/PDF小工具），
        # 1366×768 上整块有 1085 高、可视只有 583，底部会整个掉到屏幕外——
        # 「开始转换」和进度条都看不见，用户会以为程序卡死。
        # 这里把选项区放进可滚动容器，并把「开始转换 + 进度条」固定在底部。
        outer = tk.Frame(self, background=theme.BORDER_SUBTLE)
        outer.grid(row=0, column=1, sticky="nsew")
        shell = ttk.Frame(outer, style="Card.TFrame")
        shell.pack(fill="both", expand=True, padx=1, pady=1)
        shell.rowconfigure(0, weight=1)
        shell.columnconfigure(0, weight=1)

        canvas = tk.Canvas(shell, background=theme.SURFACE, highlightthickness=0,
                   bd=0, width=300)
        canvas.grid(row=0, column=0, sticky="nsew")
        vsb = ttk.Scrollbar(shell, orient="vertical", command=canvas.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=vsb.set)
        self._opt_canvas = canvas
        self._opt_vsb = vsb

        card = ttk.Frame(canvas, style="Card.TFrame", padding=12)
        self._opt_win = canvas.create_window((0, 0), window=card, anchor="nw")
        card.columnconfigure(0, weight=1)

        def _resize(_e=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # 只在内容超高时才显示滚动条
            try:
                need_h = card.winfo_reqheight()
                if need_h > canvas.winfo_height() + 2:
                    vsb.grid()
                else:
                    vsb.grid_remove()
                    canvas.yview_moveto(0)
            except Exception:  # noqa: BLE001
                pass

        card.bind("<Configure>", _resize)
        canvas.bind("<Configure>", _resize)
        # 让内层卡片跟着画布宽度走，避免横向被裁
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfigure(self._opt_win, width=e.width))

        def _wheel(event):
            if card.winfo_reqheight() > canvas.winfo_height():
                canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

        canvas.bind("<Enter>", lambda _e: canvas.bind_all("<MouseWheel>", _wheel))
        canvas.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))

        ttk.Label(card, text="转成什么", style="H2.TLabel").grid(row=0, column=0, sticky="w")
        self.target_box = ttk.Combobox(card, textvariable=self.target, state="readonly",
                                       values=[formats.TARGET_CN[t] for t in formats.TARGET_ORDER])
        self.target_box.set(formats.TARGET_CN.get(self.target.get(), formats.TARGET_CN["PDF"]))
        self.target_box.grid(row=1, column=0, sticky="w", pady=(8, 2), ipadx=6, ipady=2)
        self.target_box.bind("<<ComboboxSelected>>", lambda e: self._sync_target())
        self.target_tip = ttk.Label(card, text="", style="Mic.TLabel", wraplength=430)
        self.target_tip.grid(row=2, column=0, sticky="w")

        ttk.Label(card, text="保存到", style="H2.TLabel").grid(row=3, column=0, sticky="w", pady=(16, 0))
        row = ttk.Frame(card, style="Card.TFrame")
        row.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        row.columnconfigure(0, weight=1)
        ttk.Entry(row, textvariable=self.out_dir).grid(row=0, column=0, sticky="ew", ipady=3)
        ttk.Button(row, text="浏览…", style="Secondary.TButton",
                   command=self.choose_out).grid(row=0, column=1, padx=(6, 0))

        ttk.Separator(card, orient="horizontal").grid(row=5, column=0, sticky="ew", pady=(18, 14))

        ttk.Label(card, text="选项", style="H3.TLabel").grid(row=6, column=0, sticky="w")
        opt = ttk.Frame(card, style="Card.TFrame")
        opt.grid(row=7, column=0, sticky="w", pady=(8, 0))
        ttk.Label(opt, text="图片清晰度", style="Cap.TLabel").pack(side="left")
        ttk.Spinbox(opt, from_=72, to=600, increment=50, width=6,
                    textvariable=self.dpi).pack(side="left", padx=(6, 4))
        ttk.Label(opt, text="DPI（打印建议 300）", style="Mic.TLabel").pack(side="left")

        # ---- 转 PDF 的版面：方向 + 留白 -------------------------------
        self.pdf_box = ttk.Frame(card, style="Card.TFrame")
        self.pdf_box.grid(row=8, column=0, sticky="ew", pady=(12, 0))
        self.pdf_box.columnconfigure(0, weight=1)

        ttk.Label(self.pdf_box, text="转 PDF 的纸面", style="Cap.TLabel").grid(
            row=0, column=0, sticky="w")
        ori = ttk.Frame(self.pdf_box, style="Card.TFrame")
        ori.grid(row=1, column=0, sticky="w", pady=(4, 0))
        for label, val in (("保持原样", "auto"), ("统一竖版", "portrait"),
                           ("统一横版", "landscape")):
            ttk.Radiobutton(ori, text=label, value=val, variable=self.orientation,
                            style="Card.TRadiobutton",
                            command=self._sync_pdf_opts).pack(side="left", padx=(0, 10))

        mar = ttk.Frame(self.pdf_box, style="Card.TFrame")
        mar.grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Label(mar, text="四周留白", style="Cap.TLabel").pack(side="left")
        ttk.Spinbox(mar, from_=0, to=50, increment=5, width=5,
                    textvariable=self.margin_mm).pack(side="left", padx=(6, 4))
        ttk.Label(mar, text="毫米（0 = 不留白，内容等比缩放不裁切）",
                  style="Mic.TLabel").pack(side="left")

        self.rot_row = CheckRow(self.pdf_box, "横竖互换时把内容转 90°（宽表转竖页用）",
                                variable=self.rotate_content,
                                command=self._sync_pdf_opts)
        self.rot_row.grid(row=3, column=0, sticky="w", pady=(4, 0))

        self.pdf_tip = ttk.Label(self.pdf_box, text="", style="Mic.TLabel",
                                 wraplength=430, justify="left")
        self.pdf_tip.grid(row=4, column=0, sticky="w", pady=(4, 0))

        # ---- PDF 引擎选择 ---------------------------------------------
        # 三个选项竖排：横着排在 1366 屏幕上会被压扁，
        # 之前「自动」的宽度只剩 1px，而它恰恰是默认值。
        eng_row = ttk.Frame(self.pdf_box, style="Card.TFrame")
        eng_row.grid(row=5, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(eng_row, text="PDF 引擎", style="Cap.TLabel").pack(anchor="w")
        eng_opts = ttk.Frame(eng_row, style="Card.TFrame")
        eng_opts.pack(anchor="w", pady=(3, 0))
        for label, val in (("自动（推荐）", "auto"),
                           ("纯 Python（不依赖办公软件）", "pure"),
                           ("外部引擎（WPS/LibreOffice）", "external")):
            ttk.Radiobutton(eng_opts, text=label, value=val,
                            variable=self.pdf_engine,
                            style="Card.TRadiobutton",
                            command=self._sync_pdf_opts).pack(anchor="w")

        self.engine_tip = ttk.Label(self.pdf_box, text="", style="Mic.TLabel",
                                    wraplength=430, justify="left")
        self.engine_tip.grid(row=6, column=0, sticky="w", pady=(2, 0))

        act = ttk.Frame(card, style="Card.TFrame")
        act.grid(row=9, column=0, sticky="ew", pady=(18, 0))
        self.btn_go = ttk.Button(act, text="开始转换", style="Primary.TButton",
                                 command=self.run_convert)
        self.btn_go.pack(side="left")
        ttk.Button(act, text="打开输出目录", style="Ghost.TButton",
                   command=self.open_out).pack(side="left", padx=(6, 0))

        self.pb = ttk.Progressbar(card, mode="determinate")
        self.pb.grid(row=10, column=0, sticky="ew", pady=(14, 4))
        ttk.Label(card, textvariable=self.status, style="Cap.TLabel",
                  wraplength=430, justify="left").grid(row=11, column=0, sticky="w")
        ttk.Label(card, textvariable=self.engine, style="Mic.TLabel").grid(
            row=12, column=0, sticky="w", pady=(6, 0))

        ttk.Separator(card, orient="horizontal").grid(row=13, column=0, sticky="ew", pady=(16, 12))
        ttk.Label(card, text="PDF 小工具", style="H3.TLabel").grid(row=14, column=0, sticky="w")

        tools = ttk.Frame(card, style="Card.TFrame")
        tools.grid(row=15, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(tools, text="合并 PDF", style="Secondary.TButton",
                   command=self.tool_merge).pack(side="left")
        ttk.Button(tools, text="图片转 PDF", style="Secondary.TButton",
                   command=self.tool_images_to_pdf).pack(side="left", padx=(6, 0))

        tools2 = ttk.Frame(card, style="Card.TFrame")
        tools2.grid(row=16, column=0, sticky="ew", pady=(6, 0))
        ttk.Label(tools2, text="拆分 PDF 每", style="Cap.TLabel").pack(side="left")
        ttk.Spinbox(tools2, from_=1, to=500, width=5,
                    textvariable=self.split_every).pack(side="left", padx=(5, 4))
        ttk.Label(tools2, text="页一份", style="Cap.TLabel").pack(side="left")
        ttk.Button(tools2, text="选择并拆分", style="Secondary.TButton",
                   command=self.tool_split).pack(side="left", padx=(10, 0))

        tools3 = ttk.Frame(card, style="Card.TFrame")
        tools3.grid(row=17, column=0, sticky="ew", pady=(6, 0))
        ttk.Button(tools3, text="PDF 转图片", style="Secondary.TButton",
                   command=lambda: self.tool_pdf_to_images("png")).pack(side="left")
        ttk.Button(tools3, text="改方向／留白…", style="Secondary.TButton",
                   command=self.tool_reshape).pack(side="left", padx=(6, 0))

        tools4b = ttk.Frame(card, style="Card.TFrame")
        tools4b.grid(row=18, column=0, sticky="ew", pady=(6, 0))
        ttk.Button(tools4b, text="多份材料合并成一册（带书签）…", style="Secondary.TButton",
                   command=self.tool_merge_many).pack(side="left")

        # ---- 双层 PDF ----
        # 底层图像（看到的样子固定）+ 上层不可见文字（能搜能复制）。
        # 源文件有文字的（Word/WPS）走源文档提取，不用 OCR，100% 准确；
        # 扫描件/图片才走 OCR。
        tools4 = ttk.Frame(card, style="Card.TFrame")
        tools4.grid(row=19, column=0, sticky="ew", pady=(6, 0))
        ttk.Button(tools4, text="做双层 PDF…", style="Secondary.TButton",
                   command=self.tool_double_layer).pack(side="left")
        self.dl_tip = ttk.Label(tools4, text="", style="Mic.TLabel")
        self.dl_tip.pack(side="left", padx=(8, 0))

        self._sync_target()
        self._refresh_count()

    # ------------------------------------------------------------ 文件管理
    def add_files(self) -> None:
        picked = filedialog.askopenfilenames(title="选择要转换的文件",
                                             filetypes=APP_FILETYPES)
        if picked:
            self._add_paths([Path(p) for p in picked])

    def add_folder(self) -> None:
        d = filedialog.askdirectory(title="选择文件夹（只取里面一层）")
        if not d:
            return
        exts = {".doc", ".docx", ".docm", ".rtf", ".wps", ".xls", ".xlsx", ".xlsm",
                ".csv", ".et", ".ppt", ".pptx", ".pps", ".ppsx", ".dps", ".pdf",
                ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp", ".txt"}
        got = [Path(p) for p in sorted(Path(d).iterdir())
               if p.is_file() and p.suffix.lower() in exts]
        if not got:
            messagebox.showinfo("格式转换", "这个文件夹里没有能转换的文件。")
            return
        self._add_paths(got)

    def _add_paths(self, paths: list[Path]) -> None:
        first_new = ""
        for p in paths:
            sp = str(p)
            if sp in self.files:
                continue
            self.files.append(sp)
            try:
                size = human_size(p.stat().st_size)
            except OSError:
                size = "—"
            self.tree.insert("", "end", iid=sp,
                             values=(p.name, formats.kind_of(p), size))
            if not first_new:
                first_new = sp
        self._refresh_count()
        self._sync_target()
        # 加完自动选中第一个新文件，右栏立刻能看到它长什么样
        if first_new and not self.tree.selection():
            self.tree.selection_set(first_new)
            self.tree.see(first_new)

    def remove_selected(self) -> None:
        for iid in self.tree.selection():
            if iid in self.files:
                self.files.remove(iid)
            self.tree.delete(iid)
        self._refresh_count()
        self._sync_target()
        self._prev_iid = ""
        self.schedule_preview(120)

    def clear(self) -> None:
        self.tree.delete(*self.tree.get_children())
        self.files.clear()
        self._refresh_count()
        self._sync_target()
        self._prev_iid = ""
        self._prev_tok += 1
        self.preview.clear("选一个文件，这里显示它现在的样子")
        self.prev_note.set("")

    def _refresh_count(self) -> None:
        n = len(self.files)
        self.count.set(f"{n} 个文件")
        if n:
            self.ph.place_forget()
        else:
            self.ph.place(relx=0.5, rely=0.45, anchor="center")
        self.btn_go.configure(state="normal" if n and not self.busy else "disabled")

    # ------------------------------------------------------------ 拖拽
    def _setup_dnd(self) -> None:
        """把文件拖进来就加进列表。"""
        from . import dnd

        if not dnd.bind_targets([self.tree, self, getattr(self, "ph", None)],
                                self._on_drop):
            self.ph.configure(text="点「添加文件…」，或把文件拖到这里")

    def _on_drop(self, event) -> None:  # pragma: no cover - 需要真实拖拽
        from . import dnd

        paths = dnd.paths_from_event(self, event)
        if paths:
            self._add_paths(paths)

    # ------------------------------------------------------------ 预览
    def schedule_preview(self, delay: int = 700) -> None:
        """选中变了就排一次预览，连着点几下只跑最后一次。"""
        if self._prev_job:
            try:
                self.after_cancel(self._prev_job)
            except Exception:  # noqa: BLE001
                pass
        self._prev_job = self.after(delay, self.refresh_preview)

    def refresh_preview(self, force: bool = False) -> None:
        self._prev_job = None
        sel = self.tree.selection()
        if not sel:
            self.preview.clear("选一个文件，这里显示它现在的样子")
            self.prev_note.set("")
            self._prev_iid = ""
            return
        src = sel[0]
        if src != self._prev_iid:
            # 换了文件，先清干净，别让人以为在看新文件
            self.preview.clear("正在读这份文件…")
            self.prev_note.set("")
            self._prev_iid = src

        self._prev_tok += 1
        tok = self._prev_tok
        width = max(240, self.preview.canvas.winfo_width() - 30)

        def job() -> None:
            try:
                pdf = prevmod.any_pdf(src, force=force)
                pages, n = prevmod.render(pdf, max_width=width)
                self.msgq.put(("prev", tok, pages, n, None))
            except Exception as exc:  # noqa: BLE001
                self.msgq.put(("prev", tok, [], 0, str(exc)))

        threading.Thread(target=job, daemon=True).start()

    def _show_preview(self, tok: int, pages: list, n: int, err: str | None) -> None:
        if tok != self._prev_tok:      # 过期结果，丢掉
            return
        if err:
            self.prev_note.set("")
            self.preview.clear(err)
            return
        note = (f"{n} 页（预览前 {len(pages)} 页）"
                if len(pages) < n else f"{n} 页")
        self.prev_note.set(note)
        self.preview.set_pages(pages, note)

    # ------------------------------------------------------------ 选项
    def _target_key(self) -> str:
        label = self.target_box.get()
        for k, v in formats.TARGET_CN.items():
            if v == label:
                return k
        return self.target.get() or "PDF"

    def _sync_target(self) -> None:
        """按选中文件算出支持的目标格式，不支持的在提示里说清楚。"""
        if not self.files:
            self.target_tip.configure(text="先选文件，这里会说明这项能转成什么。")
            self._sync_pdf_opts()
            return
        key = self._target_key()
        spec = formats.TARGETS[key]
        bad = [Path(f).name for f in self.files if not formats.supported(f, key)]
        tip = spec["note"]
        if bad:
            tip += f"\n其中 {len(bad)} 个不支持：{bad[0]}" + ("…" if len(bad) > 1 else "")
        self.target_tip.configure(text=tip)
        self._sync_pdf_opts()

    # ---------------------------------------------------- 转 PDF 的版面选项
    def _sync_pdf_opts(self) -> None:
        """纸面选项只对「转成 PDF」有意义，别的目标格式下灰掉，避免误会。"""
        key = self._target_key()
        pdf_only = key == "PDF"
        state = "normal" if pdf_only else "disabled"
        for child in self.pdf_box.winfo_children():
            if isinstance(child, CheckRow):
                continue
            try:
                child.configure(state=state)
            except tk.TclError:
                for sub in child.winfo_children():
                    try:
                        sub.configure(state=state)
                    except tk.TclError:
                        pass
        self.rot_row.set_enabled(pdf_only)

        if not pdf_only:
            self.pdf_tip.configure(text="这些选项只在「转成 PDF 文档」时生效。")
            self.engine_tip.configure(text="")
            return
        mode = self.pdf_engine.get()
        if mode == "pure":
            self.engine_tip.configure(
                text="纯 Python 排版：Word/Excel/文本/图片 直接排成 PDF，"
                     "不调用任何办公软件。演示文稿（pptx）要走外部引擎。")
        elif mode == "external":
            self.engine_tip.configure(
                text=f"外部引擎：{office.engine_hint()}")
        else:
            hint = office.engine_hint()
            self.engine_tip.configure(
                text=f"自动：先用外部引擎（{hint}），"
                     f"转不了再退纯 Python。")
        ori = self.orientation.get()
        mar = self._margin()
        bits = []
        if ori == "portrait":
            bits.append("统一排成竖版")
        elif ori == "landscape":
            bits.append("统一排成横版")
        if mar > 0:
            bits.append(f"四周留白 {mar:g} 毫米")
        if not bits:
            self.pdf_tip.configure(text="保持原样：按源文件的纸张大小和方向输出。")
        else:
            self.pdf_tip.configure(
                text="、".join(bits) + "。内容等比缩放居中，不会裁掉任何东西。")

    def _margin(self) -> float:
        try:
            return max(0.0, float(self.margin_mm.get()))
        except (tk.TclError, ValueError):
            return 0.0

    def _make_options(self) -> "cv.Options":
        return cv.Options(
            dpi=int(self.dpi.get() or 150),
            orientation=self.orientation.get() or "auto",
            margin_mm=self._margin(),
            rotate_content=bool(self.rotate_content.get()),
            pdf_engine=self.pdf_engine.get() or "auto",
        )

    def choose_out(self) -> None:
        d = filedialog.askdirectory(title="输出到哪个文件夹",
                                    initialdir=self.out_dir.get() or self.last_dir)
        if d:
            self.out_dir.set(d)

    def open_out(self) -> None:
        d = self.out_dir.get().strip()
        if not d:
            messagebox.showinfo("格式转换", "还没设输出目录。")
            return
        Path(d).mkdir(parents=True, exist_ok=True)
        from ..core import sysenv

        sysenv.open_folder(d)

    # ------------------------------------------------------------ 引擎
    def _detect_engine(self) -> None:
        try:
            hint = office.engine_hint()
            eng = office.engine_name()
        except Exception:  # noqa: BLE001
            hint, eng = "未知", ""
        self.msgq.put(("engine", eng, hint))

    # ------------------------------------------------------------ 转换
    def run_convert(self) -> None:
        if self.busy:
            return
        if not self.files:
            messagebox.showinfo("格式转换", "先选几个文件。")
            return
        key = self._target_key()
        out_dir = self.out_dir.get().strip()
        if not out_dir:
            messagebox.showinfo("格式转换", "先选一个输出目录。")
            return
        # 有选中就只转选中的，没选中才转全部。
        # 以前无条件转全部：列表用 extended 模式、明摆着邀请「选一个再转」，
        # 结果选了 1 个却生成 3 个文件，是批量误操作的主要来源。
        sel = [s for s in self.tree.selection() if s in self.files]
        pool = sel if sel else list(self.files)
        todo = [f for f in pool if formats.supported(f, key)]
        skipped = len(pool) - len(todo)
        if not todo:
            messagebox.showwarning(
                "格式转换",
                f"选中的文件都转不了 {formats.TARGET_CN[key]}。\n"
                f"能转的是：{'、'.join(sorted(formats.TARGETS[key]['sources']))}")
            return
        try:
            Path(out_dir).mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            messagebox.showerror("格式转换", f"输出目录建不了：{exc}")
            return

        opt = self._make_options()
        self._set_busy(True, f"正在转换 {len(todo)} 个文件…")
        self.pb.configure(mode="determinate", maximum=len(todo), value=0)

        def job() -> None:
            done, failed = 0, []
            plain_pdf = 0
            for i, src in enumerate(todo, start=1):
                try:
                    made = cv.convert(src, key, out_dir, opt)
                    done += len(made)
                    if opt.extra.pop("pdf_to_docx_plain", False):
                        plain_pdf += 1
                except Exception as exc:  # noqa: BLE001
                    failed.append((Path(src).name, str(exc).splitlines()[0]))
                self.msgq.put(("progress", i, len(todo)))

            def finish():
                self._set_busy(False, "")
                self._after_convert(done, failed, out_dir, plain_pdf, skipped)

            self.msgq.put(("finish", finish))

        threading.Thread(target=job, daemon=True).start()

    def _after_convert(self, done: int, failed: list, out_dir: str,
                       plain_pdf: int = 0, skipped: int = 0) -> None:
        # PDF 转 Word 走的是「只抽文字」的兜底路径时，必须说清楚，
        # 否则用户会以为版式也能原样还原。
        hint = ""
        if skipped:
            # 部分文件转不了目标格式时以前是静默过滤掉的：
            # 5 份进 3 份出，用户不知道另外 2 份去哪了
            hint += (f"\n\n有 {skipped} 个文件不支持转成"
                     f"{self.target_box.get()}，已跳过。")
        if plain_pdf:
            hint = (f"\n\n注意：其中 {plain_pdf} 个 PDF 只提取了文字，"
                    f"版式（表格、图片位置）没有还原，需要在 Word 里再排一下。")
        if failed:
            lines = "\n".join(f"· {n}：{m}" for n, m in failed[:6])
            if len(failed) > 6:
                lines += f"\n… 还有 {len(failed) - 6} 个失败"
            self.status.set(f"完成 {done} 个，失败 {len(failed)} 个。")
            messagebox.showwarning(
                "格式转换",
                f"完成 {done} 个，失败 {len(failed)} 个：\n\n{lines}{hint}")
        elif done:
            self.status.set(f"完成，生成 {done} 个文件。")
            messagebox.showinfo(
                "格式转换", f"完成，生成 {done} 个文件。\n\n{out_dir}{hint}")
        else:
            self.status.set("没有生成文件，检查一下源文件。")

    def _set_busy(self, on: bool, text: str) -> None:
        self.busy = on
        if text:
            self.status.set(text)
        self.btn_go.configure(state="disabled" if on else
                              ("normal" if self.files else "disabled"))

    # ------------------------------------------------------------ 小工具
    def _tool_out_dir(self) -> str:
        d = self.out_dir.get().strip() or self.last_dir
        Path(d).mkdir(parents=True, exist_ok=True)
        return d

    def _ask_one(self, title: str, types) -> str:
        p = filedialog.askopenfilename(title=title, filetypes=types)
        return p or ""

    def _ask_many(self, title: str, types) -> list[str]:
        got = filedialog.askopenfilenames(title=title, filetypes=types)
        return list(got) if got else []

    def _unique(self, folder: str, base: str, ext: str) -> str:
        dst = Path(folder) / f"{base}{ext}"
        n = 1
        while dst.exists():
            dst = Path(folder) / f"{base}({n}){ext}"
            n += 1
        return str(dst)

    def tool_merge(self) -> None:
        files = self._ask_many("选择要合并的 PDF / Office 文档 / 图片（按选择顺序）",
                               APP_FILETYPES)
        if len(files) < 2:
            if files:
                messagebox.showinfo("合并 PDF", "至少选两个文件。")
            return
        out_dir = self._tool_out_dir()
        dst = self._unique(out_dir, "合并结果", ".pdf")
        opt = self._make_options()
        self._run_tool("合并 PDF", lambda: cv.merge_to_pdf(files, dst, opt), out_dir)

    def tool_merge_many(self) -> None:
        """多份材料合并成一册：Word/PDF/图片混着也行，每份加一个书签。"""
        files = self._ask_many("选择要合并的材料（按选择顺序）", APP_FILETYPES)
        if len(files) < 2:
            if files:
                messagebox.showinfo("合并成一册", "至少选两份材料。")
            return
        out_dir = self._tool_out_dir()
        dst = self._unique(out_dir, "合并材料", ".pdf")
        self._run_tool("合并成一册",
                       lambda: cv.merge_many(files, dst, bookmarks=True), out_dir)

    def tool_split(self) -> None:
        f = self._ask_one("选择要拆分的 PDF", PDF_FILETYPES)
        if not f:
            return
        out_dir = self._tool_out_dir()
        try:
            every = max(1, int(self.split_every.get()))
        except (tk.TclError, ValueError):
            every = 1
        self._run_tool("拆分 PDF",
                       lambda: pdf_engine.split_pdf(f, out_dir, every), out_dir)

    def tool_images_to_pdf(self) -> None:
        files = self._ask_many("选择图片（按选择顺序合成）", IMG_FILETYPES)
        if not files:
            return
        out_dir = self._tool_out_dir()
        dst = self._unique(out_dir, "图片合集", ".pdf")
        opt = self._make_options()
        self._run_tool("图片转 PDF",
                       lambda: cv.images_to_one_pdf(files, dst, opt), out_dir)

    def tool_pdf_to_images(self, fmt: str) -> None:
        f = self._ask_one("选择要导出图片的 PDF", PDF_FILETYPES)
        if not f:
            return
        out_dir = self._tool_out_dir()
        opt = cv.Options(dpi=int(self.dpi.get() or 150))
        self._run_tool(f"PDF 转 {fmt.upper()}",
                       lambda: cv.to_images(f, out_dir, fmt, opt), out_dir)

    def tool_reshape(self) -> None:
        """对现成文件改方向／留白，保留原文件，另存一份。"""
        f = self._ask_one("选择要改方向/留白的 PDF（也支持 Office 文档和图片）",
                          APP_FILETYPES)
        if not f:
            return
        opt = self._make_options()
        if opt.orientation == "auto" and opt.margin_mm <= 0:
            messagebox.showinfo("改方向／留白",
                                "先在右边选「统一竖版」或「统一横版」，"
                                "或者设一个留白毫米数。")
            return
        out_dir = self._tool_out_dir()
        dst = self._unique(out_dir, f"{Path(f).stem}-改版面", ".pdf")
        self._run_tool("改方向／留白",
                       lambda: cv.reshape(f, dst, opt), out_dir)

    def tool_double_layer(self) -> None:
        """做双层 PDF：底层图像 + 上层不可见文字，能看也能搜。

        选中的文件如果是 Word/WPS，文字直接来自源文档（不 OCR，100% 准）；
        扫描件/图片才走 OCR。
        """
        from ..core import ocr as ocrmod

        files = self._ask_many(
            "选择要做成双层 PDF 的文件（Word/WPS/PDF/图片 都可以）",
            APP_FILETYPES)
        if not files:
            return
        if len(files) > 1:
            messagebox.showinfo(
                "双层 PDF", "一次做一份（双层 PDF 是整本处理，混着做容易乱）。")
            files = files[:1]

        src = files[0]
        # 先告诉用户这一份会走哪条路——这决定了准确率
        plan = ocrmod.describe_plan(src)
        if not ocrmod.ocr_available() and "OCR" in plan:
            messagebox.showerror(
                "双层 PDF",
                "这一份需要 OCR 识别，但本机没装 OCR 引擎。\n\n"
                "装一下（在命令行里）：\n"
                "  pip install rapidocr_onnxruntime\n\n"
                "如果是 Word/WPS 文件就不需要 OCR，直接做。")
            return

        out_dir = self._tool_out_dir()
        dst = self._unique(out_dir, f"{Path(src).stem}-双层", ".pdf")
        self._run_tool(
            "双层 PDF",
            lambda: ocrmod.make_double_layer(
                src, dst, dpi=int(self.dpi.get() or 200)),
            out_dir)

    def _run_tool(self, name: str, fn, out_dir: str) -> None:
        if self.busy:
            messagebox.showinfo("格式转换", "正在处理别的任务，稍等。")
            return
        self._set_busy(True, f"{name} 处理中…")
        self.pb.configure(mode="indeterminate")
        self.pb.start(12)

        def job() -> None:
            try:
                made = fn()
                self.msgq.put(("tool_ok", name, made, out_dir))
            except Exception as exc:  # noqa: BLE001
                self.msgq.put(("tool_err", name, str(exc)))

        threading.Thread(target=job, daemon=True).start()

    # ------------------------------------------------------------ 消息泵
    def _poll(self) -> None:
        try:
            while True:
                msg = self.msgq.get_nowait()
                kind = msg[0]
                if kind == "engine":
                    self.engine.set(f"转换引擎：{msg[2]}" if msg[1]
                                    else "转换引擎：无 —— 装个 WPS 或 LibreOffice 才能转 Office 文档")
                elif kind == "prev":
                    _, tok, pages, n, err = msg
                    self._show_preview(tok, pages, n, err)
                elif kind == "progress":
                    self.pb.configure(value=msg[1])
                    self.status.set(f"正在转换第 {msg[1]} / {msg[2]} 个…")
                elif kind == "finish":
                    self.pb.stop()
                    self.pb.configure(mode="determinate", value=0)
                    msg[1]()
                elif kind == "tool_ok":
                    _, name, made, out_dir = msg
                    self.pb.stop()
                    self.pb.configure(mode="determinate", value=0)
                    self._set_busy(False, f"{name} 完成，生成 {len(made)} 个文件。")
                    messagebox.showinfo(name, f"完成，生成 {len(made)} 个文件。\n\n{out_dir}")
                elif kind == "tool_err":
                    _, name, err = msg
                    self.pb.stop()
                    self.pb.configure(mode="determinate", value=0)
                    self._set_busy(False, f"{name} 失败。")
                    messagebox.showerror(name, err[:800])
        except queue.Empty:
            pass
        finally:
            try:
                self.after(120, self._poll)
            except Exception:  # noqa: BLE001
                pass
