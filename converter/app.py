# -*- coding: utf-8 -*-
"""格式转换器主界面。"""
from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

from core import convert as cv
from core import formats, office_com, pdf_engine
from ui import theme

try:  # 拖拽是可选的，缺库也能用
    from tkinterdnd2 import DND_FILES, TkinterDnD

    _HAS_DND = True
except Exception:  # pragma: no cover
    DND_FILES = None
    TkinterDnD = None
    _HAS_DND = False

APP_NAME = "格式转换器"
APP_VER = "1.0"


def human_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024.0
    return f"{n:.1f} GB"


def open_path(path: str) -> None:
    try:
        os.startfile(path)  # type: ignore[attr-defined]
    except Exception:
        subprocess.Popen(["explorer", path])


class Base(tk.Tk if not _HAS_DND else TkinterDnD.Tk):  # type: ignore[misc]
    pass


class App(Base):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"{APP_NAME} {APP_VER}")
        self.geometry("1080x720")
        self.minsize(940, 640)
        self.fonts = theme.apply(self)
        self.files: list[str] = []
        self.busy = False
        self.msgq: queue.Queue = queue.Queue()
        self.out_dir = tk.StringVar(value=str(Path.home() / "Desktop" / "转换结果"))
        self.target = tk.StringVar(value="PDF")
        self.dpi = tk.IntVar(value=150)
        self.split_every = tk.IntVar(value=1)
        self.keep_origin = tk.BooleanVar(value=True)
        self.status = tk.StringVar(value="就绪")
        self.engine_text = tk.StringVar(value="正在检测转换引擎…")

        self._build()
        self._setup_dnd()
        self.after(80, self._poll)
        threading.Thread(target=self._detect_engine, daemon=True).start()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------ 布局
    def _build(self) -> None:
        head = ttk.Frame(self, padding=(20, 16, 20, 8))
        head.pack(fill="x")
        ttk.Label(head, text=APP_NAME, style="H1.TLabel").pack(side="left")
        ttk.Label(head, text=f"v{APP_VER}", style="Status.TLabel").pack(side="left", padx=(8, 0), pady=(4, 0))
        badge = ttk.Label(head, textvariable=self.engine_text, style="Status.TLabel")
        badge.pack(side="right", pady=(4, 0))

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=16, pady=(0, 6))
        self.tab_convert = ttk.Frame(self.nb, padding=12)
        self.tab_tools = ttk.Frame(self.nb, padding=12)
        self.tab_help = ttk.Frame(self.nb, padding=12)
        self.nb.add(self.tab_convert, text="  格式转换  ")
        self.nb.add(self.tab_tools, text="  PDF 工具箱  ")
        self.nb.add(self.tab_help, text="  使用说明  ")
        self.nb.select(self.tab_convert)

        self._build_convert()
        self._build_tools()
        self._build_help()

        bar = ttk.Frame(self, padding=(20, 4, 20, 12))
        bar.pack(fill="x")
        self.pb = ttk.Progressbar(bar, style="Brand.Horizontal.TProgressbar", mode="determinate")
        self.pb.pack(fill="x", pady=(0, 6))
        ttk.Label(bar, textvariable=self.status, style="Status.TLabel").pack(side="left")

    def _build_convert(self) -> None:
        root = self.tab_convert
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        left = ttk.Frame(root, style="Card.TFrame", padding=12)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(2, weight=1)

        top = ttk.Frame(left, style="Card.TFrame")
        top.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(top, text="待转换文件", style="H2.TLabel").pack(side="left")
        self.count_lbl = ttk.Label(top, text="0 个", style="Muted.TLabel")
        self.count_lbl.pack(side="left", padx=(8, 0))

        btns = ttk.Frame(left, style="Card.TFrame")
        btns.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(btns, text="添加文件", command=self.add_files).pack(side="left")
        ttk.Button(btns, text="添加文件夹", command=self.add_folder).pack(side="left", padx=6)
        ttk.Button(btns, text="移除选中", command=self.remove_selected).pack(side="left")
        ttk.Button(btns, text="清空", command=self.clear_files).pack(side="left", padx=6)

        wrap = ttk.Frame(left, style="Card.TFrame")
        wrap.grid(row=2, column=0, sticky="nsew")
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        cols = ("name", "kind", "size", "state")
        self.tree = ttk.Treeview(wrap, columns=cols, show="headings", selectmode="extended")
        for c, t, w, a in (
            ("name", "文件名", 380, "w"),
            ("kind", "类型", 80, "center"),
            ("size", "大小", 90, "e"),
            ("state", "状态", 110, "center"),
        ):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor=a, stretch=(c == "name"))
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.tag_configure("ok", foreground=theme.OK)
        self.tree.tag_configure("err", foreground=theme.ERR)
        self.tree.tag_configure("run", foreground=theme.BRAND)
        self.tree.bind("<Double-1>", self._open_selected)

        hint = "支持拖拽文件到这里" if _HAS_DND else "点「添加文件」选择，或把文件拖进窗口（需拖拽组件）"
        ttk.Label(left, text=hint, style="Muted.TLabel").grid(row=3, column=0, sticky="w", pady=(8, 0))

        # ---- 右侧设置
        right = ttk.Frame(root, style="Card.TFrame", padding=14)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)

        ttk.Label(right, text="转换为", style="H2.TLabel").grid(row=0, column=0, sticky="w")
        self.fmt_frame = ttk.Frame(right, style="Card.TFrame")
        self.fmt_frame.grid(row=1, column=0, sticky="ew", pady=(8, 14))
        self.fmt_btns: dict[str, ttk.Radiobutton] = {}
        for i, t in enumerate(formats.TARGET_ORDER):
            rb = ttk.Radiobutton(
                self.fmt_frame,
                text=t,
                value=t,
                variable=self.target,
                command=self._on_target,
            )
            rb.grid(row=i // 4, column=i % 4, sticky="w", padx=(0, 10), pady=3)
            self.fmt_btns[t] = rb

        self.fmt_note = ttk.Label(right, text="", style="Muted.TLabel", wraplength=300)
        self.fmt_note.grid(row=2, column=0, sticky="w", pady=(0, 12))

        ttk.Separator(right, orient="horizontal").grid(row=3, column=0, sticky="ew", pady=(0, 12))

        ttk.Label(right, text="输出位置", style="H2.TLabel").grid(row=4, column=0, sticky="w", pady=(0, 6))
        row = ttk.Frame(right, style="Card.TFrame")
        row.grid(row=5, column=0, sticky="ew")
        row.columnconfigure(0, weight=1)
        ttk.Entry(row, textvariable=self.out_dir).grid(row=0, column=0, sticky="ew")
        ttk.Button(row, text="浏览", command=self.pick_out_dir).grid(row=0, column=1, padx=(6, 0))

        opt = ttk.Frame(right, style="Card.TFrame")
        opt.grid(row=6, column=0, sticky="ew", pady=(10, 0))
        ttk.Checkbutton(opt, text="转换完成后打开输出文件夹", variable=self.keep_origin).pack(anchor="w")
        drow = ttk.Frame(opt, style="Card.TFrame")
        drow.pack(anchor="w", pady=(8, 0))
        ttk.Label(drow, text="图片清晰度", style="Muted.TLabel").pack(side="left")
        ttk.Spinbox(
            drow, from_=72, to=600, increment=50, width=6, textvariable=self.dpi, font=self.fonts["body"]
        ).pack(side="left", padx=(6, 4), ipady=3)
        ttk.Label(drow, text="DPI（打印建议 300）", style="Muted.TLabel").pack(side="left")

        right.rowconfigure(7, weight=1)

        self.go_btn = ttk.Button(right, text="开始转换", style="Accent.TButton", command=self.start)
        self.go_btn.grid(row=8, column=0, sticky="ew", pady=(16, 6))
        ttk.Button(right, text="打开输出文件夹", command=lambda: self._open_dir(self.out_dir.get())).grid(
            row=9, column=0, sticky="ew"
        )

        self._on_target()

    def _build_tools(self) -> None:
        root = self.tab_tools
        root.columnconfigure(0, weight=1, uniform="c")
        root.columnconfigure(1, weight=1, uniform="c")
        root.rowconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)

        self._tool_card(
            root, 0, 0, "合并 PDF",
            "把多个 PDF（或 Office 文档、图片）按顺序合成一个 PDF。",
            [("选择文件并合并", self.tool_merge)],
        )
        self._tool_card(
            root, 0, 1, "拆分 PDF",
            "按页数把一个 PDF 拆成多个文件，例如每 1 页一个。",
            [("选择 PDF 并拆分", self.tool_split)],
            extra=("每", self.split_every, "页拆一份"),
        )
        self._tool_card(
            root, 1, 0, "图片转 PDF",
            "多张图片合成一个 PDF，自动按图片比例排版。",
            [("选择图片", self.tool_images_to_pdf)],
        )
        self._tool_card(
            root, 1, 1, "PDF 转图片",
            "PDF 每页导出为 PNG 或 JPG，清晰度用右侧 DPI 控制。",
            [("导出为 PNG", lambda: self.tool_pdf_to_images("png")),
             ("导出为 JPG", lambda: self.tool_pdf_to_images("jpg"))],
        )

    def _tool_card(self, parent, r, c, title, desc, buttons, extra=None) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=16)
        card.grid(row=r, column=c, sticky="nsew", padx=(0, 12) if c == 0 else (0, 0), pady=(0, 12) if r == 0 else 0)
        card.columnconfigure(0, weight=1)
        ttk.Label(card, text=title, style="H2.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(card, text=desc, style="Muted.TLabel", wraplength=380).grid(row=1, column=0, sticky="w", pady=(6, 10))
        if extra:
            row = ttk.Frame(card, style="Card.TFrame")
            row.grid(row=2, column=0, sticky="w", pady=(0, 10))
            ttk.Label(row, text=extra[0], style="Muted.TLabel").pack(side="left")
            ttk.Spinbox(row, from_=1, to=500, width=5, textvariable=extra[1]).pack(side="left", padx=4)
            ttk.Label(row, text=extra[2], style="Muted.TLabel").pack(side="left")
        brow = ttk.Frame(card, style="Card.TFrame")
        brow.grid(row=3, column=0, sticky="w", pady=(4, 0))
        for text, cmd in buttons:
            ttk.Button(brow, text=text, command=cmd).pack(side="left", padx=(0, 8))

    def _build_help(self) -> None:
        root = self.tab_help
        txt = tk.Text(
            root,
            wrap="word",
            relief="flat",
            bg="#FFFFFF",
            fg=theme.INK,
            font=self.fonts["body"],
            padx=18,
            pady=16,
            spacing1=2,
            spacing3=6,
        )
        txt.pack(fill="both", expand=True)
        body = """【能做什么】
Word / Excel / PPT / PDF / 图片 之间互相转换，支持一次处理多个文件。
· 转 PDF：Word、Excel、PPT、图片、TXT 都能转成 PDF
· PDF 转 Word：还原段落、表格和大致版式
· PDF 转 Excel：自动识别表格，没有表格时按行提取文字
· PDF 转 PPT：每页转成一页幻灯片，适合演示和汇报
· 提取文字：任何文档导出为 TXT，方便检索复制
· 表格转 CSV：Excel 第一个工作表导出为 CSV
· 转图片：PDF、PPT 每页导出为 PNG / JPG

【PDF 工具箱】
合并、按页拆分、图片合成 PDF、PDF 导出图片。

【转换质量说明】
Office 文档互转和转 PDF 调用本机 WPS / Microsoft Office，版式最接近原文件；
PDF 反向转 Word / Excel 使用内置解析引擎，复杂排版可能需要微调。
图片清晰度：默认 150 DPI 适合屏幕阅读，打印建议 300 DPI。

【小技巧】
· 可以一次选中多个文件，批量转换
· 双击列表中的文件可以直接打开
· 输出文件名与源文件同名，重名会自动加 (1)(2)，不会覆盖已有文件
· 目标格式会根据选中的文件自动筛选，不支持的会置灰

【常见问题】
· 转换失败提示「请确认已安装 WPS 或 Office」：说明本机缺少办公软件，
  或该文件被密码保护、正在被其他程序独占打开。
· PDF 转 Word 后表格错位：源 PDF 本身是扫描件或分栏排版时较常见，
  可以先转成图片再核对，或改用「提取文字」得到干净文本。
· 转换过程中请不要关闭本窗口，也不要关闭弹出的办公软件。
"""
        txt.insert("1.0", body)
        txt.configure(state="disabled")

    # ------------------------------------------------------------ 文件管理
    def _setup_dnd(self) -> None:
        if not _HAS_DND:
            return
        try:
            self.drop_target_register(DND_FILES)  # type: ignore[attr-defined]
            self.dnd_bind("<<Drop>>", self._on_drop)  # type: ignore[attr-defined]
        except Exception:
            pass

    def _on_drop(self, event) -> None:  # pragma: no cover - 依赖桌面环境
        try:
            paths = self.tk.splitlist(event.data)
        except Exception:
            paths = [event.data]
        self._add_paths([p for p in paths if os.path.isfile(p)])

    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="选择要转换的文件",
            filetypes=[
                ("所有支持的文件", "*.doc *.docx *.docm *.wps *.rtf *.xls *.xlsx *.xlsm *.et *.csv "
                 "*.ppt *.pptx *.pps *.ppsx *.dps *.pdf *.txt *.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp"),
                ("Word 文档", "*.doc *.docx *.docm *.wps *.rtf"),
                ("Excel 表格", "*.xls *.xlsx *.xlsm *.et *.csv"),
                ("PPT 演示", "*.ppt *.pptx *.pps *.ppsx *.dps"),
                ("PDF 文件", "*.pdf"),
                ("图片", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp"),
                ("全部文件", "*.*"),
            ],
        )
        self._add_paths(list(paths))

    def add_folder(self) -> None:
        d = filedialog.askdirectory(title="选择文件夹（只添加其中的文件）")
        if not d:
            return
        exts = formats.OFFICE_EXT | formats.PDF_EXT | formats.IMAGE_EXT | {".txt"}
        found = [
            str(Path(d) / n)
            for n in sorted(os.listdir(d))
            if (Path(d) / n).is_file() and Path(n).suffix.lower() in exts
        ]
        if not found:
            messagebox.showinfo(APP_NAME, "这个文件夹里没有可转换的文件。")
            return
        self._add_paths(found)

    def _add_paths(self, paths: list[str]) -> None:
        added = 0
        for p in paths:
            p = os.path.abspath(p)
            if p in self.files:
                continue
            self.files.append(p)
            self.tree.insert(
                "",
                "end",
                iid=p,
                values=(os.path.basename(p), formats.kind_of(p), human_size(os.path.getsize(p)), "待转换"),
            )
            added += 1
        if added:
            self._refresh_state()
            self.status.set(f"已添加 {added} 个文件，共 {len(self.files)} 个")

    def remove_selected(self) -> None:
        for iid in self.tree.selection():
            if iid in self.files:
                self.files.remove(iid)
            self.tree.delete(iid)
        self._refresh_state()

    def clear_files(self) -> None:
        self.tree.delete(*self.tree.get_children())
        self.files.clear()
        self._refresh_state()

    def _open_selected(self, _event=None) -> None:
        sel = self.tree.selection()
        if sel:
            open_path(sel[0])

    def _refresh_state(self) -> None:
        self.count_lbl.configure(text=f"{len(self.files)} 个")
        ok = formats.targets_for(self.files)
        for t, rb in self.fmt_btns.items():
            rb.configure(state="normal" if t in ok else "disabled")
        if self.target.get() not in ok and ok:
            self.target.set(ok[0])
        self._on_target()

    def _on_target(self) -> None:
        t = self.target.get()
        info = formats.TARGETS.get(t)
        if info:
            self.fmt_note.configure(text=info["note"])
        is_img = t in ("PNG", "JPG")
        self.go_btn.configure(text="开始转换" if not self.busy else "转换中…")

    def pick_out_dir(self) -> None:
        d = filedialog.askdirectory(title="选择输出文件夹")
        if d:
            self.out_dir.set(d)

    # ------------------------------------------------------------ 引擎
    def _detect_engine(self) -> None:
        found = office_com.available()
        if not found:
            self.msgq.put(("engine", "未检测到办公软件：Office 互转不可用，PDF 功能正常"))
            return
        who = office_com.vendor() or "办公软件"
        names = {"word": "Word", "excel": "表格", "ppt": "演示"}
        cr = "/".join(names[k] for k in found)
        self.msgq.put(("engine", f"{who} 已就绪（{cr}）"))

    # ------------------------------------------------------------ 转换
    def _dpi(self) -> int:
        """DPI 允许手输，非数字时回落到 150。只在主线程调用。"""
        try:
            v = int(self.dpi.get())
        except (tk.TclError, ValueError):
            v = 150
        return max(72, min(600, v))

    def start(self) -> None:
        if self.busy:
            return
        if not self.files:
            messagebox.showinfo(APP_NAME, "请先添加要转换的文件。")
            return
        target = self.target.get()
        out_dir = self.out_dir.get().strip()
        if not out_dir:
            messagebox.showinfo(APP_NAME, "请选择输出文件夹。")
            return
        todo = [f for f in self.files if formats.supported(f, target)]
        if not todo:
            messagebox.showwarning(APP_NAME, f"选中的文件都不能转成 {target}。")
            return
        try:
            os.makedirs(out_dir, exist_ok=True)
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"无法创建输出文件夹：{exc}")
            return

        self.busy = True
        self.go_btn.configure(state="disabled", text="转换中…")
        self.pb.configure(maximum=len(todo), value=0)
        for f in self.files:
            self._set_state(f, "等待" if f in todo else "跳过")
        # Tk 变量只能在主线程读，这里先取值再交给后台线程
        opt = cv.Options(dpi=self._dpi())
        threading.Thread(target=self._worker, args=(todo, target, out_dir, opt), daemon=True).start()

    def _worker(self, todo: list[str], target: str, out_dir: str, opt: cv.Options) -> None:
        office_com.ensure_com()  # 后台线程用 COM 前必须初始化
        done = 0
        made_any: list[str] = []
        for f in todo:
            self.msgq.put(("state", (f, "转换中…")))
            try:
                made = cv.convert(f, target, out_dir, opt)
                made_any.extend(made)
                self.msgq.put(("state", (f, f"完成 · {len(made)} 个")))
                self.msgq.put(("tag", (f, "ok")))
            except Exception as exc:
                msg = str(exc) if isinstance(exc, cv.ConvertError) else f"转换失败：{exc}"
                self.msgq.put(("state", (f, "失败")))
                self.msgq.put(("tag", (f, "err")))
                self.msgq.put(("log", f"{os.path.basename(f)}：{msg}"))
            done += 1
            self.msgq.put(("progress", (done, len(todo))))
        self.msgq.put(("finished", (len(todo), out_dir, made_any)))

    def _set_state(self, iid: str, text: str) -> None:
        try:
            v = list(self.tree.item(iid, "values"))
            if len(v) >= 4:
                v[3] = text
                self.tree.item(iid, values=v)
        except tk.TclError:
            pass

    def _poll(self) -> None:
        try:
            while True:
                kind, payload = self.msgq.get_nowait()
                if kind == "state":
                    self._set_state(*payload)
                elif kind == "tag":
                    iid, tag = payload
                    try:
                        self.tree.item(iid, tags=(tag,))
                    except tk.TclError:
                        pass
                elif kind == "progress":
                    done, total = payload
                    self.pb.configure(value=done)
                    self.status.set(f"正在转换 {done} / {total}")
                elif kind == "engine":
                    self.engine_text.set(payload)
                elif kind == "log":
                    pass
                elif kind == "finished":
                    total, out_dir, made = payload
                    self.busy = False
                    self.go_btn.configure(state="normal", text="开始转换")
                    self.status.set(f"完成 {total} 个文件，生成 {len(made)} 个结果")
                    if self.keep_origin.get() and made:
                        self._open_dir(out_dir)
                elif kind == "tool_ok":
                    name, made, out_dir = payload
                    self._stop_busy()
                    self.status.set(f"{name} 完成，生成 {len(made)} 个文件")
                    if self.keep_origin.get():
                        self._open_dir(out_dir)
                elif kind == "tool_err":
                    name, msg = payload
                    self._stop_busy()
                    self.status.set(f"{name} 失败")
                    messagebox.showerror(APP_NAME, f"{name} 失败：\n{msg}")
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _stop_busy(self) -> None:
        self.busy = False
        try:
            self.pb.stop()
        except tk.TclError:
            pass
        self.pb.configure(mode="determinate", value=0)

    def _open_dir(self, path: str) -> None:
        if path and os.path.isdir(path):
            open_path(path)

    # ------------------------------------------------------------ 工具箱
    def _ask_pdf(self, multiple: bool, title: str) -> list[str]:
        if multiple:
            p = filedialog.askopenfilenames(title=title, filetypes=[("PDF 文件", "*.pdf")])
            return list(p)
        p = filedialog.askopenfilename(title=title, filetypes=[("PDF 文件", "*.pdf")])
        return [p] if p else []

    def _tool_out_dir(self) -> str:
        d = self.out_dir.get().strip() or str(Path.home() / "Desktop")
        try:
            os.makedirs(d, exist_ok=True)
        except OSError:
            d = str(Path.home() / "Desktop")
        return d

    def tool_merge(self) -> None:
        files = filedialog.askopenfilenames(
            title="选择要合并的文件（按选择顺序）",
            filetypes=[
                ("PDF 与 Office 文档", "*.pdf *.doc *.docx *.xls *.xlsx *.ppt *.pptx *.png *.jpg *.jpeg"),
                ("PDF 文件", "*.pdf"),
                ("全部文件", "*.*"),
            ],
        )
        if len(files) < 2:
            if files:
                messagebox.showinfo(APP_NAME, "至少选择两个文件才能合并。")
            return
        out_dir = self._tool_out_dir()
        dst = str(Path(out_dir) / "合并结果.pdf")
        n = 1
        while os.path.exists(dst):
            dst = str(Path(out_dir) / f"合并结果({n}).pdf")
            n += 1
        dpi = self._dpi()
        self._run_tool(
            "合并 PDF",
            lambda: cv.merge_to_pdf(list(files), dst, cv.Options(dpi=dpi)),
            out_dir,
        )

    def tool_split(self) -> None:
        files = self._ask_pdf(False, "选择要拆分的 PDF")
        if not files:
            return
        out_dir = self._tool_out_dir()
        try:
            every = max(1, int(self.split_every.get()))
        except (tk.TclError, ValueError):
            every = 1
        self._run_tool(
            "拆分 PDF",
            lambda: pdf_engine.split_pdf(files[0], out_dir, every),
            out_dir,
        )

    def tool_images_to_pdf(self) -> None:
        files = filedialog.askopenfilenames(
            title="选择图片（按选择顺序合成）",
            filetypes=[("图片", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp"), ("全部文件", "*.*")],
        )
        if not files:
            return
        out_dir = self._tool_out_dir()
        dst = str(Path(out_dir) / "图片合集.pdf")
        n = 1
        while os.path.exists(dst):
            dst = str(Path(out_dir) / f"图片合集({n}).pdf")
            n += 1
        self._run_tool("图片转 PDF", lambda: cv.images_to_one_pdf(list(files), dst, cv.Options()), out_dir)

    def tool_pdf_to_images(self, fmt: str) -> None:
        files = self._ask_pdf(False, "选择要导出图片的 PDF")
        if not files:
            return
        out_dir = self._tool_out_dir()
        dpi = self._dpi()
        self._run_tool(
            f"PDF 转 {fmt.upper()}",
            lambda: cv.to_images(files[0], out_dir, fmt, cv.Options(dpi=dpi)),
            out_dir,
        )

    def _run_tool(self, name: str, fn, out_dir: str) -> None:
        if self.busy:
            messagebox.showinfo(APP_NAME, "正在处理其他任务，请稍候。")
            return
        self.busy = True
        self.pb.configure(mode="indeterminate")
        self.pb.start(12)
        self.status.set(f"{name} 处理中…")

        def job() -> None:
            office_com.ensure_com()
            try:
                made = fn()
                self.msgq.put(("tool_ok", (name, made, out_dir)))
            except Exception as exc:
                self.msgq.put(("tool_err", (name, str(exc))))

        threading.Thread(target=job, daemon=True).start()

    # ------------------------------------------------------------ 关闭
    def _on_close(self) -> None:
        if self.busy:
            if not messagebox.askyesno(APP_NAME, "还有任务在转换，确定要退出吗？"):
                return
        self.destroy()


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
