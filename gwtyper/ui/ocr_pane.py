# -*- coding: utf-8 -*-
"""OCR 识别页：把图片 / 扫描件 / 图片型 PDF 认成文字。

三种成品，按「你要拿去干什么」选：
  - **纯文字（.txt）**      一段一行，适合检索、复制、喂给别的程序
  - **Word（按原版式）**     字体位置尽量还原，看起来还像原件，要人工核一遍
  - **Word（只保留文字）**   不还原位置，字都是正文样式，适合接着改
  - **双层 PDF**            图像 + 不可见文字，能看也能搜（归档、上报）

和「格式转换」页的分工：那页管格式互转，这页专门做「图 → 字」。
"""
from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from ..core import ocr as ocrmod
from . import pump
from ..core import ocr_text as octxt
from ..core import sysenv
from . import theme
from .check import CheckRow
from .preview_pane import PreviewPane

# 能识别的源：图片 + 文档（文档先转 PDF 再识别）
SOURCE_TYPES = [
    ("图片 / PDF / Office 文档",
     "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp *.gif *.pdf "
     "*.docx *.doc *.wps *.rtf *.xlsx *.xls *.pptx *.ppt"),
    ("图片", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp *.gif"),
    ("PDF", "*.pdf"),
    ("全部文件", "*.*"),
]

# (key, 显示名, 输出后缀, 说明)
MODES = [
    ("txt", "纯文字（.txt）", ".txt",
     "一段一行，适合检索和复制。不保留版式。"),
    ("docx_layout", "Word（按原版式）", ".docx",
     "按识别到的位置还原字号和对齐，看起来还像原件。图片和表格线还原不了。"),
    ("docx_text", "Word（只保留文字）", ".docx",
     "不还原位置，字统一是正文样式，适合认出来接着改。"),
    ("double", "双层 PDF", ".pdf",
     "先认出文字让你核对修改，再合成：下层原图、上层改好的文字，能看也能搜。适合归档、上报。"),
    ("xlsx", "表格（Excel）", ".xlsx",
     "把表里的字按行列位置放进 Excel 格子。合并单元格、跨行表头还原不了，识别完请核对。"),
]

# 这几种成品都要先给用户看/改字，再落盘
EDITABLE_MODES = ("txt", "docx_layout", "docx_text", "double")


def _xlsx_preview(path) -> str:
    """把生成的 Excel 读回来，做成纯文本预览（表格只读核对用）。"""
    try:
        from openpyxl import load_workbook

        wb = load_workbook(str(path), read_only=True, data_only=True)
        lines = []
        for ws in wb.worksheets:
            lines.append(f"===== {ws.title} =====")
            for row in ws.iter_rows(values_only=True):
                cells = ["" if c is None else str(c) for c in row]
                if any(c.strip() for c in cells):
                    lines.append("\t".join(cells).rstrip())
            lines.append("")
        wb.close()
        return "\n".join(lines).strip()
    except Exception as exc:  # noqa: BLE001
        return f"（读预览失败：{exc}）"


class ResultEditor(tk.Toplevel):
    """识别结果预览 / 编辑窗：先把认出来的字给你看，改完再保存。"""

    def __init__(self, master, name: str, text: str, ext_label: str,
                 on_save, on_cancel=None, index: int = 0, total: int = 1,
                 readonly: bool = False):
        super().__init__(master)
        self._on_save = on_save
        self._on_cancel = on_cancel
        self._readonly = readonly
        title = "识别结果（可以核对，不可改）" if readonly else "识别结果（可以直接改）"
        if total > 1:
            title += f" · 第 {index} / {total} 个"
        self.title(title)
        self.configure(background=theme.SURFACE)
        self.transient(master.winfo_toplevel())

        head = ttk.Frame(self, style="Card.TFrame", padding=(14, 12, 14, 6))
        head.pack(fill="x")
        ttk.Label(head, text=Path(name).name, style="H3.TLabel").pack(anchor="w")
        self._hint = ttk.Label(head, text="", style="Mic.TLabel")
        self._hint.pack(anchor="w", pady=(3, 0))

        body = ttk.Frame(self, style="Card.TFrame", padding=(14, 0, 14, 0))
        body.pack(fill="both", expand=True)
        self.txt = tk.Text(body, wrap="char", undo=True, height=22, width=64,
                           font=("Microsoft YaHei UI", 11), relief="flat",
                           background=theme.SUNKEN, foreground=theme.TEXT,
                           insertbackground=theme.TEXT, padx=10, pady=8,
                           selectbackground=theme.SEL_BG)
        sb = ttk.Scrollbar(body, orient="vertical", command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set)
        self.txt.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)
        self.txt.insert("1.0", text)
        if readonly:
            self.txt.configure(state="disabled")
        self.txt.bind("<<Modified>>", self._on_modified)
        self.txt.edit_modified(False)

        foot = ttk.Frame(self, style="Card.TFrame", padding=(14, 10, 14, 12))
        foot.pack(fill="x")
        ttk.Label(foot, text=f"将保存为 {ext_label}", style="Mic.TLabel").pack(side="left")
        if readonly:
            ttk.Button(foot, text="知道了", style="Primary.TButton",
                       command=self._cancel).pack(side="right")
        else:
            ttk.Button(foot, text="保存", style="Primary.TButton",
                       command=self._save).pack(side="right")
            ttk.Button(foot, text="跳过这个", style="Ghost.TButton",
                       command=self._cancel).pack(side="right", padx=(0, 8))
            ttk.Button(foot, text="全选", style="Ghost.TButton",
                       command=self._select_all).pack(side="right", padx=(0, 8))

        self._update_hint()
        try:
            from ..core import appicon

            appicon.apply_to(self)
        except Exception:  # noqa: BLE001
            pass
        self.geometry("760x620")
        self._center(master)
        self.grab_set()
        self.bind("<Escape>", lambda _e: self._cancel())
        self.bind("<Control-s>", lambda _e: (self._save(), "break")[1])
        self.txt.focus_set()

    def _center(self, master) -> None:
        try:
            self.update_idletasks()
            top = master.winfo_toplevel()
            w, h = 760, max(420, min(760, self.winfo_reqheight()))
            x = top.winfo_rootx() + (top.winfo_width() - w) // 2
            y = top.winfo_rooty() + max(40, (top.winfo_height() - h) // 3)
            self.geometry(f"{w}x{h}+{max(0, x)}+{max(0, y)}")
        except Exception:  # noqa: BLE001
            pass

    def _select_all(self) -> None:
        self.txt.tag_add("sel", "1.0", "end-1c")
        return "break"

    def _update_hint(self) -> None:
        t = self.txt.get("1.0", "end-1c")
        n = len("".join(t.split()))
        if self._readonly:
            self._hint.configure(text=f"{n} 字（不含空格）。只读预览，不能改。")
        else:
            self._hint.configure(
                text=f"{n} 字（不含空格）。改完点「保存」，Ctrl+S 也行；直接关掉等于跳过。")

    def _on_modified(self, _e=None) -> None:
        if self.txt.edit_modified():
            self.txt.edit_modified(False)
            self._update_hint()

    def _save(self) -> None:
        text = self.txt.get("1.0", "end-1c")
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()
        self._on_save(text)

    def _cancel(self) -> None:
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()
        if self._on_cancel:
            self._on_cancel()


class OcrPane(ttk.Frame):
    """OCR 选项卡。"""

    def __init__(self, master, default_dir: str = "", prefs: dict | None = None):
        super().__init__(master)
        self.files: list[str] = []
        self.busy = False
        self.msgq: "queue.Queue[tuple]" = queue.Queue()
        self.last_dir = default_dir or str(sysenv.desktop_dir())
        self._prev_tok = 0
        self._prev_job = None
        self._prefs = prefs if prefs is not None else {}

        _known = {m[0] for m in MODES}
        _mode = str(self._prefs.get("ocr_mode") or "txt")
        if _mode not in _known:
            _mode = "txt"
        _out = str(self._prefs.get("ocr_out") or "").strip()
        self.mode = tk.StringVar(value=_mode)
        self.out_dir = tk.StringVar(value=_out or str(Path(self.last_dir) / "识别结果"))
        self.dpi = tk.IntVar(value=200)
        self.editable = tk.BooleanVar(value=bool(self._prefs.get("ocr_edit", True)))
        self._edit_queue: list[tuple] = []
        self._edit_key = "txt"
        self._edit_label = ""
        self._edit_out = ""
        self._edit_done = 0
        self._edit_failed: list = []
        self.status = tk.StringVar(value="选一张图片或一份扫描件，挑个成品类型，点「开始识别」。")
        self.count = tk.StringVar(value="0 个文件")
        self.engine_text = tk.StringVar(value="")
        self.mode_tip = tk.StringVar(value="")

        self._build()
        # 恢复上次的成品类型后，把说明文字同步一遍
        self._sync_mode()
        self.after(120, self._poll)
        threading.Thread(target=self._detect_engine, daemon=True).start()

    # ------------------------------------------------------------ 布局
    def _build(self) -> None:
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
        ttk.Label(head, text="要识别的文件", style="H2.TLabel").pack(side="left")
        ttk.Label(head, textvariable=self.count, style="Mic.TLabel").pack(
            side="left", padx=(8, 0), pady=(3, 0))

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

        self.tree = ttk.Treeview(box, columns=("name", "kind"),
                                 show="headings", selectmode="extended")
        self.tree.heading("name", text="文件")
        self.tree.heading("kind", text="类型")
        self.tree.column("name", width=190, anchor="w")
        self.tree.column("kind", width=56, anchor="center")
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.bind("<Delete>", lambda e: self.remove_selected())
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.schedule_preview(220))

        self.ph = ttk.Label(box, text="把图片或扫描件拖到这里，\n或点「添加文件…」",
                            style="Mic.TLabel", justify="center", wraplength=170)
        self.ph.place(relx=0.5, rely=0.45, anchor="center")
        self._setup_dnd()

        ttk.Label(card, text="原件不动，识别结果另存。", style="Mic.TLabel",
                  wraplength=250).grid(row=3, column=0, sticky="w", pady=(6, 0))

    def _build_options(self) -> None:
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

        card = ttk.Frame(canvas, style="Card.TFrame", padding=12)
        self._opt_win = canvas.create_window((0, 0), window=card, anchor="nw")
        card.columnconfigure(0, weight=1)

        def _resize(_e=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            try:
                if card.winfo_reqheight() > canvas.winfo_height() + 2:
                    vsb.grid()
                else:
                    vsb.grid_remove()
                    canvas.yview_moveto(0)
            except Exception:  # noqa: BLE001
                pass

        card.bind("<Configure>", _resize)
        canvas.bind("<Configure>", lambda e: (
            canvas.itemconfigure(self._opt_win, width=e.width), _resize()))

        def _wheel(event):
            if card.winfo_reqheight() > canvas.winfo_height():
                canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

        canvas.bind("<Enter>", lambda _e: canvas.bind_all("<MouseWheel>", _wheel))
        canvas.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))

        ttk.Label(card, text="识别成什么", style="H2.TLabel").grid(
            row=0, column=0, sticky="w")
        opts = ttk.Frame(card, style="Card.TFrame")
        opts.grid(row=1, column=0, sticky="w", pady=(6, 0))
        for key, label, _ext, _tip in MODES:
            ttk.Radiobutton(opts, text=label, value=key, variable=self.mode,
                            style="Card.TRadiobutton",
                            command=self._sync_mode).pack(anchor="w")

        ttk.Label(card, textvariable=self.mode_tip, style="Mic.TLabel",
                  wraplength=280, justify="left").grid(
            row=2, column=0, sticky="w", pady=(6, 0))

        ttk.Label(card, text="保存到", style="H2.TLabel").grid(
            row=3, column=0, sticky="w", pady=(14, 0))
        row = ttk.Frame(card, style="Card.TFrame")
        row.grid(row=4, column=0, sticky="ew", pady=(6, 0))
        row.columnconfigure(0, weight=1)
        ttk.Entry(row, textvariable=self.out_dir).grid(
            row=0, column=0, sticky="ew", ipady=3)
        ttk.Button(row, text="浏览…", style="Secondary.TButton",
                   command=self.choose_out).grid(row=0, column=1, padx=(6, 0))

        ttk.Separator(card, orient="horizontal").grid(
            row=5, column=0, sticky="ew", pady=(14, 10))

        ttk.Label(card, text="选项", style="H3.TLabel").grid(row=6, column=0, sticky="w")
        opt2 = ttk.Frame(card, style="Card.TFrame")
        opt2.grid(row=7, column=0, sticky="w", pady=(6, 0))
        ttk.Label(opt2, text="识别清晰度", style="Cap.TLabel").pack(side="left")
        ttk.Spinbox(opt2, from_=100, to=400, increment=50, width=6,
                    textvariable=self.dpi).pack(side="left", padx=(6, 4))
        ttk.Label(opt2, text="DPI（越高越准，也越慢）", style="Mic.TLabel").pack(side="left")

        act = ttk.Frame(card, style="Card.TFrame")
        act.grid(row=8, column=0, sticky="ew", pady=(16, 0))
        self.btn_go = ttk.Button(act, text="开始识别", style="Primary.TButton",
                                 command=self.run_ocr)
        self.btn_go.pack(side="left")
        ttk.Button(act, text="打开输出目录", style="Ghost.TButton",
                   command=self.open_out).pack(side="left", padx=(6, 0))

        self.pb = ttk.Progressbar(card, mode="determinate")
        self.pb.grid(row=9, column=0, sticky="ew", pady=(12, 4))
        ttk.Label(card, textvariable=self.status, style="Cap.TLabel",
                  wraplength=280, justify="left").grid(row=10, column=0, sticky="w")
        ttk.Label(card, textvariable=self.engine_text, style="Mic.TLabel",
                  wraplength=280, justify="left").grid(
            row=11, column=0, sticky="w", pady=(6, 0))

        self._sync_mode()
        self._refresh_count()

    def _build_preview(self) -> None:
        holder = tk.Frame(self, background=theme.BORDER_SUBTLE)
        holder.grid(row=0, column=2, sticky="nsew", padx=(12, 0))
        card = ttk.Frame(holder, style="Card.TFrame", padding=12)
        card.pack(fill="both", expand=True, padx=1, pady=1)
        card.columnconfigure(0, weight=1)
        card.rowconfigure(1, weight=1)

        head = ttk.Frame(card, style="Card.TFrame")
        head.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(head, text="原图", style="H2.TLabel").pack(side="left")
        self.prev_note = tk.StringVar(value="")
        ttk.Label(head, textvariable=self.prev_note, style="Mic.TLabel").pack(
            side="left", padx=(8, 0), pady=(3, 0))
        ttk.Button(head, text="刷新", style="Ghost.TButton", width=4,
                   command=lambda: self.refresh_preview(force=True)).pack(side="right")

        self.preview = PreviewPane(card, width_hint=200)
        self.preview.grid(row=1, column=0, sticky="nsew")
        self.preview.hide_head()
        self.preview.clear("选一个文件，这里显示它的样子")

    # ------------------------------------------------------------ 文件
    def add_files(self) -> None:
        got = filedialog.askopenfilenames(title="选择要识别的文件",
                                          filetypes=SOURCE_TYPES)
        if got:
            self._add_paths([Path(p) for p in got])

    def add_folder(self) -> None:
        d = filedialog.askdirectory(title="选择文件夹（只取里面一层）")
        if not d:
            return
        exts = (ocrmod.IMAGE_SOURCES | ocrmod.TEXT_SOURCES | {".pdf"})
        got = [p for p in sorted(Path(d).iterdir())
               if p.is_file() and p.suffix.lower() in exts]
        if not got:
            messagebox.showinfo("OCR 识别", "这个文件夹里没有能识别的文件。")
            return
        self._add_paths(got)

    def _add_paths(self, paths: list[Path]) -> None:
        first = ""
        for p in paths:
            sp = str(p)
            if sp in self.files:
                continue
            self.files.append(sp)
            self.tree.insert("", "end", iid=sp,
                             values=(p.name, p.suffix.lower().lstrip(".")))
            if not first:
                first = sp
        self._refresh_count()
        if first and not self.tree.selection():
            self.tree.selection_set(first)
            self.tree.see(first)

    def remove_selected(self) -> None:
        for iid in self.tree.selection():
            if iid in self.files:
                self.files.remove(iid)
            self.tree.delete(iid)
        self._refresh_count()
        self.schedule_preview(120)

    def clear(self) -> None:
        self.tree.delete(*self.tree.get_children())
        self.files.clear()
        self._refresh_count()
        self._prev_tok += 1
        self.preview.clear("选一个文件，这里显示它的样子")
        self.prev_note.set("")

    def _refresh_count(self) -> None:
        n = len(self.files)
        self.count.set(f"{n} 个文件")
        if n:
            self.ph.place_forget()
        else:
            self.ph.place(relx=0.5, rely=0.45, anchor="center")
        self.btn_go.configure(state="normal" if n and not self.busy else "disabled")

    # ------------------------------------------------------------ 选项
    def _mode_info(self):
        """当前成品类型 -> (key, 显示名, 输出后缀, 说明)。

        一定要把 key 一起返回：run_ocr 靠它分派到四个不同的识别函数，
        之前这里只返回三元组、run_ocr 却按四元组解包，一点「开始识别」
        就抛 ValueError，被 Tk 的默认异常处理吞掉（窗口版连 stderr 都看不到），
        表现成「按钮点了没反应、OCR 不可用」。
        """
        for key, label, ext, tip in MODES:
            if key == self.mode.get():
                return key, label, ext, tip
        return MODES[0]

    def _sync_mode(self) -> None:
        _key, _label, _ext, tip = self._mode_info()
        if _key not in EDITABLE_MODES:
            tip = tip + "\n（这一种不能改字：表格是按位置填格子，改字会串行。）"
        self.mode_tip.set(tip)

    def choose_out(self) -> None:
        d = filedialog.askdirectory(title="识别结果保存到哪",
                                    initialdir=self.out_dir.get() or self.last_dir)
        if d:
            self.out_dir.set(d)

    def open_out(self) -> None:
        d = self.out_dir.get().strip()
        if not d:
            messagebox.showinfo("OCR 识别", "还没设输出目录。")
            return
        Path(d).mkdir(parents=True, exist_ok=True)
        sysenv.open_folder(d)

    # ------------------------------------------------------------ 预览
    def schedule_preview(self, delay: int = 700) -> None:
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
            self.preview.clear("选一个文件，这里显示它的样子")
            self.prev_note.set("")
            return
        src = sel[0]
        self._prev_tok += 1
        tok = self._prev_tok
        width = max(220, self.preview.canvas.winfo_width() - 30)

        def job():
            try:
                from ..core import preview as prevmod

                pdf = prevmod.any_pdf(src, force=force)
                pages, n = prevmod.render(pdf, max_width=width, max_pages=3)
                self.msgq.put(("prev", tok, pages, n, None))
            except Exception as exc:  # noqa: BLE001
                self.msgq.put(("prev", tok, [], 0, str(exc)))

        threading.Thread(target=job, daemon=True).start()

    # ------------------------------------------------------------ 引擎
    def _detect_engine(self) -> None:
        """在后台线程探测引擎。

        注意：**不要在后台线程里读 Tk 变量**（self.mode.get() 之类会报
        'main thread is not in main loop'，之前就是这样被 except 吞掉，
        界面显示成「检测失败」，其实引擎是好的）。
        这里只碰纯 Python 的东西，界面文案交给 _poll 在主线程拼。
        """
        try:
            ok = ocrmod.ocr_available()
            if ok:
                try:
                    ocrmod.engine()      # 真正加载一次，确认模型能用
                    self.msgq.put(("engine", True, "OCR 引擎：RapidOCR（随程序打包）"))
                except Exception as exc:  # noqa: BLE001
                    self.msgq.put(("engine", False,
                                   f"OCR 引擎装上了但加载失败：{exc}"))
            else:
                self.msgq.put(("engine", False,
                               "OCR 引擎未安装，识别不可用。\n"
                               "装一下：pip install rapidocr_onnxruntime"))
        except Exception as exc:  # noqa: BLE001
            self.msgq.put(("engine", False, f"OCR 引擎检测异常：{exc}"))

    # ------------------------------------------------------------ 识别
    def run_ocr(self) -> None:
        if self.busy:
            return
        if not self.files:
            messagebox.showinfo("OCR 识别", "先添加要识别的文件。")
            return
        out_dir = self.out_dir.get().strip()
        if not out_dir:
            messagebox.showinfo("OCR 识别", "先选一个输出目录。")
            return
        if not ocrmod.ocr_available():
            messagebox.showerror(
                "OCR 识别",
                "本机没有 OCR 引擎，识别不了。\n\n"
                "装一下（命令行）：\n  pip install rapidocr_onnxruntime")
            return

        sel = [s for s in self.tree.selection() if s in self.files]
        todo = sel if sel else list(self.files)
        key, label, ext, _tip = self._mode_info()
        try:
            Path(out_dir).mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            messagebox.showerror("OCR 识别", f"输出目录建不了：{exc}")
            return

        # 每种模式的输出后缀不同，同名文件加序号避免覆盖
        jobs = []
        for src in todo:
            base = Path(src).stem
            dst = Path(out_dir) / f"{base}-识别{ext}"
            n = 1
            while dst.exists():
                dst = Path(out_dir) / f"{base}-识别({n}){ext}"
                n += 1
            jobs.append((src, dst))

        dpi = max(100, min(400, int(self.dpi.get() or 200)))
        # 默认一律先给预览：纯文字/Word/双层 PDF 可改字后落盘，表格只读核对
        editable = key in EDITABLE_MODES
        self._edit_key = key
        self._edit_label = label
        self._edit_out = out_dir
        self._edit_done = 0
        self._edit_failed = []
        self._edit_total = len(jobs)
        self._edit_queue = []
        self._set_busy(True, f"正在识别 {len(jobs)} 个文件（{label}）…"
                            + ("　识别完先给你看" if editable else ""))
        self.pb.configure(mode="determinate", maximum=len(jobs), value=0)

        def job():
            if editable:
                for i, (src, dst) in enumerate(jobs, start=1):
                    try:
                        if key == "double":
                            blk = octxt.recognize_blocks(src, dpi=dpi)
                            self._edit_queue.append(
                                {"src": src, "dst": dst, "text": blk["text"],
                                 "blocks": blk})
                        else:
                            res = octxt.recognize(src, dpi=dpi)
                            self._edit_queue.append(
                                {"src": src, "dst": dst, "text": res["text"],
                                 "blocks": None})
                    except Exception as exc:  # noqa: BLE001
                        self._edit_failed.append(
                            (Path(src).name, str(exc).splitlines()[0]))
                    self.msgq.put(("progress", i, len(jobs)))

                def finish_edit():
                    self._set_busy(False, "")
                    self._edit_next()

                self.msgq.put(("finish", finish_edit))
                return

            done, failed = 0, []
            for i, (src, dst) in enumerate(jobs, start=1):
                try:
                    if key == "txt":
                        octxt.to_text(src, dst, dpi=dpi)
                    elif key == "docx_layout":
                        octxt.to_docx(src, dst, dpi=dpi)
                    elif key == "docx_text":
                        octxt.to_text_docx(src, dst, dpi=dpi)
                    elif key == "xlsx":
                        octxt.to_xlsx_table(src, dst, dpi=dpi)
                        self.msgq.put(("preview", Path(src).name,
                                       _xlsx_preview(dst), str(dst)))
                    else:
                        ocrmod.make_double_layer(src, dst, dpi=dpi)
                    done += 1
                except Exception as exc:  # noqa: BLE001
                    failed.append((Path(src).name, str(exc).splitlines()[0]))
                self.msgq.put(("progress", i, len(jobs)))

            def finish():
                self._set_busy(False, "")
                self._after(done, failed, out_dir, label)

            self.msgq.put(("finish", finish))

        threading.Thread(target=job, daemon=True).start()

    def _edit_next(self) -> None:
        """一个接一个弹「识别结果」窗；全部处理完再汇总。"""
        if not self._edit_queue:
            self._after(self._edit_done, self._edit_failed,
                        self._edit_out, self._edit_label)
            return
        item = self._edit_queue.pop(0)
        idx = self._edit_total - len(self._edit_queue)
        src, dst = item["src"], item["dst"]
        self.status.set(f"识别完成，等你看 / 改：{Path(src).name}")

        def save(text: str) -> None:
            try:
                Path(dst).parent.mkdir(parents=True, exist_ok=True)
                if self._edit_key == "txt":
                    Path(dst).write_text(str(text).rstrip() + "\n",
                                         encoding="utf-8-sig")
                elif self._edit_key == "double":
                    pages = octxt.split_edited(text)
                    octxt.to_pdf_double_from_blocks(
                        src, dst, item.get("blocks") or {}, pages)
                else:
                    octxt.text_to_docx(text, dst)
                self._edit_done += 1
            except Exception as exc:  # noqa: BLE001
                self._edit_failed.append((Path(src).name, str(exc).splitlines()[0]))
            self._edit_next()

        def cancel() -> None:
            self._edit_failed.append((Path(src).name, "跳过（没有保存）"))
            self._edit_next()

        _k, _l, ext, _t = self._mode_info()
        if self._edit_key == "double":
            label_ext = ".pdf（下层原图 + 上层你改过的文字）"
        else:
            label_ext = ext + ("（正文样式）" if self._edit_key != "txt" else "")
        ResultEditor(self, Path(src).name, item["text"], label_ext,
                     save, cancel, index=idx, total=self._edit_total)

    def _after(self, done: int, failed: list, out_dir: str, label: str) -> None:
        if failed:
            lines = "\n".join(f"· {n}：{m}" for n, m in failed[:5])
            if len(failed) > 5:
                lines += f"\n… 还有 {len(failed) - 5} 个"
            self.status.set(f"完成 {done} 个，失败 {len(failed)} 个。")
            messagebox.showwarning(
                "OCR 识别",
                f"完成 {done} 个，失败 {len(failed)} 个：\n\n{lines}")
        elif done:
            self.status.set(f"完成，生成 {done} 个文件（{label}）。")
            messagebox.showinfo(
                "OCR 识别",
                f"完成，生成 {done} 个文件（{label}）。\n\n{out_dir}")
        else:
            self.status.set("没有生成文件。")

    def _set_busy(self, on: bool, text: str) -> None:
        self.busy = on
        if text:
            self.status.set(text)
        self.btn_go.configure(state="disabled" if on else
                              ("normal" if self.files else "disabled"))

    # ------------------------------------------------------------ 拖拽
    def _setup_dnd(self) -> None:
        from . import dnd

        dnd.bind_targets([self.tree, self, getattr(self, "ph", None)], self._on_drop)

    def _on_drop(self, event) -> None:
        from . import dnd

        exts = ocrmod.IMAGE_SOURCES | ocrmod.TEXT_SOURCES | {".pdf"}
        paths = dnd.paths_from_event(self, event, accept=dnd.suffix_filter(exts))
        if paths:
            self._add_paths(paths)

    # ------------------------------------------------------------ 消息泵
    def _handle_msg(self, msg: tuple) -> None:
        kind = msg[0]
        if kind == "engine":
            self.engine_text.set(msg[2])
        elif kind == "prev":
            _, tok, pages, n, err = msg
            if tok != self._prev_tok:
                return
            if err:
                self.prev_note.set("")
                self.preview.clear(err)
            else:
                note = (f"{n} 页（预览前 {len(pages)} 页）"
                        if len(pages) < n else f"{n} 页")
                self.prev_note.set(note)
                self.preview.set_pages(pages, note)
        elif kind == "progress":
            self.pb.configure(value=msg[1])
            self.status.set(f"正在识别第 {msg[1]} / {msg[2]} 个…")
        elif kind == "preview":
            _, name, text, out_path = msg
            self.status.set(f"生成完毕，核对一下：{name}")
            ResultEditor(self, name, text or "（没有识别到文字）",
                         "已生成，下面是核对用的只读预览", None, None,
                         readonly=True)
        elif kind == "finish":
            self.pb.stop()
            self.pb.configure(mode="determinate", value=0)
            msg[1]()

    def _poll(self) -> None:
        # 逐条处理：一条坏消息不会把消息泵打死（否则表现为「识别完了界面没反应」）
        pump.drain(self.msgq, self._handle_msg)
        self.after(120, self._poll)
