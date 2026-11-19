# -*- coding: utf-8 -*-
"""PDF 工具箱：加密、去密码、瘦身、按页提取、Word 导 PDF（带书签）。"""
from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from ..core import imgpress, pdf_engine, sysenv
from . import theme

PDF_TYPES = [("PDF", "*.pdf"), ("所有文件", "*.*")]
DOC_TYPES = [("Word / WPS", "*.docx *.doc *.wps *.rtf"), ("所有文件", "*.*")]


class PdfToolsPane(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=(0, 10, 0, 0), style="Canvas.TFrame")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        card = tk.Frame(self, background=theme.BORDER_SUBTLE)
        card.grid(row=0, column=0, sticky="nsew")
        inner = ttk.Frame(card, style="Card.TFrame", padding=16)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        ttk.Label(inner, text="PDF 工具箱", style="H2.TLabel").pack(anchor="w")
        ttk.Label(inner, text="都在本机处理，不上传。加密用 AES-256；提取页支持 1-3,5 这种写法。",
                  style="Mic.TLabel").pack(anchor="w", pady=(4, 12))

        rows = [
            ("加密 PDF…", "给 PDF 加打开密码", self.do_encrypt),
            ("去掉密码…", "知道原密码才能去", self.do_decrypt),
            ("PDF 瘦身…", "清冗余、压流，能小一点是一点", self.do_compress),
            ("按页提取…", "把指定的几页抽成一个新 PDF", self.do_extract),
            ("Word 导 PDF（带书签）…", "按标题生成一/二/三级书签，方便翻阅", self.do_bookmarked),
        ]
        for text, tip, cmd in rows:
            row = ttk.Frame(inner, style="Card.TFrame")
            row.pack(fill="x", pady=4)
            ttk.Button(row, text=text, style="Secondary.TButton", width=24,
                       command=cmd).pack(side="left")
            ttk.Label(row, text=tip, style="Mic.TLabel").pack(side="left", padx=(10, 0))

        self.status = tk.StringVar(value="")
        ttk.Label(inner, textvariable=self.status, style="Cap.TLabel",
                  wraplength=760, justify="left").pack(anchor="w", pady=(14, 0))

    # ------------------------------------------------------------ 工具
    def _ask_save(self, default: str, ext: str = ".pdf") -> str | None:
        return filedialog.asksaveasfilename(
            title="存到哪", defaultextension=ext, initialfile=default,
            initialdir=str(sysenv.desktop_dir()),
            filetypes=[("PDF", "*.pdf")] if ext == ".pdf" else [("所有文件", "*.*")])

    def _run(self, label: str, fn, btn=None):
        self.status.set(f"{label}…")
        self.update_idletasks()

        def job():
            try:
                msg = fn()
                self.after(0, lambda: self._ok(msg))
            except Exception as exc:  # noqa: BLE001
                err = str(exc)
                self.after(0, lambda: self._fail(label, err))

        threading.Thread(target=job, daemon=True).start()

    def _ok(self, msg):
        self.status.set(msg)
        messagebox.showinfo("PDF 工具箱", msg)

    def _fail(self, label, err):
        self.status.set(f"{label}没做成功：{err}")
        messagebox.showerror("PDF 工具箱", f"{label}没做成功：{err}")

    def do_encrypt(self):
        src = filedialog.askopenfilename(title="选要加密的 PDF", filetypes=PDF_TYPES)
        if not src:
            return
        pw = simpledialog.askstring("加密 PDF", "设置一个打开密码：", show="*")
        if not pw:
            return
        out = self._ask_save(f"{Path(src).stem}-已加密.pdf")
        if not out:
            return
        self._run("加密", lambda: (pdf_engine.encrypt(src, out, pw),
                                   f"加密好了：{Path(out).name}")[1])

    def do_decrypt(self):
        src = filedialog.askopenfilename(title="选要解密的 PDF", filetypes=PDF_TYPES)
        if not src:
            return
        pw = simpledialog.askstring("去掉密码", "这个 PDF 的密码：", show="*")
        if not pw:
            return
        out = self._ask_save(f"{Path(src).stem}-已解密.pdf")
        if not out:
            return
        self._run("解密", lambda: (pdf_engine.decrypt(src, out, pw),
                                   f"密码已去掉：{Path(out).name}")[1])

    def do_compress(self):
        fs = filedialog.askopenfilenames(title="选要瘦身的 PDF", filetypes=PDF_TYPES)
        if not fs:
            return
        out_dir = filedialog.askdirectory(title="瘦身后的 PDF 放哪", initialdir=str(sysenv.desktop_dir()))
        if not out_dir:
            return

        def job():
            msgs = []
            for f in fs:
                o = str(Path(out_dir) / f"{Path(f).stem}-瘦身.pdf")
                _o, before, after = pdf_engine.compress(f, o)
                msgs.append(f"{Path(f).name}：{imgpress.human_size(before)} → {imgpress.human_size(after)}")
            return "瘦身完成：\n" + "\n".join(msgs)

        self._run("瘦身", job)

    def do_extract(self):
        src = filedialog.askopenfilename(title="选 PDF", filetypes=PDF_TYPES)
        if not src:
            return
        spec = simpledialog.askstring("按页提取", "要哪几页？例如 1-3,5", initialvalue="1-3")
        if not spec:
            return
        out = self._ask_save(f"{Path(src).stem}-抽取.pdf")
        if not out:
            return
        self._run("提取页", lambda: (lambda r: f"抽好了：{Path(r[0]).name}（{r[1]} 页）")(
            pdf_engine.extract_pages(src, out, spec)))

    def do_bookmarked(self):
        fs = filedialog.askopenfilenames(title="选 Word / WPS 文档", filetypes=DOC_TYPES)
        if not fs:
            return
        out_dir = filedialog.askdirectory(title="导出的 PDF 放哪", initialdir=str(sysenv.desktop_dir()))
        if not out_dir:
            return

        def job():
            from ..core import office
            msgs = []
            for f in fs:
                o = str(Path(out_dir) / f"{Path(f).stem}.pdf")
                office.word_to_pdf_bookmarked(f, o)
                msgs.append(Path(o).name)
            return "导好了（带书签）：\n" + "\n".join(msgs)

        self._run("导出 PDF", job)
