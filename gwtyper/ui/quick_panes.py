# -*- coding: utf-8 -*-
"""两个独立小页：图片批量压缩、文件查重。"""
from __future__ import annotations

import csv
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from ..core import dupfind, imgpress, sysenv
from . import dnd, theme


def _card(master, row: int, column: int, padx=(0, 12), pady=(0, 0)) -> ttk.Frame:
    outer = tk.Frame(master, background=theme.BORDER_SUBTLE)
    outer.grid(row=row, column=column, sticky="nsew", padx=padx, pady=pady)
    card = ttk.Frame(outer, style="Card.TFrame", padding=14)
    card.pack(fill="both", expand=True, padx=1, pady=1)
    return card


class ImageCompressPane(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=(0, 10, 0, 0), style="Canvas.TFrame")
        self.files: list[Path] = []
        self.q: queue.Queue = queue.Queue()
        self.busy = False
        self.columnconfigure(0, weight=1, uniform="c")
        self.columnconfigure(1, weight=1, uniform="c")
        self.rowconfigure(0, weight=1)
        self.rowconfigure(1, weight=0)
        self._build_files()
        self._build_opts()
        self._poll()

    def _build_files(self):
        card = _card(self, 0, 0, padx=(0, 12))
        card.columnconfigure(0, weight=1)
        card.rowconfigure(2, weight=1)
        ttk.Label(card, text="要压缩的图片", style="H2.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(card, text="手机拍的扫描件动辄几 MB，压到 1 MB 以内还能看清字。"
                             "只压出副本，原图不动。", style="Mic.TLabel",
                  wraplength=420, justify="left").grid(row=1, column=0, sticky="w", pady=(4, 8))
        bar = ttk.Frame(card, style="Card.TFrame")
        bar.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(bar, text="添加图片…", style="Secondary.TButton",
                   command=self.add_files).pack(side="left")
        ttk.Button(bar, text="添加文件夹…", style="Ghost.TButton",
                   command=self.add_folder).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="清空", style="Ghost.TButton",
                   command=self.clear).pack(side="left", padx=(6, 0))
        wrap = ttk.Frame(card, style="Card.TFrame")
        wrap.grid(row=2, column=0, sticky="nsew")
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(wrap, columns=("name", "size"), show="headings", height=12)
        self.tree.heading("name", text="图片")
        self.tree.heading("size", text="大小")
        self.tree.column("name", width=280, anchor="w")
        self.tree.column("size", width=90, anchor="e")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        theme.stripe(self.tree)
        self.tree.bind("<Delete>", lambda e: self.clear())

    def _build_opts(self):
        card = _card(self, 0, 1, padx=(0, 0))
        card.columnconfigure(1, weight=1)
        ttk.Label(card, text="压缩设置", style="H2.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(card, text="怎么压", style="Cap.TLabel").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(8, 2))
        self.mode = tk.StringVar(value="side")
        mode_bar = ttk.Frame(card, style="Card.TFrame")
        mode_bar.grid(row=2, column=0, columnspan=2, sticky="w")
        ttk.Radiobutton(mode_bar, text="按长边限像素（画质优先）", value="side",
                        variable=self.mode, command=self._sync_mode).pack(side="left")
        ttk.Radiobutton(mode_bar, text="按目标体积（限大小）", value="kb",
                        variable=self.mode, command=self._sync_mode).pack(side="left", padx=(14, 0))

        ttk.Label(card, text="长边最大多少像素（越大越清楚、文件越大）", style="Cap.TLabel").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(10, 2))
        self.max_side = tk.IntVar(value=1600)
        self.sb_side = ttk.Spinbox(card, from_=800, to=4000, increment=100, width=8,
                                   textvariable=self.max_side)
        self.sb_side.grid(row=4, column=0, sticky="w")
        ttk.Label(card, text="图片质量（JPEG 60～95，75 一般够用）", style="Cap.TLabel").grid(
            row=5, column=0, columnspan=2, sticky="w", pady=(10, 2))
        self.quality = tk.IntVar(value=75)
        self.sb_quality = ttk.Spinbox(card, from_=60, to=95, increment=5, width=8,
                                      textvariable=self.quality)
        self.sb_quality.grid(row=6, column=0, sticky="w")

        ttk.Label(card, text="目标体积（KB，压到不超过这么大）", style="Cap.TLabel").grid(
            row=7, column=0, columnspan=2, sticky="w", pady=(10, 2))
        self.target_kb = tk.IntVar(value=300)
        self.sb_kb = ttk.Spinbox(card, from_=30, to=5000, increment=50, width=8,
                                 textvariable=self.target_kb)
        self.sb_kb.grid(row=8, column=0, sticky="w")
        self.kb_tip = ttk.Label(card, text="先降画质，画质到底还超标就自动缩尺寸；实在压不到会给出最接近的结果。",
                                style="Mic.TLabel", wraplength=400, justify="left")
        self.kb_tip.grid(row=9, column=0, columnspan=2, sticky="w", pady=(2, 0))

        ttk.Label(card, text="存到哪个文件夹（不填就放在原图旁边）", style="Cap.TLabel").grid(
            row=10, column=0, columnspan=2, sticky="w", pady=(10, 2))
        self.out_dir = tk.StringVar(value="")
        ttk.Entry(card, textvariable=self.out_dir).grid(row=11, column=0, sticky="ew")
        ttk.Button(card, text="浏览…", style="Ghost.TButton",
                   command=self.choose_out).grid(row=11, column=1, sticky="w", padx=(6, 0))
        act = ttk.Frame(card, style="Card.TFrame")
        act.grid(row=12, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        self.btn_go = ttk.Button(act, text="开始压缩", style="Primary.TButton", command=self.run)
        self.btn_go.pack(side="left")
        ttk.Button(act, text="打开输出目录", style="Ghost.TButton",
                   command=self.open_out).pack(side="left", padx=(6, 0))
        self.pb = ttk.Progressbar(card, mode="determinate")
        self.pb.grid(row=13, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        self.status = tk.StringVar(value="")
        ttk.Label(card, textvariable=self.status, style="Cap.TLabel",
                  wraplength=420, justify="left").grid(row=14, column=0, columnspan=2, sticky="w", pady=(8, 0))
        dnd.bind_targets([self.tree, card], self._on_drop)
        self._sync_mode()

    def _sync_mode(self):
        kb = self.mode.get() == "kb"
        for w in (self.sb_side, self.sb_quality):
            w.configure(state="disabled" if kb else "normal")
        self.sb_kb.configure(state="normal" if kb else "disabled")
        self.kb_tip.configure(foreground=theme.TEXT2 if kb else theme.TEXT3)

    # ------------------------------------------------------------ 文件
    def _on_drop(self, event):
        paths = dnd.paths_from_event(self, event, accept=dnd.suffix_filter(imgpress.IMG_EXTS))
        self._add(paths)

    def add_files(self):
        fs = filedialog.askopenfilenames(
            title="选择图片",
            filetypes=[("图片", "*.jpg *.jpeg *.png *.bmp *.webp *.tif *.tiff"), ("所有文件", "*.*")])
        self._add([Path(f) for f in fs])

    def add_folder(self):
        d = filedialog.askdirectory(title="选图片所在文件夹")
        if d:
            self._add([p for p in Path(d).rglob("*") if p.is_file() and imgpress.is_image(p)])

    def _add(self, paths):
        seen = {str(p) for p in self.files}
        for p in paths:
            p = Path(p)
            if imgpress.is_image(p) and str(p) not in seen:
                self.files.append(p)
                seen.add(str(p))
                self.tree.insert("", "end", values=(p.name, imgpress.human_size(p.stat().st_size)))
        getattr(self.tree, "_stripe", lambda: None)()

    def clear(self):
        self.files.clear()
        self.tree.delete(*self.tree.get_children())
        self.status.set("")

    def choose_out(self):
        d = filedialog.askdirectory(title="压缩后的图片放哪")
        if d:
            self.out_dir.set(d)

    def open_out(self):
        d = self.out_dir.get().strip() or str(sysenv.desktop_dir())
        try:
            sysenv.open_folder(d)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("打开目录", str(exc))

    # ------------------------------------------------------------ 跑
    def run(self):
        if self.busy:
            return
        if not self.files:
            messagebox.showinfo("图片压缩", "先添加图片。")
            return
        self.busy = True
        self.btn_go.state(["disabled"])
        self.pb.configure(maximum=max(1, len(self.files)), value=0)
        out_dir = self.out_dir.get().strip() or None
        ms, qy = int(self.max_side.get() or 1600), int(self.quality.get() or 75)
        tkb = int(self.target_kb.get() or 0) if self.mode.get() == "kb" else 0

        def job():
            def prog(i, n):
                self.q.put(("prog", i, n))
            res = imgpress.compress_many(self.files, out_dir, max_side=ms, quality=qy,
                                         target_kb=tkb, progress=prog)
            self.q.put(("done", res, out_dir))

        threading.Thread(target=job, daemon=True).start()

    def _poll(self):
        try:
            while True:
                msg = self.q.get_nowait()
                if msg[0] == "prog":
                    self.pb.configure(value=msg[1])
                elif msg[0] == "done":
                    self._done(msg[1], msg[2])
        except queue.Empty:
            pass
        finally:
            self.after(120, self._poll)

    def _done(self, res, out_dir):
        self.busy = False
        self.btn_go.state(["!disabled"])
        ok = [r for r in res if r["ok"]]
        bad = [r for r in res if not r["ok"]]
        before = sum(r["before"] for r in ok)
        after = sum(r["after"] for r in ok)
        msg = f"压好 {len(ok)} 张"
        if before:
            msg += f"，{imgpress.human_size(before)} → {imgpress.human_size(after)}"
            if before > 0:
                msg += f"（省 {100 - after * 100 // max(1, before)}%）"
        if bad:
            msg += f"；{len(bad)} 张没压成（{bad[0]['msg']}）"
        self.status.set(msg)
        if out_dir:
            self.out_dir.set(str(out_dir))


class DupPane(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=(0, 10, 0, 0), style="Canvas.TFrame")
        self.files: list[Path] = []
        self.groups: list[dict] = []
        self.q: queue.Queue = queue.Queue()
        self.busy = False
        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)
        top = _card(self, 0, 0)
        ttk.Label(top, text="文件查重", style="H2.TLabel").pack(anchor="w")
        ttk.Label(top, text="同一份材料改个名字又报一遍，这里能查出来。"
                             "字节完全一样，或者正文一模一样（重新另存过也算），都会归成一组。",
                  style="Mic.TLabel", wraplength=820, justify="left").pack(anchor="w", pady=(4, 8))
        bar = ttk.Frame(top, style="Card.TFrame")
        bar.pack(anchor="w")
        ttk.Button(bar, text="添加文件…", style="Secondary.TButton",
                   command=self.add_files).pack(side="left")
        ttk.Button(bar, text="添加文件夹…", style="Ghost.TButton",
                   command=self.add_folder).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="清空", style="Ghost.TButton", command=self.clear).pack(side="left", padx=(6, 0))
        self.btn_go = ttk.Button(bar, text="开始查重", style="Primary.TButton", command=self.run)
        self.btn_go.pack(side="left", padx=(18, 0))
        ttk.Button(bar, text="导出结果 CSV…", style="Ghost.TButton",
                   command=self.export).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="打开所在文件夹", style="Ghost.TButton",
                   command=self.open_sel).pack(side="left", padx=(6, 0))
        self.status = tk.StringVar(value="")
        ttk.Label(top, textvariable=self.status, style="Cap.TLabel",
                  wraplength=820, justify="left").pack(anchor="w", pady=(8, 0))

        card = _card(self, 1, 0, pady=(12, 0))
        card.columnconfigure(0, weight=1)
        card.rowconfigure(0, weight=1)
        wrap = ttk.Frame(card, style="Card.TFrame")
        wrap.grid(row=0, column=0, sticky="nsew")
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(wrap, columns=("size",), show="tree headings")
        self.tree.heading("#0", text="文件")
        self.tree.heading("size", text="大小")
        self.tree.column("#0", width=520, anchor="w")
        self.tree.column("size", width=110, anchor="e")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.tag_configure("group", background=theme.SUNKEN)
        self.tree.bind("<Double-1>", lambda e: self.open_sel())
        dnd.bind_targets([self.tree, card], self._on_drop)

    def _on_drop(self, event):
        self._add(dnd.paths_from_event(self, event))

    def add_files(self):
        fs = filedialog.askopenfilenames(title="选择文件", filetypes=[("所有文件", "*.*")])
        self._add([Path(f) for f in fs])

    def add_folder(self):
        d = filedialog.askdirectory(title="选文件夹")
        if d:
            self._add([p for p in Path(d).rglob("*") if p.is_file()])

    def _add(self, paths):
        seen = {str(p) for p in self.files}
        for p in paths:
            p = Path(p)
            if p.is_file() and str(p) not in seen:
                self.files.append(p)
                seen.add(str(p))
        self.status.set(f"已添加 {len(self.files)} 个文件，点「开始查重」。")

    def clear(self):
        self.files.clear()
        self.groups = []
        self.tree.delete(*self.tree.get_children())
        self.status.set("")

    def run(self):
        if self.busy:
            return
        if len(self.files) < 2:
            messagebox.showinfo("文件查重", "至少加两个文件才谈得上重复。")
            return
        self.busy = True
        self.btn_go.state(["disabled"])
        self.tree.delete(*self.tree.get_children())
        self.status.set("正在比对…")
        files = list(self.files)

        def job():
            gs = dupfind.find(files)
            self.q.put(("done", gs))

        threading.Thread(target=job, daemon=True).start()

    def _poll(self):
        try:
            while True:
                msg = self.q.get_nowait()
                if msg[0] == "done":
                    self._done(msg[1])
        except queue.Empty:
            pass
        finally:
            self.after(120, self._poll)

    def _done(self, groups):
        self.busy = False
        self.btn_go.state(["!disabled"])
        self.groups = groups
        for i, g in enumerate(groups, 1):
            parent = self.tree.insert("", "end", open=True,
                                      text=f"第 {i} 组（{g['why']}，{len(g['files'])} 个文件）",
                                      values=("",), tags=("group",))
            for f in g["files"]:
                p = Path(f)
                try:
                    size = imgpress.human_size(p.stat().st_size)
                except OSError:
                    size = "-"
                self.tree.insert(parent, "end", text=p.name, values=(size,))
        if groups:
            n = sum(len(g["files"]) for g in groups)
            self.status.set(f"查出 {len(groups)} 组重复，共 {n} 个文件。重复的留一份就够了。")
        else:
            self.status.set(f"比了 {len(self.files)} 个文件，没有重复。")

    def _selected_path(self):
        sel = self.tree.selection()
        if not sel:
            return None
        iid = sel[0]
        parent = self.tree.parent(iid)
        if not parent:
            return None
        idx = self.tree.index(parent)
        row = self.tree.index(iid)
        try:
            return self.groups[idx]["files"][row]
        except Exception:  # noqa: BLE001
            return None

    def open_sel(self):
        p = self._selected_path()
        if not p:
            messagebox.showinfo("文件查重", "先在下面点一个具体文件（不要点组标题）。")
            return
        sysenv.reveal(p)

    def export(self):
        if not self.groups:
            messagebox.showinfo("文件查重", "先查一遍。")
            return
        default = str(Path(sysenv.desktop_dir()) / "文件查重结果.csv")
        out = filedialog.asksaveasfilename(title="导出查重结果", defaultextension=".csv",
                                           initialfile="文件查重结果.csv",
                                           initialdir=str(Path(default).parent),
                                           filetypes=[("CSV", "*.csv")])
        if not out:
            return
        with open(out, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["组", "重复原因", "文件名", "大小", "路径"])
            for i, g in enumerate(self.groups, 1):
                for p in g["files"]:
                    try:
                        size = Path(p).stat().st_size
                    except OSError:
                        size = ""
                    w.writerow([i, g["why"], Path(p).name, size, p])
        self.status.set(f"已导出：{out}")
        try:
            sysenv.open_folder(str(Path(out).parent))
        except Exception:  # noqa: BLE001
            pass
