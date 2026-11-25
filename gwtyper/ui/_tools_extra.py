# -*- coding: utf-8 -*-
"""给 tools_pane 追加：扫描件整理、表格生成。"""
from __future__ import annotations

import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from ..core import scanim, sysenv, villages
from ..core.formdoc import write_table_docx as _write_table_docx

IMG_EXT = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}


class ToolsExtraMixin:
    """扫描件纠偏去黑边 + 会议签到表/材料交接单生成。"""

    # -------------------------------------------------------- 扫描件整理
    def _build_deskew(self) -> None:
        card = self._card(1, 0)
        card.columnconfigure(0, weight=1)
        card.rowconfigure(2, weight=1)
        ttk.Label(card, text="扫描件摆正 / 去黑边", style="H2.TLabel").grid(
            row=0, column=0, sticky="w")
        ttk.Label(card, text="手机拍的、扫描歪了带黑边的，自动摆正并把四周黑边裁掉。"
                             "只做这两件事，不调颜色，字不会发虚。",
                  style="Mic.TLabel", wraplength=460, justify="left").grid(
            row=1, column=0, sticky="w", pady=(4, 6))

        body = ttk.Frame(card, style="Card.TFrame")
        body.grid(row=2, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)
        bar = ttk.Frame(body, style="Card.TFrame")
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Button(bar, text="添加图片…", style="Secondary.TButton",
                   command=self.scan_add).pack(side="left")
        ttk.Button(bar, text="添加文件夹…", style="Ghost.TButton",
                   command=self.scan_add_folder).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="清空", style="Ghost.TButton",
                   command=self.scan_clear).pack(side="left", padx=(6, 0))

        wrap = ttk.Frame(body, style="Card.TFrame")
        wrap.grid(row=1, column=0, sticky="nsew")
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        self.scan_tree = ttk.Treeview(wrap, columns=("name",), show="headings",
                                      height=5)
        self.scan_tree.heading("name", text="图片")
        self.scan_tree.column("name", width=300, anchor="w")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.scan_tree.yview)
        self.scan_tree.configure(yscrollcommand=sb.set)
        self.scan_tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        opt = ttk.Frame(card, style="Card.TFrame")
        opt.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        self.scan_deskew = tk.BooleanVar(value=True)
        self.scan_trim = tk.BooleanVar(value=True)
        from .check import CheckRow

        CheckRow(opt, "自动摆正", variable=self.scan_deskew).pack(side="left")
        CheckRow(opt, "裁掉黑边", variable=self.scan_trim).pack(side="left", padx=(12, 0))

        act = ttk.Frame(card, style="Card.TFrame")
        act.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(act, text="开始处理", style="Primary.TButton",
                   command=self.scan_run).pack(side="left")
        self.scan_status = tk.StringVar(value="处理完存成「原名-摆正」，原图不动。")
        ttk.Label(card, textvariable=self.scan_status, style="Cap.TLabel",
                  wraplength=460, justify="left").grid(row=5, column=0, sticky="w",
                                                       pady=(6, 0))
        from . import dnd

        dnd.bind_targets([self.scan_tree, wrap], self._on_drop_scan)

    def _on_drop_scan(self, event) -> None:
        from . import dnd

        paths = dnd.paths_from_event(self, event, accept=dnd.suffix_filter(IMG_EXT))
        if paths:
            self._add_scan([Path(p) for p in paths])

    def scan_add(self) -> None:
        fs = filedialog.askopenfilenames(
            title="选图片",
            filetypes=[("图片", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp"),
                       ("全部", "*.*")])
        if fs:
            self._add_scan([Path(f) for f in fs])

    def scan_add_folder(self) -> None:
        d = filedialog.askdirectory(title="选文件夹")
        if d:
            self._add_scan([p for p in sorted(Path(d).iterdir())
                            if p.is_file() and p.suffix.lower() in IMG_EXT])

    def _add_scan(self, paths: list[Path]) -> None:
        for p in paths:
            if p not in self.scan_files:
                self.scan_files.append(p)
                self.scan_tree.insert("", "end", values=(p.name,))
        self.scan_status.set(f"已添加 {len(self.scan_files)} 张图片。")

    def scan_clear(self) -> None:
        self.scan_files.clear()
        self.scan_tree.delete(*self.scan_tree.get_children())
        self.scan_status.set("已清空。")

    def scan_run(self) -> None:
        if not self.scan_files:
            messagebox.showinfo("扫描件整理", "先添加图片。")
            return
        ok, failed, angles = 0, [], []
        for p in self.scan_files:
            try:
                out = p.with_name(f"{p.stem}-摆正{p.suffix}")
                r = scanim.fix(p, out, deskew=bool(self.scan_deskew.get()),
                               trim=bool(self.scan_trim.get()))
                ok += 1
                if abs(r["angle"]) >= 0.3:
                    angles.append(f"{p.name} {r['angle']}°")
            except Exception as exc:  # noqa: BLE001
                failed.append(f"{p.name}：{exc}")
        msg = f"处理完 {ok} 张，存在各自文件夹里（原名-摆正）。"
        if angles:
            msg += f"　摆正了：{'、'.join(angles[:3])}"
        if failed:
            msg += f"　失败 {len(failed)} 张（{failed[0]}）"
        self.scan_status.set(msg)
        (messagebox.showwarning if failed else messagebox.showinfo)("扫描件整理", msg)

    # -------------------------------------------------------- 表格生成
    def _build_form(self) -> None:
        card = self._card(1, 1)
        card.columnconfigure(0, weight=1)
        ttk.Label(card, text="签到表 / 材料交接单", style="H2.TLabel").grid(
            row=0, column=0, sticky="w")
        ttk.Label(card, text="填个标题和名单，直接生成能打印的 Word 表格。"
                             "名单默认用「材料上交」页的村名单。",
                  style="Mic.TLabel", wraplength=460, justify="left").grid(
            row=1, column=0, sticky="w", pady=(4, 6))

        frm = ttk.Frame(card, style="Card.TFrame")
        frm.grid(row=2, column=0, sticky="nsew")
        frm.columnconfigure(1, weight=1)
        self.form_kind = tk.StringVar(value="sign")
        r = ttk.Frame(frm, style="Card.TFrame")
        r.grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Radiobutton(r, text="会议签到表", value="sign", variable=self.form_kind,
                        style="Card.TRadiobutton").pack(side="left")
        ttk.Radiobutton(r, text="材料交接单", value="hand", variable=self.form_kind,
                        style="Card.TRadiobutton").pack(side="left", padx=(12, 0))

        self.form_title = tk.StringVar(value="")
        self.form_sub = tk.StringVar(value="")
        self.form_day = tk.StringVar(value=f"{date.today():%Y年%m月%d日}")
        for i, (label, var) in enumerate(
                (("标题", self.form_title), ("单位/事由", self.form_sub),
                 ("日期", self.form_day)), start=1):
            ttk.Label(frm, text=label, style="Cap.TLabel").grid(
                row=i, column=0, sticky="w", pady=(6, 0))
            ttk.Entry(frm, textvariable=var).grid(row=i, column=1, sticky="ew",
                                                  padx=(6, 0), pady=(6, 0))
        ttk.Label(frm, text="名单", style="Cap.TLabel").grid(
            row=4, column=0, sticky="nw", pady=(6, 0))
        self.form_names = tk.Text(frm, height=6, wrap="none", relief="flat",
                                  highlightthickness=1, highlightbackground="#D3D9E2",
                                  font=("Microsoft YaHei UI", 9), padx=6, pady=4)
        self.form_names.grid(row=4, column=1, sticky="nsew", padx=(6, 0), pady=(6, 0))
        frm.rowconfigure(4, weight=1)
        names = villages.load_names()
        if names:
            self.form_names.insert("1.0", "\n".join(names))

        act = ttk.Frame(card, style="Card.TFrame")
        act.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(act, text="生成 Word…", style="Primary.TButton",
                   command=self.form_make).pack(side="left")
        ttk.Button(act, text="用「材料上交」的名单", style="Ghost.TButton",
                   command=self.form_use_names).pack(side="left", padx=(6, 0))
        self.form_status = tk.StringVar(value="")
        ttk.Label(card, textvariable=self.form_status, style="Cap.TLabel",
                  wraplength=460, justify="left").grid(row=4, column=0, sticky="w",
                                                       pady=(6, 0))

    def form_use_names(self) -> None:
        names = villages.load_names()
        self.form_names.delete("1.0", "end")
        self.form_names.insert("1.0", "\n".join(names))
        self.form_status.set(f"已带入 {len(names)} 个村（单位）。")

    def form_make(self) -> None:
        kind = self.form_kind.get()
        title = (self.form_title.get().strip()
                 or ("会议签到表" if kind == "sign" else "材料交接单"))
        sub = self.form_sub.get().strip()
        day = self.form_day.get().strip()
        names = [s.strip() for s in
                 self.form_names.get("1.0", "end-1c").split("\n") if s.strip()]
        if not names:
            messagebox.showinfo("表格生成", "名单是空的，先填几个（或用「材料上交」里的名单）。")
            return
        path = filedialog.asksaveasfilename(
            title="表格存成 Word", defaultextension=".docx",
            initialfile=f"{title}.docx", initialdir=str(sysenv.desktop_dir()),
            filetypes=[("Word", "*.docx")])
        if not path:
            return
        try:
            _write_table_docx(Path(path), kind, title, sub, day, names)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("表格生成", f"生成失败：{exc}")
            return
        self.form_status.set(f"已生成：{path}")
        messagebox.showinfo("表格生成", f"生成好了：\n{path}")
