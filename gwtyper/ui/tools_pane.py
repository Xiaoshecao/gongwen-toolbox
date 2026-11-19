# -*- coding: utf-8 -*-
"""办公小工具页：批量重命名、批量字数统计。

这两件事基层办公室天天要做：一摞材料要按「序号+名称」改名，
报材料前要统计字数。都是纯本地操作，不联网。
"""
from __future__ import annotations

import csv
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from datetime import date

from ..core import doc_source, scanim, sysenv, villages
from . import dnd
from . import theme
from .check import CheckRow
from ._tools_extra import ToolsExtraMixin

DOC_EXTS = {".docx", ".doc", ".wps", ".rtf", ".pdf", ".txt", ".md"}


def count_docx(path: Path) -> tuple[int, int]:
    """Word：返回 (页数估计, 字数)。页数拿不到就按 0。"""
    from docx import Document

    d = Document(str(path))
    chars = 0
    for p in d.paragraphs:
        chars += len("".join(p.text.split()))
    for t in d.tables:
        for row in t.rows:
            for cell in row.cells:
                chars += len("".join(cell.text.split()))
    return 0, chars


def count_pdf(path: Path) -> tuple[int, int]:
    from . import pdfmod

    fitz = pdfmod.load()
    chars = 0
    with fitz.open(str(path)) as d:
        n = d.page_count
        for page in d:
            chars += len("".join(page.get_text().split()))
    return n, chars


def count_txt(path: Path) -> tuple[int, int]:
    for enc in ("utf-8-sig", "utf-8", "gb18030", "utf-16"):
        try:
            text = path.read_text(encoding=enc)
            return 0, len("".join(text.split()))
        except UnicodeError:
            continue
    return 0, 0


def count_any(path: Path) -> tuple[int, int]:
    s = path.suffix.lower()
    if s == ".docx":
        return count_docx(path)
    if s == ".pdf":
        return count_pdf(path)
    if s in (".txt", ".md"):
        return count_txt(path)
    if s in (".doc", ".wps", ".rtf"):
        real = doc_source.to_docx(path)
        return count_docx(real)
    raise ValueError(f"统计不了这种格式：{s}")


class ToolsPane(ToolsExtraMixin, ttk.Frame):
    """办公小工具。"""

    def __init__(self, master, which: str = "all"):
        """which: all / rename / stat / deskew / form。

        单独一个功能时只建那一张卡并铺满整页，不再把四个挤在一起。
        """
        super().__init__(master, padding=(0, 8, 0, 0), style="Canvas.TFrame")
        self.which = which
        self.rename_files: list[Path] = []
        self.stat_files: list[Path] = []
        self.stat_rows: list[tuple] = []
        self.scan_files: list[Path] = []
        if which == "all":
            self.columnconfigure(0, weight=1, uniform="t")
            self.columnconfigure(1, weight=1, uniform="t")
            self.rowconfigure(0, weight=1, uniform="r")
            self.rowconfigure(1, weight=1, uniform="r")
            self._build_rename()
            self._build_stat()
            self._build_deskew()
            self._build_form()
        else:
            self.columnconfigure(0, weight=1)
            self.rowconfigure(0, weight=1)
            {"rename": self._build_rename,
             "stat": self._build_stat,
             "deskew": self._build_deskew,
             "form": self._build_form}[which]()

    # ---------------------------------------------------------- 卡片外壳
    def _card(self, row: int, column: int) -> ttk.Frame:
        outer = tk.Frame(self, background="#E4E8EE")
        if self.which == "all":
            outer.grid(row=row, column=column, sticky="nsew",
                       padx=(12 if column else 0, 12),
                       pady=(12 if row else 0, 0))
        else:
            outer.grid(row=0, column=0, sticky="nsew")
        card = ttk.Frame(outer, style="Card.TFrame", padding=14)
        card.pack(fill="both", expand=True, padx=1, pady=1)
        return card

    # ---------------------------------------------------------- 批量重命名
    def _build_rename(self) -> None:
        from . import theme

        card = self._card(0, 0)
        card.columnconfigure(0, weight=1)
        card.rowconfigure(2, weight=1)
        ttk.Label(card, text="批量重命名", style="H2.TLabel").grid(
            row=0, column=0, sticky="w")
        ttk.Label(card, text="一摞材料改成一串规整的名字，或把文件名里某个词换掉。",
                  style="Mic.TLabel", wraplength=470, justify="left").grid(
            row=1, column=0, sticky="w", pady=(4, 8))

        body = ttk.Frame(card, style="Card.TFrame")
        body.grid(row=2, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)
        bar = ttk.Frame(body, style="Card.TFrame")
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Button(bar, text="添加文件…", style="Secondary.TButton",
                   command=self.rename_add_files).pack(side="left")
        ttk.Button(bar, text="添加文件夹…", style="Ghost.TButton",
                   command=self.rename_add_folder).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="清空", style="Ghost.TButton",
                   command=self.rename_clear).pack(side="left", padx=(6, 0))

        wrap = ttk.Frame(body, style="Card.TFrame")
        wrap.grid(row=1, column=0, sticky="nsew")
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        self.rn_tree = ttk.Treeview(wrap, columns=("old", "new"), show="headings",
                                    height=6)
        self.rn_tree.heading("old", text="现在")
        self.rn_tree.heading("new", text="改成")
        self.rn_tree.column("old", width=180, anchor="w")
        self.rn_tree.column("new", width=180, anchor="w")
        theme.stripe(self.rn_tree)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.rn_tree.yview)
        self.rn_tree.configure(yscrollcommand=sb.set)
        self.rn_tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        opt = ttk.Frame(card, style="Card.TFrame")
        opt.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        self.rn_prefix = tk.StringVar(value="")
        self.rn_start = tk.IntVar(value=1)
        self.rn_width = tk.IntVar(value=2)
        self.rn_number = tk.BooleanVar(value=True)
        self.rn_find = tk.StringVar(value="")
        self.rn_repl = tk.StringVar(value="")
        self.rn_suffix = tk.StringVar(value="")
        self.rn_date = tk.BooleanVar(value=False)
        ttk.Label(opt, text="名字前缀", style="Cap.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(opt, textvariable=self.rn_prefix, width=14).grid(row=0, column=1, padx=(6, 4))
        ttk.Label(opt, text="从", style="Cap.TLabel").grid(row=0, column=2)
        self.rn_start_sp = ttk.Spinbox(opt, from_=0, to=9999, width=5,
                                       textvariable=self.rn_start)
        self.rn_start_sp.grid(row=0, column=3, padx=(4, 4))
        ttk.Label(opt, text="位数", style="Cap.TLabel").grid(row=0, column=4)
        self.rn_width_sp = ttk.Spinbox(opt, from_=1, to=6, width=4,
                                       textvariable=self.rn_width)
        self.rn_width_sp.grid(row=0, column=5, padx=(4, 0))
        self.rn_num_row = CheckRow(opt, "带序号", variable=self.rn_number,
                                   command=self._sync_rename_num)
        self.rn_num_row.grid(row=0, column=6, padx=(12, 0))
        ttk.Label(opt, text="查找", style="Cap.TLabel").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(opt, textvariable=self.rn_find, width=14).grid(row=1, column=1, padx=(6, 4), pady=(6, 0))
        ttk.Label(opt, text="替换成", style="Cap.TLabel").grid(row=1, column=2, pady=(6, 0))
        ttk.Entry(opt, textvariable=self.rn_repl, width=14).grid(row=1, column=3, columnspan=2,
                                                                 padx=(4, 0), pady=(6, 0))

        ttk.Label(opt, text="名字后缀", style="Cap.TLabel").grid(
            row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(opt, textvariable=self.rn_suffix, width=14).grid(
            row=2, column=1, padx=(6, 4), pady=(6, 0))
        ttk.Label(opt, text="（接在文件名末尾，如「-已核」；扩展名不动）",
                  style="Mic.TLabel").grid(row=2, column=2, columnspan=4,
                                           sticky="w", pady=(6, 0))

        self.rn_date_row = CheckRow(opt, "名字前面再加今天的日期（20261118-）",
                                   variable=self.rn_date)
        self.rn_date_row.grid(row=3, column=0, columnspan=6, sticky="w", pady=(6, 0))

        act = ttk.Frame(card, style="Card.TFrame")
        act.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(act, text="预览新名字", style="Secondary.TButton",
                   command=self.rename_preview).pack(side="left")
        ttk.Button(act, text="开始改名", style="Primary.TButton",
                   command=self.rename_apply).pack(side="left", padx=(6, 0))
        self.rn_status = tk.StringVar(value="先添加文件，设好规则，点「预览新名字」看看。")
        ttk.Label(card, textvariable=self.rn_status, style="Cap.TLabel",
                  wraplength=470, justify="left").grid(row=5, column=0, sticky="w", pady=(8, 0))
        dnd.bind_targets([self.rn_tree, wrap], self._on_drop_rename)

    def _sync_rename_num(self) -> None:
        """不勾「带序号」时把「从/位数」置灰。"""
        on = bool(self.rn_number.get())
        state = "normal" if on else "disabled"
        for sp in (getattr(self, "rn_start_sp", None),
                   getattr(self, "rn_width_sp", None)):
            if sp is not None:
                try:
                    sp.configure(state=state)
                except tk.TclError:
                    pass
        preview = getattr(self, "rn_status", None)
        if preview is not None and self.rename_files:
            self.rename_preview()

    def _on_drop_rename(self, event) -> None:
        paths = dnd.paths_from_event(self, event, accept=dnd.suffix_filter(set(DOC_EXTS) | {""}))
        if paths:
            self._add_rename([Path(p) for p in paths])

    def rename_add_files(self) -> None:
        fs = filedialog.askopenfilenames(title="选文件", filetypes=[("文档", "*.*")])
        if fs:
            self._add_rename([Path(f) for f in fs])

    def rename_add_folder(self) -> None:
        d = filedialog.askdirectory(title="选文件夹")
        if d:
            self._add_rename(sorted(p for p in Path(d).iterdir() if p.is_file()))

    def _add_rename(self, paths: list[Path]) -> None:
        for p in paths:
            if p not in self.rename_files:
                self.rename_files.append(p)
        self.rename_preview()

    def rename_clear(self) -> None:
        self.rename_files.clear()
        self.rn_tree.delete(*self.rn_tree.get_children())
        self.rn_status.set("已清空。")

    def _new_name(self, p: Path, i: int) -> str:
        stem = p.stem
        if self.rn_date.get():
            stem = f"{date.today():%Y%m%d}-{stem}"
        find = self.rn_find.get()
        if find:
            stem = stem.replace(find, self.rn_repl.get())
        pre = self.rn_prefix.get()
        num_on = bool(self.rn_number.get())
        if pre or num_on:
            if num_on:
                try:
                    w = max(1, int(self.rn_width.get() or 2))
                    n = int(self.rn_start.get() or 1) + i
                except (tk.TclError, ValueError):
                    w, n = 2, i + 1
                mid = f"{pre}{str(n).zfill(w)}"
            else:
                mid = pre
            stem = f"{mid}{stem}"
        suf = self.rn_suffix.get()
        if suf:
            stem = f"{stem}{suf}"
        return stem + p.suffix

    def rename_preview(self) -> None:
        self.rn_tree.delete(*self.rn_tree.get_children())
        if not self.rename_files:
            self.rn_status.set("还没有文件。")
            return
        seen: dict[str, int] = {}
        for i, p in enumerate(self.rename_files):
            new = self._new_name(p, i)
            seen[new.lower()] = seen.get(new.lower(), 0) + 1
            self.rn_tree.insert("", "end", values=(p.name, new))
        getattr(self.rn_tree, "_stripe", lambda: None)()
        dup = [k for k, c in seen.items() if c > 1]
        if dup:
            self.rn_status.set(f"注意：有 {len(dup)} 个新名字重复了，改一下前缀或序号再试。")
        else:
            self.rn_status.set(f"{len(self.rename_files)} 个文件，新名字没有重复，可以改名。")

    def rename_apply(self) -> None:
        if not self.rename_files:
            messagebox.showinfo("批量重命名", "先添加文件。")
            return
        plan = []
        for i, p in enumerate(self.rename_files):
            new = self._new_name(p, i)
            if new == p.name:
                continue
            target = p.with_name(new)
            if target.exists():
                messagebox.showwarning(
                    "批量重命名", f"「{new}」已经存在，先换个规则吧。\n（{p.name}）")
                return
            plan.append((p, target))
        if not plan:
            self.rn_status.set("没有需要改的名字。")
            return
        if not messagebox.askyesno("批量重命名", f"确定要改 {len(plan)} 个文件名吗？"):
            return
        ok, failed = 0, []
        for p, target in plan:
            try:
                p.rename(target)
                ok += 1
            except OSError as exc:
                failed.append(f"{p.name}：{exc}")
        self.rename_files = [t for _, t in plan]
        self.rename_preview()
        msg = f"改好了 {ok} 个。"
        if failed:
            msg += f"\n失败 {len(failed)} 个：\n" + "\n".join(failed[:5])
        self.rn_status.set(msg)
        (messagebox.showwarning if failed else messagebox.showinfo)("批量重命名", msg)

    # ---------------------------------------------------------- 批量字数统计
    def _build_stat(self) -> None:
        card = self._card(0, 1)
        card.columnconfigure(0, weight=1)
        card.rowconfigure(2, weight=1)
        ttk.Label(card, text="批量字数统计", style="H2.TLabel").grid(
            row=0, column=0, sticky="w")
        ttk.Label(card, text="报材料前统计一堆文件各多少字，可导出 CSV 交差。",
                  style="Mic.TLabel", wraplength=470, justify="left").grid(
            row=1, column=0, sticky="w", pady=(4, 8))

        body = ttk.Frame(card, style="Card.TFrame")
        body.grid(row=2, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)
        bar = ttk.Frame(body, style="Card.TFrame")
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Button(bar, text="添加文件…", style="Secondary.TButton",
                   command=self.stat_add_files).pack(side="left")
        ttk.Button(bar, text="添加文件夹…", style="Ghost.TButton",
                   command=self.stat_add_folder).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="清空", style="Ghost.TButton",
                   command=self.stat_clear).pack(side="left", padx=(6, 0))

        wrap = ttk.Frame(body, style="Card.TFrame")
        wrap.grid(row=1, column=0, sticky="nsew")
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        self.st_tree = ttk.Treeview(wrap, columns=("name", "pages", "chars"),
                                    show="headings", height=6)
        for c, t, w, a in (("name", "文件", 200, "w"), ("pages", "页", 50, "center"),
                           ("chars", "字数", 70, "e")):
            self.st_tree.heading(c, text=t)
            self.st_tree.column(c, width=w, anchor=a)
        theme.stripe(self.st_tree)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.st_tree.yview)
        self.st_tree.configure(yscrollcommand=sb.set)
        self.st_tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        act = ttk.Frame(card, style="Card.TFrame")
        act.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(act, text="开始统计", style="Primary.TButton",
                   command=self.stat_run).pack(side="left")
        ttk.Button(act, text="导出 CSV…", style="Ghost.TButton",
                   command=self.stat_export).pack(side="left", padx=(6, 0))
        self.st_status = tk.StringVar(value="添加文档后点「开始统计」。支持 docx / doc / pdf / txt。")
        ttk.Label(card, textvariable=self.st_status, style="Cap.TLabel",
                  wraplength=470, justify="left").grid(row=4, column=0, sticky="w", pady=(8, 0))
        dnd.bind_targets([self.st_tree, wrap], self._on_drop_stat)

    def _on_drop_stat(self, event) -> None:
        paths = dnd.paths_from_event(self, event,
                                    accept=dnd.suffix_filter(DOC_EXTS))
        if paths:
            self._add_stat([Path(p) for p in paths])

    def stat_add_files(self) -> None:
        fs = filedialog.askopenfilenames(
            title="选文件",
            filetypes=[("文档", "*.docx *.doc *.wps *.rtf *.pdf *.txt *.md"),
                       ("全部文件", "*.*")])
        if fs:
            self._add_stat([Path(f) for f in fs])

    def stat_add_folder(self) -> None:
        d = filedialog.askdirectory(title="选文件夹")
        if d:
            ps = [p for p in sorted(Path(d).iterdir())
                  if p.is_file() and p.suffix.lower() in DOC_EXTS]
            self._add_stat(ps)

    def _add_stat(self, paths: list[Path]) -> None:
        for p in paths:
            if p not in self.stat_files:
                self.stat_files.append(p)
        self.st_status.set(f"已添加 {len(self.stat_files)} 个文件。")

    def stat_clear(self) -> None:
        self.stat_files.clear()
        self.stat_rows.clear()
        self.st_tree.delete(*self.st_tree.get_children())
        self.st_status.set("已清空。")

    def stat_run(self) -> None:
        if not self.stat_files:
            messagebox.showinfo("字数统计", "先添加文件。")
            return
        self.st_tree.delete(*self.st_tree.get_children())
        self.stat_rows = []
        failed = []
        for p in self.stat_files:
            try:
                pages, chars = count_any(p)
                self.stat_rows.append((p.name, pages, chars))
                self.st_tree.insert("", "end", values=(p.name, pages or "-", chars))
            except Exception as exc:  # noqa: BLE001
                failed.append(f"{p.name}：{exc}")
        getattr(self.st_tree, "_stripe", lambda: None)()
        total = sum(r[2] for r in self.stat_rows)
        msg = f"{len(self.stat_rows)} 个文件，合计 {total} 字。"
        if failed:
            msg += f"　{len(failed)} 个没统计成（{failed[0]}）"
        self.st_status.set(msg)

    def stat_export(self) -> None:
        if not self.stat_rows:
            messagebox.showinfo("字数统计", "先点「开始统计」。")
            return
        path = filedialog.asksaveasfilename(
            title="字数统计存成 CSV",
            defaultextension=".csv",
            initialfile="字数统计.csv",
            initialdir=str(sysenv.desktop_dir()),
            filetypes=[("CSV", "*.csv")])
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(["文件", "页数", "字数"])
                for name, pages, chars in self.stat_rows:
                    w.writerow([name, pages or "", chars])
                w.writerow(["合计", "", sum(r[2] for r in self.stat_rows)])
        except OSError as exc:
            messagebox.showerror("字数统计", f"写不了文件：{exc}")
            return
        self.st_status.set(f"已导出：{path}")
        messagebox.showinfo("字数统计", f"导出了 {len(self.stat_rows)} 行：\n{path}")
