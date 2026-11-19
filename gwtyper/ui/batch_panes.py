# -*- coding: utf-8 -*-
"""批量处理四件：加水印、脱敏、数据汇总、打印体检；外加名单比对。"""
from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from ..core import imgpress, listdiff, mask, merge_tables, printcheck, sysenv, watermark
from . import dnd, theme

PDF_TYPES = [("PDF", "*.pdf"), ("所有文件", "*.*")]
ALL_TYPES = [("公文", "*.docx *.doc *.pdf *.xlsx *.xlsm"), ("所有文件", "*.*")]


def card(master, row=0, column=0, padx=(0, 0), pady=(0, 0)) -> ttk.Frame:
    outer = tk.Frame(master, background=theme.BORDER_SUBTLE)
    outer.grid(row=row, column=column, sticky="nsew", padx=padx, pady=pady)
    inner = ttk.Frame(outer, style="Card.TFrame", padding=14)
    inner.pack(fill="both", expand=True, padx=1, pady=1)
    return inner


def pick_files(types, multi=True, title="选择文件"):
    if multi:
        return list(filedialog.askopenfilenames(title=title, filetypes=types))
    f = filedialog.askopenfilename(title=title, filetypes=types)
    return [f] if f else []


class _FilePane(ttk.Frame):
    """带文件列表 + 操作区的通用页。"""

    exts: set[str] = set()
    hint = ""

    def __init__(self, master):
        super().__init__(master, padding=(0, 10, 0, 0), style="Canvas.TFrame")
        self.files: list[Path] = []
        self.busy = False
        self.columnconfigure(0, weight=1, uniform="c")
        self.columnconfigure(1, weight=1, uniform="c")
        self.rowconfigure(0, weight=1)
        self._build_left()
        self._build_right()

    def _build_left(self):
        c = card(self, 0, 0, padx=(0, 12))
        c.columnconfigure(0, weight=1)
        c.rowconfigure(2, weight=1)
        ttk.Label(c, text=self.title_text(), style="H2.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(c, text=self.hint, style="Mic.TLabel", wraplength=420,
                  justify="left").grid(row=1, column=0, sticky="w", pady=(4, 8))
        wrap = ttk.Frame(c, style="Card.TFrame")
        wrap.grid(row=2, column=0, sticky="nsew")
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(wrap, columns=("name",), show="headings", height=12)
        self.tree.heading("name", text="文件")
        self.tree.column("name", width=380, anchor="w")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        theme.stripe(self.tree)
        bar = ttk.Frame(c, style="Card.TFrame")
        bar.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(bar, text="添加文件…", style="Secondary.TButton",
                   command=self.add).pack(side="left")
        ttk.Button(bar, text="添加文件夹…", style="Ghost.TButton",
                   command=self.add_folder).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="清空", style="Ghost.TButton", command=self.clear).pack(side="left", padx=(6, 0))
        self.status = tk.StringVar(value="")
        ttk.Label(c, textvariable=self.status, style="Cap.TLabel", wraplength=420,
                  justify="left").grid(row=4, column=0, sticky="w", pady=(8, 0))
        dnd.bind_targets([self.tree, c], self._on_drop)

    def _build_right(self):
        c = card(self, 0, 1)
        self.right = c
        self.build_options(c)

    # 子类实现
    def title_text(self):
        return ""

    def build_options(self, c):
        pass

    def run_job(self):
        raise NotImplementedError

    # 通用
    def _on_drop(self, event):
        self._add(dnd.paths_from_event(self, event, accept=self._accept))

    def _accept(self, p):
        return Path(p).suffix.lower() in self.exts if self.exts else True

    def add(self):
        self._add([Path(f) for f in pick_files(ALL_TYPES)])

    def add_folder(self):
        d = filedialog.askdirectory(title="选文件夹")
        if d:
            self._add([p for p in Path(d).rglob("*") if p.is_file() and self._accept(p)])

    def _add(self, paths):
        seen = {str(p) for p in self.files}
        for p in paths:
            p = Path(p)
            if p.is_file() and self._accept(p) and str(p) not in seen:
                self.files.append(p)
                seen.add(str(p))
                self.tree.insert("", "end", values=(p.name,))
        getattr(self.tree, "_stripe", lambda: None)()
        self.status.set(f"已添加 {len(self.files)} 个文件")

    def clear(self):
        self.files.clear()
        self.tree.delete(*self.tree.get_children())
        self.status.set("")

    def _start(self, label, fn):
        if self.busy:
            return
        if not self.files:
            messagebox.showinfo(label, "先添加文件。")
            return
        self.busy = True
        self.status.set(f"{label}中…")

        def job():
            try:
                msg = fn()
            except Exception as exc:  # noqa: BLE001
                err = str(exc)
                self.after(0, lambda: self._fail(label, err))
                return
            self.after(0, lambda: self._done(label, msg))

        threading.Thread(target=job, daemon=True).start()

    def _done(self, label, msg):
        self.busy = False
        self.status.set(msg)
        messagebox.showinfo(label, msg)

    def _fail(self, label, err):
        self.busy = False
        self.status.set(f"{label}失败：{err}")
        messagebox.showerror(label, f"{label}失败：{err}")


class WatermarkPane(_FilePane):
    exts = {".pdf"}
    hint = "给 PDF 每页盖一条斜水印。只读原件，另存「-水印.pdf」。PDF 之外的格式先用「格式转换」转一下。"

    def title_text(self):
        return "批量加水印"

    def build_options(self, c):
        c.columnconfigure(1, weight=1)
        ttk.Label(c, text="水印文字", style="Cap.TLabel").grid(row=0, column=0, sticky="w")
        self.text = tk.StringVar(value="内部资料 注意保存")
        ttk.Entry(c, textvariable=self.text, width=30).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(2, 10))
        ttk.Label(c, text="透明度（0.05 很淡，0.2 明显）", style="Cap.TLabel").grid(row=2, column=0, sticky="w")
        self.opacity = tk.DoubleVar(value=0.12)
        ttk.Spinbox(c, from_=0.03, to=0.5, increment=0.02, width=8,
                    textvariable=self.opacity).grid(row=3, column=0, sticky="w", pady=(2, 10))
        ttk.Label(c, text="字号（0 = 按页面自动）", style="Cap.TLabel").grid(row=4, column=0, sticky="w")
        self.size = tk.IntVar(value=0)
        ttk.Spinbox(c, from_=0, to=120, increment=4, width=8,
                    textvariable=self.size).grid(row=5, column=0, sticky="w", pady=(2, 10))
        ttk.Label(c, text="存到哪个文件夹", style="Cap.TLabel").grid(row=6, column=0, sticky="w")
        self.out = tk.StringVar(value=str(Path(sysenv.desktop_dir()) / "加水印"))
        ttk.Entry(c, textvariable=self.out).grid(row=7, column=0, sticky="ew")
        ttk.Button(c, text="浏览…", style="Ghost.TButton",
                   command=self._choose).grid(row=7, column=1, sticky="w", padx=(6, 0))
        ttk.Button(c, text="开始加水印", style="Primary.TButton",
                   command=self.run_job).grid(row=8, column=0, sticky="w", pady=(14, 0))

    def _choose(self):
        d = filedialog.askdirectory(title="水印版 PDF 放哪")
        if d:
            self.out.set(d)

    def run_job(self):
        out = self.out.get().strip() or str(Path(sysenv.desktop_dir()) / "加水印")
        txt, op, sz = self.text.get().strip(), float(self.opacity.get() or 0.12), float(self.size.get() or 0)

        def job():
            res = watermark.apply_many(self.files, out, txt, op, sz)
            ok = sum(1 for r in res if r["ok"])
            bad = [r for r in res if not r["ok"]]
            m = f"加好了 {ok} 个，存在：{out}"
            if bad:
                m += f"；{len(bad)} 个失败（{bad[0]['msg']}）"
            return m

        self._start("加水印", job)


class MaskPane(_FilePane):
    exts = {".docx", ".xlsx", ".xlsm", ".pdf"}
    hint = "把身份证号、手机号、银行卡号中间打码，用于对外公示。只读原件，另存「-脱敏」。"

    def title_text(self):
        return "敏感信息脱敏"

    def build_options(self, c):
        from .check import CheckRow
        ttk.Label(c, text="会处理这些：", style="Cap.TLabel").pack(anchor="w")
        self.strict_id = tk.BooleanVar(value=True)
        self._strict_row = CheckRow(c, "身份证只留前 4 后 4（更严格）",
                                    variable=self.strict_id)
        self._strict_row.pack(anchor="w", pady=(0, 6))
        for line in ("身份证号：默认前 4 位 + 后 4 位保留（不勾则前 6 后 4）",
                     "手机号：前 3 位 + 后 4 位保留（中间有空格、短横也能认）",
                     "银行卡号：前 4 位 + 后 4 位保留",
                     "Word 表格 / 页眉页脚、Excel 里是数字格式的号码，也能认",
                     "支持 .docx / .xlsx / 有文字层的 .pdf"):
            ttk.Label(c, text=line, style="Mic.TLabel").pack(anchor="w")
        ttk.Label(c, text="输出文件夹", style="Cap.TLabel").pack(anchor="w", pady=(12, 2))
        self.out = tk.StringVar(value=str(Path(sysenv.desktop_dir()) / "脱敏"))
        row = ttk.Frame(c, style="Card.TFrame")
        row.pack(fill="x")
        ttk.Entry(row, textvariable=self.out).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="浏览…", style="Ghost.TButton",
                   command=lambda: self.out.set(filedialog.askdirectory() or self.out.get())).pack(side="left", padx=(6, 0))
        ttk.Button(c, text="开始脱敏", style="Primary.TButton",
                   command=self.run_job).pack(anchor="w", pady=(14, 0))

    def run_job(self):
        out_dir = Path(self.out.get().strip() or (Path(sysenv.desktop_dir()) / "脱敏"))
        out_dir.mkdir(parents=True, exist_ok=True)

        strict = bool(self.strict_id.get())

        def job():
            ok, bad, hits, nid, nph = 0, [], 0, 0, 0
            for p in self.files:
                try:
                    r = mask.run(p, out_dir / f"{p.stem}-脱敏{p.suffix}", strict)
                    ok += 1
                    hits += r.get("hits", 0)
                    nid += r.get("id", 0)
                    nph += r.get("phone", 0)
                except Exception as exc:  # noqa: BLE001
                    bad.append(f"{p.name}（{exc}）")
            m = (f"处理了 {ok} 个，共打码 {hits} 处"
                 f"（身份证 {nid}、手机号 {nph}），存在：{out_dir}")
            if bad:
                m += "\n没成功的：" + "；".join(bad[:5])
            return m

        self._start("脱敏", job)


class MergeTablePane(_FilePane):
    exts = {".xlsx", ".xlsm"}
    hint = "各村报的同格式表格，合成一张总表，自动加「来源文件」列，没报的能从文件名看出来。"

    def title_text(self):
        return "数据汇总"

    def build_options(self, c):
        ttk.Label(c, text="汇总表存成", style="Cap.TLabel").pack(anchor="w")
        self.out = tk.StringVar(value=str(Path(sysenv.desktop_dir()) / "汇总表.xlsx"))
        row = ttk.Frame(c, style="Card.TFrame")
        row.pack(fill="x", pady=(2, 10))
        ttk.Entry(row, textvariable=self.out).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="另存为…", style="Ghost.TButton",
                   command=self._choose).pack(side="left", padx=(6, 0))
        ttk.Label(c, text="每个文件取第一个工作表，第一行当表头。", style="Mic.TLabel").pack(anchor="w")
        ttk.Button(c, text="开始汇总", style="Primary.TButton",
                   command=self.run_job).pack(anchor="w", pady=(14, 0))

    def _choose(self):
        f = filedialog.asksaveasfilename(title="汇总表存成", defaultextension=".xlsx",
                                         initialfile="汇总表.xlsx",
                                         initialdir=str(Path(sysenv.desktop_dir())),
                                         filetypes=[("Excel", "*.xlsx")])
        if f:
            self.out.set(f)

    def run_job(self):
        out = self.out.get().strip()

        def job():
            r = merge_tables.summarize(self.files, out)
            m = f"汇总好 {r['files']} 个文件、{r['rows']} 行、{r['cols']} 列。\n{out}"
            if r["skipped"]:
                m += "\n跳过：" + "；".join(f"{n}（{w}）" for n, w in r["skipped"][:5])
            return m

        self._start("数据汇总", job)


class PrintCheckPane(_FilePane):
    exts = {".docx", ".doc", ".pdf"}
    hint = "看看印出来会不会难看：末页只剩一两行、空白页、表格被页边切断。"

    def title_text(self):
        return "打印体检"

    def build_options(self, c):
        c.columnconfigure(0, weight=1)
        c.rowconfigure(2, weight=1)
        ttk.Label(c, text="体检结果", style="Cap.TLabel").grid(row=0, column=0, sticky="w")
        self.result = tk.Text(c, height=16, wrap="word", relief="flat",
                              background=theme.SUNKEN, foreground=theme.TEXT)
        self.result.grid(row=2, column=0, sticky="nsew", pady=(4, 8))
        bar = ttk.Frame(c, style="Card.TFrame")
        bar.grid(row=3, column=0, sticky="ew")
        ttk.Button(bar, text="开始体检", style="Primary.TButton", command=self.run_job).pack(side="left")
        ttk.Button(bar, text="导出结果…", style="Ghost.TButton",
                   command=self.export).pack(side="left", padx=(6, 0))

    def run_job(self):
        def job():
            lines, total = [], 0
            for p in self.files:
                try:
                    fs = printcheck.check(p)
                except Exception as exc:  # noqa: BLE001
                    lines.append(f"■ {p.name}：查不了（{exc}）")
                    continue
                if not fs:
                    lines.append(f"■ {p.name}：没发现问题")
                else:
                    total += len(fs)
                    lines.append(f"■ {p.name}：")
                    for f in fs:
                        where = f"第{f['page']}页" if f["page"] else "整篇"
                        lines.append(f"    [{f['level']}] {where} {f['msg']}")
            head = f"共 {len(self.files)} 份，发现 {total} 条。\n\n"
            self.after(0, lambda: self._show(head + "\n".join(lines)))

        self._start("打印体检", job)

    def _show(self, text):
        self.busy = False
        self.result.delete("1.0", "end")
        self.result.insert("1.0", text)
        self.status.set("体检完成")

    def export(self):
        txt = self.result.get("1.0", "end").strip()
        if not txt:
            messagebox.showinfo("打印体检", "先体检一遍。")
            return
        f = filedialog.asksaveasfilename(title="导出体检结果", defaultextension=".txt",
                                         initialfile="打印体检.txt",
                                         initialdir=str(Path(sysenv.desktop_dir())),
                                         filetypes=[("文本", "*.txt")])
        if f:
            Path(f).write_text(txt, encoding="utf-8-sig")
            self.status.set(f"已导出：{f}")


class ListDiffPane(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=(0, 10, 0, 0), style="Canvas.TFrame")
        self.columnconfigure(0, weight=1, uniform="c")
        self.columnconfigure(1, weight=1, uniform="c")
        self.rowconfigure(0, weight=1)
        left = card(self, 0, 0, padx=(0, 12))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(2, weight=1)
        ttk.Label(left, text="名单比对", style="H2.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(left, text="上级下发的名单贴左边，自己手里的贴右边，点比对，"
                             "一眼看出谁少了、谁多了。", style="Mic.TLabel",
                  wraplength=420, justify="left").grid(row=1, column=0, sticky="w", pady=(4, 8))
        panes = ttk.Frame(left, style="Card.TFrame")
        panes.grid(row=2, column=0, sticky="nsew")
        panes.columnconfigure(0, weight=1, uniform="e")
        panes.columnconfigure(1, weight=1, uniform="e")
        panes.rowconfigure(1, weight=1)
        ttk.Label(panes, text="上级名单", style="Cap.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(panes, text="我的名单", style="Cap.TLabel").grid(row=0, column=1, sticky="w", padx=(8, 0))
        self.a = tk.Text(panes, height=14, wrap="none", relief="flat", background=theme.SUNKEN)
        self.b = tk.Text(panes, height=14, wrap="none", relief="flat", background=theme.SUNKEN)
        self.a.grid(row=1, column=0, sticky="nsew")
        self.b.grid(row=1, column=1, sticky="nsew", padx=(8, 0))
        bar = ttk.Frame(left, style="Card.TFrame")
        bar.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(bar, text="比对", style="Primary.TButton", command=self.compare).pack(side="left")
        ttk.Button(bar, text="导入上级名单…", style="Ghost.TButton",
                   command=lambda: self._load(self.a)).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="导入我的名单…", style="Ghost.TButton",
                   command=lambda: self._load(self.b)).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="导出结果…", style="Ghost.TButton",
                   command=self.export).pack(side="right")
        self.status = tk.StringVar(value="")
        ttk.Label(left, textvariable=self.status, style="Cap.TLabel", wraplength=420,
                  justify="left").grid(row=4, column=0, sticky="w", pady=(8, 0))

        right = card(self, 0, 1)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)
        ttk.Label(right, text="比对结果", style="Cap.TLabel").grid(row=0, column=0, sticky="w")
        self.out = tk.Text(right, wrap="word", relief="flat", background=theme.SUNKEN)
        self.out.grid(row=1, column=0, sticky="nsew", pady=(4, 0))
        self.last = ""

    def _load(self, box):
        f = filedialog.askopenfilename(title="选名单文件",
                                       filetypes=[("文本/CSV", "*.txt *.csv"), ("所有文件", "*.*")])
        if not f:
            return
        raw = Path(f).read_text(encoding="utf-8-sig", errors="replace")
        box.delete("1.0", "end")
        box.insert("1.0", raw)

    def compare(self):
        a = listdiff.parse(self.a.get("1.0", "end"))
        b = listdiff.parse(self.b.get("1.0", "end"))
        if not a or not b:
            messagebox.showinfo("名单比对", "两边都要有名单。")
            return
        r = listdiff.diff(a, b)
        self.last = listdiff.to_text(a, b, r)
        self.out.delete("1.0", "end")
        self.out.insert("1.0", self.last)
        self.status.set(f"上级 {len(a)} 个，我的 {len(b)} 个，少了 {len(r['only_a'])} 个，多了 {len(r['only_b'])} 个。")

    def export(self):
        if not self.last:
            messagebox.showinfo("名单比对", "先比对一次。")
            return
        f = filedialog.asksaveasfilename(title="导出比对结果", defaultextension=".txt",
                                         initialfile="名单比对.txt",
                                         initialdir=str(Path(sysenv.desktop_dir())),
                                         filetypes=[("文本", "*.txt")])
        if f:
            Path(f).write_text(self.last, encoding="utf-8-sig")
            self.status.set(f"已导出：{f}")
