# -*- coding: utf-8 -*-
"""公文模板库：左边选模板，右边填空，直接出符合版式的 .docx。"""
from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from ..core import sysenv, templates
from . import theme


class TemplatePane(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=(0, 10, 0, 0), style="Canvas.TFrame")
        self.key: str | None = None
        self.vars: dict[str, tk.Variable] = {}
        self._texts: dict[str, tk.Text] = {}
        self.columnconfigure(0, weight=0, minsize=260)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        self._build_left()
        self._build_right()
        self.tree.selection_set(self.tree.get_children()[0])

    def _build_left(self):
        outer = tk.Frame(self, background=theme.BORDER_SUBTLE)
        outer.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        c = ttk.Frame(outer, style="Card.TFrame", padding=12)
        c.pack(fill="both", expand=True, padx=1, pady=1)
        c.columnconfigure(0, weight=1)
        c.rowconfigure(1, weight=1)
        ttk.Label(c, text="公文模板库", style="H2.TLabel").grid(row=0, column=0, sticky="w")
        wrap = ttk.Frame(c, style="Card.TFrame")
        wrap.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(wrap, columns=("name",), show="tree headings", height=14)
        self.tree.heading("#0", text="模板")
        self.tree.column("#0", width=220, anchor="w")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        for k, t in templates.TEMPLATES.items():
            self.tree.insert("", "end", iid=k, text=f"{t['name']}")
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._on_pick())
        ttk.Label(c, text="填好空点「生成 Word」，版式已按公文规范设好。",
                  style="Mic.TLabel", wraplength=210, justify="left").grid(row=2, column=0, sticky="w", pady=(8, 0))

    def _build_right(self):
        outer = tk.Frame(self, background=theme.BORDER_SUBTLE)
        outer.grid(row=0, column=1, sticky="nsew")
        self.card = ttk.Frame(outer, style="Card.TFrame", padding=14)
        self.card.pack(fill="both", expand=True, padx=1, pady=1)
        self.card.columnconfigure(0, weight=1)
        self.card.rowconfigure(1, weight=1)
        self.title_lbl = ttk.Label(self.card, text="", style="H2.TLabel")
        self.title_lbl.grid(row=0, column=0, sticky="w")
        self.body = ttk.Frame(self.card, style="Card.TFrame")
        self.body.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        self.body.columnconfigure(1, weight=1)
        bar = ttk.Frame(self.card, style="Card.TFrame")
        bar.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        ttk.Button(bar, text="生成 Word…", style="Primary.TButton",
                   command=self.generate).pack(side="left")
        self.status = tk.StringVar(value="")
        ttk.Label(self.card, textvariable=self.status, style="Cap.TLabel",
                  wraplength=700, justify="left").grid(row=3, column=0, sticky="w", pady=(8, 0))

    def _on_pick(self):
        sel = self.tree.selection()
        if not sel:
            return
        self.key = sel[0]
        t = templates.TEMPLATES[self.key]
        self.title_lbl.configure(text=f"{t['name']}")
        for w in self.body.winfo_children():
            w.destroy()
        self.vars.clear()
        self._texts.clear()
        row = 0
        for key, label, default, multi in t["fields"]:
            ttk.Label(self.body, text=label, style="Cap.TLabel").grid(
                row=row, column=0, sticky="nw", pady=(6, 2), padx=(0, 10))
            if multi:
                txt = tk.Text(self.body, height=4, wrap="word", relief="flat",
                              background=theme.SUNKEN)
                txt.grid(row=row, column=1, sticky="ew", pady=(6, 2))
                txt.insert("1.0", default or "")
                self._texts[key] = txt
            else:
                var = tk.StringVar(value=default or "")
                ttk.Entry(self.body, textvariable=var).grid(row=row, column=1, sticky="ew", pady=(6, 2))
                self.vars[key] = var
            row += 1
        ttk.Label(self.body, text=t["desc"], style="Mic.TLabel",
                  wraplength=640, justify="left").grid(row=row, column=0, columnspan=2, sticky="w", pady=(10, 0))

    def _values(self) -> dict:
        v = {k: var.get() for k, var in self.vars.items()}
        for k, txt in self._texts.items():
            v[k] = txt.get("1.0", "end")
        return v

    def generate(self):
        if not self.key:
            return
        t = templates.TEMPLATES[self.key]
        title = (self.vars.get("title").get() if "title" in self.vars
                 else (self.vars["name"].get() if "name" in self.vars else t["name"]))
        default_name = f"{title or t['name']}.docx".replace("/", "-").replace("\\", "-")
        out = filedialog.asksaveasfilename(
            title="模板生成到哪", defaultextension=".docx", initialfile=default_name,
            initialdir=str(sysenv.desktop_dir()), filetypes=[("Word", "*.docx")])
        if not out:
            return
        try:
            p = templates.build(self.key, out, self._values())
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("公文模板库", f"没生成出来：{exc}")
            return
        self.status.set(f"已生成：{p}")
        messagebox.showinfo("公文模板库", f"已生成：\n{p}")
        try:
            sysenv.open_folder(str(Path(out).parent))
        except Exception:  # noqa: BLE001
            pass
