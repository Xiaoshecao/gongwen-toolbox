# -*- coding: utf-8 -*-
"""规则集编辑器：把 yaml 摊平成表格，允许直接改数值并存盘。

用户的核心要求是「以 GB/T 9704 为基础，参数可替换」——所以任何字段都要能改，
改完存成新的规则集，原 gb9704.yaml 不动。
"""
from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import yaml

from ..core import fontsizes
from ..core import rules as rulesmod
from ..core.model import ROLE_CN
from . import theme

# 可编辑字段与显示名、类型
FIELDS: list[tuple[str, str, str]] = [
    ("font", "字体", "str"),
    ("size_pt", "字号", "size"),
    ("bold", "加粗", "bool"),
    ("align", "对齐", "align"),
    ("line_pt", "行距(磅)", "num"),
    ("first_line_chars", "首行缩进(字符)", "num"),
    ("right_indent_chars", "右缩进(字符)", "num"),
    ("space_before_pt", "段前(磅)", "num"),
    ("space_after_pt", "段后(磅)", "num"),
]
ALIGNS = ["justify", "center", "right", "left", "distribute"]
ALIGN_CN = {"justify": "两端对齐", "center": "居中", "right": "右对齐",
            "left": "左对齐", "distribute": "分散对齐"}

PAGE_FIELDS = [("top_cm", "上边距cm"), ("bottom_cm", "下边距cm"),
               ("left_cm", "左边距cm"), ("right_cm", "右边距cm"),
               ("footer_cm", "页脚距cm")]


def _parse_size(raw: str) -> float | None:
    """把「五号（10.5磅）」「五号」「10.5」「10.5磅」都解析成磅值。"""
    if not raw:
        return None
    for cn, pt in fontsizes.choices():
        if raw.startswith(cn):
            return pt
    try:
        return float(raw.replace("磅", "").strip())
    except ValueError:
        return None


class RulesEditor(tk.Toplevel):
    def __init__(self, master, rules: dict, on_saved=None) -> None:
        super().__init__(master)
        self.title(f"格式标准设置 —— {rules.get('name', '')}")
        try:
            from ..core import appicon

            appicon.apply_to(self, "格式标准设置")
        except Exception:  # noqa: BLE001
            pass
        self.geometry("980x620")
        self.transient(master)
        self.rules = yaml.safe_load(yaml.safe_dump(rules, allow_unicode=True))  # 深拷贝
        self.on_saved = on_saved
        self.vars: dict[tuple[str, str], tk.Variable] = {}
        self._size_combos: dict[tuple[str, str], ttk.Combobox] = {}
        self._build()

    # ------------------------------------------------------------ 界面
    def _build(self) -> None:
        outer = ttk.Frame(self, padding=14)
        outer.pack(fill="both", expand=True)
        self.configure(background=theme.CANVAS)

        head = ttk.Frame(outer, style="Canvas.TFrame")
        head.pack(fill="x", pady=(0, 10))
        ttk.Label(head, text=f"格式标准：{self.rules.get('name', '')}",
                  style="H1.TLabel").pack(side="left")
        ttk.Label(head, text="改完点「另存为」，原文件不动",
                  style="Status.TLabel").pack(side="left", padx=(12, 0), pady=(6, 0))

        nb = ttk.Notebook(outer)
        nb.pack(fill="both", expand=True)
        nb.add(self._page_tab(nb), text="页面")
        nb.add(self._roles_tab(nb), text="段落角色")

        bar = ttk.Frame(outer, style="Canvas.TFrame")
        bar.pack(fill="x", pady=(12, 0))
        ttk.Button(bar, text="取消", style="Ghost.TButton",
                   command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(bar, text="另存为…", style="Secondary.TButton",
                   command=self.save_as).pack(side="right")
        ttk.Button(bar, text="应用到当前标准", style="Primary.TButton",
                   command=self.save_inplace).pack(side="right", padx=(0, 8))

    def _page_tab(self, parent) -> ttk.Frame:
        f = ttk.Frame(parent, padding=16)
        f.columnconfigure(1, weight=1)
        page = self.rules.setdefault("page", {})
        for i, (k, label) in enumerate(PAGE_FIELDS):
            ttk.Label(f, text=label).grid(row=i, column=0, sticky="w", pady=4, padx=(0, 12))
            var = tk.StringVar(value="" if page.get(k) is None else f"{page.get(k):g}")
            ttk.Entry(f, textvariable=var, width=14).grid(row=i, column=1, sticky="w")
            self.vars[("page", k)] = var

        ft = self.rules.setdefault("features", {})
        row = len(PAGE_FIELDS)
        ttk.Separator(f, orient="horizontal").grid(row=row, column=0, columnspan=2,
                                                  sticky="ew", pady=10)
        row += 1
        v_pn = tk.BooleanVar(value=bool(ft.get("page_number", True)))
        ttk.Checkbutton(f, text="排页码（页脚居中「— 1 —」）",
                        variable=v_pn).grid(row=row, column=0, columnspan=2, sticky="w")
        self.vars[("features", "page_number")] = v_pn
        row += 1
        ttk.Label(f, text="页码格式").grid(row=row, column=0, sticky="w", pady=4)
        v_fmt = tk.StringVar(value=ft.get("page_number_fmt", "— {n} —"))
        ttk.Entry(f, textvariable=v_fmt, width=20).grid(row=row, column=1, sticky="w")
        self.vars[("features", "page_number_fmt")] = v_fmt
        ttk.Label(f, text="（{n} 代表数字）", style="Status.TLabel").grid(
            row=row, column=1, sticky="w", padx=(170, 0))
        row += 1
        ttk.Label(f, text="页数上限").grid(row=row, column=0, sticky="w", pady=4)
        v_max = tk.StringVar(value="" if not ft.get("max_pages") else str(ft["max_pages"]))
        ttk.Entry(f, textvariable=v_max, width=14).grid(row=row, column=1, sticky="w")
        self.vars[("features", "max_pages")] = v_max
        ttk.Label(f, text="（留空不限）", style="Status.TLabel").grid(
            row=row, column=1, sticky="w", padx=(170, 0))

        tbl = self.rules.setdefault("table", {})
        row += 1
        ttk.Separator(f, orient="horizontal").grid(row=row, column=0, columnspan=2,
                                                  sticky="ew", pady=10)
        row += 1
        ttk.Label(f, text="表格字体").grid(row=row, column=0, sticky="w", pady=4)
        v_tf = tk.StringVar(value=tbl.get("font", "仿宋_GB2312"))
        ttk.Entry(f, textvariable=v_tf, width=20).grid(row=row, column=1, sticky="w")
        self.vars[("table", "font")] = v_tf
        row += 1
        ttk.Label(f, text="表格字号").grid(row=row, column=0, sticky="w", pady=4)
        # 同样用号数下拉，别让人算磅值
        _ts_pt = tbl.get("size_pt", 14)
        _ts_init = fontsizes.label(float(_ts_pt), with_pt=True) or f"{_ts_pt:g}磅"
        v_ts = tk.StringVar(value=_ts_init)
        cb_ts = ttk.Combobox(
            f, textvariable=v_ts, width=14, state="normal",
            values=[f"{cn}（{pt:g}磅）" for cn, pt in fontsizes.choices()])
        cb_ts.grid(row=row, column=1, sticky="w")
        self.vars[("table", "size_pt")] = v_ts
        self._size_combos[("table", "size_pt")] = cb_ts
        return f

    def _roles_tab(self, parent) -> ttk.Frame:
        f = ttk.Frame(parent, padding=10)
        f.columnconfigure(0, weight=1)
        f.rowconfigure(0, weight=1)

        cols = ["role"] + [c for c, _l, _t in FIELDS]
        heads = ["角色"] + [l for _c, l, _t in FIELDS]
        tv = ttk.Treeview(f, columns=cols, show="headings", height=16)
        for c, h in zip(cols, heads):
            tv.heading(c, text=h)
            tv.column(c, width=140 if c == "font" else (74 if c == "role" else 86),
                      anchor="w" if c in ("role", "font") else "center",
                      stretch=(c == "font"))
        tv.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(f, orient="vertical", command=tv.yview)
        sb.grid(row=0, column=1, sticky="ns")
        tv.configure(yscrollcommand=sb.set)
        tv.bind("<Double-1>", lambda e: self._edit_cell(e, tv))
        self.tv = tv

        roles = self.rules.get("roles", {})
        order = ["title", "subtitle", "h1", "h2", "h3", "lead", "body",
                 "table_caption", "note", "attach", "signature", "date"]
        keys = [r for r in order if r in roles] + [r for r in roles if r not in order]
        for r in keys:
            spec = roles[r] or {}
            vals = [ROLE_CN.get(r, r)]
            for c, _l, t in FIELDS:
                vals.append(self._fmt(spec.get(c), t))
            tv.insert("", "end", iid=r, values=vals)
        ttk.Label(f, text="双击单元格修改；字体留空表示沿用正文",
                  style="Status.TLabel").grid(row=1, column=0, sticky="w", pady=(6, 0))
        return f

    @staticmethod
    def _fmt(v, kind: str) -> str:
        if v is None:
            return ""
        if kind == "bool":
            return "是" if v else "否"
        if kind == "align":
            return ALIGN_CN.get(v, v)
        if kind == "size":
            # 字号按中文字号显示：16 磅 → 「三号」，基层文书看得懂
            try:
                return fontsizes.label(float(v), with_pt=False) or f"{v:g}磅"
            except (TypeError, ValueError):
                return str(v)
        if kind == "num":
            return f"{v:g}"
        return str(v)

    def _edit_size(self, tv, row: str, col: str, x: int, y: int,
                   w: int, h: int, field: str) -> None:
        """字号单元格：给一个「三号(16磅)」下拉，不让人算磅值。"""
        names = [f"{cn}（{pt:g}磅）" for cn, pt in fontsizes.choices()]
        cur = tv.set(row, col)

        # 先把当前值翻成下拉里的写法
        spec = self.rules.get("roles", {}).get(row, {})
        cur_pt = spec.get(field)
        init = ""
        if isinstance(cur_pt, (int, float)):
            for cn, pt in fontsizes.choices():
                if abs(float(cur_pt) - pt) <= 0.35:
                    init = f"{cn}（{pt:g}磅）"
                    break
            if not init:
                init = f"{cur_pt:g} 磅"

        box = ttk.Combobox(tv, values=names, state="normal")
        box.set(init or cur)
        box.place(x=x, y=y, width=max(w, 120), height=h)
        box.focus_set()
        try:
            box.select_range(0, "end")
        except Exception:  # noqa: BLE001
            pass

        def commit(_e=None):
            val = box.get().strip()
            box.destroy()
            spec2 = self.rules.setdefault("roles", {}).setdefault(row, {})
            if not val:
                spec2.pop(field, None)
                tv.set(row, col, "")
                return
            pt = None
            for cn, value in fontsizes.choices():
                if val.startswith(cn) or val == f"{value:g} 磅" or val == f"{value:g}":
                    pt = value
                    break
            if pt is None:                    # 自己填了数字就按磅值收
                try:
                    pt = float(val.replace("磅", "").strip())
                except ValueError:
                    return
            spec2[field] = int(pt) if pt == int(pt) else pt
            tv.set(row, col, self._fmt(spec2[field], "size"))

        box.bind("<<ComboboxSelected>>", commit)
        box.bind("<Return>", commit)
        box.bind("<FocusOut>", commit)
        box.bind("<Escape>", lambda _e: box.destroy())

    def _edit_cell(self, event, tv) -> None:
        row = tv.identify_row(event.y)
        col = tv.identify_column(event.x)
        if not row or not col:
            return
        idx = int(col.lstrip("#")) - 1
        if idx <= 0:
            return
        field, label, kind = FIELDS[idx - 1]
        x, y, w, h = tv.bbox(row, col)
        cur = tv.set(row, col)

        if kind == "bool":
            doc = "是" if cur == "否" else "否"
            tv.set(row, col, doc)
            spec = self.rules.setdefault("roles", {}).setdefault(row, {})
            spec[field] = (doc == "是")
            return
        if kind == "align":
            order = ALIGNS
            try:
                nxt = order[(order.index({v: k for k, v in ALIGN_CN.items()}.get(cur, cur)) + 1)
                            % len(order)]
            except (ValueError, KeyError):
                nxt = order[0]
            tv.set(row, col, ALIGN_CN[nxt])
            self.rules.setdefault("roles", {}).setdefault(row, {})[field] = nxt
            return
        if kind == "size":
            self._edit_size(tv, row, col, x, y, w, h, field)
            return

        box = ttk.Entry(tv)
        box.insert(0, cur)
        box.select_range(0, "end")
        box.place(x=x, y=y, width=w, height=h)
        box.focus_set()

        def commit(_e=None):
            val = box.get().strip()
            box.destroy()
            spec = self.rules.setdefault("roles", {}).setdefault(row, {})
            if not val:
                spec.pop(field, None)
                tv.set(row, col, "")
            elif kind == "num":
                try:
                    num = float(val)
                except ValueError:
                    return
                spec[field] = int(num) if num == int(num) else num
                tv.set(row, col, f"{spec[field]:g}")
            else:
                spec[field] = val
                tv.set(row, col, val)

        box.bind("<Return>", commit)
        box.bind("<FocusOut>", commit)
        box.bind("<Escape>", lambda _e: box.destroy())

    # ------------------------------------------------------------ 收集
    def collect(self) -> dict:
        page = self.rules.setdefault("page", {})
        for k, _l in PAGE_FIELDS:
            var = self.vars[("page", k)]
            v = var.get().strip()
            if v:
                try:
                    page[k] = float(v)
                except ValueError:
                    pass
            else:
                page.pop(k, None)

        ft = self.rules.setdefault("features", {})
        ft["page_number"] = bool(self.vars[("features", "page_number")].get())
        fmt = self.vars[("features", "page_number_fmt")].get().strip()
        ft["page_number_fmt"] = fmt or "— {n} —"
        mx = self.vars[("features", "max_pages")].get().strip()
        if mx.isdigit():
            ft["max_pages"] = int(mx)
        else:
            ft.pop("max_pages", None)

        tbl = self.rules.setdefault("table", {})
        tbl["font"] = self.vars[("table", "font")].get().strip() or "仿宋_GB2312"
        raw = self.vars[("table", "size_pt")].get().strip()
        pt = _parse_size(raw)
        if pt is None:
            tbl.pop("size_pt", None)
        else:
            tbl["size_pt"] = pt
        if not tbl:
            self.rules.pop("table", None)
        return self.rules

    def _write(self, path: Path) -> None:
        data = self.collect()
        data.pop("_path", None)
        data.pop("_key", None)
        body = yaml.safe_dump(data, allow_unicode=True, sort_keys=False,
                              default_flow_style=False)
        path.write_text("# 由「格式标准设置」修改保存\n" + body, encoding="utf-8")

    def save_as(self) -> None:
        name = filedialog.asksaveasfilename(
            title="另存格式标准", defaultextension=".yaml", initialdir=str(rulesmod.USER_DIR),
            initialfile=f"{self.rules.get('_key', 'custom')}.yaml",
            filetypes=[("规则集", "*.yaml")])
        if not name:
            return
        try:
            self._write(Path(name))
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("保存失败", str(e), parent=self)
            return
        if self.on_saved:
            self.on_saved(Path(name))
        self.destroy()

    def save_inplace(self) -> None:
        p = self.rules.get("_path")
        if not p or not Path(p).exists():
            messagebox.showinfo("提示", "当前标准来自临时路径，请用「另存为」", parent=self)
            return
        src = Path(p)
        # 内置规则集随程序发布（打包后可能是只读的），改动一律存成用户目录里的同名覆盖
        try:
            builtin = rulesmod.RULES_DIR.resolve()
        except OSError:
            builtin = rulesmod.RULES_DIR
        target = rulesmod.save_target(src.stem) if src.parent.resolve() == builtin else src
        if target == src:
            tip = f"覆盖写入：\n{p}\n\n原文件内容会被替换，建议先「另存为」备份。"
        else:
            tip = f"内置标准保持原样，改动存成你的覆盖版本：\n{target}"
        if not messagebox.askyesno("确认", tip, parent=self):
            return
        try:
            self._write(target)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("保存失败", str(e), parent=self)
            return
        if self.on_saved:
            self.on_saved(target)
        self.destroy()
