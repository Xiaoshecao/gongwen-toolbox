# -*- coding: utf-8 -*-
"""「材料上交」页：按村名单核对谁交了、谁没交。

- 村名单可以贴进来，自动记住（下次打开还在）
- 可以手动双击标记，也可以选一个文件夹，按文件名自动认哪个村交了
- 一键导出未交清单（txt / csv），方便群里催
"""
from __future__ import annotations

import csv
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from ..core import sysenv, villages
from . import theme

TICK = "\u2714"
CROSS = "\u2716"


class ChecklistPane(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=(0, 8, 0, 0), style="Canvas.TFrame")
        self.names: list[str] = villages.load_names()
        self.status: dict[str, dict] = villages.load_status()
        self.last_dir = ""
        self.columnconfigure(0, weight=2, uniform="v")
        self.columnconfigure(1, weight=3, uniform="v")
        self.rowconfigure(0, weight=1)
        self._build_names()
        self._build_list()
        self._refresh_presets()
        self.refresh()

    # ------------------------------------------------------------ 卡片
    def _card(self, column: int, padx) -> ttk.Frame:
        outer = tk.Frame(self, background="#E4E8EE")
        outer.grid(row=0, column=column, sticky="nsew", padx=padx)
        card = ttk.Frame(outer, style="Card.TFrame", padding=14)
        card.pack(fill="both", expand=True, padx=1, pady=1)
        return card

    # ------------------------------------------------------------ 左：名单
    def _build_names(self) -> None:
        card = self._card(0, (12, 12))
        card.columnconfigure(0, weight=1)
        card.rowconfigure(2, weight=1)
        ttk.Label(card, text="村（单位）名单", style="H2.TLabel").grid(
            row=0, column=0, sticky="w")
        ttk.Label(card, text="一行一个（村、社区、站所、科室、学校、公司都行，"
                             "不必以「村」结尾）。点「保存名单」，下次打开自动带出来；"
                             "也可以从 Excel 里解析，解析完先让你核对。",
                  style="Mic.TLabel", wraplength=460, justify="left").grid(
            row=1, column=0, sticky="w", pady=(4, 8))

        pre = ttk.Frame(card, style="Card.TFrame")
        pre.grid(row=5, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(pre, text="预设", style="Cap.TLabel").pack(side="left")
        self.preset = tk.StringVar(value="")
        self.preset_box = ttk.Combobox(pre, textvariable=self.preset, width=14,
                                       state="readonly")
        self.preset_box.pack(side="left", padx=(6, 0))
        self.preset_box.bind("<<ComboboxSelected>>", lambda e: self.load_preset())
        ttk.Button(pre, text="载入", style="Ghost.TButton",
                   command=self.load_preset).pack(side="left", padx=(6, 0))
        ttk.Button(pre, text="存为预设…", style="Ghost.TButton",
                   command=self.save_as_preset).pack(side="left", padx=(6, 0))
        ttk.Button(pre, text="删", style="Ghost.TButton",
                   command=self.delete_preset).pack(side="left", padx=(4, 0))
        wrap = ttk.Frame(card, style="Card.TFrame")
        wrap.grid(row=2, column=0, sticky="nsew")
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        self.txt = tk.Text(wrap, height=18, wrap="none", undo=True,
                           relief="flat", highlightthickness=1,
                           highlightbackground="#D3D9E2", highlightcolor="#7EA6F2",
                           font=("Microsoft YaHei UI", 10), padx=8, pady=6)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set)
        self.txt.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self.txt.insert("1.0", "\n".join(self.names))

        act = ttk.Frame(card, style="Card.TFrame")
        act.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(act, text="保存名单", style="Primary.TButton",
                   command=self.save_names).pack(side="left")
        ttk.Button(act, text="导入 txt…", style="Ghost.TButton",
                   command=self.import_names).pack(side="left", padx=(6, 0))
        ttk.Button(act, text="从 Excel 解析…", style="Secondary.TButton",
                   command=self.import_excel).pack(side="left", padx=(6, 0))
        ttk.Button(act, text="清空", style="Ghost.TButton",
                   command=self.clear_names).pack(side="left", padx=(6, 0))
        self.name_status = tk.StringVar(value="")
        ttk.Label(card, textvariable=self.name_status, style="Cap.TLabel",
                  wraplength=460, justify="left").grid(row=6, column=0, sticky="w",
                                                       pady=(8, 0))

    def _names_from_text(self) -> list[str]:
        out, seen = [], set()
        for line in self.txt.get("1.0", "end-1c").replace("\r", "\n").split("\n"):
            s = villages.strip_ordinal(line.strip())
            if not s:
                continue
            k = villages.normalize(s)
            if k and k not in seen:
                seen.add(k)
                out.append(s)
        return out

    def save_names(self) -> None:
        names = self._names_from_text()
        try:
            p = villages.save_names(names)
        except OSError as exc:
            messagebox.showerror("村名单", f"存不了：{exc}")
            return
        self.names = names
        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", "\n".join(names))
        self.name_status.set(f"已保存 {len(names)} 个村（{p}）")
        self.refresh()

    def import_names(self) -> None:
        f = filedialog.askopenfilename(title="选村名单",
                                       filetypes=[("文本", "*.txt *.csv"), ("全部", "*.*")])
        if not f:
            return
        try:
            raw = Path(f).read_text(encoding="utf-8-sig")
        except UnicodeError:
            try:
                raw = Path(f).read_text(encoding="gb18030")
            except (OSError, UnicodeError) as exc:
                messagebox.showerror("村名单", f"读不了：{exc}")
                return
        except OSError as exc:
            messagebox.showerror("村名单", f"读不了：{exc}")
            return
        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", raw.strip())
        self.name_status.set("已导入，记得点「保存名单」。")

    def _refresh_presets(self) -> None:
        names = villages.list_presets()
        self.preset_box["values"] = names
        if names and not self.preset.get():
            self.preset.set(names[0])

    def load_preset(self) -> None:
        name = self.preset.get().strip()
        if not name:
            return
        names = villages.load_preset(name)
        if not names:
            messagebox.showinfo("预设名单", f"预设「{name}」是空的。")
            return
        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", "\n".join(names))
        self.name_status.set(f"已载入预设「{name}」（{len(names)} 个），记得点「保存名单」生效。")

    def save_as_preset(self) -> None:
        names = self._names_from_text()
        if not names:
            messagebox.showinfo("预设名单", "名单是空的，先把名称贴进来。")
            return
        name = simpledialog.askstring("存为预设", "给这份名单起个名字（比如：2026各村）：",
                                      initialvalue=self.preset.get() or "各单位名单")
        if not name:
            return
        try:
            villages.save_preset(name, names)
        except OSError as exc:
            messagebox.showerror("预设名单", f"存不了：{exc}")
            return
        self.preset.set(villages._safe(name))
        self._refresh_presets()
        self.name_status.set(f"已存为预设「{name}」（{len(names)} 个）。")

    def delete_preset(self) -> None:
        name = self.preset.get().strip()
        if not name:
            return
        if not messagebox.askyesno("预设名单", f"删掉预设「{name}」？（当前名单不受影响）"):
            return
        villages.delete_preset(name)
        self.preset.set("")
        self._refresh_presets()
        self.name_status.set(f"已删除预设「{name}」。")

    def import_excel(self) -> None:
        f = filedialog.askopenfilename(
            title="选 Excel（村里报上来的表就行）",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("全部", "*.*")])
        if not f:
            return
        try:
            got = villages.parse_excel_names(f)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("从 Excel 解析", f"解析不了：{exc}")
            return
        names = got.get("names") or []
        if not names:
            messagebox.showinfo("从 Excel 解析",
                                "这张表里没找到像单位名称的内容。\n"
                                "可以把名单整理成一列（第一行写「名称」也行），再试一次。")
            return
        dlg = NameConfirmDialog(self, names, got.get("sheet", ""), got.get("column", ""))
        self.wait_window(dlg)
        if not dlg.result:
            self.name_status.set("已取消，名单没动。")
            return
        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", "\n".join(dlg.result))
        self.name_status.set(
            f"解析出 {len(dlg.result)} 个（表「{got.get('sheet')}」第 {got.get('column')} 列），"
            "核对没问题就点「保存名单」；想留着下次用就点「存为预设…」。")

    def clear_names(self) -> None:
        if not messagebox.askyesno("村名单", "清空名单？勾选记录也会跟着清掉。"):
            return
        self.txt.delete("1.0", "end")
        self.names = []
        self.status = {}
        villages.save_names([])
        villages.save_status({})
        self.name_status.set("已清空。")
        self.refresh()

    # ------------------------------------------------------------ 右：台账
    def _build_list(self) -> None:
        card = self._card(1, (0, 12))
        card.columnconfigure(0, weight=1)
        card.rowconfigure(2, weight=1)
        ttk.Label(card, text="上交情况", style="H2.TLabel").grid(
            row=0, column=0, sticky="w")
        ttk.Label(card, text="双击某一行切换「已交 / 未交」；选文件夹时，"
                             "文件名里含这个名称就算它交了（不必是「某某村」）。",
                  style="Mic.TLabel", wraplength=520, justify="left").grid(
            row=1, column=0, sticky="w", pady=(4, 8))

        wrap = ttk.Frame(card, style="Card.TFrame")
        wrap.grid(row=2, column=0, sticky="nsew")
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(wrap, columns=("name", "state", "why"),
                                 show="headings", height=16)
        for c, t, w, a in (("name", "村（单位）", 200, "w"),
                           ("state", "上交", 70, "center"),
                           ("why", "依据", 240, "w")):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor=a)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        theme.stripe(self.tree)
        self.tree.tag_configure("done", foreground="#0E9F6E")
        self.tree.tag_configure("todo", foreground="#B01C1C")
        self.tree.bind("<Double-1>", self._on_double)

        act = ttk.Frame(card, style="Card.TFrame")
        act.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(act, text="选文件夹按文件名认…", style="Primary.TButton",
                   command=self.detect_from_files).pack(side="left")
        ttk.Button(act, text="全部已交", style="Ghost.TButton",
                   command=lambda: self._set_all(True)).pack(side="left", padx=(6, 0))
        ttk.Button(act, text="全部未交", style="Ghost.TButton",
                   command=lambda: self._set_all(False)).pack(side="left", padx=(6, 0))
        ttk.Button(act, text="导出未交清单…", style="Secondary.TButton",
                   command=self.export_todo).pack(side="right")

        act2 = ttk.Frame(card, style="Card.TFrame")
        act2.grid(row=4, column=0, sticky="ew", pady=(6, 0))
        ttk.Button(act2, text="生成通报 Word…", style="Secondary.TButton",
                   command=self.make_notice).pack(side="left")
        ttk.Button(act2, text="按台账归档…", style="Secondary.TButton",
                   command=self.archive_files).pack(side="left", padx=(6, 0))

        self.sum_text = tk.StringVar(value="")
        ttk.Label(card, textvariable=self.sum_text, style="Cap.TLabel",
                  wraplength=560, justify="left").grid(row=5, column=0, sticky="w",
                                                       pady=(8, 0))

    # ------------------------------------------------------------ 逻辑
    def refresh(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for n in self.names:
            st = self.status.get(n) or {}
            done = bool(st.get("done"))
            why = str(st.get("src") or "")
            self.tree.insert("", "end", iid=n, tags=("done" if done else "todo",),
                             values=(n, f"{TICK} 已交" if done else f"{CROSS} 未交", why))
        getattr(self.tree, "_stripe", lambda: None)()
        self._update_sum()

    def _update_sum(self) -> None:
        total = len(self.names)
        done = sum(1 for n in self.names if (self.status.get(n) or {}).get("done"))
        if not total:
            self.sum_text.set("还没有名单。左边贴一份，点「保存名单」。")
        else:
            self.sum_text.set(
                f"共 {total} 个，已交 {done} 个，未交 {total - done} 个。"
                + ("" if done < total else "　全部交齐了。"))

    def _persist(self) -> None:
        try:
            villages.save_status(self.status)
        except OSError:
            pass

    def _on_double(self, event) -> None:
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        st = self.status.setdefault(iid, {"done": False, "src": ""})
        st["done"] = not st.get("done")
        st["src"] = "手动标记" if st["done"] else ""
        self._persist()
        self.refresh()

    def _set_all(self, done: bool) -> None:
        for n in self.names:
            self.status[n] = {"done": done, "src": "批量标记" if done else ""}
        self._persist()
        self.refresh()

    def detect_from_files(self) -> None:
        if not self.names:
            messagebox.showinfo("材料上交", "先在左边保存名单。")
            return
        d = filedialog.askdirectory(title="选材料所在的文件夹", initialdir=self.last_dir or None)
        if not d:
            return
        self.last_dir = d
        files = [str(p) for p in sorted(Path(d).iterdir()) if p.is_file()]
        if not files:
            # 很多人的材料是一个村一个子文件夹，往里再扫一层
            files = [str(p) for p in sorted(Path(d).rglob("*")) if p.is_file()]
        if not files:
            messagebox.showinfo("材料上交", "这个文件夹里没有文件。")
            return
        hits = villages.match_files(self.names, files)
        hit = 0
        for n, fname in hits.items():
            if fname:
                self.status[n] = {"done": True, "src": f"文件名：{fname}"}
                hit += 1
        self._persist()
        self.refresh()
        miss = [n for n in self.names if not (self.status.get(n) or {}).get("done")]
        self.sum_text.set(
            f"扫了 {len(files)} 个文件，认出 {hit} 个村已交；还有 {len(miss)} 个没认到。"
            + ("" if not miss else "　（认不到的可以双击手动标）"))

    def make_notice(self) -> None:
        """按台账生成一份《材料上报情况通报》Word。"""
        if not self.names:
            messagebox.showinfo("生成通报", "还没有名单，先在上面贴一份。")
            return
        title = simpledialog.askstring("生成通报", "通报标题：",
                                       initialvalue="关于材料上报情况的通报")
        if not title:
            return
        deadline = simpledialog.askstring("生成通报", "截止时间（可留空）：", initialvalue="") or ""
        issuer = simpledialog.askstring("生成通报", "落款单位（可留空）：", initialvalue="") or ""
        out = filedialog.asksaveasfilename(
            title="通报存到哪", defaultextension=".docx",
            initialfile=f"{title}.docx", initialdir=str(sysenv.desktop_dir()),
            filetypes=[("Word", "*.docx")])
        if not out:
            return
        try:
            from ..core import notice

            notice.build(out, title, self.names, self.status, deadline=deadline, issuer=issuer)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("生成通报", f"没生成出来：{exc}")
            return
        self.sum_text.set(f"通报已生成：{out}")
        messagebox.showinfo("生成通报", f"已生成：\n{out}")
        try:
            sysenv.open_folder(str(Path(out).parent))
        except Exception:  # noqa: BLE001
            pass

    def archive_files(self) -> None:
        """把收到的材料按台账归到「单位名」子文件夹。"""
        if not self.names:
            messagebox.showinfo("按台账归档", "还没有名单。")
            return
        files = filedialog.askopenfilenames(title="选收到的材料（可多选）")
        if not files:
            return
        out = filedialog.askdirectory(title="归档到哪个文件夹（会在里面按单位建子文件夹）")
        if not out:
            return
        try:
            from ..core import archive

            r = archive.archive(files, self.names, out)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("按台账归档", f"归档失败：{exc}")
            return
        n = sum(len(v) for v in r["done"].values())
        msg = f"归档了 {n} 个文件，放进了 {len(r['done'])} 个单位文件夹。"
        if r["unknown"]:
            msg += f"\n有 {len(r['unknown'])} 个没认出是哪个单位：\n" + "\n".join(
                Path(x).name for x in r["unknown"][:8])
        self.sum_text.set(msg.replace("\n", " "))
        messagebox.showinfo("按台账归档", msg)
        try:
            sysenv.open_folder(r["out"])
        except Exception:  # noqa: BLE001
            pass

    def export_todo(self) -> None:
        todo = [n for n in self.names if not (self.status.get(n) or {}).get("done")]
        if not self.names:
            messagebox.showinfo("材料上交", "还没有名单。")
            return
        path = filedialog.asksaveasfilename(
            title="导出未交清单", defaultextension=".txt",
            initialfile=f"未交清单-{date.today():%Y%m%d}.txt",
            initialdir=str(sysenv.desktop_dir()),
            filetypes=[("文本", "*.txt"), ("CSV", "*.csv")])
        if not path:
            return
        try:
            if path.lower().endswith(".csv"):
                with open(path, "w", newline="", encoding="utf-8-sig") as f:
                    w = csv.writer(f)
                    w.writerow(["村（单位）", "上交", "依据"])
                    for n in self.names:
                        st = self.status.get(n) or {}
                        w.writerow([n, "已交" if st.get("done") else "未交",
                                    st.get("src") or ""])
            else:
                lines = [f"材料上交情况（{date.today():%Y年%m月%d日}）",
                         f"共 {len(self.names)} 个，已交 {len(self.names) - len(todo)} 个，"
                         f"未交 {len(todo)} 个。", ""]
                if todo:
                    lines.append("未交：")
                    lines += [f"  {i}. {n}" for i, n in enumerate(todo, 1)]
                else:
                    lines.append("全部交齐。")
                Path(path).write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
        except OSError as exc:
            messagebox.showerror("材料上交", f"写不了文件：{exc}")
            return
        self.sum_text.set(f"已导出：{path}")
        messagebox.showinfo("材料上交", f"已导出（未交 {len(todo)} 个）：\n{path}")


class NameConfirmDialog(tk.Toplevel):
    """解析结果先看一眼：能改、能删行，确认了才往名单里填。"""

    def __init__(self, master, names: list[str], sheet: str = "", column: str = ""):
        super().__init__(master)
        self.result: list[str] | None = None
        self.title("核对解析出来的名称")
        self.configure(background="#EDEFF3")
        self.transient(master)
        frm = ttk.Frame(self, style="Card.TFrame", padding=14)
        frm.pack(fill="both", expand=True, padx=1, pady=1)
        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(2, weight=1)

        head = "从 Excel 解析出下面这些名称，核对一下：" if sheet else "解析出下面这些名称，核对一下："
        ttk.Label(frm, text=head, style="H2.TLabel").grid(row=0, column=0, sticky="w")
        where = ""
        if sheet:
            where = f"（表「{sheet}」第 {column} 列）"
        ttk.Label(frm, text=f"{where}多一行少一行都可以改；确认后才会填进名单。",
                  style="Mic.TLabel", wraplength=460, justify="left").grid(
            row=1, column=0, sticky="w", pady=(4, 8))

        wrap = ttk.Frame(frm, style="Card.TFrame")
        wrap.grid(row=2, column=0, sticky="nsew")
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        self.txt = tk.Text(wrap, height=14, wrap="none", relief="flat",
                           highlightthickness=1, highlightbackground="#D3D9E2",
                           highlightcolor="#7EA6F2",
                           font=("Microsoft YaHei UI", 10), padx=8, pady=6)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set)
        self.txt.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self.txt.insert("1.0", "\n".join(names))

        act = ttk.Frame(frm, style="Card.TFrame")
        act.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        self.count = tk.StringVar(value=f"共 {len(names)} 行")
        ttk.Label(act, textvariable=self.count, style="Cap.TLabel").pack(side="left")
        ttk.Button(act, text="取消", style="Ghost.TButton",
                   command=self._cancel).pack(side="right")
        ttk.Button(act, text="就用这些", style="Primary.TButton",
                   command=self._ok).pack(side="right", padx=(0, 6))

        self.txt.bind("<KeyRelease>", lambda e: self._update())
        self.bind("<Escape>", lambda e: self._cancel())
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self._center(master)
        self.grab_set()
        self.txt.focus_set()

    def _center(self, master) -> None:
        self.update_idletasks()
        w = max(420, self.winfo_reqwidth())
        h = max(360, self.winfo_reqheight())
        try:
            px, py = master.winfo_rootx(), master.winfo_rooty()
            pw, ph = master.winfo_width(), master.winfo_height()
        except tk.TclError:
            px = py = 0
            pw, ph = w, h
        self.geometry(f"{w}x{h}+{px + max(0, (pw - w) // 2)}+{py + max(0, (ph - h) // 3)}")

    def _names(self) -> list[str]:
        out, seen = [], set()
        for line in self.txt.get("1.0", "end-1c").replace("\r", "\n").split("\n"):
            s = villages.strip_ordinal(line.strip())
            if not s:
                continue
            k = villages.normalize(s)
            if k and k not in seen:
                seen.add(k)
                out.append(s)
        return out

    def _update(self) -> None:
        self.count.set(f"共 {len(self._names())} 行")

    def _ok(self) -> None:
        names = self._names()
        if not names:
            messagebox.showinfo("核对名称", "一个都没有，先在框里填几个。", parent=self)
            return
        self.result = names
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()
