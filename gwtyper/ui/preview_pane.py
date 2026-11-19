# -*- coding: utf-8 -*-
"""页面预览控件：把渲染好的 PNG 字节按页竖排到一块可滚动的画布上。"""
from __future__ import annotations

import io
import tkinter as tk
from tkinter import ttk

try:
    from PIL import Image, ImageTk

    _HAS_PIL = True
except Exception:  # pragma: no cover
    _HAS_PIL = False

from . import theme

GAP = 14        # 页间距
MARGIN = 12     # 画布内边距


class PreviewPane(ttk.Frame):
    """只负责显示，转 PDF / 渲染在 app 的后台线程里做。"""

    def __init__(self, master, width_hint: int = 420, **kw):
        super().__init__(master, **kw)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        head = ttk.Frame(self)
        head.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        ttk.Label(head, text="版面预览", style="H2.TLabel").pack(side="left")
        self.info = ttk.Label(head, text="", style="Muted.TLabel")
        self.info.pack(side="left", padx=(8, 0))
        self.head = head

        self.canvas = tk.Canvas(self, background=theme.INSET, highlightthickness=0,
                                width=width_hint)
        self.canvas.grid(row=1, column=0, sticky="nsew")
        sb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        sb.grid(row=1, column=1, sticky="ns")
        self.sb = sb
        self.canvas.configure(yscrollcommand=sb.set)

        self._pages: list[bytes] = []
        self._photos: list[object] = []
        self._items: list[int] = []
        self._layout_job = None
        self._msg = "导入文件后显示这里"

        self.canvas.bind("<Configure>", self._on_resize)
        self._bind_wheel()

    # ------------------------------------------------------------ 滚轮
    def _bind_wheel(self) -> None:
        def on_wheel(event):
            delta = -1 if getattr(event, "delta", 0) > 0 else 1
            self.canvas.yview_scroll(delta * 3, "units")

        def enter(_e):
            self.canvas.bind_all("<MouseWheel>", on_wheel)

        def leave(_e):
            self.canvas.unbind_all("<MouseWheel>")

        self.canvas.bind("<Enter>", enter)
        self.canvas.bind("<Leave>", leave)

    def _on_resize(self, _e=None) -> None:
        if self._layout_job:
            self.after_cancel(self._layout_job)
        self._layout_job = self.after(180, self._redraw)

    # ------------------------------------------------------------ 接口
    def hide_head(self) -> None:
        """外部已有标题栏时，去掉自带那行，把画布提到第一行。"""
        self.head.grid_remove()
        self.canvas.grid_configure(row=0)
        self.sb.grid_configure(row=0)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)

    def set_pages(self, pages: list[bytes], note: str = "") -> None:
        self._pages = list(pages)
        self._msg = ""
        self.info.configure(text=note or (f"{len(pages)} 页" if pages else ""))
        self._redraw()

    def clear(self, msg: str = "") -> None:
        self._pages = []
        self._msg = msg or ""
        self.info.configure(text="")
        self._redraw()

    def set_info(self, text: str) -> None:
        self.info.configure(text=text)

    # ------------------------------------------------------------ 绘制
    def _free_photos(self) -> None:
        """显式释放上一批 PhotoImage。

        Windows 上 Tk 的 PhotoImage 在后台持有 GDI 位图（CreateDIBSection）。
        只 `self._photos.clear()` 不会立刻回收——引用一断，回收时机由 GC 决定，
        而预览是「每次重绘整份重做」，翻几次就攒出几百张位图，
        最后弹 `Tk_GetPixmap: Error from CreateDIBSSection —— 内存资源不足`，
        GDI 句柄耗尽，整个界面画不出来。
        这里主动置 None 并调 delete()，让 Tk 立刻释放底层资源。
        """
        for ph in self._photos:
            try:
                if hasattr(ph, "delete"):
                    ph.delete()       # PIL ImageTk.PhotoImage 有 delete()
            except Exception:  # noqa: BLE001
                pass
        self._photos.clear()

    def _redraw(self) -> None:
        self._layout_job = None
        self.canvas.delete("all")
        self._free_photos()
        self._items.clear()

        w = max(200, self.canvas.winfo_width() - MARGIN * 2)
        if not self._pages:
            self.canvas.create_text(
                self.canvas.winfo_width() // 2 or 200, 40, text=self._msg,
                fill=theme.TEXT2, font=(self._font(), 10),
                width=max(120, w))          # 加宽度约束，长句两边才不会被切
            self.canvas.configure(scrollregion=(0, 0, w, 120))
            return

        y = MARGIN
        for raw in self._pages:
            photo, h = self._to_photo(raw, w)
            if photo is None:
                continue
            self._photos.append(photo)
            self._items.append(self.canvas.create_image(w // 2 + MARGIN, y, anchor="n", image=photo))
            y += h + GAP
        self.canvas.configure(scrollregion=(0, 0, w + MARGIN * 2, y + MARGIN))

    def _to_photo(self, raw: bytes, target_w: int):
        if not _HAS_PIL:
            return None, 0
        img = None
        try:
            img = Image.open(io.BytesIO(raw))
            img.load()
            if img.width > target_w:
                ratio = target_w / img.width
                # 先算出目标尺寸再 resize，然后立刻丢掉原图——
                # 一张 A4@150dpi 是 1240×1754，不释放的话几十页就把内存吃光
                small = img.resize(
                    (target_w, max(1, int(img.height * ratio))), Image.LANCZOS)
                img.close()
                img = small
            photo = ImageTk.PhotoImage(img)
            return photo, img.height
        except Exception:  # noqa: BLE001
            return None, 0
        finally:
            if img is not None:
                try:
                    img.close()
                except Exception:  # noqa: BLE001
                    pass

    def _font(self) -> str:
        try:
            import tkinter.font as tkfont

            return tkfont.nametofont("TkDefaultFont").actual("family")
        except Exception:  # noqa: BLE001
            return "Microsoft YaHei UI"
