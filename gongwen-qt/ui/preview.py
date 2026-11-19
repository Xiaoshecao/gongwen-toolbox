# -*- coding: utf-8 -*-
"""版面预览：把 docx / pdf 渲染成图片，翻页看。"""
from __future__ import annotations

import hashlib
import os
from pathlib import Path

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (QHBoxLayout, QLabel, QPushButton, QScrollArea,
                               QVBoxLayout, QWidget)

from PySide6.QtCore import QSize


def _cache_dir() -> Path:
    from gwtyper.core import sysenv
    d = Path(sysenv.cache_dir("qt-preview"))
    d.mkdir(parents=True, exist_ok=True)
    return d


def to_pdf(src: Path) -> Path | None:
    """把 docx/pdf 变成 pdf（纯 Python，不依赖办公软件）。"""
    from gwtyper.core import pdfwrite
    src = Path(src)
    if src.suffix.lower() == ".pdf":
        return src
    key = hashlib.sha1(f"{src}|{src.stat().st_mtime_ns}".encode()).hexdigest()[:12]
    dst = _cache_dir() / f"{src.stem}-{key}.pdf"
    if dst.exists() and dst.stat().st_size > 0:
        return dst
    try:
        pdfwrite.docx_to_pdf(str(src), str(dst))
    except Exception:  # noqa: BLE001
        return None
    return dst if dst.exists() and dst.stat().st_size > 0 else None


class PdfPreview(QWidget):
    """一个能翻页的 PDF 预览面板。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._doc = None
        self._page = 0
        self._pdf: Path | None = None
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)

        bar = QHBoxLayout()
        self.btn_prev = QPushButton("上一页")
        self.btn_next = QPushButton("下一页")
        self.lbl = QLabel("还没有可预览的内容")
        self.lbl.setObjectName("Hint")
        bar.addWidget(self.btn_prev)
        bar.addWidget(self.btn_next)
        bar.addWidget(self.lbl, 1)
        lay.addLayout(bar)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        self.view = QLabel("")
        self.view.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        self.scroll.setWidget(self.view)
        lay.addWidget(self.scroll, 1)

        self.btn_prev.clicked.connect(lambda: self._go(-1))
        self.btn_next.clicked.connect(lambda: self._go(1))
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.setInterval(120)
        self._timer.timeout.connect(self._render)

    # ---------------- API ----------------
    def clear(self, msg="还没有可预览的内容"):
        self._doc = None
        self._pdf = None
        self._page = 0
        self.view.setPixmap(QPixmap())
        self.view.setText("")
        self.lbl.setText(msg)

    def show_file(self, path) -> bool:
        pdf = to_pdf(Path(path))
        if not pdf:
            self.clear("这个文件预览不了（可能不是 Word/PDF）")
            return False
        return self.show_pdf(pdf)

    def show_pdf(self, path) -> bool:
        import fitz
        try:
            if self._doc:
                self._doc.close()
            self._doc = fitz.open(str(path))
            if self._doc.page_count == 0:
                self.clear("空白文档")
                return False
            self._pdf = Path(path)
            self._page = 0
            self._render()
            return True
        except Exception:  # noqa: BLE001
            self.clear("打不开这个文件")
            return False

    # ---------------- 内部 ----------------
    def _go(self, delta):
        if not self._doc:
            return
        p = self._page + delta
        if 0 <= p < self._doc.page_count:
            self._page = p
            self._render()

    def resizeEvent(self, ev):
        super().resizeEvent(ev)
        self._timer.start()

    def _render(self):
        if not self._doc:
            return
        try:
            page = self._doc.load_page(self._page)
            avail = max(320, self.scroll.viewport().width() - 24)
            zoom = avail / page.rect.width
            zoom = max(0.2, min(zoom, 3.0))
            pix = page.get_pixmap(matrix=__import__("fitz").Matrix(zoom, zoom), alpha=False)
            img = QPixmap()
            img.loadFromData(pix.tobytes("png"))
            self.view.setPixmap(img)
            self.view.setMinimumSize(img.size())
            self.lbl.setText(f"第 {self._page + 1} / {self._doc.page_count} 页　{Path(self._pdf).name}")
        except Exception:  # noqa: BLE001
            self.lbl.setText("这一页渲染失败")
