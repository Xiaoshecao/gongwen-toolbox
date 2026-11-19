# -*- coding: utf-8 -*-
"""通用小部件：卡片 + 可拖拽的文件列表面板。"""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (QAbstractItemView, QFileDialog, QFrame,
                               QHBoxLayout, QLabel, QListWidget,
                               QListWidgetItem, QPushButton, QVBoxLayout,
                               QWidget)


def card(title: str | None = None) -> tuple[QFrame, QVBoxLayout]:
    f = QFrame()
    f.setObjectName("Card")
    lay = QVBoxLayout(f)
    lay.setContentsMargins(14, 12, 14, 12)
    lay.setSpacing(8)
    if title:
        h = QLabel(title)
        h.setObjectName("H2")
        lay.addWidget(h)
    return f, lay


def hint(text: str = "") -> QLabel:
    lb = QLabel(text)
    lb.setObjectName("Hint")
    lb.setWordWrap(True)
    return lb


class FilePanel(QWidget):
    """左栏：文件列表 + 添加文件/文件夹/清空。支持拖入。"""

    changed = Signal()

    def __init__(self, title="文件", accept=None, multiple=True, parent=None):
        super().__init__(parent)
        self.files: list[Path] = []
        self._accept = accept  # 可调用：path -> bool
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)

        bar = QHBoxLayout()
        h = QLabel(title)
        h.setObjectName("H2")
        bar.addWidget(h)
        bar.addStretch(1)
        self.btn_add = QPushButton("添加文件…")
        self.btn_dir = QPushButton("添加文件夹…")
        self.btn_clear = QPushButton("清空")
        for b in (self.btn_add, self.btn_dir, self.btn_clear):
            bar.addWidget(b)
        lay.addLayout(bar)

        self.list = QListWidget()
        self.list.setSelectionMode(QAbstractItemView.ExtendedSelection if multiple
                                   else QAbstractItemView.SingleSelection)
        self.list.setAcceptDrops(True)
        self.list.setDragDropMode(QAbstractItemView.DropOnly)
        self.list.viewport().setAcceptDrops(True)
        lay.addWidget(self.list, 1)
        self.lbl = hint("0 个文件")
        lay.addWidget(self.lbl)

        self.btn_add.clicked.connect(self.add_files)
        self.btn_dir.clicked.connect(self.add_folder)
        self.btn_clear.clicked.connect(self.clear)

    # --- 拖放 ---
    def dragEnterEvent(self, ev):
        if ev.mimeData().hasUrls():
            ev.acceptProposedAction()

    def dropEvent(self, ev):
        paths = [Path(u.toLocalFile()) for u in ev.mimeData().urls() if u.isLocalFile()]
        self._add(paths)
        ev.acceptProposedAction()

    # --- 操作 ---
    def add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "选择文件", "", "所有文件 (*)")
        self._add([Path(p) for p in paths])

    def add_folder(self):
        d = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if d:
            self._add(sorted(p for p in Path(d).rglob("*") if p.is_file()))

    def _add(self, paths):
        for p in paths:
            try:
                if not p.is_file():
                    continue
                if self._accept and not self._accept(p):
                    continue
                if p not in self.files:
                    self.files.append(p)
            except OSError:
                continue
        self.refresh()

    def set_files(self, paths):
        self._add([Path(p) for p in paths])

    def clear(self):
        self.files.clear()
        self.refresh()

    def selected_or_all(self) -> list[Path]:
        sel = [self.list.row(i) for i in self.list.selectedItems()]
        return [self.files[i] for i in sorted(sel)] if sel else list(self.files)

    def refresh(self):
        self.list.clear()
        for p in self.files:
            it = QListWidgetItem(p.name)
            it.setToolTip(str(p))
            self.list.addItem(it)
        self.lbl.setText(f"{len(self.files)} 个文件")
        self.changed.emit()


def accept_suffix(exts) -> "callable":
    exts = {e.lower() for e in exts}

    def _ok(p: Path) -> bool:
        return p.suffix.lower() in exts
    return _ok
