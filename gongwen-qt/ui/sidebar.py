# -*- coding: utf-8 -*-
"""左侧竖排菜单：分组 + 条目；底部常驻「关于」。"""
from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (QButtonGroup, QFrame, QLabel, QPushButton,
                               QScrollArea, QVBoxLayout, QWidget)

from . import theme


class Sidebar(QFrame):
    currentChanged = Signal(int)
    aboutClicked = Signal()

    def __init__(self, spec: list[tuple[str, list[str]]], parent=None):
        super().__init__(parent)
        self.setObjectName("Sidebar")
        self.setFixedWidth(198)
        self._names: list[str] = []
        self._group = QButtonGroup(self)
        self._group.setExclusive(True)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.viewport().setAutoFillBackground(False)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        body = QWidget()
        body.setObjectName("NavBody")
        lay = QVBoxLayout(body)
        lay.setContentsMargins(8, 8, 8, 8)
        lay.setSpacing(2)

        idx = 0
        for group, names in spec:
            if group:
                g = QLabel(group)
                g.setObjectName("NavGroup")
                lay.addWidget(g)
            for name in names:
                b = QPushButton(name)
                b.setObjectName("NavItem")
                b.setCheckable(True)
                b.setCursor(Qt.PointingHandCursor)
                b.clicked.connect(lambda _=False, i=idx: self._pick(i))
                self._group.addButton(b, idx)
                lay.addWidget(b)
                self._names.append(name)
                idx += 1
        lay.addStretch(1)
        scroll.setWidget(body)
        outer.addWidget(scroll, 1)

        bottom = QFrame()
        bottom.setObjectName("SidebarBottom")
        bl = QVBoxLayout(bottom)
        bl.setContentsMargins(8, 6, 8, 10)
        bl.setSpacing(0)
        self.btn_about = QPushButton("关于")
        self.btn_about.setObjectName("NavAbout")
        self.btn_about.setCursor(Qt.PointingHandCursor)
        self.btn_about.clicked.connect(self.aboutClicked.emit)
        bl.addWidget(self.btn_about)
        outer.addWidget(bottom)

    # ---- API ----
    def names(self) -> list[str]:
        return list(self._names)

    def select(self, key) -> None:
        if isinstance(key, str):
            key = self._names.index(key)
        btn = self._group.button(int(key))
        if btn is not None:
            btn.setChecked(True)
            self.currentChanged.emit(int(key))

    def _pick(self, i: int) -> None:
        self.currentChanged.emit(i)
