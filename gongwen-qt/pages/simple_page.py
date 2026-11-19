# -*- coding: utf-8 -*-
"""尚未迁移的页面占位：说明这一页将来做什么，避免空白。"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFrame, QLabel, QVBoxLayout, QWidget


class SimplePage(QWidget):
    def __init__(self, name: str, desc: str = "", parent=None):
        super().__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        title = QLabel(name)
        title.setObjectName("H1")
        root.addWidget(title)

        card = QFrame()
        card.setObjectName("Card")
        cv = QVBoxLayout(card)
        cv.setContentsMargins(16, 16, 16, 16)
        cv.setSpacing(8)
        h = QLabel("这一页还在迁移中")
        h.setObjectName("H2")
        cv.addWidget(h)
        sub = QLabel(desc or "功能逻辑沿用原来那套内核，界面换成 Qt 后逐步接上。")
        sub.setObjectName("Muted")
        sub.setWordWrap(True)
        cv.addWidget(sub)
        cv.addStretch(1)
        root.addWidget(card, 1)
