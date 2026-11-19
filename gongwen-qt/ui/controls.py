# -*- coding: utf-8 -*-
"""常用小控件：DPI 选择（固定档位 + 自定义）。"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QComboBox, QHBoxLayout, QSpinBox, QWidget

PRESETS = (96, 150, 200, 300, 600)


def widen_popup(combo: QComboBox, extra: int = 40, floor: int = 170) -> None:
    """让下拉列表按最长条目撑开（否则「150 DPI」会被省略成「15…PI」）。"""
    try:
        view = combo.view()
        need = view.sizeHintForColumn(0) + extra
        view.setMinimumWidth(max(floor, need))
    except Exception:
        pass


class DpiPicker(QWidget):
    """下拉选常用档位，最后一档「自定义…」时出现数字框。"""

    def __init__(self, value: int = 150, minimum: int = 72, maximum: int = 600,
                 parent=None):
        super().__init__(parent)
        self._min, self._max = minimum, maximum
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)

        self.combo = QComboBox()
        for v in PRESETS:
            self.combo.addItem(f"{v} DPI", v)
        self.combo.addItem("自定义", None)
        lay.addWidget(self.combo)
        widen_popup(self.combo)

        self.spin = QSpinBox()
        self.spin.setRange(minimum, maximum)
        self.spin.setValue(int(value))
        self.spin.setSuffix(" DPI")
        self.spin.setVisible(False)
        lay.addWidget(self.spin)

        self.combo.currentIndexChanged.connect(self._sync)
        self.set_value(value)

    def _sync(self):
        custom = self.combo.currentData() is None
        self.spin.setVisible(custom)
        if not custom:
            self.spin.setValue(int(self.combo.currentData()))

    def value(self) -> int:
        data = self.combo.currentData()
        return int(self.spin.value()) if data is None else int(data)

    def set_value(self, value: int) -> None:
        value = max(self._min, min(self._max, int(value)))
        idx = self.combo.findData(value)
        if idx >= 0:
            self.combo.setCurrentIndex(idx)
            self.spin.setValue(value)
            self.spin.setVisible(False)
        else:
            self.combo.setCurrentIndex(self.combo.count() - 1)  # 自定义
            self.spin.setValue(value)
            self.spin.setVisible(True)
