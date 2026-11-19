# -*- coding: utf-8 -*-
"""通用后台工作线程：跑一个函数，回报进度/结果/错误。"""
from __future__ import annotations

from PySide6.QtCore import QThread, Signal


class FuncWorker(QThread):
    progress = Signal(int, int, str)
    done = Signal(object)
    failed = Signal(str)

    def __init__(self, fn, parent=None):
        super().__init__(parent)
        self._fn = fn

    def run(self):
        try:
            self.done.emit(self._fn(self.progress.emit))
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(str(exc))


class BatchWorker(QThread):
    """逐份处理：fn(src) -> 说明；out_dir 由调用方在 fn 里决定。"""

    progress = Signal(int, int, str)
    done = Signal(int, list)
    failed = Signal(str)

    def __init__(self, files, fn, parent=None):
        super().__init__(parent)
        self.files, self._fn = list(files), fn

    def run(self):
        ok, bad = 0, []
        n = len(self.files)
        for i, f in enumerate(self.files, 1):
            self.progress.emit(i, n, str(f).rsplit("\\", 1)[-1])
            try:
                self._fn(f)
                ok += 1
            except Exception as exc:  # noqa: BLE001
                bad.append(f"{str(f).rsplit(chr(92), 1)[-1]}：{exc}")
        self.done.emit(ok, bad)
