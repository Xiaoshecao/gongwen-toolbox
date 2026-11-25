# -*- coding: utf-8 -*-
"""格式转换页（Qt 版）：文件列表 + 目标格式 + 输出目录 + 后台转换。

内核直接复用 gwtyper.core.convert（含纯 Python PDF 引擎、横竖版、留白）。
"""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QThread, Qt, Signal
from ui.controls import DpiPicker
from PySide6.QtWidgets import (QComboBox, QFileDialog, QFrame, QHBoxLayout,
                               QLabel, QLineEdit, QListWidget, QListWidgetItem,
                               QMessageBox, QProgressBar, QPushButton,
                               QSpinBox, QVBoxLayout, QWidget)

from gwtyper.core import formats as fmt
from gwtyper.core import sysenv
from gwtyper.core.convert import Options, convert


class ConvertWorker(QThread):
    progress = Signal(int, int, str)
    one_done = Signal(str, object)
    one_failed = Signal(str, str)
    finished_all = Signal(int, int)

    def __init__(self, files, target, out_dir, opt, parent=None):
        super().__init__(parent)
        self.files = files
        self.target = target
        self.out_dir = out_dir
        self.opt = opt

    def run(self):
        ok, bad = 0, 0
        n = len(self.files)
        for i, f in enumerate(self.files, 1):
            self.progress.emit(i, n, Path(f).name)
            try:
                made = convert(f, self.target, self.out_dir, self.opt)
                ok += 1
                self.one_done.emit(f, made)
            except Exception as exc:  # noqa: BLE001
                bad += 1
                self.one_failed.emit(f, str(exc))
        self.finished_all.emit(ok, bad)


class ConvertPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.files: list[Path] = []
        self.worker: ConvertWorker | None = None

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        bar = QHBoxLayout()
        t = QLabel("格式转换")
        t.setObjectName("H1")
        bar.addWidget(t)
        bar.addSpacing(12)
        hint = QLabel("Word / Excel / PPT / 图片 / PDF 互转；PDF 由内置引擎生成，不依赖办公软件")
        hint.setObjectName("Muted")
        bar.addWidget(hint)
        bar.addStretch(1)
        self.btn_add = QPushButton("添加文件…")
        self.btn_dir = QPushButton("添加文件夹…")
        self.btn_clear = QPushButton("清空")
        for b in (self.btn_add, self.btn_dir, self.btn_clear):
            bar.addWidget(b)
        root.addLayout(bar)

        body = QHBoxLayout()
        body.setSpacing(10)
        # 左：文件
        left = QFrame()
        left.setObjectName("Card")
        lv = QVBoxLayout(left)
        lv.setContentsMargins(12, 12, 12, 12)
        h = QLabel("要转换的文件")
        h.setObjectName("H2")
        lv.addWidget(h)
        self.list = QListWidget()
        lv.addWidget(self.list, 1)
        self.lbl_cnt = QLabel("0 个文件")
        self.lbl_cnt.setObjectName("Hint")
        lv.addWidget(self.lbl_cnt)
        body.addWidget(left, 2)

        # 右：设置
        right = QFrame()
        right.setObjectName("Card")
        rv = QVBoxLayout(right)
        rv.setContentsMargins(14, 12, 14, 12)
        rv.setSpacing(8)
        h2 = QLabel("转成什么")
        h2.setObjectName("H2")
        rv.addWidget(h2)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("目标格式"))
        self.combo = QComboBox()
        self.combo.setMinimumWidth(220)
        row1.addWidget(self.combo, 1)
        rv.addLayout(row1)

        self.lbl_note = QLabel("")
        self.lbl_note.setObjectName("Hint")
        self.lbl_note.setWordWrap(True)
        rv.addWidget(self.lbl_note)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("保存到"))
        self.out_edit = QLineEdit(str(Path(sysenv.export_dir()) / "转换结果"))
        row2.addWidget(self.out_edit, 1)
        self.btn_browse = QPushButton("浏览…")
        row2.addWidget(self.btn_browse)
        rv.addLayout(row2)

        row3 = QHBoxLayout()
        row3.addWidget(QLabel("图片清晰度"))
        self.dpi = DpiPicker(150, 72, 600)
        row3.addWidget(self.dpi)
        row3.addStretch(1)
        rv.addLayout(row3)

        rv.addStretch(1)
        self.pb = QProgressBar()
        self.pb.setVisible(False)
        rv.addWidget(self.pb)
        self.lbl_status = QLabel("先添加文件")
        self.lbl_status.setObjectName("Hint")
        self.lbl_status.setWordWrap(True)
        rv.addWidget(self.lbl_status)
        self.btn_go = QPushButton("开始转换")
        self.btn_go.setObjectName("Primary")
        self.btn_go.setEnabled(False)
        rv.addWidget(self.btn_go)
        right.setMinimumWidth(400)
        body.addWidget(right, 1)
        root.addLayout(body, 1)

        self.btn_add.clicked.connect(self.add_files)
        self.btn_dir.clicked.connect(self.add_folder)
        self.btn_clear.clicked.connect(self.clear)
        self.btn_browse.clicked.connect(self.pick_out)
        self.btn_go.clicked.connect(self.run_convert)
        self.combo.currentIndexChanged.connect(self._sync_note)
        self._refresh()

    # ---- 文件 ----
    def add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "选择文件", "", "所有文件 (*)")
        self._add([Path(p) for p in paths])

    def add_folder(self):
        d = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if not d:
            return
        exts = set()
        for t in fmt.TARGETS.values():
            exts |= set(t["sources"])
        self._add([p for p in Path(d).rglob("*") if p.suffix.lower() in exts])

    def _add(self, paths):
        for p in paths:
            if p.is_file() and p not in self.files:
                self.files.append(p)
        self._refresh()

    def clear(self):
        self.files.clear()
        self._refresh()

    def set_files(self, paths):
        self._add([Path(p) for p in paths])

    def _refresh(self):
        self.list.clear()
        for p in self.files:
            it = QListWidgetItem(f"{p.name}　·　{fmt.kind_of(str(p))}")
            it.setToolTip(str(p))
            self.list.addItem(it)
        self.lbl_cnt.setText(f"{len(self.files)} 个文件")
        self.btn_go.setEnabled(bool(self.files))
        # 目标格式：取所有文件的交集
        cur = self.combo.currentData()
        self.combo.blockSignals(True)
        self.combo.clear()
        if not self.files:
            self.combo.addItem("先添加文件", None)
            self.combo.setEnabled(False)
        else:
            self.combo.setEnabled(True)
            targets = fmt.targets_for([str(p) for p in self.files])
            for key in targets:
                self.combo.addItem(fmt.TARGET_CN.get(key, key), key)
            idx = self.combo.findData(cur) if cur else 0
            self.combo.setCurrentIndex(max(0, idx))
        self.combo.blockSignals(False)
        self._sync_note()

    def _sync_note(self):
        key = self.combo.currentData()
        if not self.files:
            self.lbl_note.setText("文档、表格、演示、图片都能互转；PDF 由内置引擎生成，不依赖办公软件。")
            return
        self.lbl_note.setText(fmt.TARGETS.get(key, {}).get("note", "") if key else "")

    def pick_out(self):
        d = QFileDialog.getExistingDirectory(self, "选择保存位置")
        if d:
            self.out_edit.setText(d)

    # ---- 转换 ----
    def run_convert(self):
        if self.worker and self.worker.isRunning():
            return
        key = self.combo.currentData()
        out = self.out_edit.text().strip()
        if not key or not out:
            QMessageBox.information(self, "还差一点", "先选目标格式和保存位置。")
            return
        opt = Options(dpi=self.dpi.value())
        self.pb.setVisible(True)
        self.pb.setRange(0, len(self.files))
        self.pb.setValue(0)
        self.btn_go.setEnabled(False)
        self.lbl_status.setText("正在转换…")
        self.worker = ConvertWorker([str(p) for p in self.files], key, out, opt, self)
        self.worker.progress.connect(lambda i, n, name: (self.pb.setValue(i - 1),
                                                        self.lbl_status.setText(f"[{i}/{n}] {name}")))
        self.worker.one_failed.connect(lambda f, msg: self._fail(f, msg))
        self.worker.finished_all.connect(self._done)
        self._fails: list[str] = []
        self.worker.start()

    def _fail(self, f, msg):
        self._fails.append(f"{Path(f).name}：{msg}")

    def _done(self, ok, bad):
        self.pb.setValue(self.pb.maximum())
        self.btn_go.setEnabled(True)
        lock = f"完成 {ok} 个" + (f"，失败 {bad} 个" if bad else "")
        self.lbl_status.setText(lock)
        if bad:
            QMessageBox.warning(self, "部分失败", "\n".join(self._fails[:8]))
