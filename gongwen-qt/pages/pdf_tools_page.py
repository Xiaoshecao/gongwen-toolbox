# -*- coding: utf-8 -*-
"""PDF 工具箱（Qt 版）：加密 / 去密码 / 瘦身 / 按页提取 / Word 带书签导 PDF。"""
from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (QComboBox, QFileDialog, QGridLayout, QHBoxLayout,
                               QLabel, QLineEdit, QMessageBox, QProgressBar,
                               QPushButton, QVBoxLayout, QWidget)

from gwtyper.core import office, sysenv
from gwtyper.core import pdf_engine as pe
from ui.filepanel import FilePanel, accept_suffix, card, hint
from pages.workers import BatchWorker
from pages.tools_pages import _page_head

OPS = [
    ("encrypt", "加密（设个密码）"),
    ("decrypt", "去掉密码"),
    ("compress", "瘦身（压缩体积）"),
    ("extract", "按页提取"),
    ("bookmark", "Word 导 PDF（带书签）"),
]


class PdfToolsPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)
        root.addLayout(_page_head("PDF 工具箱", "加密、去密码、瘦身、抽页；Word 导 PDF 可带标题书签"))

        body = QHBoxLayout()
        body.setSpacing(10)
        left, ll = card()
        self.panel = FilePanel("要处理的文件", accept=accept_suffix({".pdf", ".docx", ".doc"}))
        ll.addWidget(self.panel)
        body.addWidget(left, 2)

        right, rl = card("做什么")
        self.op = QComboBox()
        for k, label in OPS:
            self.op.addItem(label, k)
        rl.addWidget(self.op)

        grid = QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(8)
        self.password = QLineEdit()
        self.password.setEchoMode(QLineEdit.Password)
        self.password.setPlaceholderText("密码")
        self.pages = QLineEdit()
        self.pages.setPlaceholderText("例如 1-3,5（不填就是全部）")
        grid.addWidget(QLabel("密码"), 0, 0)
        grid.addWidget(self.password, 0, 1)
        grid.addWidget(QLabel("页码"), 1, 0)
        grid.addWidget(self.pages, 1, 1)
        rl.addLayout(grid)

        row = QHBoxLayout()
        row.addWidget(QLabel("保存到"))
        self.out = QLineEdit(str(Path(sysenv.desktop_dir()) / "PDF 处理结果"))
        row.addWidget(self.out, 1)
        btn = QPushButton("浏览…")
        row.addWidget(btn)
        rl.addLayout(row)

        def browse():
            d = QFileDialog.getExistingDirectory(None, "选择保存位置")
            if d:
                self.out.setText(d)
        btn.clicked.connect(browse)

        self.pb = QProgressBar()
        self.pb.setVisible(False)
        rl.addWidget(self.pb)
        self.status = hint("先添加文件")
        rl.addWidget(self.status)
        self.btn_go = QPushButton("开始处理")
        self.btn_go.setObjectName("Primary")
        rl.addWidget(self.btn_go)
        rl.addWidget(hint("注：Word 导 PDF 带书签要用本机 WPS / Office；其他操作不依赖办公软件。"))
        rl.addStretch(1)
        body.addWidget(right, 3)
        root.addLayout(body, 1)

        self.op.currentIndexChanged.connect(self._sync)
        self.btn_go.clicked.connect(self.run)
        self.worker = None
        self._sync()

    def _sync(self):
        k = self.op.currentData()
        self.password.setEnabled(k in ("encrypt", "decrypt"))
        self.pages.setEnabled(k == "extract")

    def run(self):
        if self.worker and self.worker.isRunning():
            return
        if not self.panel.files:
            QMessageBox.information(self, "还没有文件", "先添加文件。")
            return
        k = self.op.currentData()
        out = Path(self.out.text().strip())
        out.mkdir(parents=True, exist_ok=True)
        pwd = self.password.text()
        spec = self.pages.text().strip()
        if k in ("encrypt", "decrypt") and not pwd:
            QMessageBox.information(self, "还差一点", "填个密码。")
            return

        def one(src):
            src = Path(src)
            dst = str(out / (src.stem + {"encrypt": "-加密", "decrypt": "-去密码",
                                         "compress": "-瘦身", "extract": "-抽页",
                                         "bookmark": ""}[k] + ".pdf"))
            if k == "encrypt":
                pe.encrypt(str(src), dst, pwd)
            elif k == "decrypt":
                pe.decrypt(str(src), dst, pwd)
            elif k == "compress":
                pe.compress(str(src), dst)
            elif k == "extract":
                total = pe.pdf_page_count(str(src))
                pe.extract_pages(str(src), dst, spec or f"1-{total}")
            elif k == "bookmark":
                office.word_to_pdf_bookmarked(str(src), dst)

        self.pb.setVisible(True)
        self.pb.setRange(0, len(self.panel.files))
        self.btn_go.setEnabled(False)
        self.worker = BatchWorker([str(p) for p in self.panel.files], one, self)
        self.worker.progress.connect(lambda i, n, name: (self.pb.setValue(i - 1),
                                                         self.status.setText(f"[{i}/{n}] {name}")))
        self.worker.done.connect(self._done)
        self.worker.failed.connect(lambda m: QMessageBox.warning(self, "处理失败", m))
        self.worker.start()

    def _done(self, ok, bad):
        self.pb.setValue(self.pb.maximum())
        self.btn_go.setEnabled(True)
        self.status.setText(f"处理 {ok} 份" + (f"，失败 {len(bad)} 份" if bad else ""))
        if bad:
            QMessageBox.warning(self, "部分失败", "\n".join(bad[:8]))
        else:
            QMessageBox.information(self, "完成", f"好了 {ok} 份，在：{self.out.text()}")
