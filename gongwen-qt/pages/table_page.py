# -*- coding: utf-8 -*-
"""表格排版：从一份样表学出版式（行高列宽/字体/页边距/边框），再套到别的表格上。"""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (
    QComboBox, QFileDialog, QHBoxLayout, QInputDialog, QLabel, QLineEdit,
    QMessageBox, QPushButton, QVBoxLayout, QWidget,
)

from gwtyper.core import table_style as TS

from pages.workers import BatchWorker  # noqa: F401  (统一消息泵)
from ui.filepanel import FilePanel, card, hint


class ApplyTableWorker(QThread):
    progress = Signal(int, int, str)
    done = Signal(list, list)

    def __init__(self, files: list[str], style: dict, out_dir: Path):
        super().__init__()
        self.files, self.style, self.out_dir = files, style, out_dir

    def run(self):
        ok, bad = [], []
        n = len(self.files)
        for i, f in enumerate(self.files, 1):
            self.progress.emit(i, n, Path(f).name)
            try:
                dst = self.out_dir / (Path(f).stem + ("-表格排版.xlsx" if Path(f).suffix.lower() in (".xlsx", ".xlsm") else "-表格排版.docx"))
                TS.apply_table(f, dst, self.style)
                ok.append(str(dst))
            except Exception as exc:  # noqa: BLE001
                bad.append((str(f), str(exc)))
        self.done.emit(ok, bad)


class TablePage(QWidget):
    def __init__(self):
        super().__init__()
        self.worker = None
        self._style: dict | None = None

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        # ---- 样式栏 ----
        c1, l1 = card()
        bar = QHBoxLayout()
        t = QLabel("表格排版")
        t.setObjectName("H1")
        bar.addWidget(t)
        bar.addSpacing(14)
        bar.addWidget(QLabel("表格样式"))
        self.combo = QComboBox()
        self.combo.setMinimumWidth(260)
        bar.addWidget(self.combo)
        self.btn_learn = QPushButton("从样表学习…")
        self.btn_learn.setToolTip("给一份已经排好的表格（Excel .xlsx 或 Word .docx），学出它的行高列宽、字体、边框和页边距")
        self.btn_del = QPushButton("删除样式")
        bar.addWidget(self.btn_learn)
        bar.addWidget(self.btn_del)
        bar.addStretch(1)
        l1.addLayout(bar)
        self.lbl_sum = QLabel("")
        self.lbl_sum.setObjectName("Hint")
        self.lbl_sum.setWordWrap(True)
        l1.addWidget(self.lbl_sum)
        root.addWidget(c1)
        root.addWidget(hint("先用「从样表学习…」学一份 Excel 样表；以后把要排的表格（Excel）拖进来，点下面按钮就套上同样的列宽行高、字体和页边距。"))

        # ---- 文件 ----
        self.panel = FilePanel("要排版的表格文件", accept=(".xlsx", ".xlsm", ".xls", ".et"))
        root.addWidget(self.panel, 1)

        # ---- 底部 ----
        c2, l2 = card()
        row = QHBoxLayout()
        row.addWidget(QLabel("输出到"))
        self.out = QLineEdit()
        self.out.setPlaceholderText("默认存到和原文件同一个文件夹")
        row.addWidget(self.out, 1)
        self.btn_pick = QPushButton("选择…")
        row.addWidget(self.btn_pick)
        l2.addLayout(row)
        row2 = QHBoxLayout()
        self.btn_go = QPushButton("按样式排版表格")
        self.btn_go.setObjectName("Primary")
        self.btn_go.setEnabled(False)
        self.btn_open = QPushButton("打开输出文件夹")
        row2.addWidget(self.btn_go)
        row2.addWidget(self.btn_open)
        row2.addStretch(1)
        self.status = QLabel("先学一个表格样式，再添加要排的文件")
        self.status.setObjectName("Hint")
        row2.addWidget(self.status)
        l2.addLayout(row2)
        root.addWidget(c2)

        self.btn_learn.clicked.connect(self.learn)
        self.btn_del.clicked.connect(self.delete_style)
        self.btn_pick.clicked.connect(self.pick_out)
        self.btn_go.clicked.connect(self.run)
        self.btn_open.clicked.connect(self.open_out)
        self.combo.currentIndexChanged.connect(self._on_pick)
        self.panel.changed.connect(self._sync)
        self._reload()

    # ---------- 样式 ----------
    def _reload(self, keep: str | None = None):
        self.combo.blockSignals(True)
        self.combo.clear()
        for st in TS.list_styles():
            self.combo.addItem(st.get("name") or "表格样式", st.get("_path"))
        self.combo.blockSignals(False)
        if keep:
            i = self.combo.findData(keep)
            if i >= 0:
                self.combo.setCurrentIndex(i)
        self._on_pick()

    def _on_pick(self, *_a):
        p = self.combo.currentData()
        self._style = None
        if not p:
            self.lbl_sum.setText("还没有表格样式——先点「从样表学习…」学一份。")
            self._sync()
            return
        try:
            self._style = TS.load_style(p)
            self.lbl_sum.setText(TS.describe(self._style))
        except Exception as exc:  # noqa: BLE001
            self.lbl_sum.setText(f"样式读不出来：{exc}")
        self._sync()

    def learn(self):
        f, _ = QFileDialog.getOpenFileName(self, "选一份样表", "", "Excel 表格 (*.xlsx *.xlsm);;Word 文档 (*.docx)")
        if not f:
            return
        try:
            st = TS.learn_table(f)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "学不了", str(exc))
            return
        name, ok = QInputDialog.getText(self, "给这个表格样式起个名字", "名字：",
                                        text=st.get("name") or "表格样式")
        if not ok:
            return
        try:
            p = TS.save_style(st, (name or "").strip() or None)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "存不了", str(exc))
            return
        self._reload(keep=str(p))
        QMessageBox.information(self, "学好了",
                                f"已保存表格样式「{self.combo.currentText()}」。\n\n{TS.describe(self._style or {})}")

    def delete_style(self):
        p = self.combo.currentData()
        if not p:
            return
        if QMessageBox.question(self, "删掉这个样式？", f"确定删除「{self.combo.currentText()}」吗？") != QMessageBox.Yes:
            return
        TS.delete_style(p)
        self._reload()

    # ---------- 排版 ----------
    def pick_out(self):
        d = QFileDialog.getExistingDirectory(self, "输出到哪个文件夹")
        if d:
            self.out.setText(d)

    def _sync(self):
        self.btn_go.setEnabled(bool(self._style) and bool(self.panel.files) and self.worker is None)

    def run(self):
        files = [str(p) for p in (self.panel.selected_or_all() or [])]
        if not files or not self._style:
            return
        out_dir = Path(self.out.text().strip()) if self.out.text().strip() else Path(files[0]).parent
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "这个文件夹用不了", str(exc))
            return
        self.worker = ApplyTableWorker(files, self._style, out_dir)
        self.worker.progress.connect(lambda i, n, name: self.status.setText(f"[{i}/{n}] {name}"))
        self.worker.done.connect(lambda ok, bad: self._done(ok, bad, out_dir))
        self.btn_go.setEnabled(False)
        self.status.setText("开始排版…")
        self.worker.start()

    def _done(self, ok, bad, out_dir: Path):
        self.worker = None
        self._sync()
        self.status.setText(f"完成 {len(ok)} 个，失败 {len(bad)} 个")
        msg = f"排版完成 {len(ok)} 个文件。\n输出目录：{out_dir}"
        if bad:
            msg += "\n\n失败：\n" + "\n".join(f"· {Path(f).name}：{e}" for f, e in bad[:10])
        QMessageBox.information(self, "表格排版", msg)

    def open_out(self):
        d = self.out.text().strip()
        if not d:
            files = self.panel.selected_or_all()
            if files:
                d = str(Path(str(files[0])).parent)
        if d:
            from PySide6.QtCore import QUrl
            from PySide6.QtGui import QDesktopServices
            QDesktopServices.openUrl(QUrl.fromLocalFile(d))
