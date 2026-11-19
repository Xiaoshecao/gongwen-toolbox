# -*- coding: utf-8 -*-
"""公文模板库（Qt 版）：通知/请示/报告/函/纪要/情况汇报/证明。"""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QFileDialog, QFormLayout, QHBoxLayout, QLabel,
                               QLineEdit, QListWidget, QListWidgetItem,
                               QMessageBox, QPlainTextEdit, QPushButton,
                               QVBoxLayout, QWidget)

from gwtyper.core import sysenv
from gwtyper.core import templates as T
from ui.filepanel import card, hint
from pages.tools_pages import _page_head


class TemplatePage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)
        root.addLayout(_page_head("公文模板库", "选一种，把空填上，直接出排好版的 Word"))

        body = QHBoxLayout()
        body.setSpacing(10)
        left, ll = card("模板")
        self.list = QListWidget()
        for key, tpl in T.TEMPLATES.items():
            it = QListWidgetItem(tpl.get("name", key))
            it.setData(1, key)
            self.list.addItem(it)
        ll.addWidget(self.list, 1)
        body.addWidget(left, 1)

        right, rl = card("填内容")
        self.lbl = hint("")
        rl.addWidget(self.lbl)
        self.form_host = QWidget()
        self.form = QFormLayout(self.form_host)
        self.form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.form.setVerticalSpacing(12)
        self.form.setContentsMargins(0, 0, 0, 0)
        rl.addWidget(self.form_host)
        rl.addStretch(1)
        self.btn_go = QPushButton("生成 Word…")
        self.btn_go.setObjectName("Primary")
        rl.addWidget(self.btn_go)
        self.status = hint("")
        rl.addWidget(self.status)
        body.addWidget(right, 2)
        root.addLayout(body, 1)

        self.list.currentRowChanged.connect(self._pick)
        self.btn_go.clicked.connect(self.build)
        self._widgets = {}
        self.list.setCurrentRow(0)

    def _pick(self, row):
        while self.form.rowCount():
            self.form.removeRow(0)
        self._widgets = {}
        item = self.list.item(row)
        if not item:
            return
        key = item.data(1)
        tpl = T.TEMPLATES[key]
        self._key = key
        self.lbl.setText(tpl.get("desc", ""))
        for f in tpl.get("fields", []):
            fkey, label, default, multi = f
            if multi:
                w = QPlainTextEdit()
                w.setPlainText(default or "")
                w.setFixedHeight(90)
            else:
                w = QLineEdit(default or "")
            self.form.addRow(label, w)
            self._widgets[fkey] = w

    def build(self):
        if not getattr(self, "_key", None):
            return
        values = {}
        for k, w in self._widgets.items():
            values[k] = w.toPlainText() if isinstance(w, QPlainTextEdit) else w.text()
        name = values.get("title") or T.TEMPLATES[self._key].get("name", "公文")
        dst, _ = QFileDialog.getSaveFileName(
            self, "保存", str(Path(sysenv.desktop_dir()) / f"{name}.docx"), "Word (*.docx)")
        if not dst:
            return
        try:
            T.build(self._key, dst, values)
            self.status.setText(f"已生成：{dst}")
            QMessageBox.information(self, "完成", f"生成好了：\n{dst}")
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "生成失败", str(exc))
