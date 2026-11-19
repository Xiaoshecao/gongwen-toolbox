# -*- coding: utf-8 -*-
"""材料上交页（Qt 版）：名单 + 台账 + 按文件名识别 + 导出未交清单。

内核复用 gwtyper.core.villages（名单记忆、状态记忆、文件名匹配）。
"""
from __future__ import annotations

import csv
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QComboBox, QFileDialog, QFrame, QHBoxLayout,
                               QInputDialog, QLabel, QMessageBox, QPlainTextEdit,
                               QPushButton, QTableWidget, QTableWidgetItem,
                               QVBoxLayout, QWidget)

from gwtyper.core import sysenv
from gwtyper.core import villages as V

COL_NAME, COL_STATE, COL_WHY = 0, 1, 2


class ChecklistPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.names: list[str] = []
        self.status: dict[str, dict] = {}

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        bar = QHBoxLayout()
        t = QLabel("材料上交")
        t.setObjectName("H1")
        bar.addWidget(t)
        bar.addSpacing(12)
        tip = QLabel("一行一个（村、社区、站所、科室、学校、公司都行，不必以「村」结尾）")
        tip.setObjectName("Muted")
        bar.addWidget(tip)
        bar.addStretch(1)
        self.lbl_sum = QLabel("")
        self.lbl_sum.setObjectName("Muted")
        bar.addWidget(self.lbl_sum)
        root.addLayout(bar)

        body = QHBoxLayout()
        body.setSpacing(10)

        # 左：名单
        left = QFrame()
        left.setObjectName("Card")
        lv = QVBoxLayout(left)
        lv.setContentsMargins(12, 12, 12, 12)
        lv.setSpacing(6)
        h = QLabel("名单")
        h.setObjectName("H2")
        lv.addWidget(h)

        pre = QHBoxLayout()
        self.preset = QComboBox()
        self.preset.setMinimumWidth(150)
        pre.addWidget(self.preset, 1)
        self.btn_load = QPushButton("载入")
        self.btn_save_preset = QPushButton("存为预设…")
        pre.addWidget(self.btn_load)
        pre.addWidget(self.btn_save_preset)
        lv.addLayout(pre)

        self.text = QPlainTextEdit()
        self.text.setPlaceholderText("一行一个，例如：\n古家兰村\n河南村\n多长村")
        lv.addWidget(self.text, 1)

        row = QHBoxLayout()
        self.btn_save = QPushButton("保存名单")
        self.btn_save.setObjectName("Primary")
        self.btn_excel = QPushButton("从 Excel 解析…")
        row.addWidget(self.btn_save)
        row.addWidget(self.btn_excel)
        lv.addLayout(row)
        body.addWidget(left, 1)

        # 右：台账
        right = QFrame()
        right.setObjectName("Card")
        rv = QVBoxLayout(right)
        rv.setContentsMargins(12, 12, 12, 12)
        rv.setSpacing(6)
        h2 = QLabel("上交情况")
        h2.setObjectName("H2")
        rv.addWidget(h2)
        hint = QLabel("双击一行可手动切换「已交 / 未交」；文件名里含这个名称就算它交了")
        hint.setObjectName("Hint")
        rv.addWidget(hint)

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["名称", "状态", "怎么认出来的"])
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setColumnWidth(0, 220)
        self.table.setColumnWidth(1, 90)
        rv.addWidget(self.table, 1)

        row2 = QHBoxLayout()
        self.btn_scan = QPushButton("选文件夹，按文件名认…")
        self.btn_all = QPushButton("全部已交")
        self.btn_none = QPushButton("全部未交")
        self.btn_export = QPushButton("导出未交清单…")
        for b in (self.btn_scan, self.btn_all, self.btn_none, self.btn_export):
            row2.addWidget(b)
        rv.addLayout(row2)
        body.addWidget(right, 2)
        root.addLayout(body, 1)

        self.btn_save.clicked.connect(self.save_names)
        self.btn_excel.clicked.connect(self.import_excel)
        self.btn_load.clicked.connect(self.load_preset)
        self.btn_save_preset.clicked.connect(self.save_preset)
        self.btn_scan.clicked.connect(self.scan_folder)
        self.btn_all.clicked.connect(lambda: self.set_all(True))
        self.btn_none.clicked.connect(lambda: self.set_all(False))
        self.btn_export.clicked.connect(self.export_todo)
        self.table.cellDoubleClicked.connect(self._toggle_row)

        self.reload()

    # ---- 数据 ----
    def reload(self):
        self.names = V.load_names()
        self.status = V.load_status()
        self.text.setPlainText("\n".join(self.names))
        self.refresh_presets()
        self.refresh()

    def refresh_presets(self):
        self.preset.clear()
        try:
            for n in V.list_presets():
                self.preset.addItem(n, n)
        except Exception:  # noqa: BLE001
            pass

    def refresh(self):
        self.table.setRowCount(len(self.names))
        done = 0
        for r, name in enumerate(self.names):
            st = self.status.get(name) or {}
            ok = bool(st.get("done"))
            done += ok
            it = QTableWidgetItem(name)
            it.setFlags(it.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(r, COL_NAME, it)
            s = QTableWidgetItem("已交" if ok else "未交")
            s.setFlags(s.flags() & ~Qt.ItemIsEditable)
            s.setForeground(Qt.darkGreen if ok else Qt.darkRed)
            self.table.setItem(r, COL_STATE, s)
            w = QTableWidgetItem(st.get("why", "") if ok else "")
            w.setFlags(w.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(r, COL_WHY, w)
        self.lbl_sum.setText(f"共 {len(self.names)}，已交 {done}，未交 {len(self.names) - done}")

    # ---- 名单 ----
    def save_names(self, quiet=False):
        names = [ln.strip() for ln in self.text.toPlainText().splitlines() if ln.strip()]
        self.names = names
        V.save_names(names)
        keep = {k: v for k, v in self.status.items() if k in names}
        self.status = keep
        V.save_status(keep)
        self.refresh()
        if not quiet:
            QMessageBox.information(self, "已保存", f"名单已保存：{len(names)} 个。")

    def load_preset(self):
        key = self.preset.currentData()
        if not key:
            return
        try:
            names = V.load_preset(key)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "载入失败", str(exc))
            return
        self.text.setPlainText("\n".join(names))
        QMessageBox.information(self, "已载入", f"「{key}」共 {len(names)} 个，确认后点「保存名单」。")

    def save_preset(self):
        if not self.text.toPlainText().strip():
            QMessageBox.information(self, "名单是空的", "先写几个名称再存预设。")
            return
        name, ok = QInputDialog.getText(self, "存为预设", "给这套名单起个名字：")
        if not ok or not name.strip():
            return
        names = [ln.strip() for ln in self.text.toPlainText().splitlines() if ln.strip()]
        V.save_preset(name.strip(), names)
        self.refresh_presets()
        self.preset.setCurrentText(name.strip())
        QMessageBox.information(self, "已保存", f"预设「{name.strip()}」（{len(names)} 个）已保存。")

    def import_excel(self):
        p, _ = QFileDialog.getOpenFileName(self, "选择 Excel", "", "Excel (*.xlsx *.xlsm)")
        if not p:
            return
        try:
            res = V.parse_excel_names(p)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "解析失败", str(exc))
            return
        names = res.get("names") or []
        if not names:
            QMessageBox.information(self, "没读到名称", "这个表里没找到像名称的内容。")
            return
        self.text.setPlainText("\n".join(names))
        QMessageBox.information(
            self, "解析完成",
            f"从第「{res.get('column', '')}」列读到 {len(names)} 个名称，已填进左边。\n"
            "核对没问题就点「保存名单」；想留着下次用就点「存为预设」。")

    # ---- 台账 ----
    def _toggle_row(self, row, _col):
        if not (0 <= row < len(self.names)):
            return
        name = self.names[row]
        st = self.status.get(name) or {}
        done = not bool(st.get("done"))
        self.status[name] = {"done": done, "why": "手动标记" if done else "", "src": "manual"}
        V.save_status(self.status)
        self.refresh()

    def set_all(self, done: bool):
        for n in self.names:
            self.status[n] = {"done": done, "why": "手动标记" if done else "", "src": "manual"}
        V.save_status(self.status)
        self.refresh()

    def scan_folder(self):
        d = QFileDialog.getExistingDirectory(self, "选一个文件夹，按文件名认")
        if not d:
            return
        root = Path(d)
        files = [p for p in root.iterdir() if p.is_file()]
        if not files:
            files = [p for p in root.rglob("*") if p.is_file()]
        hits = V.match_files(self.names, [str(p) for p in files]) or {}
        hits = {k: v for k, v in hits.items() if v}
        for name, hit in hits.items():
            self.status[name] = {"done": True, "why": Path(str(hit)).name, "src": "filename"}
        V.save_status(self.status)
        self.refresh()
        got = len(hits)
        QMessageBox.information(
            self, "识别完成",
            f"扫了 {len(files)} 个文件，认出 {got} 个已交；还有 {len(self.names) - got} 个没认到。")

    def export_todo(self):
        todo = [n for n in self.names if not (self.status.get(n) or {}).get("done")]
        if not todo:
            QMessageBox.information(self, "都交了", "全部交齐，没有未交的。")
            return
        p, flt = QFileDialog.getSaveFileName(
            self, "导出未交清单", str(Path(sysenv.desktop_dir()) / "未交清单.txt"),
            "文本 (*.txt);;CSV (*.csv)")
        if not p:
            return
        if p.lower().endswith(".csv"):
            with open(p, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.writer(f)
                w.writerow(["名称", "状态"])
                for n in todo:
                    w.writerow([n, "未交"])
        else:
            Path(p).write_text("未交材料名单\n\n" + "\n".join(f"{i}. {n}" for i, n in enumerate(todo, 1)),
                               encoding="utf-8-sig")
        QMessageBox.information(self, "已导出", f"{len(todo)} 个未交，已存到：\n{p}")
