# -*- coding: utf-8 -*-
"""批量处理页（Qt 版）：图片压缩 / 文件查重 / 批量水印 / 脱敏 / 数据汇总 / 打印体检 / 名单比对。"""
from __future__ import annotations

import csv
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QCheckBox, QComboBox, QDoubleSpinBox,
                               QFileDialog, QGridLayout, QHBoxLayout, QLabel,
                               QLineEdit, QMessageBox, QPlainTextEdit,
                               QProgressBar, QPushButton, QSpinBox,
                               QTableWidget, QTableWidgetItem, QVBoxLayout,
                               QWidget)

from gwtyper.core import (convert as cv, dupfind, imgpress, listdiff, mask,
                          merge_tables, printcheck, sysenv, watermark)
from gwtyper.core import pdf_engine
from ui.filepanel import FilePanel, accept_suffix, card, hint
from pages.workers import BatchWorker, FuncWorker
from pages.tools_pages import IMG_EXT, DOC_EXT, _page_head


def _out_row(lay, default_name: str):
    row = QHBoxLayout()
    row.addWidget(QLabel("保存到"))
    ed = QLineEdit(str(Path(sysenv.export_dir()) / default_name))
    row.addWidget(ed, 1)
    btn = QPushButton("浏览…")
    row.addWidget(btn)

    def browse():
        d = QFileDialog.getExistingDirectory(None, "选择保存位置")
        if d:
            ed.setText(d)
    btn.clicked.connect(browse)
    lay.addLayout(row)
    return ed


def _run_bar(lay, go_text: str):
    pb = QProgressBar()
    pb.setVisible(False)
    lay.addWidget(pb)
    st = hint("")
    lay.addWidget(st)
    btn = QPushButton(go_text)
    btn.setObjectName("Primary")
    lay.addWidget(btn)
    return pb, st, btn


# ---------------------------------------------------------------- 图片压缩
class ImageCompressPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)
        root.addLayout(_page_head("图片压缩", "缩尺寸省空间；可「按长边」也可「按目标体积」压"))

        body = QHBoxLayout()
        body.setSpacing(10)
        left, ll = card()
        self.panel = FilePanel("要压缩的图片", accept=accept_suffix(IMG_EXT))
        ll.addWidget(self.panel)
        body.addWidget(left, 2)

        right, rl = card("怎么压")
        self.mode = QComboBox()
        self.mode.addItem("按长边 + 质量", "side")
        self.mode.addItem("按目标体积（KB）", "size")
        rl.addWidget(self.mode)
        grid = QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(8)
        self.side = QSpinBox()
        self.side.setRange(400, 8000)
        self.side.setValue(1600)
        self.side.setSingleStep(100)
        self.quality = QSpinBox()
        self.quality.setRange(40, 95)
        self.quality.setValue(75)
        self.target = QSpinBox()
        self.target.setRange(20, 5000)
        self.target.setValue(200)
        grid.addWidget(QLabel("长边不超过"), 0, 0)
        grid.addWidget(self.side, 0, 1)
        grid.addWidget(QLabel("质量"), 0, 2)
        grid.addWidget(self.quality, 0, 3)
        grid.addWidget(QLabel("目标体积"), 1, 0)
        grid.addWidget(self.target, 1, 1)
        grid.addWidget(QLabel("KB"), 1, 2)
        rl.addLayout(grid)
        self.out = _out_row(rl, "压缩后的图片")
        self.tbl = QTableWidget(0, 3)
        self.tbl.setHorizontalHeaderLabels(["文件", "原来", "压缩后"])
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.horizontalHeader().setStretchLastSection(True)
        rl.addWidget(self.tbl, 1)
        self.pb, self.status, self.btn_go = _run_bar(rl, "开始压缩")
        body.addWidget(right, 3)
        root.addLayout(body, 1)

        self.mode.currentIndexChanged.connect(self._sync)
        self.btn_go.clicked.connect(self.run)
        self.worker = None
        self._sync()

    def _sync(self):
        by_size = self.mode.currentData() == "size"
        for w in (self.side, self.quality):
            w.setEnabled(not by_size)
        self.target.setEnabled(by_size)

    def run(self):
        if self.worker and self.worker.isRunning():
            return
        if not self.panel.files:
            QMessageBox.information(self, "还没有图片", "先添加图片。")
            return
        out = self.out.text().strip()
        Path(out).mkdir(parents=True, exist_ok=True)
        by_size = self.mode.currentData() == "size"
        side, q, kb = self.side.value(), self.quality.value(), self.target.value()
        self._rows = []
        self.tbl.setRowCount(0)

        def one(src):
            if by_size:
                dst, before, after = imgpress.compress_to_size(src, str(Path(out) / Path(src).name),
                                                              target_kb=kb, max_side=side)
            else:
                dst, before, after = imgpress.compress(src, str(Path(out) / Path(src).name),
                                                       max_side=side, quality=q)
            self._rows.append((Path(src).name, before, after))

        self.pb.setVisible(True)
        self.pb.setRange(0, len(self.panel.files))
        self.btn_go.setEnabled(False)
        self.worker = BatchWorker([str(p) for p in self.panel.files], one, self)
        self.worker.progress.connect(lambda i, n, name: (self.pb.setValue(i - 1),
                                                         self.status.setText(f"[{i}/{n}] {name}")))
        self.worker.done.connect(self._done)
        self.worker.failed.connect(lambda m: QMessageBox.warning(self, "压缩失败", m))
        self.worker.start()

    def _done(self, ok, bad):
        self.pb.setValue(self.pb.maximum())
        self.btn_go.setEnabled(True)
        self.tbl.setRowCount(0)
        for name, before, after in getattr(self, "_rows", []):
            r = self.tbl.rowCount()
            self.tbl.insertRow(r)
            self.tbl.setItem(r, 0, QTableWidgetItem(name))
            self.tbl.setItem(r, 1, QTableWidgetItem(imgpress.human_size(before)))
            self.tbl.setItem(r, 2, QTableWidgetItem(imgpress.human_size(after)))
        self.status.setText(f"压好 {ok} 张" + (f"，失败 {len(bad)} 张" if bad else ""))
        if bad:
            QMessageBox.warning(self, "部分失败", "\n".join(bad[:8]))


# ---------------------------------------------------------------- 文件查重
class DupPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)
        root.addLayout(_page_head("文件查重", "按内容找重复：同一份文件、或正文一模一样的 Word/PDF"))

        body = QHBoxLayout()
        body.setSpacing(10)
        left, ll = card()
        self.panel = FilePanel("要查的文件", accept=accept_suffix(DOC_EXT | IMG_EXT))
        ll.addWidget(self.panel)
        body.addWidget(left, 2)

        right, rl = card("重复的")
        self.tbl = QTableWidget(0, 2)
        self.tbl.setHorizontalHeaderLabels(["怎么算重复", "文件"])
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.horizontalHeader().setStretchLastSection(True)
        rl.addWidget(self.tbl, 1)
        row = QHBoxLayout()
        self.btn_go = QPushButton("开始查重")
        self.btn_go.setObjectName("Primary")
        self.btn_csv = QPushButton("导出 CSV…")
        row.addWidget(self.btn_go)
        row.addWidget(self.btn_csv)
        rl.addLayout(row)
        self.pb = QProgressBar()
        self.pb.setVisible(False)
        rl.addWidget(self.pb)
        self.status = hint("先添加文件")
        rl.addWidget(self.status)
        body.addWidget(right, 3)
        root.addLayout(body, 1)

        self.btn_go.clicked.connect(self.run)
        self.btn_csv.clicked.connect(self.export)
        self.worker = None
        self._groups = []

    def run(self):
        if self.worker and self.worker.isRunning():
            return
        if len(self.panel.files) < 2:
            QMessageBox.information(self, "文件太少", "至少加两个文件才谈得上重复。")
            return
        files = [str(p) for p in self.panel.files]
        self.pb.setVisible(True)
        self.pb.setRange(0, 0)
        self.btn_go.setEnabled(False)
        self.status.setText("正在比对…")
        self.worker = FuncWorker(lambda prog: dupfind.find(files), self)
        self.worker.done.connect(self._done)
        self.worker.failed.connect(lambda m: QMessageBox.warning(self, "查重失败", m))
        self.worker.start()

    def _done(self, groups):
        self._groups = groups or []
        self.pb.setVisible(False)
        self.btn_go.setEnabled(True)
        self.tbl.setRowCount(0)
        for g in self._groups:
            for f in g["files"]:
                r = self.tbl.rowCount()
                self.tbl.insertRow(r)
                self.tbl.setItem(r, 0, QTableWidgetItem(g["why"]))
                self.tbl.setItem(r, 1, QTableWidgetItem(Path(f).name))
        self.status.setText(f"找到 {len(self._groups)} 组重复" if self._groups else "没发现重复")

    def export(self):
        if not self._groups:
            QMessageBox.information(self, "还没有结果", "先点「开始查重」。")
            return
        dst, _ = QFileDialog.getSaveFileName(self, "导出 CSV",
                                             str(Path(sysenv.export_dir()) / "查重结果.csv"),
                                             "CSV (*.csv)")
        if not dst:
            return
        with open(dst, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["怎么算重复", "文件"])
            for g in self._groups:
                for p in g["files"]:
                    w.writerow([g["why"], p])
        self.status.setText(f"已导出：{dst}")


# ---------------------------------------------------------------- 批量水印
class WatermarkPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)
        root.addLayout(_page_head("批量加水印", "给 PDF 打斜向灰字水印；Word 会先转成 PDF 再打"))

        body = QHBoxLayout()
        body.setSpacing(10)
        left, ll = card()
        self.panel = FilePanel("要加水印的文件", accept=accept_suffix({".pdf", ".docx", ".doc"}))
        ll.addWidget(self.panel)
        body.addWidget(left, 2)

        right, rl = card("水印")
        self.text = QLineEdit("内部资料 注意保存")
        rl.addWidget(QLabel("水印文字"))
        rl.addWidget(self.text)
        row = QHBoxLayout()
        row.addWidget(QLabel("透明度"))
        self.opacity = QDoubleSpinBox()
        self.opacity.setRange(0.03, 0.6)
        self.opacity.setSingleStep(0.02)
        self.opacity.setValue(0.12)
        row.addWidget(self.opacity)
        row.addStretch(1)
        rl.addLayout(row)
        self.out = _out_row(rl, "加了水印")
        self.pb, self.status, self.btn_go = _run_bar(rl, "开始加")
        rl.addStretch(1)
        body.addWidget(right, 2)
        root.addLayout(body, 1)

        self.btn_go.clicked.connect(self.run)
        self.worker = None

    def run(self):
        if self.worker and self.worker.isRunning():
            return
        if not self.panel.files:
            QMessageBox.information(self, "还没有文件", "先添加文件。")
            return
        text = self.text.text().strip()
        if not text:
            QMessageBox.information(self, "还差一点", "写个水印文字。")
            return
        out = Path(self.out.text().strip())
        out.mkdir(parents=True, exist_ok=True)
        op = float(self.opacity.value())

        def one(src):
            src = Path(src)
            pdf = src
            if src.suffix.lower() != ".pdf":
                tmp = out / (src.stem + ".pdf")
                cv.to_pdf(str(src), str(tmp), cv.Options())
                pdf = tmp
            watermark.apply(str(pdf), str(out / f"{src.stem}-水印.pdf"), text, opacity=op)

        self.pb.setVisible(True)
        self.pb.setRange(0, len(self.panel.files))
        self.btn_go.setEnabled(False)
        self.worker = BatchWorker([str(p) for p in self.panel.files], one, self)
        self.worker.progress.connect(lambda i, n, name: (self.pb.setValue(i - 1),
                                                         self.status.setText(f"[{i}/{n}] {name}")))
        self.worker.done.connect(self._done)
        self.worker.failed.connect(lambda m: QMessageBox.warning(self, "加不上", m))
        self.worker.start()

    def _done(self, ok, bad):
        self.pb.setValue(self.pb.maximum())
        self.btn_go.setEnabled(True)
        self.status.setText(f"加了 {ok} 份" + (f"，失败 {len(bad)} 份" if bad else ""))
        if bad:
            QMessageBox.warning(self, "部分失败", "\n".join(bad[:8]))


# ---------------------------------------------------------------- 敏感信息脱敏
class MaskPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)
        root.addLayout(_page_head("敏感信息脱敏", "把手机号、身份证号、银行卡号打码；输出到新文件"))

        body = QHBoxLayout()
        body.setSpacing(10)
        left, ll = card()
        self.panel = FilePanel("要脱敏的文件", accept=accept_suffix(
            {".pdf", ".docx", ".xlsx", ".csv"}))
        ll.addWidget(self.panel)
        body.addWidget(left, 2)

        right, rl = card("怎么脱")
        self.chk_strict = QCheckBox("身份证只留前 4 后 4（更严格）")
        self.chk_strict.setChecked(True)
        rl.addWidget(self.chk_strict)
        rl.addWidget(hint("手机 138****5678；身份证按上面勾的来；银行卡只留前后 4 位。"))
        self.out = _out_row(rl, "脱敏结果")
        self.pb, self.status, self.btn_go = _run_bar(rl, "开始脱敏")
        rl.addStretch(1)
        body.addWidget(right, 3)
        root.addLayout(body, 1)

        self.btn_go.clicked.connect(self.run)
        self.worker = None

    def run(self):
        if self.worker and self.worker.isRunning():
            return
        if not self.panel.files:
            QMessageBox.information(self, "还没有文件", "先添加文件。")
            return
        out = Path(self.out.text().strip())
        out.mkdir(parents=True, exist_ok=True)
        strict = self.chk_strict.isChecked()
        tally = {"id": 0, "phone": 0, "bank": 0}

        def one(src):
            src = Path(src)
            dst = out / (src.stem + "-脱敏" + src.suffix)
            res = mask.run(str(src), str(dst), strict_id=strict) or {}
            for k in tally:
                tally[k] += int(res.get(k, 0) or 0)

        self.pb.setVisible(True)
        self.pb.setRange(0, len(self.panel.files))
        self.btn_go.setEnabled(False)
        self.worker = BatchWorker([str(p) for p in self.panel.files], one, self)
        self.worker.progress.connect(lambda i, n, name: (self.pb.setValue(i - 1),
                                                         self.status.setText(f"[{i}/{n}] {name}")))
        self.worker.done.connect(lambda ok, bad: self._done(ok, bad, tally))
        self.worker.failed.connect(lambda m: QMessageBox.warning(self, "脱敏失败", m))
        self.worker.start()

    def _done(self, ok, bad, tally):
        self.pb.setValue(self.pb.maximum())
        self.btn_go.setEnabled(True)
        self.status.setText(f"处理 {ok} 份（身份证 {tally['id']}、手机 {tally['phone']}、银行卡 {tally['bank']}）"
                            + (f"，失败 {len(bad)}" if bad else ""))
        if bad:
            QMessageBox.warning(self, "部分失败", "\n".join(bad[:8]))


# ---------------------------------------------------------------- 数据汇总
class MergeTablePage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)
        root.addLayout(_page_head("数据汇总", "把多个 Excel/CSV 合成一张表，列名自动对齐（并集）"))

        body = QHBoxLayout()
        body.setSpacing(10)
        left, ll = card()
        self.panel = FilePanel("要汇总的表", accept=accept_suffix({".xlsx", ".xlsm", ".csv"}))
        ll.addWidget(self.panel)
        body.addWidget(left, 2)

        right, rl = card("汇总成一张")
        row = QHBoxLayout()
        row.addWidget(QLabel("输出文件"))
        self.out = QLineEdit(str(Path(sysenv.export_dir()) / "数据汇总.xlsx"))
        row.addWidget(self.out, 1)
        btn = QPushButton("另存为…")
        row.addWidget(btn)
        rl.addLayout(row)

        def browse():
            p, _ = QFileDialog.getSaveFileName(None, "保存汇总表", self.out.text(), "Excel (*.xlsx)")
            if p:
                self.out.setText(p)
        btn.clicked.connect(browse)
        self.pb, self.status, self.btn_go = _run_bar(rl, "开始汇总")
        rl.addStretch(1)
        body.addWidget(right, 3)
        root.addLayout(body, 1)

        self.btn_go.clicked.connect(self.run)
        self.worker = None

    def run(self):
        if self.worker and self.worker.isRunning():
            return
        if not self.panel.files:
            QMessageBox.information(self, "还没有表", "先添加表格文件。")
            return
        files = [str(p) for p in self.panel.files]
        dst = self.out.text().strip()
        self.pb.setVisible(True)
        self.pb.setRange(0, 0)
        self.btn_go.setEnabled(False)
        self.status.setText("正在汇总…")
        self.worker = FuncWorker(lambda prog: merge_tables.summarize(files, dst), self)
        self.worker.done.connect(self._done)
        self.worker.failed.connect(lambda m: QMessageBox.warning(self, "汇总失败", m))
        self.worker.start()

    def _done(self, res):
        self.pb.setVisible(False)
        self.btn_go.setEnabled(True)
        self.status.setText(f"汇总 {res.get('files', 0)} 个表、{res.get('rows', 0)} 行、{res.get('cols', 0)} 列"
                            f" → {self.out.text()}")


# ---------------------------------------------------------------- 打印体检
class PrintCheckPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)
        root.addLayout(_page_head("打印体检", "打出来才发现有问题？先查越界、字号太小、行距太挤、图片太糊"))

        body = QHBoxLayout()
        body.setSpacing(10)
        left, ll = card()
        self.panel = FilePanel("要检查的文件", accept=accept_suffix(DOC_EXT))
        ll.addWidget(self.panel)
        body.addWidget(left, 2)

        right, rl = card("结果")
        self.tbl = QTableWidget(0, 3)
        self.tbl.setHorizontalHeaderLabels(["文件", "等级", "问题"])
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.horizontalHeader().setStretchLastSection(True)
        rl.addWidget(self.tbl, 1)
        self.pb, self.status, self.btn_go = _run_bar(rl, "开始检查")
        body.addWidget(right, 3)
        root.addLayout(body, 1)

        self.btn_go.clicked.connect(self.run)
        self.worker = None

    def run(self):
        if self.worker and self.worker.isRunning():
            return
        if not self.panel.files:
            QMessageBox.information(self, "还没有文件", "先添加文件。")
            return
        files = [str(p) for p in self.panel.files]
        self.pb.setVisible(True)
        self.pb.setRange(0, 0)
        self.btn_go.setEnabled(False)
        self.status.setText("正在检查…")

        def job(prog):
            rows = []
            for f in files:
                for fd in printcheck.check(f):
                    rows.append((Path(f).name, fd.get("level", ""), fd.get("msg", "")))
            return rows

        self.worker = FuncWorker(job, self)
        self.worker.done.connect(self._done)
        self.worker.failed.connect(lambda m: QMessageBox.warning(self, "检查失败", m))
        self.worker.start()

    def _done(self, rows):
        self.pb.setVisible(False)
        self.btn_go.setEnabled(True)
        self.tbl.setRowCount(0)
        for name, level, msg in rows:
            r = self.tbl.rowCount()
            self.tbl.insertRow(r)
            self.tbl.setItem(r, 0, QTableWidgetItem(name))
            self.tbl.setItem(r, 1, QTableWidgetItem({"error": "必改", "warn": "建议"}.get(level, level)))
            self.tbl.setItem(r, 2, QTableWidgetItem(msg))
        self.status.setText(f"{len(rows)} 条问题" if rows else "没发现问题，可以直接打")


# ---------------------------------------------------------------- 名单比对
class ListDiffPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)
        root.addLayout(_page_head("名单比对", "两份名单对一下：谁没交、多的、重的"))

        body = QHBoxLayout()
        body.setSpacing(10)
        left, ll = card("上级名单")
        self.a = QPlainTextEdit()
        self.a.setPlaceholderText("一行一个；也可以点下面按钮读文件")
        ll.addWidget(self.a, 1)
        ba = QPushButton("读文件…")
        ll.addWidget(ba)
        body.addWidget(left, 1)

        mid, ml = card("我的名单")
        self.b = QPlainTextEdit()
        self.b.setPlaceholderText("一行一个")
        ml.addWidget(self.b, 1)
        bb = QPushButton("读文件…")
        ml.addWidget(bb)
        body.addWidget(mid, 1)

        right, rl = card("比对结果")
        self.tbl = QTableWidget(0, 2)
        self.tbl.setHorizontalHeaderLabels(["情况", "名称"])
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.horizontalHeader().setStretchLastSection(True)
        rl.addWidget(self.tbl, 1)
        row = QHBoxLayout()
        self.btn_go = QPushButton("开始比对")
        self.btn_go.setObjectName("Primary")
        self.btn_txt = QPushButton("导出结果…")
        row.addWidget(self.btn_go)
        row.addWidget(self.btn_txt)
        rl.addLayout(row)
        self.status = hint("")
        rl.addWidget(self.status)
        body.addWidget(right, 2)
        root.addLayout(body, 1)

        self.btn_go.clicked.connect(self.run)
        self.btn_txt.clicked.connect(self.export)
        ba.clicked.connect(lambda: self._load(self.a))
        bb.clicked.connect(lambda: self._load(self.b))
        self._res = None

    def _load(self, target: QPlainTextEdit):
        p, _ = QFileDialog.getOpenFileName(self, "选择名单文件", "",
                                           "名单 (*.txt *.docx *.xlsx);;所有文件 (*)")
        if not p:
            return
        try:
            names = listdiff.parse(p)
            target.setPlainText("\n".join(names))
            self.status.setText(f"读进 {len(names)} 个名字")
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "读不进来", str(exc))

    def run(self):
        a = listdiff.parse(self.a.toPlainText())
        b = listdiff.parse(self.b.toPlainText())
        res = listdiff.diff(a, b)
        self._res = (a, b, res)
        self.tbl.setRowCount(0)
        labels = [("only_a", "上级有、我这没有"), ("only_b", "我有、上级没有"),
                  ("same", "两边都有"), ("dup_a", "上级名单里重了"), ("dup_b", "我的名单里重了")]
        n = 0
        for key, label in labels:
            for name in res.get(key, []):
                r = self.tbl.rowCount()
                self.tbl.insertRow(r)
                self.tbl.setItem(r, 0, QTableWidgetItem(label))
                self.tbl.setItem(r, 1, QTableWidgetItem(name))
                n += 1
        self.status.setText(f"上级 {len(a)} 个、我的 {len(b)} 个；差异 {len(res.get('only_a', [])) + len(res.get('only_b', []))} 个")

    def export(self):
        if not self._res:
            QMessageBox.information(self, "还没比对", "先点「开始比对」。")
            return
        dst, _ = QFileDialog.getSaveFileName(self, "导出结果",
                                             str(Path(sysenv.export_dir()) / "名单比对.txt"),
                                             "文本 (*.txt)")
        if not dst:
            return
        a, b, res = self._res
        Path(dst).write_text(listdiff.to_text(a, b, res), encoding="utf-8-sig")
        self.status.setText(f"已导出：{dst}")
