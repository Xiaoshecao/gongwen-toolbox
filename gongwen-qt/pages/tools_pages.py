# -*- coding: utf-8 -*-
"""办公小工具四页（Qt 版）：批量重命名 / 字数统计 / 扫描件纠偏 / 签到表·交接单。"""
from __future__ import annotations

import csv
import re
from datetime import date
from pathlib import Path

from PySide6.QtCore import QThread, Qt, Signal
from PySide6.QtWidgets import (QCheckBox, QComboBox, QFileDialog, QGridLayout,
                               QHBoxLayout, QLabel, QLineEdit, QMessageBox,
                               QPlainTextEdit, QProgressBar, QPushButton,
                               QSpinBox, QTableWidget, QTableWidgetItem,
                               QVBoxLayout, QWidget)

from gwtyper.core import scanim, sysenv
from gwtyper.core import pdf_engine
from gwtyper.core.formdoc import write_table_docx as _write_table_docx
from ui.filepanel import FilePanel, accept_suffix, card, hint

IMG_EXT = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}
DOC_EXT = {".docx", ".doc", ".wps", ".rtf", ".pdf", ".txt", ".md", ".xlsx", ".pptx"}


def _page_head(title: str, desc: str) -> QHBoxLayout:
    bar = QHBoxLayout()
    t = QLabel(title)
    t.setObjectName("H1")
    bar.addWidget(t)
    bar.addSpacing(12)
    d = hint(desc)
    bar.addWidget(d)
    bar.addStretch(1)
    return bar


# ---------------------------------------------------------------- 重命名
class RenamePage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)
        root.addLayout(_page_head("批量重命名", "改名字不动内容；扩展名不参与改名"))

        body = QHBoxLayout()
        body.setSpacing(10)
        left, ll = card()
        self.panel = FilePanel("要改名的文件", accept=accept_suffix(DOC_EXT | IMG_EXT))
        ll.addWidget(self.panel)
        left_lay = ll
        body.addWidget(left, 2)

        right, rl = card("怎么改")
        grid = QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(8)
        self.prefix = QLineEdit()
        self.prefix.setPlaceholderText("加在名字前面，可留空")
        self.suffix = QLineEdit()
        self.suffix.setPlaceholderText("接在名字后面，如 -已核")
        self.find = QLineEdit()
        self.find.setPlaceholderText("原名里要替换掉的字，可留空")
        self.repl = QLineEdit()
        self.repl.setPlaceholderText("替换成")
        self.chk_num = QCheckBox("带序号")
        self.chk_num.setChecked(True)
        self.start = QSpinBox()
        self.start.setRange(0, 9999)
        self.start.setValue(1)
        self.width = QSpinBox()
        self.width.setRange(1, 6)
        self.width.setValue(2)
        grid.addWidget(QLabel("名字前缀"), 0, 0)
        grid.addWidget(self.prefix, 0, 1, 1, 3)
        grid.addWidget(QLabel("名字后缀"), 1, 0)
        grid.addWidget(self.suffix, 1, 1, 1, 3)
        grid.addWidget(QLabel("查找"), 2, 0)
        grid.addWidget(self.find, 2, 1)
        grid.addWidget(QLabel("替换为"), 2, 2)
        grid.addWidget(self.repl, 2, 3)
        grid.addWidget(self.chk_num, 3, 0)
        grid.addWidget(QLabel("从"), 3, 1)
        grid.addWidget(self.start, 3, 2)
        grid.addWidget(QLabel("位数"), 3, 3)
        grid.addWidget(self.width, 3, 4)
        rl.addLayout(grid)
        self.tbl = QTableWidget(0, 2)
        self.tbl.setHorizontalHeaderLabels(["原名", "改成"])
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.horizontalHeader().setStretchLastSection(True)
        rl.addWidget(self.tbl, 1)
        row = QHBoxLayout()
        self.btn_preview = QPushButton("预览新名字")
        self.btn_apply = QPushButton("开始改名")
        self.btn_apply.setObjectName("Primary")
        row.addWidget(self.btn_preview)
        row.addWidget(self.btn_apply)
        rl.addLayout(row)
        self.status = hint("先添加文件，再点「预览新名字」看一眼")
        rl.addWidget(self.status)
        body.addWidget(right, 3)
        root.addLayout(body, 1)

        self.btn_preview.clicked.connect(self.preview)
        self.btn_apply.clicked.connect(self.apply)

    def _new_name(self, p: Path, i: int) -> str:
        stem = p.stem
        f, r = self.find.text(), self.repl.text()
        if f:
            stem = stem.replace(f, r)
        pre = self.prefix.text()
        suf = self.suffix.text()
        num_on = self.chk_num.isChecked()
        num = str(self.start.value() + i).zfill(self.width.value()) if num_on else ""
        if pre or num_on:
            stem = f"{pre}{num}{stem}"
        return f"{stem}{suf}{p.suffix}"

    def preview(self):
        files = self.panel.files
        self.tbl.setRowCount(len(files))
        seen = {}
        for i, p in enumerate(files):
            new = self._new_name(p, i)
            self.tbl.setItem(i, 0, QTableWidgetItem(p.name))
            self.tbl.setItem(i, 1, QTableWidgetItem(new))
            key = str(p.parent / new)
            seen[key] = seen.get(key, 0) + 1
        dup = [k for k, v in seen.items() if v > 1]
        self.status.setText(f"{len(files)} 个文件" + ("；有重名，改不了，先调一下规则" if dup else "，看着没问题就点「开始改名」"))

    def apply(self):
        files = self.panel.files
        if not files:
            QMessageBox.information(self, "还没有文件", "先添加文件。")
            return
        self.preview()
        plans = []
        for i, p in enumerate(files):
            target = p.with_name(self._new_name(p, i))
            if target != p:
                plans.append((p, target))
        exists = [t for _, t in plans if t.exists()]
        if exists:
            QMessageBox.warning(self, "有同名文件", "这些名字已经存在，先改规则：\n" + "\n".join(x.name for x in exists[:8]))
            return
        done, bad = 0, []
        for src, dst in plans:
            try:
                src.rename(dst)
                done += 1
            except Exception as exc:  # noqa: BLE001
                bad.append(f"{src.name}：{exc}")
        self.status.setText(f"改了 {done} 个" + (f"，失败 {len(bad)} 个" if bad else ""))
        if bad:
            QMessageBox.warning(self, "部分失败", "\n".join(bad[:8]))
        self.panel.clear()
        self.tbl.setRowCount(0)


# ---------------------------------------------------------------- 字数统计
class StatWorker(QThread):
    progress = Signal(int, int, str)
    done = Signal(list, int)

    def __init__(self, files, parent=None):
        super().__init__(parent)
        self.files = files

    @staticmethod
    def _count(p: Path) -> tuple[int, int]:
        ext = p.suffix.lower()
        if ext == ".docx":
            import docx
            d = docx.Document(str(p))
            chars = sum(len(x.text.strip()) for x in d.paragraphs)
            for t in d.tables:
                for row in t.rows:
                    for c in row.cells:
                        chars += len(c.text.strip())
            return 0, chars
        if ext == ".pdf":
            pages = pdf_engine.pdf_pages_text(str(p))
            return len(pages), sum(len(re.sub(r"\s", "", t)) for t in pages)
        if ext in (".txt", ".md", ".csv"):
            try:
                txt = pdf_engine.txt_guess_encoding(str(p))
            except Exception:  # noqa: BLE001
                txt = ""
            return 0, len(re.sub(r"\s", "", txt))
        if ext in (".doc", ".wps", ".rtf"):
            from gwtyper.core import doc_source
            return 0, sum(len(x.text.strip()) for x in __import__("docx").Document(
                str(doc_source.to_docx(p))).paragraphs)
        raise RuntimeError("这种格式暂时统计不了")

    def run(self):
        rows = []
        total = 0
        n = len(self.files)
        for i, f in enumerate(self.files, 1):
            self.progress.emit(i, n, Path(f).name)
            try:
                pages, chars = self._count(Path(f))
                rows.append((Path(f).name, pages, chars, ""))
                total += chars
            except Exception as exc:  # noqa: BLE001
                rows.append((Path(f).name, 0, 0, str(exc)))
        self.done.emit(rows, total)


class StatPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)
        root.addLayout(_page_head("字数统计", "统计正文净字数（不含空格），可导出 CSV"))

        body = QHBoxLayout()
        body.setSpacing(10)
        left, ll = card()
        self.panel = FilePanel("要统计的文件", accept=accept_suffix(DOC_EXT))
        ll.addWidget(self.panel)
        body.addWidget(left, 2)

        right, rl = card("结果")
        self.tbl = QTableWidget(0, 3)
        self.tbl.setHorizontalHeaderLabels(["文件", "页数", "字数"])
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.horizontalHeader().setStretchLastSection(True)
        rl.addWidget(self.tbl, 1)
        self.pb = QProgressBar()
        self.pb.setVisible(False)
        rl.addWidget(self.pb)
        row = QHBoxLayout()
        self.btn_go = QPushButton("开始统计")
        self.btn_go.setObjectName("Primary")
        self.btn_csv = QPushButton("导出 CSV…")
        row.addWidget(self.btn_go)
        row.addWidget(self.btn_csv)
        rl.addLayout(row)
        self.status = hint("先添加文件")
        rl.addWidget(self.status)
        body.addWidget(right, 3)
        root.addLayout(body, 1)

        self.btn_go.clicked.connect(self.run)
        self.btn_csv.clicked.connect(self.export)
        self.worker = None
        self._rows = []

    def run(self):
        if self.worker and self.worker.isRunning():
            return
        if not self.panel.files:
            QMessageBox.information(self, "还没有文件", "先添加文件。")
            return
        self.pb.setVisible(True)
        self.pb.setRange(0, len(self.panel.files))
        self.btn_go.setEnabled(False)
        self.worker = StatWorker([str(p) for p in self.panel.files], self)
        self.worker.progress.connect(lambda i, n, name: (self.pb.setValue(i - 1),
                                                         self.status.setText(f"[{i}/{n}] {name}")))
        self.worker.done.connect(self._done)
        self.worker.start()

    def _done(self, rows, total):
        self._rows = rows
        self.tbl.setRowCount(len(rows))
        for i, (name, pages, chars, err) in enumerate(rows):
            self.tbl.setItem(i, 0, QTableWidgetItem(name))
            self.tbl.setItem(i, 1, QTableWidgetItem(str(pages) if not err else "-"))
            self.tbl.setItem(i, 2, QTableWidgetItem(str(chars) if not err else err))
        self.pb.setValue(self.pb.maximum())
        self.btn_go.setEnabled(True)
        self.status.setText(f"{len(rows)} 份，合计 {total} 字")

    def export(self):
        if not self._rows:
            QMessageBox.information(self, "还没统计", "先点「开始统计」。")
            return
        dst, _ = QFileDialog.getSaveFileName(self, "导出 CSV",
                                             str(Path(sysenv.export_dir()) / "字数统计.csv"),
                                             "CSV (*.csv)")
        if not dst:
            return
        with open(dst, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["文件", "页数", "字数"])
            total = 0
            for name, pages, chars, err in self._rows:
                w.writerow([name, pages, chars if not err else err])
                total += chars
            w.writerow(["合计", "", total])
        self.status.setText(f"已导出：{dst}")


# ---------------------------------------------------------------- 扫描件纠偏
class DeskewWorker(QThread):
    progress = Signal(int, int, str)
    done = Signal(int, list)

    def __init__(self, files, out_dir, deskew, trim, parent=None):
        super().__init__(parent)
        self.files, self.out_dir, self.deskew, self.trim = files, out_dir, deskew, trim

    def run(self):
        ok, bad = 0, []
        n = len(self.files)
        for i, f in enumerate(self.files, 1):
            src = Path(f)
            self.progress.emit(i, n, src.name)
            try:
                Path(self.out_dir).mkdir(parents=True, exist_ok=True)
                out = Path(self.out_dir) / f"{src.stem}-摆正{src.suffix}"
                scanim.fix(str(src), str(out), deskew=self.deskew, trim=self.trim)
                ok += 1
            except Exception as exc:  # noqa: BLE001
                bad.append(f"{src.name}：{exc}")
        self.done.emit(ok, bad)


class DeskewPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)
        root.addLayout(_page_head("扫描件纠偏", "把拍歪的、带黑边的扫描页摆正、裁边（不改颜色）"))

        body = QHBoxLayout()
        body.setSpacing(10)
        left, ll = card()
        self.panel = FilePanel("要处理的图片", accept=accept_suffix(IMG_EXT))
        ll.addWidget(self.panel)
        body.addWidget(left, 2)

        right, rl = card("怎么处理")
        self.chk_deskew = QCheckBox("摆正（倾斜超过 0.3° 才转）")
        self.chk_deskew.setChecked(True)
        self.chk_trim = QCheckBox("裁掉四周黑边/白边")
        self.chk_trim.setChecked(True)
        rl.addWidget(self.chk_deskew)
        rl.addWidget(self.chk_trim)
        row = QHBoxLayout()
        row.addWidget(QLabel("保存到"))
        self.out = QLineEdit(str(Path(sysenv.export_dir()) / "扫描件整理"))
        row.addWidget(self.out, 1)
        self.btn_browse = QPushButton("浏览…")
        row.addWidget(self.btn_browse)
        rl.addLayout(row)
        rl.addStretch(1)
        self.pb = QProgressBar()
        self.pb.setVisible(False)
        rl.addWidget(self.pb)
        self.status = hint("先添加图片")
        rl.addWidget(self.status)
        self.btn_go = QPushButton("开始处理")
        self.btn_go.setObjectName("Primary")
        rl.addWidget(self.btn_go)
        body.addWidget(right, 2)
        root.addLayout(body, 1)

        self.btn_browse.clicked.connect(self._browse)
        self.btn_go.clicked.connect(self.run)
        self.worker = None

    def _browse(self):
        d = QFileDialog.getExistingDirectory(self, "选择保存位置")
        if d:
            self.out.setText(d)

    def run(self):
        if self.worker and self.worker.isRunning():
            return
        if not self.panel.files:
            QMessageBox.information(self, "还没有图片", "先添加图片。")
            return
        self.pb.setVisible(True)
        self.pb.setRange(0, len(self.panel.files))
        self.btn_go.setEnabled(False)
        self.worker = DeskewWorker([str(p) for p in self.panel.files], self.out.text().strip(),
                                   self.chk_deskew.isChecked(), self.chk_trim.isChecked(), self)
        self.worker.progress.connect(lambda i, n, name: (self.pb.setValue(i - 1),
                                                         self.status.setText(f"[{i}/{n}] {name}")))
        self.worker.done.connect(self._done)
        self.worker.start()

    def _done(self, ok, bad):
        self.pb.setValue(self.pb.maximum())
        self.btn_go.setEnabled(True)
        self.status.setText(f"处理 {ok} 张" + (f"，失败 {len(bad)} 张" if bad else ""))
        if bad:
            QMessageBox.warning(self, "部分失败", "\n".join(bad[:8]))
        else:
            QMessageBox.information(self, "完成", f"处理好了 {ok} 张，在：{self.out.text()}")


# ---------------------------------------------------------------- 签到表 / 交接单
class FormPage(QWidget):
    KINDS = [("sign", "会议签到表"), ("hand", "材料交接单")]

    def __init__(self, parent=None):
        super().__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)
        root.addLayout(_page_head("签到表 · 材料交接单", "填一下标题和名单，直接生成能打印的 Word 表"))

        body = QHBoxLayout()
        body.setSpacing(10)
        left, ll = card("表的样子")
        grid = QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(8)
        self.kind = QComboBox()
        for k, label in self.KINDS:
            self.kind.addItem(label, k)
        self.title = QLineEdit("")
        self.sub = QLineEdit("")
        self.day = QLineEdit(date.today().strftime("%Y年%m月%d日"))
        grid.addWidget(QLabel("类型"), 0, 0)
        grid.addWidget(self.kind, 0, 1)
        grid.addWidget(QLabel("标题"), 1, 0)
        grid.addWidget(self.title, 1, 1)
        grid.addWidget(QLabel("副标题"), 2, 0)
        grid.addWidget(self.sub, 2, 1)
        grid.addWidget(QLabel("日期"), 3, 0)
        grid.addWidget(self.day, 3, 1)
        ll.addLayout(grid)
        ll.addWidget(hint("名单一行一个：姓名、单位、职务，可以用顿号或逗号分开"))
        self.names = QPlainTextEdit()
        self.names.setPlaceholderText("张三、党政办、主任\n李四、财政所、所长")
        ll.addWidget(self.names, 1)
        self.btn_go = QPushButton("生成 Word 表…")
        self.btn_go.setObjectName("Primary")
        ll.addWidget(self.btn_go)
        self.status = hint("")
        ll.addWidget(self.status)
        body.addWidget(left, 3)

        right, rl = card("预览")
        self.preview = QPlainTextEdit()
        self.preview.setReadOnly(True)
        rl.addWidget(self.preview, 1)
        body.addWidget(right, 2)
        root.addLayout(body, 1)

        self.btn_go.clicked.connect(self.build)
        self.kind.currentIndexChanged.connect(self._sync)
        self._sync()

    def _sync(self):
        k = self.kind.currentData()
        if k == "sign":
            self.title.setText("会 议 签 到 表")
            self.sub.setText("")
        else:
            self.title.setText("材 料 交 接 单")
            self.sub.setText("")
        self._render()

    def _parse_names(self):
        out = []
        for line in self.names.toPlainText().splitlines():
            line = line.strip()
            if not line:
                continue
            parts = [x for x in re.split(r"[、,，\t]", line) if x.strip()]
            out.append(parts)
        return out

    def _render(self):
        k = self.kind.currentData()
        head = ["序号", "姓名", "单位", "职务", "签名"] if k == "sign" else ["序号", "材料名称", "份数", "备注"]
        lines = [self.title.text(), self.sub.text(), self.day.text(), "  ".join(head)]
        for i, parts in enumerate(self._parse_names(), 1):
            lines.append(f"{i}  " + "  ".join(parts))
        self.preview.setPlainText("\n".join(x for x in lines if x))

    def build(self):
        rows = self._parse_names()
        if not rows:
            QMessageBox.information(self, "还没名单", "先在左边把名单一行一个填上。")
            return
        dst, _ = QFileDialog.getSaveFileName(
            self, "保存 Word 表",
            str(Path(sysenv.export_dir()) / f"{self.title.text().replace(' ', '') or '表格'}.docx"),
            "Word (*.docx)")
        if not dst:
            return
        names = ["、".join(p[:2]) if len(p) >= 2 else p[0] for p in rows]
        # _write_table_docx 期望 names 为「姓名、单位、职务」这种一行文本
        names = ["、".join(p) for p in rows]
        try:
            _write_table_docx(Path(dst), self.kind.currentData(), self.title.text(),
                              self.sub.text(), self.day.text(), names)
            self.status.setText(f"已生成：{dst}")
            QMessageBox.information(self, "完成", f"生成好了：\n{dst}")
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "生成失败", str(exc))
