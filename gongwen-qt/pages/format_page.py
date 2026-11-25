# -*- coding: utf-8 -*-
"""公文排版页（Qt 版）：添加文件 → 只读体检 → 结果表。

性能设计：
- 检查放在后台线程（QThread），界面不卡；
- 解析结果按「路径 + 修改时间」缓存，重复检查同一份文件不再重新解析；
- 同一次检查里多份文件并行解析（线程池），只用 CPU 解析，结果仍按顺序回填。
"""
from __future__ import annotations

import os
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from PySide6.QtCore import QAbstractTableModel, QThread, Qt, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtCore import QUrl
from PySide6.QtWidgets import (QCheckBox, QComboBox, QDialog, QDialogButtonBox,
                               QFileDialog, QFrame, QHBoxLayout, QLineEdit,
                               QLabel, QListWidget, QListWidgetItem,
                               QMessageBox, QPushButton, QRadioButton,
                               QSplitter, QTableView, QVBoxLayout, QWidget)

from gwtyper.core import sysenv

from gwtyper.core import apply as applymod
from gwtyper.core import check as checkmod
from gwtyper.core import classify as clsmod
from gwtyper.core import parse_docx
from gwtyper.core import rules as rulesmod
from gwtyper.core.model import ROLE_CN

COLS = ["等级", "位置", "角色", "检查项", "现在", "应为"]
LEVEL_CN = {"error": "必改", "warn": "建议", "info": "提示"}
PROFILES = [("gb9704", "国家标准 GB/T 9704 公文格式"),
            ("news", "新闻稿 / 通讯"),
            ("report", "述职 / 总结报告")]

# 解析缓存：{路径: (mtime, size, Doc)}
_CACHE: dict[str, tuple[float, int, object]] = {}


def _load_doc(path: str):
    st = os.stat(path)
    hit = _CACHE.get(path)
    if hit and hit[0] == st.st_mtime and hit[1] == st.st_size:
        return hit[2], True
    doc = parse_docx.parse(path)
    _CACHE[path] = (st.st_mtime, st.st_size, doc)
    return doc, False


def _check_one(path: str, profile: str):
    """返回 (findings, parsed_from_cache)。注意：Doc 会被缓存，重新检查要重置角色。"""
    doc, cached = _load_doc(path)
    auto, review = clsmod.classify(doc.paras)   # 会就地写回 p.role
    fs = checkmod.check(doc, rulesmod.load(profile))
    return fs, cached


class FindingsModel(QAbstractTableModel):
    def __init__(self, rows=None):
        super().__init__()
        self.rows = rows or []

    def set_rows(self, rows):
        self.beginResetModel()
        self.rows = rows
        self.endResetModel()

    def rowCount(self, parent=None):
        return len(self.rows)

    def columnCount(self, parent=None):
        return len(COLS)

    def headerData(self, s, orient, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and orient == Qt.Horizontal:
            return COLS[s]
        return None

    def data(self, idx, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        f = self.rows[idx.row()]
        loc = f"第{f.index + 1}段" if f.index is not None else "整篇"
        vals = [LEVEL_CN.get(f.level, f.level), loc, ROLE_CN.get(f.role, f.role),
                f.field, f.actual, f.expected]
        return vals[idx.column()]


class CheckWorker(QThread):
    done = Signal(object, int, int, float, int)   # rows, err, warn, ms, cached
    failed = Signal(object)

    def __init__(self, files: list[str], profile: str, parent=None):
        super().__init__(parent)
        self.files = files
        self.profile = profile

    def run(self):
        t0 = time.perf_counter()
        rows, errors, cached_n = [], [], 0
        try:
            # 多份并行解析（最多 4 个线程），结果按文件顺序汇总
            with ThreadPoolExecutor(max_workers=min(4, max(1, len(self.files)))) as ex:
                results = list(ex.map(lambda p: _check_one(p, self.profile), self.files))
            for fs, cached in results:
                rows.extend(fs)
                if cached:
                    cached_n += 1
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(f"{exc}")
            return
        err = sum(1 for f in rows if f.level == "error")
        warn = sum(1 for f in rows if f.level == "warn")
        self.done.emit(rows, err, warn, (time.perf_counter() - t0) * 1000, cached_n)


class ApplyWorker(QThread):
    progress = Signal(int, int, str)
    done = Signal(int, list, str)          # ok, failed[(名字, 原因)], 输出目录
    failed = Signal(str)

    def __init__(self, files, profile, out_dir, same_dir, export_pdf, parent=None):
        super().__init__(parent)
        self.files, self.profile = list(files), profile
        self.out_dir, self.same_dir, self.export_pdf = out_dir, same_dir, export_pdf

    def run(self):
        try:
            rules = rulesmod.load(self.profile)
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(f"格式标准读不出来：{exc}")
            return
        ok = 0
        bad: list = []
        n = len(self.files)
        for i, src in enumerate(self.files, 1):
            src = Path(src)
            self.progress.emit(i, n, src.name)
            try:
                folder = src.parent if self.same_dir else Path(self.out_dir)
                folder.mkdir(parents=True, exist_ok=True)
                out = folder / f"{src.stem}-排版后.docx"
                applymod.apply(str(src), rules, str(out), 0.45)
                if self.export_pdf:
                    try:
                        from gwtyper.core import pdfwrite
                        pdfwrite.docx_to_pdf(str(out), str(out.with_suffix(".pdf")))
                    except Exception as pexc:  # noqa: BLE001
                        bad.append((out.stem + ".pdf", f"导 PDF 失败：{pexc}"))
                ok += 1
            except Exception as exc:  # noqa: BLE001
                bad.append((src.name, str(exc)))
        self.done.emit(ok, bad, str(self.out_dir if not self.same_dir else "和原文件放一起"))


class ApplyDialog(QDialog):
    """按规则排版：选输出位置、是否顺带导出 PDF。"""

    def __init__(self, files, profile_label, parent=None):
        super().__init__(parent)
        self.setWindowTitle("按规则排版")
        self.resize(560, 420)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(18, 16, 18, 14)
        lay.setSpacing(10)

        t = QLabel(f"把 {len(files)} 份公文按「{profile_label}」重排一遍")
        t.setObjectName("H2")
        lay.addWidget(t)
        tip = QLabel("只写新文件，原文件不动。输出文件名统一为「原名-排版后.docx」。")
        tip.setObjectName("Hint")
        tip.setWordWrap(True)
        lay.addWidget(tip)

        self.rb_same = QRadioButton("和原文件放在一起")
        self.rb_same.setChecked(True)
        lay.addWidget(self.rb_same)
        row = QHBoxLayout()
        self.rb_dir = QRadioButton("放到指定文件夹")
        row.addWidget(self.rb_dir)
        self.ed_dir = QLineEdit(str(Path(sysenv.export_dir()) / "排版结果"))
        self.ed_dir.setEnabled(False)
        row.addWidget(self.ed_dir, 1)
        self.btn_browse = QPushButton("浏览…")
        self.btn_browse.setEnabled(False)
        row.addWidget(self.btn_browse)
        lay.addLayout(row)

        self.chk_pdf = QCheckBox("同时导出一份 PDF")
        self.chk_pdf.setChecked(True)
        lay.addWidget(self.chk_pdf)

        lst = QListWidget()
        for f in files:
            lst.addItem(f"{Path(f).name}　→　{Path(f).stem}-排版后.docx")
        lst.setSelectionMode(QListWidget.NoSelection)
        lay.addWidget(lst, 1)

        btns = QDialogButtonBox()
        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        okbtn = QPushButton("开始排版")
        okbtn.setObjectName("Primary")
        okbtn.clicked.connect(self.accept)
        btns.addButton(cancel, QDialogButtonBox.RejectRole)
        btns.addButton(okbtn, QDialogButtonBox.AcceptRole)
        lay.addWidget(btns)

        self.rb_dir.toggled.connect(self._sync)
        self.btn_browse.clicked.connect(self._browse)

    def _sync(self):
        on = self.rb_dir.isChecked()
        self.ed_dir.setEnabled(on)
        self.btn_browse.setEnabled(on)

    def _browse(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出文件夹")
        if d:
            self.ed_dir.setText(d)

    def settings(self):
        return {
            "same_dir": self.rb_same.isChecked(),
            "out_dir": self.ed_dir.text().strip(),
            "export_pdf": self.chk_pdf.isChecked(),
        }


from ui.preview import PdfPreview


class FormatPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.files: list[Path] = []
        self.worker: CheckWorker | None = None
        self.apply_worker: ApplyWorker | None = None

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        bar = QHBoxLayout()
        title = QLabel("公文排版")
        title.setObjectName("H1")
        bar.addWidget(title)
        bar.addSpacing(14)
        bar.addWidget(QLabel("格式标准"))
        self.combo = QComboBox()
        self._fill_profiles()
        bar.addWidget(self.combo)
        self.btn_rules = QPushButton("格式标准设置…")
        self.btn_rules.setToolTip("调页边距、页码、各段落的字体字号行距（可另存为自己的标准）")
        bar.addWidget(self.btn_rules)
        self.btn_learn_rules = QPushButton("从文件学标准…")
        self.btn_learn_rules.setToolTip("给一份已经排好的公文，我学出它的页边距、字号、行距，存成一套新标准")
        bar.addWidget(self.btn_learn_rules)
        bar.addStretch(1)
        self.btn_add = QPushButton("添加文件…")
        self.btn_apply = QPushButton("按规则排版…")
        self.btn_apply.setObjectName("Primary")
        self.btn_apply.setEnabled(False)
        self.btn_check = QPushButton("检查（只读）")
        self.btn_check.setEnabled(False)
        bar.addWidget(self.btn_add)
        bar.addWidget(self.btn_check)
        bar.addWidget(self.btn_apply)
        root.addLayout(bar)

        split = QSplitter(Qt.Horizontal)
        left = QFrame()
        left.setObjectName("Card")
        lv = QVBoxLayout(left)
        lv.setContentsMargins(12, 12, 12, 12)
        lv.setSpacing(6)
        lv.addWidget(self._h2("文件"))
        self.list = QListWidget()
        lv.addWidget(self.list, 1)
        self.lbl_sum = QLabel("还没有文件")
        self.lbl_sum.setObjectName("Hint")
        lv.addWidget(self.lbl_sum)

        right = QFrame()
        right.setObjectName("Card")
        rv = QVBoxLayout(right)
        rv.setContentsMargins(12, 12, 12, 12)
        rv.setSpacing(6)
        head = QHBoxLayout()
        head.addWidget(self._h2("检查结果"))
        head.addStretch(1)
        self.lbl_time = QLabel("")
        self.lbl_time.setObjectName("Hint")
        head.addWidget(self.lbl_time)
        rv.addLayout(head)
        self.table = QTableView()
        self.model = FindingsModel()
        self.table.setModel(self.model)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        rv.addWidget(self.table, 1)

        prev = QFrame()
        prev.setObjectName("Card")
        pv = QVBoxLayout(prev)
        pv.setContentsMargins(12, 12, 12, 12)
        pv.setSpacing(6)
        ph = QHBoxLayout()
        ph.addWidget(self._h2("版面预览"))
        ph.addStretch(1)
        self.prev_mode = QComboBox()
        self.prev_mode.addItem("原件", "src")
        self.prev_mode.addItem("排版后", "out")
        ph.addWidget(self.prev_mode)
        pv.addLayout(ph)
        self.preview = PdfPreview()
        pv.addWidget(self.preview, 1)

        split.addWidget(left)
        split.addWidget(right)
        split.addWidget(prev)
        split.setSizes([260, 560, 460])
        root.addWidget(split, 1)

        self.btn_add.clicked.connect(self.add_files)
        self.btn_check.clicked.connect(self.run_check)
        self.btn_apply.clicked.connect(self.run_apply)
        self.btn_rules.clicked.connect(self.open_rules_editor)
        self.btn_learn_rules.clicked.connect(self.learn_from_file)
        self.combo.currentIndexChanged.connect(lambda _i: None)
        self.list.currentRowChanged.connect(lambda _i: self._preview_current())
        self.prev_mode.currentIndexChanged.connect(lambda _i: self._preview_current())
        self._last_out: Path | None = None

    def _preview_current(self):
        row = self.list.currentRow()
        if row < 0 or row >= len(self.files):
            self.preview.clear("点左边一个文件就能看版面")
            return
        if self.prev_mode.currentData() == "out":
            src = self.files[row]
            out = src.with_name(f"{src.stem}-排版后{src.suffix}")
            if out.exists():
                self.preview.show_file(out)
            else:
                self.preview.show_file(src)
                self.preview.lbl.setText("还没排版过，现在显示的是原件")
            return
        self.preview.show_file(self.files[row])

    def _h2(self, text):
        w = QLabel(text)
        w.setObjectName("H2")
        return w

    def learn_from_file(self):
        """从一份已排好的公文里学出一套格式标准（页边距、字号、行距、缩进）。"""
        from PySide6.QtWidgets import QFileDialog, QInputDialog, QMessageBox
        f, _ = QFileDialog.getOpenFileName(self, "选一份已经排好的公文", "", "Word 文档 (*.docx)")
        if not f:
            return
        try:
            from gwtyper.core import learn as learnmod
            prof = learnmod.learn(f)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "学不了", f"读这份文件时出错：{exc}")
            return
        name, ok = QInputDialog.getText(self, "给这套标准起个名字", "名字：",
                                        text=str(prof.get("name") or "参照件标准"))
        if not ok:
            return
        prof["name"] = (name or "").strip() or prof.get("name")
        try:
            from gwtyper.core import rules as rulesmod
            target = rulesmod.save_target(prof["name"])
            learnmod.save(prof, target)
            rulesmod._load_path.cache_clear()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "存不了", str(exc))
            return
        self._fill_profiles(keep=target.stem)
        QMessageBox.information(self, "学好了",
                                f"已按这份文件生成格式标准「{prof['name']}」，并切换过去。\n"
                                "可以直接点「按规则排版…」用这套标准排别的公文。")

    # ---- 格式标准 ----
    def _profiles(self):
        """按 rules 目录（内置 + 用户的）列出可选的格式标准。"""
        out = []
        try:
            for key in rulesmod.available():
                label = key
                try:
                    r = rulesmod.load(key)
                    label = str(r.get("name") or key)
                except Exception:  # noqa: BLE001
                    pass
                out.append([key, label])
        except Exception:  # noqa: BLE001
            out = [[k, v] for k, v in PROFILES]
        if not out:
            out = [[k, v] for k, v in PROFILES]
        # 两个标准同名时带上文件名区分，免得下拉里两条一模一样
        seen = {}
        for row in out:
            seen[row[1]] = seen.get(row[1], 0) + 1
        for row in out:
            if seen[row[1]] > 1:
                row[1] = f"{row[1]}（{row[0]}）"
        return [(k, v) for k, v in out]

    def _fill_profiles(self, keep=None):
        cur = keep or (self.combo.currentData() if self.combo.count() else None)
        self.combo.blockSignals(True)
        self.combo.clear()
        for key, label in self._profiles():
            self.combo.addItem(label, key)
        if cur:
            i = self.combo.findData(cur)
            if i >= 0:
                self.combo.setCurrentIndex(i)
        self.combo.blockSignals(False)

    def open_rules_editor(self):
        key = self.combo.currentData() or "gb9704"
        try:
            from ui.rules_editor import RulesDialog
        except Exception:  # noqa: BLE001
            try:
                from gongwen_qt.ui.rules_editor import RulesDialog  # type: ignore
            except Exception as exc:  # noqa: BLE001
                QMessageBox.warning(self, "打不开", f"格式标准设置加载失败：{exc}")
                return
        dlg = RulesDialog(self, key)
        if dlg.exec() and dlg.saved_key:
            self._fill_profiles(keep=dlg.saved_key)
            self._refresh() if hasattr(self, "_refresh") else None
            if self.files:
                self.run_check()
            self.lbl_time.setText("格式标准已更新")

    # ---- 行为 ----
    def add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "选择公文", "", "公文 (*.docx *.doc *.wps *.rtf);;所有文件 (*)")
        for p in paths:
            self._add(Path(p))

    def _add(self, p: Path):
        if not p.exists() or p in self.files:
            return
        self.files.append(p)
        it = QListWidgetItem(p.name)
        it.setToolTip(str(p))
        self.list.addItem(it)
        self.btn_check.setEnabled(True)
        self.btn_apply.setEnabled(True)
        self.lbl_sum.setText(f"{len(self.files)} 份")
        if self.list.currentRow() < 0:
            self.list.setCurrentRow(0)

    def run_check(self):
        if self.worker and self.worker.isRunning():
            return
        profile = self.combo.currentData() or "gb9704"
        self.btn_check.setEnabled(False)
        self.btn_add.setEnabled(False)
        self.lbl_time.setText("检查中…")
        self.worker = CheckWorker([str(p) for p in self.files], profile, self)
        self.worker.done.connect(self._on_done)
        self.worker.failed.connect(self._on_failed)
        self.worker.start()

    def _on_done(self, rows, err, warn, ms, cached_n):
        self.model.set_rows(rows)
        self.table.resizeColumnsToContents()
        self.lbl_sum.setText(f"{len(self.files)} 份｜必改 {err}／建议 {warn}（共 {len(rows)} 项）")
        extra = f"，缓存命中 {cached_n}" if cached_n else ""
        self.lbl_time.setText(f"用时 {ms / 1000:.2f} 秒{extra}")
        self.btn_check.setEnabled(True)
        self.btn_add.setEnabled(True)
        self.btn_apply.setEnabled(True)

    # ---- 按规则排版 ----
    def _selected_or_all(self) -> list[str]:
        idx = [self.list.row(i) for i in self.list.selectedItems()]
        if idx:
            return [str(self.files[i]) for i in sorted(idx)]
        return [str(p) for p in self.files]

    def run_apply(self):
        if self.apply_worker and self.apply_worker.isRunning():
            return
        files = self._selected_or_all()
        if not files:
            QMessageBox.information(self, "还没有文件", "先添加文件。")
            return
        dlg = ApplyDialog(files, self.combo.currentText(), self)
        if dlg.exec() != QDialog.Accepted:
            return
        st = dlg.settings()
        if not st["same_dir"] and not st["out_dir"]:
            QMessageBox.information(self, "还差一点", "选了「指定文件夹」但没填路径。")
            return
        self.btn_apply.setEnabled(False)
        self.btn_check.setEnabled(False)
        self.lbl_time.setText("排版中…")
        self.apply_worker = ApplyWorker(files, self.combo.currentData() or "gb9704",
                                        st["out_dir"], st["same_dir"], st["export_pdf"], self)
        self.apply_worker.progress.connect(
            lambda i, n, name: self.lbl_time.setText(f"[{i}/{n}] {name}"))
        self.apply_worker.done.connect(self._on_applied)
        self.apply_worker.failed.connect(self._on_apply_failed)
        self.apply_worker.start()

    def _on_applied(self, ok, bad, where):
        try:
            if self.list.currentRow() < 0 and self.files:
                self.list.setCurrentRow(0)
            self.prev_mode.setCurrentIndex(1)
            self._preview_current()
        except Exception:  # noqa: BLE001
            pass
        self.lbl_time.setText(f"排版完成 {ok} 份" + (f"，失败 {len(bad)} 份" if bad else ""))
        self.btn_apply.setEnabled(True)
        self.btn_check.setEnabled(True)
        box = QMessageBox(self)
        box.setWindowTitle("排版完成")
        text = f"生成 {ok} 份，放在：{where}"
        if bad:
            text += "\n\n没成功的：\n" + "\n".join(f"· {n}：{m}" for n, m in bad[:8])
        box.setText(text)
        b_open = box.addButton("打开文件夹", QMessageBox.ActionRole)
        box.addButton("好", QMessageBox.AcceptRole)
        box.exec()
        if box.clickedButton() is b_open:
            target = where if Path(where).is_dir() else str(Path(self._selected_or_all()[0]).parent)
            QDesktopServices.openUrl(QUrl.fromLocalFile(target))

    def _on_apply_failed(self, msg):
        self.lbl_time.setText("排版失败")
        self.btn_apply.setEnabled(True)
        self.btn_check.setEnabled(True)
        QMessageBox.warning(self, "排版失败", str(msg))

    def _on_failed(self, msg):
        self.lbl_time.setText("检查失败")
        self.btn_check.setEnabled(True)
        self.btn_add.setEnabled(True)
        QMessageBox.warning(self, "检查失败", str(msg))
