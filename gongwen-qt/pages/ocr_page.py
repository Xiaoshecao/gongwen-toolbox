# -*- coding: utf-8 -*-
"""OCR 识别页（Qt 版）：图片/PDF/Office → 文字 / Word / 双层 PDF。

识别结果先给用户核对、可以直接改，再落盘（与原版一致）。
"""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QThread, Qt, Signal
from ui.controls import DpiPicker
from PySide6.QtWidgets import (QComboBox, QDialog, QDialogButtonBox, QFileDialog,
                               QFrame, QHBoxLayout, QLabel, QLineEdit,
                               QListWidget, QListWidgetItem, QMessageBox,
                               QPlainTextEdit, QProgressBar, QPushButton,
                               QSpinBox, QVBoxLayout, QWidget)

from gwtyper.core import ocr_text as ocr
from gwtyper.core import sysenv

MODES = [
    ("txt", "纯文字（.txt）", ".txt"),
    ("docx_layout", "Word（按原版式，不逐句改）", ".docx"),
    ("docx_text", "Word（一段一行，可以改字）", ".docx"),
    ("double", "双层 PDF（能看也能搜）", ".pdf"),
]
EXT = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp", ".gif", ".pdf",
       ".docx", ".doc", ".wps", ".xlsx", ".pptx"}


class OcrWorker(QThread):
    progress = Signal(int, int, str)
    one = Signal(str, str, object)     # src, mode, result(dict)
    one_failed = Signal(str, str)
    finished_all = Signal(int, int)

    def __init__(self, files, mode, dpi, out_dir, parent=None):
        super().__init__(parent)
        self.files, self.mode, self.dpi, self.out_dir = files, mode, dpi, out_dir

    def run(self):
        ok = bad = 0
        n = len(self.files)
        for i, f in enumerate(self.files, 1):
            self.progress.emit(i, n, Path(f).name)
            try:
                stem = Path(f).stem
                if self.mode in ("txt", "docx_layout", "docx_text"):
                    res = ocr.recognize(f, dpi=self.dpi)
                    self.one.emit(f, self.mode, res)
                elif self.mode == "docx_layout":
                    out = str(Path(self.out_dir) / f"{stem}.docx")
                    res = ocr.to_docx(f, out, dpi=self.dpi)
                    self.one.emit(f, "docx_layout_done", res)
                elif self.mode == "double":
                    out = str(Path(self.out_dir) / f"{stem}-双层.pdf")
                    res = ocr.to_pdf_double(f, out, dpi=self.dpi)
                    self.one.emit(f, "double", res)
                else:
                    raise RuntimeError("不认识的方式")
                ok += 1
            except Exception as exc:  # noqa: BLE001
                bad += 1
                self.one_failed.emit(f, str(exc))
        self.finished_all.emit(ok, bad)


class EditDialog(QDialog):
    """识别结果核对窗：可以改字再保存。"""

    def __init__(self, name: str, text: str, ext_label: str, index=0, total=1, parent=None):
        super().__init__(parent)
        self.setWindowTitle("识别结果（可以直接改）")
        self.resize(820, 640)
        self.result_text: str | None = None

        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 14, 16, 14)
        lay.setSpacing(8)

        head = QHBoxLayout()
        t = QLabel(name)
        t.setObjectName("H2")
        head.addWidget(t)
        head.addStretch(1)
        cnt = QLabel(f"{index}/{total}")
        cnt.setObjectName("Hint")
        head.addWidget(cnt)
        lay.addLayout(head)

        self.editor = QPlainTextEdit()
        self.editor.setPlainText(text)
        lay.addWidget(self.editor, 1)

        self.lbl_n = QLabel("")
        self.lbl_n.setObjectName("Hint")
        lay.addWidget(self.lbl_n)
        self.editor.textChanged.connect(self._update_n)
        self._update_n()

        box = QDialogButtonBox()
        b_skip = QPushButton("跳过这个")
        b_skip.clicked.connect(self.reject)
        b_save = QPushButton(f"保存为 {ext_label}")
        b_save.setObjectName("Primary")
        b_save.clicked.connect(self._save)
        box.addButton(b_skip, QDialogButtonBox.RejectRole)
        box.addButton(b_save, QDialogButtonBox.AcceptRole)
        lay.addWidget(box)

    def _update_n(self):
        t = self.editor.toPlainText()
        self.lbl_n.setText(f"{len(t.replace(chr(10), '').replace(' ', ''))} 字")

    def _save(self):
        self.result_text = self.editor.toPlainText()
        self.accept()


class OcrPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.files: list[Path] = []
        self.worker: OcrWorker | None = None
        self._queue: list[tuple[str, dict]] = []

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        bar = QHBoxLayout()
        t = QLabel("OCR 识别")
        t.setObjectName("H1")
        bar.addWidget(t)
        bar.addSpacing(12)
        tip = QLabel("扫描件、照片、PDF 都能认；识别完先给你核对，改完再保存")
        tip.setObjectName("Muted")
        bar.addWidget(tip)
        bar.addStretch(1)
        self.btn_add = QPushButton("添加文件…")
        self.btn_clear = QPushButton("清空")
        bar.addWidget(self.btn_add)
        bar.addWidget(self.btn_clear)
        root.addLayout(bar)

        body = QHBoxLayout()
        body.setSpacing(10)
        left = QFrame()
        left.setObjectName("Card")
        lv = QVBoxLayout(left)
        lv.setContentsMargins(12, 12, 12, 12)
        h = QLabel("要识别的文件")
        h.setObjectName("H2")
        lv.addWidget(h)
        self.list = QListWidget()
        lv.addWidget(self.list, 1)
        self.lbl_cnt = QLabel("0 个文件")
        self.lbl_cnt.setObjectName("Hint")
        lv.addWidget(self.lbl_cnt)
        body.addWidget(left, 1)

        right = QFrame()
        right.setObjectName("Card")
        rv = QVBoxLayout(right)
        rv.setContentsMargins(14, 12, 14, 12)
        rv.setSpacing(8)
        h2 = QLabel("要什么成品")
        h2.setObjectName("H2")
        rv.addWidget(h2)
        row = QHBoxLayout()
        row.addWidget(QLabel("成品"))
        self.mode = QComboBox()
        for key, label, _ in MODES:
            self.mode.addItem(label, key)
        row.addWidget(self.mode, 1)
        rv.addLayout(row)
        self.lbl_tip = QLabel("")
        self.lbl_tip.setObjectName("Hint")
        self.lbl_tip.setWordWrap(True)
        rv.addWidget(self.lbl_tip)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("保存到"))
        self.out = QLineEdit(str(Path(sysenv.export_dir()) / "识别结果"))
        row2.addWidget(self.out, 1)
        self.btn_browse = QPushButton("浏览…")
        row2.addWidget(self.btn_browse)
        rv.addLayout(row2)

        row3 = QHBoxLayout()
        row3.addWidget(QLabel("识别精度"))
        self.dpi = DpiPicker(200, 120, 400)
        row3.addWidget(self.dpi)
        row3.addStretch(1)
        rv.addLayout(row3)

        rv.addStretch(1)
        self.pb = QProgressBar()
        self.pb.setVisible(False)
        rv.addWidget(self.pb)
        self.status = QLabel("先添加文件")
        self.status.setObjectName("Hint")
        self.status.setWordWrap(True)
        rv.addWidget(self.status)
        self.btn_go = QPushButton("开始识别")
        self.btn_go.setObjectName("Primary")
        self.btn_go.setEnabled(False)
        rv.addWidget(self.btn_go)
        right.setFixedWidth(420)
        body.addWidget(right, 0)
        root.addLayout(body, 1)

        self.btn_add.clicked.connect(self.add_files)
        self.btn_clear.clicked.connect(self.clear)
        self.btn_browse.clicked.connect(self.pick_out)
        self.btn_go.clicked.connect(self.run_ocr)
        self.mode.currentIndexChanged.connect(self._sync_tip)
        self._sync_tip()
        self._refresh()

    # ---- 文件 ----
    def add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "选择文件", "",
            "支持的文件 (*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp *.pdf *.docx *.doc *.wps *.xlsx *.pptx);;所有文件 (*)")
        for p in paths:
            self._add(Path(p))

    def _add(self, p: Path):
        if p.is_file() and p not in self.files:
            self.files.append(p)
        self._refresh()

    def set_files(self, paths):
        for p in paths:
            self._add(Path(p))

    def clear(self):
        self.files.clear()
        self._refresh()

    def _refresh(self):
        self.list.clear()
        for p in self.files:
            self.list.addItem(QListWidgetItem(p.name))
        self.lbl_cnt.setText(f"{len(self.files)} 个文件")
        self.btn_go.setEnabled(bool(self.files))

    def _sync_tip(self):
        key = self.mode.currentData()
        tips = {
            "txt": "只出文字，一段一行，方便再粘到别处。",
            "docx_layout": "按原来的版面排成 Word（字号、居中、缩进按位置估），可能需人工微调；这一种不弹改字窗。",
            "docx_text": "只保留文字，一段一行，统一仿宋三号；保存前可以改字。",
            "double": "先认字 → 你核对修改 → 合成：下层原图、上层你改过的文字，能看也能搜。",
        }
        self.lbl_tip.setText(tips.get(key, ""))

    def pick_out(self):
        d = QFileDialog.getExistingDirectory(self, "选择保存位置")
        if d:
            self.out.setText(d)

    # ---- 识别 ----
    def run_ocr(self):
        if self.worker and self.worker.isRunning():
            return
        mode = self.mode.currentData()
        out = self.out.text().strip()
        if not out:
            QMessageBox.information(self, "还差一点", "先选保存位置。")
            return
        Path(out).mkdir(parents=True, exist_ok=True)
        self.pb.setVisible(True)
        self.pb.setRange(0, len(self.files))
        self.pb.setValue(0)
        self.btn_go.setEnabled(False)
        self.status.setText("正在识别…")
        self._queue.clear()
        self._fails: list[str] = []
        self.worker = OcrWorker([str(p) for p in self.files], mode, self.dpi.value(), out, self)
        self.worker.progress.connect(lambda i, n, name: (self.pb.setValue(i - 1),
                                                        self.status.setText(f"[{i}/{n}] {name}")))
        self.worker.one.connect(self._collect)
        self.worker.one_failed.connect(lambda f, m: self._fails.append(f"{Path(f).name}：{m}"))
        self.worker.finished_all.connect(lambda ok, bad: self._review(out))
        self.worker.start()

    def _collect(self, src, mode, res):
        self._queue.append((src, {"mode": mode, "res": res}))

    def _review(self, out_dir: str):
        self.pb.setValue(self.pb.maximum())
        saved = 0
        total = len(self._queue)
        for i, (src, item) in enumerate(self._queue, 1):
            mode, res = item["mode"], item["res"]
            stem = Path(src).stem
            if mode in ("double", "docx_layout_done"):
                saved += 1
                continue
            text = res.get("text", "")
            ext_label = "文本" if mode == "txt" else "Word"
            dlg = EditDialog(Path(src).name, text, ext_label, i, total, self)
            if dlg.exec() != QDialog.Accepted or dlg.result_text is None:
                continue
            try:
                if mode == "txt":
                    Path(out_dir, f"{stem}.txt").write_text(dlg.result_text, encoding="utf-8-sig")
                else:
                    ocr.text_to_docx(dlg.result_text, str(Path(out_dir, f"{stem}.docx")))
                saved += 1
            except Exception as exc:  # noqa: BLE001
                self._fails.append(f"{stem}：{exc}")
        self.btn_go.setEnabled(True)
        msg = f"完成 {saved} 个"
        if self._fails:
            msg += f"，失败 {len(self._fails)} 个"
        self.status.setText(msg)
        if self._fails:
            QMessageBox.warning(self, "部分失败", "\n".join(self._fails[:8]))
        else:
            QMessageBox.information(self, "完成", f"{msg}。\n文件在：{out_dir}")
