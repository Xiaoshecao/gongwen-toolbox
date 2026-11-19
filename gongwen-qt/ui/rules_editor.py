# -*- coding: utf-8 -*-
"""格式标准设置（自定义规则）——Qt 版。

对应旧版 tkinter 的「规则集设置」：页面边距 / 页码 / 表格字号 + 各段落角色的
字体、字号、加粗、对齐、行距、缩进、段前段后。改完可以「应用到当前标准」或
「另存为新的标准」。
"""
from __future__ import annotations

from pathlib import Path

import yaml
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QFontDatabase
from PySide6.QtWidgets import (QCheckBox, QComboBox, QDialog, QDialogButtonBox,
                               QFormLayout, QGridLayout,
                               QHBoxLayout, QInputDialog, QLabel, QLineEdit,
                               QDoubleSpinBox, QMessageBox, QPushButton, QSpinBox, QTabWidget,
                               QTableWidget, QTableWidgetItem, QVBoxLayout,
                               QWidget)

from gwtyper.core import rules as rulesmod
from gwtyper.core.model import ROLE_CN

ALIGNS = [("left", "左对齐"), ("center", "居中"), ("right", "右对齐"),
          ("justify", "两端对齐")]

# 中文字号 → 磅值（公文常用）
CH_SIZES = [("初号", 42.0), ("小初", 36.0), ("一号", 26.0), ("小一", 24.0),
            ("二号", 22.0), ("小二", 18.0), ("三号", 16.0), ("小三", 15.0),
            ("四号", 14.0), ("小四", 12.0), ("五号", 10.5), ("小五", 9.0),
            ("六号", 7.5), ("小六", 6.5)]


class SizeBox(QComboBox):
    """字号选择：显示「二号」这种说法，值仍是磅数。"""

    def __init__(self, pt=None, parent=None):
        super().__init__(parent)
        for label, val in CH_SIZES:
            self.addItem(label, val)
        self.addItem("自定义", None)
        self.set_pt(pt)

    def set_pt(self, pt):
        pt = _num(pt, 16.0)
        idx = [i for i, (_, v) in enumerate(CH_SIZES) if abs(v - pt) < 0.01]
        if idx:
            self.setCurrentIndex(idx[0])
        else:
            self.setItemText(self.count() - 1, f"自定义 {pt:g} 磅")
            self.setCurrentIndex(self.count() - 1)

    def value(self) -> float:
        data = self.currentData()
        if data is None:
            # 自定义：从文字里抠数字
            import re
            m = re.search(r"([\d.]+)", self.currentText())
            return float(m.group(1)) if m else 16.0
        return float(data)


ROLE_COLS = ["角色", "字体", "字号", "加粗", "对齐", "行距", "首行缩进",
             "右缩进", "段前", "段后"]


def _num(v, default=0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default



class FontBox(QComboBox):
    """字体选择：可下拉也可直接输入字体名（QFontComboBox 会把不存在的字体
    静默替换成默认字体，导致保存时字体名被改坏）。"""

    def __init__(self, family="仿宋_GB2312", parent=None):
        super().__init__(parent)
        self.setEditable(True)
        try:
            fams = QFontDatabase.families()
        except Exception:  # noqa: BLE001
            fams = []
        self.addItems(fams)
        self.set_family(family)

    def set_family(self, family):
        family = str(family or "仿宋_GB2312")
        i = self.findText(family)
        if i < 0:
            self.insertItem(0, family)
            i = 0
        self.setCurrentIndex(i)

    def family(self) -> str:
        return (self.currentText() or "").strip() or "仿宋_GB2312"

class RulesDialog(QDialog):
    """编辑一个格式标准。ok 后 config 里是完整规则 dict。"""

    def __init__(self, parent, key: str):
        super().__init__(parent)
        self.key = key
        self.data = rulesmod.load(key) or {}
        self.setWindowTitle(f"格式标准设置 · {self.data.get('name', key)}")
        self.resize(980, 640)

        lay = QVBoxLayout(self)
        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_page(), "页面")
        self.tabs.addTab(self._build_roles(), "段落角色")
        lay.addWidget(self.tabs, 1)

        bar = QHBoxLayout()
        bar.addStretch(1)
        self.btn_as = QPushButton("另存为新的标准…")
        self.btn_new = QPushButton("保存为一份新标准…")
        self.btn_ok = QPushButton("应用到当前标准")
        self.btn_ok.setObjectName("Primary")
        self.btn_cancel = QPushButton("取消")
        bar.addWidget(self.btn_as)
        bar.addWidget(self.btn_ok)
        bar.addWidget(self.btn_cancel)
        lay.addLayout(bar)

        self.btn_as.clicked.connect(self.save_as_new)
        self.btn_ok.clicked.connect(self._apply_current)
        self.btn_cancel.clicked.connect(self.reject)
        self.saved_path: Path | None = None
        self.saved_key: str | None = None

    # ---------------- 页面页 ----------------
    def _build_page(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        page = self.data.get("page") or {}
        feats = self.data.get("features") or {}
        table = self.data.get("table") or {}

        f1 = QFormLayout()
        self.sp_top = self._spin(page.get("top_cm"), 0, 10, 0.1, " cm")
        self.sp_bottom = self._spin(page.get("bottom_cm"), 0, 10, 0.1, " cm")
        self.sp_left = self._spin(page.get("left_cm"), 0, 10, 0.1, " cm")
        self.sp_right = self._spin(page.get("right_cm"), 0, 10, 0.1, " cm")
        self.sp_footer = self._spin(page.get("footer_cm"), 0, 10, 0.1, " cm")
        f1.addRow("上边距", self.sp_top)
        f1.addRow("下边距", self.sp_bottom)
        f1.addRow("左边距", self.sp_left)
        f1.addRow("右边距", self.sp_right)
        f1.addRow("页脚距（页码位置）", self.sp_footer)
        lay.addLayout(f1)

        h = QLabel("页码")
        h.setObjectName("H2")
        lay.addWidget(h)
        f2 = QFormLayout()
        self.chk_pageno = QCheckBox("排页码")
        self.chk_pageno.setChecked(bool(feats.get("page_number", True)))
        f2.addRow("", self.chk_pageno)
        self.ed_fmt = QLineEdit(str(feats.get("page_number_fmt", "— {n} —")))
        self.ed_fmt.setToolTip("用 {n} 代表页码，例如 — {n} —")
        f2.addRow("页码格式", self.ed_fmt)
        self.ed_pnfont = FontBox(feats.get("page_number_font", "宋体"))
        f2.addRow("页码字体", self.ed_pnfont)
        self.sp_pnsize = SizeBox(feats.get("page_number_size_pt", 14))
        f2.addRow("页码字号", self.sp_pnsize)
        lay.addLayout(f2)

        h2 = QLabel("表格")
        h2.setObjectName("H2")
        lay.addWidget(h2)
        f3 = QFormLayout()
        self.ed_tfont = FontBox(table.get("font", "仿宋_GB2312"))
        f3.addRow("表内字体", self.ed_tfont)
        self.sp_tsize = SizeBox(table.get("size_pt", 14))
        f3.addRow("表内字号", self.sp_tsize)
        lay.addLayout(f3)
        lay.addStretch(1)
        return w

    @staticmethod
    def _spin(value, lo, hi, step, suffix="", decimals=1):
        sp = QDoubleSpinBox()
        sp.setDecimals(decimals)
        sp.setRange(float(lo), float(hi))
        sp.setSingleStep(float(step))
        sp.setValue(float(_num(value, lo)))
        sp.setSuffix(suffix)
        return sp

    # ---------------- 角色页 ----------------
    def _build_roles(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.addWidget(QLabel("每个角色的格式；字体是判断段落层级的主要依据"))
        roles = self.data.get("roles") or {}
        keys = list(roles.keys())
        self.tbl = QTableWidget(len(keys), len(ROLE_COLS))
        self.tbl.setHorizontalHeaderLabels(ROLE_COLS)
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.horizontalHeader().setStretchLastSection(True)
        self._role_keys = keys
        for r, k in enumerate(keys):
            spec = roles.get(k) or {}
            it = QTableWidgetItem(ROLE_CN.get(k, k) if isinstance(ROLE_CN, dict) else str(k))
            it.setData(Qt.UserRole, k)
            it.setFlags(it.flags() & ~Qt.ItemIsEditable)
            self.tbl.setItem(r, 0, it)

            fc = FontBox(spec.get("font", "仿宋_GB2312"))
            self.tbl.setCellWidget(r, 1, fc)

            sp = SizeBox(spec.get("size_pt", 16))
            self.tbl.setCellWidget(r, 2, sp)

            cb = QCheckBox()
            cb.setChecked(bool(spec.get("bold", False)))
            cellw = QWidget()
            cl = QHBoxLayout(cellw)
            cl.setContentsMargins(0, 0, 0, 0)
            cl.addWidget(cb, 0, Qt.AlignCenter)
            self.tbl.setCellWidget(r, 3, cellw)

            al = QComboBox()
            for val, label in ALIGNS:
                al.addItem(label, val)
            cur = str(spec.get("align", "justify"))
            idx = [i for i, (v, _) in enumerate(ALIGNS) if v == cur]
            al.setCurrentIndex(idx[0] if idx else 0)
            self.tbl.setCellWidget(r, 4, al)

            for col, fkey, default, lo, hi, step in (
                (5, "line_pt", 28, 10, 60, 0.5),
                (6, "first_line_chars", 2, 0, 10, 1),
                (7, "right_indent_chars", 0, 0, 10, 1),
                (8, "space_before_pt", 0, 0, 60, 1),
                (9, "space_after_pt", 0, 0, 60, 1),
            ):
                dec = 1 if fkey in ("line_pt", "space_before_pt", "space_after_pt") else 0
                sp2 = QDoubleSpinBox()
                sp2.setDecimals(dec)
                sp2.setRange(float(lo), float(hi))
                sp2.setSingleStep(float(step))
                sp2.setValue(float(_num(spec.get(fkey, default), default)))
                if fkey in ("line_pt", "space_before_pt", "space_after_pt"):
                    sp2.setSuffix(" 磅")
                self.tbl.setCellWidget(r, col, sp2)
        self.tbl.resizeColumnsToContents()
        lay.addWidget(self.tbl, 1)
        return w

    # ---------------- 收集 & 保存 ----------------
    def collect(self) -> dict:
        d = dict(self.data)
        page = dict(d.get("page") or {})
        page.update({
            "top_cm": round(self.sp_top.value(), 1),
            "bottom_cm": round(self.sp_bottom.value(), 1),
            "left_cm": round(self.sp_left.value(), 1),
            "right_cm": round(self.sp_right.value(), 1),
            "footer_cm": round(self.sp_footer.value(), 1),
        })
        d["page"] = page

        feats = dict(d.get("features") or {})
        feats.update({
            "page_number": self.chk_pageno.isChecked(),
            "page_number_fmt": self.ed_fmt.text().strip() or "— {n} —",
            "page_number_font": self.ed_pnfont.family(),
            "page_number_size_pt": round(self.sp_pnsize.value(), 1),
        })
        d["features"] = feats

        d["table"] = {
            "font": self.ed_tfont.family(),
            "size_pt": round(self.sp_tsize.value(), 1),
            "align": (d.get("table") or {}).get("align", "center"),
        }

        roles = {}
        for r, k in enumerate(self._role_keys):
            fc = self.tbl.cellWidget(r, 1)
            sp = self.tbl.cellWidget(r, 2)
            cbw = self.tbl.cellWidget(r, 3)
            cb = cbw.findChild(QCheckBox) if cbw else None
            al = self.tbl.cellWidget(r, 4)
            spec = {
                "font": fc.family() if fc else "仿宋_GB2312",
                "size_pt": round(sp.value(), 1) if sp else 16,
                "align": al.currentData() if al else "justify",
                "line_pt": round(self.tbl.cellWidget(r, 5).value(), 1),
                "first_line_chars": int(self.tbl.cellWidget(r, 6).value()),
                "right_indent_chars": int(self.tbl.cellWidget(r, 7).value()),
                "space_before_pt": round(self.tbl.cellWidget(r, 8).value(), 1),
                "space_after_pt": round(self.tbl.cellWidget(r, 9).value(), 1),
            }
            if cb is not None and cb.isChecked():
                spec["bold"] = True
            # 去掉等于 0 的可选项，保持 yaml 干净
            for k2 in ("right_indent_chars", "space_before_pt"):
                if not spec.get(k2):
                    spec.pop(k2, None)
            roles[k] = spec
        d["roles"] = roles
        return d

    def _dump(self, d: dict, path: Path) -> None:
        clean = {k: v for k, v in d.items() if not k.startswith("_")}
        text = ("# 由「格式标准设置」保存\n"
                + yaml.safe_dump(clean, allow_unicode=True, sort_keys=False))
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8")
        rulesmod._load_path.cache_clear()

    def _apply_current(self):
        try:
            d = self.collect()
            target = rulesmod.save_target(self.key)
            self._dump(d, target)
            self.saved_path = target
            self.saved_key = Path(target).stem
            QMessageBox.information(self, "已保存",
                                    f"改好了，存到：\n{target}\n\n（原来的内置标准没动，这是你的覆盖版本）")
            self.accept()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "保存失败", str(exc))

    def save_as_new(self):
        name, ok = QInputDialog.getText(self, "另存为新的标准", "起个名字：")
        if not ok or not name.strip():
            return
        try:
            d = self.collect()
            d["name"] = name.strip()
            target = rulesmod.save_target(name.strip())
            self._dump(d, target)
            self.saved_path = target
            self.saved_key = Path(target).stem
            QMessageBox.information(self, "已保存", f"新标准存到：\n{target}")
            self.accept()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "保存失败", str(exc))
