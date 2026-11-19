# -*- coding: utf-8 -*-
"""Qt 版主题：与 tkinter 版同一套设计令牌，集中在这里。"""
from __future__ import annotations

from pathlib import Path

from PySide6.QtGui import QFont, QFontDatabase

_ARROW = (Path(__file__).resolve().parent.parent / 'assets' / 'chevron.png').as_posix()
_ARROW_UP = (Path(__file__).resolve().parent.parent / 'assets' / 'chevron_up.png').as_posix()

# 设计令牌（与 gwtyper/ui/theme.py 对齐）
C = {
    "canvas": "#EDEFF3",
    "surface": "#FFFFFF",
    "sunken": "#F6F7F9",
    "inset": "#E3E7ED",
    "hover": "#F3F5F9",
    "active": "#E9EDF4",
    "row_even": "#FAFBFC",
    "border": "#E4E8EE",
    "border2": "#D3D9E2",
    "text": "#1B2230",
    "text2": "#5A6675",
    "text3": "#8A94A3",
    "primary": "#2563EB",
    "primary_hover": "#1D4ED8",
    "primary_press": "#1E40AF",
    "primary_soft": "#E8EFFC",
    "sel_bg": "#DCE7FB",
    "sel_fg": "#12305F",
    "ok": "#0E9F6E",
    "ok_tx": "#0A7A54",
    "ok_soft": "#E6F6F0",
    "warn": "#D97706",
    "warn_tx": "#9A5B08",
    "warn_soft": "#FDF3E4",
    "err": "#DC2626",
    "err_tx": "#B01C1C",
    "err_soft": "#FDECEC",
}

FAMILY = "Microsoft YaHei UI"


def pick_family() -> str:
    fams = set(QFontDatabase.families())
    for name in ("Microsoft YaHei UI", "微软雅黑", "Microsoft YaHei", "PingFang SC",
                 "Noto Sans CJK SC", "Noto Sans SC", "WenQuanYi Micro Hei", "SimHei"):
        if name in fams:
            return name
    return "Sans Serif"


def app_font(size: int = 10, bold: bool = False) -> QFont:
    f = QFont(FAMILY, size)
    f.setBold(bold)
    return f


def stylesheet() -> str:
    return f"""
    QWidget {{ background: {C['canvas']}; color: {C['text']}; }}
    /* 文字类控件不要自己刷底色，否则放在白卡片上会出现一块灰底 */
    QLabel, QCheckBox, QRadioButton {{ background: transparent; }}
    QLabel#H1, QLabel#H2, QLabel#Muted, QLabel#Hint {{ background: transparent; }}
    QMainWindow, QDialog {{ background: {C['canvas']}; }}

    /* 侧边栏 */
    QFrame#Sidebar {{ background: {C['surface']}; border-right: 1px solid {C['border']}; }}
    QWidget#NavBody {{ background: transparent; }}
    QFrame#Sidebar {{ background: {C['surface']}; border-right: 1px solid {C['border']}; }}
    QLabel#NavGroup {{ color: #7C8797; font-size: 11px; padding: 9px 14px 3px 14px;
                       background: transparent; }}
    QFrame#SidebarBottom {{ background: {C['surface']}; border-top: 1px solid {C['border']}; }}
    QPushButton#NavAbout {{
        text-align: left; padding: 7px 14px; border: none; border-radius: 6px;
        color: {C['text2']}; background: transparent;
    }}
    QPushButton#NavAbout:hover {{ background: {C['hover']}; color: {C['text']}; }}
    QPushButton#NavItem {{
        text-align: left; padding: 6px 14px; border: none; border-radius: 6px;
        color: {C['text']}; background: transparent;
    }}
    QPushButton#NavItem:hover {{ background: {C['hover']}; }}
    QPushButton#NavItem:checked {{
        background: {C['primary_soft']}; color: {C['primary']}; font-weight: bold;
        border-left: 3px solid {C['primary']}; padding-left: 11px;
    }}

    /* 卡片 */
    QFrame#Card {{ background: {C['surface']}; border: 1px solid {C['border']}; border-radius: 8px; }}
    QLabel#H1 {{ font-size: 17px; font-weight: bold; }}
    QLabel#H2 {{ font-size: 13px; font-weight: bold; }}
    QLabel#Muted {{ color: {C['text2']}; }}
    QLabel#Hint {{ color: {C['text3']}; font-size: 11px; }}

    /* 按钮 */
    QPushButton {{
        background: {C['surface']}; border: 1px solid {C['border2']}; border-radius: 6px;
        padding: 6px 14px;
    }}
    QPushButton:hover {{ background: {C['hover']}; }}
    QPushButton:disabled {{ color: {C['text3']}; background: {C['sunken']}; border-color: {C['border']}; }}
    QPushButton#Primary {{
        background: {C['primary']}; color: #FFFFFF; border: 1px solid {C['primary']};
    }}
    QPushButton#Primary:hover {{ background: {C['primary_hover']}; }}
    QPushButton#Primary:pressed {{ background: {C['primary_press']}; }}


    /* 下拉框 */
    QComboBox {{
        background: {C['surface']}; border: 1px solid {C['border2']}; border-radius: 6px;
        padding: 5px 8px 5px 10px; min-height: 18px;
    }}
    QComboBox:hover {{ border-color: #B9C2CF; }}
    QComboBox:focus {{ border-color: {C['primary']}; }}
    QComboBox::drop-down {{
        width: 24px; border: none; background: transparent;
    }}
    QComboBox::down-arrow {{ image: url("{_ARROW}"); width: 13px; height: 13px; margin-right: 6px; }}
    QComboBox QAbstractItemView::item {{ min-height: 28px; padding: 2px 8px; }}

    /* 数字框 */
    QSpinBox, QDoubleSpinBox {{
        background: {C['surface']}; border: 1px solid {C['border2']}; border-radius: 6px;
        padding: 4px 4px 4px 8px; min-height: 20px;
        selection-background-color: {C['primary_soft']}; selection-color: {C['primary']};
    }}
    QSpinBox:hover, QDoubleSpinBox:hover {{ border-color: #B9C2CF; }}
    QSpinBox:focus, QDoubleSpinBox:focus {{ border-color: {C['primary']}; }}
    QSpinBox::up-button, QDoubleSpinBox::up-button {{
        subcontrol-origin: border; subcontrol-position: top right;
        width: 20px; border: none; border-left: 1px solid {C['border']};
        border-top-right-radius: 6px; background: transparent;
    }}
    QSpinBox::down-button, QDoubleSpinBox::down-button {{
        subcontrol-origin: border; subcontrol-position: bottom right;
        width: 20px; border: none; border-left: 1px solid {C['border']};
        border-bottom-right-radius: 6px; background: transparent;
    }}
    QSpinBox::up-button:hover, QDoubleSpinBox::up-button:hover,
    QSpinBox::down-button:hover, QDoubleSpinBox::down-button:hover {{ background: {C['hover']}; }}
    QSpinBox::up-arrow, QDoubleSpinBox::up-arrow {{
        image: url("{_ARROW_UP}"); width: 10px; height: 10px;
    }}
    QSpinBox::down-arrow, QDoubleSpinBox::down-arrow {{
        image: url("{_ARROW}"); width: 10px; height: 10px;
    }}
    QComboBox QAbstractItemView {{
        background: {C['surface']}; border: 1px solid {C['border2']}; border-radius: 8px;
        padding: 5px; outline: none; min-width: 170px;
        selection-background-color: {C['primary_soft']}; selection-color: {C['primary']};
    }}
    QComboBox QAbstractItemView::item {{
        min-height: 28px; padding: 3px 10px; border-radius: 6px; color: {C['text']};
    }}
    QComboBox QAbstractItemView::item:hover {{ background: {C['hover']}; }}
    QComboBox QAbstractItemView::item:selected {{
        background: {C['primary_soft']}; color: {C['primary']};
    }}

    /* 表格 */
    QTableView, QTableWidget {{
        background: {C['surface']}; alternate-background-color: {C['row_even']};
        gridline-color: {C['border']}; border: 1px solid {C['border']}; border-radius: 6px;
        selection-background-color: {C['sel_bg']}; selection-color: {C['sel_fg']};
    }}
    QHeaderView::section {{
        background: {C['sunken']}; color: {C['text2']}; border: none;
        border-bottom: 1px solid {C['border']}; padding: 6px 8px; font-weight: bold;
    }}
    QListWidget {{
        background: {C['surface']}; border: 1px solid {C['border']}; border-radius: 6px;
        selection-background-color: {C['sel_bg']}; selection-color: {C['sel_fg']};
    }}
    QLineEdit, QComboBox, QSpinBox, QPlainTextEdit, QTextEdit {{
        background: {C['surface']}; border: 1px solid {C['border2']}; border-radius: 6px;
        padding: 4px 6px;
    }}

    /* 滚动条：细、圆角、悬停加深，不用系统那种灰块 */
    QScrollBar:vertical {{
        background: transparent; width: 10px; margin: 2px 1px 2px 1px; border: none;
    }}
    QScrollBar::handle:vertical {{
        background: {C['border2']}; min-height: 30px; border-radius: 4px;
    }}
    QScrollBar::handle:vertical:hover {{ background: #A9B4C2; }}
    QScrollBar::handle:vertical:pressed {{ background: {C['text3']}; }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0; width: 0; background: transparent; border: none;
    }}
    QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: transparent; }}

    QScrollBar:horizontal {{
        background: transparent; height: 10px; margin: 1px 2px 1px 2px; border: none;
    }}
    QScrollBar::handle:horizontal {{
        background: {C['border2']}; min-width: 30px; border-radius: 4px;
    }}
    QScrollBar::handle:horizontal:hover {{ background: #A9B4C2; }}
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
        height: 0; width: 0; background: transparent; border: none;
    }}
    QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{ background: transparent; }}

    QScrollArea {{ background: transparent; border: none; }}
    QScrollArea > QWidget > QWidget {{ background: transparent; }}
    QScrollArea > QWidget > QScrollBar {{ background: transparent; }}
    QStatusBar {{ background: {C['surface']}; border-top: 1px solid {C['border']}; color: {C['text2']}; }}
    QMenuBar {{ background: {C['surface']}; }}
    QMenuBar::item:selected {{ background: {C['hover']}; }}
    """


def install_combo_fix() -> None:
    """全局修下拉列表宽度：Qt 默认按输入框宽度开弹窗，
    条目一长就被省略成「15…PI」。开弹窗前按最长条目重设一次宽度。"""
    from PySide6.QtWidgets import QComboBox

    if getattr(QComboBox, "_gw_fixed", False):
        return
    orig = QComboBox.showPopup

    def _show(self):
        try:
            view = self.view()
            need = view.sizeHintForColumn(0) + 36
            view.setMinimumWidth(max(self.width(), need, 150))
        except Exception:
            pass
        orig(self)

    QComboBox.showPopup = _show
    QComboBox._gw_fixed = True
