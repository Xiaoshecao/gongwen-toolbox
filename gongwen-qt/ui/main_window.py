# -*- coding: utf-8 -*-
"""主窗口：左侧竖排菜单 + 右侧页面堆栈；「关于」在帮助菜单与侧边栏底部两处。"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QHBoxLayout, QLabel, QMainWindow, QStackedWidget,
                               QStatusBar, QVBoxLayout, QWidget)

from . import theme
from .about import AboutDialog
from .sidebar import Sidebar

NAV = [
    ("公文处理", ["公文排版", "表格排版", "公文模板库"]),
    ("文档处理", ["格式转换", "OCR 识别"]),
    ("办公小工具", ["批量重命名", "字数统计", "扫描件纠偏", "签到表·交接单", "图片压缩", "文件查重"]),
    ("PDF 工具", ["PDF 工具箱"]),
    ("批量处理", ["批量加水印", "敏感信息脱敏", "数据汇总", "打印体检"]),
    ("材料管理", ["材料上交", "名单比对"]),
]


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        from .about import _app_version
        self.setWindowTitle(f"公文工具箱 {_app_version()}（Qt 版）")
        self.resize(1360, 860)

        central = QWidget()
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        head = QWidget()
        hv = QVBoxLayout(head)
        hv.setContentsMargins(20, 14, 20, 10)
        hv.setSpacing(2)
        t = QLabel("公文工具箱")
        t.setObjectName("H1")
        sub = QLabel("公文排版 · 模板库 · 格式转换 · OCR · 办公小工具 · PDF 工具 · 批量处理 · 材料管理")
        sub.setObjectName("Muted")
        hv.addWidget(t)
        hv.addWidget(sub)
        root.addWidget(head)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(0)
        self.sidebar = Sidebar(NAV)
        self.stack = QStackedWidget()

        from pages.checklist_page import ChecklistPage
        from pages.convert_page import ConvertPage
        from pages.format_page import FormatPage
        from pages.ocr_page import OcrPage
        from pages.table_page import TablePage
        from pages.template_page import TemplatePage
        from pages.tools_pages import DeskewPage, FormPage, RenamePage, StatPage
        from pages.batch_pages import (DupPage, ImageCompressPage, ListDiffPage,
                                       MaskPage, MergeTablePage, PrintCheckPage,
                                       WatermarkPage)
        from pages.pdf_tools_page import PdfToolsPage
        from pages.simple_page import SimplePage

        pages = []
        desc = {
            "公文模板库": "7 种公文模板（通知/请示/报告/函/会议纪要/情况汇报/证明），填字段直接出 Word。",
            "格式转换": "Word/Excel/PPT/图片/PDF 互转，含横竖版、留白、合并带书签。",
            "OCR 识别": "扫描件识别成文字/Word/双层 PDF，识别结果可先改再落盘。",
            "PDF 工具箱": "加密、去密码、瘦身、按页提取。",
            "材料上交": "按名单核对各村（单位）材料交没交，可导出未交清单。",
        }
        for group, names in NAV:
            for name in names:
                if name == "公文排版":
                    pages.append(FormatPage())
                elif name == "格式转换":
                    pages.append(ConvertPage())
                elif name == "OCR 识别":
                    pages.append(OcrPage())
                elif name == "表格排版":
                    pages.append(TablePage())
                elif name == "公文模板库":
                    pages.append(TemplatePage())
                elif name == "批量重命名":
                    pages.append(RenamePage())
                elif name == "字数统计":
                    pages.append(StatPage())
                elif name == "扫描件纠偏":
                    pages.append(DeskewPage())
                elif name == "签到表·交接单":
                    pages.append(FormPage())
                elif name == "图片压缩":
                    pages.append(ImageCompressPage())
                elif name == "文件查重":
                    pages.append(DupPage())
                elif name == "PDF 工具箱":
                    pages.append(PdfToolsPage())
                elif name == "批量加水印":
                    pages.append(WatermarkPage())
                elif name == "敏感信息脱敏":
                    pages.append(MaskPage())
                elif name == "数据汇总":
                    pages.append(MergeTablePage())
                elif name == "打印体检":
                    pages.append(PrintCheckPage())
                elif name == "名单比对":
                    pages.append(ListDiffPage())
                elif name == "材料上交":
                    pages.append(ChecklistPage())
                else:
                    pages.append(SimplePage(name, desc.get(name, "")))
        for p in pages:
            self.stack.addWidget(p)

        row.addWidget(self.sidebar)
        holder = QWidget()
        hl = QVBoxLayout(holder)
        hl.setContentsMargins(14, 6, 14, 10)
        hl.addWidget(self.stack)
        row.addWidget(holder, 1)
        root.addLayout(row, 1)
        self.setCentralWidget(central)

        self.sidebar.currentChanged.connect(self.stack.setCurrentIndex)
        self.sidebar.aboutClicked.connect(self.show_about)

        sb = QStatusBar()
        sb.showMessage("就绪")
        self.setStatusBar(sb)

        # 默认停在第一页
        self.sidebar.select(0)

    def show_about(self):
        AboutDialog(self).exec()
