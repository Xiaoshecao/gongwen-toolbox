# -*- coding: utf-8 -*-
"""「关于」对话框 —— 独立入口（帮助菜单 + 侧边栏底部），不埋在二级菜单里。"""
from __future__ import annotations

import platform
import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QDesktopServices
from PySide6.QtCore import QUrl
from PySide6.QtWidgets import (QDialog, QDialogButtonBox, QFrame, QLabel,
                               QPushButton, QTextBrowser, QVBoxLayout)

from . import theme

REPO_URL = "https://github.com/Xiaoshecao/gongwen-toolbox"
RELEASES_URL = REPO_URL + "/releases"


def _root() -> Path:
    """仓库根目录（有 LICENSE 的地方）。"""
    here = Path(__file__).resolve()
    for p in [here] + list(here.parents):
        if (p / "LICENSE").exists():
            return p
    return here.parents[2]


def _app_version() -> str:
    """优先读 gongwen-qt/VERSION，其次复用 gwtyper 的版本号。"""
    for p in [Path(__file__).resolve().parents[1] / "VERSION"]:
        if p.exists():
            try:
                return p.read_text(encoding="utf-8").strip()
            except OSError:
                pass
    try:
        from gwtyper.app import APP_VER
        return APP_VER
    except Exception:  # noqa: BLE001
        return "4.0"


def platform_text() -> str:
    """按当前运行的平台写，不写死。"""
    machine = (platform.machine() or "").lower()
    if sys.platform.startswith("win"):
        return "Windows 10 / 11（64 位）"
    if sys.platform == "darwin":
        return f"macOS（{machine or 'arm64'}）"
    if sys.platform.startswith("linux"):
        if machine in ("aarch64", "arm64"):
            return "Linux · ARM64（麒麟 V10 / 统信 UOS 等国产系统）"
        if machine in ("x86_64", "amd64"):
            return "Linux · x86_64"
        return f"Linux · {machine or '未知架构'}"
    return sys.platform


class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("关于 公文工具箱")
        self.setMinimumWidth(560)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 18, 20, 16)
        lay.setSpacing(10)

        title = QLabel("公文工具箱")
        title.setObjectName("H1")
        lay.addWidget(title)

        ver = QLabel(f"版本 {_app_version()}　·　{platform_text()}")
        ver.setObjectName("Muted")
        lay.addWidget(ver)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet(f"color: {theme.C['border']};")
        lay.addWidget(line)

        text = QTextBrowser()
        text.setOpenExternalLinks(True)
        text.setMinimumHeight(230)
        text.setHtml(f"""
        <p><b>给基层办公室用的公文工具。</b>排版检查按 GB/T 9704 国标，格式标准可自定义；
        另外带格式转换、OCR 识别、批量处理、材料上交台账等日常功能。</p>
        <p><b>许可证：</b>GNU Affero 通用公共许可证第 3 版（AGPL-3.0-or-later）。<br>
        你可以自由使用、修改、再分发；<b>分发（含修改版）时须同样以 AGPL-3.0 开源并保留声明</b>。
        仅在自己单位内部使用不产生额外义务。</p>
        <p><b>第三方组件：</b>完整清单见仓库
        <code>THIRD_PARTY_NOTICES.md</code>。主要几项：<br>
        · <b>PyMuPDF</b>：AGPL-3.0 或 Artifex 商业双许可；<br>
        · <b>PySide6 / Qt 6</b>（本版界面框架）：LGPL-3.0，动态链接，可按 LGPL 要求替换库文件；<br>
        · <b>Qt PDF</b>（版面预览）：GPL-3.0，与本项目的 AGPL-3.0 兼容；<br>
        · Linux 包内置的 <b>LibreOffice</b>：MPL-2.0（文件级 copyleft，不影响本程序）；<br>
        · <b>文泉驿字体</b>：GPL-2.0 + 字体嵌入例外。</p>
        <p><b>源码与下载：</b><a href="{REPO_URL}">{REPO_URL}</a><br>
        <a href="{RELEASES_URL}">{RELEASES_URL}</a></p>
        """)
        lay.addWidget(text, 1)

        box = QDialogButtonBox()
        btn_license = QPushButton("查看完整许可证")
        btn_license.clicked.connect(self.open_license)
        btn_src = QPushButton("打开源码页")
        btn_src.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(REPO_URL)))
        btn_ok = QPushButton("关闭")
        btn_ok.setObjectName("Primary")
        btn_ok.clicked.connect(self.accept)
        box.addButton(btn_license, QDialogButtonBox.ActionRole)
        box.addButton(btn_src, QDialogButtonBox.ActionRole)
        box.addButton(btn_ok, QDialogButtonBox.AcceptRole)
        lay.addWidget(box)

    def open_license(self) -> None:
        lic = _root() / "LICENSE"
        if lic.exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(lic)))
        else:
            QDesktopServices.openUrl(QUrl(REPO_URL + "/blob/main/LICENSE"))
