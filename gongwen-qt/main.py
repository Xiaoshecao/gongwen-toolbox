# -*- coding: utf-8 -*-
"""公文工具箱 Qt 预览版入口。

复用 gwtyper/core 的纯 Python 内核（解析/规则/检查/转换/OCR），界面换 PySide6。
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if getattr(sys, "frozen", False):          # PyInstaller 打包后
    BASE = Path(getattr(sys, "_MEIPASS", ROOT))
    REPO = BASE
else:
    BASE = ROOT
    REPO = ROOT.parent                     # D:\公文处理
for _p in (str(REPO), str(BASE)):
    if _p not in sys.path:
        sys.path.insert(0, _p)             # 让 `import gwtyper.core` / `import ui.xxx` 可用

from PySide6.QtWidgets import QApplication  # noqa: E402

from ui import theme  # noqa: E402


def main() -> int:
    app = QApplication(sys.argv)
    theme.FAMILY = theme.pick_family()
    app.setFont(theme.app_font(10))
    app.setStyleSheet(theme.stylesheet())
    theme.install_combo_fix()

    # Windows 任务栏要单独声明 AppUserModelID，否则任务栏显示的是 python 的图标、
    # 而且会和别的 python 程序并成一组。
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "Xiaoshecao.GongwenToolbox.4"
            )
        except Exception:  # noqa: BLE001
            pass

    # 图标：Windows 用 .ico（多尺寸、任务栏/标题栏都认），其他平台优先 .png
    try:
        from PySide6.QtGui import QIcon
        names = ("icon.ico", "icon.png") if sys.platform == "win32" else ("icon.png", "icon.ico")
        icon = QIcon()
        for d in (BASE / "assets", BASE / "gwtyper" / "assets",
                  REPO / "gwtyper" / "assets"):
            for nm in names:
                f = d / nm
                if f.exists():
                    icon = QIcon(str(f))
                    if not icon.isNull():
                        break
            if not icon.isNull():
                break
        if not icon.isNull():
            app.setWindowIcon(icon)
    except Exception:  # noqa: BLE001
        pass

    from ui.main_window import MainWindow
    w = MainWindow()
    w.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
