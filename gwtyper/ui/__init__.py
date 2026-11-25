# -*- coding: utf-8 -*-
"""tkinter 界面包。

注意：Qt 版（gongwen-qt）只复用 ``gwtyper.core``，不导入本包；
这里对 theme 做容错导入，避免在没有 tkinter 的环境（如精简的 Linux 运行时）导入即崩。
"""
try:  # pragma: no cover
    from . import theme  # noqa: F401
except Exception:  # noqa: BLE001
    theme = None  # type: ignore[assignment]
