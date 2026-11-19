# -*- coding: utf-8 -*-
"""打包入口（Windows exe 与 Linux AppImage 共用）。

源码运行时把项目根目录挂进 sys.path；PyInstaller 打出来的包里
`gwtyper` 已经是内置模块，不需要这些补救。
"""
import sys
from pathlib import Path

if not getattr(sys, "frozen", False):
    root = Path(__file__).resolve().parents[2]     # D:\公文处理
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))

from gwtyper.main import main  # noqa: E402

if __name__ == "__main__":
    main()
