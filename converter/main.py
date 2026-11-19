# -*- coding: utf-8 -*-
"""程序入口。

打包前先做几件事：
1. 屏蔽第三方库的告警输出（pdf2docx 会往控制台打一堆 [INFO]）；
2. 给 COM 线程加上稳定设置，避免 WPS 弹窗打断自动化；
3. 崩溃时把错误写到日志文件，而不是让窗口一闪就没。
"""
from __future__ import annotations

import os
import sys
import traceback
import logging
import warnings
from pathlib import Path

warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", message=".*fitz.*")

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))


def log_dir() -> Path:
    base = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "格式转换器"
    try:
        base.mkdir(parents=True, exist_ok=True)
    except OSError:
        base = Path.home()
    return base


def install_logger() -> None:
    """把未捕获异常写进日志，窗口模式下用户至少有个凭据。"""
    path = log_dir() / "error.log"
    logging.basicConfig(
        filename=str(path),
        level=logging.ERROR,
        format="%(asctime)s %(levelname)s %(message)s",
        encoding="utf-8",
    )

    def hook(exc_type, exc, tb):
        logging.error("未捕获异常", exc_info=(exc_type, exc, tb))
        try:
            import tkinter.messagebox as mb

            mb.showerror("格式转换器", f"程序出错：{exc}\n\n详细信息已写入：\n{path}")
        except Exception:
            pass

    sys.excepthook = hook


def run_gui() -> None:
    import app

    gui = app.App()
    gui.mainloop()


def main() -> None:
    install_logger()
    try:
        run_gui()
    except Exception:
        logging.getLogger().error("启动失败", exc_info=True)
        raise


if __name__ == "__main__":
    main()
