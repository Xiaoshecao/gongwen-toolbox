# -*- coding: utf-8 -*-
"""公文工具箱入口。"""
from __future__ import annotations

import logging
import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore", category=FutureWarning)

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))


def log_dir() -> Path:
    from gwtyper.core import sysenv

    try:
        return sysenv.data_dir()
    except OSError:
        return Path.home()


def install_logger() -> Path:
    path = log_dir() / "error.log"
    logging.basicConfig(filename=str(path), level=logging.ERROR,
                        format="%(asctime)s %(levelname)s %(message)s", encoding="utf-8")

    def hook(exc_type, exc, tb):
        logging.error("未捕获异常", exc_info=(exc_type, exc, tb))
        try:
            import tkinter.messagebox as mb
            mb.showerror("公文工具箱", f"程序出错：{exc}\n\n详细信息已写入：\n{path}")
        except Exception:
            pass

    sys.excepthook = hook
    return path


def run_cli() -> int:
    from gwtyper.core.cli import main as cli_main
    return cli_main(sys.argv[1:])


_CLI_COMMANDS = ("check", "format", "learn", "pdf", "doctor",
                 "double", "-h", "--help")


def run_selftest(rest: list) -> int:
    """在真实界面里跑一遍 OCR 识别流程，结果写文件（窗口版没有控制台）。

    为什么要这个：exe 是窗口程序，`sys.stdout` 是空的，界面回调里抛的
    异常连一行都看不到。OCR 页签就因为一个元组解包错误整整废了一版，
    症状只是「点按钮没反应」。留一个能自证的入口，出问题一跑就知道。
    """
    import tempfile
    import time

    install_logger()
    from gwtyper.ui import theme

    theme.enable_dpi_awareness()
    from gwtyper.app import App

    src = Path(rest[0]).resolve() if rest else None
    lines: list[str] = []

    def log(s: str) -> None:
        lines.append(s)
        try:
            print(s, flush=True)
        except Exception:  # noqa: BLE001
            pass

    out_dir = Path(tempfile.mkdtemp(prefix="gwtyper-selftest-"))
    app = App()
    app.withdraw()

    problems: list[str] = []
    app.report_callback_exception = (
        lambda t, e, tb: problems.append(f"{t.__name__}: {e}"))

    try:
        import tkinter.messagebox as mb

        for name in ("showinfo", "showwarning", "showerror"):
            setattr(mb, name, lambda *a, **k: None)

        log(f"版本    : {app.title()}")
        log(f"页签    : {app.nb.tabs()}")
        try:
            from .app import _HAS_DND

            by_tk = any(getattr(c, "__module__", "").startswith("tkinterdnd2")
                        for c in type(app).__mro__)
            log(f"拖拽    : {'可用' if (_HAS_DND and by_tk) else '不可用'}"
                f"（模块 {_HAS_DND}，窗口基类 {by_tk}）")
        except Exception as exc:  # noqa: BLE001
            log(f"拖拽    : 检测失败 {exc}")
        log(f"转换引擎: {app.engine_hint if hasattr(app, 'engine_hint') else ''}")

        pane = app.ocr_pane
        for key in ("txt", "docx_layout", "docx_text", "double"):
            pane.mode.set(key)
            info = pane._mode_info()
            log(f"  模式 {key:12s} -> {info[1]:16s} {info[2]:6s} 长度 {len(info)}")
            if len(info) != 4 or info[0] != key:
                problems.append(f"模式 {key} 的 _mode_info 不对：{info}")

        if src and src.exists():
            pane.mode.set("txt")
            pane.out_dir.set(str(out_dir))
            pane._add_paths([src])
            # 识别完会弹编辑窗；自检里换成立刻保存的假窗，避免自检卡在模态窗上
            from gwtyper.ui import ocr_pane as opmod

            class _AutoSave:
                def __init__(self, master, name, text, ext_label, on_save,
                             on_cancel=None, index=0, total=1, readonly=False):
                    if on_save:
                        on_save(text)

            opmod.ResultEditor = _AutoSave
            app.update()
            log(f"文件    : {src.name}（列表 {len(pane.files)} 个，"
                f"按钮 {pane.btn_go.cget('state')}）")
            pane.run_ocr()
            t0 = time.time()
            while pane.busy and time.time() - t0 < 300:
                app.update()
                time.sleep(0.1)
            made = list(out_dir.glob("*"))
            log(f"状态    : {pane.status.get()}")
            log(f"产物    : {[p.name for p in made]}")
            if made:
                txt = made[0].read_text(encoding="utf-8", errors="replace")
                log(f"字数    : {len(txt)}")
                log(f"开头    : {txt[:60]!r}")
            else:
                problems.append("没有生成任何文件")

            # 再验一遍「识别后可编辑」这条链：recognize 不落盘 + text_to_docx
            try:
                from gwtyper.core import ocr_text as octxt

                rec = octxt.recognize(src, dpi=150)
                edit_out = out_dir / "编辑后.docx"
                octxt.text_to_docx(rec["text"] + "\n自检补一句", edit_out)
                ok_edit = edit_out.exists() and edit_out.stat().st_size > 0
                log(f"可编辑  : 识别 {rec['chars']} 字 → "
                    f"{'docx 生成成功' if ok_edit else 'docx 没生成'}")
                if not ok_edit:
                    problems.append("可编辑链路：text_to_docx 没产出文件")
            except Exception as exc:  # noqa: BLE001
                problems.append(f"可编辑链路 {type(exc).__name__}: {exc}")
        else:
            log("（没给文件，只做了模式检查）")
    except Exception as exc:  # noqa: BLE001
        import traceback

        problems.append(f"{type(exc).__name__}: {exc}")
        log(traceback.format_exc())
    finally:
        try:
            app.destroy()
        except Exception:  # noqa: BLE001
            pass

    log(f"回调异常: {problems if problems else '无'}")
    log("结论    : " + ("OK" if not problems else "有问题"))

    try:
        from gwtyper.core import sysenv

        report = sysenv.data_dir() / "selftest.txt"
    except Exception:  # noqa: BLE001
        report = Path(tempfile.gettempdir()) / "gwtyper-selftest.txt"
    report.write_text("\n".join(lines), encoding="utf-8")
    return 0 if not problems else 1


def main() -> None:
    if len(sys.argv) > 1 and sys.argv[1] == "--selftest":
        sys.exit(run_selftest(sys.argv[2:]))
    if len(sys.argv) > 1 and sys.argv[1] in _CLI_COMMANDS:
        sys.exit(run_cli())
    install_logger()
    from gwtyper.ui import theme

    theme.enable_dpi_awareness()      # 必须在建窗口之前
    from gwtyper.app import App
    App().mainloop()


if __name__ == "__main__":
    main()
