# -*- coding: utf-8 -*-
"""统一的消息泵：把后台线程塞进 queue 的消息逐条交给界面处理。

三条规矩（都是踩过坑才写进来的）：
1. 每条消息单独 try —— 一条坏消息不能把整个泵打死，否则界面此后再也不更新；
2. 无论处理成功还是失败，都要重新注册 after —— 否则泵停了，进度条不动、
   按钮永久灰、后台结果再也取不出来（console=False 时用户一个提示都看不到）；
3. 空队列不是错误，直接返回。
"""
from __future__ import annotations

import logging
import queue


def drain(q: "queue.Queue", handler) -> int:
    """把队列清空，逐条交给 handler。返回处理条数。"""
    n = 0
    while True:
        try:
            msg = q.get_nowait()
        except queue.Empty:
            return n
        n += 1
        try:
            handler(msg)
        except Exception:  # noqa: BLE001
            logging.exception("处理消息失败：%r", (msg[:1] if isinstance(msg, tuple) else msg))


def start(widget, q: "queue.Queue", handler, interval: int = 120) -> None:
    """在 widget 上起一个定时泵。"""

    def tick() -> None:
        drain(q, handler)
        try:
            widget.after(interval, tick)
        except Exception:  # noqa: BLE001  —— 窗口已销毁
            pass

    widget.after(interval, tick)
