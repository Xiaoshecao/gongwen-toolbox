# -*- coding: utf-8 -*-
"""段落角色识别：特征打分，不猜。低于阈值进「待确认」。"""
from __future__ import annotations

import re

from .model import Para, Role
from .parse_docx import (RE_ATTACH, RE_DATE, RE_H1, RE_H2, RE_H3, RE_LEAD_PREFIX,
                         RE_NOTE, RE_TABLE_CAPTION)

HEI = {"黑体", "黑体_GB2312", "方正黑体简体", "方正黑体_GBK", "微软雅黑", "思源黑体"}
KAI_PREFIX = "楷体"
SONG_XBS = {"方正小标宋简体", "方正小标宋_GBK", "小标宋", "华文中宋"}
FANGSONG_PREFIX = "仿宋"

MAX_HEAD_CHARS = 45


def _f(p: Para) -> str:
    return p.fmt.font or ""


def classify_one(p: Para) -> Para:
    t = p.text.strip()
    f = p.fmt
    font = _f(p)
    size = f.size_pt or 0
    n = len(t)
    score: dict[Role, float] = {}
    ev: list[str] = []

    def add(role: Role, w: float, why: str):
        score[role] = score.get(role, 0) + w
        ev.append(f"{why}({w:+.1f})")

    if not t:
        p.role, p.confidence, p.evidence = Role.BLANK, 1.0, ["空段"]
        return p

    short = n <= MAX_HEAD_CHARS
    ends_dot = t[-1] in "。；！？"
    centered = f.align == "center"
    right = f.align == "right"

    # —— 表题 / 表注 / 附件：优先判定，避免误判成大标题
    if RE_TABLE_CAPTION.match(t) and n <= 40:
        add(Role.TABLE_CAPTION, 8, "表/图编号开头")
    if RE_NOTE.match(t):
        add(Role.NOTE, 8, "注：开头")
    if RE_ATTACH.match(t) and n <= 30:
        add(Role.ATTACH, 8, "附件标记")
    if size and size <= 14 and n > 30:
        add(Role.NOTE, 3, "小字号长段")

    # —— 条目引导语（同段内楷体引导 + 仿宋正文）
    if p.lead_prefix:
        add(Role.LEAD, 9, "段首楷体引导语")
    if RE_LEAD_PREFIX.match(t) and n > 20 and not ends_dot:
        add(Role.LEAD, 2, "一是/——开头")

    # —— 对齐与位置
    if right and RE_DATE.match(t):
        add(Role.DATE, 9, "右对齐+日期形态")
    if right and short and not ends_dot:
        add(Role.SIGNATURE, 5, "右对齐短句")
        add(Role.DATE, 1, "右对齐短句")
    if centered and size >= 20:
        add(Role.TITLE, 6, f"居中{size:g}磅")
    if centered and 14 <= size < 20:
        add(Role.SUBTITLE, 3, "居中中小号")
    if centered and p.index <= 5 and not RE_TABLE_CAPTION.match(t):
        add(Role.SUBTITLE, 2, "靠前居中")

    # —— 编号形态
    if RE_H1.match(t):
        add(Role.H1, 7, "一、编号")
    if RE_H2.match(t):
        add(Role.H2, 7, "（一）编号")
    if RE_H3.match(t):
        add(Role.H3, 6, "1.编号")
    if RE_DATE.match(t):
        add(Role.DATE, 8, "日期形态")

    # —— 字体（公文层级主要靠字体承载，最强信号）
    if font in SONG_XBS:
        if size >= 20:
            add(Role.TITLE, 6, f"字体{font}+{size:g}磅")
        else:
            add(Role.TABLE_CAPTION, 5, f"字体{font}但只有{size:g}磅")
    if font in HEI:
        add(Role.H1, 5, f"字体{font}")
        if short and not ends_dot:
            add(Role.H1, 2, "短且无句号")
    if font.startswith(KAI_PREFIX):
        add(Role.H2, 2, f"字体{font}")
    if font.startswith(FANGSONG_PREFIX):
        add(Role.BODY, 3, f"字体{font}")

    # —— 长度/标点
    if not short:
        add(Role.BODY, 3, "长段")
    if ends_dot:
        add(Role.BODY, 3, "句号收尾")
        add(Role.H1, -2, "标题不以句号收尾")
        add(Role.H2, -2, "标题不以句号收尾")
    if short and not ends_dot:
        add(Role.H1, 1, "短且无句号")

    if not score:
        p.role, p.confidence, p.evidence = Role.OTHER, 0.2, ["无特征"]
        return p

    role, top = max(score.items(), key=lambda kv: kv[1])
    pos = sum(v for v in score.values() if v > 0) or 1
    conf = round(min(1.0, top / pos + min(top, 9) / 18), 2)
    p.role, p.confidence, p.evidence = role, conf, ev
    return p


def classify(paras: list[Para], threshold: float = 0.45,
             overrides: dict[int, str] | None = None) -> tuple[list[Para], list[Para]]:
    """overrides: {段落 index: 角色名} —— 用户在界面上手工指定的，一律优先。"""
    overrides = overrides or {}
    auto, review = [], []
    for p in paras:
        if p.index in overrides:
            try:
                p.role = Role(overrides[p.index])
            except ValueError:
                classify_one(p)
            else:
                p.confidence = 1.0
                p.evidence = ["手动指定"]
                auto.append(p)
                continue
        classify_one(p)
        if p.role == Role.BLANK or p.confidence >= threshold:
            auto.append(p)
        else:
            review.append(p)
    return auto, review
