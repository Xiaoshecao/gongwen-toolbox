# -*- coding: utf-8 -*-
"""中文字号 ↔ 磅值对照（GB/T 9704 用的那几档）。

中文「号数」是印刷界的老叫法，和磅值不是线性的：三号=16磅、小三=15磅、
四号=14磅、小四=12磅…… 公文里常用的是：
  标题 二号(22磅) / 正文与各级标题 三号(16磅) / 表格 小四或五号
基层文书大多只会说「三号仿宋」，不知道 16 磅。所以界面上优先显示号数。
"""
from __future__ import annotations

# (中文名, 磅值)，按磅值从大到小
_SIZES: list[tuple[str, float]] = [
    ("初号", 42.0),
    ("小初", 36.0),
    ("一号", 26.0),
    ("小一", 24.0),
    ("二号", 22.0),
    ("小二", 18.0),
    ("三号", 16.0),
    ("小三", 15.0),
    ("四号", 14.0),
    ("小四", 12.0),
    ("五号", 10.5),
    ("小五", 9.0),
    ("六号", 7.5),
    ("小六", 6.5),
    ("七号", 5.5),
    ("八号", 5.0),
]

# 磅值 → 中文名（容差 0.3 磅，够吸收 15.9999 这种浮点误差）
_BY_PT: list[tuple[float, str]] = [(pt, name) for name, pt in _SIZES]
TOL = 0.35


def name_of(pt: float | None) -> str:
    """16.0 → '三号'。对不上任何标准号数时返回空串。"""
    if pt is None:
        return ""
    for value, name in _BY_PT:
        if abs(float(pt) - value) <= TOL:
            return name
    return ""


def pt_of(name: str) -> float | None:
    """'三号' → 16.0。认不出返回 None。"""
    for cn, pt in _SIZES:
        if cn == name:
            return pt
    return None


def label(pt: float | None, with_pt: bool = True) -> str:
    """界面显示用：'三号（16磅）'。

    with_pt=False 时只给 '三号'；对不上号数的自定义字号（比如 17 磅）
    退回 '17 磅'，不然用户不知道是什么。
    """
    if pt is None:
        return ""
    name = name_of(pt)
    num = f"{pt:g}"
    if not name:
        return f"{num} 磅"
    return f"{name}（{num}磅）" if with_pt else name


def choices() -> list[tuple[str, float]]:
    """给下拉框用的完整列表（字号名, 磅值）。"""
    return list(_SIZES)
