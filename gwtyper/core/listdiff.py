# -*- coding: utf-8 -*-
"""两份名单比对：上级下发的名单 vs 自己手里的名单。"""
from __future__ import annotations

import re

from . import villages

_SPLIT = re.compile(r"[\n\r,，、;；\t]+")


def parse(text_or_path) -> list[str]:
    """把一串文本（或一个 txt/csv 文件路径）拆成名字列表，去序号、去重。"""
    from pathlib import Path

    text = ""
    p = Path(str(text_or_path))
    if p.exists() and p.is_file():
        raw = p.read_bytes()
        for enc in ("utf-8-sig", "utf-8", "gb18030", "utf-16"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
    else:
        text = str(text_or_path)
    out: list[str] = []
    seen: set[str] = set()
    for line in _SPLIT.split(text):
        name = villages.strip_ordinal(line.strip())
        if not name or not villages._looks_like_name(name):
            continue
        if name not in seen:
            seen.add(name)
            out.append(name)
    return out


def _dups(names: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    for n in names:
        seen[n] = seen.get(n, 0) + 1
    return [n for n, c in seen.items() if c > 1]


def diff(a: list[str], b: list[str]) -> dict:
    """a = 上级名单，b = 自己名单。"""
    sa, sb = set(a), set(b)
    return {
        "only_a": [n for n in a if n not in sb],   # 上级有、自己少（最要紧）
        "only_b": [n for n in b if n not in sa],   # 自己多出来的
        "same": [n for n in a if n in sb],
        "dup_a": _dups(a),
        "dup_b": _dups(b),
    }


def to_text(a, b, result: dict, title_a="上级名单", title_b="我的名单") -> str:
    lines = [f"{title_a} {len(a)} 个，{title_b} {len(b)} 个。", ""]
    oa, ob = result["only_a"], result["only_b"]
    lines.append(f"【{title_a}有、{title_b}没有】（{len(oa)} 个）—— 优先补这些：")
    lines += [f"  {i}. {n}" for i, n in enumerate(oa, 1)] or ["  无"]
    lines.append("")
    lines.append(f"【{title_b}多出来的】（{len(ob)} 个）—— 核对是不是写错了：")
    lines += [f"  {i}. {n}" for i, n in enumerate(ob, 1)] or ["  无"]
    if result["dup_a"] or result["dup_b"]:
        lines.append("")
        lines.append("【重复】")
        if result["dup_a"]:
            lines.append(f"  {title_a}：" + "、".join(result["dup_a"]))
        if result["dup_b"]:
            lines.append(f"  {title_b}：" + "、".join(result["dup_b"]))
    lines.append("")
    lines.append(f"两边一致 {len(result['same'])} 个。")
    return "\n".join(lines)
