# -*- coding: utf-8 -*-
"""参照件学习：读一份已排好的公文，反推出一份规则集（derived.yaml）。

用法：python -m gwtyper learn 参照件.docx -o rules/derived.yaml
思路：先按通用分类器认出每段角色，再按角色聚合实测格式，取众数为规则值。
"""
from __future__ import annotations

from collections import Counter, defaultdict
from pathlib import Path

import yaml

from . import classify
from .model import Doc, Para, Role
from .parse_docx import parse

# 只学这些属性；对齐/缩进等结构性强的也学，但首行缩进优先用字符
FIELDS = ["font", "size_pt", "align", "line_pt", "line_rule",
          "first_line_chars", "space_before_pt", "space_after_pt", "bold"]


def _mode(values):
    vals = [v for v in values if v not in (None, "", "none")]
    if not vals:
        return None
    c = Counter(vals)
    return c.most_common(1)[0][0]


def _para_indent_chars(p: Para) -> float | None:
    f = p.fmt
    if f.first_indent_chars is not None:
        return f.first_indent_chars
    if f.first_indent_pt:
        size = f.size_pt or 16
        return round(f.first_indent_pt / size, 2)
    return 0.0 if f.first_indent_pt is None else None


def learn(path: str | Path, name: str | None = None) -> dict:
    doc: Doc = parse(path)
    classify.classify(doc.paras)

    buckets: dict[Role, list[Para]] = defaultdict(list)
    for p in doc.paras:
        if p.role == Role.BLANK or not p.text.strip():
            continue
        buckets[p.role].append(p)

    roles_out: dict[str, dict] = {}
    for role, ps in buckets.items():
        if role in (Role.OTHER,):
            continue
        spec: dict = {}
        for field in FIELDS:
            if field == "font":
                v = _mode([p.dominant.font if p.dominant else p.fmt.font for p in ps])
            elif field == "size_pt":
                v = _mode([p.dominant.size_pt if p.dominant else p.fmt.size_pt for p in ps])
            elif field == "first_line_chars":
                v = _mode([_para_indent_chars(p) for p in ps])
            elif field == "bold":
                v = _mode([p.dominant.bold if p.dominant else p.fmt.bold for p in ps])
            else:
                v = _mode([getattr(p.fmt, field) for p in ps])
            if v is not None:
                spec[field] = v
        # 行距若众数是 exactly，则只保留 line_pt；否则 line_rule 单倍之类无意义
        if spec.get("line_rule") != "exactly":
            spec.pop("line_pt", None)
        if spec.get("line_rule") in (None, "none"):
            spec.pop("line_rule", None)
        if spec.get("bold") is False:
            spec.pop("bold")
        roles_out[role.value] = spec

    return {
        "name": name or f"参照件：{Path(path).stem}",
        "source": str(Path(path).name),
        "page": {
            "size": "A4",
            "top_cm": doc.page.top_cm,
            "bottom_cm": doc.page.bottom_cm,
            "left_cm": doc.page.left_cm,
            "right_cm": doc.page.right_cm,
        },
        "roles": roles_out,
    }


def save(profile: dict, out: str | Path) -> Path:
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    header = ("# 由参照件自动推导，请人工复核后使用。\n"
              "# 生成：python -m gwtyper learn <参照件.docx>\n")
    body = yaml.safe_dump(profile, allow_unicode=True, sort_keys=False, default_flow_style=False)
    out.write_text(header + body, encoding="utf-8")
    return out
