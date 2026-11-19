# -*- coding: utf-8 -*-
"""文档角色模型：解析与渲染之间唯一的中间层。"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum


class Role(str, Enum):
    TITLE = "title"                  # 大标题 二号方正小标宋
    SUBTITLE = "subtitle"            # 副标题/肩题 三号楷体
    H1 = "h1"                        # 一级标题 一、 三号黑体
    H2 = "h2"                        # 二级标题 （一） 三号楷体
    H3 = "h3"                        # 三级标题 1. 三号仿宋加粗
    BODY = "body"                    # 正文 三号仿宋 28 磅固定
    LEAD = "lead"                    # 条目引导语「——坚持……」楷体引导
    TABLE_CAPTION = "table_caption"  # 表题
    NOTE = "note"                    # 表注 / 说明
    ATTACH = "attach"                # 附件说明 / 附件标记
    SIGNATURE = "signature"          # 署名/落款单位
    DATE = "date"                    # 成文日期
    BLANK = "blank"                  # 空段
    OTHER = "other"                  # 未归类


ROLE_CN: dict[str, str] = {
    "title": "大标题", "subtitle": "副标题", "h1": "一级标题", "h2": "二级标题",
    "h3": "三级标题", "body": "正文", "lead": "引导语", "table_caption": "表题",
    "note": "表注", "attach": "附件", "signature": "署名", "date": "日期",
    "blank": "空", "other": "未归类",
}


@dataclass
class RunFmt:
    text: str
    font: str | None = None
    font_ascii: str | None = None
    size_pt: float | None = None
    bold: bool | None = None


@dataclass
class Fmt:
    """段落级实测格式。单位 pt / cm / 字符。"""
    font: str | None = None          # 主 run 的 eastAsia 字体
    size_pt: float | None = None
    bold: bool | None = None
    align: str | None = None
    line_pt: float | None = None
    line_rule: str | None = None
    first_indent_pt: float | None = None
    first_indent_chars: float | None = None
    space_before_pt: float | None = None
    space_after_pt: float | None = None
    left_indent_pt: float | None = None
    right_indent_chars: float | None = None


@dataclass
class Para:
    index: int
    text: str
    fmt: Fmt
    runs: list[RunFmt] = field(default_factory=list)
    n_runs: int = 0
    has_drawing: bool = False
    role: Role = Role.OTHER
    confidence: float = 0.0
    evidence: list[str] = field(default_factory=list)
    mixed: bool = False              # run 之间字体/字号不一致
    lead_prefix: str = ""            # 段首楷体引导语片段（LEAD 用）

    @property
    def dominant(self) -> RunFmt | None:
        """字数最多的 run —— 决定段落主体字体，别被楷体引导语带偏。"""
        cs = [r for r in self.runs if r.text.strip()]
        return max(cs, key=lambda r: len(r.text)) if cs else None


@dataclass
class PageSpec:
    width_cm: float = 21.0
    height_cm: float = 29.7
    top_cm: float = 3.7
    bottom_cm: float = 3.5
    left_cm: float = 2.8
    right_cm: float = 2.6
    footer_cm: float | None = None


@dataclass
class Doc:
    path: str
    page: PageSpec
    paras: list[Para]
    n_tables: int = 0
    fonts_used: set[str] = field(default_factory=set)
    source: str = ""                 # 用户看到的那个文件（可能是 .doc，path 是转换后的 docx）

    @property
    def name(self) -> str:
        import os

        return os.path.basename(self.source or self.path)


@dataclass
class Finding:
    index: int | None
    role: str
    field: str
    actual: str
    expected: str
    level: str = "error"             # error 必改 / warn 建议 / info 提示
    note: str = ""

    def as_row(self) -> str:
        loc = f"第{self.index + 1}段" if self.index is not None else "整篇"
        return f"[{self.level}] {loc} {ROLE_CN.get(self.role, self.role)} · {self.field}：{self.actual} → 应为 {self.expected}"
