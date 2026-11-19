# -*- coding: utf-8 -*-
"""解析 docx -> Doc。只读，不改文件。"""
from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from . import doc_source
from .model import Doc, Fmt, PageSpec, Para, RunFmt

_ALIGN = {0: "left", 1: "center", 2: "right", 3: "justify", 4: "distribute"}
_LINE_RULE = {0: "single", 1: "at_least", 2: "double", 3: "multiple", 4: "exactly"}


def _east_asia(run) -> str | None:
    rpr = run._element.rPr
    if rpr is None:
        return None
    rf = rpr.rFonts
    return rf.get(qn("w:eastAsia")) if rf is not None else None


def _runs(p) -> list[RunFmt]:
    out = []
    for r in p.runs:
        out.append(RunFmt(
            text=r.text,
            font=_east_asia(r),
            font_ascii=r.font.name,
            size_pt=round(r.font.size.pt, 2) if r.font.size else None,
            bold=r.bold,
        ))
    return out


def _style_fallback(p) -> tuple[str | None, float | None]:
    """段落样式上的字体字号（直接格式没写时用）。"""
    try:
        st = p.style
        el = st.element
        font = None
        if el.rPr is not None and el.rPr.rFonts is not None:
            font = el.rPr.rFonts.get(qn("w:eastAsia"))
        size = round(st.font.size.pt, 2) if st.font.size else None
        return font, size
    except Exception:
        return None, None


def _para_fmt(p) -> Fmt:
    pf = p.paragraph_format
    f = Fmt(
        align=_ALIGN.get(int(pf.alignment)) if pf.alignment is not None else None,
        line_rule=_LINE_RULE.get(int(pf.line_spacing_rule)) if pf.line_spacing_rule is not None else None,
        first_indent_pt=round(pf.first_line_indent.pt, 2) if pf.first_line_indent is not None else None,
        space_before_pt=round(pf.space_before.pt, 2) if pf.space_before is not None else None,
        space_after_pt=round(pf.space_after.pt, 2) if pf.space_after is not None else None,
        left_indent_pt=round(pf.left_indent.pt, 2) if pf.left_indent is not None else None,
    )
    ls = pf.line_spacing
    if ls is not None:
        f.line_pt = round(ls.pt, 2) if hasattr(ls, "pt") else round(float(ls), 3)

    ind = p._element.pPr.find(qn("w:ind")) if p._element.pPr is not None else None
    if ind is not None:
        if ind.get(qn("w:firstLineChars")):
            f.first_indent_chars = round(int(ind.get(qn("w:firstLineChars"))) / 100, 2)
        if ind.get(qn("w:rightChars")):
            f.right_indent_chars = round(int(ind.get(qn("w:rightChars"))) / 100, 2)
        if ind.get(qn("w:right")) and f.right_indent_chars is None:
            v = int(ind.get(qn("w:right")))
            f.right_indent_chars = round(v / 20 / 16, 2)  # twips -> 磅 -> 三号字宽近似
    return f


def parse(path: str | Path) -> Doc:
    """老格式（.doc/.wps/.rtf）会先悄悄转成 docx 再读，source 保留原文件名。"""
    src = Path(path)
    real = doc_source.to_docx(src)
    try:
        d = Document(str(real))
    except Exception as exc:  # noqa: BLE001
        raise doc_source.DocSourceError(
            f"这个文件打不开（{type(exc).__name__}），可能已损坏：{src.name}") from exc
    if d.sections is None or not len(d.sections):
        raise doc_source.DocSourceError(f"这个文件里没有正文，读不出页面设置：{src.name}")
    sec = d.sections[0]
    page = PageSpec(
        width_cm=round(sec.page_width.cm, 2),
        height_cm=round(sec.page_height.cm, 2),
        top_cm=round(sec.top_margin.cm, 2),
        bottom_cm=round(sec.bottom_margin.cm, 2),
        left_cm=round(sec.left_margin.cm, 2),
        right_cm=round(sec.right_margin.cm, 2),
        footer_cm=round(sec.footer_distance.cm, 2) if sec.footer_distance else None,
    )

    paras: list[Para] = []
    fonts: set[str] = set()
    st_font, st_size = (None, None)

    for i, p in enumerate(d.paragraphs):
        f = _para_fmt(p)
        runs = _runs(p)
        body = [r for r in runs if r.text.strip()]
        if body:
            dom = max(body, key=lambda r: len(r.text))
            f.font, f.size_pt, f.bold = dom.font, dom.size_pt, dom.bold
            mixed = len({(r.font, r.size_pt) for r in body}) > 1
            fonts.update(r.font for r in body if r.font)
        else:
            mixed = False

        if f.font is None or f.size_pt is None:
            sf, ss = _style_fallback(p)
            f.font = f.font or sf
            f.size_pt = f.size_pt or ss

        # 楷体引导语：段首楷体 run 后接非楷体 run
        lead = ""
        if len(body) >= 2:
            first, rest = body[0], body[1]
            if first.font and first.font.startswith("楷体") and first.font != (rest.font or ""):
                if first.text.rstrip()[-1:] in "。：；" and len(first.text) <= 40:
                    lead = first.text.strip()

        has_draw = bool(p._element.findall(".//" + qn("w:drawing")))
        paras.append(Para(index=i, text=p.text, fmt=f, runs=runs, n_runs=len(p.runs),
                          has_drawing=has_draw, mixed=mixed, lead_prefix=lead))

    return Doc(path=str(real), page=page, paras=paras, n_tables=len(d.tables),
               fonts_used=fonts, source=str(src))


# —— 编号与形态正则
RE_H1 = re.compile(r"^[一二三四五六七八九十]{1,3}、")
RE_H2 = re.compile(r"^（[一二三四五六七八九十]{1,3}）|^\([一二三四五六七八九十]{1,3}\)")
RE_H3 = re.compile(r"^\d{1,2}[.．、](?!\d)")
RE_DATE = re.compile(
    r"^[0-9〇一二三四五六七八九十]{2,4}\s*年\s*\d{1,2}\s*月(\s*\d{1,2}\s*日)?$|^\d{4}-\d{1,2}-\d{1,2}$")
RE_TABLE_CAPTION = re.compile(r"^(表|附表|上图|下图|图)\s*\d*")
RE_NOTE = re.compile(r"^(注|说明|备注)\s*[：:]")
RE_ATTACH = re.compile(r"^附件\s*[：:0-9一二三四五六七八九十]?\s*$|^附件\s*[：:]")
RE_LEAD_PREFIX = re.compile(r"^(一是|二是|三是|四是|五是|六是|——)")
