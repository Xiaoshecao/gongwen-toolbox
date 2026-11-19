# -*- coding: utf-8 -*-
"""按台账生成《材料上报情况通报》Word。"""
from __future__ import annotations

from datetime import date
from pathlib import Path

TITLE_FONT = "方正小标宋简体"
BODY_FONT = "仿宋_GB2312"
KAI = "楷体_GB2312"


def _set_run(run, font: str, size_pt: float, bold: bool = False) -> None:
    from docx.oxml.ns import qn
    from docx.shared import Pt

    run.font.name = font
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), font)
    rfonts.set(qn("w:ascii"), font)
    rfonts.set(qn("w:hAnsi"), font)


def _para(doc, text: str, font=BODY_FONT, size=16, bold=False, align=None,
          indent_chars=2, line_pt=28, space_before=0, space_after=0):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt

    p = doc.add_paragraph()
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.line_spacing = Pt(line_pt)
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if indent_chars:
        pf.first_line_indent = Pt(size * indent_chars)
    run = p.add_run(text)
    _set_run(run, font, size, bold)
    return p


def build(path, title: str, names: list[str], status: dict, deadline: str = "",
          unit: str = "各村（单位）", issuer: str = "", day: str = "") -> Path:
    """status = {name: {"done": True/False}}"""
    from docx import Document
    from docx.shared import Cm

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin, sec.bottom_margin = Cm(3.7), Cm(3.5)
    sec.left_margin, sec.right_margin = Cm(2.8), Cm(2.6)

    _para(doc, title, font=TITLE_FONT, size=22, align="center", indent_chars=0,
          line_pt=32, space_after=18)

    _para(doc, f"{unit}：", indent_chars=0)
    d = day or f"{date.today().year}年{date.today().month}月{date.today().day}日"
    done = [n for n in names if (status.get(n) or {}).get("done")]
    todo = [n for n in names if not (status.get(n) or {}).get("done")]
    _para(doc, f"现将截至{d}，材料上报情况通报如下：")
    _para(doc, f"应交 {len(names)} 个，已交 {len(done)} 个，未交 {len(todo)} 个。")
    if todo:
        _para(doc, "未交的单位（%d 个）：%s。" % (len(todo), "、".join(todo)))
        tail = f"请以上单位于{deadline}前报送" if deadline else "请以上单位抓紧报送"
        _para(doc, f"{tail}，逾期将另行通报。")
    else:
        _para(doc, "各单位均已报送，特此通报表扬。")

    _para(doc, "", indent_chars=0)  # 空一行
    _para(doc, issuer or "", font=KAI, align="right", indent_chars=0)
    _para(doc, d, font=KAI, align="right", indent_chars=0)

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))
    return path
