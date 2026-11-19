# -*- coding: utf-8 -*-
"""把规则套回 docx（排版动作，会写文件）。只在副本上操作，绝不覆盖原件。"""
from __future__ import annotations

import shutil
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from . import classify
from .model import Doc, Para, Role
from .parse_docx import parse

ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "distribute": WD_ALIGN_PARAGRAPH.DISTRIBUTE,
}
ASCII_FONT = "Times New Roman"


def _set_run_font(run, font: str | None, size_pt: float | None,
                  bold: bool | None = None, ascii_font: str = ASCII_FONT) -> None:
    if font:
        run.font.name = ascii_font
        rpr = run._element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.insert(0, rf)
        rf.set(qn("w:eastAsia"), font)
        rf.set(qn("w:ascii"), ascii_font)
        rf.set(qn("w:hAnsi"), ascii_font)
    if size_pt:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.font.bold = bold
    # 清掉主题字体，否则 eastAsia 会被 theme 覆盖
    rpr = run._element.rPr
    if rpr is not None:
        rf = rpr.find(qn("w:rFonts"))
        if rf is not None:
            for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
                if rf.get(qn(a)) is not None:
                    del rf.attrib[qn(a)]


def _set_indent(p, chars: float | None, size_pt: float, right_chars: float | None) -> None:
    pf = p.paragraph_format
    ppr = p._element.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if chars is not None:
        if ind is None:
            ind = OxmlElement("w:ind")
            ppr.append(ind)
        if chars <= 0:
            for a in ("w:firstLine", "w:firstLineChars"):
                if ind.get(qn(a)) is not None:
                    del ind.attrib[qn(a)]
            ind.set(qn("w:firstLineChars"), "0")
            ind.set(qn("w:firstLine"), "0")
            pf.first_line_indent = Pt(0)
        else:
            ind.set(qn("w:firstLineChars"), str(int(round(chars * 100))))
            ind.set(qn("w:firstLine"), str(int(round(chars * size_pt * 20))))
            pf.first_line_indent = Pt(chars * size_pt)
    if right_chars is not None and ind is not None:
        ind.set(qn("w:rightChars"), str(int(round(right_chars * 100))))
        ind.set(qn("w:right"), str(int(round(right_chars * size_pt * 20))))


def _no_snap(p) -> None:
    """关掉网格对齐，否则 28 磅固定行距会被 docGrid 吸成 28.8。"""
    ppr = p._element.get_or_add_pPr()
    sn = ppr.find(qn("w:snapToGrid"))
    if sn is None:
        sn = OxmlElement("w:snapToGrid")
        ppr.append(sn)
    sn.set(qn("w:val"), "0")


def _set_page(doc, page: dict) -> None:
    for sec in doc.sections:
        for key, attr in (("top_cm", "top_margin"), ("bottom_cm", "bottom_margin"),
                          ("left_cm", "left_margin"), ("right_cm", "right_margin")):
            if key in page:
                setattr(sec, attr, Cm(page[key]))
        if page.get("footer_cm"):
            sec.footer_distance = Cm(page["footer_cm"])
        # 去掉文档网格吸附
        sec_el = sec._sectPr
        grid = sec_el.find(qn("w:docGrid"))
        if grid is not None:
            grid.set(qn("w:type"), "none")


def _style_para(p, spec: dict, dom_size: float, role: Role, para: Para) -> None:
    font = spec.get("font")
    size = spec.get("size_pt") or dom_size
    bold = spec.get("bold")

    pf = p.paragraph_format
    if spec.get("align") in ALIGN:
        pf.alignment = ALIGN[spec["align"]]
    if spec.get("line_pt"):
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(spec["line_pt"])
        _no_snap(p)
    if spec.get("space_before_pt") is not None:
        pf.space_before = Pt(spec["space_before_pt"])
    if spec.get("space_after_pt") is not None:
        pf.space_after = Pt(spec["space_after_pt"])
    _set_indent(p, spec.get("first_line_chars"), size, spec.get("right_indent_chars"))

    # 清除段落样式带来的字体影响：直接格式覆盖
    runs = [r for r in p.runs if r.text.strip()]
    if not runs:
        return

    if role == Role.LEAD and para.lead_prefix:
        # 楷体引导语 + 正文字体
        body_font = None
        for r in runs:
            if r.text.strip() and para.lead_prefix in r.text:
                _set_run_font(r, font, size, bold)
                body_font = "仿宋_GB2312"
            else:
                _set_run_font(r, body_font or spec.get("body_font") or "仿宋_GB2312", size, bold)
        return

    for r in runs:
        _set_run_font(r, font, size, bold)


def _field_run(par, instr: str, size_pt: float, font: str) -> None:
    """插入一个域（如 PAGE），公文页码必须用域才能自动编号。"""
    run = par.add_run()
    _set_run_font(run, font, size_pt)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), instr)
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "1"
    r.append(t)
    fld.append(r)
    par._p.append(fld)


def _set_page_number(doc, features: dict) -> None:
    """页脚页码：— 1 —，四号宋体，居中。GB/T 9704 要求一字线包夹。"""
    if not features.get("page_number"):
        return
    fmt = features.get("page_number_fmt") or "— {n} —"
    pre, _, post = fmt.partition("{n}")
    size = features.get("page_number_size_pt") or 14
    font = features.get("page_number_font") or "宋体"
    for sec in doc.sections:
        foot = sec.footer
        foot.is_linked_to_previous = False
        par = foot.paragraphs[0] if foot.paragraphs else foot.add_paragraph()
        for r in list(par.runs):
            r._element.getparent().remove(r._element)
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.paragraph_format.first_line_indent = Pt(0)
        if pre:
            r = par.add_run(pre)
            _set_run_font(r, font, size)
        _field_run(par, " PAGE ", size, font)
        if post:
            r = par.add_run(post)
            _set_run_font(r, font, size)


def _style_table(doc, tspec: dict) -> int:
    """表格内文字：GB/T 9704 只要求「与正文一致或小一号」，默认五号仿宋居中。"""
    if not tspec:
        return 0
    font = tspec.get("font") or "仿宋_GB2312"
    size = tspec.get("size_pt") or 14
    align = ALIGN.get(tspec.get("align") or "center", WD_ALIGN_PARAGRAPH.CENTER)
    n = 0
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if not p.text.strip():
                        continue
                    p.paragraph_format.alignment = align
                    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                    p.paragraph_format.first_line_indent = Pt(0)
                    for r in p.runs:
                        _set_run_font(r, font, size)
                    n += 1
    return n


def apply(src: str | Path, rules: dict, out: str | Path | None = None,
          threshold: float = 0.45, overrides: dict[int, str] | None = None
          ) -> tuple[Path, Doc, list[Para]]:
    """排版并另存。默认输出 <原名>-排版后.docx，不覆盖原文件。

    老格式（.doc/.wps）会先转 docx 再处理，输出一律是 .docx。
    """
    src = Path(src)
    out = Path(out) if out else src.with_name(f"{src.stem}-排版后.docx")
    if out.resolve() == src.resolve():
        raise ValueError("输出路径不能与原文件相同")

    doc: Doc = parse(src)
    auto, review = classify.classify(doc.paras, threshold, overrides)
    roles = rules.get("roles", {})
    skip = {x.index for x in review}

    d = Document(doc.path)
    if rules.get("page"):
        _set_page(d, rules["page"])
    _set_page_number(d, rules.get("features") or {})
    if rules.get("table"):
        _style_table(d, rules["table"])

    tot = 0
    for para, parsed in zip(d.paragraphs, doc.paras):
        spec = roles.get(parsed.role.value)
        if not spec or parsed.role == Role.BLANK:
            continue
        if parsed.index in skip:
            continue
        dom_size = (parsed.dominant.size_pt if parsed.dominant and parsed.dominant.size_pt
                    else parsed.fmt.size_pt or 16)
        _style_para(para, spec, dom_size, parsed.role, parsed)
        tot += 1

    d.save(str(out))
    return out, doc, review


def copy_side_by_side(src: Path, out: Path) -> None:
    shutil.copy2(src, out)
