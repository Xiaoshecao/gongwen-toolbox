# -*- coding: utf-8 -*-
"""生成公文习惯的表格 Word（签到表 / 材料交接单），不依赖任何 GUI 库。"""
from __future__ import annotations

from pathlib import Path

def write_table_docx(path: Path, kind: str, title: str, sub: str,
                      day: str, names: list[str]) -> None:
    """按公文习惯排一张能直接打印的表。"""
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt

    d = Document()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin, sec.bottom_margin = Cm(3.7), Cm(3.5)
    sec.left_margin, sec.right_margin = Cm(2.8), Cm(2.6)
    normal = d.styles["Normal"]
    normal.font.name = "仿宋_GB2312"
    normal.font.size = Pt(16)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), "仿宋_GB2312")

    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title)
    run.font.name = "方正小标宋简体"
    run.element.rPr.rFonts.set(qn("w:eastAsia"), "方正小标宋简体")
    run.font.size = Pt(22)
    if sub:
        sp = d.add_paragraph()
        sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sr = sp.add_run(sub)
        sr.font.name = "楷体_GB2312"
        sr.element.rPr.rFonts.set(qn("w:eastAsia"), "楷体_GB2312")
        sr.font.size = Pt(14)

    if kind == "sign":
        heads = ["序号", "姓名 / 单位", "职务", "签名", "备注"]
        widths = [1.4, 4.6, 3.0, 3.6, 3.0]
    else:
        heads = ["序号", "材料名称", "份数", "备注"]
        widths = [1.4, 8.6, 2.2, 3.4]

    table = d.add_table(rows=1, cols=len(heads))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(heads):
        cell = table.rows[0].cells[i]
        cell.text = ""
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = para.add_run(h)
        r.font.name = "黑体"
        r.element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
        r.font.size = Pt(12)
        cell.width = Cm(widths[i])

    for idx, n in enumerate(names, start=1):
        row = table.add_row().cells
        vals = ([str(idx), n, "", "", ""] if kind == "sign"
                else [str(idx), n, "", ""])
        for j, v in enumerate(vals):
            row[j].text = ""
            para = row[j].paragraphs[0]
            para.alignment = (WD_ALIGN_PARAGRAPH.CENTER if j != 1
                              else WD_ALIGN_PARAGRAPH.LEFT)
            r = para.add_run(v)
            r.font.name = "仿宋_GB2312"
            r.element.rPr.rFonts.set(qn("w:eastAsia"), "仿宋_GB2312")
            r.font.size = Pt(12)
            row[j].width = Cm(widths[j])

    # 表格行高放宽一点，手写签名不挤
    for row in table.rows:
        row.height = Cm(0.95)

    tail = d.add_paragraph()
    tail.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tr = tail.add_run(day)
    tr.font.name = "仿宋_GB2312"
    tr.element.rPr.rFonts.set(qn("w:eastAsia"), "仿宋_GB2312")
    tr.font.size = Pt(14)
    path.parent.mkdir(parents=True, exist_ok=True)
    d.save(str(path))
