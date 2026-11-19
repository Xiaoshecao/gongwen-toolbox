# -*- coding: utf-8 -*-
"""公文模板库：填空就出一份符合 GB/T 9704 版式的稿子。

模板都是本地的 python-docx 生成，不联网、不调模型。
"""
from __future__ import annotations

from datetime import date
from pathlib import Path

TITLE_FONT = "方正小标宋简体"
BODY_FONT = "仿宋_GB2312"
H_FONT = "黑体"
KAI = "楷体_GB2312"

# key: (名称, 说明, [(字段key, 显示名, 默认值, 是否多行)])
TEMPLATES: dict[str, dict] = {
    "notice": {
        "name": "通知",
        "desc": "关于××的通知：主送单位、事项、时间地点、要求",
        "fields": [
            ("title", "标题（关于……的通知）", "关于做好当前工作的通知", False),
            ("to", "主送单位", "", False),
            ("body", "正文（可多行，一段一行）", "", True),
            ("issuer", "落款单位", "", False),
            ("day", "成文日期", "", False),
        ],
    },
    "request": {
        "name": "请示",
        "desc": "请示事项：理由 + 具体请求",
        "fields": [
            ("title", "标题（关于……的请示）", "关于申请工作经费的请示", False),
            ("to", "主送机关", "", False),
            ("reason", "理由（可多行）", "", True),
            ("ask", "请示事项", "以上请示，请予批准。", False),
            ("issuer", "落款单位", "", False),
            ("day", "成文日期", "", False),
        ],
    },
    "report": {
        "name": "报告",
        "desc": "工作情况报告：做法、成效、下一步",
        "fields": [
            ("title", "标题（关于……的报告）", "关于近期工作情况的报告", False),
            ("to", "主送机关", "", False),
            ("basic", "基本情况（可多行）", "", True),
            ("next", "下一步打算（可多行）", "", True),
            ("issuer", "落款单位", "", False),
            ("day", "成文日期", "", False),
        ],
    },
    "letter": {
        "name": "函",
        "desc": "平级或不相隶属单位之间商洽工作",
        "fields": [
            ("title", "标题（关于……的函）", "关于商请协助有关事项的函", False),
            ("to", "主送单位", "", False),
            ("body", "正文（可多行）", "", True),
            ("issuer", "落款单位", "", False),
            ("day", "成文日期", "", False),
        ],
    },
    "minutes": {
        "name": "会议纪要",
        "desc": "会议基本情况 + 议定事项",
        "fields": [
            ("title", "标题（……会议纪要）", "工作会议纪要", False),
            ("when", "时间", "", False),
            ("where", "地点", "", False),
            ("host", "主持人", "", False),
            ("clerk", "记录人", "", False),
            ("who", "参加人员", "", False),
            ("topics", "会议议定事项（可多行，一行一条）", "", True),
            ("issuer", "落款单位", "", False),
            ("day", "成文日期", "", False),
        ],
    },
    "brief": {
        "name": "情况汇报",
        "desc": "向上级汇报某项工作：情况、做法、成效",
        "fields": [
            ("title", "标题（关于……的情况汇报）", "关于重点工作推进情况的汇报", False),
            ("to", "主送机关", "", False),
            ("basic", "基本情况（可多行）", "", True),
            ("effect", "主要成效（可多行）", "", True),
            ("next", "下一步（可多行）", "", True),
            ("issuer", "落款单位", "", False),
            ("day", "成文日期", "", False),
        ],
    },
    "certify": {
        "name": "证明",
        "desc": "兹证明××：姓名、身份证号、事由",
        "fields": [
            ("name", "姓名", "", False),
            ("idno", "身份证号", "", False),
            ("matter", "证明事项（可多行）", "", True),
            ("issuer", "落款单位（盖章）", "", False),
            ("day", "成文日期", "", False),
        ],
    },
}


def _set_run(run, font: str, size: float, bold: bool = False) -> None:
    from docx.oxml.ns import qn
    from docx.shared import Pt

    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    for attr in ("w:eastAsia", "w:ascii", "w:hAnsi"):
        rfonts.set(qn(attr), font)


def _p(doc, text, font=BODY_FONT, size=16, bold=False, align="justify",
       indent=2, line=28, before=0, after=0):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt

    p = doc.add_paragraph()
    p.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER,
                   "right": WD_ALIGN_PARAGRAPH.RIGHT}.get(align, WD_ALIGN_PARAGRAPH.JUSTIFY)
    pf = p.paragraph_format
    pf.line_spacing = Pt(line)
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if indent:
        pf.first_line_indent = Pt(size * indent)
    r = p.add_run(text)
    _set_run(r, font, size, bold)
    return p


def _new_doc():
    from docx import Document
    from docx.shared import Cm

    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)
    s.top_margin, s.bottom_margin = Cm(3.7), Cm(3.5)
    s.left_margin, s.right_margin = Cm(2.8), Cm(2.6)
    return doc


def _lines(text: str) -> list[str]:
    return [ln.strip() for ln in str(text or "").splitlines() if ln.strip()]


def _sign(doc, issuer: str, day: str) -> None:
    d = day.strip() or f"{date.today().year}年{date.today().month}月{date.today().day}日"
    _p(doc, "", indent=0)
    if issuer.strip():
        _p(doc, issuer.strip(), font=KAI, align="right", indent=0)
    _p(doc, d, font=KAI, align="right", indent=0)


def build(key: str, out, v: dict) -> Path:
    kinds = TEMPLATES[key]["name"]
    doc = _new_doc()

    if key == "certify":
        _p(doc, "证　明", font=TITLE_FONT, size=22, align="center", indent=0, line=32, after=18)
        body = []
        who = v.get("name", "").strip()
        idno = v.get("idno", "").strip()
        first = "兹证明"
        if who:
            first += who
        if idno:
            first += f"（身份证号：{idno}）"
        first += "，"
        body.append(first)
        body += _lines(v.get("matter"))
        for i, ln in enumerate(body):
            _p(doc, ln)
        _sign(doc, v.get("issuer", ""), v.get("day", ""))
    elif key == "minutes":
        _p(doc, v.get("title", "会议纪要"), font=TITLE_FONT, size=22, align="center", indent=0, line=32, after=18)
        head = [("时间", v.get("when")), ("地点", v.get("where")), ("主持人", v.get("host")),
                ("记录人", v.get("clerk")), ("参加人员", v.get("who"))]
        for label, val in head:
            if str(val or "").strip():
                _p(doc, f"{label}：{str(val).strip()}", indent=0)
        _p(doc, "", indent=0)
        _p(doc, "会议议定事项如下：", indent=2)
        for i, ln in enumerate(_lines(v.get("topics")), 1):
            _p(doc, f"{i}. {ln}" if not ln[:1].isdigit() else ln)
        _sign(doc, v.get("issuer", ""), v.get("day", ""))
    else:
        _p(doc, v.get("title", kinds), font=TITLE_FONT, size=22, align="center", indent=0, line=32, after=18)
        to = str(v.get("to") or "").strip()
        if to:
            _p(doc, to if to.endswith("：") else to + "：", indent=0)
        order = ["body", "reason", "ask", "basic", "next", "effect"]
        for f in order:
            for ln in _lines(v.get(f)):
                _p(doc, ln)
        # 请示固定收尾
        if key == "request" and not _lines(v.get("ask")):
            _p(doc, "以上请示，请予批准。")
        _sign(doc, v.get("issuer", ""), v.get("day", ""))

    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    return out
