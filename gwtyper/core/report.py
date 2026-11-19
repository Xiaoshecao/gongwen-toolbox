# -*- coding: utf-8 -*-
"""把体检结果写成可读报告（文本 / Markdown）。CLI 与界面共用。"""
from __future__ import annotations

from collections import Counter
from datetime import datetime
from pathlib import Path

from .model import Doc, Finding, Para, ROLE_CN, Role

LEVEL_CN = {"error": "必改", "warn": "建议", "info": "提示"}
ORDER = ["title", "subtitle", "h1", "h2", "h3", "lead", "body",
         "table_caption", "note", "attach", "signature", "date", "other"]


def score(findings: list[Finding]) -> tuple[int, int, int]:
    """健康分：从 100 往下扣。必改扣得多，建议扣得少，底分 0。

    只是给人一个"这份稿子大致什么水平"的直观量级，不参与任何判定。
    """
    n_err = sum(1 for f in findings if f.level == "error")
    n_warn = sum(1 for f in findings if f.level == "warn")
    pts = 100 - min(70, 7 * n_err) - min(25, 2 * n_warn)
    return max(0, pts), n_err, n_warn


def score_cn(pts: int) -> str:
    if pts >= 95:
        return "很规整"
    if pts >= 85:
        return "良好"
    if pts >= 70:
        return "基本可用"
    if pts >= 50:
        return "需要修改"
    return "问题较多"


def _role_counts(paras: list[Para]) -> list[tuple[str, int]]:
    c = Counter(p.role.value for p in paras if p.text.strip())
    rows = [(r, c[r]) for r in ORDER if c.get(r)]
    return rows


def build(doc: Doc, findings: list[Finding], review: list[Para],
          profile_name: str = "", title: str = "公文格式体检报告") -> str:
    n_err = sum(1 for f in findings if f.level == "error")
    n_warn = sum(1 for f in findings if f.level == "warn")
    L: list[str] = []
    L.append(title)
    L.append("=" * 46)
    L.append(f"文件：{doc.name}")
    L.append(f"时间：{datetime.now():%Y-%m-%d %H:%M}")
    L.append(f"规则集：{profile_name or '—'}")
    _pts, _ne, _nw = score(findings)
    L.append(f"健康分：{_pts} 分（{score_cn(_pts)}）"
             f"　必改 {_ne} 项，建议 {_nw} 项")
    L.append("")
    L.append("一、基本情况")
    L.append(f"  段落 {len([p for p in doc.paras if p.text.strip()])} 段｜表格 {doc.n_tables} 个")
    L.append(f"  页边距 上{doc.page.top_cm:g} 下{doc.page.bottom_cm:g} "
             f"左{doc.page.left_cm:g} 右{doc.page.right_cm:g} cm")
    L.append(f"  字体 {'、'.join(sorted(doc.fonts_used)) or '（无）'}")
    L.append("")
    L.append("二、段落角色分布")
    rows = _role_counts(doc.paras)
    L.append("  " + "｜".join(f"{ROLE_CN.get(r, r)} {n}" for r, n in rows) if rows else "  无")
    L.append("")
    L.append(f"三、问题清单（必改 {n_err}／建议 {n_warn}）")
    if not findings:
        L.append("  未发现问题。")
    else:
        by_role: dict[str, list[Finding]] = {}
        for f in findings:
            by_role.setdefault(f.role, []).append(f)
        for role, items in by_role.items():
            L.append(f"  [{ROLE_CN.get(role, role)}] {len(items)} 处")
            for f in items:
                loc = f"第{f.index + 1}段" if f.index is not None else "整篇"
                L.append(f"    ({LEVEL_CN.get(f.level, f.level)}) {loc} {f.field}："
                         f"{f.actual} → 应为 {f.expected}")
    if review:
        L.append("")
        L.append(f"四、待确认（{len(review)} 段，置信度不足，排版时跳过）")
        for p in review:
            L.append(f"  第{p.index + 1}段 {p.confidence:.2f}「{p.text[:36]}」")
            L.append(f"    证据：{'，'.join(p.evidence[:4]) or '—'}")
    L.append("")
    L.append("—" * 46)
    L.append("本报告只读生成，未改动原文件。")
    return "\n".join(L)


def _w(s: str) -> int:
    """显示宽度：全角算 2。"""
    return sum(2 if ord(ch) > 0x2E7F else 1 for ch in s)


def _pad(s: str, width: int, right: bool = False) -> str:
    fill = " " * max(0, width - _w(s))
    return fill + s if right else s + fill


def build_batch(entries: list[tuple[Doc, list[Finding], list[Para]]],
                profile_name: str = "") -> str:
    """多份文件汇总：先总表，再逐份明细。"""
    tot_e = sum(1 for _d, fs, _r in entries for f in fs if f.level == "error")
    tot_w = sum(1 for _d, fs, _r in entries for f in fs if f.level == "warn")
    L = ["公文格式批量体检汇总", "=" * 46,
         f"时间：{datetime.now():%Y-%m-%d %H:%M}",
         f"规则集：{profile_name or '—'}",
         f"文件 {len(entries)} 份｜合计 必改 {tot_e}／建议 {tot_w}", ""]
    L.append(_pad("文件", 42) + _pad("必改", 6, True) + _pad("建议", 6, True)
             + _pad("待确认", 8, True))
    L.append("-" * 62)
    for doc, fs, rv in entries:
        e = sum(1 for f in fs if f.level == "error")
        w = sum(1 for f in fs if f.level == "warn")
        name = doc.name
        if _w(name) > 40:
            while _w(name) > 39:
                name = name[:-1]
            name += "…"
        L.append(_pad(name, 42) + _pad(str(e), 6, True) + _pad(str(w), 6, True)
                 + _pad(str(len(rv)), 8, True))
    for doc, fs, rv in entries:
        if not fs and not rv:
            continue
        L.append("")
        L.append("=" * 46)
        L.extend(build(doc, fs, rv, profile_name, title=doc.name).splitlines()[3:])
    return "\n".join(L)


# ---------------------------------------------------------------- Word / PDF

def _md_doc(doc: Doc, findings: list[Finding], review: list[Para],
            profile_name: str, title: str):
    """把一份报告造成 python-docx 文档（不落盘，方便复用）。"""
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt

    d = Document()
    st = d.styles["Normal"]
    st.font.name = "仿宋_GB2312"
    st.font.size = Pt(12)
    st.element.rPr.rFonts.set(
        __import__("docx").oxml.ns.qn("w:eastAsia"), "仿宋_GB2312")

    h = d.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run(title)
    r.bold = True
    r.font.size = Pt(18)

    n_err = sum(1 for f in findings if f.level == "error")
    n_warn = sum(1 for f in findings if f.level == "warn")
    meta = [
        f"文件：{doc.name}",
        f"时间：{datetime.now():%Y-%m-%d %H:%M}",
        f"规则集：{profile_name or '—'}",
        f"健康分：{score(findings)[0]} 分（{score_cn(score(findings)[0])}）",
    ]
    for m in meta:
        d.add_paragraph(m)

    d.add_paragraph("一、基本情况")
    d.add_paragraph(
        f"段落 {len([p for p in doc.paras if p.text.strip()])} 段｜表格 {doc.n_tables} 个；"
        f"页边距 上{doc.page.top_cm:g} 下{doc.page.bottom_cm:g} "
        f"左{doc.page.left_cm:g} 右{doc.page.right_cm:g} cm")
    d.add_paragraph("字体：" + ("、".join(sorted(doc.fonts_used)) or "（无）"))

    d.add_paragraph("二、段落角色分布")
    rows = _role_counts(doc.paras)
    d.add_paragraph("｜".join(f"{ROLE_CN.get(r, r)} {n}" for r, n in rows) if rows else "无")

    d.add_paragraph(f"三、问题清单（必改 {n_err}／建议 {n_warn}）")
    if not findings:
        d.add_paragraph("未发现问题。")
    else:
        t = d.add_table(rows=1, cols=5)
        t.style = "Table Grid"
        for i, txt in enumerate(("等级", "位置", "角色", "检查项", "实际情况 → 应为")):
            t.rows[0].cells[i].text = txt
        for f in findings:
            loc = f"第{f.index + 1}段" if f.index is not None else "整篇"
            cells = t.add_row().cells
            cells[0].text = LEVEL_CN.get(f.level, f.level)
            cells[1].text = loc
            cells[2].text = ROLE_CN.get(f.role, f.role)
            cells[3].text = f.field
            cells[4].text = f"{f.actual} → {f.expected}"

    if review:
        d.add_paragraph(f"四、待确认（{len(review)} 段，置信度不足，排版时跳过）")
        for p in review:
            d.add_paragraph(f"第{p.index + 1}段 {p.confidence:.2f}「{p.text[:36]}」")

    d.add_paragraph("")
    d.add_paragraph("本报告只读生成，未改动原文件。")
    return d


def build_docx(path, doc: Doc, findings: list[Finding], review: list[Para],
               profile_name: str = "", title: str = "公文格式体检报告"):
    d = _md_doc(doc, findings, review, profile_name, title)
    d.save(str(path))
    return path


def build_batch_docx(path, entries: list[tuple[Doc, list[Finding], list[Para]]],
                     profile_name: str = ""):
    """多份汇总的 Word 版：一张总表 + 每份的问题表。"""
    from docx import Document
    from docx.shared import Pt

    d = Document()
    d.styles["Normal"].font.size = Pt(11)
    tot_e = sum(1 for _d, fs, _r in entries for f in fs if f.level == "error")
    tot_w = sum(1 for _d, fs, _r in entries for f in fs if f.level == "warn")
    p = d.add_paragraph()
    r = p.add_run("公文格式批量体检汇总")
    r.bold = True
    r.font.size = Pt(18)
    d.add_paragraph(f"时间：{datetime.now():%Y-%m-%d %H:%M}　规则集：{profile_name or '—'}")
    d.add_paragraph(f"文件 {len(entries)} 份｜合计 必改 {tot_e}／建议 {tot_w}")

    t = d.add_table(rows=1, cols=4)
    t.style = "Table Grid"
    for i, txt in enumerate(("文件", "必改", "建议", "待确认")):
        t.rows[0].cells[i].text = txt
    for doc, fs, rv in entries:
        cells = t.add_row().cells
        cells[0].text = doc.name
        cells[1].text = str(sum(1 for f in fs if f.level == "error"))
        cells[2].text = str(sum(1 for f in fs if f.level == "warn"))
        cells[3].text = str(len(rv))

    for doc, fs, rv in entries:
        if not fs and not rv:
            continue
        p = d.add_paragraph()
        r = p.add_run(doc.name)
        r.bold = True
        if fs:
            tt = d.add_table(rows=1, cols=5)
            tt.style = "Table Grid"
            for i, txt in enumerate(("等级", "位置", "角色", "检查项", "实际情况 → 应为")):
                tt.rows[0].cells[i].text = txt
            for f in fs:
                loc = f"第{f.index + 1}段" if f.index is not None else "整篇"
                cells = tt.add_row().cells
                cells[0].text = LEVEL_CN.get(f.level, f.level)
                cells[1].text = loc
                cells[2].text = ROLE_CN.get(f.role, f.role)
                cells[3].text = f.field
                cells[4].text = f"{f.actual} → {f.expected}"
        if rv:
            d.add_paragraph(f"待确认 {len(rv)} 段")
    d.save(str(path))
    return path


def save(path, doc: Doc, findings: list[Finding], review: list[Para],
         profile_name: str = "", entries=None) -> str:
    """按后缀导出报告：.txt/.md 纯文本，.docx Word，.pdf 先造 Word 再转。"""
    from pathlib import Path

    out = Path(path)
    suffix = out.suffix.lower()
    if entries is not None:
        if suffix == ".docx":
            return str(build_batch_docx(out, entries, profile_name))
        if suffix == ".pdf":
            tmp = out.with_suffix(".docx")
            build_batch_docx(tmp, entries, profile_name)
            _docx_to_pdf(tmp, out)
            tmp.unlink(missing_ok=True)
            return str(out)
        out.write_text(build_batch(entries, profile_name), encoding="utf-8-sig")
        return str(out)

    if suffix == ".docx":
        return str(build_docx(out, doc, findings, review, profile_name))
    if suffix == ".pdf":
        tmp = out.with_suffix(".docx")
        build_docx(tmp, doc, findings, review, profile_name)
        _docx_to_pdf(tmp, out)
        tmp.unlink(missing_ok=True)
        return str(out)
    out.write_text(build(doc, findings, review, profile_name), encoding="utf-8-sig")
    return str(out)


def _docx_to_pdf(src, dst):
    from . import convert as cv
    from . import office

    try:
        office.word_to_pdf(str(src), str(dst))
    except Exception:  # noqa: BLE001  —— 没有 WPS/Office 时用纯 Python 排字
        cv.to_pdf(str(src), str(dst.parent), cv.Options(pdf_engine="pure"))
        made = dst.parent / (src.stem + ".pdf")
        if made.exists() and made != dst:
            made.replace(dst)
