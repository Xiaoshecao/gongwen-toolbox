# -*- coding: utf-8 -*-
"""文本层体检：逐段比对实测格式与规则，产出问题清单。只读。"""
from __future__ import annotations

from .model import Doc, Finding, Role

CHECK_FIELDS = ["font", "size_pt", "line_pt", "align", "first_indent_chars"]
FIELD_CN = {
    "font": "字体", "size_pt": "字号", "line_pt": "行距", "align": "对齐",
    "first_indent_chars": "首行缩进", "space_before_pt": "段前", "space_after_pt": "段后",
    "bold": "加粗", "right_indent_chars": "右缩进",
}
ALIGN_CN = {"left": "左对齐", "center": "居中", "right": "右对齐", "justify": "两端对齐",
            "distribute": "分散对齐", None: "未设"}

NO_INDENT = {Role.TITLE, Role.SUBTITLE, Role.SIGNATURE, Role.DATE,
             Role.TABLE_CAPTION, Role.NOTE}
# 段内允许多字体（引导语楷体 + 正文仿宋）的角色
MULTI_FONT_OK = {Role.LEAD}

FONT_ALIAS = {
    "仿宋_GB2312": {"仿宋", "FangSong", "FangSong_GB2312", "fangsong"},
    "楷体_GB2312": {"楷体", "KaiTi", "KaiTi_GB2312", "kaiti"},
    "方正小标宋简体": {"宋体", "SimSun", "华文中宋", "STZhongsong", "方正小标宋_GBK"},
    "黑体": {"SimHei", "黑体_GB2312"},
}


def _size_cn(v):
    """字号显示：优先中文字号（「三号（16磅）」），对不上号数才只写磅值。

    基层文书只知道「三号」「小四」，看到「16 磅」还得换算，
    所以体检结果里把号数放在前面。
    """
    if not isinstance(v, (int, float)):
        return str(v)
    try:
        from . import fontsizes

        return fontsizes.label(float(v))
    except Exception:  # noqa: BLE001
        return f"{v:g}磅"


def _font_level(actual: str | None, want: str) -> str:
    """字体不一致：本机替换字体算 warn，别的一律 error。"""
    return "warn" if (actual or "") in FONT_ALIAS.get(want, set()) else "error"


def check(doc: Doc, rules: dict, tol: float = 0.6, text: bool = True) -> list[Finding]:
    """text=True 时一并做文本层校对（标点/数字/错词/序号）。"""
    out: list[Finding] = []
    roles = rules.get("roles", {})

    for p in doc.paras:
        if p.role == Role.BLANK:
            continue
        spec = roles.get(p.role.value)
        if not spec:
            continue
        f = p.fmt
        dom = p.dominant

        # 字体：以主体 run 为准；引导语段落主体是正文，楷体只看段首 run
        if spec.get("font"):
            if p.role == Role.LEAD:
                first = next((r for r in p.runs if r.text.strip()), None)
                if first and not (first.font or "").startswith("楷体"):
                    out.append(Finding(p.index, p.role.value, "引导语字体",
                                       first.font or "未设", spec["font"], level="warn"))
                body_want = (roles.get("body") or {}).get("font")
                got = (dom.font if dom else None) or f.font
                if body_want and got != body_want:
                    out.append(Finding(p.index, p.role.value, "正文字体", got or "未设",
                                       body_want, level=_font_level(got, body_want)))
            else:
                got = (dom.font if dom else None) or f.font
                if got != spec["font"]:
                    out.append(Finding(p.index, p.role.value, "字体", got or "未设",
                                       spec["font"], level=_font_level(got, spec["font"])))

        if spec.get("size_pt") and dom and dom.size_pt:
            if abs(dom.size_pt - spec["size_pt"]) > 0.1:
                out.append(Finding(p.index, p.role.value, "字号", _size_cn(dom.size_pt),
                                   _size_cn(spec["size_pt"])))

        if spec.get("line_pt") and f.line_pt is not None:
            if f.line_rule != "exactly":
                out.append(Finding(p.index, p.role.value, "行距规则", f.line_rule or "未设",
                                   "固定值", level="warn"))
            elif abs(f.line_pt - spec["line_pt"]) > tol:
                out.append(Finding(p.index, p.role.value, "行距", f"{f.line_pt:g}磅",
                                   f"{spec['line_pt']:g}磅固定值"))

        if spec.get("align") and f.align != spec["align"]:
            out.append(Finding(p.index, p.role.value, "对齐",
                               ALIGN_CN.get(f.align, str(f.align)),
                               ALIGN_CN.get(spec["align"], spec["align"])))

        # 首行缩进
        if p.role not in NO_INDENT and "first_line_chars" in spec:
            want = spec["first_line_chars"]
            size = (dom.size_pt if dom and dom.size_pt else None) or spec.get("size_pt") or 16
            if f.first_indent_chars is not None:
                if abs(f.first_indent_chars - want) > 0.05:
                    out.append(Finding(p.index, p.role.value, "首行缩进",
                                       f"{f.first_indent_chars:g}字符", f"{want:g}字符"))
            elif f.first_indent_pt is not None:
                got = round(f.first_indent_pt / size, 2)
                if abs(got - want) > 0.15:
                    out.append(Finding(p.index, p.role.value, "首行缩进",
                                       f"{got:g}字符({f.first_indent_pt:g}磅)", f"{want:g}字符"))
            elif want > 0:
                out.append(Finding(p.index, p.role.value, "首行缩进", "无", f"{want:g}字符"))

        for k in ("space_before_pt", "space_after_pt"):
            if spec.get(k) is not None:
                got = getattr(f, k) or 0
                if abs(got - spec[k]) > 1.0:
                    out.append(Finding(p.index, p.role.value, FIELD_CN[k],
                                       f"{got:g}磅", f"{spec[k]:g}磅", level="warn"))

        if spec.get("bold") is True and dom and dom.bold is not True:
            out.append(Finding(p.index, p.role.value, "加粗", "否", "是", level="warn"))

        if p.mixed and p.role not in MULTI_FONT_OK:
            kinds = "、".join(sorted({(r.font or "?") for r in p.runs if r.text.strip()}))
            out.append(Finding(p.index, p.role.value, "段内格式", f"混合：{kinds}",
                               "统一为一种字体字号", level="warn",
                               note="常见于复制粘贴"))

    # 页面
    pg = rules.get("page", {})
    ps = doc.page
    for k, cn in (("top_cm", "上边距"), ("bottom_cm", "下边距"),
                  ("left_cm", "左边距"), ("right_cm", "右边距")):
        if k in pg:
            got = getattr(ps, k)
            if abs(got - pg[k]) > 0.05:
                out.append(Finding(None, "page", cn, f"{got:g}cm", f"{pg[k]:g}cm"))

    if text:
        from .doc_check import check_elements
        from .text_check import check_text

        out.extend(check_text(doc))
        out.extend(check_elements(doc))
    return out


def check_features(doc: Doc, rules: dict, pages: int = 0) -> list[Finding]:
    """整篇级要素：页码域、篇末截止线、页数上限。需要读原始 docx，故单独成函数。"""
    out: list[Finding] = []
    ft = rules.get("features") or {}
    if not ft:
        return out
    try:
        from docx import Document

        d = Document(doc.path)
        has_page = any("PAGE" in (s.footer._element.xml if s.footer else "")
                       for s in d.sections)
    except Exception:  # noqa: BLE001
        return out

    if ft.get("page_number") and not has_page:
        out.append(Finding(None, "page", "页码", "无页码域", "页脚居中「— 1 —」", "warn",
                           note="用域才能自动编号"))
    if ft.get("page_number") is False and has_page:
        out.append(Finding(None, "page", "页码", "有页码", "本类文种不加页码", "warn"))
    if ft.get("max_pages") and pages > ft["max_pages"]:
        out.append(Finding(None, "page", "页数", f"{pages} 页",
                           f"控制在一页内（现超出 {pages - ft['max_pages']} 页）", "warn"))
    return out


def summary(findings: list[Finding]) -> dict:
    return {
        "total": len(findings),
        "error": sum(1 for x in findings if x.level == "error"),
        "warn": sum(1 for x in findings if x.level == "warn"),
    }
