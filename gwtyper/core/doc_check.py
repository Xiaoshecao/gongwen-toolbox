# -*- coding: utf-8 -*-
"""公文要素完整性检查（GB/T 9704-2012 的结构要素）。

只管「该有的有没有、顺序对不对」，不管字体字号（那是 check.py 的事）。
只读，产出 Finding。
"""
from __future__ import annotations

import re

from .model import Doc, Finding, Para, Role

RE_SEND_TO = re.compile(r"[：:]\s*$")
RE_ATTACH = re.compile(r"^附件\s*[:：\d]")
RE_CC = re.compile(r"^抄送\s*[:：]")
RE_PRINT = re.compile(r"印发|印发机关")
RE_DATE_ARABIC = re.compile(r"^\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日$")
RE_DATE_CN = re.compile(r"^[〇零○一二三四五六七八九十]{4}\s*年")
RE_DATE_ANY = re.compile(r"(19|20)?\d{2}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日")
RE_TITLE_PUNCT = re.compile(r"[。．.]$")

# 公文中「不写主送机关」的文种（命令、公告、通告、决定等面向社会公布的）
NO_SEND_TO_TYPES = ("公告", "通告", "命令", "决定", "决议", "公报")


def _text(p: Para) -> str:
    return (p.text or "").strip()


def _first_index(paras: list[Para], pred) -> int | None:
    for p in paras:
        if pred(p):
            return p.index
    return None


def _body_start(paras: list[Para], after: int) -> int | None:
    for p in paras:
        if p.index <= after:
            continue
        if p.role in (Role.BODY, Role.LEAD) and len(_text(p)) >= 10:
            return p.index
    return None


def check_elements(doc: Doc) -> list[Finding]:
    """检查标题、主送机关、正文、附件说明、署名、成文日期、抄送、印发的完整性与顺序。"""
    out: list[Finding] = []
    paras = [p for p in doc.paras if _text(p)]
    if not paras:
        out.append(Finding(None, "other", "内容", "空文档", "至少要有标题和正文",
                           "error", "整篇没有文字"))
        return out

    def add(index, field, actual, expected, level="error", note=""):
        out.append(Finding(index, "other", field, actual, expected, level, note))

    title = next((p for p in paras if p.role == Role.TITLE), None)
    if title is None:
        add(None, "标题", "未识别到标题", "标题（二号方正小标宋，居中）", "error",
            "第一段是标题，字号明显大于正文、居中")

    # —— 标题不加句号
    if title is not None and RE_TITLE_PUNCT.search(_text(title)):
        add(title.index, "标题标点", _text(title)[-12:], "标题末尾不加句号", "warn")

    # —— 主送机关：标题之后、正文之前，顶格、以冒号结尾
    after = title.index if title is not None else -1
    send_to = None
    body_start = _body_start(paras, after)
    for p in paras:
        if p.index <= after:
            continue
        if body_start is not None and p.index >= body_start:
            break
        t = _text(p)
        if 2 <= len(t) <= 60 and RE_SEND_TO.search(t) and p.role not in (
                Role.TITLE, Role.SUBTITLE, Role.SIGNATURE, Role.DATE):
            send_to = p
            break

    has_sig = any(p.role == Role.SIGNATURE for p in paras)
    has_date = any(p.role == Role.DATE for p in paras) or bool(
        _first_index(paras, lambda p: RE_DATE_ANY.match(_text(p))))
    looks_formal = bool(title) and (has_sig or has_date)
    ttl = _text(title) if title is not None else ""
    exempt = any(k in ttl for k in NO_SEND_TO_TYPES)

    if send_to is None and looks_formal and not exempt:
        add(None, "主送机关", "未发现主送机关",
            "主送机关（顶格、机关名称后加冒号）", "warn",
            "只有公告/通告/命令/决定这类面向社会公布的文种才不写主送机关")

    # —— 发文机关署名与成文日期
    if has_sig and not has_date:
        add(None, "成文日期", "有署名、无日期", "成文日期（阿拉伯数字，署名下一行右空四字）", "error")
    if has_date and not has_sig:
        add(None, "发文机关署名", "有日期、无署名", "发文机关署名（与标题中的机关一致）", "error")
    if looks_formal and not has_sig and not has_date:
        add(None, "落款", "既无署名也无成文日期", "发文机关署名 + 成文日期", "error")

    date_par = next((p for p in paras if p.role == Role.DATE), None)
    if date_par is None:
        idx = _first_index(paras, lambda p: RE_DATE_ANY.match(_text(p)))
        date_par = next((p for p in paras if p.index == idx), None) if idx is not None else None
    if date_par is not None:
        t = _text(date_par)
        if RE_DATE_CN.match(t):
            add(date_par.index, "成文日期写法", t, "用阿拉伯数字，如 2026年11月18日",
                "warn", "GB/T 9704 规定成文日期用阿拉伯数字")
        elif not RE_DATE_ARABIC.match(t):
            add(date_par.index, "成文日期格式", t, "YYYY年M月D日（月、日不补零）", "warn")

    # —— 顺序：署名在日期之前
    sig = next((p for p in paras if p.role == Role.SIGNATURE), None)
    if sig is not None and date_par is not None and date_par.index < sig.index:
        add(date_par.index, "落款顺序", "日期在署名之前", "先发文机关署名，后成文日期", "error")

    # —— 附件说明：应在正文之后、署名之前
    att = next((p for p in paras if RE_ATTACH.match(_text(p))), None)
    if att is not None:
        if sig is not None and att.index > sig.index:
            add(att.index, "附件说明位置", "出现在署名之后", "附件说明应在正文之后、署名之前", "error")
        elif body_start is not None and att.index < body_start:
            add(att.index, "附件说明位置", "出现在正文之前", "附件说明应在正文之后、署名之前", "warn")

    # —— 抄送机关 / 印发机关（版记，必须在最后）
    cc = next((p for p in paras if RE_CC.match(_text(p))), None)
    if cc is not None and body_start is not None and cc.index < body_start:
        add(cc.index, "抄送机关位置", "出现在正文之前", "抄送机关在版记（文末）", "error")

    pr = next((p for p in paras if RE_PRINT.search(_text(p))), None)
    if pr is not None and body_start is not None and pr.index < body_start:
        add(pr.index, "印发机关位置", "出现在正文之前", "印发机关和印发日期在版记（文末）", "error")

    # —— 正文
    if not any(p.role in (Role.BODY, Role.LEAD) for p in paras):
        add(None, "正文", "未识别到正文段落", "正文（三号仿宋）", "error")

    return out
