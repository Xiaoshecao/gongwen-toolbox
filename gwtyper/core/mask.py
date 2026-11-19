# -*- coding: utf-8 -*-
"""敏感信息脱敏：身份证号、手机号、银行卡号打码，用于对外公示。

要点（v3.4 重写）：
- 手机号允许中间有空格或短横（138-1234-5678、138 1234 5678、+86 13812345678）
- 身份证按「地区码 + 8 位生日 + 3 位顺序 + 校验位」校验，15 位老号码也认
- Excel 里如果号码是「数字格式」（不是文本），同样打码
- Word 表格、页眉页脚也处理；号码被拆成几段文字（Word 常见）也能认出来
"""
from __future__ import annotations

import re
from pathlib import Path

# 18 位身份证：地区码 + 合法生日 + 顺序码 + 校验位（末位可为 X）
RE_ID18 = re.compile(
    r"(?<![0-9A-Za-z])([1-9]\d{5})"
    r"((?:18|19|20)\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01]))"
    r"(\d{3}[\dXx])(?![0-9A-Za-z])"
)
# 15 位老身份证：地区码 + 6 位生日 + 3 位顺序
RE_ID15 = re.compile(
    r"(?<!\d)([1-9]\d{5})(\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01]))(\d{3})(?!\d)"
)
# 手机号：可带 +86、空格、短横
RE_PHONE = re.compile(
    r"(?<!\d)(?:\+?86[-\s]?)?(1[3-9]\d)[-\s]?(\d{4})[-\s]?(\d{4})(?!\d)"
)
# 银行卡：16~19 位，可带空格/短横分组
RE_BANK = re.compile(
    r"(?<!\d)(\d{4})[-\s]?(\d{4})[-\s]?(\d{4})[-\s]?(\d{4,7})(?!\d)"
)

RE_ANY = re.compile(
    "|".join([RE_ID18.pattern, RE_ID15.pattern, RE_PHONE.pattern, RE_BANK.pattern])
)


def _stars(n: int) -> str:
    return "*" * n


def mask_text(s: str, strict_id: bool = True) -> str:
    """返回打码后的文本。strict_id=True 时身份证只留前 4 后 4。"""
    if not s:
        return s

    def _id18(m):
        head, last = m.group(1), m.group(3)
        if strict_id:
            return f"{head[:4]}{_stars(10)}{last}"
        return f"{head}{_stars(8)}{last}"

    def _id15(m):
        head, last = m.group(1), m.group(3)
        if strict_id:
            return f"{head[:4]}{_stars(7)}{last}"
        return f"{head}{_stars(6)}{last}"

    def _bank(m):
        digits = re.sub(r"\D", "", m.group(0))
        return f"{digits[:4]}{_stars(max(1, len(digits) - 8))}{digits[-4:]}"

    s = RE_ID18.sub(_id18, s)
    s = RE_ID15.sub(_id15, s)
    s = RE_PHONE.sub(lambda m: f"{m.group(1)}****{m.group(3)}", s)
    s = RE_BANK.sub(_bank, s)
    return s


def _hits_detail(s: str) -> dict:
    """先摘掉身份证再数手机/银行卡，避免把 18 位身份证重复算成银行卡。"""
    id_hits = len(RE_ID18.findall(s)) + len(RE_ID15.findall(s))
    rest = RE_ID15.sub(" ", RE_ID18.sub(" ", s))
    return {
        "id": id_hits,
        "phone": len(RE_PHONE.findall(rest)),
        "bank": len(RE_BANK.findall(rest)),
    }


def count_hits(s: str) -> int:
    d = _hits_detail(s)
    return d["id"] + d["phone"] + d["bank"]


def _cell_text(v):
    """Excel 单元格取文本：数字不带 .0，日期等非字符串返回 None。"""
    if isinstance(v, bool):
        return None
    if isinstance(v, str):
        return v
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else str(v)
    return None


def _mask_paragraph(p, strict_id: bool) -> int:
    """把一段（含多 run）合并判断，命中就把打码文本写进第一个 run。"""
    runs = [r for r in p.runs]
    full = "".join(r.text for r in runs)
    if not full.strip():
        return 0
    new = mask_text(full, strict_id)
    if new == full:
        return 0
    n = count_hits(full)
    if runs:
        runs[0].text = new
        for r in runs[1:]:
            r.text = ""
    return n


def _iter_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
    for sec in doc.sections:
        for part in (sec.header, sec.footer, sec.first_page_header,
                     sec.first_page_footer, sec.even_page_header,
                     sec.even_page_footer):
            try:
                for p in part.paragraphs:
                    yield p
                for t in part.tables:
                    for row in t.rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                yield p
            except Exception:  # noqa: BLE001
                continue


def mask_docx(src, out, strict_id: bool = True) -> dict:
    from docx import Document

    doc = Document(str(src))
    hits, detail = 0, {"id": 0, "phone": 0, "bank": 0}
    for p in _iter_paragraphs(doc):
        full = "".join(r.text for r in p.runs)
        n = _mask_paragraph(p, strict_id)
        if n:
            hits += n
            d = _hits_detail(full)
            for k in detail:
                detail[k] += d[k]
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    return {"out": str(out), "hits": hits, **detail}


def mask_xlsx(src, out, strict_id: bool = True) -> dict:
    from openpyxl import load_workbook

    wb = load_workbook(str(src))
    hits, detail = 0, {"id": 0, "phone": 0, "bank": 0}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                text = _cell_text(cell.value)
                if not text:
                    continue
                new = mask_text(text, strict_id)
                if new == text:
                    continue
                n = count_hits(text)
                if n:
                    hits += n
                    d = _hits_detail(text)
                    for k in detail:
                        detail[k] += d[k]
                    cell.value = new
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    return {"out": str(out), "hits": hits, **detail}


def mask_pdf(src, out, strict_id: bool = True) -> dict:
    """只处理有文字层的 PDF；扫描件请先用 OCR 生成文字层。"""
    from . import pdfmod

    fitz = pdfmod.load()
    doc = fitz.open(str(src))
    hits, detail = 0, {"id": 0, "phone": 0, "bank": 0}
    try:
        for page in doc:
            text = page.get_text()
            seen = set()
            for m in RE_ANY.finditer(text):
                raw = m.group(0)
                key = (raw, m.start())
                if key in seen:
                    continue
                seen.add(key)
                masked = mask_text(raw, strict_id)
                if masked == raw:
                    continue
                rects = page.search_for(raw)
                if not rects:
                    # 号码可能被断成两行/带空格，退一步只搜数字串
                    digits = re.sub(r"[^\dXx]", "", raw)
                    rects = page.search_for(digits) if digits else []
                for rect in rects:
                    page.add_redact_annot(
                        rect, text=masked,
                        fontsize=max(6, min(11, rect.height * 0.9)),
                        fill=(1, 1, 1))
                    hits += 1
                d = _hits_detail(raw)
                for k in detail:
                    detail[k] += d[k]
            try:
                page.apply_redactions()
            except Exception:  # noqa: BLE001
                pass
        out = Path(out)
        out.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(out), garbage=3, deflate=True)
    finally:
        doc.close()
    return {"out": str(out), "hits": hits, **detail}


def run(src, out, strict_id: bool = True) -> dict:
    suf = Path(src).suffix.lower()
    if suf == ".docx":
        return mask_docx(src, out, strict_id)
    if suf in (".xlsx", ".xlsm"):
        return mask_xlsx(src, out, strict_id)
    if suf == ".pdf":
        return mask_pdf(src, out, strict_id)
    raise ValueError("只支持 .docx / .xlsx / .pdf。老的 .doc 请先另存为 .docx。")
