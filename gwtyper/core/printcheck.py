# -*- coding: utf-8 -*-
"""打印体检：排版完先看看印出来会不会难看。

只看几件最容易出丑的事：
- 末页只剩一两行（翻过来就一行字）
- 有大片空白页
- 疑似表格/段落被页边切断
"""
from __future__ import annotations

import tempfile
from pathlib import Path


def _fitz():
    from . import pdfmod

    return pdfmod.load()


def _to_pdf(src: Path, tmp: Path) -> Path:
    from . import office

    dst = tmp / (src.stem + ".pdf")
    if src.suffix.lower() == ".pdf":
        return src
    office.word_to_pdf(src, dst)
    return dst


def check(path) -> list[dict]:
    """返回 [{"level": "warn"/"建议", "page": n or 0, "msg": "..."}]，page=0 表示整篇。"""
    fitz = _fitz()
    src = Path(path)
    findings: list[dict] = []
    with tempfile.TemporaryDirectory(prefix="gwtyper_pc_") as td:
        pdf = _to_pdf(src, Path(td))
        doc = fitz.open(str(pdf))
        try:
            n = doc.page_count
            if n == 0:
                return [{"level": "error", "page": 0, "msg": "这份文件没有页面。"}]
            last = doc[n - 1]
            r = last.rect
            blocks = [b for b in last.get_text("blocks") if str(b[4]).strip()]
            used = 0.0
            if blocks:
                top = min(b[1] for b in blocks)
                bottom = max(b[3] for b in blocks)
                used = max(0.0, bottom - top)
            ratio = used / max(1.0, r.height)
            if n > 1 and ratio < 0.12:
                findings.append({"level": "warn", "page": n,
                                 "msg": f"末页只有一点内容（约占页面 {ratio * 100:.0f}%），印出来翻过去只有一两行，建议调一下让上一页收住。"})
            for i, page in enumerate(doc, 1):
                txt = page.get_text().strip()
                if not txt and not page.get_images():
                    findings.append({"level": "warn", "page": i, "msg": "这一页是空白页。"})
                    continue
                bs = [b for b in page.get_text("blocks") if str(b[4]).strip()]
                if not bs:
                    continue
                bottom = max(b[3] for b in bs)
                if bottom > page.rect.height - 24 and i < n:
                    nxt = doc[i]
                    nb = [b for b in nxt.get_text("blocks") if str(b[4]).strip()]
                    if nb and min(b[1] for b in nb) < 90:
                        findings.append({"level": "建议", "page": i,
                                         "msg": "这一页底部贴着页边、下一页又从最上面接着写，像是有表格或段落被切断了，核对一下。"})
        finally:
            doc.close()
    return findings
