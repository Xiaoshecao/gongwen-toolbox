# -*- coding: utf-8 -*-
"""版面体检：转 PDF 后看真实落版。越界、表格超宽、图片出界、末页孤行、空白页。

PDF 通道走 `office.py`（Windows 优先 WPS/Office COM，其余平台用 LibreOffice），
拿不到 PDF 时返回 skipped，不抛异常。
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

from . import office, pdfmod

CM = 72 / 2.54          # pt per cm
MARGIN_TOL_PT = 6.0     # 允许 6 磅溢出（标点悬挂等正常现象）
EDGE_TOL_PT = 2.0


@dataclass
class PageIssue:
    page: int
    kind: str
    detail: str
    level: str = "error"

    def as_row(self) -> str:
        return f"[{self.level}] 第{self.page}页 {self.kind}：{self.detail}"


@dataclass
class LayoutReport:
    ok: bool
    pages: int = 0
    issues: list[PageIssue] = field(default_factory=list)
    skipped: str = ""
    pdf: str = ""

    def as_rows(self) -> list[str]:
        return [i.as_row() for i in self.issues]


def _to_pdf_engine() -> str:
    return office.engine_name()


def to_pdf(docx_path: str | Path, out_dir: str | Path | None = None) -> str:
    src = Path(docx_path)
    out_dir = Path(out_dir or src.parent / ".pdfcache")
    out_dir.mkdir(parents=True, exist_ok=True)
    dst = out_dir / f"{src.stem}.pdf"
    while dst.exists() and dst.stat().st_mtime < src.stat().st_mtime:
        dst.unlink()
    if not office.engine_name():
        raise RuntimeError("没有可用的转换引擎（装 WPS 或 LibreOffice）")
    office.word_to_pdf(src, dst)
    if not dst.exists():
        raise RuntimeError("转换未产出 PDF")
    return str(dst)


def check_pdf(pdf_path: str, page_spec: dict, n_tables: int = 0,
              min_last_page_ratio: float = 0.25) -> LayoutReport:
    """按页边距框判越界；再看末页是否只剩一两行。"""
    try:
        pymupdf = pdfmod.load()
    except pdfmod.PdfMissing:
        return LayoutReport(ok=True, skipped="未安装 PyMuPDF")

    top = page_spec.get("top_cm", 3.7) * CM
    bottom = page_spec.get("bottom_cm", 3.5) * CM
    left = page_spec.get("left_cm", 2.8) * CM
    right = page_spec.get("right_cm", 2.6) * CM

    issues: list[PageIssue] = []
    with pymupdf.open(pdf_path) as doc:
        n = doc.page_count
        for pno, page in enumerate(doc, start=1):
            w, h = page.rect.width, page.rect.height
            box = pymupdf.Rect(left, top, w - right, h - bottom)
            body_h = box.height

            # 正文区内的文字块：查越界
            blocks = [b for b in page.get_text("blocks") if b[4].strip()]
            for b in blocks:
                r = pymupdf.Rect(b[:4])
                if r.x1 > box.x1 + MARGIN_TOL_PT:
                    over = (r.x1 - box.x1) / CM
                    issues.append(PageIssue(pno, "文字超出右边距", f"超出 {over:.2f}cm", "error"))
                if r.x0 < box.x0 - MARGIN_TOL_PT:
                    over = (box.x0 - r.x0) / CM
                    issues.append(PageIssue(pno, "文字超出左边距", f"超出 {over:.2f}cm", "error"))
                if r.y1 > box.y1 + MARGIN_TOL_PT * 3:
                    over = (r.y1 - box.y1) / CM
                    issues.append(PageIssue(pno, "文字压到页脚区", f"超出 {over:.2f}cm", "warn"))

            # 图片/线条出界
            for img in page.get_images(full=True):
                for r in page.get_image_rects(img[0]):
                    if r.x0 < box.x0 - EDGE_TOL_PT or r.x1 > box.x1 + EDGE_TOL_PT \
                            or r.y0 < box.y0 - EDGE_TOL_PT * 3 or r.y1 > box.y1 + EDGE_TOL_PT * 6:
                        issues.append(PageIssue(pno, "图片超出正文区",
                                                f"图框 {r.x0 / CM:.1f},{r.y0 / CM:.1f} → "
                                                f"{r.x1 / CM:.1f},{r.y1 / CM:.1f} cm", "error"))

            # 表格线超宽
            for dr in page.get_drawings():
                rr = dr["rect"]
                if rr.width > box.width + MARGIN_TOL_PT * 2:
                    issues.append(PageIssue(pno, "表格/图形超宽",
                                            f"宽 {rr.width / CM:.2f}cm > 版心 {box.width / CM:.2f}cm",
                                            "error"))
                    break

            # 末页孤行
            if pno == n and n > 1:
                txt_blocks = [b for b in blocks]
                if txt_blocks:
                    y_max = max(b[3] for b in txt_blocks)
                    used = (y_max - box.y0) / body_h if body_h else 1
                    if 0 < used < min_last_page_ratio and n_tables == 0:
                        issues.append(PageIssue(pno, "末页孤行",
                                                f"只用到版心 {used * 100:.0f}%，建议调整行距或续页",
                                                "warn"))

    # 同类问题页面合并，避免刷屏
    merged: dict[tuple[str, str], list[int]] = {}
    for i in issues:
        merged.setdefault((i.kind, i.detail.split("超出")[0]), []).append(i.page)
    out: list[PageIssue] = []
    seen = set()
    for i in issues:
        key = (i.kind, i.level)
        if key in seen:
            continue
        same = [x for x in issues if x.kind == i.kind and x.level == i.level]
        pages = sorted({x.page for x in same})
        seen.add(key)
        first = same[0]
        if len(pages) == 1:
            out.append(first)
        else:
            out.append(PageIssue(pages[0], f"{i.kind}（{len(pages)} 页："
                                  + "、".join(map(str, pages[:8])) + ("…" if len(pages) > 8 else "")
                                  + "）", first.detail, i.level))
    return LayoutReport(ok=not out, pages=n, issues=out, pdf=pdf_path)


def check(docx_path: str | Path, page_spec: dict, n_tables: int = 0,
          keep_pdf: bool = False) -> LayoutReport:
    try:
        pdf = to_pdf(docx_path)
    except Exception as e:  # noqa: BLE001
        return LayoutReport(ok=True, skipped=f"转 PDF 失败（{type(e).__name__}: {e}）—— 版面未检查")
    try:
        return check_pdf(pdf, page_spec, n_tables)
    finally:
        if not keep_pdf:
            try:
                os.unlink(pdf)
            except OSError:
                pass
