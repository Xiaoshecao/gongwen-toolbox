# -*- coding: utf-8 -*-
"""纯 Python 转换：不依赖 WPS/Office，装到哪台电脑都能跑。"""
from __future__ import annotations

import csv
import io
import os
from pathlib import Path

import pymupdf


# ---------------------------------------------------------------- PDF 基础
def pdf_page_count(path: str) -> int:
    with pymupdf.open(path) as doc:
        return doc.page_count


def pdf_to_images(path: str, out_dir: str, fmt: str = "png", dpi: int = 150) -> list[str]:
    """每页导出一张图片，返回生成的文件列表。"""
    fmt = fmt.lower().lstrip(".")
    if fmt == "jpeg":
        fmt = "jpg"
    made: list[str] = []
    stem = Path(path).stem
    with pymupdf.open(path) as doc:
        for i, page in enumerate(doc, start=1):
            pix = page.get_pixmap(dpi=dpi, alpha=False)
            dst = Path(out_dir) / f"{stem}_第{i}页.{fmt}"
            n = 1
            while dst.exists():
                dst = Path(out_dir) / f"{stem}_第{i}页({n}).{fmt}"
                n += 1
            pix.save(str(dst))
            made.append(str(dst))
    return made


def pdf_to_text(path: str) -> str:
    chunks: list[str] = []
    with pymupdf.open(path) as doc:
        for i, page in enumerate(doc, start=1):
            chunks.append(f"===== 第 {i} 页 =====\n{page.get_text('text').strip()}")
    return "\n\n".join(chunks).strip() + "\n"


def pdf_to_docx(path: str, out: str) -> None:
    """PDF 转 Word：优先用 pdf2docx 还原版式。"""
    from pdf2docx import Converter

    cv = Converter(path)
    try:
        cv.convert(out, start=0, end=None, multi_processing=False)
    finally:
        cv.close()


def images_to_pdf(paths: list[str], out: str) -> None:
    """多张图片合成一个 PDF，按文件名顺序，自动适配页面尺寸。"""
    doc = pymupdf.open()
    try:
        for p in paths:
            img = pymupdf.open(p)
            try:
                pdfbytes = img.convert_to_pdf()
            finally:
                img.close()
            with pymupdf.open("pdf", pdfbytes) as one:
                doc.insert_pdf(one)
        doc.save(out, deflate=True, garbage=3)
    finally:
        doc.close()


def merge_pdfs(paths: list[str], out: str) -> None:
    doc = pymupdf.open()
    try:
        for p in paths:
            with pymupdf.open(p) as one:
                doc.insert_pdf(one)
        doc.save(out, deflate=True, garbage=3)
    finally:
        doc.close()


def split_pdf(path: str, out_dir: str, every: int = 1) -> list[str]:
    """按每 N 页拆一个文件。"""
    made: list[str] = []
    stem = Path(path).stem
    with pymupdf.open(path) as doc:
        total = doc.page_count
        step = max(1, int(every))
        part = 1
        for start in range(0, total, step):
            end = min(start + step, total)
            dst = Path(out_dir) / f"{stem}_第{part}部分.pdf"
            n = 1
            while dst.exists():
                dst = Path(out_dir) / f"{stem}_第{part}部分({n}).pdf"
                n += 1
            with pymupdf.open() as one:
                one.insert_pdf(doc, from_page=start, to_page=end - 1)
                one.save(str(dst), deflate=True, garbage=3)
            made.append(str(dst))
            part += 1
    return made


def pdf_pages_to_one(path: str, pages: list[int], out: str) -> None:
    """抽取指定页码（从 1 开始）合并成一个 PDF。"""
    with pymupdf.open(path) as doc:
        with pymupdf.open() as one:
            for p in pages:
                if 1 <= p <= doc.page_count:
                    one.insert_pdf(doc, from_page=p - 1, to_page=p - 1)
            one.save(out, deflate=True, garbage=3)


# ---------------------------------------------------------------- Word
def docx_to_text(path: str) -> str:
    import docx

    d = docx.Document(path)
    parts: list[str] = []
    for para in d.paragraphs:
        if para.text.strip():
            parts.append(para.text)
    for table in d.tables:
        for row in table.rows:
            cells = [c.text.strip().replace("\n", " ") for c in row.cells]
            if any(cells):
                parts.append("\t".join(cells))
    return "\n".join(parts) + "\n"


# ---------------------------------------------------------------- Excel
def xlsx_to_csv(path: str, out: str) -> None:
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        ws = wb.worksheets[0]
        with open(out, "w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.writer(fh)
            for row in ws.iter_rows(values_only=True):
                if row is None:
                    continue
                if all(v is None for v in row):
                    continue
                writer.writerow(["" if v is None else v for v in row])
    finally:
        wb.close()


def xlsx_to_text(path: str) -> str:
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    out = io.StringIO()
    try:
        for ws in wb.worksheets:
            out.write(f"===== 工作表：{ws.title} =====\n")
            for row in ws.iter_rows(values_only=True):
                if row is None or all(v is None for v in row):
                    continue
                out.write("\t".join("" if v is None else str(v) for v in row) + "\n")
            out.write("\n")
    finally:
        wb.close()
    return out.getvalue()


# ---------------------------------------------------------------- PPT
def pptx_to_text(path: str) -> str:
    from pptx import Presentation

    prs = Presentation(path)
    out = io.StringIO()
    for i, slide in enumerate(prs.slides, start=1):
        out.write(f"===== 第 {i} 页 =====\n")
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                out.write(shape.text_frame.text.strip() + "\n")
            if getattr(shape, "has_table", False) and shape.has_table:
                for row in shape.table.rows:
                    out.write("\t".join(c.text.strip() for c in row.cells) + "\n")
        out.write("\n")
    return out.getvalue()


def pptx_to_images(path: str, out_dir: str) -> list[str]:
    """PPT 转图片：借 WPS 先转 PDF 再渲染，保证排版正确。"""
    from . import office_com

    tmp_pdf = str(Path(out_dir) / f".{Path(path).stem}_tmp.pdf")
    office_com.ppt_to(path, tmp_pdf, pdf=True)
    try:
        return pdf_to_images(tmp_pdf, out_dir, "png")
    finally:
        try:
            os.remove(tmp_pdf)
        except OSError:
            pass


def txt_guess_encoding(path: str) -> str:
    raw = Path(path).read_bytes()
    for enc in ("utf-8-sig", "utf-8", "gb18030", "utf-16"):
        try:
            raw.decode(enc)
            return enc
        except UnicodeDecodeError:
            continue
    return "gb18030"
