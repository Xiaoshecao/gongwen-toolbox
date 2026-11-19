# -*- coding: utf-8 -*-
"""纯 Python 的 PDF / 文档处理：不依赖 WPS 或 Office，装到哪台电脑都能跑。

从格式转换器搬过来的，只改了两处：
- `import pymupdf` 改成走 `pdfmod.load()`（Debian 上这个库只叫 `fitz`）。
- 不再往 `office_com` 里扎，需要转 PDF 时调 `office.word_to_pdf`。
"""
from __future__ import annotations

import csv
import io
import os
from pathlib import Path

from . import pdfmod


def _fitz():
    return pdfmod.load()


# ---------------------------------------------------------------- PDF 基础
def pdf_page_count(path: str) -> int:
    with _fitz().open(path) as doc:
        return doc.page_count


def pdf_to_images(path: str, out_dir: str, fmt: str = "png", dpi: int = 150) -> list[str]:
    """每页导出一张图片，返回生成的文件列表。"""
    fmt = fmt.lower().lstrip(".")
    if fmt == "jpeg":
        fmt = "jpg"
    made: list[str] = []
    stem = Path(path).stem
    with _fitz().open(path) as doc:
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
    for i, text in enumerate(pdf_pages_text(path), start=1):
        chunks.append(f"===== 第 {i} 页 =====\n{text}")
    return "\n\n".join(chunks).strip() + "\n"


def pdf_pages_text(path: str) -> list[str]:
    """逐页取纯文字，返回列表（页序）。

    比 `pdf_to_text` 更适合程序处理：那个返回的是带 `===== 第 N 页 =====`
    标记的整段字符串，调用方一旦按空行切块，标记会和正文粘在同一块里，
    想丢标记就会连正文一起丢掉（这个坑真踩过）。
    """
    out: list[str] = []
    with _fitz().open(path) as doc:
        for page in doc:
            out.append(page.get_text("text").strip())
    return out


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
    fitz = _fitz()
    doc = fitz.open()
    try:
        for p in paths:
            img = fitz.open(p)
            try:
                pdfbytes = img.convert_to_pdf()
            finally:
                img.close()
            with fitz.open("pdf", pdfbytes) as one:
                doc.insert_pdf(one)
        doc.save(out, deflate=True, garbage=3)
    finally:
        doc.close()


def merge_pdfs(paths: list[str], out: str) -> None:
    fitz = _fitz()
    doc = fitz.open()
    try:
        for p in paths:
            with fitz.open(p) as one:
                doc.insert_pdf(one)
        doc.save(out, deflate=True, garbage=3)
    finally:
        doc.close()


def split_pdf(path: str, out_dir: str, every: int = 1) -> list[str]:
    """按每 N 页拆一个文件。"""
    fitz = _fitz()
    made: list[str] = []
    stem = Path(path).stem
    with fitz.open(path) as doc:
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
            with fitz.open() as one:
                one.insert_pdf(doc, from_page=start, to_page=end - 1)
                one.save(str(dst), deflate=True, garbage=3)
            made.append(str(dst))
            part += 1
    return made


def reshape_pdf(path: str, out: str, orientation: str = "auto",
                margin_mm: float = 0.0, rotate_content: bool = False) -> None:
    """改纸张方向和留白：把每页内容按比例缩放后重新摆到新纸面上。

    orientation: auto 保持原样 / portrait 统一竖版 / landscape 统一横版
    margin_mm:   四周留白（毫米）。0 表示只换方向不缩放。
    rotate_content: 横竖方向变了时，把内容整体转 90°（宽表格转竖页时用），
                    否则内容会缩得很小但字是正的。

    内容用 `keep_proportion=True` 等比缩放居中，绝不裁切。
    """
    fitz = _fitz()
    m = max(0.0, float(margin_mm)) * 72.0 / 25.4
    src_doc = fitz.open(path)
    out_doc = fitz.open()
    try:
        for i in range(src_doc.page_count):
            page = src_doc[i]
            w, h = page.rect.width, page.rect.height
            if orientation == "portrait":
                tw, th = min(w, h), max(w, h)
            elif orientation == "landscape":
                tw, th = max(w, h), min(w, h)
            else:
                tw, th = w, h

            rot = 0
            if rotate_content and (w > h) != (tw > th):
                rot = 90

            newp = out_doc.new_page(width=tw, height=th)
            rect = fitz.Rect(m, m, tw - m, th - m)
            if rect.width < 20 or rect.height < 20:      # 留白比纸还大，退回满页
                rect = fitz.Rect(0, 0, tw, th)
            newp.show_pdf_page(rect, src_doc, i, rotate=rot, keep_proportion=True)
        out_doc.save(out, deflate=True, garbage=3)
    finally:
        out_doc.close()
        src_doc.close()


def pdf_pages_to_one(path: str, pages: list[int], out: str) -> None:
    """抽取指定页码（从 1 开始）合并成一个 PDF。"""
    fitz = _fitz()
    with fitz.open(path) as doc:
        with fitz.open() as one:
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
    """PPT 转图片：先转 PDF 再渲染，保证排版正确。"""
    from . import office

    tmp_pdf = str(Path(out_dir) / f".{Path(path).stem}_tmp.pdf")
    office.word_to_pdf(path, tmp_pdf)
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

# ---------------------------------------------------------------- 安全 / 体积 / 抽页
def encrypt(src, out, password: str, owner: str = "") -> str:
    """加打开密码（AES-256）。"""
    fitz = _fitz()
    doc = fitz.open(str(src))
    try:
        doc.save(str(out), encryption=fitz.PDF_ENCRYPT_AES_256,
                 owner_pw=owner or password, user_pw=password, garbage=3, deflate=True)
    finally:
        doc.close()
    return str(out)


def decrypt(src, out, password: str) -> str:
    """去掉密码（要有原密码）。"""
    fitz = _fitz()
    doc = fitz.open(str(src))
    try:
        if doc.needs_pass and not doc.authenticate(password):
            raise RuntimeError("密码不对，打不开这个 PDF。")
        doc.save(str(out), garbage=3, deflate=True)
    finally:
        doc.close()
    return str(out)


def compress(src, out) -> tuple[str, int, int]:
    """瘦身：清冗余对象 + 压缩流。返回 (out, 原字节, 新字节)。"""
    fitz = _fitz()
    before = Path(str(src)).stat().st_size
    doc = fitz.open(str(src))
    try:
        doc.save(str(out), garbage=4, deflate=True, clean=True)
    finally:
        doc.close()
    return str(out), before, Path(str(out)).stat().st_size


def parse_pages(spec: str, total: int) -> list[int]:
    """'1-3,5' → [1,2,3,5]（1 起，超范围自动丢掉）。"""
    pages: list[int] = []
    for chunk in str(spec).replace("，", ",").split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        if "-" in chunk:
            a, _, b = chunk.partition("-")
            try:
                lo, hi = int(a), int(b)
            except ValueError:
                continue
            pages.extend(range(min(lo, hi), max(lo, hi) + 1))
        else:
            try:
                pages.append(int(chunk))
            except ValueError:
                continue
    return sorted({p for p in pages if 1 <= p <= total})


def extract_pages(src, out, spec: str) -> tuple[str, int]:
    fitz = _fitz()
    doc = fitz.open(str(src))
    try:
        pages = parse_pages(spec, doc.page_count)
        if not pages:
            raise RuntimeError("页码没写对。比如想抽第 1 到 3 页和第 5 页，就写 1-3,5")
        new = fitz.open()
        for p in pages:
            new.insert_pdf(doc, from_page=p - 1, to_page=p - 1)
        new.save(str(out), garbage=3, deflate=True)
        new.close()
    finally:
        doc.close()
    return str(out), len(pages)
