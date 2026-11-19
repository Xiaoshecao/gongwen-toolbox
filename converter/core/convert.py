# -*- coding: utf-8 -*-
"""转换调度：把「一个源文件 -> 一种目标格式」拆成可回退的策略链。"""
from __future__ import annotations

import os
import traceback
from dataclasses import dataclass, field
from pathlib import Path

from . import office_com, pdf_engine
from .formats import IMAGE_EXT, PDF_EXT, TARGETS, out_name, suffix


class ConvertError(RuntimeError):
    """转换失败，消息直接展示给用户。"""


@dataclass
class Options:
    dpi: int = 150
    split_every: int = 1
    merge: bool = False
    keep_source: bool = True
    extra: dict = field(default_factory=dict)


def _need(out: str) -> None:
    if not os.path.exists(out) or os.path.getsize(out) == 0:
        raise ConvertError("转换没有产生有效文件")


def _com_or_raise(src: str, out: str, what: str) -> None:
    try:
        office_com.to_pdf(src, out)
    except Exception as exc:
        raise ConvertError(
            f"{what}转 PDF 失败：{exc}。请确认本机已安装 WPS 或 Microsoft Office 且能正常打开该文件。"
        ) from exc
    _need(out)


# ------------------------------------------------------------------ 目标 PDF
def to_pdf(src: str, out: str, opt: Options) -> list[str]:
    s = suffix(src)
    if s in IMAGE_EXT:
        pdf_engine.images_to_pdf([src], out)
        _need(out)
        return [out]
    if s in PDF_EXT:
        raise ConvertError("源文件已经是 PDF")
    _com_or_raise(src, out, "Office 文档")
    return [out]


def images_to_one_pdf(paths: list[str], out: str, opt: Options) -> list[str]:
    pdf_engine.images_to_pdf(paths, out)
    _need(out)
    return [out]


def merge_to_pdf(paths: list[str], out: str, opt: Options) -> list[str]:
    """多个文件合并成一个 PDF：Office 先各自转 PDF，再拼接。"""
    import tempfile

    tmpdir = tempfile.mkdtemp(prefix="fmt_merge_")
    parts: list[str] = []
    try:
        for p in paths:
            s = suffix(p)
            if s in PDF_EXT:
                parts.append(p)
            elif s in IMAGE_EXT:
                tmp = os.path.join(tmpdir, f"{len(parts)}.pdf")
                pdf_engine.images_to_pdf([p], tmp)
                parts.append(tmp)
            else:
                tmp = os.path.join(tmpdir, f"{len(parts)}.pdf")
                office_com.to_pdf(p, tmp)
                if not os.path.exists(tmp):
                    raise ConvertError(f"{Path(p).name} 转 PDF 失败")
                parts.append(tmp)
        pdf_engine.merge_pdfs(parts, out)
        _need(out)
        return [out]
    finally:
        for f in parts:
            if f.startswith(tmpdir):
                try:
                    os.remove(f)
                except OSError:
                    pass
        try:
            os.rmdir(tmpdir)
        except OSError:
            pass


# ------------------------------------------------------------------ 目标 DOCX
def _pdf_to_docx_plain(src: str, out: str) -> None:
    """pdf2docx 失败时的兜底：按页抽文字，简单排版成 docx。"""
    import docx
    from docx.shared import Pt

    d = docx.Document()
    style = d.styles["Normal"]
    style.font.name = "宋体"
    style.font.size = Pt(10.5)
    text = pdf_engine.pdf_to_text(src)
    for block in text.split("\n\n"):
        block = block.strip()
        if not block:
            continue
        if block.startswith("===== 第"):
            d.add_page_break()
            continue
        for line in block.splitlines():
            d.add_paragraph(line)
    d.save(out)


def to_docx(src: str, out: str, opt: Options) -> list[str]:
    s = suffix(src)
    if s in PDF_EXT:
        try:
            pdf_engine.pdf_to_docx(src, out)
            _need(out)
        except Exception:
            _pdf_to_docx_plain(src, out)
            _need(out)
        return [out]
    # doc/rtf/wps -> docx 交给办公软件，保留版式
    try:
        office_com.upgrade(src, out)
    except Exception as exc:
        raise ConvertError(f"转换为 DOCX 失败：{exc}") from exc
    _need(out)
    return [out]


# ------------------------------------------------------------------ 目标 XLSX
def _pdf_tables_to_xlsx(src: str, out: str) -> int:
    import openpyxl

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    count = 0
    with pdf_engine.pymupdf.open(src) as doc:
        for pno, page in enumerate(doc, start=1):
            tables = []
            if hasattr(page, "find_tables"):
                try:
                    tables = page.find_tables().tables
                except Exception:
                    tables = []
            if not tables:
                continue
            for tno, table in enumerate(tables, start=1):
                rows = table.extract() or []
                if not rows:
                    continue
                ws = wb.create_sheet(f"第{pno}页_{tno}")
                for row in rows:
                    ws.append(["" if c is None else str(c) for c in row])
                count += 1
    if count == 0:
        raise ConvertError("这一页里没识别到表格")
    wb.save(out)
    return count


def _pdf_text_to_xlsx(src: str, out: str) -> None:
    """没有表格结构时，按行拆成两列兜底，至少能编辑。"""
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "文字"
    ws.append(["页码", "内容"])
    with pdf_engine.pymupdf.open(src) as doc:
        for pno, page in enumerate(doc, start=1):
            for line in page.get_text("text").splitlines():
                if line.strip():
                    ws.append([pno, line.strip()])
    wb.save(out)


def to_xlsx(src: str, out: str, opt: Options) -> list[str]:
    s = suffix(src)
    if s in PDF_EXT:
        try:
            _pdf_tables_to_xlsx(src, out)
        except ConvertError:
            _pdf_text_to_xlsx(src, out)
        _need(out)
        return [out]
    if s in {".xls", ".et"}:
        try:
            office_com.upgrade(src, out)
        except Exception as exc:
            raise ConvertError(f"转换为 XLSX 失败：{exc}") from exc
        _need(out)
        return [out]
    if s == ".csv":
        import openpyxl

        wb = openpyxl.Workbook()
        ws = wb.active
        enc = pdf_engine.txt_guess_encoding(src)
        with open(src, "r", encoding=enc, newline="") as fh:
            import csv as _csv

            for row in _csv.reader(fh):
                ws.append(row)
        wb.save(out)
        _need(out)
        return [out]
    raise ConvertError(f"不支持 {s} 转 XLSX")


# ------------------------------------------------------------------ 目标 PPTX
def to_pptx(src: str, out: str, opt: Options) -> list[str]:
    s = suffix(src)
    if s in PDF_EXT:
        from pptx import Presentation
        from pptx.util import Emu, Inches

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        blank = prs.slide_layouts[6]
        with pdf_engine.pymupdf.open(src) as doc:
            for page in doc:
                pix = page.get_pixmap(dpi=opt.dpi, alpha=False)
                img_path = str(Path(out).with_suffix(f".p{page.number + 1}.png"))
                pix.save(img_path)
                slide = prs.slides.add_slide(blank)
                # 按原始比例居中铺满，避免拉伸变形
                pw, ph = pix.width, pix.height
                box_w, box_h = prs.slide_width, prs.slide_height
                scale = min(box_w / pw, box_h / ph)
                w, h = int(pw * scale), int(ph * scale)
                left, top = int((box_w - w) / 2), int((box_h - h) / 2)
                slide.shapes.add_picture(img_path, Emu(left), Emu(top), Emu(w), Emu(h))
                try:
                    os.remove(img_path)
                except OSError:
                    pass
        prs.save(out)
        _need(out)
        return [out]
    try:
        office_com.upgrade(src, out)
    except Exception as exc:
        raise ConvertError(f"转换为 PPTX 失败：{exc}") from exc
    _need(out)
    return [out]


# ------------------------------------------------------------------ 目标 TXT / CSV
def to_txt(src: str, out: str, opt: Options) -> list[str]:
    s = suffix(src)
    if s in PDF_EXT:
        text = pdf_engine.pdf_to_text(src)
    elif s in {".docx", ".docm", ".dotx"}:
        text = pdf_engine.docx_to_text(src)
    elif s in {".xlsx", ".xlsm"}:
        text = pdf_engine.xlsx_to_text(src)
    elif s in {".pptx", ".ppsx"}:
        text = pdf_engine.pptx_to_text(src)
    elif s in {".txt", ".md", ".rtf"}:
        enc = pdf_engine.txt_guess_encoding(src)
        text = Path(src).read_text(encoding=enc, errors="replace")
    else:
        # 老格式先升到新格式再抽文字
        import tempfile

        tmpdir = tempfile.mkdtemp(prefix="fmt_txt_")
        try:
            if s in {".doc", ".wps", ".dot"}:
                tmp = os.path.join(tmpdir, "a.docx")
                office_com.upgrade(src, tmp)
                text = pdf_engine.docx_to_text(tmp)
            elif s in {".xls", ".et"}:
                tmp = os.path.join(tmpdir, "a.xlsx")
                office_com.upgrade(src, tmp)
                text = pdf_engine.xlsx_to_text(tmp)
            elif s in {".ppt", ".pps", ".dps"}:
                tmp = os.path.join(tmpdir, "a.pptx")
                office_com.upgrade(src, tmp)
                text = pdf_engine.pptx_to_text(tmp)
            else:
                raise ConvertError(f"不支持 {s} 转 TXT")
        finally:
            for name in os.listdir(tmpdir):
                try:
                    os.remove(os.path.join(tmpdir, name))
                except OSError:
                    pass
            try:
                os.rmdir(tmpdir)
            except OSError:
                pass
    Path(out).write_text(text, encoding="utf-8-sig")
    _need(out)
    return [out]


def to_csv(src: str, out: str, opt: Options) -> list[str]:
    s = suffix(src)
    if s in {".xlsx", ".xlsm"}:
        pdf_engine.xlsx_to_csv(src, out)
    elif s in {".xls", ".et"}:
        import tempfile

        tmpdir = tempfile.mkdtemp(prefix="fmt_csv_")
        try:
            tmp = os.path.join(tmpdir, "a.xlsx")
            office_com.upgrade(src, tmp)
            pdf_engine.xlsx_to_csv(tmp, out)
        finally:
            for name in os.listdir(tmpdir):
                try:
                    os.remove(os.path.join(tmpdir, name))
                except OSError:
                    pass
            try:
                os.rmdir(tmpdir)
            except OSError:
                pass
    else:
        raise ConvertError(f"不支持 {s} 转 CSV")
    _need(out)
    return [out]


# ------------------------------------------------------------------ 目标图片
def to_images(src: str, out_dir: str, fmt: str, opt: Options) -> list[str]:
    s = suffix(src)
    if s in PDF_EXT:
        made = pdf_engine.pdf_to_images(src, out_dir, fmt, opt.dpi)
    elif s in {".pptx", ".ppsx", ".ppt", ".pps", ".dps"}:
        if s in {".ppt", ".pps", ".dps"}:
            import tempfile

            tmpdir = tempfile.mkdtemp(prefix="fmt_img_")
            try:
                tmp_pdf = os.path.join(tmpdir, "a.pdf")
                office_com.to_pdf(src, tmp_pdf)
                made = pdf_engine.pdf_to_images(tmp_pdf, out_dir, fmt, opt.dpi)
            finally:
                for name in os.listdir(tmpdir):
                    try:
                        os.remove(os.path.join(tmpdir, name))
                    except OSError:
                        pass
                try:
                    os.rmdir(tmpdir)
                except OSError:
                    pass
        else:
            tmp_pdf = str(Path(out_dir) / f".{Path(src).stem}_tmp.pdf")
            office_com.ppt_to(src, tmp_pdf, pdf=True)
            try:
                made = pdf_engine.pdf_to_images(tmp_pdf, out_dir, fmt, opt.dpi)
            finally:
                try:
                    os.remove(tmp_pdf)
                except OSError:
                    pass
    else:
        raise ConvertError(f"不支持 {s} 转图片")
    if not made:
        raise ConvertError("没有导出任何页面")
    return made


# ------------------------------------------------------------------ 入口
def convert(src: str, target: str, out_dir: str, opt: Options | None = None) -> list[str]:
    """转换单个文件，返回生成的文件列表。失败抛 ConvertError。"""
    opt = opt or Options()
    if not os.path.exists(src):
        raise ConvertError("文件不存在或已被移动")
    os.makedirs(out_dir, exist_ok=True)
    dst = str(out_name(src, target, out_dir))

    try:
        if target == "PDF":
            return to_pdf(src, dst, opt)
        if target == "DOCX":
            return to_docx(src, dst, opt)
        if target == "XLSX":
            return to_xlsx(src, dst, opt)
        if target == "PPTX":
            return to_pptx(src, dst, opt)
        if target == "TXT":
            return to_txt(src, dst, opt)
        if target == "CSV":
            return to_csv(src, dst, opt)
        if target in ("PNG", "JPG"):
            return to_images(src, out_dir, target.lower(), opt)
        raise ConvertError(f"未知目标格式 {target}")
    except ConvertError:
        raise
    except Exception as exc:  # 把底层异常收敛成一句人话
        detail = traceback.format_exc(limit=2)
        raise ConvertError(f"转换失败：{exc}") from exc
