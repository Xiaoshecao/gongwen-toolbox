# -*- coding: utf-8 -*-
"""格式转换：一个源文件 → 一种目标格式，按策略链逐步回退。

从原来的「格式转换器」搬过来，改动：
- Office 通道统一走 `office.to_pdf` / `office.upgrade`（Windows COM / Linux WPS /
  LibreOffice 三选一），不再直接 win32com，所以 Linux 上也能转 PDF。
  `office` 按扩展名分 word / excel / ppt 三类分别派发，Excel 用 KET、
  PPT 用 KWPP，不会再出现「拿 WPS 文字去开 xlsx」这种错。
- pymupdf 统一走 `pdfmod.load()`。
- 失败信息里不再写死「请确认已安装 WPS 或 Office」，改成报当前引擎。
"""
from __future__ import annotations

import os
import tempfile
import traceback
from dataclasses import dataclass, field
from pathlib import Path

from . import office, pdf_engine, pdfmod
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
    # 转 PDF 时的版面选项（见 pdf_engine.reshape_pdf）
    orientation: str = "auto"        # auto / portrait / landscape
    margin_mm: float = 0.0           # 四周留白，毫米
    rotate_content: bool = False     # 横竖互换时把内容转 90°
    # 用哪个 PDF 引擎：pure=纯 Python（不依赖办公软件）/ external=WPS·Office·LibreOffice
    # auto = 先试纯 Python，不行再退外部引擎
    pdf_engine: str = "auto"

    def wants_reshape(self) -> bool:
        return self.orientation != "auto" or self.margin_mm > 0


def _need(out: str) -> None:
    if not os.path.exists(out) or os.path.getsize(out) == 0:
        raise ConvertError("转换没有产生有效文件")


def _com_or_raise(src: str, out: str, what: str) -> None:
    try:
        office.to_pdf(src, out)
    except Exception as exc:  # noqa: BLE001
        raise ConvertError(
            f"{what}转 PDF 失败：{exc}\n"
            f"当前转换引擎：{office.engine_hint()}。装一个 WPS 或 LibreOffice 就能转。"
        ) from exc
    _need(out)


def _tmpdir(prefix: str) -> str:
    return tempfile.mkdtemp(prefix=prefix)


def _cleanup(dirpath: str) -> None:
    try:
        for name in os.listdir(dirpath):
            try:
                os.remove(os.path.join(dirpath, name))
            except OSError:
                pass
        os.rmdir(dirpath)
    except OSError:
        pass


# ------------------------------------------------------------------ 目标 PDF
def _reshape_if_needed(raw: str, out: str, opt: Options) -> str:
    """需要改方向/留白时，把 raw 重排成 out；不需要就原样返回 raw。"""
    if not opt.wants_reshape():
        return raw
    pdf_engine.reshape_pdf(raw, out, opt.orientation, opt.margin_mm,
                           opt.rotate_content)
    _need(out)
    return out


def _pure_pdf(src: str, out: str, opt: Options) -> bool:
    """用纯 Python 引擎排 PDF。支持的类型返回 True，不支持的返回 False。

    这条路完全不碰 WPS / Office / LibreOffice，Linux 上尤其重要：
    LibreOffice 常常只装了 writer，转 xlsx/pptx 会报
    `source file could not be loaded`，看着像文件坏了。
    """
    s = suffix(src)
    try:
        from . import pdfwrite
    except ImportError:
        return False
    if not pdfwrite.available():
        return False

    if s in {".docx", ".docm", ".dotx"}:
        page = None
        if opt.orientation != "auto" or opt.margin_mm > 0:
            # 纸张方向交给 reshape 统一处理，这里只按原尺寸排
            page = None
        pdfwrite.docx_to_pdf(src, out, page)
        return True
    if s in {".xlsx", ".xlsm"}:
        pdfwrite.xlsx_to_pdf(src, out)
        return True
    if s in {".txt", ".md"}:
        pdfwrite.text_to_pdf(src, out)
        return True
    if s in {".csv", ".et"}:
        # csv 先并成 xlsx 再排，避免再写一套
        import openpyxl

        tmpdir = _tmpdir("gwtyper_csvpdf_")
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            enc = pdf_engine.txt_guess_encoding(src)
            import csv as _csv

            with open(src, "r", encoding=enc, newline="") as fh:
                for row in _csv.reader(fh):
                    ws.append(row)
            tmp = os.path.join(tmpdir, "a.xlsx")
            wb.save(tmp)
            pdfwrite.xlsx_to_pdf(tmp, out)
            return True
        finally:
            _cleanup(tmpdir)
    return False


def to_pdf(src: str, out: str, opt: Options) -> list[str]:
    s = suffix(src)
    if s in IMAGE_EXT:
        if not opt.wants_reshape():
            pdf_engine.images_to_pdf([src], out)
            _need(out)
            return [out]
        tmpdir = _tmpdir("gwtyper_pdf_")
        try:
            raw = os.path.join(tmpdir, "a.pdf")
            pdf_engine.images_to_pdf([src], raw)
            return [_reshape_if_needed(raw, out, opt)]
        finally:
            _cleanup(tmpdir)
    if s in PDF_EXT:
        if not opt.wants_reshape():
            raise ConvertError("源文件已经是 PDF")
        return [_reshape_if_needed(src, out, opt)]

    want_pure = opt.pdf_engine == "pure"
    want_auto = opt.pdf_engine == "auto"

    # ---- auto：**先外部引擎**（Windows 的 WPS/Office COM、Linux 的 LibreOffice），
    #      转不了再用纯 Python。外部引擎出来的东西跟办公软件里看到的完全一致，
    #      优先它更符合直觉；纯 Python 只在没有办公引擎或它失败时兜底。----
    if want_auto:
        try:
            if not opt.wants_reshape():
                _com_or_raise(src, out, "Office 文档")
                return [out]
            tmpdir = _tmpdir("gwtyper_pdf_")
            try:
                raw = os.path.join(tmpdir, "a.pdf")
                _com_or_raise(src, raw, "Office 文档")
                return [_reshape_if_needed(raw, out, opt)]
            finally:
                _cleanup(tmpdir)
        except ConvertError:
            pass          # 外部引擎不行，落到下面的纯 Python

    # ---- 纯 Python（pure 模式，或 auto 的外部引擎失败后）----
    if want_pure or want_auto:
        tmpdir = _tmpdir("gwtyper_purepdf_")
        try:
            raw = os.path.join(tmpdir, "a.pdf") if opt.wants_reshape() else out
            try:
                ok = _pure_pdf(src, raw, opt)
            except Exception as exc:  # noqa: BLE001
                if want_pure:
                    raise ConvertError(f"纯 Python 排版失败：{exc}") from exc
                ok = False
            if ok:
                _need(raw)
                if opt.wants_reshape():
                    return [_reshape_if_needed(raw, out, opt)]
                return [out]
        finally:
            _cleanup(tmpdir)
        if want_pure:
            raise ConvertError(
                "纯 Python 引擎转不了这个格式（支持 docx / xlsx / csv / txt / 图片 / PDF）。\n"
                "演示文稿（pptx/ppt）请把「PDF 引擎」改成「外部引擎」。")

    # ---- external：只走外部引擎 ----
    if not opt.wants_reshape():
        _com_or_raise(src, out, "Office 文档")
        return [out]

    tmpdir = _tmpdir("gwtyper_pdf_")
    try:
        raw = os.path.join(tmpdir, "a.pdf")
        _com_or_raise(src, raw, "Office 文档")
        return [_reshape_if_needed(raw, out, opt)]
    finally:
        _cleanup(tmpdir)


def images_to_one_pdf(paths: list[str], out: str, opt: Options) -> list[str]:
    if not opt.wants_reshape():
        pdf_engine.images_to_pdf(paths, out)
        _need(out)
        return [out]
    tmpdir = _tmpdir("gwtyper_ipdf_")
    try:
        raw = os.path.join(tmpdir, "a.pdf")
        pdf_engine.images_to_pdf(paths, raw)
        return [_reshape_if_needed(raw, out, opt)]
    finally:
        _cleanup(tmpdir)


def _merge_raw(paths: list[str], out: str) -> None:
    """只做拼接，不做版面调整。"""
    tmpdir = _tmpdir("gwtyper_merge_")
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
                office.to_pdf(p, tmp)
                if not os.path.exists(tmp):
                    raise ConvertError(f"{Path(p).name} 转 PDF 失败")
                parts.append(tmp)
        pdf_engine.merge_pdfs(parts, out)
        _need(out)
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


def merge_to_pdf(paths: list[str], out: str, opt: Options) -> list[str]:
    """多个文件合并成一个 PDF：Office 先各自转 PDF，再拼接。"""
    if not opt.wants_reshape():
        _merge_raw(paths, out)
        return [out]
    tmpdir = _tmpdir("gwtyper_merge_r_")
    try:
        raw = os.path.join(tmpdir, "a.pdf")
        _merge_raw(paths, raw)
        return [_reshape_if_needed(raw, out, opt)]
    finally:
        _cleanup(tmpdir)


def merge_many(paths: list[str], out: str, bookmarks: bool = True,
               progress=None) -> list[str]:
    """把一摞材料（Word/PDF/图片混着也行）按顺序合成一个 PDF。

    bookmarks=True 时给每个源文件加一个书签（标题就是文件名），
    长材料翻起来能直接跳。
    """
    if not paths:
        raise ConvertError("没有要合并的文件。")
    tmpdir = _tmpdir("gwtyper_merge_many_")
    parts: list[tuple[str, str]] = []   # (pdf 路径, 显示名)
    try:
        for i, p in enumerate(paths, start=1):
            if progress:
                progress(f"第 {i}/{len(paths)} 份：{Path(p).name}")
            s = suffix(p)
            if s in PDF_EXT:
                parts.append((p, Path(p).stem))
                continue
            tmp = os.path.join(tmpdir, f"{i}.pdf")
            if s in IMAGE_EXT:
                pdf_engine.images_to_pdf([p], tmp)
            else:
                office.to_pdf(p, tmp)
            if not os.path.exists(tmp) or os.path.getsize(tmp) == 0:
                raise ConvertError(f"{Path(p).name} 转 PDF 失败")
            parts.append((tmp, Path(p).stem))
        pdf_engine.merge_pdfs([p for p, _ in parts], out)
        _need(out)
        if bookmarks:
            _add_merge_bookmarks(out, parts)
    finally:
        _cleanup(tmpdir)
    return [out]


def _add_merge_bookmarks(out: str, parts: list[tuple[str, str]]) -> None:
    """给合并后的 PDF 按源文件加一级书签。页码算不出来就跳过。"""
    try:
        from . import pdfmod

        fitz = pdfmod.load()
        starts, cur = [], 1
        for p, name in parts:
            starts.append((name, cur))
            with fitz.open(str(p)) as d:
                cur += d.page_count
        doc = fitz.open(str(out))
        try:
            toc = [[1, name, page] for name, page in starts]
            doc.set_toc(toc)
            doc.saveIncr()
        finally:
            doc.close()
    except Exception:  # noqa: BLE001
        log.debug("加书签失败（不影响合并结果）", exc_info=True)


# ------------------------------------------------------------------ 目标 DOCX
def _pdf_to_docx_plain(src: str, out: str) -> None:
    """pdf2docx 不可用时的兜底：逐页抽文字，简单排版成 docx。

    注意别用 `pdf_to_text()` 的返回值去 split("\\n\\n")：那里的每个块都是
    「===== 第 N 页 ===== + 正文」粘在一起，按标记过滤会把正文一起扔掉，
    结果得到一个只有分页符的空文档（而且不报错）。这里改用逐页接口。
    """
    import docx
    from docx.shared import Pt

    d = docx.Document()
    style = d.styles["Normal"]
    style.font.name = "宋体"
    style.font.size = Pt(10.5)

    pages = pdf_engine.pdf_pages_text(src)
    total = 0
    for i, page_text in enumerate(pages):
        if i:
            d.add_page_break()
        for line in page_text.splitlines():
            if line.strip():
                d.add_paragraph(line.rstrip())
                total += len(line.strip())
    if total == 0:
        raise ConvertError(
            "这份 PDF 里取不到文字（多半是扫描件，整页都是图片）。\n"
            "可以改用「转成图片」，或者先做 OCR 再转 Word。")
    d.save(out)


def to_docx(src: str, out: str, opt: Options) -> list[str]:
    s = suffix(src)
    if s in PDF_EXT:
        # 优先 pdf2docx（能还原表格和大致版式）；它没打包或失败时，
        # 退到「逐页抽文字」——这条只保证文字不丢，版式要自己再排一遍。
        # 注意 pdf2docx 依赖 numpy/opencv，打包时被排除了，所以实际
        # 走的多半是兜底那条，要提前告诉用户。
        try:
            pdf_engine.pdf_to_docx(src, out)
            _need(out)
        except Exception:  # noqa: BLE001
            _pdf_to_docx_plain(src, out)
            _need(out)
            opt.extra["pdf_to_docx_plain"] = True
        return [out]
    # doc/rtf/wps -> docx 交给办公软件，保留版式
    try:
        office.upgrade(src, out)
    except Exception as exc:  # noqa: BLE001
        raise ConvertError(f"转换为 DOCX 失败：{exc}") from exc
    _need(out)
    return [out]


# ------------------------------------------------------------------ 目标 XLSX
def _pdf_tables_to_xlsx(src: str, out: str) -> int:
    import openpyxl

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    count = 0
    with pdfmod.load().open(src) as doc:
        for pno, page in enumerate(doc, start=1):
            tables = []
            if hasattr(page, "find_tables"):
                try:
                    tables = page.find_tables().tables
                except Exception:  # noqa: BLE001
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
    with pdfmod.load().open(src) as doc:
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
            office.upgrade(src, out)
        except Exception as exc:  # noqa: BLE001
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
        with pdfmod.load().open(src) as doc:
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
        office.upgrade(src, out)
    except Exception as exc:  # noqa: BLE001
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
        tmpdir = _tmpdir("gwtyper_txt_")
        try:
            if s in {".doc", ".wps", ".dot"}:
                tmp = os.path.join(tmpdir, "a.docx")
                office.upgrade(src, tmp)
                text = pdf_engine.docx_to_text(tmp)
            elif s in {".xls", ".et"}:
                tmp = os.path.join(tmpdir, "a.xlsx")
                office.upgrade(src, tmp)
                text = pdf_engine.xlsx_to_text(tmp)
            elif s in {".ppt", ".pps", ".dps"}:
                tmp = os.path.join(tmpdir, "a.pptx")
                office.upgrade(src, tmp)
                text = pdf_engine.pptx_to_text(tmp)
            else:
                raise ConvertError(f"不支持 {s} 转 TXT")
        finally:
            _cleanup(tmpdir)
    Path(out).write_text(text, encoding="utf-8-sig")
    _need(out)
    return [out]


def to_csv(src: str, out: str, opt: Options) -> list[str]:
    s = suffix(src)
    if s in {".xlsx", ".xlsm"}:
        pdf_engine.xlsx_to_csv(src, out)
    elif s in {".xls", ".et"}:
        tmpdir = _tmpdir("gwtyper_csv_")
        try:
            tmp = os.path.join(tmpdir, "a.xlsx")
            office.upgrade(src, tmp)
            pdf_engine.xlsx_to_csv(tmp, out)
        finally:
            _cleanup(tmpdir)
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
            tmpdir = _tmpdir("gwtyper_img_")
            try:
                tmp_pdf = os.path.join(tmpdir, "a.pdf")
                office.to_pdf(src, tmp_pdf)
                made = pdf_engine.pdf_to_images(tmp_pdf, out_dir, fmt, opt.dpi)
            finally:
                _cleanup(tmpdir)
        else:
            tmp_pdf = str(Path(out_dir) / f".{Path(src).stem}_tmp.pdf")
            office.to_pdf(src, tmp_pdf)
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


# ------------------------------------------------------------------ 改版面
def reshape(src: str, out: str, opt: Options) -> list[str]:
    """把已有 PDF（或 Office/图片）改成指定方向和留白，另存一份。"""
    s = suffix(src)
    if s in PDF_EXT:
        if opt.orientation == "auto" and opt.margin_mm <= 0:
            raise ConvertError("方向和留白都没改，不用转。")
        pdf_engine.reshape_pdf(src, out, opt.orientation, opt.margin_mm,
                              opt.rotate_content)
        _need(out)
        return [out]

    tmpdir = _tmpdir("gwtyper_resh_")
    try:
        raw = os.path.join(tmpdir, "a.pdf")
        if s in IMAGE_EXT:
            pdf_engine.images_to_pdf([src], raw)
        else:
            _com_or_raise(src, raw, "Office 文档")
        pdf_engine.reshape_pdf(raw, out, opt.orientation or "auto",
                              opt.margin_mm, opt.rotate_content)
        _need(out)
        return [out]
    finally:
        _cleanup(tmpdir)


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
    except Exception as exc:  # noqa: BLE001
        # 把底层异常收敛成一句人话，详细堆栈留在 extra 里
        raise ConvertError(f"转换失败：{exc}\n{traceback.format_exc(limit=2)}") from exc
