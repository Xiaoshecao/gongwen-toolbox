# -*- coding: utf-8 -*-
"""OCR 纯文字 / 原版式输出：从图片或扫描件里把字认出来，落成 docx 或 txt。

三种产出：
  - `to_text(...)`       纯文字（txt），一段一行，适合检索、喂给别的程序
  - `to_docx(...)`       按原版式排回 Word：保留每行文字的位置、字号、粗体倾向，
                         尽量让结果看起来还像原来那一页
  - `make_double_layer()`（在 ocr.py 里）图像层 + 不可见文字层

版式还原的做法和限制：
  拿 OCR 的边界框，按行聚成段落（行距接近的合并），
  字号按框高换算，缩进按左边界和页宽比例换算，居中的行判成居中。
  **做不到**：还原图片、表格线、精确的字间距、字体区分（OCR 只给文字和框）。
  所以文档里会说明「这是按位置近似还原的，需要人工再核一遍」。
"""
from __future__ import annotations

import io
import logging
import shutil
import tempfile
from pathlib import Path

from . import ocr as ocrmod
from . import pdfmod

log = logging.getLogger(__name__)


class OcrTextError(RuntimeError):
    """识别失败，消息直接给用户看。"""


def _pages_of(src: str | Path, dpi: int = 200, progress=None,
              tmpdir: Path | None = None) -> list[dict]:
    """把源文件变成「逐页的 OCR 结果」。

    返回 [{"width": 页宽pt, "height": 页高pt, "items": [(x0,y0,x1,y1,text),...]}, ...]

    **临时目录由调用方创建和清理**（tmpdir 参数）：本函数只往里写中间图，
    不删任何东西。以前在这里 finally 清目录，结果把用户给的原始扫描件
    也一起删了。
    """
    src = Path(src)
    s = src.suffix.lower()
    if not src.is_file():
        raise OcrTextError(f"找不到文件：{src}")
    if not ocrmod.ocr_available():
        raise OcrTextError(
            "没装 OCR 引擎，做不了识别。\n"
            "装一下：pip install rapidocr_onnxruntime")

    fitz = pdfmod.load()
    pages: list[dict] = []
    if tmpdir is None:
        import tempfile

        tmpdir = Path(tempfile.mkdtemp(prefix="gwoct_"))
    tmpdir.mkdir(parents=True, exist_ok=True)

    if s in ocrmod.IMAGE_SOURCES:
        if progress:
            progress(f"{src.name}：识别中…")
        w_px, h_px = ocrmod._image_px(src)
        items = ocrmod._ocr_items(str(src), 72)
        pages.append({"width": float(w_px or 595),
                      "height": float(h_px or 842), "items": items})
        return pages

    # PDF / Office：先统一成 PDF
    if s in ocrmod.TEXT_SOURCES:
        from . import office

        mid = tmpdir / "mid.pdf"
        if progress:
            progress("转 PDF…")
        office.to_pdf(src, mid)
        pdf_path = mid
    elif s == ".pdf":
        pdf_path = src
    else:
        raise OcrTextError(f"这个格式不支持：{s}")

    doc = fitz.open(str(pdf_path))
    try:
        n = doc.page_count
        for i in range(n):
            page = doc[i]
            if progress:
                progress(f"第 {i+1}/{n} 页：识别中…")
            pix = page.get_pixmap(dpi=dpi)
            png = tmpdir / f"p{i}.png"
            pix.save(str(png))
            items = ocrmod._ocr_items(str(png), dpi)
            pages.append({"width": page.rect.width,
                          "height": page.rect.height, "items": items})
    finally:
        doc.close()
    return pages


def _group_lines(items: list[tuple]) -> list[list[tuple]]:
    """把 OCR 结果按行分组（同一水平带上的算一行）。

    OCR 给的粒度是「一行」（不是逐字），但同一段里可能有多个框，
    这里按 y 中心排序后再合并 y 相近的。
    """
    if not items:
        return []
    rows: list[list[tuple]] = []
    for it in sorted(items, key=lambda x: ((x[1] + x[3]) / 2, x[0])):
        cy = (it[1] + it[3]) / 2
        h = it[3] - it[1]
        placed = False
        for row in rows:
            ry0 = min(r[1] for r in row)
            ry1 = max(r[3] for r in row)
            if ry0 - h * 0.3 <= cy <= ry1 + h * 0.3:
                row.append(it)
                placed = True
                break
        if not placed:
            rows.append([it])
    # 行内按 x 排序
    for row in rows:
        row.sort(key=lambda x: x[0])
    return rows


def _group_paragraphs(rows: list[list[tuple]]) -> list[list[list[tuple]]]:
    """把行聚成段落。

    公文正文是两端对齐的，所以**段末那行一定短**（右边留白），
    这是比行距更可靠的断段信号。四种判据按可靠度排：

    1. 上一行右侧明显没顶到版心右边 → 上一行是段末，从这里断开。
    2. 当前行左边界比上一行右移了约两个字 → 首行缩进，新段开始。
    3. 相邻行间距 / 上一行行高 > 1.15 → 段间距（正文约 0.76~0.97，
       段间约 1.4）。不能用全页平均行高当基准，正文行会把均值拉高，
       阈值就永远不触发（踩过）。
    4. 字号突变（>1.35 或 <0.75）→ 标题/脚注另起。
    """
    if not rows:
        return []

    def _top(r):
        return min(x[1] for x in r)

    def _bot(r):
        return max(x[3] for x in r)

    def _h(r):
        return max(1.0, _bot(r) - _top(r))

    def _left(r):
        return min(x[0] for x in r)

    def _right(r):
        return max(x[2] for x in r)

    # 版心右边线：整页最靠右的那个点。留白太多（比如个别行溢出）时退到 95 分位。
    rights = sorted(_right(r) for r in rows)
    right_edge = rights[int(len(rights) * 0.95) - 1] if len(rights) >= 3 else rights[-1]

    paras: list[list[list[tuple]]] = [[rows[0]]]
    for prev, cur in zip(rows, rows[1:]):
        gap = _top(cur) - _bot(prev)
        ratio = gap / _h(prev)
        h_prev = _h(prev)
        new_para = False

        # 1) 上一行没顶到右边线 → 上一行是段末（公文两端对齐，最可靠）
        #    容忍一个字宽，避免标点占位差异误判
        if _right(prev) < right_edge - h_prev * 1.1:
            new_para = True
        # 2) 首行缩进：左边界右移约两字（行高的 0.85~2.0 倍之间）
        elif h_prev * 0.85 < (_left(cur) - _left(prev)) < h_prev * 2.0:
            new_para = True
        # 3) 段间距
        elif ratio > 1.15:
            new_para = True
        # 4) 字号突变
        elif _h(cur) > h_prev * 1.35 or _h(cur) < h_prev * 0.75:
            new_para = True

        if new_para:
            paras.append([cur])
        else:
            paras[-1].append(cur)
    return paras


def recognize(src: str | Path, dpi: int = 200, progress=None) -> dict:
    """只识别、不落盘。给「识别完先给我看、可以改字」用。

    返回 {"text": 全文, "pages": 页数, "chars": 字数}。
    多页时页与页之间留一行「===== 第 N 页 =====」做分隔（保存时可以去掉）。
    """
    import tempfile

    tmpdir = Path(tempfile.mkdtemp(prefix="gwoct_"))
    try:
        pages = _pages_of(src, dpi, progress, tmpdir=tmpdir)
        chunks: list[str] = []
        total = 0
        for i, pg in enumerate(pages, start=1):
            if len(pages) > 1:
                chunks.append(f"===== 第 {i} 页 =====")
            for para in _group_paragraphs(_group_lines(pg["items"])):
                line = "".join(it[4] for row in para for it in row).strip()
                if line:
                    chunks.append(line)
                    total += len(line)
                chunks.append("")
        return {"text": "\n".join(chunks).strip(), "pages": len(pages),
                "chars": total}
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def text_to_docx(text: str, out: str | Path, font: str = "仿宋_GB2312",
                 size_pt: float = 16, drop_page_marks: bool = True) -> dict:
    """把（人工改过的）纯文字写成 Word，正文样式统一。

    每行一段；`drop_page_marks=True` 时把「===== 第 N 页 =====」这种分隔行去掉。
    """
    import docx
    from docx.oxml.ns import qn
    from docx.shared import Pt

    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    d = docx.Document()
    st = d.styles["Normal"]
    st.font.name = font
    st.font.size = Pt(size_pt)
    try:
        st.element.rPr.rFonts.set(qn("w:eastAsia"), font)
    except Exception:  # noqa: BLE001
        log.debug("设置中文字体失败", exc_info=True)
    chars = 0
    for line in str(text).replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        s = line.rstrip()
        if drop_page_marks and s.startswith("===== 第") and s.endswith("页 ====="):
            continue
        d.add_paragraph(s)
        if s.strip():
            chars += len(s.strip())
    d.save(str(out))
    return {"out": str(out), "chars": chars}


def recognize_blocks(src: str | Path, dpi: int = 200, progress=None) -> dict:
    """识别并保留段落框：给「先看/改字，最后再生成双层 PDF」用。

    返回 {"pages": [{"width","height","paras":[{"rect":(x0,y0,x1,y1),"text":...}]}],
          "text": 按页拼接的全文（页间有「===== 第 N 页 =====」）, "chars": N}
    """
    import tempfile

    tmpdir = Path(tempfile.mkdtemp(prefix="gwoct_"))
    try:
        pages = _pages_of(src, dpi, progress, tmpdir=tmpdir)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    out_pages: list[dict] = []
    chunks: list[str] = []
    total = 0
    for i, pg in enumerate(pages, start=1):
        if len(pages) > 1:
            chunks.append(f"===== 第 {i} 页 =====")
        paras = []
        for para in _group_paragraphs(_group_lines(pg["items"])):
            text = "".join(it[4] for row in para for it in row).strip()
            if not text:
                continue
            xs0 = [it[0] for row in para for it in row]
            ys0 = [it[1] for row in para for it in row]
            xs1 = [it[2] for row in para for it in row]
            ys1 = [it[3] for row in para for it in row]
            paras.append({"rect": (min(xs0), min(ys0), max(xs1), max(ys1)),
                          "text": text})
            chunks.append(text)
            total += len(text)
        chunks.append("")
        out_pages.append({"width": pg["width"], "height": pg["height"],
                          "paras": paras})
    return {"pages": out_pages, "text": "\n".join(chunks).strip(), "chars": total}


def split_edited(text: str) -> list[list[str]]:
    """把编辑框里的全文拆成「每页一组段落」（空行分段）。"""
    pages: list[list[str]] = [[]]
    for raw in str(text).replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        s = raw.strip()
        if s.startswith("===== 第") and s.endswith("页 ====="):
            pages.append([])
            continue
        if s:
            pages[-1].append(s)
    return pages


def to_pdf_double_from_blocks(src: str | Path, out: str | Path, blocks: dict,
                              edited_pages: list[list[str]],
                              dpi: int = 200, progress=None) -> dict:
    """双层 PDF：底层原图，上层用（人工改过的）文字，按原段落框放。

    段落数和原来不一样时，多出来的文字放到该页最后一段的框里（可能挤，但只影响
    搜索高亮，不影响显示）。
    """
    fitz = pdfmod.load()
    src, out = Path(src), Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)

    tmpdir = Path(tempfile.mkdtemp(prefix="gw2l_"))
    try:
        src_doc = _open_as_pdf(src, tmpdir)
        try:
            dst = fitz.open()
            n = min(src_doc.page_count, max(1, len(blocks.get("pages") or [])))
            lines = 0
            for i in range(n):
                page = src_doc[i]
                w, h = page.rect.width, page.rect.height
                pix = page.get_pixmap(dpi=dpi)
                png = tmpdir / f"p{i}.png"
                pix.save(str(png))
                np_ = dst.new_page(width=w, height=h)
                np_.insert_image(np_.rect, filename=str(png))
                if progress:
                    progress(f"第 {i + 1}/{n} 页：铺图 + 写字…")
                texts = edited_pages[i] if i < len(edited_pages) else []
                boxes = (blocks["pages"][i].get("paras") or []) if i < len(blocks["pages"]) else []
                lines += _put_text(np_, texts, boxes, w, h)
            dst.save(str(out), garbage=3, deflate=True)
            dst.close()
        finally:
            src_doc.close()
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
    return {"out": str(out), "pages": n, "lines": lines}


def _open_as_pdf(src: Path, tmpdir: Path):
    """图片/PDF/Office 统一开成 fitz 文档（Office 先转 PDF）。"""
    fitz = pdfmod.load()
    s = src.suffix.lower()
    if s == ".pdf":
        return fitz.open(str(src))
    if s in ocrmod.IMAGE_SOURCES:
        img = fitz.open(str(src))
        try:
            pdf_bytes = img.convert_to_pdf()
        finally:
            img.close()
        d = fitz.open("pdf", pdf_bytes)
        return d
    from . import office

    mid = tmpdir / "mid.pdf"
    office.word_to_pdf(str(src), str(mid))
    return fitz.open(str(mid))


def _put_text(page, texts: list[str], boxes: list[dict], w: float, h: float) -> int:
    """把每段文字以「不可见文字层」放进对应框；返回写入段数。"""
    fitz = pdfmod.load()
    if not texts:
        return 0
    placed = 0
    for k, txt in enumerate(texts):
        if k < len(boxes):
            x0, y0, x1, y1 = boxes[k]["rect"]
        elif boxes:
            x0, y0, x1, y1 = boxes[-1]["rect"]
            y0, y1 = y1, min(h - 24, y1 + max(14, (y1 - y0)))
        else:
            x0, y0, x1, y1 = 36, 36 + k * 18, w - 36, 36 + k * 18 + 16
        rect = fitz.Rect(max(1, x0), max(1, y0), min(w - 1, x1 + 4), min(h - 1, y1 + 3))
        if rect.width < 12 or rect.height < 6:
            rect = fitz.Rect(36, 36 + k * 16, w - 36, 36 + k * 16 + 14)
        fs = max(5.0, min(28.0, rect.height * 0.86))
        done = False
        for scale in (1.0, 0.8, 0.62):
            try:
                rc = page.insert_textbox(rect, txt, fontsize=fs * scale,
                                          fontname="china-s", render_mode=3,
                                          color=(0, 0, 0), align=0)
            except Exception:  # noqa: BLE001
                log.debug("insert_textbox 失败，改 insert_text", exc_info=True)
                rc = -1
            if rc is not None and rc >= 0:
                done = True
                break
        if not done:
            try:
                page.insert_text((rect.x0 + 1, rect.y0 + fs), txt, fontsize=fs,
                                 fontname="china-s", render_mode=3)
                done = True
            except Exception:  # noqa: BLE001
                log.warning("文字层写入失败：%s", txt[:20])
        if done:
            placed += 1
    return placed


def to_text(src: str | Path, out: str | Path, dpi: int = 200,
            progress=None) -> dict:
    """OCR → 纯文字 txt。一段一行，保留空行分隔段落。"""
    import tempfile

    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    tmpdir = Path(tempfile.mkdtemp(prefix="gwoct_"))
    pages = _pages_of(src, dpi, progress, tmpdir=tmpdir)
    try:
        chunks: list[str] = []
        total_chars = 0
        for i, pg in enumerate(pages, start=1):
            if len(pages) > 1:
                chunks.append(f"===== 第 {i} 页 =====")
            rows = _group_lines(pg["items"])
            for para in _group_paragraphs(rows):
                line = "".join(t for row in para for t in
                               [it[4] for it in row])
                if line.strip():
                    chunks.append(line)
                    total_chars += len(line)
                chunks.append("")
        text = "\n".join(chunks).rstrip() + "\n"
        out.write_text(text, encoding="utf-8-sig")
        if total_chars == 0:
            raise OcrTextError("一个字都没认出来，可能图片太模糊或是空白页")
        return {"pages": len(pages), "chars": total_chars, "out": str(out)}
    finally:
        ocrmod._cleanup(tmpdir)


def to_docx(src: str | Path, out: str | Path, dpi: int = 200,
            page_size: tuple[float, float] | None = None,
            progress=None) -> dict:
    """OCR → 按原版式排回 Word（近似）。

    尽量保留：段落划分、字号、居中/右对齐、首行缩进。
    做不到：图片、表格线、字体区分、精确字距。
    """
    import docx
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt, Cm

    import tempfile

    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    tmpdir = Path(tempfile.mkdtemp(prefix="gwoct_"))
    pages = _pages_of(src, dpi, progress, tmpdir=tmpdir)
    try:
        d = docx.Document()
        # 页面按源尺寸设（pt → cm）
        sec = d.sections[0]
        if pages:
            w_pt = pages[0]["width"]
            h_pt = pages[0]["height"]
            sec.page_width = Cm(w_pt * 2.54 / 72)
            sec.page_height = Cm(h_pt * 2.54 / 72)
        sec.top_margin = Cm(2.8)
        sec.bottom_margin = Cm(2.8)
        sec.left_margin = Cm(2.8)
        sec.right_margin = Cm(2.6)

        total_chars = 0
        total_paras = 0
        for pi, pg in enumerate(pages):
            if pi:
                d.add_page_break()
            rows = _group_lines(pg["items"])
            paras = _group_paragraphs(rows)
            page_w = pg["width"]
            # 版心左右边界（按比例估）
            for para in paras:
                xs0 = min(it[0] for row in para for it in row)
                xs1 = max(it[2] for row in para for it in row)
                hs = [it[3] - it[1] for row in para for it in row]
                text = "".join(it[4] for row in para for it in row)
                if not text.strip():
                    continue
                p = d.add_paragraph()
                run = p.add_run(text)
                # 字号：按框高换算（框高 ≈ 字高，稍放大一点更接近原样）
                size_pt = max(9.0, min(30.0, (sum(hs) / len(hs)) * 1.15))
                run.font.size = Pt(round(size_pt * 2) / 2)
                # 对齐：两端留白接近 → 居中；右侧明显短于左侧 → 居左
                left_gap = xs0
                right_gap = page_w - xs1
                if abs(left_gap - right_gap) < page_w * 0.06:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif right_gap < page_w * 0.08 and left_gap > page_w * 0.25:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    # 首行缩进：左边界比「常见左边界」多出约两个字符 → 首行缩进
                    if left_gap > page_w * 0.10:
                        p.paragraph_format.first_line_indent = Pt(size_pt * 2)
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                total_chars += len(text)
                total_paras += 1

        # 末尾加一行说明，提醒这是近似还原
        if total_paras:
            note = d.add_paragraph()
            r = note.add_run("（本文件由图片识别生成，版式按位置近似还原，请人工核对）")
            r.font.size = Pt(9)
            note.alignment = WD_ALIGN_PARAGRAPH.CENTER

        d.save(str(out))
        if total_chars == 0:
            raise OcrTextError("一个字都没认出来，可能图片太模糊或是空白页")
        return {"pages": len(pages), "chars": total_chars,
                "paras": total_paras, "out": str(out)}
    finally:
        ocrmod._cleanup(tmpdir)


def to_text_docx(src: str | Path, out: str | Path, dpi: int = 200,
                 progress=None) -> dict:
    """OCR → Word，但**只保留文字**（不分栏、不还原位置，适合再排版）。

    和 to_docx 的区别：这个不做位置还原，字都是正文样式，
    适合「认出来接着改」；to_docx 适合「看起来还像原件」。
    """
    import docx
    from docx.shared import Pt

    import tempfile

    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    tmpdir = Path(tempfile.mkdtemp(prefix="gwoct_"))
    pages = _pages_of(src, dpi, progress, tmpdir=tmpdir)
    try:
        d = docx.Document()
        st = d.styles["Normal"]
        st.font.name = "仿宋_GB2312"
        st.font.size = Pt(16)
        total_chars = 0
        for pi, pg in enumerate(pages):
            if pi:
                d.add_page_break()
            rows = _group_lines(pg["items"])
            for para in _group_paragraphs(rows):
                text = "".join(it[4] for row in para for it in row)
                if text.strip():
                    d.add_paragraph(text)
                    total_chars += len(text)
        d.save(str(out))
        if total_chars == 0:
            raise OcrTextError("一个字都没认出来")
        return {"pages": len(pages), "chars": total_chars, "out": str(out)}
    finally:
        ocrmod._cleanup(tmpdir)


def to_pdf_double(src: str | Path, out: str | Path, dpi: int = 200,
                  progress=None) -> dict:
    """现成的 PDF（图片型）→ 双层 PDF。就是 ocr.from_pdf 的别名。"""
    return ocrmod.from_pdf(src, out, dpi=dpi, progress=progress)


def _columns(rows: list[list[tuple]], tol: float) -> list[float]:
    """按左边界聚出列位置（表格的列左对齐得很好）。"""
    lefts: list[float] = []
    for row in rows:
        for it in row:
            lefts.append(it[0])
    lefts.sort()
    cols: list[float] = []
    for x in lefts:
        if not cols or x - cols[-1] > tol:
            cols.append(x)
        else:
            cols[-1] = (cols[-1] + x) / 2
    return cols


def _mid_pdf(src: str | Path, tmpdir: Path, progress=None) -> Path:
    """把任意支持的文件变成 PDF（图片就地转一页 PDF），供找表格用。"""
    src = Path(src)
    s = src.suffix.lower()
    if s == ".pdf":
        return src
    if s in ocrmod.IMAGE_SOURCES:
        import fitz as _fitz

        img = _fitz.open(str(src))
        out = tmpdir / "img.pdf"
        pdf_bytes = img.convert_to_pdf()
        img.close()
        out.write_bytes(pdf_bytes)
        return out
    from . import office

    if progress:
        progress("转 PDF…")
    mid = tmpdir / "mid.pdf"
    office.to_pdf(src, mid)
    return mid


def to_xlsx_table(src: str | Path, out: str | Path, dpi: int = 200,
                  progress=None) -> dict:
    """把表格（图片/扫描件/PDF）识别成 Excel。逐页一个工作表。

    两条路：PDF 里有真正的表格线时用 PyMuPDF 的 find_tables（最准，
    合并单元格和换行都能对上）；没有表格线（扫描件、拍照）就退回
    OCR + 行列聚类。复杂表头仍可能需要人工核对。
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter

    tmpdir = None
    try:
        import tempfile

        tmpdir = Path(tempfile.mkdtemp(prefix="gwxlsx_"))
        fitz = pdfmod.load()

        # ---- 先试有表格线的 PDF（准确得多）
        try:
            pdf = _mid_pdf(src, tmpdir, progress)
            doc = fitz.open(str(pdf))
            try:
                found = [(i, pg.find_tables().tables) for i, pg in enumerate(doc)]
                found = [(i, ts) for i, ts in found if ts]
                if found:
                    wb = Workbook()
                    wb.remove(wb.active)
                    chars = 0
                    for pi, ts in found:
                        ws = wb.create_sheet(title=f"第{pi + 1}页"[:31])
                        ri = 0
                        for t in ts:
                            data = t.extract() or []
                            for row in data:
                                ri += 1
                                for ci, val in enumerate(row, start=1):
                                    txt = (val or "").replace("\n", " ").strip()
                                    if not txt:
                                        continue
                                    c = ws.cell(row=ri, column=ci, value=txt)
                                    c.alignment = Alignment(vertical="center",
                                                            wrap_text=True)
                                    chars += len(txt)
                            ri += 1      # 多个表之间空一行
                    for ws in wb.worksheets:
                        widths: dict[int, int] = {}
                        for row in ws.iter_rows():
                            for c in row:
                                if c.value:
                                    widths[c.column] = max(
                                        widths.get(c.column, 6),
                                        min(40, len(str(c.value)) * 2))
                        for ci, w in widths.items():
                            ws.column_dimensions[get_column_letter(ci)].width = w
                    if wb.sheetnames:
                        wb.save(str(out))
                        return {"pages": doc.page_count, "chars": chars,
                                "out": str(out), "via": "表格线"}
            finally:
                doc.close()
        except Exception as exc:  # noqa: BLE001
            log.debug("find_tables 不可用，改用 OCR：%s", exc)

        # ---- 退回 OCR + 行列聚类
        pages = _pages_of(src, dpi=dpi, progress=progress, tmpdir=tmpdir)

        wb = Workbook()
        wb.remove(wb.active)
        total = 0
        for pi, pg in enumerate(pages, start=1):
            rows = _group_lines(pg["items"])
            ws = wb.create_sheet(title=f"第{pi}页"[:31])
            if not rows:
                ws["A1"] = "（这一页没识别到文字）"
                continue
            hs = [(it[3] - it[1]) for row in rows for it in row] or [12.0]
            hs.sort()
            med_h = hs[len(hs) // 2] or 12.0
            cols = _columns(rows, tol=max(6.0, med_h * 0.9))
            if not cols:
                cols = [min(it[0] for row in rows for it in row)]

            def _col_of(x0: float) -> int:
                best, bi = 1e9, 1
                for i, cx in enumerate(cols, start=1):
                    d = abs(x0 - cx)
                    if d < best:
                        best, bi = d, i
                return bi

            page_rows: list[dict[int, list[str]]] = []
            for row in rows:
                cells: dict[int, list[str]] = {}
                for it in row:
                    cells.setdefault(_col_of(it[0]), []).append(it[4])
                page_rows.append(cells)
            # 判断到底是不是"表格"：如果绝大多数行只有一列，说明这是一段
            # 带缩进的正文（不是表），别硬拆成三列，按单列顺序写下来更好读。
            multi = sum(1 for c in page_rows if len(c) > 1)
            if multi < max(2, int(len(page_rows) * 0.3)):
                page_rows = [{1: " ".join(it[4] for it in row)} for row in rows]
            for ri, cells in enumerate(page_rows, start=1):
                for ci, parts in cells.items():
                    val = parts if isinstance(parts, str) else " ".join(parts)
                    c = ws.cell(row=ri, column=ci, value=val)
                    c.alignment = Alignment(vertical="center", wrap_text=True)
                    total += len(val)

            # 列宽按最长内容粗估
            widths: dict[int, int] = {}
            for row in ws.iter_rows():
                for c in row:
                    if c.value:
                        widths[c.column] = max(widths.get(c.column, 6),
                                               min(40, len(str(c.value)) * 2))
            for ci, w in widths.items():
                ws.column_dimensions[get_column_letter(ci)].width = w
            for c in ws[1]:
                c.font = Font(bold=True)

        if not wb.sheetnames:
            wb.create_sheet(title="第1页")
        wb.save(str(out))
        return {"pages": len(pages), "chars": total, "out": str(out)}
    finally:
        if tmpdir is not None:
            ocrmod._cleanup(tmpdir)
