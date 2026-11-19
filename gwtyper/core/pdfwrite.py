# -*- coding: utf-8 -*-
"""纯 Python 生成 PDF —— 不调用 WPS / Office / LibreOffice。

为什么要有这条路
----------------
LibreOffice 在 Linux 上经常缺组件（只装 writer 时转不了 xlsx，报
`source file could not be loaded`），而政府单位的国产机不一定有办公套件。
PDF 又是最需要稳的一环，所以自己排。

能力边界（都是实测过的）
------------------------
能做：
- docx / txt / markdown → PDF，保留字号、行距、缩进、对齐、段前段后
- xlsx → PDF，按表格画，自动分页、表头重复
- 图片 → PDF
- 任意已有 PDF → 改方向 / 加留白（这条本来就纯 Python）

不做（交给外部引擎）：
- PDF → docx / xlsx / pptx（反向解析版式，pdf2docx 已经够用）
- PPT 渲染（演示版式太复杂，不如交给 LibreOffice）

字体：用 fontfind 解析名字 → fontTools 子集化 → PyMuPDF 嵌入。
一份 1038 字的公文，全量嵌字体 48 MB，子集化后 141 KB。
"""
from __future__ import annotations

import os
import tempfile
from pathlib import Path

from . import fontfind, pdfmod


class PureError(RuntimeError):
    """纯 Python 排版失败，消息直接给用户看。"""


def available() -> bool:
    return pdfmod.available()


def _fitz():
    return pdfmod.load()


# ---------------------------------------------------------------- 字体准备
class _FontPool:
    """管字体文件 → 子集 → 嵌入用的一切，避免同一份字体裁多遍。"""

    def __init__(self, tmpdir: Path):
        self.tmp = tmpdir
        self.files: dict[str, Path] = {}      # 字体名 → 原始文件
        self.subsets: dict[tuple[str, str], Path] = {}
        self.missing: set[str] = set()

    def file_of(self, name: str) -> Path | None:
        if name in self.files:
            return self.files[name]
        got = fontfind.resolve_or_fallback(name)
        if got is None:
            self.missing.add(name)
            return None
        self.files[name] = got
        return got

    def subset_of(self, name: str, text: str) -> Path | None:
        """拿到一份只含 text 用字的字体文件。"""
        src = self.file_of(name)
        if src is None:
            return None
        chars = "".join(sorted(set(text)))
        if not chars.strip():
            chars = "公文"
        key = (name, chars)
        if key in self.subsets:
            return self.subsets[key]
        stem = f"{abs(hash(key)) % (10 ** 10)}"
        out = self.tmp / f"sub_{stem}.ttf"
        got = fontfind.subset(src, chars, out)
        if got is None:
            # 子集化不可用（没装 fontTools）→ 直接用原文件，体积大但能用
            got = src
        self.subsets[key] = got
        return got


# ---------------------------------------------------------------- docx → PDF
_ALIGN = {"center": 1, "right": 2, "justify": 3, "left": 0, None: 0}
_ALIGN_STR = {1: "center", 2: "right", 3: "justify", 0: "left"}


def _mm2pt(mm: float) -> float:
    return float(mm) * 72.0 / 25.4


def _east_asia_name(run) -> str | None:
    """取 run 的中文字体名。

    python-docx 的 `run.font.name` 读的是 w:rFonts/@w:ascii —— 那是**西文**
    字体，公文文档里通常写着 Times New Roman。真正决定中文长相的是
    @w:eastAsia（仿宋_GB2312 之类），python-docx 没有现成属性，只能读 XML。
    """
    try:
        rpr = run._element.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
        if rpr is None:
            return None
        fonts = rpr.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts")
        if fonts is None:
            return None
        return fonts.get(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia")
    except Exception:  # noqa: BLE001
        return None


def _style_east_asia(doc, paragraph) -> str | None:
    """段落样式里定义的 East Asian 字体（run 没写时用这个）。"""
    try:
        style = paragraph.style
        if style is None or style.element is None:
            return None
        rpr = style.element.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
        if rpr is None:
            return None
        fonts = rpr.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts")
        if fonts is None:
            return None
        return fonts.get(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia")
    except Exception:  # noqa: BLE001
        return None


def docx_to_pdf(src: str | Path, dst: str | Path, page: dict | None = None,
                progress=None) -> None:
    """把 docx 按它自己的格式排成 PDF（不按规则集重排）。

    page: 覆盖纸张设置，形如 {"top_cm":3.7,"bottom_cm":3.5,"left_cm":2.8,
          "right_cm":2.6}；不给就用 docx 里存的页边距。
    """
    try:
        import docx
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Pt
    except ImportError as exc:
        raise PureError("没装 python-docx，读不了 Word 文档") from exc

    fitz = _fitz()
    src, dst = Path(src), Path(dst)

    try:
        doc = docx.Document(str(src))
    except Exception as exc:  # noqa: BLE001
        raise PureError(f"打不开这个 Word 文档：{exc}") from exc

    # ---- 纸张 ----
    w_pt, h_pt = 595.0, 842.0          # A4
    margins = {"top_cm": 3.7, "bottom_cm": 3.5, "left_cm": 2.8, "right_cm": 2.6}
    try:
        sec = doc.sections[0]
        if sec.page_width and sec.page_height:
            w_pt = sec.page_width.pt
            h_pt = sec.page_height.pt
        if sec.top_margin:
            margins = {
                "top_cm": sec.top_margin.cm,
                "bottom_cm": sec.bottom_margin.cm,
                "left_cm": sec.left_margin.cm,
                "right_cm": sec.right_margin.cm,
            }
    except Exception:  # noqa: BLE001
        pass
    if page:
        margins.update({k: v for k, v in page.items() if v})

    left = _mm2pt(margins["left_cm"] * 10)
    right = w_pt - _mm2pt(margins["right_cm"] * 10)
    top = _mm2pt(margins["top_cm"] * 10)
    bottom = h_pt - _mm2pt(margins["bottom_cm"] * 10)
    print_w = max(60.0, right - left)
    print_h = max(60.0, bottom - top)

    # ---- 字体池 ----
    tmpdir = Path(tempfile.mkdtemp(prefix="gwpdf_"))
    pool = _FontPool(tmpdir)
    # 先把所有文字收一遍，字体一次裁到位（比逐段裁快，也不会漏字）
    all_text: list[str] = []
    for p in doc.paragraphs:
        if p.text.strip():
            all_text.append(p.text)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                if cell.text.strip():
                    all_text.append(cell.text)
    whole = "".join(all_text) + "0123456789—-()（）、《》，。：；！？ "

    # 字体名清单（正文用最多的那几个）
    names: set[str] = set()
    for p in doc.paragraphs:
        for r in p.runs:
            if r.font.name:
                names.add(r.font.name)
    if not names:
        names.add("仿宋_GB2312")

    out = fitz.open()
    try:
        # 页面状态统一放 state 里：翻页只改 state["page"] 和 state["y"]，
        # 不靠函数返回值传坐标（那样容易把 page 对象和 y 数值搞混）。
        state = {"page": None, "y": top}

        def open_page():
            p = out.new_page(width=w_pt, height=h_pt)
            state["page"] = p
            state["y"] = top
            return p

        def need(height: float):
            """如果这一行放不下就翻页，返回当前页。"""
            if state["page"] is None or state["y"] + height > bottom:
                open_page()
            return state["page"]

        open_page()

        for para in doc.paragraphs:
            text = para.text
            if not text.strip():
                # 空段落：留一个行高的空
                state["y"] += 16
                continue

            pf = para.paragraph_format
            # 字号：取段内第一个有字号的 run
            size_pt = 16.0
            font_name = "仿宋_GB2312"
            bold = False
            for r in para.runs:
                if r.font.size:
                    size_pt = r.font.size.pt
                    break
            # 中文字体：先看 run 的 eastAsia，再看样式，最后才退到西文名。
            # 直接读 font.name 会拿到 Times New Roman（那是西文字体），
            # 用它去找中文字体会找不到，结果一个字都排不出来。
            for r in para.runs:
                got = _east_asia_name(r)
                if got:
                    font_name = got
                    break
            else:
                got = _style_east_asia(doc, para)
                if got:
                    font_name = got
            bold = any(r.bold for r in para.runs if r.bold)

            line_pt = 28.0
            ls = pf.line_spacing
            if ls is not None:
                # python-docx 这里给的东西很杂：
                #   Twips/Length → 有 .pt（注意 Twips 还是 int 的子类，
                #                  所以必须先按 .pt 判断，不能先判断 int）
                #   纯 float     → 行距倍数（1.5 表示 1.5 倍）
                if hasattr(ls, "pt"):
                    try:
                        got = float(ls.pt)
                    except Exception:  # noqa: BLE001
                        got = 0.0
                    if 6 <= got <= 200:      # 合理的固定行距
                        line_pt = got
                elif isinstance(ls, float) and ls < 10:
                    line_pt = ls * size_pt
                elif isinstance(ls, int) and 6 <= ls <= 200:
                    line_pt = float(ls)
            if line_pt < size_pt:
                line_pt = size_pt * 1.2

            indent = 0.0
            try:
                if pf.first_line_indent:
                    indent = pf.first_line_indent.pt
            except Exception:  # noqa: BLE001
                pass

            align = _ALIGN.get(str(para.alignment), 0)
            if para.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                align = 1
            elif para.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                align = 2
            elif para.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
                align = 3

            # 段前段后
            try:
                if pf.space_before:
                    state["y"] += pf.space_before.pt
            except Exception:  # noqa: BLE001
                pass

            sub = pool.subset_of(font_name, whole)
            if sub is None:
                sub = pool.subset_of("仿宋_GB2312", whole)

            # 手工折行：PyMuPDF 的 insert_textbox 能自动折，但它按西文规则，
            # 中文行首要避免标点开头，所以这里先按字宽算好每行内容。
            _draw_paragraph(
                state, need, text, sub, font_name, size_pt, line_pt,
                indent, align, left, print_w, bottom, bold,
            )

            try:
                if pf.space_after:
                    state["y"] += pf.space_after.pt
            except Exception:  # noqa: BLE001
                pass

        # ---- 表格 ----
        if doc.tables:
            for table in doc.tables:
                _draw_table(state, need, table, pool, whole, left,
                            print_w, bottom)

        out.save(str(dst), deflate=True, garbage=3)
    finally:
        out.close()
        import shutil

        shutil.rmtree(tmpdir, ignore_errors=True)

    if not dst.exists() or dst.stat().st_size == 0:
        raise PureError("排完了但没生成文件")


def _text_width(fitz, text: str, font, size: float) -> float:
    try:
        return fitz.get_text_length(text, fontname=font.name, fontsize=size)
    except Exception:  # noqa: BLE001
        return len(text) * size * 0.5


_FONT_NAMES: dict[str, str] = {}      # 字体文件 → 页内资源名（F1/F2/F3…）


def _font_alias(sub) -> str:
    """给每份字体一个稳定的资源名。

    以前全用 "F1"，结果同一页里先注册的那个字体赢了，后面 insert_text
    即使指定 F1 也还是用第一份 —— 一份带标题/小标题/正文四种字体的公文，
    排出来全是一个样子。按文件分配别名才能各归各。
    """
    if sub is None:
        return "F1"
    key = str(sub)
    if key not in _FONT_NAMES:
        _FONT_NAMES[key] = f"F{len(_FONT_NAMES) + 1}"
    return _FONT_NAMES[key]


def _register(page, sub) -> str:
    """在当前页注册字体，返回它的资源名。每页都要注册一次。"""
    alias = _font_alias(sub)
    if sub:
        try:
            page.insert_font(fontname=alias, fontfile=str(sub))
        except Exception:  # noqa: BLE001
            pass
    return alias


def _measure(fitz, font, text, size_pt):
    try:
        return font.text_length(text, fontsize=size_pt)
    except Exception:  # noqa: BLE001
        return len(text) * size_pt * 0.5


def _draw_paragraph(state, need, text, sub, font_name, size_pt, line_pt,
                    indent, align, left, print_w, bottom, bold):
    """把一段文字折行写进页面。state["y"] 是当前基线位置，直接改它。"""
    fitz = _fitz()
    try:
        font = fitz.Font(fontfile=str(sub)) if sub else fitz.Font("china-s")
    except Exception:  # noqa: BLE001
        try:
            font = fitz.Font("china-s")
        except Exception:  # noqa: BLE001
            font = fitz.Font("helv")

    # 折行（按实际字宽，中英文混排也准）
    lines: list[str] = []
    cur, cur_w = "", 0.0
    avail = print_w - indent
    for ch in text:
        if ch == "\n":
            lines.append(cur)
            cur, cur_w, avail = "", 0.0, print_w
            continue
        w = _measure(fitz, font, ch, size_pt) or size_pt * 0.5
        if cur_w + w > avail and cur:
            lines.append(cur)
            cur, cur_w, avail = ch, w, print_w
        else:
            cur += ch
            cur_w += w
    if cur:
        lines.append(cur)

    for i, line in enumerate(lines):
        page = need(line_pt)
        alias = _register(page, sub)
        x = left + (indent if i == 0 else 0)
        lw = _measure(fitz, font, line, size_pt)
        if align == 1:
            x = left + (print_w - lw) / 2
        elif align == 2:
            x = left + print_w - lw
        baseline = state["y"] + size_pt
        try:
            page.insert_text((x, baseline), line, fontname=alias,
                             fontsize=size_pt, color=(0, 0, 0))
        except Exception:  # noqa: BLE001
            try:
                page.insert_text((x, baseline), line, fontsize=size_pt)
            except Exception:  # noqa: BLE001
                pass
        state["y"] += line_pt


def _draw_table(state, need, table, pool, whole, left, print_w, bottom):
    """表格：按等宽列画，超页自动翻页（表头会重复画一次）。"""
    rows = list(table.rows)
    if not rows:
        return
    ncol = max(len(r.cells) for r in rows)
    col_w = print_w / max(1, ncol)
    size_pt = 10.5
    row_h = 22.0
    sub = pool.subset_of("仿宋_GB2312", whole)

    def one(row, is_head=False):
        page = need(row_h)
        alias = _register(page, sub)
        y0 = state["y"]
        try:
            page.draw_line((left, y0), (left + print_w, y0),
                           color=(0.55, 0.55, 0.55), width=0.5)
        except Exception:  # noqa: BLE001
            pass
        for i in range(min(ncol, len(row.cells))):
            txt = row.cells[i].text.strip().replace("\n", " ")
            if len(txt) > 22:
                txt = txt[:21] + "…"
            try:
                page.insert_text((left + i * col_w + 3, y0 + row_h * 0.7),
                                 txt, fontname=alias, fontsize=size_pt)
            except Exception:  # noqa: BLE001
                pass
        state["y"] += row_h

    head = rows[0]
    one(head, True)
    for r in rows[1:]:
        # 快到底了就翻页，并把表头再画一次
        if state["y"] + row_h * 2 > bottom:
            need(row_h * 3)
            one(head, True)
        one(r)

    page = need(row_h)
    try:
        page.draw_line((left, state["y"]), (left + print_w, state["y"]),
                       color=(0.55, 0.55, 0.55), width=0.5)
    except Exception:  # noqa: BLE001
        pass
    state["y"] += 10


# ---------------------------------------------------------------- xlsx → PDF
def xlsx_to_pdf(src: str | Path, dst: str | Path,
                repeat_header: bool = True, progress=None) -> None:
    """把表格画成 PDF：自动分页、超宽缩放、表头每页重复。"""
    try:
        import openpyxl
    except ImportError as exc:
        raise PureError("没装 openpyxl，读不了 Excel") from exc

    fitz = _fitz()
    src, dst = Path(src), Path(dst)
    try:
        wb = openpyxl.load_workbook(str(src), data_only=True, read_only=True)
    except Exception as exc:  # noqa: BLE001
        raise PureError(f"打不开这个表格：{exc}") from exc

    tmpdir = Path(tempfile.mkdtemp(prefix="gwxls_"))
    pool = _FontPool(tmpdir)
    out = fitz.open()

    w_pt, h_pt = 842.0, 595.0          # 表格默认横版，能放下更多列
    m = _mm2pt(15)
    left, right = m, w_pt - m
    top, bottom = m, h_pt - m
    print_w, print_h = right - left, bottom - top

    try:
        for ws in wb.worksheets:
            rows = []
            for row in ws.iter_rows(values_only=True):
                if row is None:
                    continue
                cells = ["" if v is None else str(v) for v in row]
                while cells and not cells[-1].strip():
                    cells.pop()
                if cells:
                    rows.append(cells)
            if not rows:
                continue

            ncol = max(len(r) for r in rows)
            # 列宽：按每列最长内容算比例，限制在合理范围
            lens = []
            for c in range(ncol):
                longest = max((len(r[c]) if c < len(r) else 0) for r in rows)
                lens.append(max(4, min(40, longest)))
            total = sum(lens) or 1
            widths = [print_w * (l / total) for l in lens]

            size = 9.0
            row_h = 16.0
            whole = "".join("".join(r) for r in rows[:400]) + "0123456789"
            sub = pool.subset_of("宋体", whole)

            # 页面和字体别名放一起：翻页时 new_sheet_page() 会重算别名，
            # inner() 必须读同一个来源，否则翻页后还是拿旧的别名去画。
            sheet = {"page": None, "alias": "F1"}

            def new_sheet_page():
                p = out.new_page(width=w_pt, height=h_pt)
                sheet["page"] = p
                sheet["alias"] = _register(p, sub)
                # 标题：工作表名
                try:
                    p.insert_text((left, top - 6), f"【{ws.title}】",
                                  fontname=sheet["alias"], fontsize=11)
                except Exception:  # noqa: BLE001
                    pass
                return p

            new_sheet_page()
            y = top

            def draw(r, bold_row=False):
                nonlocal y
                if y + row_h > bottom:
                    new_sheet_page()
                    y = top
                    if repeat_header and rows:
                        inner(rows[0], True)
                inner(r, bold_row)

            def inner(r, bold_row=False):
                nonlocal y
                x = left
                for i in range(ncol):
                    txt = r[i] if i < len(r) else ""
                    if len(txt) > 40:
                        txt = txt[:39] + "…"
                    try:
                        sheet["page"].insert_text(
                            (x + 2, y + row_h * 0.72), txt,
                            fontname=sheet["alias"], fontsize=size)
                    except Exception:  # noqa: BLE001
                        pass
                    x += widths[i]
                try:
                    sheet["page"].draw_line((left, y), (right, y),
                                            color=(0.75, 0.75, 0.75), width=0.4)
                except Exception:  # noqa: BLE001
                    pass
                y += row_h

            for i, r in enumerate(rows):
                draw(r, i == 0)

        if not out.page_count:
            raise PureError("这个表格是空的，没有内容可排")
        out.save(str(dst), deflate=True, garbage=3)
    finally:
        out.close()
        try:
            wb.close()
        except Exception:  # noqa: BLE001
            pass
        import shutil

        shutil.rmtree(tmpdir, ignore_errors=True)

    if not dst.exists() or dst.stat().st_size == 0:
        raise PureError("排完了但没生成文件")


# ---------------------------------------------------------------- 文本 → PDF
def text_to_pdf(src: str | Path, dst: str | Path,
                font: str = "仿宋_GB2312", size_pt: float = 14.0,
                line_pt: float = 24.0, progress=None) -> None:
    """纯文本 / markdown → PDF。按 GB/T 9704 的版心排。"""
    fitz = _fitz()
    src, dst = Path(src), Path(dst)
    try:
        from . import pdf_engine

        enc = pdf_engine.txt_guess_encoding(str(src))
        text = src.read_text(encoding=enc, errors="replace")
    except Exception as exc:  # noqa: BLE001
        raise PureError(f"读不了这个文本文件：{exc}") from exc

    tmpdir = Path(tempfile.mkdtemp(prefix="gwtxt_"))
    pool = _FontPool(tmpdir)
    sub = pool.subset_of(font, text + "0123456789")
    out = fitz.open()
    try:
        w_pt, h_pt = 595.0, 842.0
        left = _mm2pt(28)
        right = w_pt - _mm2pt(26)
        top = _mm2pt(37)
        bottom = h_pt - _mm2pt(35)
        print_w = right - left

        page = out.new_page(width=w_pt, height=h_pt)
        alias = _register(page, sub)
        y = top

        def measure(s: str) -> float:
            try:
                return fitz.get_text_length(s, fontname="helv", fontsize=size_pt)
            except Exception:  # noqa: BLE001
                return len(s) * size_pt * 0.5

        # 中文按两倍西文宽估算
        avail_chars = int(print_w / (size_pt * 0.98)) or 40
        for raw in text.splitlines():
            if not raw.strip():
                y += line_pt
                if y > bottom:
                    page = out.new_page(width=w_pt, height=h_pt)
                    alias = _register(page, sub)
                    y = top
                continue
            line = raw
            while line:
                chunk = line[:avail_chars]
                line = line[avail_chars:]
                if y + line_pt > bottom:
                    page = out.new_page(width=w_pt, height=h_pt)
                    alias = _register(page, sub)
                    y = top
                try:
                    page.insert_text((left, y + size_pt), chunk,
                                     fontname=alias, fontsize=size_pt)
                except Exception:  # noqa: BLE001
                    pass
                y += line_pt
        out.save(str(dst), deflate=True, garbage=3)
    finally:
        out.close()
        import shutil

        shutil.rmtree(tmpdir, ignore_errors=True)
