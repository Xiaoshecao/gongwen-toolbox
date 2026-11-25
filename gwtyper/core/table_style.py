# -*- coding: utf-8 -*-
"""表格样式：从一份「样表」里学出表格的版式参数，再套到别的表格上。

学的项目：
  · 页面：四边页边距、纸张方向
  · 表格：整表对齐、表格宽度、边框粗细/线型/颜色、单元格内边距
  · 列宽（逐列，厘米）、行高（逐行，厘米）
  · 表头行：字体、字号、加粗、水平对齐、垂直对齐
  · 正文行：字体、字号、水平对齐、垂直对齐、行距

用法（内核层）：
    st = learn_table("样表.docx")            # -> dict
    save_style(st, "我的表格样式")            # 存成 yaml
    apply_table("待排版.docx", "输出.docx", st)
"""
from __future__ import annotations

import copy
from pathlib import Path

import yaml
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from . import sysenv

TWIP_CM = 2.54 / 1440.0          # 1 twip = 1/20 磅 = 2.54/1440 cm
EMU_CM = 1 / 360000.0            # 1 cm = 360000 EMU


# ---------------------------------------------------------------- 存储位置
def styles_dir() -> Path:
    d = sysenv.data_dir() / "table_styles"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _safe(name: str) -> str:
    keep = [c for c in name if c not in '\\/:*?"<>|']
    s = "".join(keep).strip() or "表格样式"
    return s[:40]


def save_style(style: dict, name: str | None = None) -> Path:
    if name:
        style = dict(style, name=name)
    p = styles_dir() / f"{_safe(style.get('name') or '表格样式')}.yaml"
    p.write_text(
        "# 由「从样表学习」自动生成，可手工微调。\n"
        + yaml.safe_dump(style, allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )
    return p


def list_styles() -> list[dict]:
    out = []
    for p in sorted(styles_dir().glob("*.yaml")):
        try:
            d = yaml.safe_load(p.read_text(encoding="utf-8")) or {}
        except Exception:
            continue
        d.setdefault("name", p.stem)
        d["_path"] = str(p)
        out.append(d)
    return out


def load_style(name_or_path: str) -> dict:
    p = Path(name_or_path)
    if not p.exists():
        p = styles_dir() / f"{_safe(name_or_path)}.yaml"
    d = yaml.safe_load(p.read_text(encoding="utf-8")) or {}
    d.setdefault("name", p.stem)
    d["_path"] = str(p)
    return d


def delete_style(name_or_path: str) -> bool:
    p = Path(name_or_path)
    if not p.exists():
        p = styles_dir() / f"{_safe(name_or_path)}.yaml"
    if p.exists():
        p.unlink()
        return True
    return False


# ---------------------------------------------------------------- 读取工具
def _el(parent, tag):
    if parent is None:
        return None
    return parent.find(qn(tag))


def _int(v, default=None):
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def _twip_cm(v):
    v = _int(v)
    return None if v is None else round(v * TWIP_CM, 2)


def _cell_font(cell):
    """取单元格里第一个有文字的 run 的字体/字号/加粗（中文字体优先 eastAsia）。"""
    for p in cell.paragraphs:
        for r in p.runs:
            if not (r.text or "").strip():
                continue
            rpr = r._element.find(qn("w:rPr"))
            east = None
            if rpr is not None:
                rf = rpr.find(qn("w:rFonts"))
                if rf is not None:
                    east = rf.get(qn("w:eastAsia")) or rf.get(qn("w:ascii"))
            size = r.font.size.pt if r.font.size else None
            bold = r.bold
            if east or size or bold is not None:
                return {"font": east, "size_pt": size, "bold": bool(bold) if bold is not None else None}
    # 退回段落样式
    p = cell.paragraphs[0]
    st = p.style.font if p.style is not None else None
    return {
        "font": (st.name if st is not None else None),
        "size_pt": (st.size.pt if st is not None and st.size else None),
        "bold": None,
    }


def _cell_align(cell):
    p = cell.paragraphs[0]
    a = p.alignment
    if a is None:
        ppr = p._p.find(qn("w:pPr"))
        jc = _el(ppr, "w:jc")
        if jc is not None:
            v = jc.get(qn("w:val"))
            return {"left": "left", "center": "center", "right": "right",
                    "both": "justify"}.get(v, v)
        return None
    return {WD_ALIGN_PARAGRAPH.LEFT: "left", WD_ALIGN_PARAGRAPH.CENTER: "center",
            WD_ALIGN_PARAGRAPH.RIGHT: "right", WD_ALIGN_PARAGRAPH.JUSTIFY: "justify"}.get(a)


def _cell_valign(cell):
    tcpr = cell._tc.find(qn("w:tcPr"))
    va = _el(tcpr, "w:vAlign")
    if va is None:
        return None
    return {"top": "top", "center": "center", "bottom": "bottom"}.get(va.get(qn("w:val")))


def _borders(tbl_pr):
    b = _el(tbl_pr, "w:tblBorders")
    if b is None:
        return None
    out = {}
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = _el(b, f"w:{side}")
        if e is None:
            continue
        out[side] = {
            "val": e.get(qn("w:val")),
            "sz": _int(e.get(qn("w:sz"))),
            "color": e.get(qn("w:color")),
        }
    return out or None


def _cell_margins(tbl_pr):
    m = _el(tbl_pr, "w:tblCellMar")
    if m is None:
        return None
    out = {}
    for side in ("top", "left", "bottom", "right"):
        e = _el(m, f"w:{side}")
        if e is not None:
            out[side] = _twip_cm(e.get(qn("w:w")))
    return out or None


# ---------------------------------------------------------------- 学习
def learn_table(path: str | Path, name: str | None = None) -> dict:
    """读一份样表（Excel .xlsx，或 Word .docx 里的第一个表格），学出版式参数。"""
    _suf = Path(path).suffix.lower()
    if _suf in XLSX_EXT:
        return learn_xlsx(path, name)
    if _suf in XLS_EXT:
        raise ValueError(_OLD_XLS_MSG)
    doc = Document(str(path))
    if not doc.tables:
        raise ValueError("这份文件里没找到表格。请给一份带表格的 Word 文件（.docx）。")

    sec = doc.sections[0]
    style: dict = {
        "name": name or f"样表：{Path(path).stem}",
        "source": Path(path).name,
        "page": {
            "top_cm": round(sec.top_margin.cm, 2) if sec.top_margin else None,
            "bottom_cm": round(sec.bottom_margin.cm, 2) if sec.bottom_margin else None,
            "left_cm": round(sec.left_margin.cm, 2) if sec.left_margin else None,
            "right_cm": round(sec.right_margin.cm, 2) if sec.right_margin else None,
            "orientation": "landscape" if sec.page_width and sec.page_height and sec.page_width > sec.page_height else "portrait",
        },
    }

    t = doc.tables[0]
    tbl = t._tbl
    tbl_pr = tbl.find(qn("w:tblPr"))

    jc = _el(tbl_pr, "w:jc")
    align = None
    if jc is not None:
        align = {"left": "left", "center": "center", "right": "right"}.get(jc.get(qn("w:val")))
    tw = _el(tbl_pr, "w:tblW")
    width = None
    if tw is not None:
        width = {"type": tw.get(qn("w:type")), "w": _int(tw.get(qn("w:w")))}

    col_widths = []
    grid = _el(tbl, "w:tblGrid")
    if grid is not None:
        for gc in grid.findall(qn("w:gridCol")):
            col_widths.append(_twip_cm(gc.get(qn("w:w"))))

    row_heights = []
    for tr in tbl.findall(qn("w:tr")):
        trpr = _el(tr, "w:trPr")
        h = _el(trpr, "w:trHeight")
        row_heights.append(_twip_cm(h.get(qn("w:val"))) if h is not None else None)

    header_cell = t.rows[0].cells[0]
    body_cell = t.rows[1].cells[0] if len(t.rows) > 1 else header_cell
    hf = _cell_font(header_cell)
    bf = _cell_font(body_cell)

    style["table"] = {
        "align": align,
        "width": width,
        "borders": _borders(tbl_pr),
        "cell_margins_cm": _cell_margins(tbl_pr),
        "col_widths_cm": col_widths,
        "row_heights_cm": row_heights,
        "header_rows": 1,
        "header": dict(hf, align=_cell_align(header_cell), valign=_cell_valign(header_cell)),
        "body": dict(bf, align=_cell_align(body_cell), valign=_cell_valign(body_cell)),
    }
    return style


# ---------------------------------------------------------------- Excel（.xlsx）支持
XLSX_EXT = {".xlsx", ".xlsm"}
XLS_EXT = {".xls", ".et", ".ett"}
_OLD_XLS_MSG = "老 .xls / .et 读不了，请先在 Excel 或 WPS 里另存为 .xlsx 再学。"


def _is_xlsx(path) -> bool:
    return Path(path).suffix.lower() in XLSX_EXT


def _cm(inch_val):
    try:
        return round(float(inch_val) * 2.54, 2)
    except (TypeError, ValueError):
        return None


def _cell_style_info(cell) -> dict:
    a = cell.alignment
    f = cell.font
    return {
        "font": f.name,
        "size_pt": float(f.sz) if f.sz else None,
        "bold": bool(f.bold),
        "align": a.horizontal,
        "valign": a.vertical,
        "wrap": bool(a.wrap_text),
    }


def learn_xlsx(path, name=None) -> dict:
    """从一份 Excel 样表学：列宽、行高、表头/正文字体字号、对齐、边框、打印页边距。"""
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter

    wb = load_workbook(str(path))
    ws = wb.active
    pm = ws.page_margins
    max_row = max(1, ws.max_row or 1)
    max_col = max(1, ws.max_column or 1)

    widths = []
    for i in range(1, max_col + 1):
        w = ws.column_dimensions[get_column_letter(i)].width
        widths.append(round(float(w), 2) if w else None)

    heights = []
    for r in range(1, max_row + 1):
        h = ws.row_dimensions[r].height
        heights.append(round(float(h), 1) if h else None)

    head = ws.cell(1, 1)
    body = ws.cell(2, 1) if max_row > 1 else head
    b = head.border
    borders = {}
    for side in ("left", "right", "top", "bottom"):
        sd = getattr(b, side, None)
        if sd is not None and sd.style:
            borders[side] = {"style": sd.style, "color": (sd.color.rgb if sd.color is not None and sd.color.rgb else None)}

    return {
        "name": name or f"样表：{Path(path).stem}",
        "source": Path(path).name,
        "kind": "xlsx",
        "page": {
            "top_cm": _cm(pm.top), "bottom_cm": _cm(pm.bottom),
            "left_cm": _cm(pm.left), "right_cm": _cm(pm.right),
            "orientation": ws.page_setup.orientation or "portrait",
        },
        "table": {
            "align": None,
            "col_widths_ch": widths,
            "row_heights_pt": heights,
            "header_rows": 1,
            "header": _cell_style_info(head),
            "body": _cell_style_info(body),
            "borders": borders,
        },
    }


def apply_xlsx(src, out, style: dict) -> dict:
    """把样式套到 Excel 的每个工作表上（值不动，只改列宽行高字体对齐边框页边距）。"""
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Border, Font, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.page import PageMargins

    wb = load_workbook(str(src))
    spec = style.get("table") or {}
    pg = style.get("page") or {}
    widths = spec.get("col_widths_ch") or []
    heights = spec.get("row_heights_pt") or []
    head = spec.get("header") or {}
    body = spec.get("body") or {}
    hrows = int(spec.get("header_rows") or 1)
    bd = spec.get("borders") or {}

    def _bd():
        if not bd:
            return None
        return Border(**{k: Side(style=v.get("style"), color=v.get("color")) for k, v in bd.items()})

    cells = 0
    for ws in wb.worksheets:
        max_row = max(1, ws.max_row or 1)
        max_col = max(1, ws.max_column or 1)
        for i, w in enumerate(widths[:max_col], start=1):
            if w:
                ws.column_dimensions[get_column_letter(i)].width = w
        for r, h in enumerate(heights[:max_row], start=1):
            if h:
                ws.row_dimensions[r].height = h
        for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
            for cell in row:
                info = head if cell.row <= hrows else body
                if info.get("font") or info.get("size_pt") or info.get("bold") is not None:
                    cell.font = Font(name=info.get("font") or cell.font.name,
                                     sz=info.get("size_pt") or cell.font.sz,
                                     bold=info.get("bold"))
                cell.alignment = Alignment(horizontal=info.get("align"),
                                           vertical=info.get("valign"),
                                           wrap_text=bool(info.get("wrap")))
                b = _bd()
                if b is not None:
                    cell.border = b
                cells += 1
        if any(pg.get(k) for k in ("top_cm", "bottom_cm", "left_cm", "right_cm")):
            ws.page_margins = PageMargins(
                left=(pg.get("left_cm") or 2.54) / 2.54,
                right=(pg.get("right_cm") or 2.54) / 2.54,
                top=(pg.get("top_cm") or 2.54) / 2.54,
                bottom=(pg.get("bottom_cm") or 2.54) / 2.54,
            )
        if pg.get("orientation"):
            ws.page_setup.orientation = pg["orientation"]

    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    return {"out": str(out), "sheets": len(wb.worksheets), "cells": cells}


def describe(style: dict) -> str:
    """给界面显示的一行摘要。"""
    t = style.get("table", {})
    pg = style.get("page", {})
    h = t.get("header", {})
    b = t.get("body", {})
    if style.get("kind") == "xlsx":
        cols = len([c for c in t.get("col_widths_ch", []) if c])
        h = t.get("header") or {}
        b = t.get("body") or {}
        return (
            f"Excel 样表；页边距 上{pg.get('top_cm')} 下{pg.get('bottom_cm')} 左{pg.get('left_cm')} 右{pg.get('right_cm')} cm；"
            f"{cols} 列；表头 {h.get('font') or ''} {h.get('size_pt') or ''}磅"
            f"{' 加粗' if h.get('bold') else ''}；正文 {b.get('font') or ''} {b.get('size_pt') or ''}磅"
        )
    cols = len([c for c in t.get("col_widths_cm", []) if c])
    return (
        f"页边距 上{pg.get('top_cm')} 下{pg.get('bottom_cm')} 左{pg.get('left_cm')} 右{pg.get('right_cm')} cm；"
        f"{cols} 列；表头 {h.get('font') or ''} {h.get('size_pt') or ''}磅"
        f"{' 加粗' if h.get('bold') else ''}；正文 {b.get('font') or ''} {b.get('size_pt') or ''}磅"
    )


# ---------------------------------------------------------------- 套用
_ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}
_TBL_ALIGN = {
    "left": WD_TABLE_ALIGNMENT.LEFT,
    "center": WD_TABLE_ALIGNMENT.CENTER,
    "right": WD_TABLE_ALIGNMENT.RIGHT,
}


def _set_run_font(run, font, size_pt, bold=None):
    if font:
        run.font.name = font
        rpr = run._element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = rpr.makeelement(qn("w:rFonts"), {})
            rpr.append(rf)
        for attr in ("w:eastAsia", "w:ascii", "w:hAnsi"):
            rf.set(qn(attr), font)
        for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            if rf.get(qn(attr)) is not None:
                del rf.attrib[qn(attr)]
    if size_pt:
        run.font.size = Pt(float(size_pt))
    if bold is not None:
        run.font.bold = bool(bold)


def _set_cell_valign(cell, valign):
    if not valign:
        return
    tcpr = cell._tc.get_or_add_tcPr()
    old = tcpr.find(qn("w:vAlign"))
    if old is not None:
        tcpr.remove(old)
    e = tcpr.makeelement(qn("w:vAlign"), {qn("w:val"): valign})
    tcpr.append(e)


def _apply_borders(tbl_pr, borders):
    if not borders:
        return
    old = tbl_pr.find(qn("w:tblBorders"))
    if old is not None:
        tbl_pr.remove(old)
    b = tbl_pr.makeelement(qn("w:tblBorders"), {})
    for side, spec in borders.items():
        e = b.makeelement(qn(f"w:{side}"), {})
        if spec.get("val"):
            e.set(qn("w:val"), spec["val"])
        if spec.get("sz") is not None:
            e.set(qn("w:sz"), str(spec["sz"]))
        if spec.get("color"):
            e.set(qn("w:color"), spec["color"])
        b.append(e)
    tbl_pr.append(b)


def _apply_cell_margins(tbl_pr, margins):
    if not margins:
        return
    old = tbl_pr.find(qn("w:tblCellMar"))
    if old is not None:
        tbl_pr.remove(old)
    m = tbl_pr.makeelement(qn("w:tblCellMar"), {})
    for side in ("top", "left", "bottom", "right"):
        v = margins.get(side)
        if v is None:
            continue
        e = m.makeelement(qn(f"w:{side}"), {qn("w:w"): str(int(round(v / TWIP_CM))), qn("w:type"): "dxa"})
        m.append(e)
    tbl_pr.append(m)


def _apply_table(t, spec: dict):
    tbl = t._tbl
    tbl_pr = tbl.find(qn("w:tblPr"))
    if tbl_pr is None:
        tbl_pr = tbl.makeelement(qn("w:tblPr"), {})
        tbl.insert(0, tbl_pr)

    if spec.get("align") in _TBL_ALIGN:
        t.alignment = _TBL_ALIGN[spec["align"]]
    _apply_borders(tbl_pr, spec.get("borders"))
    _apply_cell_margins(tbl_pr, spec.get("cell_margins_cm"))

    cols = spec.get("col_widths_cm") or []
    heights = spec.get("row_heights_cm") or []
    hdr = spec.get("header") or {}
    body = spec.get("body") or {}
    header_rows = int(spec.get("header_rows") or 1)

    for ri, row in enumerate(t.rows):
        is_head = ri < header_rows
        fmt = hdr if is_head else body
        # 行高
        if ri < len(heights) and heights[ri]:
            trpr = row._tr.get_or_add_trPr()
            old = trpr.find(qn("w:trHeight"))
            if old is not None:
                trpr.remove(old)
            trpr.append(trpr.makeelement(qn("w:trHeight"), {
                qn("w:val"): str(int(round(heights[ri] / TWIP_CM))),
                qn("w:hRule"): "atLeast",
            }))
        for ci, cell in enumerate(row.cells):
            # 列宽
            if ci < len(cols) and cols[ci]:
                cell.width = Cm(float(cols[ci]))
            _set_cell_valign(cell, fmt.get("valign"))
            for p in cell.paragraphs:
                if fmt.get("align") in _ALIGN:
                    p.alignment = _ALIGN[fmt["align"]]
                if not p.runs:
                    # 空段落也要占位设字体，便于以后输入
                    r = p.add_run("")
                    _set_run_font(r, fmt.get("font"), fmt.get("size_pt"), fmt.get("bold"))
                for r in p.runs:
                    _set_run_font(r, fmt.get("font"), fmt.get("size_pt"), fmt.get("bold"))


def apply_table(src: str | Path, out: str | Path, style: dict) -> dict:
    """把样式套到 src 里所有表格上，写到 out。返回处理统计。"""
    _suf = Path(src).suffix.lower()
    if _suf in XLSX_EXT:
        return apply_xlsx(src, out, style)
    if _suf in XLS_EXT:
        raise ValueError(_OLD_XLS_MSG)
    doc = Document(str(src))
    if not doc.tables:
        raise ValueError("这份文件里没有表格，没什么可排的。")

    pg = style.get("page") or {}
    for sec in doc.sections:
        if pg.get("top_cm"):
            sec.top_margin = Cm(float(pg["top_cm"]))
        if pg.get("bottom_cm"):
            sec.bottom_margin = Cm(float(pg["bottom_cm"]))
        if pg.get("left_cm"):
            sec.left_margin = Cm(float(pg["left_cm"]))
        if pg.get("right_cm"):
            sec.right_margin = Cm(float(pg["right_cm"]))

    spec = style.get("table") or {}
    for t in doc.tables:
        _apply_table(t, spec)

    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    return {"out": str(out), "tables": len(doc.tables)}
