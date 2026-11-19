# -*- coding: utf-8 -*-
"""数据汇总：把各村报上来的同格式 Excel 合成一张总表。"""
from __future__ import annotations

from pathlib import Path


def _rows(path) -> list[list]:
    from openpyxl import load_workbook

    wb = load_workbook(str(path), read_only=True, data_only=True)
    try:
        ws = wb.active
        out = []
        for row in ws.iter_rows(values_only=True):
            vals = ["" if v is None else v for v in row]
            if any(str(v).strip() for v in vals):
                out.append(vals)
        return out
    finally:
        wb.close()


def summarize(files, out, source_col: str = "来源文件", progress=None) -> dict:
    """每个文件取第一个工作表，表头取第一条非空行，纵向拼接。

    返回 {"files": n, "rows": 数据行数, "cols": 列数, "skipped": [(文件, 原因)]}
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font

    files = [Path(f) for f in files]
    skipped: list[tuple[str, str]] = []
    headers: list[str] = []
    data: list[list] = []
    total = len(files)

    for i, f in enumerate(files, 1):
        try:
            rows = _rows(f)
        except Exception as exc:  # noqa: BLE001
            skipped.append((f.name, f"读不了：{exc}"))
            continue
        if not rows:
            skipped.append((f.name, "空表"))
            continue
        head = [str(v).strip() for v in rows[0]]
        body = rows[1:]
        for h in head:
            if h and h not in headers:
                headers.append(h)
        for r in body:
            rec = {h: (r[j] if j < len(r) else "") for j, h in enumerate(head) if h}
            rec[source_col] = f.name
            data.append(rec)
        if progress:
            try:
                progress(i, total)
            except Exception:  # noqa: BLE001
                pass

    if not data:
        raise ValueError("这些文件里没读到数据。确认是 .xlsx 表格，且第一行是表头。")

    cols = headers + [source_col]
    wb = Workbook()
    ws = wb.active
    ws.title = "汇总"
    ws.append(cols)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
    for rec in data:
        ws.append([rec.get(c, "") for c in cols])
    widths: dict[str, int] = {}
    for rec in data:
        for c in cols:
            widths[c] = max(widths.get(c, len(str(c))), len(str(rec.get(c, ""))))
    from openpyxl.utils import get_column_letter

    for j, c in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(j)].width = min(40, max(8, widths.get(c, 8) + 2))
    ws.freeze_panes = "A2"

    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    return {"files": total - len(skipped), "rows": len(data), "cols": len(cols), "skipped": skipped}
