# -*- coding: utf-8 -*-
"""给 PDF 批量加斜向水印（内部资料、注意保存、禁止外传这类）。"""
from __future__ import annotations

from pathlib import Path


def _fitz():
    from . import pdfmod

    return pdfmod.load()


def _apply_and_save(src, dst, text: str, opacity: float = 0.12, size: float = 0) -> str:
    fitz = _fitz()
    doc = fitz.open(str(src))
    try:
        if not text.strip():
            raise ValueError("水印文字不能为空。")
        for page in doc:
            r = page.rect
            fs = float(size) if size else max(18.0, min(r.width, r.height) * 0.075)
            cx, cy = (r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2
            tw = fitz.TextWriter(r)
            font = fitz.Font("china-s")
            width = font.text_length(text, fontsize=fs)
            tw.append((cx - width / 2, cy), text, font=font, fontsize=fs)
            try:
                tw.write_text(page, color=(0.45, 0.45, 0.45), opacity=float(opacity),
                              morph=(fitz.Point(cx, cy), fitz.Matrix(-30)))
            except TypeError:
                tw.write_text(page, color=(0.55, 0.55, 0.55))
        Path(dst).parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(dst), garbage=3, deflate=True)
    finally:
        doc.close()
    return str(dst)


def apply(src, out, text: str, opacity: float = 0.12, size: float = 0) -> str:
    return _apply_and_save(src, out, text, opacity, size)


def apply_many(paths, out_dir, text: str, opacity: float = 0.12,
               size: float = 0, progress=None) -> list[dict]:
    res = []
    total = len(paths)
    for i, p in enumerate(paths, 1):
        p = Path(p)
        try:
            dst = Path(out_dir) / f"{p.stem}-水印.pdf"
            _apply_and_save(p, dst, text, opacity, size)
            res.append({"src": str(p), "out": str(dst), "ok": True, "msg": ""})
        except Exception as exc:  # noqa: BLE001
            res.append({"src": str(p), "out": "", "ok": False, "msg": str(exc)})
        if progress:
            try:
                progress(i, total)
            except Exception:  # noqa: BLE001
                pass
    return res
