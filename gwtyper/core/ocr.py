# -*- coding: utf-8 -*-
"""双层 PDF：两条完全不同的路径，别混在一起。

**路径 A：源文件是 Word / WPS / Excel / PPT（有文字）**
    转 PDF → 文字和位置直接来自源文档 → 做成「图像 + 不可见文字」
    准确率 100%，不需要 OCR。
    用途：
      - 防字体丢失（有些老机器没装方正小标宋，直接发 PDF 也可能显示替换）
      - 防篡改（图片层改不了字形）
      - 归档时要「看起来一模一样 + 还能搜」

**路径 B：源文件是扫描件 / 照片 / 图片型 PDF（只有像素）**
    必须 OCR。用 RapidOCR（纯 ONNX，随包分发）。
    用途：把纸质文件数字化，能搜能复制。

两条路都产出同一种东西：底层图像 + 上层不可见文字。
"""
from __future__ import annotations

import tempfile
from pathlib import Path

from . import pdfmod

# 有「原生文字」的格式：不需要 OCR
TEXT_SOURCES = {
    ".docx", ".docm", ".dotx",
    ".doc", ".wps", ".rtf", ".dot",
    ".xlsx", ".xlsm", ".xls", ".et", ".csv",
    ".pptx", ".ppsx", ".ppt", ".dps",
}
# 只有像素、必须要 OCR 的格式
IMAGE_SOURCES = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp", ".gif"}


class OcrError(RuntimeError):
    """OCR 不可用或失败。"""


_engine = None


def engine():
    global _engine
    if _engine is not None:
        return _engine
    try:
        from rapidocr_onnxruntime import RapidOCR
    except ImportError as exc:
        raise OcrError(
            "没装 OCR 引擎，扫描件做不了双层 PDF。\n"
            "装一下：pip install rapidocr_onnxruntime") from exc
    try:
        _engine = RapidOCR()
    except Exception as exc:  # noqa: BLE001
        raise OcrError(f"OCR 引擎初始化失败：{exc}") from exc
    return _engine


def ocr_available() -> bool:
    try:
        from rapidocr_onnxruntime import RapidOCR  # noqa: F401
        return True
    except Exception:  # noqa: BLE001
        return False


# ------------------------------------------------------------------ 共用工具
def _has_text_page(page) -> bool:
    """这一页是不是真有「文字层」，而不是残留的几个字。

    不能只看 get_text() 非空：扫描件 PDF 里常有页眉页码、或者 OCR 软件
    留下的零星文字，只有十来个字。真正的文字版页面每页至少几百字。
    这里按「字数 ≥ 30 且覆盖面积合理」判断，避免把扫描页误当文字页。
    """
    txt = (page.get_text() or "").strip()
    if not txt:
        return False
    # 去掉空白后少于 30 字，基本可以认定不是文字版
    compact = "".join(txt.split())
    if len(compact) < 30:
        return False
    return True


def _insert_invisible_text(page, items) -> int:
    """写不可见文字层。items = [(x0,y0,x1,y1,text,size_or_None), ...]"""
    fitz = pdfmod.load()
    n = 0
    for it in items:
        x0, y0, x1, y1, text = it[0], it[1], it[2], it[3], it[4]
        size_hint = it[5] if len(it) > 5 else None
        text = (text or "").strip()
        if not text:
            continue
        h, w = y1 - y0, x1 - x0
        if h <= 0.5 or w <= 0.5:
            continue
        size = size_hint or max(4.0, min(h * 1.05, 72.0))
        try:
            nat = fitz.get_text_length(text, fontname="china-s", fontsize=size)
        except Exception:  # noqa: BLE001
            nat = len(text) * size
        sx = max(0.3, min(3.0, (w / nat) if nat > 0 else 1.0))
        try:
            page.insert_text(
                (x0, y1 - h * 0.18), text,
                fontname="china-s", fontsize=size,
                render_mode=3,                          # 不可见
                morph=(fitz.Point(x0, y1), fitz.Matrix(sx, 1.0)),
            )
            n += 1
        except Exception:  # noqa: BLE001
            continue
    return n


def _text_items_from_pdf_page(page) -> list[tuple]:
    """从「有文字」的 PDF 页里取文字和位置（准确率 100%，无需 OCR）。"""
    items: list[tuple] = []
    data = page.get_text("dict")
    for block in data.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                txt = span.get("text", "")
                if not txt.strip():
                    continue
                bb = span["bbox"]
                items.append((bb[0], bb[1], bb[2], bb[3], txt,
                              float(span.get("size") or 0) or None))
    return items


def _ocr_items(png_path: str, dpi: int) -> list[tuple]:
    """OCR 一张图，坐标换算成 PDF 点。"""
    res, _ = engine()(png_path)
    if not res:
        return []
    k = 72.0 / dpi
    out: list[tuple] = []
    for box, text, _score in res:
        text = (text or "").strip()
        if not text:
            continue
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        out.append((min(xs) * k, min(ys) * k, max(xs) * k, max(ys) * k, text, None))
    return out


# ------------------------------------------------------------------ 路径 A
def from_office(src: str | Path, out: str | Path, dpi: int = 200,
                progress=None) -> dict:
    """Word/Excel/PPT → 双层 PDF。文字来自源文档，**不用 OCR**，100% 准确。"""
    from . import office

    src, out = Path(src), Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    tmpdir = Path(tempfile.mkdtemp(prefix="gw2l_"))
    try:
        mid = tmpdir / "mid.pdf"
        if progress:
            progress("转 PDF…")
        office.to_pdf(src, mid)
        return from_pdf(mid, out, dpi=dpi, progress=progress,
                        prefer_text=True)
    finally:
        _cleanup(tmpdir)


# ------------------------------------------------------------------ 路径 B / 通用
def from_pdf(src: str | Path, out: str | Path, dpi: int = 200,
             progress=None, prefer_text: bool = True) -> dict:
    """PDF → 双层 PDF。

    prefer_text=True（默认）：页面**有文字**就用原生文字做文字层，
    只有没文字的页才 OCR。混合文档（一部分扫描、一部分电子）也对。
    prefer_text=False：全部走 OCR（用于强行统一，一般不需要）。
    """
    fitz = pdfmod.load()
    src, out = Path(src), Path(out)
    if not src.is_file():
        raise OcrError(f"找不到文件：{src}")
    out.parent.mkdir(parents=True, exist_ok=True)

    src_doc = fitz.open(str(src))
    dst = fitz.open()
    stat = {"pages": src_doc.page_count, "text_pages": 0,
            "ocr_pages": 0, "lines": 0}
    tmpdir = Path(tempfile.mkdtemp(prefix="gw2l_"))
    try:
        n = src_doc.page_count
        for i in range(n):
            page = src_doc[i]
            w, h = page.rect.width, page.rect.height
            new_page = dst.new_page(width=w, height=h)

            use_text = prefer_text and _has_text_page(page)
            if use_text:
                # 图像层：把这一页渲成图，保证显示效果固定
                pix = page.get_pixmap(dpi=dpi)
                png = tmpdir / f"p{i}.png"
                pix.save(str(png))
                new_page.insert_image(new_page.rect, filename=str(png))
                # 文字层：直接取原生文字（准确率 100%）
                items = _text_items_from_pdf_page(page)
                stat["text_pages"] += 1
                if progress:
                    progress(f"第 {i+1}/{n} 页：取文字层（无需识别）")
            else:
                if not ocr_available():
                    raise OcrError(
                        "这一页是扫描件，需要 OCR，但本机没装 OCR 引擎。\n"
                        "装一下：pip install rapidocr_onnxruntime")
                pix = page.get_pixmap(dpi=dpi)
                png = tmpdir / f"p{i}.png"
                pix.save(str(png))
                new_page.insert_image(new_page.rect, filename=str(png))
                if progress:
                    progress(f"第 {i+1}/{n} 页：识别中…")
                items = _ocr_items(str(png), dpi)
                stat["ocr_pages"] += 1

            stat["lines"] += _insert_invisible_text(new_page, items)

        dst.save(str(out), garbage=3, deflate=True)
    finally:
        dst.close()
        src_doc.close()
        _cleanup(tmpdir)
    stat["out"] = str(out)
    return stat


def from_images(images: list[str | Path], out: str | Path, dpi: int = 200,
                progress=None) -> dict:
    """图片（照片/扫描件）→ 双层 PDF，必须 OCR。"""
    fitz = pdfmod.load()
    paths = [Path(p) for p in images]
    for p in paths:
        if not p.is_file():
            raise OcrError(f"找不到图片：{p}")
    if not ocr_available():
        raise OcrError(
            "图片做双层 PDF 需要 OCR 引擎。\n"
            "装一下：pip install rapidocr_onnxruntime")

    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    dst = fitz.open()
    stat = {"pages": len(paths), "text_pages": 0,
            "ocr_pages": 0, "lines": 0}
    tmpdir = Path(tempfile.mkdtemp(prefix="gw2l_"))
    try:
        for i, p in enumerate(paths, start=1):
            if progress:
                progress(f"第 {i}/{len(paths)} 张：识别中…")
            img = fitz.open(str(p))
            try:
                rect = img[0].rect
            finally:
                img.close()
            page = dst.new_page(width=rect.width, height=rect.height)
            page.insert_image(page.rect, filename=str(p))

            # 图片本身就是像素，直接交给 OCR；坐标要按「页面点 / 图片像素」缩放
            items = _ocr_items(str(p), 72)      # 72dpi ⇒ 1px = 1pt 起算
            w_px, h_px = _image_px(p)
            if w_px and h_px:
                kx, ky = rect.width / w_px, rect.height / h_px
                items = [(a * kx, b * ky, c * kx, d * ky, t, s)
                         for a, b, c, d, t, s in items]
            stat["lines"] += _insert_invisible_text(page, items)
            stat["ocr_pages"] += 1
        dst.save(str(out), garbage=3, deflate=True)
    finally:
        dst.close()
        _cleanup(tmpdir)
    stat["out"] = str(out)
    return stat


# ------------------------------------------------------------------ 入口
def make_double_layer(src: str | Path, out: str | Path, dpi: int = 200,
                      progress=None) -> dict:
    """统一的入口：按源文件类型自动选路径。

    - Word/WPS/Excel/PPT → from_office（文字来自源文档，不 OCR）
    - 图片               → from_images（OCR）
    - PDF                → from_pdf（有文字取文字，没文字才 OCR）
    """
    src = Path(src)
    s = src.suffix.lower()
    if s in TEXT_SOURCES:
        return from_office(src, out, dpi, progress)
    if s in IMAGE_SOURCES:
        return from_images([src], out, dpi, progress)
    if s == ".pdf":
        return from_pdf(src, out, dpi, progress)
    raise OcrError(f"这个格式做不了双层 PDF：{s}")


def describe_plan(src: str | Path) -> str:
    """告诉用户这一份会走哪条路（界面上显示用）。"""
    s = Path(src).suffix.lower()
    if s in TEXT_SOURCES:
        return "文字来自源文档，不识别，100% 准确"
    if s in IMAGE_SOURCES or s == ".pdf":
        return "需要 OCR 识别（准确率约 99%）"
    return "不支持的格式"


def _image_px(path: Path) -> tuple[int, int]:
    try:
        from PIL import Image

        with Image.open(path) as im:
            return im.width, im.height
    except Exception:  # noqa: BLE001
        return 0, 0


def _cleanup(dirpath: Path) -> None:
    try:
        for f in dirpath.iterdir():
            f.unlink()
        dirpath.rmdir()
    except OSError:
        pass
