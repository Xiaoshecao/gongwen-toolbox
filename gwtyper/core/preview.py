# -*- coding: utf-8 -*-
"""版面预览：把 docx 转成 PDF 再逐页渲染成图片，供界面显示。

两个用途：
- 原件视角：直接渲染用户给的文件，看它现在长什么样；
- 规则视角：先按规则集把内容排到临时文件里，再渲染，看改完会变成什么样。

转换走 COM（和版面体检同一条通道），大约 1~2 秒一份，所以界面里必须异步调用。
"""
from __future__ import annotations

import time
from pathlib import Path

from . import pdfmod, sysenv


def _cache_dir() -> Path:
    return sysenv.cache_dir("preview")


PAGE_GAP = 12          # 页与页之间的留白（像素）
DPI_RENDER = 150       # 先按这个 dpi 栅格化，再按窗口宽度缩放


class PreviewError(RuntimeError):
    pass


def _stamp(p: Path) -> tuple[int, int]:
    try:
        st = p.stat()
        return st.st_size, int(st.st_mtime)
    except OSError:
        return 0, 0


def _cache_key(prefix: str, path: Path) -> str:
    """缓存键：路径 + 大小 + 修改时间。

    只用文件名（stem）会出事：两个不同目录下的同名同大小文件
    （公文模板复制出来很常见）会命中同一份缓存，预览显示成别人的内容。
    doc_source 那边早就用了 sha1(路径|大小|mtime)，这里对齐。
    """
    import hashlib

    size, mtime = _stamp(path)
    raw = f"{path.resolve()}|{size}|{mtime}"
    digest = hashlib.sha1(raw.encode("utf-8")).hexdigest()[:12]
    return f"{prefix}-{path.stem}-{digest}"


def make_docx(src: str | Path, rules: dict | None = None,
              overrides: dict[int, str] | None = None,
              threshold: float = 0.45) -> Path:
    """返回要渲染的 docx。

    rules 为空 → 原件本身；给了规则 → 按规则排到缓存目录里的一份副本。
    """
    src = Path(src)
    if not rules:
        real = src
        if src.suffix.lower() != ".docx":
            from .doc_source import to_docx

            real = to_docx(src)
        return real

    from .apply import apply

    cache = _cache_dir()
    cache.mkdir(parents=True, exist_ok=True)
    out = cache / f"{src.stem}-规则预览.docx"
    if out.exists():
        try:
            out.unlink()
        except OSError:
            pass
    path, _doc, _review = apply(src, rules, out, threshold, overrides)
    return path


def pdf_of(docx_path: str | Path, force: bool = False) -> Path:
    """docx → pdf，按「大小+修改时间」复用缓存。"""
    from . import office

    docx_path = Path(docx_path)
    cache = _cache_dir()
    cache.mkdir(parents=True, exist_ok=True)
    key = _cache_key("pdf", docx_path)
    safe = "".join(ch if ch.isalnum() or ch in "._-（）()" else "_" for ch in key)
    dst = cache / f"{safe}.pdf"
    if dst.exists() and not force and dst.stat().st_size > 0:
        return dst

    if not office.engine_name():
        raise PreviewError(
            "没有可用的转换引擎，看不了版面预览。\n"
            "装一个 WPS 或 LibreOffice 就行。")
    try:
        office.word_to_pdf(docx_path, dst)
    except Exception as exc:  # noqa: BLE001
        raise PreviewError(f"转 PDF 失败（{exc}）。") from exc
    if not dst.exists() or dst.stat().st_size == 0:
        raise PreviewError("转 PDF 没有产出文件，可能文件损坏或受密码保护")
    return dst


MAX_PREVIEW_PAGES = 30      # 最多渲这么多页，超出的只报页数


def render(pdf_path: str | Path, max_width: int = 0, dpi: int = DPI_RENDER,
           max_pages: int = MAX_PREVIEW_PAGES) -> tuple[list[bytes], int]:
    """把每页渲染成 PNG 字节，返回 (页图列表, 总页数)。

    max_width>0 时按宽度等比缩放——窗口多宽就渲多大，别拿 150dpi 的整幅图
    去缩略显示。一张 A4@150dpi 是 1240×1754，40 页就是 8700 万像素，
    Windows 上很容易把 GDI 位图配额用光（表现为
    `Tk_GetPixmap: Error from CreateDIBSSection`）。

    max_pages 限制渲染页数；超出部分不渲，只在页数里体现。
    """
    try:
        pymupdf = pdfmod.load()
    except pdfmod.PdfMissing as exc:
        raise PreviewError("未安装 PyMuPDF") from exc

    doc = pymupdf.open(str(pdf_path))
    try:
        n = len(doc)
        limit = max(1, min(n, max_pages))
        pages: list[bytes] = []
        for i in range(limit):
            page = doc[i]
            use_dpi = dpi
            if max_width > 0:
                w = page.rect.width * dpi / 72.0
                if w > max_width:
                    use_dpi = dpi * max_width / w
            # 72dpi 以下细节会糊，但比崩掉好
            pix = page.get_pixmap(dpi=max(60, min(dpi, int(use_dpi))))
            pages.append(pix.tobytes("png"))
        return pages, n
    finally:
        doc.close()


def any_pdf(src: str | Path, force: bool = False) -> Path:
    """把任意能转的文件（Office / PDF / 老格式）弄成一份 PDF，给右下预览用。

    和 `pdf_of` 的区别：这里不管 docx 规则，只要一份能看的 PDF；
    源文件本身是 PDF 就直接用，图片就合成一份。
    """
    from . import office

    src = Path(src)
    s = src.suffix.lower()
    cache = _cache_dir()
    cache.mkdir(parents=True, exist_ok=True)
    key = _cache_key("any", src)
    safe = "".join(ch if ch.isalnum() or ch in "._-（）()" else "_" for ch in key)
    dst = cache / f"{safe}.pdf"
    if dst.exists() and not force and dst.stat().st_size > 0:
        return dst

    if s == ".pdf":
        return src

    if s in {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp", ".gif"}:
        from . import pdf_engine

        pdf_engine.images_to_pdf([str(src)], str(dst))
    elif s in {".txt", ".md"}:
        raise PreviewError("纯文本没有版面，看不了")
    else:
        if not office.engine_name():
            raise PreviewError(
                "没有可用的转换引擎，看不了预览。\n装一个 WPS 或 LibreOffice 就行。")
        try:
            office.to_pdf(src, dst)
        except Exception as exc:  # noqa: BLE001
            raise PreviewError(f"转 PDF 失败（{exc}）。") from exc

    if not dst.exists() or dst.stat().st_size == 0:
        raise PreviewError("转了但没有产出文件，可能文件损坏或加了密码")
    return dst


def page_count(pdf_path: str | Path) -> int:
    try:
        pymupdf = pdfmod.load()

        d = pymupdf.open(str(pdf_path))
        try:
            return len(d)
        finally:
            d.close()
    except Exception:  # noqa: BLE001
        return 0


def prune(max_age_days: int = 7, keep: int = 40) -> int:
    base = _cache_dir()
    if not base.exists():
        return 0
    now = time.time()
    files = sorted(base.glob("*"), key=lambda p: p.stat().st_mtime, reverse=True)
    n = 0
    for i, f in enumerate(files):
        if not f.is_file():
            continue
        try:
            if i >= keep or (now - f.stat().st_mtime) > max_age_days * 86400:
                f.unlink()
                n += 1
        except OSError:
            continue
    return n
