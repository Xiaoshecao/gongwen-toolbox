# -*- coding: utf-8 -*-
"""图片批量压缩：手机拍的扫描件动辄 5～10MB，压到 1MB 以内还能看清。"""
from __future__ import annotations

import io
import os
from pathlib import Path

IMG_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tif", ".tiff"}


class CompressError(RuntimeError):
    pass


def is_image(path) -> bool:
    return Path(path).suffix.lower() in IMG_EXTS


def compress(src, out=None, max_side: int = 1600, quality: int = 75) -> tuple[Path, int, int]:
    """压一张图。返回 (输出路径, 原字节, 新字节)。"""
    from PIL import Image, ImageOps

    src = Path(src)
    if not src.exists():
        raise CompressError(f"找不到文件：{src}")
    before = src.stat().st_size
    try:
        im = Image.open(src)
    except Exception as exc:  # noqa: BLE001
        raise CompressError(f"这个图片打不开：{src.name}（{exc}）") from exc

    im = ImageOps.exif_transpose(im)
    if im.mode not in ("RGB", "L"):
        # PNG 透明底压成 JPG 会变黑，先铺白底
        bg = Image.new("RGB", im.size, (255, 255, 255))
        if im.mode in ("RGBA", "LA"):
            bg.paste(im, mask=im.split()[-1])
        else:
            bg.paste(im.convert("RGB"))
        im = bg
    elif im.mode == "L":
        im = im.convert("RGB")

    w, h = im.size
    long_side = max(w, h)
    if long_side > max_side:
        scale = max_side / float(long_side)
        im = im.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.LANCZOS)

    if out is None:
        out = src.with_name(f"{src.stem}-压缩.jpg")
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    im.save(out, "JPEG", quality=int(quality), optimize=True, progressive=True)
    return out, before, out.stat().st_size


def compress_to_size(src, out=None, target_kb: int = 200, max_side: int = 4000,
                     min_quality: int = 25) -> tuple[Path, int, int]:
    """按目标体积压（KB）。先降质量，质量到底还大就缩尺寸。返回 (输出, 原字节, 新字节)。"""
    from PIL import Image, ImageOps

    src = Path(src)
    if not src.exists():
        raise CompressError(f"找不到文件：{src}")
    before = src.stat().st_size
    target = max(10, int(target_kb)) * 1024
    try:
        im = Image.open(src)
    except Exception as exc:  # noqa: BLE001
        raise CompressError(f"这个图片打不开：{src.name}（{exc}）") from exc
    im = ImageOps.exif_transpose(im)
    if im.mode not in ("RGB", "L"):
        bg = Image.new("RGB", im.size, (255, 255, 255))
        if im.mode in ("RGBA", "LA"):
            bg.paste(im, mask=im.split()[-1])
        else:
            bg.paste(im.convert("RGB"))
        im = bg
    elif im.mode == "L":
        im = im.convert("RGB")

    if out is None:
        out = src.with_name(f"{src.stem}-压缩.jpg")
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)

    data = b""
    cur = im
    for _round in range(7):
        w, h = cur.size
        long_side = max(w, h)
        if long_side > max_side:
            sc = max_side / float(long_side)
            cur = cur.resize((max(1, int(w * sc)), max(1, int(h * sc))), Image.LANCZOS)
        q = 88
        while q >= min_quality:
            buf = io.BytesIO()
            cur.save(buf, "JPEG", quality=q, optimize=True, progressive=True)
            data = buf.getvalue()
            if len(data) <= target:
                break
            q -= 6
        if len(data) <= target:
            break
        nw, nh = int(cur.size[0] * 0.82), int(cur.size[1] * 0.82)
        if min(nw, nh) < 320:
            break
        cur = cur.resize((max(1, nw), max(1, nh)), Image.LANCZOS)

    out.write_bytes(data)
    return out, before, out.stat().st_size


def compress_many(paths, out_dir=None, max_side: int = 1600, quality: int = 75,
                  target_kb: int = 0, progress=None) -> list[dict]:
    """批量压。out_dir 为空时压在原文件旁边。target_kb>0 时改为按体积压。"""
    out = []
    total = len(paths)
    for i, p in enumerate(paths, 1):
        p = Path(p)
        dst = None
        if out_dir:
            dst = Path(out_dir) / f"{p.stem}-压缩.jpg"
        try:
            if target_kb:
                o, before, after = compress_to_size(p, dst, target_kb=target_kb,
                                                    max_side=max(3200, max_side))
            else:
                o, before, after = compress(p, dst, max_side=max_side, quality=quality)
            out.append({"src": str(p), "out": str(o), "before": before, "after": after,
                        "ok": True, "msg": ""})
        except Exception as exc:  # noqa: BLE001
            out.append({"src": str(p), "out": "", "before": 0, "after": 0,
                        "ok": False, "msg": str(exc)})
        if progress:
            try:
                progress(i, total)
            except Exception:  # noqa: BLE001
                pass
    return out


def human_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024.0
    return f"{n:.1f} GB"
