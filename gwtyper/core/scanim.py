# -*- coding: utf-8 -*-
"""扫描件/拍照件整理：自动摆正（纠偏）+ 去掉黑边。

原则：只做「摆正」和「裁掉边框」，不增强对比度、不磨皮——
公文要的是能看清、不变形，过度处理反而让字发虚。
"""
from __future__ import annotations

from pathlib import Path

IMG_EXT = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}


class ScanError(RuntimeError):
    pass


def fix(src: str | Path, out: str | Path, deskew: bool = True,
        trim: bool = True, max_angle: float = 15.0) -> dict:
    """摆正 + 去黑边，另存为新文件（不动原件）。"""
    import cv2
    import numpy as np

    src, out = Path(src), Path(out)
    if src.suffix.lower() not in IMG_EXT:
        raise ScanError(f"这个格式处理不了：{src.suffix}")
    data = np.fromfile(str(src), dtype=np.uint8)
    img = cv2.imdecode(data, cv2.IMREAD_COLOR)
    if img is None:
        raise ScanError(f"打不开这张图：{src.name}")

    angle = 0.0
    if deskew:
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        thr = cv2.threshold(gray, 0, 255,
                            cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
        pts = cv2.findNonZero(thr)
        if pts is not None and len(pts) >= 200:
            rect = cv2.minAreaRect(pts)
            a = rect[-1]
            a = -(90 + a) if a < -45 else -a
            if 0.3 < abs(a) <= max_angle:
                angle = a
                h, w = img.shape[:2]
                m = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0)
                cos, sin = abs(m[0, 0]), abs(m[0, 1])
                nw, nh = int(h * sin + w * cos), int(h * cos + w * sin)
                m[0, 2] += nw / 2 - w / 2
                m[1, 2] += nh / 2 - h / 2
                img = cv2.warpAffine(img, m, (nw, nh),
                                     flags=cv2.INTER_CUBIC,
                                     borderMode=cv2.BORDER_CONSTANT,
                                     borderValue=(255, 255, 255))

    cropped = False
    if trim:
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        ink = cv2.threshold(gray, 235, 255, cv2.THRESH_BINARY_INV)[1]
        pts = cv2.findNonZero(ink)
        if pts is not None:
            x, y, w, h = cv2.boundingRect(pts)
            pad = 14
            x0, y0 = max(0, x - pad), max(0, y - pad)
            x1 = min(img.shape[1], x + w + pad)
            y1 = min(img.shape[0], y + h + pad)
            if (x1 - x0) * (y1 - y0) < img.shape[0] * img.shape[1] * 0.985:
                img = img[y0:y1, x0:x1]
                cropped = True

    out.parent.mkdir(parents=True, exist_ok=True)
    ext = out.suffix.lower() or ".png"
    ok, buf = cv2.imencode(ext if ext in IMG_EXT else ".png", img)
    if not ok:
        raise ScanError("保存失败（图片编码出错）")
    buf.tofile(str(out))
    return {"out": str(out), "angle": round(angle, 2), "cropped": cropped,
            "size": (img.shape[1], img.shape[0])}
