# -*- coding: utf-8 -*-
"""格式登记表：哪些源格式能转成哪些目标格式。"""
from __future__ import annotations

from pathlib import Path

WORD_EXT = {".doc", ".docx", ".docm", ".rtf", ".txt", ".wps", ".dot", ".dotx"}
EXCEL_EXT = {".xls", ".xlsx", ".xlsm", ".csv", ".et"}
PPT_EXT = {".ppt", ".pptx", ".pps", ".ppsx", ".dps"}
IMAGE_EXT = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp", ".gif"}
PDF_EXT = {".pdf"}
OFFICE_EXT = WORD_EXT | EXCEL_EXT | PPT_EXT

# 目标格式 -> 允许的源格式
TARGETS: dict[str, dict] = {
    "PDF": {
        "ext": ".pdf",
        "sources": WORD_EXT | EXCEL_EXT | PPT_EXT | IMAGE_EXT | PDF_EXT,
        "note": "文档、表格、演示、图片都能合成 PDF",
    },
    "DOCX": {
        "ext": ".docx",
        "sources": PDF_EXT | {".doc", ".rtf", ".wps", ".txt", ".docm"},
        "note": "PDF 转 Word，也可把 doc/rtf 升级为 docx",
    },
    "XLSX": {
        "ext": ".xlsx",
        "sources": PDF_EXT | {".xls", ".csv", ".et"},
        "note": "PDF 表格提取为 Excel，或 xls/csv 升级为 xlsx",
    },
    "PPTX": {
        "ext": ".pptx",
        "sources": PDF_EXT | {".ppt", ".pps", ".dps"},
        "note": "PDF 按页转为幻灯片，或 ppt 升级为 pptx",
    },
    "TXT": {
        "ext": ".txt",
        "sources": PDF_EXT | WORD_EXT | PPT_EXT | {".xlsx"},
        "note": "提取纯文字，便于检索和复制",
    },
    "CSV": {
        "ext": ".csv",
        "sources": {".xlsx", ".xls", ".et"},
        "note": "表格首个工作表导出为 CSV",
    },
    "PNG": {
        "ext": ".png",
        "sources": PDF_EXT | PPT_EXT,
        "note": "PDF/PPT 每页导出为图片",
    },
    "JPG": {
        "ext": ".jpg",
        "sources": PDF_EXT | PPT_EXT,
        "note": "PDF/PPT 每页导出为图片",
    },
}

# 目标是 PDF 时，源格式如果是 PDF 就没意义
TARGET_ORDER = ["PDF", "DOCX", "XLSX", "PPTX", "TXT", "CSV", "PNG", "JPG"]


def suffix(path: str | Path) -> str:
    return Path(path).suffix.lower()


def kind_of(path: str | Path) -> str:
    """给文件一个中文类别名，用于列表展示。"""
    s = suffix(path)
    if s in PDF_EXT:
        return "PDF"
    if s in {".doc", ".docx", ".docm", ".wps", ".rtf", ".dot", ".dotx"}:
        return "Word"
    if s in {".xls", ".xlsx", ".xlsm", ".csv", ".et"}:
        return "Excel"
    if s in {".ppt", ".pptx", ".pps", ".ppsx", ".dps"}:
        return "PPT"
    if s in IMAGE_EXT:
        return "图片"
    if s == ".txt":
        return "文本"
    return "其他"


def targets_for(paths: list[str]) -> list[str]:
    """根据选中的文件，算出可用的目标格式。"""
    if not paths:
        return TARGET_ORDER[:]
    allowed = set(TARGET_ORDER)
    for p in paths:
        s = suffix(p)
        ok = {t for t in TARGET_ORDER if s in TARGETS[t]["sources"]}
        allowed &= ok
    return [t for t in TARGET_ORDER if t in allowed]


def supported(src: str | Path, target: str) -> bool:
    return suffix(src) in TARGETS[target]["sources"]


def out_name(src: str | Path, target: str, out_dir: str | Path) -> Path:
    """生成输出文件名，重名时自动加 (1)(2)。"""
    p = Path(src)
    ext = TARGETS[target]["ext"]
    base = p.stem
    dst = Path(out_dir) / f"{base}{ext}"
    n = 1
    while dst.exists():
        dst = Path(out_dir) / f"{base}({n}){ext}"
        n += 1
    return dst
