# -*- coding: utf-8 -*-
"""文件查重：同一份材料改个名字又报一遍，能查出来。

两级判断：
1. 字节完全相同（sha1）——最可靠；
2. docx/pdf/txt 去掉空白后正文完全相同——改了文件名、重新另存过也能认出来。
"""
from __future__ import annotations

import hashlib
from pathlib import Path

TEXT_LIKE = {".docx", ".pdf", ".txt", ".md", ".csv"}


def _sha1(path: Path, block: int = 1 << 20) -> str:
    h = hashlib.sha1()
    with open(path, "rb") as f:
        while True:
            b = f.read(block)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def _text_key(path: Path) -> str:
    """取正文（去空白）再做 sha1；取不到就返回空串。"""
    try:
        suf = path.suffix.lower()
        text = ""
        if suf == ".docx":
            from docx import Document
            d = Document(str(path))
            parts = [p.text for p in d.paragraphs]
            for t in d.tables:
                for row in t.rows:
                    for c in row.cells:
                        parts.append(c.text)
            text = "\n".join(parts)
        elif suf == ".pdf":
            from . import pdfmod
            fitz = pdfmod.load()
            with fitz.open(str(path)) as doc:
                text = "\n".join(page.get_text() for page in doc)
        elif suf in (".txt", ".md", ".csv"):
            raw = path.read_bytes()
            for enc in ("utf-8-sig", "utf-8", "gb18030", "utf-16"):
                try:
                    text = raw.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
        text = "".join(text.split())
        return hashlib.sha1(text.encode("utf-8")).hexdigest() if len(text) >= 30 else ""
    except Exception:  # noqa: BLE001
        return ""


def find(paths, progress=None) -> list[dict]:
    """返回重复组：[{"why": "字节相同"/"内容相同", "files": [..]}, …]"""
    files = [Path(p) for p in paths if Path(p).is_file()]
    by_hash: dict[str, list[str]] = {}
    by_text: dict[str, list[str]] = {}
    total = len(files)
    for i, p in enumerate(files, 1):
        try:
            by_hash.setdefault(_sha1(p), []).append(str(p))
            if p.suffix.lower() in TEXT_LIKE:
                k = _text_key(p)
                if k:
                    by_text.setdefault(k, []).append(str(p))
        except OSError:
            pass
        if progress:
            try:
                progress(i, total)
            except Exception:  # noqa: BLE001
                pass

    groups: list[dict] = []
    seen_sets: list[set[str]] = []
    for why, table in (("字节相同", by_hash), ("内容相同", by_text)):
        for _k, group in table.items():
            if len(group) < 2:
                continue
            s = set(group)
            if any(s <= old for old in seen_sets):
                continue  # 字节相同的必然内容相同，不重复报
            seen_sets.append(s)
            groups.append({"why": why, "files": sorted(group)})
    groups.sort(key=lambda g: (-len(g["files"]), g["files"][0]))
    return groups
