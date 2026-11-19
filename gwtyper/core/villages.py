# -*- coding: utf-8 -*-
"""村名单与「材料上交」台账。

名单存在 `%LOCALAPPDATA%\公文排版工具\villages.txt`（一行一个村名），
下次打开自动载回来；勾选状态存在同目录 `village_status.json`。
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from . import sysenv

_WS = re.compile(r"\s+")


def strip_ordinal(s: str) -> str:
    """去掉行首的序号，但保住「102团」这种名字。

    只有「数字 + 分隔符（. 、 , ， ) ））」才算序号；
    「102团」「3号站」这种数字跟字连在一起的是名字的一部分，不动。
    """
    return re.sub(r"^\s*\d{1,3}(?:\s*[.、,，)）]\s*|\s+)", "", str(s or ""))


def normalize(s: str) -> str:
    """去掉所有空白，便于「文件名里有没有这个村名」的比较。"""
    return _WS.sub("", str(s or ""))


def names_path() -> Path:
    return sysenv.data_dir() / "villages.txt"


def status_path() -> Path:
    return sysenv.data_dir() / "village_status.json"


def load_names() -> list[str]:
    p = names_path()
    if not p.is_file():
        return []
    try:
        raw = p.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeError):
        return []
    out, seen = [], set()
    for line in raw.replace("\r", "\n").split("\n"):
        s = strip_ordinal(line.strip())
        if not s or s.startswith("#"):
            continue
        k = normalize(s)
        if k and k not in seen:
            seen.add(k)
            out.append(s)
    return out


def save_names(names: list[str]) -> Path:
    p = names_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text("\n".join(n.strip() for n in names if n.strip()) + "\n",
                 encoding="utf-8")
    return p


def load_status() -> dict[str, dict]:
    p = status_path()
    if not p.is_file():
        return {}
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_status(status: dict[str, dict]) -> None:
    p = status_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(status, ensure_ascii=False, indent=1), encoding="utf-8")


_PUNCT = re.compile(r"[\s（）()【】\[\]/\\|·．.\-—_、,，]+")
_SUFFIXES = ("村民委员会", "村委员会", "村委会", "社区居民委员会", "社区居委会",
             "居委会", "人民政府", "镇政府", "办事处", "办公室",
             # 「村」「社区」也算后缀：名单写「示例村」，文件叫「示例.jpg」也要能对上
             "村", "社区")


def _key(s: str) -> str:
    """比较用的键：去空白、去常见标点。"""
    return _PUNCT.sub("", str(s or ""))


def _aliases(name: str) -> list[str]:
    """「东村村委会」也能被「东村总结.docx」命中；「财政所（东片）」去掉括号也能比。"""
    out = []
    base = _key(name)
    if base:
        out.append(base)
    for suf in _SUFFIXES:
        if not base.endswith(suf):
            continue
        stem = base[: -len(suf)]
        # 去掉后缀后至少要留 2 个字（「东村」只剩「东」，就不拿它去乱撞）
        need = 2
        if len(stem) >= need:
            out.append(stem)
    return out


def match_files(names: list[str], files: list[str]) -> dict[str, str | None]:
    """按文件名找村：返回 {村名: 命中的文件名 或 None}。

    规则：文件名去掉空白和扩展名后，包含这个名称（去空白）即算命中。
    名称不必以「村」结尾——社区、站所、科室、学校、公司都照这个规则比。
    多个文件命中同一个名称时取第一个；一个文件可命中多个（如「甲村、乙村总结」）。
    """
    pairs: dict[str, str | None] = {n: None for n in names}
    for f in files:
        stem = _key(Path(f).stem)
        for n in names:
            if pairs[n] is not None:
                continue
            if any(a and a in stem for a in _aliases(n)):
                pairs[n] = Path(f).name
    return pairs


# ---------------------------------------------------------------- 预设名单
def presets_dir() -> Path:
    return sysenv.data_dir() / "villages_presets"


def list_presets() -> list[str]:
    d = presets_dir()
    if not d.is_dir():
        return []
    return sorted(p.stem for p in d.glob("*.txt"))


def load_preset(name: str) -> list[str]:
    p = presets_dir() / f"{name}.txt"
    if not p.is_file():
        return []
    return _lines_to_names(p.read_text(encoding="utf-8-sig"))


def save_preset(name: str, names: list[str]) -> Path:
    d = presets_dir()
    d.mkdir(parents=True, exist_ok=True)
    p = d / f"{_safe(name)}.txt"
    p.write_text("\n".join(names) + "\n", encoding="utf-8")
    return p


def delete_preset(name: str) -> bool:
    p = presets_dir() / f"{name}.txt"
    if p.is_file():
        p.unlink()
        return True
    return False


def _safe(name: str) -> str:
    bad = '\\/:*?"<>|'
    s = "".join("_" if c in bad else c for c in str(name)).strip()
    return s or "未命名"


def _lines_to_names(text: str) -> list[str]:
    out, seen = [], set()
    for line in text.replace("\r", "\n").split("\n"):
        s = strip_ordinal(line.strip())
        if not s or s.startswith("#"):
            continue
        k = normalize(s)
        if k and k not in seen:
            seen.add(k)
            out.append(s)
    return out


# ---------------------------------------------------------------- Excel 解析
HEADER_WORDS = {
    "序号", "编号", "村名", "村", "村庄", "单位", "单位名称", "名称", "备注",
    "姓名", "联系人", "电话", "联系方式", "统计", "数量", "合计", "小计",
    "是否上交", "上交情况", "材料", "材料名称", "填报单位", "填报人",
}


def _looks_like_name(v) -> bool:
    """像不像一个单位名（村、社区、站所、科室、学校、公司都算）。

    名称**不必以「村」结尾**；「第一社区」「102团」「财政所（东片）」
    这类带数字、带括号的也要认，只挡掉纯数字/日期和整句的话。
    """
    if not isinstance(v, str):
        return False
    s = v.strip()
    if not (2 <= len(s) <= 30):
        return False
    if s in HEADER_WORDS:
        return False
    if not re.search(r"[\u4e00-\u9fff]", s):      # 至少要有一个汉字
        return False
    if re.fullmatch(r"[\d\s.\-—/_年月日]+", s):   # 纯数字、日期
        return False
    if any(ch in "。；，,：:！？!?、" for ch in s):   # 整句的话
        return False
    return True


def parse_excel_names(path: str | Path, sheet: str | None = None,
                      column: str | None = None) -> dict:
    """从 Excel 里把村名挑出来。

    返回 {"names": [...], "sheet": 用到的表名, "column": 用到的列号或"整表",
          "rows": 扫了多少行}
    规则：优先按列打分（哪一列像名字的格子最多就用哪一列），
    表头（序号/村名/单位…）、纯数字、带标点的一律跳过。
    """
    from openpyxl import load_workbook
    from openpyxl.utils import column_index_from_string

    path = Path(path)
    if path.suffix.lower() not in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        raise ValueError("只认 .xlsx / .xlsm；老 .xls 请先在 Excel 里另存为 .xlsx")
    wb = load_workbook(str(path), read_only=True, data_only=True)
    try:
        sheets = [wb[sheet]] if sheet and sheet in wb.sheetnames else wb.worksheets
        best = {"names": [], "sheet": "", "column": "", "rows": 0}
        for ws in sheets:
            rows = 0
            col_hits: dict[int, list[str]] = {}
            for row in ws.iter_rows(values_only=True):
                rows += 1
                if rows > 5000:
                    break
                for idx, v in enumerate(row, start=1):
                    if _looks_like_name(v):
                        col_hits.setdefault(idx, []).append(v.strip())
            if column:
                target = column_index_from_string(column.upper())
                picks = col_hits.get(target, [])
                cand = {"names": _dedupe(picks), "sheet": ws.title,
                        "column": column.upper(), "rows": rows}
            else:
                if not col_hits:
                    continue
                idx = max(col_hits, key=lambda k: len(col_hits[k]))
                cand = {"names": _dedupe(col_hits[idx]), "sheet": ws.title,
                        "column": _col_letter(idx), "rows": rows}
            if len(cand["names"]) > len(best["names"]):
                best = cand
        return best
    finally:
        try:
            wb.close()
        except Exception:  # noqa: BLE001
            pass


def _dedupe(items: list[str]) -> list[str]:
    out, seen = [], set()
    for s in items:
        k = normalize(s)
        if k and k not in seen:
            seen.add(k)
            out.append(s)
    return out


def _col_letter(idx: int) -> str:
    s = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s
