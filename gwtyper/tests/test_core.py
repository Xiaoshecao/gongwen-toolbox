# -*- coding: utf-8 -*-
"""最小回归测试：不依赖 pytest，直接 python -m gwtyper.tests.test_core 就能跑。

覆盖：解析+识别+体检（三份真实公文）、公文要素检查、文本校对新增规则、
报告导出（txt/docx）、转换格式表。改代码后先跑这个，别让老功能悄悄坏掉。
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from gwtyper.core import rules as rulesmod          # noqa: E402
from gwtyper.core.check import check                # noqa: E402
from gwtyper.core.classify import classify          # noqa: E402
from gwtyper.core.doc_check import check_elements   # noqa: E402
from gwtyper.core.model import Doc, Fmt, PageSpec, Para, Role  # noqa: E402
from gwtyper.core.parse_docx import parse           # noqa: E402
from gwtyper.core import report as reportmod        # noqa: E402
from gwtyper.core.text_check import check_text      # noqa: E402

SAMPLES = [
    # 前两份本身是合规稿；第三份是「有毛病的原件」，已知必改 4 处，
    # 排完版再体检才是 0（那一步在别处验证过）。
    ("发城镇重阳节前夕走访慰问老年人.docx", "news", 0),
    ("述职述廉报告-谭钊（最终版）.docx", "report", 0),
    ("发城镇新时代文明实践项目培育工作机制（试行）.docx", "gb9704", 4),
]

_fails: list[str] = []


def check_(name, cond, detail=""):
    if cond:
        print(f"  ✓ {name}")
    else:
        print(f"  ✗ {name} {detail}")
        _fails.append(name)


def _para(i, text, role):
    return Para(index=i, text=text, fmt=Fmt(), role=role)


def test_samples():
    print("[1] 三份黄金样本：体检应为 0 问题")
    for name, prof, want in SAMPLES:
        p = ROOT / name
        if not p.exists():
            print(f"  - 跳过（文件不在）：{name}")
            continue
        doc = parse(str(p))
        classify(doc.paras, 0.45)
        fs = check(doc, rulesmod.load(prof))
        errs = [f for f in fs if f.level == "error"]
        check_(f"{name} 必改 {len(errs)}（应为 {want}）", len(errs) == want,
               f"实际 {len(errs)}：{errs[:2]}")


def test_elements():
    print("[2] 公文要素检查")
    # 缺主送机关
    d = Doc(path="x", page=PageSpec(), paras=[
        _para(0, "关于开展某项工作的通知", Role.TITLE),
        _para(1, "为做好相关工作，现将有关事项通知如下。", Role.BODY),
        _para(2, "某某镇人民政府", Role.SIGNATURE),
        _para(3, "2026年11月18日", Role.DATE),
    ])
    fields = [f.field for f in check_elements(d)]
    check_("发现缺主送机关", "主送机关" in fields, str(fields))

    # 日期早于署名
    d2 = Doc(path="x", page=PageSpec(), paras=[
        _para(0, "通知", Role.TITLE),
        _para(1, "某某单位：", Role.BODY),
        _para(2, "正文内容写在这里。", Role.BODY),
        _para(3, "2026年11月18日", Role.DATE),
        _para(4, "某某镇人民政府", Role.SIGNATURE),
    ])
    fs = check_elements(d2)
    check_("日期在署名之前被报出", any(f.field == "落款顺序" for f in fs),
           str([f.field for f in fs]))

    # 成文日期用中文数字
    d3 = Doc(path="x", page=PageSpec(), paras=[
        _para(0, "某某单位：", Role.BODY),
        _para(1, "正文。", Role.BODY),
        _para(2, "某某镇人民政府", Role.SIGNATURE),
        _para(3, "二〇二六年十一月十八日", Role.DATE),
    ])
    check_("中文数字日期被报出",
          any(f.field == "成文日期写法" for f in check_elements(d3)))


def test_text():
    print("[3] 文本校对新增规则")
    d = Doc(path="x", page=PageSpec(), paras=[
        _para(0, "他说“要抓紧。", Role.BODY),
        _para(1, "这项工作二〇二六年完成。", Role.BODY),
        _para(2, "计划26年11月开工。", Role.BODY),
    ])
    fields = [f.field for f in check_text(d)]
    check_("引号不成对被报出", "引号" in fields)
    check_("中文数字年份被报出", fields.count("年份写法") >= 2, str(fields))


def test_report():
    print("[4] 报告导出")
    d = Doc(path="x", page=PageSpec(), paras=[
        _para(0, "标题", Role.TITLE),
        _para(1, "正文。", Role.BODY),
    ])
    with tempfile.TemporaryDirectory() as td:
        txt = Path(td) / "r.txt"
        reportmod.save(txt, d, [], [], "测试")
        check_("txt 报告有内容", txt.exists() and txt.stat().st_size > 0)
        docx = Path(td) / "r.docx"
        reportmod.save(docx, d, [], [], "测试")
        check_("docx 报告有内容", docx.exists() and docx.stat().st_size > 1000,
               f"{docx.stat().st_size if docx.exists() else 0}B")


def test_parallel_helper():
    print("[5] 并行助手")
    from gwtyper.app import _parallel

    seen = []
    _parallel([1, 2, 3, 4, 5], seen.append)
    check_("全部执行", sorted(seen) == [1, 2, 3, 4, 5], str(seen))


def main() -> int:
    test_samples()
    test_elements()
    test_text()
    test_report()
    test_parallel_helper()
    print()
    if _fails:
        print(f"失败 {len(_fails)} 项：{_fails}")
        return 1
    print("全部通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
