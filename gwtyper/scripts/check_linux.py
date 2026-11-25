# -*- coding: utf-8 -*-
"""把 Linux 代码路径跑一遍，并在能真转文件时做一次真实转换。

用法（在 `D:\\公文处理` 下）：
    python gwtyper/scripts/check_linux.py           # 任一台机器，验逻辑
    python3 gwtyper/scripts/check_linux.py          # 在真 Linux 上，会多跑真实转换

前一半：设好 HOME / XDG_DATA_HOME / XDG_CACHE_HOME，把 sys.platform 改成 'linux'，
再导入各模块，检查路径落点、字体枚举兜底、规则集双目录、无引擎时的报错文案。
后一半（仅当真 Linux + 装了 LibreOffice/字体时）：拿工作区里的真实公文转 PDF，
验证 soffice 通道端到端可用。
"""
import os
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]        # D:\公文处理
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

TMP = Path(tempfile.mkdtemp(prefix="gwlin-"))
os.environ["HOME"] = str(TMP / "home")
os.environ["XDG_DATA_HOME"] = str(TMP / "xdg-data")
os.environ["XDG_CACHE_HOME"] = str(TMP / "xdg-cache")
(TMP / "home").mkdir(parents=True, exist_ok=True)

REAL_LINUX = sys.platform.startswith("linux")
if not REAL_LINUX:
    sys.platform = "linux"      # 让所有分支按 Linux 走
for _m in [k for k in list(sys.modules) if k.startswith("gwtyper")]:
    del sys.modules[_m]

from gwtyper.core import doc_source, fonts, office, preview, rules, sysenv  # noqa: E402

ok = True


def chk(name, cond, extra=""):
    global ok
    if not cond:
        ok = False
    print(f"{'[OK]' if cond else '[!!]'} {name} {extra}")


print(f"== sysenv (linux 分支, 真机={REAL_LINUX}) ==")
print("  IS_WIN/IS_LINUX:", sysenv.IS_WIN, sysenv.IS_LINUX, "| note:", sysenv.platform_note())
print("  data_dir:", sysenv.data_dir())
print("  cache_dir:", sysenv.cache_dir("preview"))
print("  user_rules:", sysenv.user_rules_dir())
chk("platform_note 是 Linux", sysenv.platform_note() == "Linux")
chk("data_dir 走 XDG_DATA_HOME", str(sysenv.data_dir()).startswith(str(TMP / "xdg-data")))
chk("cache_dir 走 XDG_CACHE_HOME", "xdg-cache" in str(sysenv.cache_dir("preview")))
chk("data_dir 已建", sysenv.data_dir().is_dir())

print("== fonts (linux) ==")
inst = fonts.installed()
print("  枚举到:", len(inst), "| 目录建议:", [str(p) for p in fonts.font_dir_candidates()])
chk("字体枚举不抛错", isinstance(inst, set))
chk("字体目录建议含 /usr/share/fonts",
    any("share" in p.parts and "fonts" in p.parts for p in fonts.font_dir_candidates()),
    f"({[str(p) for p in fonts.font_dir_candidates()]})")
if REAL_LINUX and inst:
    has_cjk = any(("Noto" in f and "CJK" in f) or "WenQuanYi" in f for f in inst)
    chk("有中文字体（否则中文会变方块）", has_cjk, f"(共 {len(inst)} 个族)")
else:
    chk("无 fc-list 时返回空集不误报缺字", fonts.missing(["方正小标宋简体"]) == [] or bool(inst))

print("== office (linux) ==")
eng = office.engine_name()
print("  engine:", repr(eng), "| hint:", office.engine_hint(),
      "| wps:", office.wps_linux(), "| soffice:", office.soffice())
chk("Linux 下不走 COM", not office.com_available())
chk("Linux 下的引擎只可能是 wps / libreoffice / 空", eng in ("wps", "libreoffice", ""))
if REAL_LINUX:
    print("  WPS 优先，其次 LibreOffice；都没有则为空 —— 都算通过")
    chk("engine_name 与引擎存在性一致",
        bool(eng) == bool(office.wps_linux() or office.soffice()))

if not eng:
    try:
        office.word_to_pdf(ROOT / "不存在.docx", TMP / "x.pdf")
        chk("无引擎时抛 OfficeError", False, "没抛异常")
    except office.OfficeError as exc:
        chk("无引擎时抛 OfficeError", True, f"({str(exc)[:40]}…)")
    except Exception as exc:  # noqa: BLE001
        chk("无引擎时抛 OfficeError", False, f"抛了 {type(exc).__name__}: {exc}")
    try:
        preview.pdf_of(TMP / "nope.docx")
        chk("preview 抛 PreviewError", False, "没抛")
    except preview.PreviewError as exc:
        chk("preview 抛 PreviewError", True, f"({str(exc)[:24]}…)")
    except Exception as exc:  # noqa: BLE001
        chk("preview 抛 PreviewError", False, f"抛了 {type(exc).__name__}: {exc}")
    fake = TMP / "old.doc"
    fake.write_bytes(b"not a real doc")
    try:
        doc_source.to_docx(fake)
        chk("doc_source 无引擎时报错", False, "没抛")
    except doc_source.DocSourceError as exc:
        chk("doc_source 无引擎时报错", True, f"({str(exc)[:20]}…)")
    except Exception as exc:  # noqa: BLE001
        chk("doc_source 无引擎时报错", False, f"抛了 {type(exc).__name__}: {exc}")

print("== rules 双目录 (linux) ==")
print("  内置:", rules.RULES_DIR)
print("  用户:", rules.USER_DIR)
print("  可用:", rules.available())
chk("内置规则仍可读", rules.find("gb9704") is not None)
chk("save_target 指向用户目录", rules.save_target("t").parent == rules.USER_DIR)
chk("gb9704 内容正常", rules.load("gb9704")["roles"]["body"]["font"] == "仿宋_GB2312")

# ---------------------------------------------------------------- 真机端到端
if REAL_LINUX and eng:
    print("== 端到端：真实 docx → PDF（soffice 通道） ==")
    sample = ROOT / "发城镇重阳节前夕走访慰问老年人.docx"
    if not sample.exists():
        print("  跳过：样本不存在")
    else:
        pdf = TMP / "out.pdf"
        try:
            office.word_to_pdf(sample, pdf)
            size = pdf.stat().st_size if pdf.exists() else 0
            chk("转出了 PDF", size > 10000, f"({size} 字节)")
            import fitz
            doc = fitz.open(str(pdf))
            chk("PDF 页数 >= 1", doc.page_count >= 1, f"({doc.page_count} 页)")
            txt = "".join(doc[i].get_text() for i in range(doc.page_count))
            chk("PDF 里有正常中文（不是方块/空）", "发城" in txt or "老人" in txt, f"({len(txt)} 字)")
            doc.close()
        except Exception as exc:  # noqa: BLE001
            chk("端到端转换", False, f"{type(exc).__name__}: {exc}")

    print("== 端到端：体检 + 排版（不依赖 Office 的部分） ==")
    sample = ROOT / "发城镇重阳节前夕走访慰问老年人.docx"
    if sample.exists():
        from gwtyper.core import apply as applymod
        from gwtyper.core import check as checkmod
        from gwtyper.core import parse_docx
        doc = parse_docx.parse(sample)
        findings = checkmod.check(doc, rules.load("news"), text=True)
        chk("体检能跑", isinstance(findings, list), f"({len(findings)} 条)")
        out = TMP / "fmt.docx"
        applymod.apply(sample, rules.load("news"), out=out)
        doc2 = parse_docx.parse(out)
        chk("排版产出可回读", doc2.page.left_cm > 0, f"({len(doc2.paras)} 段)")
        findings2 = checkmod.check(doc2, rules.load("news"), text=True)
        chk("排版后体检干净", len(findings2) == 0, f"({len(findings2)} 条)")

print("\n结果:", "全部通过" if ok else "有问题")
sys.exit(0 if ok else 1)
