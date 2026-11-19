# -*- coding: utf-8 -*-
"""命令行：体检 / 排版 / 学参照件。

    python -m gwtyper check  文件.docx -p gb9704      # 只读体检
    python -m gwtyper format 文件.docx -p gb9704      # 排版另存副本
    python -m gwtyper learn  参照件.docx -o rules/x.yaml
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from . import classify, rules as rulesmod
from .check import check, check_features, summary
from .model import ROLE_CN, Role
from .parse_docx import parse

ROLE_ORDER = ["title", "subtitle", "h1", "h2", "h3", "lead", "body",
              "table_caption", "note", "attach", "signature", "date", "page", "other"]


def _files(args) -> list[Path]:
    from .doc_source import is_supported

    out: list[Path] = []
    for a in args:
        p = Path(a)
        if p.is_dir():
            out.extend(sorted(x for x in p.rglob("*") if x.is_file() and is_supported(x)))
        else:
            out.append(p)
    return [f for f in out if not f.name.startswith("~$")]


def report(path: Path, profile, threshold: float, show_all: bool, quiet_ok=False) -> int:
    doc = parse(path)
    rules = rulesmod.load(profile) if isinstance(profile, str) else profile
    auto, review = classify.classify(doc.paras, threshold)
    findings = check(doc, rules)
    findings.extend(check_features(doc, rules))
    s = summary(findings)

    if not (quiet_ok and s["total"] == 0):
        print(f"\n{'=' * 68}\n{doc.name}")
        print(f"规则集 {rules.get('name', profile)} ｜ 段落 {len(doc.paras)} ｜ 表格 {doc.n_tables}")
        print(f"页边距 上{doc.page.top_cm} 下{doc.page.bottom_cm} "
              f"左{doc.page.left_cm} 右{doc.page.right_cm} cm")
        print(f"中文字体：{'、'.join(sorted(doc.fonts_used)) or '（无）'}")

        if show_all:
            print("\n── 角色识别 ──")
            for p in doc.paras:
                if not p.text.strip():
                    continue
                flag = "  " if p.confidence >= threshold else "??"
                print(f"{flag} {p.index + 1:>3} {ROLE_CN[p.role.value]:<5} "
                      f"{p.confidence:.2f}  {p.text[:34]}")

        if review:
            print(f"\n待确认 {len(review)} 段（置信度 < {threshold}，未自动排版）：")
            for p in review:
                print(f"   - 第{p.index + 1}段 「{p.text[:30]}」 证据：{'，'.join(p.evidence[:4])}")

        print(f"\n── 体检 ── 问题 {s['total']}（必改 {s['error']}／建议 {s['warn']}）")
        by_role: dict[str, list] = {}
        for f in findings:
            by_role.setdefault(f.role, []).append(f)
        for role in ROLE_ORDER:
            items = by_role.pop(role, None)
            if not items:
                continue
            print(f"\n[{ROLE_CN.get(role, role)}] {len(items)} 处")
            for f in items[:12]:
                print("  " + f.as_row())
            if len(items) > 12:
                print(f"  …… 另有 {len(items) - 12} 处同类")
        for role, items in by_role.items():
            print(f"\n[{ROLE_CN.get(role, role)}] {len(items)} 处")
            for f in items:
                print("  " + f.as_row())
    return 0 if s["error"] == 0 else 1


def cmd_check(args) -> int:
    from . import report as reportmod

    files = _files(args.paths)
    if not files:
        print("没有找到 doc/docx")
        return 2
    rules = rulesmod.load(args.profile)
    rc = 0
    entries = []

    for f in files:
        try:
            doc = parse(f)
            auto, review = classify.classify(doc.paras, args.threshold)
            findings = check(doc, rules)
            entries.append((doc, findings, review))
            rc |= report(f, rules, args.threshold, args.all)
        except Exception as e:  # noqa: BLE001
            print(f"\n!! {f.name} 解析失败：{type(e).__name__}: {e}")
            rc = 2

    if args.out and entries:
        text = (reportmod.build_batch(entries, rules.get("name", ""))
                if len(entries) > 1 else
                reportmod.build(entries[0][0], entries[0][1], entries[0][2],
                                rules.get("name", "")))
        out = Path(args.out)
        if out.suffix.lower() not in (".txt", ".md"):
            name = "公文体检报告.txt" if len(entries) > 1 else \
                f"{Path(entries[0][0].name).stem}-体检报告.txt"
            out = out / name
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(text, encoding="utf-8-sig")
        print(f"\n报告已保存：{out}")
    return rc


def cmd_format(args) -> int:
    from .apply import apply

    files = _files(args.paths)
    if not files:
        print("没有找到 doc/docx")
        return 2
    rules = rulesmod.load(args.profile)
    rc = 0
    for f in files:
        out = Path(args.out) if (args.out and len(files) == 1) else None
        try:
            path, doc, review = apply(f, rules, out, args.threshold)
            print(f"\n已排版：{path}")
            if review:
                print(f"  跳过 {len(review)} 段待确认：")
                for p in review:
                    print(f"    - 第{p.index + 1}段 「{p.text[:28]}」")
            # 排版后自检
            rc |= report(path, args.profile, args.threshold, False, quiet_ok=True)
        except Exception as e:  # noqa: BLE001
            print(f"\n!! {f.name} 排版失败：{type(e).__name__}: {e}")
            import traceback
            traceback.print_exc()
            rc = 2
    return rc


def cmd_learn(args) -> int:
    from .learn import learn, save

    p = Path(args.ref)
    prof = learn(p, args.name)
    out = Path(args.out) if args.out else rulesmod.save_target(p.stem)
    saved = save(prof, out)
    print(f"已生成规则集：{saved}")
    print(f"  识别角色：{'、'.join(ROLE_CN.get(k, k) for k in prof['roles'])}")
    print("  请打开核对数值后再使用。")
    return 0


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser("gwtyper", description="公文格式体检与排版")
    sub = ap.add_subparsers(dest="cmd")

    names = "、".join(rulesmod.available())

    c = sub.add_parser("check", help="只读体检")
    c.add_argument("paths", nargs="+")
    c.add_argument("-p", "--profile", default="gb9704", help="规则集名或 yaml 路径")
    c.add_argument("-t", "--threshold", type=float, default=0.45)
    c.add_argument("-a", "--all", action="store_true", help="列出全部段落角色")
    c.add_argument("-o", "--out", help="把报告另存为 txt（目录或不带扩展名的路径）")

    f = sub.add_parser("format", help="排版并另存副本")
    f.add_argument("paths", nargs="+")
    f.add_argument("-p", "--profile", default="gb9704")
    f.add_argument("-o", "--out", help="输出路径（仅单文件时有效）")
    f.add_argument("-t", "--threshold", type=float, default=0.45)

    l = sub.add_parser("learn", help="从参照件反推规则集")
    l.add_argument("ref")
    l.add_argument("-o", "--out")
    l.add_argument("-n", "--name")

    p = sub.add_parser("pdf", help="转 PDF（可选纯 Python 引擎，不依赖办公软件）")
    p.add_argument("paths", nargs="+")
    p.add_argument("-o", "--out", default=".", help="输出目录")
    p.add_argument("-e", "--engine", default="auto",
                   choices=["pure", "external", "auto"],
                   help="pure=纯 Python / external=内置·系统办公引擎 / auto=先试纯 Python")
    p.add_argument("--portrait", action="store_true", help="统一竖版")
    p.add_argument("--landscape", action="store_true", help="统一横版")
    p.add_argument("--margin", type=float, default=0.0, help="四周留白（毫米）")

    d = sub.add_parser("doctor", help="自检：看各引擎是否可用")

    dl = sub.add_parser("double", help="做双层 PDF（图像层 + 可搜索文字层）")
    dl.add_argument("src", help="源文件：Word/WPS/Excel/PPT/PDF/图片")
    dl.add_argument("-o", "--out", help="输出 PDF 路径")
    dl.add_argument("--dpi", type=int, default=200, help="渲染精度，默认 200")

    args = ap.parse_args(argv)
    if args.cmd is None:
        ap.print_help()
        print(f"\n可用规则集：{names}")
        return 0
    return {"check": cmd_check, "format": cmd_format, "learn": cmd_learn,
            "pdf": cmd_pdf, "doctor": cmd_doctor,
            "double": cmd_double}[args.cmd](args)


def cmd_double(args) -> int:
    """命令行做双层 PDF。"""
    from . import ocr as ocrmod

    src = Path(args.src)
    out = Path(args.out) if args.out else src.with_name(f"{src.stem}-双层.pdf")
    print(f"源文件：{src.name}")
    print(f"方式  ：{ocrmod.describe_plan(src)}")
    try:
        st = ocrmod.make_double_layer(
            src, out, dpi=args.dpi,
            progress=lambda m: print(f"  {m}", flush=True))
    except Exception as exc:  # noqa: BLE001
        print(f"失败：{exc}")
        return 1
    print(f"完成：{out}")
    print(f"  {st.get('pages', 0)} 页｜文字页 {st.get('text_pages', 0)}｜"
          f"OCR 页 {st.get('ocr_pages', 0)}｜文字 {st.get('lines', 0)} 段")
    after = Path(st["out"])
    if after.exists():
        print(f"  体积 {after.stat().st_size/1024:.0f} KB")
    return 0


def cmd_doctor(args) -> int:
    """打印各 PDF 引擎的可用情况——排查「为什么转不了」用。"""
    from . import lo_bundle, office

    print("=" * 60)
    print("PDF 引擎自检")
    print("=" * 60)
    print(f"平台           : {sys.platform}")
    try:
        from . import sysenv

        print(f"资源目录       : {sysenv.resource_dir()}")
    except Exception:  # noqa: BLE001
        pass

    print(f"\n首选引擎       : {office.engine_name()!r}")
    print(f"说明           : {office.engine_hint()}")

    print("\n--- 内置 LibreOffice ---")
    try:
        st = lo_bundle.status()
        print(f"  包内位置     : {st['bundled'] or '（没有）'}")
        print(f"  可用         : {st['usable']}")
        print(f"  unshare      : {st['unshare']}（免 root 的私有挂载）")
        print(f"  可提权       : {st['can_elevate']}（unshare 不可用时用 root 挂载）")
        print(f"  已全局挂载   : {st['mounted']}")
    except Exception as exc:  # noqa: BLE001
        print(f"  检测失败：{exc}")

    print("\n--- 纯 Python 排版 ---")
    try:
        from . import pdfwrite

        print(f"  可用         : {pdfwrite.available()}")
        try:
            from . import fontfind

            for name in ("仿宋_GB2312", "楷体_GB2312", "黑体", "方正小标宋简体"):
                got = fontfind.resolve(name)
                print(f"  {name:14s}: {got or '（缺，会用兜底字体）'}")
        except Exception as exc:  # noqa: BLE001
            print(f"  字体检测失败：{exc}")
    except Exception as exc:  # noqa: BLE001
        print(f"  检测失败：{exc}")

    print("\n--- 图标 ---")
    try:
        from . import appicon

        paths = appicon.icon_paths()
        if not paths:
            print("  ！找不到任何图标文件")
        for p in paths:
            print(f"  {p}  {p.stat().st_size} 字节")
    except Exception as exc:  # noqa: BLE001
        print(f"  检测失败：{exc}")

    print("\n--- OCR（双层 PDF 用）---")
    try:
        from . import ocr as ocrmod

        ok = ocrmod.ocr_available()
        print(f"  可用         : {ok}")
        if ok:
            try:
                eng = ocrmod.engine()
                print(f"  引擎         : {type(eng).__name__}")
                print("  模型         : 随程序打包，无需另装")
            except Exception as exc:  # noqa: BLE001
                print(f"  引擎初始化失败：{exc}")
        else:
            print("  没装 rapidocr_onnxruntime，扫描件做不了双层 PDF")
            print("  （Word/WPS 文件仍可做，因为它不需要 OCR）")
    except Exception as exc:  # noqa: BLE001
        print(f"  检测失败：{exc}")

    print("\n--- 系统办公软件 ---")
    print(f"  WPS(Linux)   : {office.wps_linux() or '（没有）'}")
    print(f"  系统 soffice : {office.soffice() or '（没有）'}")
    print(f"  Windows COM  : {office.com_available()}")
    print("=" * 60)
    return 0


def cmd_pdf(args) -> int:
    """把文件转成 PDF。--engine pure 时完全不碰 WPS/Office/LibreOffice。"""
    from . import convert as cv

    opt = cv.Options(
        pdf_engine=args.engine,
        orientation="portrait" if args.portrait else
                    ("landscape" if args.landscape else "auto"),
        margin_mm=args.margin,
    )
    out_dir = args.out
    os.makedirs(out_dir, exist_ok=True)
    failed = 0
    for path in args.paths:
        try:
            made = cv.convert(path, "PDF", out_dir, opt)
            for m in made:
                size = os.path.getsize(m)
                print(f"  已生成 {m}（{size / 1024:.0f} KB）")
        except Exception as exc:  # noqa: BLE001
            failed += 1
            print(f"  失败 {path}：{str(exc).splitlines()[0]}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
