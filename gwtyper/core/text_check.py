# -*- coding: utf-8 -*-
"""文本层校对：标点、空格、序号连续性、常见错词。只读，产出 Finding。

原则：宁可少报也不误报。公文正文里合法的写法（如「1.5 万亩」「2026年」）
不能被当成问题，所以每条规则都限定上下文。
"""
from __future__ import annotations

import re

from .model import Doc, Finding, Para, Role

CN = r"\u4e00-\u9fff"

# 半角标点出现在中文之间：,;:?! 与 () 分开处理（括号可能是英文缩写）
RE_HALF_PUNCT = re.compile(rf"[{CN}][,;:?!][{CN}]")
RE_HALF_PAREN = re.compile(rf"[{CN}]\([^)]*[{CN}][^)]*\)[{CN}]?")
RE_FULLWIDTH_DIGIT = re.compile(r"[０-９]+")
RE_FULLWIDTH_ALPHA = re.compile(r"[Ａ-Ｚａ-ｚ]+")
RE_MULTI_SPACE = re.compile(rf"[{CN}][ \t\u3000]{{2,}}[{CN}]")
RE_SPACE_BEFORE_PUNCT = re.compile(rf"[ \t\u3000]+[，。、；：？！）】」』]")
RE_DUP_PUNCT = re.compile(r"([，。、；：？！])\1+")
RE_ELLIPSIS = re.compile(r"\.{3,}|。{3,}")
# 数字范围：GB/T 15835 要求用浪纹线「～」。半角 ~ 或一个连字符都不行。
RE_RANGE_TILDE = re.compile(r"(\d)\s*[~]\s*(\d)")
# 百分数前面不该有空格
RE_PCT_SPACE = re.compile(r"\d+\s+[%％]")
# 用半角连字符表示数字范围（排除日期 2026-11-18、电话、编号）
RE_RANGE_HYPHEN = re.compile(r"(?<![\d\-])(\d{1,3})\s*-\s*(\d{1,3})(?![\d\-月日年])")

# 「的/地/得」只查常见固定搭配，不做通用语法判断（通用判断误报太多）
DE_PHRASES = {
    "做的很好": "做得很好", "做的不错": "做得不错", "做的不好": "做得不好",
    "说的对": "说得对", "说的很对": "说得很对", "说的好": "说得好",
    "跑的很快": "跑得很快", "发展的很快": "发展得很快", "长的很": "长得很",
    "变的很": "变得很", "写的很好": "写得很好", "干的很好": "干得很好",
    "紧密的团结": "紧密地团结", "认真的学习": "认真地学习",
    "高兴的说": "高兴地说", "不断的提高": "不断地提高",
    "逐步的完善": "逐步地完善", "更好的服务": "更好地服务",
    "有效的推进": "有效地推进", "全面的落实": "全面地落实",
}

# 常见公文错词（左错右对），只查明确无歧义的
WRONG_WORDS = {
    "截止目前": "截至目前", "截止到": "截至", "截止2026": "截至2026",
    "帐目": "账目", "帐户": "账户", "记帐": "记账",
    "其它": "其他", "做为": "作为", "按装": "安装", "即使如此": "即便如此",
    "必需品的": "必需品", "供献": "贡献", "布署": "部署", "复盖": "覆盖",
    "既使": "即使", "以经": "已经", "因该": "应该", "必须要": "必须",
    "座谈会上上": "座谈会上", "的的": "的", "了了": "了", "是是": "是",
    "按排": "安排", "予会": "与会", "给与": "给予", "走头无路": "走投无路",
    "迫不急待": "迫不及待", "严竣": "严峻", "攻艰": "攻坚", "受益非浅": "受益匪浅",
    "一慨": "一概", "渡过难关": "度过难关", "闻名暇迩": "闻名遐迩",
}

# 需要成对出现的符号：名称 → (开, 闭)
PAIRS = {
    "引号": ("“", "”"),
    "单引号": ("‘", "’"),
    "书名号": ("《", "》"),
    "括号": ("（", "）"),
}

# 中文数字写年份（GB/T 9704 要求成文日期用阿拉伯数字）
RE_CN_YEAR = re.compile(r"[〇零○一二三四五六七八九]{2,4}\s*年")
# 两位数字年份（如「26年」），前面不能紧跟数字或年月
RE_SHORT_YEAR = re.compile(r"(?<![\d年])(\d{2})\s*年(?=[\d一二三四五六七八九十]{1,2}\s*月)")


# 序号族：(角色, 提取函数, 名称)
NUM_H1 = re.compile(r"^([一二三四五六七八九十]{1,3})、")
NUM_H2 = re.compile(r"^[（(]([一二三四五六七八九十]{1,3})[）)]")
NUM_H3 = re.compile(r"^(\d{1,2})[.．、]")
CN_NUM = {c: i for i, c in enumerate("一二三四五六七八九十", start=1)}

TEXT_ROLES = {Role.TITLE, Role.SUBTITLE, Role.H1, Role.H2, Role.H3,
              Role.BODY, Role.LEAD, Role.NOTE, Role.SIGNATURE}


def _cn2int(s: str) -> int | None:
    s = s.strip()
    if not s:
        return None
    if s == "十":
        return 10
    if s.startswith("十"):
        return 10 + CN_NUM.get(s[1], 0) if len(s) > 1 else 10
    if "十" in s:
        a, _, b = s.partition("十")
        return CN_NUM.get(a, 0) * 10 + (CN_NUM.get(b, 0) if b else 0)
    return CN_NUM.get(s)


def _check_para(p: Para) -> list[Finding]:
    out: list[Finding] = []
    t = p.text
    role = p.role.value

    def add(field, actual, expected, level="error", note=""):
        out.append(Finding(p.index, role, field, actual, expected, level=level, note=note))

    m = RE_HALF_PUNCT.search(t)
    if m:
        seg = t[max(0, m.start() - 4):m.end() + 4]
        add("标点", f"半角「{t[m.start() + 1]}」→ {seg}", "用全角标点", "warn",
            "中文语境下应用全角标点")

    m = RE_HALF_PAREN.search(t)
    if m:
        add("标点", m.group(0)[:16], "中文括号（", "warn", "中文内容用全角括号")

    m = RE_FULLWIDTH_DIGIT.search(t)
    if m:
        add("数字", f"全角数字「{m.group(0)}」", "半角数字", "warn")

    m = RE_FULLWIDTH_ALPHA.search(t)
    if m:
        add("字母", f"全角字母「{m.group(0)}」", "半角字母", "warn")

    m = RE_MULTI_SPACE.search(t)
    if m and p.role not in (Role.TITLE, Role.SUBTITLE):
        seg = m.group(0).replace(" ", "·")
        add("空格", f"多余空格「{seg}」", "删掉多余空格", "warn")

    m = RE_SPACE_BEFORE_PUNCT.search(t)
    if m:
        add("空格", f"标点前有空格「{m.group(0)}」", "去掉空格", "warn")

    m = RE_DUP_PUNCT.search(t)
    if m:
        add("标点", f"重复标点「{m.group(0)}」", "保留一个", "error")

    m = RE_ELLIPSIS.search(t)
    if m:
        add("标点", f"省略号写法「{m.group(0)}」", "……（六点）", "warn")

    m = RE_RANGE_TILDE.search(t)
    if m:
        add("数字用法", f"数字范围「{m.group(0)}」", "用浪纹线「～」，如 3～5（GB/T 15835）",
            "warn", "数值范围用浪纹线，不用半角波浪线")

    m = RE_RANGE_HYPHEN.search(t)
    if m:
        add("数字用法", f"数字范围「{m.group(0)}」", "用浪纹线「～」，如 3～5",
            "warn", "半角连字符不是数字范围号")

    m = RE_PCT_SPACE.search(t)
    if m:
        add("数字用法", f"百分号前有空格「{m.group(0)}」", "写成 50%，中间不留空", "warn")

    # —— 成对符号：数量不等就提示（成对缺失比多打更容易漏）
    for name, (op, cl) in PAIRS.items():
        n_op, n_cl = t.count(op), t.count(cl)
        if n_op != n_cl:
            add(name, f"开了 {n_op} 个、闭了 {n_cl} 个",
                f"{op}{cl} 成对出现", "warn", "引号/书名号/括号必须成对")

    # —— 年份写法
    m = RE_CN_YEAR.search(t)
    if m and p.role in (Role.BODY, Role.LEAD, Role.DATE, Role.SIGNATURE):
        add("年份写法", f"「{m.group(0)}」", "用阿拉伯数字，如 2026年", "warn",
            "公文中的年份用阿拉伯数字")
    m = RE_SHORT_YEAR.search(t)
    if m:
        add("年份写法", f"「{m.group(0)}」", f"补全世纪，如 20{m.group(1)}年", "warn",
            "两位年份容易误读")

    for bad, good in WRONG_WORDS.items():
        if bad in t:
            add("用词", bad, good, "warn")

    for bad, good in DE_PHRASES.items():
        if bad in t:
            add("的地得", bad, good, "warn", "结构助词用错")

    if p.role in (Role.BODY, Role.LEAD) \
            and t.strip() and t.strip()[-1] not in "。！？：；…”」』）":
        add("句末", f"「{t.strip()[-12:]}」", "补句号", "warn", "段落一般以句号收尾")

    return out


def _check_sequence(paras: list[Para]) -> list[Finding]:
    """序号连续性：一、二、… 与（一）（二）… 与 1. 2. …"""
    out: list[Finding] = []

    def scan(regex, label, cn=True):
        seen: list[tuple[int, int, Para]] = []   # (n, para index, para)
        for p in paras:
            m = regex.match(p.text.strip())
            if not m:
                continue
            n = _cn2int(m.group(1)) if cn else int(m.group(1))
            if n:
                seen.append((n, p.index, p))
        if len(seen) < 2:
            return
        for k in range(1, len(seen)):
            prev, cur = seen[k - 1][0], seen[k][0]
            if cur == prev + 1:
                continue
            if cur == 1:
                continue                        # 新的一组重头开始，正常
            if prev == 1 and cur > 2:
                out.append(Finding(seen[k][1], "h2" if label == "（一）" else "h1", "序号",
                                   f"{label}第 {cur} 项", f"应为 {label}2（中间漏了）", "warn",
                                   note="或确认是否新起一组"))
            elif cur > prev + 1:
                out.append(Finding(seen[k][1], "h2" if label == "（一）" else "h1", "序号",
                                   f"跳过 {prev + 1}～{cur - 1}", f"接续 {label}{prev + 1}", "warn"))
            elif cur <= prev:
                out.append(Finding(seen[k][1], "h2" if label == "（一）" else "h1", "序号",
                                   f"{label}{cur} 与前重复/回退", "按顺序递增", "warn"))

    scan(NUM_H1, "")
    scan(NUM_H2, "（一）")
    scan(NUM_H3, "", cn=False)
    return out


def check_text(doc: Doc) -> list[Finding]:
    out: list[Finding] = []
    texty = [p for p in doc.paras if p.text.strip() and p.role in TEXT_ROLES]
    for p in texty:
        out.extend(_check_para(p))
    out.extend(_check_sequence(texty))
    return out
