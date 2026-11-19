# -*- coding: utf-8 -*-
"""规则集：YAML 加载 + 合并 + 查找。可加载任意路径（参照件推导出的 derived.yaml）。

查找顺序：显式路径 → 用户规则集目录（可写，学出来的规则存这里）
→ 内置 rules/（随程序发布，只读）。用户目录同名会覆盖内置的。
"""
from __future__ import annotations

from functools import lru_cache
from pathlib import Path

import yaml

from . import sysenv

RULES_DIR = sysenv.resource_dir() / "rules"      # 内置，只读
USER_DIR = sysenv.user_rules_dir()               # 用户自建，可写


@lru_cache(maxsize=16)
def _load_path(p: str) -> dict:
    with open(p, "r", encoding="utf-8") as fh:
        data = yaml.safe_load(fh) or {}
    if not isinstance(data, dict):
        raise ValueError(f"规则集格式错误（顶层应为映射）：{p}")
    data.setdefault("roles", {})
    data["_path"] = p
    data["_key"] = Path(p).stem
    return data


def _dirs() -> list[Path]:
    out = []
    for d in (USER_DIR, RULES_DIR):
        if d and d.is_dir() and d not in out:
            out.append(d)
    return out


def find(name: str) -> Path | None:
    """按名字找规则文件，用户目录优先。"""
    stem = Path(name).stem if Path(name).suffix in (".yaml", ".yml") else name
    for d in _dirs():
        p = d / f"{stem}.yaml"
        if p.exists():
            return p
        p2 = d / f"{stem}.yml"
        if p2.exists():
            return p2
    return None


def load(name: str = "gb9704") -> dict:
    """name 可以是规则名（gb9704）或 .yaml 路径。"""
    cand = Path(name)
    if cand.suffix in (".yaml", ".yml") and cand.exists():
        return dict(_load_path(str(cand.resolve())))
    p = find(name)
    if p is None:
        raise FileNotFoundError(f"规则集不存在：{name}（可用：{'、'.join(available())}）")
    return dict(_load_path(str(p.resolve())))


def save_target(name: str) -> Path:
    """新建/覆盖规则集时应写到哪里：总是用户目录。"""
    stem = Path(name).stem
    return USER_DIR / f"{stem}.yaml"


def merge(base: dict, override: dict | None) -> dict:
    """项目级覆盖：只做一层深合并，够用且可预测。"""
    if not override:
        return base
    out = {k: (dict(v) if isinstance(v, dict) else v) for k, v in base.items()}
    for k, v in override.items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k].update(v)
        else:
            out[k] = v
    return out


def roles(rules: dict) -> dict:
    return rules.get("roles", {})


def get(rules: dict, role: str) -> dict:
    return roles(rules).get(role, {})


def available() -> list[str]:
    names: set[str] = set()
    for d in _dirs():
        names.update(p.stem for p in d.glob("*.yaml"))
        names.update(p.stem for p in d.glob("*.yml"))
    return sorted(names)
