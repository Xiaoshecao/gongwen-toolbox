#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""修正 arm64 产物里「内容对但文件名还带 x86_64 标签」的扩展模块。

为什么需要：PyInstaller 在 x86 上分析时，把扩展模块的**文件名**原样抄了过来
（`onnxruntime_pybind11_state.cpython-312-x86_64-linux-gnu.so`），
交叉构建只替换了文件内容（ELF 头变成 aarch64），名字没改。
而 Python 的 import 机制是**按 filename tag 匹配**的：
ARM 机器上它只会找 `*aarch64-linux-gnu.so`，名字不对就 ImportError。

用法（构建后跑）：
    python3 fix_ext_tags.py <onedir 目录>
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

# 文件名里可能出现、需要替换掉的平台 tag
TAGS = [
    ("x86_64-linux-gnu", "aarch64-linux-gnu"),
    ("i386-linux-gnu", "aarch64-linux-gnu"),
    ("x86_64-linux-musl", "aarch64-linux-musl"),
]


def elf_machine(path: Path) -> int | None:
    try:
        with open(path, "rb") as f:
            h = f.read(20)
        if len(h) < 20 or h[:4] != b"\x7fELF":
            return None
        return h[18]
    except OSError:
        return None


def main(root: Path) -> int:
    if not root.is_dir():
        print(f"目录不存在：{root}")
        return 2
    renamed = 0
    for p in list(root.rglob("*.so")) + list(root.rglob("*.pyd")):
        if not p.is_file():
            continue
        name = p.name
        new = name
        for old_tag, new_tag in TAGS:
            if old_tag in name:
                new = name.replace(old_tag, new_tag)
                break
        else:
            # 没有 -linux-gnu 这种 tag，但是 .so 且名字里带 x86_64
            if "x86_64" in name:
                new = name.replace("x86_64", "aarch64")
        if new == name:
            continue
        # 只改「内容确实是 arm64」的，避免把真 x86 文件改坏名字
        machine = elf_machine(p)
        if machine != 0xB7:
            print(f"  跳过（不是 arm64，machine={machine}）：{name}")
            continue
        target = p.with_name(new)
        if target.exists():
            # 已经有人建了正确的，删掉这个错名的
            try:
                p.unlink()
                print(f"  删除重复 {name}（已有 {new}）")
                renamed += 1
            except OSError:
                pass
            continue
        try:
            p.rename(target)
            print(f"  改名 {name} -> {new}")
            renamed += 1
        except OSError as exc:
            print(f"  改名失败 {name}: {exc}")
    print(f"共修正 {renamed} 个文件名")
    return 0


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(Path(sys.argv[1]).resolve()))
