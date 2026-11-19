#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""列出若干 ELF 文件的动态依赖，并在指定目录里递归解析出真实库文件路径。

为什么需要它：PyInstaller 在 qemu+proot 里跑 `ldd` 会报
"you do not have read permission"，于是 tkinter 的 .so 能进包、它依赖的
libtk8.6/libtcl8.6/libX11 却一个都收不进来，打出来的包一启动就崩。

这个脚本自己解析 ELF 的 .dynamic 段（DT_NEEDED），再用“按目录查找”
的方式递归展开，不依赖 ldd / lddtree。用法：

    python3 elfdeps.py /usr/bin/python3.7 /path/to/_tkinter.so \
        --dirs /usr/lib/aarch64-linux-gnu /lib/aarch64-linux-gnu

输出：每行一个“真实文件绝对路径”（已解软链）。
"""
from __future__ import annotations

import os
import struct
import sys

DT_NEEDED = 1
DT_STRTAB = 5
DT_STRSZ = 10
PT_LOAD = 1
PT_DYNAMIC = 2


class ElfError(Exception):
    pass


def _parse64(data, off):
    """返回 (needed_list, ). 只处理 ELF64 小端（aarch64/x86_64）。"""
    if data[:4] != b"\x7fELF":
        raise ElfError("not ELF")
    ei_class = data[4]
    ei_data = data[5]
    if ei_class != 2 or ei_data != 1:
        raise ElfError("only ELF64 little-endian supported")

    e_phoff = struct.unpack_from("<Q", data, 0x20)[0]
    e_phentsize = struct.unpack_from("<H", data, 0x36)[0]
    e_phnum = struct.unpack_from("<H", data, 0x38)[0]

    loads = []
    dyn_off = dyn_size = None
    for i in range(e_phnum):
        base = e_phoff + i * e_phentsize
        p_type = struct.unpack_from("<I", data, base)[0]
        p_offset = struct.unpack_from("<Q", data, base + 0x08)[0]
        p_vaddr = struct.unpack_from("<Q", data, base + 0x10)[0]
        p_filesz = struct.unpack_from("<Q", data, base + 0x20)[0]
        if p_type == PT_LOAD:
            loads.append((p_vaddr, p_offset, p_filesz))
        elif p_type == PT_DYNAMIC:
            dyn_off, dyn_size = p_offset, p_filesz

    if dyn_off is None:
        return []

    entries = []
    n = dyn_size // 16
    for i in range(n):
        tag, val = struct.unpack_from("<QQ", data, dyn_off + i * 16)
        entries.append((tag, val))
        if tag == 0:
            break

    strtab_vaddr = None
    strsz = 0
    needed_offsets = []
    for tag, val in entries:
        if tag == DT_NEEDED:
            needed_offsets.append(val)
        elif tag == DT_STRTAB:
            strtab_vaddr = val
        elif tag == DT_STRSZ:
            strsz = val

    if strtab_vaddr is None:
        return []

    # vaddr -> file offset
    strtab_off = None
    for vaddr, offset, filesz in loads:
        if vaddr <= strtab_vaddr < vaddr + filesz:
            strtab_off = offset + (strtab_vaddr - vaddr)
            break
    if strtab_off is None:
        return []

    out = []
    for off in needed_offsets:
        start = strtab_off + off
        end = data.find(b"\x00", start)
        if end < 0:
            continue
        out.append(data[start:end].decode("utf-8", "replace"))
    return out


def needed(path):
    with open(path, "rb") as f:
        data = f.read()
    return _parse64(data, 0)


def resolve(name, dirs):
    for d in dirs:
        p = os.path.join(d, name)
        if os.path.exists(p):
            return os.path.realpath(p)
    return None


def walk(roots, dirs, skip=()):
    """从 roots 出发递归收集所有依赖的真实路径。"""
    found = {}
    queue = []
    for r in roots:
        real = os.path.realpath(r)
        try:
            queue.extend(needed(real))
        except Exception as exc:  # noqa: BLE001
            print("WARN cannot read %s: %s" % (real, exc), file=sys.stderr)
        found[os.path.basename(real)] = real
    while queue:
        name = queue.pop()
        if name in found:
            continue
        path = resolve(name, dirs)
        if not path:
            print("MISS %s" % name, file=sys.stderr)
            found[name] = None
            continue
        if any(name.startswith(s) for s in skip):
            found[name] = None
            continue
        found[name] = path
        try:
            queue.extend(needed(path))
        except Exception as exc:  # noqa: BLE001
            print("WARN cannot read %s: %s" % (path, exc), file=sys.stderr)
    return [p for p in found.values() if p]


def main(argv):
    roots, dirs = [], []
    i = 0
    while i < len(argv):
        a = argv[i]
        if a == "--dirs":
            i += 1
            while i < len(argv) and not argv[i].startswith("--"):
                dirs.append(argv[i])
                i += 1
            continue
        if a.startswith("--"):
            i += 1
            continue
        roots.append(a)
        i += 1
    if not dirs:
        dirs = ["/usr/lib/aarch64-linux-gnu", "/lib/aarch64-linux-gnu",
                "/usr/lib", "/lib", "/usr/lib64", "/lib64", "/usr/local/lib"]
    # 这些是 glibc / 基础库，目标系统一定有，不必打进包
    skip = ("libc.so", "libm.so", "libdl.so", "libpthread.so", "librt.so",
            "ld-linux", "libresolv.so", "libutil.so", "libnsl.so",
            "libcrypt.so", "libanl.so")
    for p in walk(roots, dirs, skip):
        print(p)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
