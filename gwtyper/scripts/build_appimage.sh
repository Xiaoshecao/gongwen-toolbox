#!/usr/bin/env bash
# 打 Linux AppImage（支持跨架构，两种打包模式）。
#
# 前置（Debian/Ubuntu 系，打包机上）：
#   sudo apt install -y python3 python3-tk python3-docx python3-yaml \
#                       python3-fitz python3-pil squashfs-tools
#   python3 -m pip install --break-system-packages pyinstaller   # 仅 full 模式需要
#
# 用法：
#   bash gwtyper/scripts/build_appimage.sh              # 按本机架构打，默认 full
#   bash gwtyper/scripts/build_appimage.sh aarch64      # 指定架构
#   MODE=thin bash gwtyper/scripts/build_appimage.sh    # 薄包（不打包 Python 运行时）
# 产物：gwtyper/dist/公文排版工具-<ver>-<arch>.AppImage
#
# 两种模式的区别：
#   full（默认，推荐分发用）：PyInstaller 把 Python 解释器、tcl/tk、python-docx、
#     PyMuPDF 全部打进包里，约 130MB。换一台干净机器（没装 python3-tk 也行）直接
#     双击就能跑。
#   thin：只打纯 Python 源码，约 1MB。要求目标机自己装 python3 + tkinter + 那几个库。
#     适合几十台同构的机器统一用 apt 铺一遍依赖。
#
# 跨架构：full 模式里的 Python 解释器和 .so 是**架构相关**的，所以 aarch64 包
#   必须在 aarch64 机器（或 qemu）上打，不能在你手上这台 x86 上交叉生成。
#   thin 模式才是真正的架构无关（纯 .py），换 runtime 就行。
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PKG="$(cd "$HERE/.." && pwd)"          # .../gwtyper
ROOT="$(cd "$PKG/.." && pwd)"          # 仓库根（含 gwtyper/）
PY_BIN="${PY_BIN:-python3}"
APP="公文工具箱"
MODE="${MODE:-full}"
# 已经有 PyInstaller 产物时，直接指过来就能拼包，不用再打一遍。
# 用在交叉打包上：arm64 的产物是在 qemu 里打出来的（见下），拼 AppImage 这步
# 只是 mksquashfs + cat，跟架构无关，回到 x86 上做又快又省事。
DIST_SRC="${DIST_SRC:-}"

# ---------------------------------------------------------------- 架构
ARCH="${1:-${ARCH:-}}"
if [ -z "$ARCH" ]; then
    case "$(uname -m)" in
        x86_64|amd64)  ARCH="x86_64" ;;
        aarch64|arm64) ARCH="aarch64" ;;
        armv7l|armhf)  ARCH="armhf" ;;
        *)             ARCH="$(uname -m)" ;;
    esac
fi
case "$ARCH" in
    amd64) ARCH="x86_64" ;;
    arm64) ARCH="aarch64" ;;
esac
HOST="$(uname -m)"
# 注意：set -e 下不能用 `[ 条件 ] && 赋值`，条件为假时整条返回 1 会让脚本静默退出
case "$HOST" in
    amd64|x86_64) HOST="x86_64" ;;
    arm64|aarch64) HOST="aarch64" ;;
esac
APPDIR="$PKG/build/AppDir-$ARCH"

command -v "$PY_BIN" >/dev/null || { echo "缺少 $PY_BIN"; exit 1; }

cd "$ROOT"

# 版本号从 app.py 的 APP_VER 抠出来，免得两处对不上（-m1 避免 sed 被 SIGPIPE 打断，
# set -o pipefail 下 `sed | head` 会直接退出）
VER="$(grep -m1 '^APP_VER = ' gwtyper/app.py | sed 's/.*= *"\([^"]*\)".*/\1/')"
[ -n "$VER" ] || VER="0.0"

echo "== 模式：$MODE ｜ 架构：$ARCH（本机 $HOST）｜ 版本：$VER =="

if [ "$MODE" = "full" ] && [ "$ARCH" != "$HOST" ] && [ -z "$DIST_SRC" ]; then
    cat >&2 <<EOF
错误：full 模式不能交叉打包。
  包里带的是本机架构的 Python 解释器和 .so，在 $HOST 上只能产出 $HOST 的包。
  要出 $ARCH 版，有两条路：
    1. 在 $ARCH 机器（或 proot + qemu-aarch64 环境）里先跑 PyInstaller，
       再用 DIST_SRC=/path/to/linux-dist/gwtyper bash $0 $ARCH 只拼包；
    2. 目标机装好 python3 + 那几个库，改用：MODE=thin bash $0 $ARCH
EOF
    exit 1
fi

echo "== 清理 =="
rm -rf "$APPDIR"
# 逐个建：一条 mkdir 带多个路径时，只要其中一个父目录正被清理就会整条失败
for d in "$APPDIR/usr/bin" "$APPDIR/usr/share/applications" \
         "$APPDIR/usr/share/icons/hicolor/256x256/apps" \
         "$APPDIR/usr/share/gwtyper" "$PKG/build"; do
    mkdir -p "$d" || { echo "建目录失败：$d"; exit 1; }
done

if [ "$MODE" = "full" ]; then
    # ------------------------------------------------------------ 自带运行时
    if [ -n "$DIST_SRC" ]; then
        echo "== 复用现成的 PyInstaller 产物：$DIST_SRC =="
        DIST="$DIST_SRC"
    else
        echo "== PyInstaller 打包（含 Python 运行时）=="
        "$PY_BIN" -m PyInstaller --version >/dev/null 2>&1 || {
            echo "需要 pyinstaller：$PY_BIN -m pip install --break-system-packages pyinstaller"
            exit 1
        }
        # 缺依赖时 PyInstaller 只在 hiddenimport 那里打一行 ERROR 就跳过，
        # 最后打出一个「能开、但一转 PDF 就崩」的残包。所以先自己查一遍。
        miss="$("$PY_BIN" - <<'PY' 2>/dev/null
mods = ["fontTools", "fontTools.subset", "openpyxl", "pptx", "docx",
        "yaml", "fitz", "PIL"]
bad = []
for m in mods:
    try:
        __import__(m)
    except Exception:
        bad.append(m)
print(" ".join(bad))
PY
)"
        if [ -n "$miss" ]; then
            cat >&2 <<EOF
缺少打包所需的 Python 模块：$miss
装一下再来（Debian/Ubuntu 上系统 python 要加 --break-system-packages）：
  $PY_BIN -m pip install --break-system-packages \\
      fonttools openpyxl python-pptx python-docx pyyaml pymupdf pillow
EOF
            exit 1
        fi
        rm -rf "$PKG/build/linux-dist" "$PKG/build/linux-work"
        "$PY_BIN" -m PyInstaller --noconfirm --clean \
            --distpath "$PKG/build/linux-dist" \
            --workpath "$PKG/build/linux-work" \
            "$PKG/build_linux.spec" > "$PKG/build/pyinstaller.log" 2>&1 || {
                echo "PyInstaller 构建失败，日志：$PKG/build/pyinstaller.log"
                tail -20 "$PKG/build/pyinstaller.log" >&2
                exit 1
            }
        # 包打完了再确认一次关键模块真的进去了
        for probe in fontTools openpyxl; do
            if [ ! -e "$PKG/build/linux-dist/gwtyper/_internal/$probe" ]; then
                echo "警告：$probe 没进包，纯 Python PDF 功能会不可用" >&2
                echo "      看 $PKG/build/pyinstaller.log 里的 hidden import 报错" >&2
            fi
        done
        DIST="$PKG/build/linux-dist/gwtyper"
    fi
    [ -x "$DIST/gwtyper" ] || { echo "PyInstaller 没产出 $DIST/gwtyper"; exit 1; }

    echo "== 复制运行时到 AppDir =="
    mkdir -p "$APPDIR/usr/lib/gwtyper-dist"
    cp -a "$DIST/." "$APPDIR/usr/lib/gwtyper-dist/"

    cat > "$APPDIR/usr/bin/gwtyper" <<'LAUNCH'
#!/bin/sh
HERE="$(dirname "$(readlink -f "$0")")/.."
exec "$HERE/lib/gwtyper-dist/gwtyper" "$@"
LAUNCH
    chmod +x "$APPDIR/usr/bin/gwtyper"

    # 包一份开源中文字体，保证干净机器上界面和文档至少不出方块
    echo "== 附带走一遍中文字体 =="
    FONTDIR="$APPDIR/usr/share/fonts/truetype/gwtyper"
    mkdir -p "$FONTDIR"
    for f in /usr/share/fonts/truetype/wqy/wqy-zenhei.ttc \
             /usr/share/fonts/truetype/wqy/wqy-microhei.ttc; do
        [ -f "$f" ] && cp "$f" "$FONTDIR/" && echo "  带了 $(basename "$f")"
    done
    cat > "$APPDIR/usr/share/fonts/gwtyper-fonts.conf" <<'FC'
<?xml version="1.0"?>
<!DOCTYPE fontconfig SYSTEM "fonts.dtd">
<fontconfig>
  <!-- 关键：先把系统原有的配置整份包含进来。
       只写一个 <dir> 会把系统字体全挡住，用户机器上装的真「仿宋_GB2312」反而认不出来。 -->
  <include ignore_missing="yes">/etc/fonts/fonts.conf</include>

  <!-- 包里自带一份开源中文字体，只作「一台字都没有」时的兜底。
       刻意不写 alias：一旦把「仿宋_GB2312」别名到文泉驿，
       会盖掉系统里真正的方正/微软字体映射，版面反而走形。
       真缺字体时程序会在界面上老实提示，这比悄悄换字更负责。 -->
  <dir prefix="relative">truetype/gwtyper</dir>
</fontconfig>
FC

    # ---- 内嵌 LibreOffice（Linux 不再要求目标机装办公套件）----
    # LibreOffice 的可执行文件把 /usr/lib/libreoffice 写死在编译期，所以
    # 包里必须保持这个相对结构，运行时由 lo_bundle.py 用 unshare bind
    # （免 root）或 root 挂载把它贴到系统路径去。
    if [ "${BUNDLE_LO:-1}" = "0" ]; then
        echo "== 跳过内嵌 LibreOffice（BUNDLE_LO=0）=="
    elif [ -x "${LO_SRC:-/usr/lib/libreoffice}/program/soffice" ]; then
        LO_SRC="${LO_SRC:-/usr/lib/libreoffice}"
        echo "== 内嵌 LibreOffice（来源：$LO_SRC）=="
        mkdir -p "$APPDIR/usr/lib"
        cp -a "$LO_SRC" "$APPDIR/usr/lib/libreoffice"
        # 画廊/模板/向导对「转 PDF」没用，删掉省 25MB 左右。
        # 注意 share/config 和 share/registry 不能动——删了 soffice 直接起不来。
        rm -rf "$APPDIR/usr/lib/libreoffice/share/gallery" \
               "$APPDIR/usr/lib/libreoffice/share/template" \
               "$APPDIR/usr/lib/libreoffice/share/wizards" 2>/dev/null
        echo "  LibreOffice: $(du -sh "$APPDIR/usr/lib/libreoffice" | cut -f1)"

        # LibreOffice 依赖一批系统库（libuno_* / libicu / libboost…），
        # 目标机不一定有，得一起带上。
        #
        # 关键：必须放进 program/ 目录，**不能**用 LD_LIBRARY_PATH 指向别处。
        # 否则 LibreOffice 会把那个目录当成自己的安装根，去找 <那里>/unorc，
        # 找不到就 "The component manager is not available" 然后段错误（rc=139）。
        # 这个坑踩过，写在注释里免得以后手贱改回去。
        PROG="$APPDIR/usr/lib/libreoffice/program"
        _soffice="$PROG/soffice.bin"
        # 这一段用 `set +e` 包住：交叉打包时 ldd 读不了别的架构，命令会返回
        # 非零；script 开了 set -e，直接就把整个打包中断了（表现为日志停在
        # "LibreOffice: 287M" 然后 rc=1，看不出原因）。
        set +e
        if [ -x "$_soffice" ]; then
            # 交叉打包时（arm64），系统库要从对应 rootfs 里找，不能用本机的
            SYSROOT="${SYSROOT:-}"
            { ldd "$_soffice" 2>/dev/null | awk '/=> \//{print $3}';
              ldd "$PROG/libmergedlo.so" 2>/dev/null | awk '/=> \//{print $3}';
              ldd "$PROG/libsclo.so" 2>/dev/null | awk '/=> \//{print $3}';
              ldd "$PROG/libswlo.so" 2>/dev/null | awk '/=> \//{print $3}';
              ldd "$PROG/libsdlo.so" 2>/dev/null | awk '/=> \//{print $3}';
            } | sort -u | while read -r so; do
                [ -f "$so" ] || continue
                base=$(basename "$so")
                case "$base" in
                    libc.so*|libm.so*|libdl.so*|libpthread.so*|librt.so*|\
                    ld-linux*|libgcc_s.so*|libstdc++.so*) continue ;;
                esac
                [ -e "$PROG/$base" ] || cp -L "$so" "$PROG/$base" 2>/dev/null || true
            done
        fi
        # 交叉打包：ldd 读不了别的架构，直接从 rootfs 里按名字捞
        if [ -n "${SYSROOT:-}" ] && [ -d "$SYSROOT/usr/lib" ]; then
            for hit in $(find "$SYSROOT/usr/lib" "$SYSROOT/lib" -maxdepth 3 \
                         \( -name 'libuno*' -o -name 'libicu*.so.*' \
                            -o -name 'libboost_*.so.*' \) 2>/dev/null); do
                base=$(basename "$hit")
                [ -e "$PROG/$base" ] || cp -L "$hit" "$PROG/$base" 2>/dev/null || true
            done
        fi
        # LibreOffice 自己的 .so 里还有 dlopen 的组件，按名字补齐。
        # 交叉打包时优先从 SYSROOT（arm64 rootfs）里找，找不到才看本机。
        _find_roots="/usr/lib"
        if [ -n "${SYSROOT:-}" ]; then
            _find_roots="$SYSROOT/usr/lib $SYSROOT/lib"
        fi
        for so in libuno_sal.so.* libuno_salhelpergcc3.so.* libuno_cppu.so.* \
                  libuno_cppuhelpergcc3.so.* libuno_purpenvhelpergcc3.so.* \
                  libicuuc.so.* libicui18n.so.* libicudata.so.* \
                  libboost_system.so.* libboost_filesystem.so.* \
                  libxml2.so.* libxslt.so.* libcairo.so.* libcups.so.* \
                  libdbus-1.so.* libnss3.so libnspr4.so libplc4.so libplds4.so \
                  libsmime3.so libnssutil3.so libexpat.so.* libfontconfig.so.* \
                  libfreetype.so.* libharfbuzz.so.* libgraphite2.so.* \
                  libjpeg.so.* libpng16.so.* libtiff.so.* liblcms2.so.* \
                  libpoppler.so.* libgpgme.so.* libgpg-error.so.* \
                  libabw-0.1.so.* libwps-0.4.so.* libwpd-0.10.so.* \
                  libvisio-0.1.so.* libetonyek-0.1.so.* libmwaw-0.3.so.* \
                  libexttextcat-2.0.so.* libhyphen.so.* libmythes-1.2.so.* \
                  libclucene-core.so.* libclucene-shared.so.* \
                  libnspr4.so libnssutil3.so libsmime3.so libsqlite3.so.* \
                  libenchant-2.so.* libhunspell-1.7.so.* libmariadb.so.* \
                  libcurl.so.* libldap.so.* libsasl2.so.* liblber.so.*; do
            for hit in $(find $_find_roots -maxdepth 3 -name "$so" 2>/dev/null); do
                base=$(basename "$hit")
                [ -e "$PROG/$base" ] || cp -L "$hit" "$PROG/$base" 2>/dev/null || true
            done
        done
        set -e
        echo "  program 里的 .so: $(ls "$PROG"/*.so* 2>/dev/null | wc -l) 个"
    else
        echo "== 本机没装 LibreOffice，跳过内嵌 ==" >&2
        echo "   要内嵌请先：sudo apt install libreoffice-writer libreoffice-calc libreoffice-impress" >&2
    fi
elif [ "$MODE" = "thin" ]; then
    # ------------------------------------------------------------ 薄包
    echo "== 复制纯 Python 源码（thin）=="
    DEST="$APPDIR/usr/share/gwtyper/gwtyper"
    mkdir -p "$DEST"
    # APPDIR 本身就在 $PKG/build 下，直接 `cp -r $PKG` 会把目录拷进自己。
    # 用 tar 流 + 排除项复制，既避开递归，也顺手丢掉 __pycache__/dist/build。
    tar -C "$PKG" -cf - \
        --exclude='./build' --exclude='./dist' --exclude='__pycache__' \
        --exclude='./*.spec' . | tar -C "$DEST" -xf -

    cat > "$APPDIR/usr/bin/gwtyper" <<'LAUNCH'
#!/bin/sh
HERE="$(dirname "$(readlink -f "$0")")/../share/gwtyper"
exec python3 "$HERE/gwtyper/packaging/entry.py" "$@"
LAUNCH
    chmod +x "$APPDIR/usr/bin/gwtyper"
fi

cp "$PKG/packaging/gwtyper.desktop" "$APPDIR/usr/share/applications/"
cp "$PKG/packaging/gwtyper.desktop" "$APPDIR/"
cp "$PKG/packaging/AppRun" "$APPDIR/AppRun"
chmod +x "$APPDIR/AppRun"

if [ -f "$ROOT/converter/assets/icon.png" ]; then
    cp "$ROOT/converter/assets/icon.png" \
       "$APPDIR/usr/share/icons/hicolor/256x256/apps/gwtyper.png"
    cp "$ROOT/converter/assets/icon.png" "$APPDIR/gwtyper.png"
fi

if [ "$MODE" = "thin" ]; then
    echo "== 自检：能不能 import =="
    PYTHONPATH="$APPDIR/usr/share/gwtyper" "$PY_BIN" -c \
      "import gwtyper.core.sysenv as s; print('sysenv ok ->', s.platform_note())" \
      || { echo "自检失败：缺 python3-docx / python3-yaml / python3-fitz / python3-tk 之类"; exit 1; }
fi

echo "== 打 AppImage =="
if [ "$MODE" = "full" ]; then
    OUT="$PKG/dist/${APP}-${VER}-${ARCH}.AppImage"
else
    OUT="$PKG/dist/${APP}-${VER}-${ARCH}-thin.AppImage"
fi
mkdir -p "$PKG/dist"

# 不用 appimagetool：它在国内网络下取 runtime 老失败，且它的 --runtime-file
# 在这些版本里参数解析有毛病。AppImage 本质是「runtime 可执行文件 + squashfs」，
# 自己拼反而更稳、也更好跨架构（换对应架构的 runtime 就行）。
case "$ARCH" in
    x86_64)  RTSUF="x86_64" ;;
    aarch64) RTSUF="aarch64" ;;
    armhf)   RTSUF="armhf" ;;
    *)       RTSUF="" ;;
esac
[ -n "$RTSUF" ] || { echo "不认识的架构：$ARCH（可用 x86_64 / aarch64 / armhf）"; exit 1; }

if ! command -v mksquashfs >/dev/null; then
    echo "需要 squashfs-tools：sudo apt install -y squashfs-tools"
    exit 1
fi

RT="$PKG/build/runtime-$ARCH"
if [ ! -s "$RT" ]; then
    echo "  取 runtime（$RTSUF）"
    mkdir -p "$PKG/build"
    RT_PATH="AppImage/type2-runtime/releases/download/continuous/runtime-$RTSUF"
    got=""
    # GitHub 在国内经常掉，按顺序试镜像
    for M in "https://github.com" \
             "https://ghfast.top/https://github.com" \
             "https://gh-proxy.com/https://github.com" \
             "https://ghproxy.net/https://github.com"; do
        if curl -fL --retry 2 --connect-timeout 20 -o "$RT.tmp" "$M/$RT_PATH"; then
            mv "$RT.tmp" "$RT"; chmod +x "$RT"; got="$M"; break
        fi
        rm -f "$RT.tmp"
    done
    [ -n "$got" ] || { echo "runtime 下载失败（GitHub 与镜像都不通），可从 $RT_PATH 手工取"; exit 1; }
    echo "  runtime 来自 $got"
fi

# .DirIcon 让文件管理器能显示图标；缺了不影响运行
[ -f "$APPDIR/gwtyper.png" ] && ln -sf gwtyper.png "$APPDIR/.DirIcon"

SQ="$PKG/build/AppDir-$ARCH.squashfs"
rm -f "$SQ" "$OUT"
mksquashfs "$APPDIR" "$SQ" -root-owned -noappend -comp zstd -quiet
cat "$RT" "$SQ" > "$OUT"
chmod +x "$OUT"
rm -f "$SQ"

echo "完成：$OUT"
ls -lh "$OUT"
