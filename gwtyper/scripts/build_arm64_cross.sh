#!/bin/bash
# 一步打出 Linux ARM64 自带运行时包（面向麒麟 V10 / glibc 2.28 及以上）
#
# 用法:
#   bash gwtyper/scripts/build_arm64_cross.sh              # 只出便携版 tar.gz
#   bash gwtyper/scripts/build_arm64_cross.sh --appimage   # 另出一个 AppImage
#   ROOTFS=/opt/arm64deb10 bash gwtyper/scripts/build_arm64_cross.sh
#
# 原理
# ----
# 目标机是麒麟 V10（glibc 2.28）。用新发行版（glibc 2.39）打出来的包，
# 在麒麟上会报 GLIBC_2.34 not found，用户看到的就是“缺 python 组件”。
# 所以只在 glibc 2.28 的 arm64 rootfs 里打——用 qemu 用户态模拟跑，
# 慢（约 5 分钟），但出来的包真能跑。
#
# 依赖库怎么收
# ------------
# PyInstaller 靠 ldd 收依赖，而 qemu+proot 里 ldd 对任何文件都报
# "you do not have read permission"，结果 _tkinter.so 进了包、它依赖的
# libtk8.6/libtcl8.6/libX11 一个都没进，GUI 一启动就崩。
# 这里先用 scripts/elfdeps.py 自己解析 ELF 的 DT_NEEDED，把真实库清单
# 写进 /tmp/so_list.txt，再由 build_linux.spec 显式塞进 binaries。
#
# 注意：旧版本这个脚本用的是“伪装 PLATFORM 在 x86 上跑 + 替换 ELF”，
# 1 分钟出包，但找不到 arm64 对应版本时会静默 unlink 丢库，打出来的包
# 缺 tkinter / X11，正是上次国产机上“打不开、说缺 python 组件”的原因。
# 那个做法已废弃，不要再启用（git 历史里还能翻到旧版）。
set -o pipefail
export LANG=C.UTF-8

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO="$(cd "$HERE/../.." && pwd)"
cd "$REPO" || exit 1

ROOTFS="${ROOTFS:-/opt/arm64deb10}"
DO_APPIMAGE=0
for a in "$@"; do
  case "$a" in
    --appimage) DO_APPIMAGE=1 ;;
    --help|-h) sed -n '2,30p' "$0"; exit 0 ;;
  esac
done

VER="$(grep -m1 '^APP_VER = ' gwtyper/app.py | sed 's/.*= *"\([^"]*\)".*/\1/')"
VER="${VER:-0.0}"
DIST="gwtyper/build/linux-dist-arm64-deb10"
WORK="gwtyper/build/linux-work-arm64-deb10"
LOG=".tmp/arm64_build.log"
SO_LIST="/tmp/so_list_arm64.txt"
BUILD_ROOT="${BUILD_ROOT:-/root/gw-build}"
mkdir -p .tmp gwtyper/build

# 为什么不在 /mnt/d 里直接构建：
#   /mnt/d 是 Windows 盘经 9p 共享给 WSL，几万个小文件来回读写会被卡成龟速，
#   实测 qemu 进程的 CPU 只有 13%~20%，瓶颈全在 I/O（32 分钟里大半在等磁盘）。
#   把源码同步到 WSL 原生 ext4（/root/gw-build）再构建，通常能快 3~8 倍。
#   产物打完再拷回 repo 的 gwtyper/build/ 下，后续打包/AppImage 逻辑不变。
Q="proot -q /usr/bin/qemu-aarch64-static -S $ROOTFS -b /dev -b /proc -b /sys -b /etc/resolv.conf -b $BUILD_ROOT:/work -w /work"
E="/usr/bin/env PATH=/usr/sbin:/usr/bin:/sbin:/bin LANG=C.UTF-8"

echo "== 0/6 检查 rootfs =="
if [ ! -x "$ROOTFS/usr/bin/python3" ]; then
  echo "找不到 arm64 rootfs：$ROOTFS"
  echo "先建一个 Debian 10 arm64 的 rootfs（glibc 2.28），装 python3 python3-tk python3-dev binutils 和 pyinstaller，"
  echo "详见 gwtyper/README.md 的「麒麟 V10 / glibc 2.28」一节。"
  exit 1
fi
$Q $E /usr/bin/python3 -c 'import sys,PyInstaller;print("rootfs python",sys.version.split()[0],"| PyInstaller",PyInstaller.__version__)' || {
  echo "rootfs 里缺 python3 / PyInstaller"; exit 1; }

echo "== 1/6 同步源码到 WSL 原生盘（避开 /mnt 的 9p 慢速 I/O）=="
mkdir -p "$BUILD_ROOT"
if command -v rsync >/dev/null 2>&1; then
  rsync -a --delete \
    --exclude build --exclude dist --exclude __pycache__ --exclude '*.pyc' \
    --exclude .tmp --exclude docs \
    "$REPO/gwtyper/" "$BUILD_ROOT/gwtyper/"
else
  rm -rf "$BUILD_ROOT/gwtyper"
  mkdir -p "$BUILD_ROOT/gwtyper"
  ( cd "$REPO/gwtyper" && tar -cf - \
      --exclude=build --exclude=dist --exclude=__pycache__ --exclude='*.pyc' \
      --exclude=.tmp --exclude=docs . ) | tar -xf - -C "$BUILD_ROOT/gwtyper"
fi
echo "  源码 → $BUILD_ROOT/gwtyper（$(du -sh "$BUILD_ROOT/gwtyper" | cut -f1)）"

echo "== 2/6 解析依赖库（elfdeps，绕开 ldd）=="
SITE=/root/.local/lib/python3.7/site-packages
# Pillow / PyMuPDF 等会带自己的一份 .libs 目录，里面是改了名字的库
# （libjpeg-0b62f035.so.62.3.0 这种），必须一起作为查找目录。
# 注意：整个命令要在 rootfs 的 shell 里跑，否则 _tkinter*.so 这种通配符
# 不会被展开（proot 直接 exec 不会做 glob），tk 相关库就会漏收。
$Q $E /bin/sh -c "/usr/bin/python3 /work/gwtyper/scripts/elfdeps.py \
    /usr/bin/python3 \
    /usr/lib/python3.7/lib-dynload/_tkinter*.so \
    $SITE/PIL/*.so $SITE/lxml/*.so $SITE/fitz/*.so \
    $SITE/cv2/*.so $SITE/cv2/.libs/*.so* \
    $SITE/onnxruntime/capi/*.so* \
    $SITE/numpy/core/*.so $SITE/numpy.libs/*.so* \
    $SITE/pyclipper/*.so $SITE/shapely/*.so $SITE/shapely.libs/*.so* \
    --dirs /usr/lib/aarch64-linux-gnu /lib/aarch64-linux-gnu /usr/lib /lib \
           $SITE/*.libs $SITE" \
    > "$SO_LIST" 2>".tmp/elfdeps.err"
NLIB=$(wc -l < "$SO_LIST" | tr -d ' ')
echo "  依赖库 $NLIB 个 → $SO_LIST"
# Debian 10 的库比新发行版少（没有 libicu74 / libbrotli 这些），40 个上下算正常
if [ "$NLIB" -lt 30 ]; then echo "库数量异常（$NLIB），多半 elfdeps 失败："; tail -5 .tmp/elfdeps.err; exit 1; fi
if ! grep -q "libtk8.6" "$SO_LIST"; then
  echo "依赖清单里没有 libtk8.6 —— tkinter 会起不来，先查 elfdeps 的 _tkinter*.so 通配"; exit 1; fi

echo "== 3/6 PyInstaller（qemu 里真跑）=="
rm -f "$LOG"
$Q /usr/bin/env PATH=/usr/sbin:/usr/bin:/sbin:/bin LANG=C.UTF-8 GWTYPER_SO_LIST="$SO_LIST" \
   /usr/bin/python3 -m PyInstaller --noconfirm --clean \
   --distpath /work/"$DIST" --workpath /work/"$WORK" \
   /work/gwtyper/build_linux.spec > "$LOG" 2>&1
rc=$?
if [ $rc -ne 0 ]; then echo "构建失败，看 $LOG"; tail -25 "$LOG"; exit 1; fi
grep -E "显式收集|本机缺这些库" "$LOG" | tail -2

echo "  把产物拷回仓库：$DIST"
rm -rf "$DIST"
mkdir -p "$(dirname "$DIST")"
cp -a "$BUILD_ROOT/$DIST" "$DIST"

echo "== 4/6 校验架构（必须全是 aarch64）=="
APP="$DIST/gwtyper"
if [ ! -x "$APP/gwtyper" ]; then echo "没找到产物 $APP/gwtyper"; exit 1; fi
TOTAL=$(find "$APP" -type f -print0 | xargs -0 file 2>/dev/null | grep -c "ELF")
BAD=$(find "$APP" -type f -print0 | xargs -0 file 2>/dev/null | grep "ELF" | grep -vc "ARM aarch64")
echo "  ELF 共 $TOTAL 个，非 aarch64：$BAD"
if [ "$TOTAL" -lt 50 ] || [ "$BAD" -ne 0 ]; then
  echo "架构不对，拒绝出包："
  find "$APP" -type f -print0 | xargs -0 file 2>/dev/null | grep "ELF" | grep -v "ARM aarch64" | head -10
  exit 1
fi

echo "== 5/6 打包便携版 tar.gz =="
PKGDIR=".tmp/portable-arm64/公文工具箱-$VER-arm64"
rm -rf .tmp/portable-arm64
mkdir -p "$PKGDIR"
cp -a "$APP" "$PKGDIR/app"
cat > "$PKGDIR/公文工具箱" <<'LAUNCH'
#!/bin/sh
# 不需要 PATH 里有任何东西，也不依赖 dirname
HERE=${0%/*}
case "$HERE" in "$0") HERE=.;; esac
exec "$HERE/app/gwtyper" "$@"
LAUNCH
chmod +x "$PKGDIR/公文工具箱"
cat > "$PKGDIR/使用说明.txt" <<EOF
公文工具箱 $VER（Linux ARM64 / 麒麟 V10 等 glibc 2.28 及以上）

解压后：
  ./公文工具箱                                    打开界面
  ./公文工具箱 check 某文件.docx -p gb9704          只体检，不改原文件
  ./公文工具箱 format 某文件.docx -p gb9704         排版另存副本

不需要安装 Python，也不需要联网。装 WPS 或 LibreOffice 才能导出 PDF、看版面预览。
若提示没有执行权限：chmod +x 公文工具箱
EOF
OUT_TGZ="gwtyper/dist/公文工具箱-$VER-linux-arm64-便携版.tar.gz"
tar -C .tmp/portable-arm64 -czf "$OUT_TGZ" "公文工具箱-$VER-arm64"
echo "  → $OUT_TGZ（$(du -h "$OUT_TGZ" | cut -f1)）"

if [ "$DO_APPIMAGE" = "1" ]; then
  echo "== 6/6 打 AppImage =="
  DIST_SRC="$REPO/$APP" bash gwtyper/scripts/build_appimage.sh aarch64 || echo "AppImage 打包失败（不影响便携版）"
else
  echo "== 6/6 跳过 AppImage（加 --appimage 才打）=="
fi

cat <<EOF

完成。
  便携版: $OUT_TGZ
  产物目录: $APP
自检（在一台没有 python 的 arm64 机器上，或 proot 里）：
  ./公文工具箱 check 某个公文.docx -p gb9704
EOF
