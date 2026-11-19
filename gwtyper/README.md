# 公文工具箱

两个页签，一个程序：

- **公文排版** —— 以 GB/T 9704—2012《党政机关公文格式》为基准，所有参数可替换的体检与排版
- **格式转换** —— 公文/表格/演示/图片互转，外加合并、拆分、图片与 PDF 互换

## 启动

三种方式，任选：

| 场景 | 用什么 |
|---|---|
| Windows，给同事用 | 双击 `dist\公文工具箱.exe`，免装 Python |
| Windows，开发 | 双击 `启动公文排版工具.bat`（必须从 `D:\公文处理` 起，否则找不到包） |
| Linux x86_64（推荐） | 双击 `公文工具箱-<ver>-x86_64.AppImage`（180MB，自带 Python 运行时 **和 LibreOffice**，干净机器直接能跑，不用装任何办公软件） |
| Linux ARM64（国产机、树莓派、飞腾/鲲鹏） | `公文工具箱-<ver>-aarch64.AppImage`（175MB，同样自带运行时和 arm64 版 LibreOffice） |
| 只需要基本 PDF / 批量转换 | 用 `-thin.AppImage`（1MB）+ 一次 `apt install` 铺依赖 |

## 跨平台

内核（读 docx、认角色、比对、写回）是纯 python-docx，三个平台一模一样。
平台差异全都收在三个文件里，别的地方不许再判断 `sys.platform`：

| 文件 | 管什么 |
|---|---|
| `core/sysenv.py` | 数据目录、缓存目录、日志、打开文件/文件夹 |
| `core/office.py` | 转换引擎：Windows 走 WPS/Office 的 COM（按扩展名分 Word/Excel/PPT 三套 ProgID）；Linux/macOS 先试 **WPS**，没有再退 LibreOffice |
| `core/pdfwrite.py` | 纯 Python 排版：docx/xlsx/txt → PDF，**完全不调用办公软件** |
| `core/fontfind.py` | 中文字体名 → 实际文件，并用 fontTools 只嵌入用到的字 |
| `core/fonts.py` | 枚举本机字体：Windows 用 GDI，其余用 `fc-list` |

规则集分两层：内置的随程序发布（只读），你自己学出来或改出来的存到用户目录，同名优先。

- Windows：`%LOCALAPPDATA%\公文排版工具\rules`
- Linux：`~/.local/share/公文排版工具/rules`
- macOS：`~/Library/Application Support/公文排版工具/rules`

### Linux 依赖

```bash
sudo apt install -y python3 python3-tk python3-docx python3-yaml python3-fitz python3-pil \
                    fonts-noto-cjk
# 转换引擎二选一（政府单位多用 WPS，装了会优先用 WPS）：
sudo apt install -y wps-office            # 或去 wps.cn 下 deb 包
sudo apt install -y libreoffice-writer    # 没有 WPS 时用它兜底
# 想要和 Windows 完全一致的字体观感，再手动装方正小标宋、仿宋_GB2312、楷体_GB2312、黑体
# 放到 ~/.local/share/fonts/ 后 fc-cache -f
```

引擎优先级是 **WPS → LibreOffice**（`office.engine_name()` 反映实际在用的那个）。
WPS 的命令行转换不是每个版本都有：一旦调用失败，程序会记下并在本次运行内不再重试，自动退到 LibreOffice。
`libreoffice-writer` 是版面预览和 .doc 转换的前提，两者都没有就只能体检不能看版面。
PDF 那块 Ubuntu 叫 **`python3-fitz`**（模块名 `fitz`），pip 装的 PyMuPDF 才叫 `pymupdf`——
代码里 `core/pdfmod.py` 两个名字都试，所以两种装法都能跑。

### 打包

```bat
:: Windows → gwtyper\dist\公文工具箱.exe
python -m PyInstaller --noconfirm --clean gwtyper\build_win.spec
```

```bash
# Linux → gwtyper/dist/公文工具箱-<ver>-<arch>.AppImage
bash gwtyper/scripts/build_appimage.sh              # 按本机 uname -m 推架构，默认 full（自带运行时）
MODE=thin bash gwtyper/scripts/build_appimage.sh    # 薄包，约 1MB
```

**两种包，别搞混：**

| 模式 | 体积 | 里面有什么 | 干净机器能跑吗 |
|---|---|---|---|
| `full`（默认） | ~70MB | Python 3.12 解释器、tcl/tk、python-docx、PyMuPDF、lxml、一份开源中文字体 | **能**，双击就行 |
| `thin` | ~1MB | 只有 .py 源码 | **不能**，要先 `apt install` 一堆依赖 |

`full` 是给人用的，`thin` 是给运维批量铺的。早期版本只有 thin，1MB 看着可疑但确实是这个数——
里面就是几十个 .py 文件。现在默认改成 full，因为「拿到就能用」比「省 70MB」重要。

**full 模式不能交叉打包**：包里带的是打包机架构的 Python 解释器和 `.so`，脚本会直接拒绝
在 x86 上打 aarch64 的 full 包。要国产机（飞腾/鲲鹏/树莓派）版本，有三条路：

1. **在国产机上打**——装上 `python3-tk libtk8.6 binutils libpython3.12` 和
   `pip install pyinstaller python-docx pyyaml pymupdf pillow`，然后同样一条命令
2. **在 x86 机上用 proot + qemu 交叉打**（我就是这么打出 aarch64 包的，见下）
3. 用 thin 包 + 目标机 `apt install` 依赖

proot 交叉打包（不需要 root、不需要 binfmt，**别碰 `binfmt_misc`，会把 WSL 一起搞坏**）：

> **更快的办法**：在 qemu 里直接跑 PyInstaller 要 **31 分钟**（Analysis 阶段
> 光 `hook-pickle` 就 181 秒、`hook-heapq` 165 秒，全是纯 Python 逻辑被逐条翻译字节码）。
> 用 `bash gwtyper/scripts/build_arm64_cross.sh` 只要 **2 分钟**——它在 x86 上跑
> Analysis，然后把 .so / 解释器 / bootloader 换成 arm64 的，最后那步 `objcopy`
> 也转给 qemu 执行（本机 objcopy 认不出 arm64 文件）。脚本会自动拼好 AppImage。
> 前提是 `/opt/arm64root` 已经存在（下面 1~3 步）。

```bash
# 1. 一个 arm64 的 Ubuntu 根文件系统（29MB 的 base 就够了）
curl -L -o /tmp/arm64.tar.gz \
  https://cdimage.ubuntu.com/ubuntu-base/releases/24.04/release/ubuntu-base-24.04.5-base-arm64.tar.gz
sudo mkdir -p /opt/arm64root && sudo tar -xzf /tmp/arm64.tar.gz -C /opt/arm64root
sudo cp /usr/bin/qemu-aarch64-static /opt/arm64root/usr/bin/

# 2. 在里面装依赖（源必须用 http，https 在 proot 里过不了证书）
Q="proot -q /usr/bin/qemu-aarch64-static -S /opt/arm64root -b /dev -b /proc -b /sys -w /root"
$Q apt-get install -y python3 python3-tk libtk8.6 libpython3.12 binutils zlib1g

# 3. Python 依赖要在主机上下 aarch64 的 wheel，再离线装（qemu 里直接 pip 会挂死）
python3 -m pip download -d /tmp/wheels --only-binary=:all: \
  --platform manylinux2014_aarch64 --python-version 3.12 --implementation cp --abi cp312 \
  pyinstaller python-docx pyyaml pymupdf pillow fonttools openpyxl python-pptx
$Q pip install --break-system-packages --no-index --find-links /tmp/wheels \
  pyinstaller python-docx pyyaml pymupdf pillow fonttools openpyxl python-pptx

# 4. 一条命令出包（交叉构建，约 2 分钟）
bash gwtyper/scripts/build_arm64_cross.sh
```

三个必踩的坑，都记在这里省得再走一遍：`binutils`（PyInstaller 要 `objdump`）、
`libpython3.12`（否则报 `Python shared library was not found`）、
`pyinstaller` 的 wheel 必须带 `--platform manylinux2014_aarch64`（主机默认下的是 x86 那个，装进去报 `No matching distribution found`）。

第四个坑最阴：**qemu+proot 里 `ldd` 用不了**（对任何文件都报 `you do not have read permission`），
PyInstaller 靠 ldd 收集依赖，于是 `_tkinter.so` 能进包、它依赖的 `libtk8.6`/`libtcl8.6`/`libX11`
一个都没进——CLI 一切正常，GUI 一启动就 `Failed to execute script 'entry'`。
`build_linux.spec` 里因此写死了 55 个 .so 名字，用 `sysconfig` 的 MULTIARCH 找路径、
`os.path.realpath` 解软链后显式塞进 `binaries`，跟 ldd 彻底无关，两个架构通吃。
构建日志里会打一行 `[build_linux.spec] 显式收集 55 个 .so`，看到它才算过。

AppImage 的结构是「runtime 可执行文件 + squashfs」，脚本自己拼（`mksquashfs` 后 `cat`），
不走 `appimagetool`（它把 runtime 写死在自己逻辑里，`--runtime-file` 参数解析有问题，WSL 里还缺 fusermount）。
runtime 从 GitHub 拉，按 `github.com` → `ghfast.top` → `gh-proxy.com` → `ghproxy.net` 顺序试镜像，实测只有 `ghfast.top` 通。

### 自带运行时到底打了什么（实测）

```bash
$ ls .../_internal/libpython3.12.so.1.0
$ ls .../_internal/libtk8.6.so
$ ldd .../gwtyper | grep "not found"     # 空
```

只有 4 个动态依赖，全是 glibc 级别的（`libc`/`libdl`/`libz`/`libpthread`），
任何 2018 年以后的发行版都有。**把系统 python3 换成一定失败的假程序，CLI 和 GUI 照样能跑**——
这是判断「是否真的自带」最硬的一条证据，我自己就是这么验的：

```bash
printf '#!/bin/sh\nexit 99\n' > /tmp/fakebin/python3 && chmod +x /tmp/fakebin/python3
PATH="/tmp/fakebin:$PATH" ./公文工具箱-0.9-x86_64.AppImage check 某公文.docx -p news
```

### Linux 分支怎么验的

真机验过了（Ubuntu 24.04 / WSL2 / LibreOffice 24.2）：

| 验的东西 | 结果 |
|---|---|
| `office.engine_name()` | `libreoffice`，`soffice` 在 `/usr/bin/soffice` |
| 真实 docx → PDF | 3 页、142KB、中文正常不乱码 |
| 版面预览 `preview.render` | 3 页，首页 72KB PNG |
| 版面体检 `layout_check` | 正常出结果（报 2 条，见下） |
| 字体枚举 | 32 个族，能认出 Noto Sans CJK SC / WenQuanYi |
| 规则集双目录 | 内置可读、`save_target` 落 `~/.local/share/公文排版工具/rules` |
| GUI（WSLg） | 窗口起来，标题「公文工具箱 0.9」；拖拽退化为无（`_HAS_DND=False`） |
| AppImage x86_64（full） | 70MB，自带 Python 3.12 + tcl/tk + PyMuPDF；把系统 python3 换成假货后 CLI、GUI 仍然能跑 |
| AppImage aarch64（full） | 104MB，内部 56 个 .so（含 `libtk8.6`/`libtcl8.6`/`libX11`）；在本机用 proot+qemu 跑，假 python3 下 `check` 问题 0、GUI 存活 15 秒 |

两条版面问题都是**真的**，不是误报：LibreOffice 没有方正小标宋/仿宋_GB2312，用替代字体后
字宽变了，第 1 页文字超出右边距 0.57cm，末页只剩几行。装上那四款字体就消失。
这也说明版面体检在 Linux 上是有效的——它抓到的是排版结果，不是字体名。

临时性的平台差异只有一处要额外兜：**PyMuPDF 的模块名**。见上面 `python3-fitz` 说明。


## 两个动作是分开的

| 动作 | 会不会改文件 | 用途 |
|---|---|---|
| 检查（只读） | 不改 | 出问题清单，先看再动手 |
| 按规则排版… | 写新文件 `原名-排版后.docx` | 原件永不覆盖；写之前会弹确认，提示会跳过几段、目标文件是否已存在 |

界面里还多两个批量用的（「更多 ▾」菜单里）：

- **检查全部** —— 左侧所有文件一起查，右侧按文件分组列问题，左侧每行显示「必改 N」
- **导出报告** —— 写成 txt（UTF-8 BOM，Excel 直接打开不乱码），单份是明细，多份是汇总＋明细

## 界面

顶部四个页签：**公文排版** / **格式转换** / **OCR 识别** / **办公小工具**。

### 公文排版

四栏：左边是公文列表和动作，中间是段落角色，然后是检查结果，右边是版面预览。

- 空的时候每个栏都有提示文字，不会只剩一片白
- 主按钮是**按规则排版…**（写新文件），**检查（只读）**是次要按钮；两个在没打开文件时都是灰的
- 段落角色表里「确定度」分三档：确定 / 可能 / 拿不准。双击改角色，右键可以撤销单段或全部的手工指定
- 检查结果表四列：位置 / 检查项 / **现在** / **应为**，左边一个状态点（● 必改、▲ 建议、· 通过），不整行染色
- 排完版结果条上直接有 **打开 / 转成 PDF / 所在文件夹**——报材料要 PDF 时不用再切到「格式转换」页转一遍
- 「更多 ▾」里第一条是 **一键处理（体检 + 排版 + 导 PDF + 报告）…**，然后是 **按规则排版全部文件…**（左边所有文件各排一份，可选放原件旁边或统一放一个文件夹）和 **把当前文件转成 PDF…**
- 「更多 ▾」里还有：**导出检查结果（CSV）…**（Excel 直接打开，便于分工改稿）、**对比排版前后差异…**（列出哪几段文字或角色变了）、**备份设置… / 恢复设置…**（格式标准 + 设置打成一个 zip，换电脑直接搬）、**清理后台 WPS 残留进程**（批量转换后可能留下无窗口的 wps.exe，一键收尾；你自己开着的文档不会被碰）
- 检查结果上面有 **健康分**（100 分往下扣：必改扣 7、建议扣 2，底 0），后面跟「很规整 / 良好 / 基本可用 / 需要修改 / 问题较多」
- 检查结果里双击一条问题，除了下面角色表跳到那一段，还会在 **Word / WPS 里打开文档并选中那一段**（打不开就只留在界面跳转，不打扰）
- 导出报告可选 **Word（.docx）/ PDF / 纯文本** 三种
- 批量检查、批量排版都是多线程并行（数多时明显快）
- 窗口按界面实际需要开，不超出可用桌面（150% 缩放、4K 屏也不会被裁）
- 勾选框是自绘的 `☑ / ☐`——clam 主题下 ttk 的勾选框会把选中画成 `✗`，容易被读成出错

### 格式转换

三栏：左边放要转的文件，中间设目标格式，右边是**预览**。

放文件有三种方式，效果一样：点「添加文件…」、点「添加文件夹…」（取里面一层）、
或者**直接把文件或文件夹拖进左边任意位置**。加完会自动选中第一个，右栏立刻出预览。

| 转成 | 能转的源 | 说明 |
|---|---|---|
| PDF 文档 | Word / Excel / PPT / 图片 / 文本 | 文档、表格、演示、图片都能合成 PDF |
| Word 文档（.docx） | PDF / 老格式 doc | PDF 走 pdf2docx，老格式走引擎另存 |
| Excel 表格（.xlsx） | PDF / Excel | 先试表格识别，识别不出就按文本两列导 |
| PPT 演示（.pptx） | PDF | 每页渲成图铺满一页 16:9 |
| 纯文本（.txt） | Word / PDF / PPT / Excel | — |
| CSV 表格（.csv） | Excel | — |
| 图片（.png / .jpg） | PDF / PPT | 清晰度可调（72~600 DPI，默认 150） |

下拉框只列**这批文件真的能转的**目标（按文件类型取交集），选不了就不会误导。

#### 预览

选中左边任一文件，右栏就把它的真实版面渲染出来（走同一套转换引擎转 PDF 再栅格化）。
转换前先看一眼对不对，比转完才发现排版乱了省事。切换文件自动刷新，
转换前的表格、演示都能看；纯文本没有版面，会直接说明而不是给个空白。

#### 转 PDF 时的纸面选项

选「PDF 文档」时，下面会多出一块「转 PDF 的纸面」，用来收拾那些纸面不齐的文件：

| 选项 | 干什么 |
|---|---|
| 保持原样 / 统一竖版 / 统一横版 | 把混着 A4 竖、A4 横、A5 的散件统一成一种纸型 |
| 四周留白 | 0~50 毫米。给要装订、要签字、要盖章的文件留出边 |
| 横竖互换时把内容转 90° | 宽表转竖页时用。勾上内容整体转 90°，不勾就等比缩小（字会变小但方向正） |

内容一律 `keep_proportion=True` 等比缩放居中，**绝不裁切**——宁可留白也不吃掉一个字。
留白设得比纸还大时会自动退回满页，不会把内容挤到看不见。

选别的目标格式时这块会灰掉——这些选项只跟「输出 PDF」有关。

#### PDF 引擎：Linux 包自带 LibreOffice，另有纯 Python 兜底

「转 PDF」有三种引擎可选，在右侧「PDF 引擎」那一行切换：

| 模式 | 谁在干活 | 什么时候用 |
|---|---|---|
| 外部引擎 | **Linux：包内自带 LibreOffice** / **Windows：WPS·Office COM** | 要跟办公软件里一模一样（pptx 只能走这条） |
| 纯 Python | 程序自己排版，不调用任何外部软件 | 不想让外部程序参与；或机器上确实没有办公引擎 |
| **自动（默认）** | **先试外部引擎，转不了自动退纯 Python** | 不知道选哪个就用这个 |

「自动」的顺序和平台无关，但「外部引擎」在两边是不同的东西：

- **Windows**：WPS 或 Microsoft Office 的 COM —— 转出来的就跟在 WPS 里另存 PDF 一样
- **Linux**：**包里自带的那个 LibreOffice**（不是系统装的）—— 目标机不用另装办公软件

外部引擎失败（没装、文件加密、格式不支持）时，自动模式会静默改用纯 Python 再试一次，
不会直接报错。

**Linux 的 AppImage 里已经打包了 LibreOffice**（约 180 MB 里占 100 MB），
所以目标机**不用再装任何办公软件**，docx / xlsx / pptx 全都能转。
Linux 版的 WPS 不再考虑——实测它不支持命令行转换，还会常驻 `wpscloudsvr` 进程。

包内 LibreOffice 怎么跑起来的，有两个坑值得记下：

1. **它的可执行文件把 `/usr/lib/libreoffice` 写死在编译期**，放到别的目录直接起不来。
   所以运行时用 `unshare -m --map-root-user` 建一个私有 mount namespace，
   把包内那份 bind 到 `/usr/lib/libreoffice`——免 root、不动系统、退出即消失。
   极少数加固过的发行版禁用了非特权 userns，这时用 `pkexec`/`sudo` 请求一次 root
   做全局挂载（已挂过就复用，不重复挂）。
2. **依赖库必须放进 `program/` 目录本身，不能用 `LD_LIBRARY_PATH` 指过去。**
   因为 LibreOffice 会把 `LD_LIBRARY_PATH` 的第一个目录当成自己的安装根，
   去那里找 `unorc`，找不到就报 `The component manager is not available`
   然后段错误（rc=139）。这个坑排查了很久，注释写在 `build_appimage.sh` 里了。

纯 Python 这条路能转：**docx / xlsx / csv / txt / 图片 / 图片合集**，以及已有的 PDF 改纸面。
不能转：**ppt / pptx**（演示版式太复杂），它会明确让你切到外部引擎，不会静默失败。

技术上就是 PyMuPDF 画、fontTools 裁字体，两个都是纯 Python，随程序打包、不需要额外安装。
字体按公文里的东亚字体名（`w:eastAsia`，不是西文那个 Times New Roman）逐个解析——
方正小标宋、仿宋_GB2312、楷体_GB2312、黑体各归各的，四种字体都正确嵌入。

体积上也做了处理：中文字体一整份 10 MB，全量嵌进 PDF 会变成 **48 MB**。
现在只裁进用到的字，同一份公文 **304 KB**（WPS 转出来是 127 KB，同一量级）。

想快速看这台机器上各个引擎到底能不能用，跑一句就行：

```bash
公文工具箱 doctor
```

它会列出：包内 LibreOffice 在哪、unshare 能不能用、是否已全局挂载、
纯 Python 引擎是否就绪、四种公文字体各自解析到哪个文件、系统里有没有 WPS/LibreOffice。

实测同一份公文两种引擎对比：

| | 页数 | 体积 | 耗时 | 字体 |
|---|---|---|---|---|
| 纯 Python | 3 页 | 304 KB | 1.45s | 方正小标宋 / 楷体 / 黑体 / 仿宋，4 种都对 |
| WPS | 3 页 | 127 KB | 1.38s | 同上 |

页数、字体、字号、缩进、行距都对得上；行末断字位置有毫米级差异（字宽算法不同），不影响阅读。

下面还有六个 PDF 小工具，独立于上面的转换：合并多个 PDF、图片转 PDF、按每 N 页拆分、PDF 转图片、
**改方向／留白**（对现成的 PDF/Office/图片改纸面，另存一份）、
**多份材料合并成一册（带书签）**（Word / PDF / 图片混着也能合，每份自动一个书签，长材料翻起来能直接跳）。

转换结果默认写到「保存到」那一栏（默认桌面），重名自动加 `(1)(2)`，**原文件不动**。

> 排版页（公文排版那个页签）的版面预览仍然走外部引擎——它要看的是「Word 里长什么样」，
> 那是办公软件的渲染结果，换成自己排就失去意义了。



### OCR 识别

三栏：左边放扫描件或图片，中间选成品类型，右边是原图预览（转换前先确认清楚）。

**认完一律先给你看**（没有开关，默认就是这么走）：纯文字 / Word / **双层 PDF** 三种成品认完不会直接落盘，
先弹一个「识别结果」窗口把认出来的字摆出来——认错的字当场改，字数实时统计，Ctrl+S 或点「保存」落盘，
「跳过这个」就不保存。多个文件会一个接一个弹。

**双层 PDF 的流程就是「先认字 → 你改 → 最后合成」**：底层铺原来的扫描图，上层把你改过的文字
按原来的段落位置放进去，所以打印还是原样、电脑里能搜能复制。改了字也不会跑位——文字层是按原框放的。
表格（Excel）是唯一不能改字的：它是按位置往格子里填，改字会串行，生成后弹一个只读预览供核对。

能认的源：图片（png / jpg / bmp / tif / webp）、PDF（含扫描件）、Word / Excel / PPT
（Office 文件直接读里面的文字，不用 OCR，所以是 100% 准确的）。

| 成品 | 用途 |
|---|---|
| 纯文字（.txt） | 只要字，不要格式；拷进别的文件里最省事 |
| Word（按原版式） | 尽量还原原来的版式：字号按字高推算，居中 / 右对齐 / 首行缩进按位置判断 |
| Word（只保留文字） | 干净的一份 Word，正文仿宋_GB2312 三号，自己再套格式 |
| 双层 PDF | 上面是原来的扫描图，底下藏一层看不见的文字，可以搜索、可以复制 |
| 表格（Excel） | 把表里的字按行列位置放进 Excel 格子。**PDF 里有真的表格线时用 PyMuPDF 的 find_tables（最准）**；扫描件、拍照才退回 OCR 行列聚类，合并单元格和跨行表头可能要人工再核一遍 |

**双层 PDF** 是给"原件必须保持扫描件样子、但要能搜"的场合用的：打印出来还是原样，
在电脑里 Ctrl+F 能搜到字，也能整段复制。也可以拿一份**已经是 PDF 的文件**（比如别人给的
纯图片 PDF）转成双层，不用先拆成图片。

### 办公小工具

第四页，四件办公室天天要干的小事，纯本地、不联网：

- **批量重命名**：把一摞材料改成「前缀 + 序号 + 原名」（如 `材料01重阳节.docx`），也可以只做
  查找替换；勾上「再加今天的日期」就变成 `20261118-材料01重阳节.docx`。先点「预览新名字」看一遍，
  有重名会拦下来，确认后再「开始改名」。
- **批量字数统计**：docx / doc / wps / rtf / pdf / txt 都能统计，列出每个文件的页数和字数，
  可以导出 CSV（带合计行）交差。
- **扫描件摆正 / 去黑边**：手机拍的、扫描歪了带黑边的，自动摆正（0.3°～15°）并把四周黑边裁掉。
  只做这两件事，不调颜色，字不会发虚；处理完存成「原名-摆正」，原图不动。
- **签到表 / 材料交接单**：填个标题、单位、日期，名单默认用「材料上交」页的村名单，
  直接生成能打印的 Word 表格（A4、标题小标宋二号、表格黑体表头 + 仿宋正文、行高 0.95cm）。

### 材料上交

第五页，专治「催材料」：按你给的村（单位）名单，核对谁交了、谁没交。

- 左边把名单贴进来（一行一个），点「保存名单」——**下次打开自动带出来**（存在本机，不上传）。
  也可以从 txt / csv 导入。
- **从 Excel 解析**：点「从 Excel 解析…」选一张村里报上来的表，程序自己找「哪一列像村名」
  （表头「序号/村名/单位/备注」、纯数字、带标点的都会跳过），**解析完先弹一个窗口让你核对**，
  多一行少一行都能当场改，点「就用这些」才填进名单——不会悄悄改你的名单。只认 .xlsx / .xlsm（老 .xls 请先另存）。
- **预设名单**：点「存为预设…」把当前名单存成一套（如「2026各村」），下拉框里随时「载入」，
  下次换一批材料不用重贴；旁边的「删」只删预设，不动当前名单。
- 名单里**不必都是「某某村」**：社区、站所、科室、学校、公司、'102团'这种都行，规则是「文件名里含这个名称就算交了」。
  名称带编号（「102团」「3号站」）、带括号（「财政所（东片）」）也照常认；写「西村村民委员会」时，
  文件名只写「西村」也能对上（会自动去掉村委会/居委会/人民政府/办事处这类后缀再比）。
- 右边是台账：**双击某行**就切换「已交 / 未交」，依据列会写「手动标记」。
- 点「选文件夹按文件名认…」，挑一个装材料（或材料清单）的文件夹，**文件名里含哪个村名就算哪个村交了**，
  依据列写「文件名：xxx」。一份文件里同时提到两个村（如「西村、北村联合报表」），两个村都会算已交。
- 「全部已交 / 全部未交」一键批量标；底部实时显示「共 N 个，已交 X，未交 Y」。
- 「导出未交清单…」可存 txt 或 csv（txt 带日期抬头，直接发群里催；csv 方便再加工）。


DPI 默认 200。实测 150 / 200 / 250 / 300 的识别准确率都在 99.6% 以上，
150 最快；字特别小或者扫描得糊的时候往上加。

识别引擎是 RapidOCR（纯 ONNX，三个模型共 15.6MB），**随程序一起打包**，
不需要联网，也不用装别的软件。

已知的粗糙之处：按位置猜段落。公文正文两端对齐，所以"段末那行一定短"是最可靠的断段信号，
据此判断；但排版特别松散的件（比如一行只有一个词的清单）可能切得不准，
识别完请对照原图扫一眼。

## 格式标准

`rules/` 下三份现成的：

- `gb9704` 国家标准 GB/T 9704 公文格式（红头文件、通知、机制办法）—— 页边距 3.7/3.5/2.8/2.6
- `report` 述职/总结报告 —— 页边距 3.0/2.8/2.8/2.8
- `news` 新闻稿/通讯（一页内、可配图、不加页码）

**格式标准设置…**（更多 ▾ 菜单）可以在界面里直接改字号、行距、缩进、页边距、页码格式，改完「另存为」新标准。
改内置那三份时不会动原件，而是存成用户目录里的覆盖版本。

**从参照件学一套标准** —— 拿一份已经排好的公文当样本，反推出一份格式标准。生成的 yaml 建议先人工核一遍数值。

## 老格式（.doc / .wps / .rtf）

拖进来就能用，不用先手动另存。程序会用本机的转换引擎（Windows 上是 WPS 或 Office 的 COM，
Linux/macOS 上是 WPS，没有就退 LibreOffice）悄悄转成 docx 放在缓存目录，再走后面的流程。**原件一个字节都不动。**

- 同一个文件（大小和修改时间没变）第二次直接用缓存，不会重转
- 排版输出一律是 `.docx`
- 一个引擎都没有、或文件带密码时，会明确告诉你转不了，而不是给一堆看不懂的报错

## 段落角色

公文的层级是靠**字体**扛的，不是靠样式：黑体＝一级标题、楷体＝二级标题、方正小标宋＝大标题、仿宋＝正文。分类器就按这个信号加序号形态打分。

置信度低于 0.45 的段落会标黄，进「待确认」，排版时**跳过不猜**。中间栏双击那一段就能手工指定角色，改完立刻重算。

## 体检查什么

文本层（`core/check.py`）：

- 字体、字号、行距、对齐、首行缩进、段前段后、加粗
- 整篇页边距、有没有页码域、页数是否超上限
- 段内混排字体（多半是复制粘贴带进来的）

文字层（`core/text_check.py`）：半角标点、全角数字字母、多余空格、重复标点、省略号写法、引号／单引号／书名号／括号是否成对、年份写成中文数字或两位（如「二〇二六年」「26年」）、常见错词（截止／截至、按装／安装、予会／与会…）、一级二级序号是否连续、**数字用法（数字范围要用浪纹线「3～5」、半角 `3-5` 和 `3~5` 都会提示、百分号前不留空格）**、**「的/地/得」常见搭配错用**（做的很好→做得很好、紧密的团结→紧密地团结）。

要素层（`core/doc_check.py`）：标题有没有、标题末尾有没有句号、**缺主送机关**（顶格、机关名后加冒号；公告/通告/命令这类文种不查）、有署名没日期或有日期没署名、成文日期是否用阿拉伯数字且格式对、日期有没有跑到署名前面、附件说明是否在正文之后署名之前、抄送和印发机关是否落在版记位置、有没有正文。

版面层（`core/layout_check.py`，勾选后才会跑，要转 PDF）：文字越左右边距、压到页脚区、图片出界、表格超宽、末页孤行。

## 一键处理（批量流水线）

「更多 ▾」→ **一键处理（体检 + 排版 + 导 PDF + 报告）…**：选一个输出文件夹，一次跑完

1. 逐份体检（并行）
2. 按当前格式标准各排一份 `原名-排版后.docx`
3. 逐份导出 PDF（走 WPS/Office，串行，避免抢 COM）
4. 生成一份 `体检汇总.docx`

原件一律不动，结果全在新文件夹里。跑完自动打开那个文件夹，失败的会列出来。

## 版面预览

右栏是版面预览，用转换引擎把 docx 转成 PDF 再渲染成图，看到的就是真正落版的样子：

- **当前原件**：这份文件现在长什么样（不做任何修改）
- **排版效果**（默认）：按当前格式标准排完长什么样（排到缓存里的副本，不动原文件）

改格式标准、改角色、换文件都会自动重渲染（防抖 900ms，后台线程跑，不卡界面）。转一次 PDF 大约 1~2 秒，页图渲染只要 0.02 秒。

## 设置会记住什么

关窗口时把这几项写进 `%LOCALAPPDATA%\公文排版工具\prefs.json`（Linux 在 `~/.local/share/…`），下次打开照旧：

| 记住的 | 说明 |
| --- | --- |
| 界面密度 | 常规 / 紧凑 |
| 格式标准 | 上次选的哪套，比如「新闻稿/通讯」 |
| 窗口大小和位置 | 最大化也记；换了更小的屏幕会自动夹回可用桌面 |
| 格式转换：目标格式、输出目录 | 打开就是上次那套 |
| OCR：成品类型、输出目录 | 同上 |

想全部恢复默认，删掉那个 `prefs.json` 就行（不会影响你自己的格式标准）。

## 命令行

```bat
python -m gwtyper check  文件.docx -p gb9704 -o 报告.txt
python -m gwtyper format 文件.docx -p gb9704
python -m gwtyper learn  参照件.docx -o rules/自定义.yaml
```

`-p` 可以是规则名，也可以是 yaml 的绝对路径。

## 结构

```
gwtyper/
  core/
    model.py        角色枚举 + 段落/格式/问题数据类（唯一的中间层）
    parse_docx.py   读 docx → Doc（run 级字体、页边距、缩进、行距）
    classify.py     段落角色打分、置信度、人工覆盖
    check.py        文本层比对
    text_check.py   标点/错词/序号校对
    layout_check.py 转 PDF 看真实落版
    apply.py        把规则写回 docx（含页码域、表格字体）
    learn.py        参照件反推规则
    report.py       体检报告生成
    rules.py        规则集加载/合并
    fonts.py        本机字体是否齐备
    sysenv.py       平台差异（目录、打开文件、路径）
    office.py       转换引擎（COM → WPS → LibreOffice）
    pdfmod.py       PyMuPDF 导入兼容（pymupdf / fitz 两个模块名）
    fontfind.py     字体名 → 实际文件（含 fontTools 子集化）
    pdfwrite.py     纯 Python 排版：docx/xlsx/txt → PDF，不依赖办公软件
    doc_source.py   .doc/.wps/.rtf → docx（转换 + 缓存）
    preview.py      版面预览（原件/按规则 → PDF → 页图）
    formats.py      格式转换的目标矩阵（哪种源能转成哪种）
    convert.py      格式转换实现（PDF/Word/Excel/PPT/文本/CSV/图片）
    pdf_engine.py   PDF 页数、拆页、渲图、抽文字、合并
    ocr.py          RapidOCR 引擎 + 双层 PDF 合成
    ocr_text.py     识别结果 → txt / Word / 双层 PDF / 表格
    doc_check.py    公文要素检查（标题/主送/落款/日期/附件/抄送）
    wordjump.py     检查结果双击定位到 Word 段落
    villages.py     村名单与材料上交台账（名单、预设、Excel 解析、勾选状态、文件名匹配）
    scanim.py       扫描件摆正 + 去黑边
    cli.py          命令行
  ui/
    theme.py        设计令牌与 ttk 样式（配色、字号、DPI）
    check.py        自绘勾选框（clam 主题会把 ✓ 画成 ✗）
    rules_editor.py 规则集可视化编辑
    preview_pane.py 预览画布（滚动、缩放）
    convert_pane.py 格式转换页签
    ocr_pane.py     OCR 识别页签（含识别结果编辑窗）
    tools_pane.py   办公小工具页签（重命名/字数/扫描件/表格）
    _tools_extra.py 办公小工具的扫描件与表格生成部分
    checklist_pane.py 材料上交页签
  rules/            gb9704 / report / news（内置只读；用户目录同名可覆盖）
  packaging/        entry.py + AppRun + .desktop（打包用）
  scripts/          build_appimage.sh（full/thin）+ build_arm64_cross.sh（arm64 交叉构建）
  build_win.spec    PyInstaller 配置（Windows，单文件 exe）
  build_linux.spec  PyInstaller 配置（Linux，onedir + 自带运行时）
  app.py            主界面
  main.py           入口（GUI / CLI）
```

## 已知限制

- 排版只吃 `.docx`、`.doc`、`.wps`、`.rtf`（后几种靠转换引擎转 docx：Windows 用 WPS/Office，其他平台优先 WPS、退 LibreOffice）
- **Linux 上已不再使用 WPS**（12.1.2 实测：`--convert-to` 返回 0 但不产出文件，
  二进制里也搜不到该参数，还会常驻 `wpscloudsvr` 后台进程）。
  Linux 包的引擎顺序是：**包内 LibreOffice → 系统 LibreOffice → 纯 Python**。
  Windows 上的 WPS 走 COM，一直正常
- 版面检查与预览依赖转换引擎，没有就会跳过并明确提示
- 缺字体（方正小标宋、仿宋_GB2312 等）会被系统静默替换——界面右下角会提示缺哪个。
  Linux 上尤其明显：没有那四款字体时版面会略走形（实测超出右边距 0.57cm），装上就恢复
- AppImage 分两种：`full`（70MB，自带 Python 运行时，干净机器直接跑）、`thin`（1MB，要先装依赖）
- 转换引擎（WPS/LibreOffice）永远走系统那份，不打进包里——几百兆的办公套件塞进去不划算，
  而且和用户已有的文档环境保持一致更可靠
- PDF 转 Word（`pdf2docx`）对扫描件无能为力，纯图片 PDF 只能得到空白或图片版
- Windows 上 Excel/PPT 转 PDF 走的是 `ExportAsFixedFormat` / `Presentations.SaveAs`，
  **转换时会短暂弹出 WPS 窗口**（WPS 演示不允许完全后台打开），转完自动关掉
- **COM 转换有 120 秒超时**：WPS 开着十几个窗口或有看不见的对话框时会无限期卡住，
  超时后自动改用内置/系统 LibreOffice，并提示「关掉多余的 WPS 再试」。
  可用环境变量 `GWTYPER_COM_TIMEOUT` 调整秒数
- WPS 的 `Quit()` 有时不回收进程，反复转换会留下没有窗口的 `wps.exe`/`et.exe`。
  这些是自动化实例，可以安全结束；有窗口标题的那个才是你自己打开着的文档



## 3.1 新增（模板库 / 批量处理 / 名单比对 / 按体积压图）

| 页 | 干什么 | 说明 |
|---|---|---|
| 公文模板库 | 通知、请示、报告、函、会议纪要、情况汇报、**证明** | 左边选模板、右边填空，直接出标准版式 .docx（标题方正小标宋 22 磅、正文仿宋_GB2312 三号、固定行距 28 磅、首行缩进 2 字） |
| 批量加水印 | PDF 加斜向灰色水印文字 | 只加一层可见文字，不改原内容；适合「内部资料 注意保存」这类 |
| 敏感信息脱敏 | PDF / Word / Excel 里的手机号、身份证号打码 | 手机号 → `138****5678`，身份证保留前 6 后 4；默认输出副本 |
| 数据汇总 | 多张同结构 Excel/CSV 合并成一张 | 自动对齐表头、最后一列写「来源文件」 |
| 打印体检 | 查排版前能不能直接打印 | 版心越界、字号过小、行距过密、图片分辨率不足等 |
| 名单比对 | 两份名单找差异 | 只在 A、只在 B、两边都有、各自重复项 |
| 图片压缩 | 新增「按目标体积」模式 | 先降画质，画质到底还超标就自动缩尺寸，尽量压到指定 KB 以内 |

## 3.0 竖排菜单（按类别分组）

左侧一条竖菜单，右侧是内容区，不再横排在顶上。分组：

| 分组 | 页 |
|---|---|
| 公文处理 | 公文排版、公文模板库 |
| 文档处理 | 格式转换、OCR 识别 |
| 办公小工具 | 改名与统计、图片压缩（可按长边或按目标体积）、文件查重 |
| PDF 工具 | PDF 工具箱（加密 / 去密码 / 瘦身 / 按页提取 / Word 导 PDF 带书签） |
| 批量处理 | 批量加水印、敏感信息脱敏、数据汇总、打印体检 |
| 材料管理 | 材料上交（含生成通报 Word、按台账归档）、名单比对 |

- **生成通报 Word**：按台账出《材料上报情况通报》，应交/已交/未交数量 + 未交名单 + 截止时间，符合公文格式。
- **按台账归档**：把收到的材料按「单位名」建子文件夹复制进去，没认出来的单列。
- **图片压缩**：长边默认 1600、质量 75，只出副本，原图不动。
- **文件查重**：字节完全相同、或 docx/pdf/txt 正文完全相同（另存过也算）都归一组，可导出 CSV。
- **PDF 工具箱**：加密用 AES-256；提取页支持 `1-3,5`；Word 导 PDF 让 Word 按标题生成书签，WPS 不支持时自动退回普通导出。

2.5 横排菜单版留在 `gwtyper/dist/旧版本/公文工具箱-2.5-横版菜单.exe`，两边都能用。
