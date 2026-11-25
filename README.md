# 公文工具箱（Gongwen Toolbox）

面向基层办公的中文公文处理桌面工具：**公文排版体检、格式转换、OCR 识别、办公小工具、PDF 工具、材料台账**。
规则基准为 **GB/T 9704**（党政机关公文格式），所有参数都可自定义替换。

- 运行平台：Windows 10/11、Linux（x86_64、ARM64／麒麟 V10）
- 语言：Python 3（打包版自带运行时，用户无需安装 Python，也不需要装 Office）
- 界面：**4.0 起为 PySide6（Qt）**；3.x 为 tkinter（源码保留在 `gwtyper/`，功能相同）

> **4.0 是什么**：界面从 tkinter 重写为 Qt，功能与 3.5 一致，**启动更快、大文件表格不卡**（实测窗口构建 3149ms → 约 300–600ms，1 万行表格滚动 7.5ms/次）。内核一行没改，仍复用 `gwtyper/core`，规则与输出完全一致。Qt 版源码在 `gongwen-qt/`。

## 功能一览

| 页签 | 能做什么 |
|---|---|
| 公文排版 | 按 GB/T 9704 体检（字体/字号/行距/缩进/页边距/页码/要素），一键排版另存副本，排版前后差异对比 |
| 公文模板库 | 通知／请示／报告／函／会议纪要／情况汇报／证明，填表即出 Word |
| 格式转换 | docx/xlsx/pptx/pdf/图片 互转；转 PDF 可统一横竖版、加留白；多份材料合并成一册（带书签） |
| OCR 识别 | 扫描件/图片/PDF → 纯文字 / Word（按原版式）/ Word（只留文字）/ 双层 PDF / 表格 Excel；识别结果可改 |
| 批量重命名 / 字数统计 / 扫描件纠偏 / 签到表·交接单 | 办公小工具 |
| PDF 工具箱 | 加密、去密码、瘦身、按页提取、Word 导 PDF（带书签） |
| 批量加水印 / 敏感信息脱敏 / 数据汇总 / 打印体检 | 批量处理 |
| 材料上交 / 名单比对 | 按名单核对各单位交没交、自动归档、生成通报；两份名单比对差异 |

## 目录结构

```
gwtyper/            3.x 主程序（Python 包，tkinter 界面；内核仍被 4.0 复用）
  app.py            主窗口（tkinter）
  core/             内核：解析/分类/检查/排版/转换/PDF/OCR/规则
  ui/               各页签界面
  rules/            格式标准（YAML，可自定义）
  scripts/          各平台打包脚本（含 arm64 交叉构建）
  tests/            回归用例
  README.md         详细使用说明（启动方式、规则说明、打包步骤、已知限制）
gongwen-qt/         4.0 主程序（PySide6 界面）
  main.py           入口
  ui/               主题、侧边栏、关于、控件、预览
  pages/            各功能页（公文排版 / 格式转换 / OCR / 小工具 / PDF / 模板库 …）
  build_win.spec    Windows onedir 打包配置
  build_win_onefile.spec  Windows 单文件 exe（含 runtime_tmpdir，避免中文路径问题）
  build_linux.spec  Linux x86_64 打包配置
  PERF.md           与 3.x 的性能对比实测
converter/          早期的「格式转换器」（能力已整合进 gwtyper，保留源码备查）
重构方案-架构建议.md   重构的架构选型与计划（已按方案甲落地为 4.0）
```

## 快速开始（源码运行）

```bash
# Windows：双击 启动公文排版工具.bat，或
python -m gwtyper.main            # 在仓库根目录执行

# 命令行（只读体检，不改原文件）
python -m gwtyper check 某公文.docx -p gb9704
python -m gwtyper format 某公文.docx -p gb9704     # 排版另存副本
python -m gwtyper doctor                            # 自查环境（引擎/字体/OCR）
```

详细说明见 [gwtyper/README.md](gwtyper/README.md)。

## 下载现成版本

编译好的 Windows exe、Linux ARM64 便携版等，见本仓库的 **Releases** 页面（仓库本身只放源码）。

## 规则与数据文件

- 格式标准：`gwtyper/rules/*.yaml`（`gb9704` 国家标准 / `news` 新闻稿 / `report` 述职报告）
- 用户自定义标准存放在 `%LOCALAPPDATA%\公文排版工具\rules`（Linux 为 `~/.local/share/公文排版工具/rules`），与仓库里的默认值互不影响

本项目以 **GNU Affero 通用公共许可证第 3 版（AGPL-3.0-or-later）** 发布：可自由使用、修改、再分发；分发（含修改版）时须同样以 AGPL-3.0 开源并保留声明。许可证全文见 [LICENSE](LICENSE)，随包第三方组件见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。

## 下载（国内加速）

GitHub 直连慢时，在下面任一前缀后面直接拼上 release 文件地址即可（前缀 + 完整 URL）：

| 代理前缀 | 实测速度 |
|---|---|
| `https://gh-proxy.com/` | 约 2.6 MB/s |
| `https://ghfast.top/` | 约 1.5 MB/s |
| `https://edgeone.gh-proxy.com/` | 约 1.8 MB/s |
| `https://ghproxy.net/` | 可用 |

例（Windows 版）：

```
https://gh-proxy.com/https://github.com/Xiaoshecao/gongwen-toolbox/releases/download/v4.1/gongwen-toolbox-4.1-windows-x64.exe
```

当前版本 **[Releases v4.1](https://github.com/Xiaoshecao/gongwen-toolbox/releases/tag/v4.1)** 共 5 个文件：

| 系统 | 文件 | 大小 |
|---|---|---|
| Windows 10/11 x64 | `gongwen-toolbox-4.1-windows-x64.exe` | 217 MB |
| Linux x86_64 | `gongwen-toolbox-4.1-x86_64.AppImage` | 287 MB |
| Linux x86_64 | `gongwen-toolbox-4.1-linux-x86_64.tar.gz` | 318 MB |
| 麒麟 V10 / 统信 UOS（ARM64） | `gongwen-toolbox-4.1-arm64.AppImage` | 139 MB |
| 麒麟 V10 / 统信 UOS（ARM64） | `gongwen-toolbox-4.1-linux-arm64.tar.gz` | 156 MB |

> Windows 版是单文件 exe，**双击运行**，会在同目录建 `rt` 解包目录，请放在有写权限的文件夹里。Linux 版先 `chmod +x` 再双击；tar.gz 解压后运行 `运行.sh`。

### 4.1 新增了什么

- **表格排版**：给一份排好的 Excel 样表（.xlsx），学出列宽、行高、表头/正文字体字号、对齐、边框和打印页边距，再把别的表格一键套成同样的版式（值不动，只改版式）。
- **从文件学格式标准**：给一份已排好的公文，直接学出页边距、字体字号、行距、缩进，存成自己的格式标准。
- **修复**：模板生成不再需要管理员权限（默认输出目录会自动挑一个能写的位置）；输出目录统一走同样逻辑；打包补齐 `fontTools`。
