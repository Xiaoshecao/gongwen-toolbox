# 格式转换器 · 使用说明

办公文件与 PDF 互转的 Windows 小工具，单文件 EXE，双击即用，免安装。

## 一、能做什么

| 目标格式 | 可以转换的来源 |
| --- | --- |
| **PDF** | Word（doc/docx/wps/rtf）、Excel（xls/xlsx/et/csv）、PPT（ppt/pptx/pps/dps）、图片（png/jpg/bmp/tif/webp）、TXT |
| **DOCX** | PDF、doc、rtf、wps、txt |
| **XLSX** | PDF、xls、et、csv |
| **PPTX** | PDF、ppt、pps、dps |
| **TXT** | PDF、Word、Excel、PPT —— 只要文字，便于检索复制 |
| **CSV** | xlsx、xls、et（导出第一个工作表） |
| **PNG / JPG** | PDF、PPT 每页导出一张图 |

**PDF 工具箱**：合并 PDF、按页拆分、图片合成 PDF、PDF 导出图片。

## 二、怎么用

1. 双击 `格式转换器.exe`。
2. 「添加文件」或直接把文件拖进窗口；也可以「添加文件夹」批量导入。
3. 右侧点选目标格式（不支持当前文件的格式会自动置灰）。
4. 确认输出位置，点「开始转换」。
5. 列表里每个文件都会显示「完成 / 失败」，完成后自动打开输出文件夹。

## 三、转换质量

- **转 PDF、Office 之间互转**：调用本机 WPS 或 Microsoft Office，版式与原文件一致，最适合公文、报表、汇报材料。
- **PDF 转 Word / Excel**：使用内置 PDF 解析引擎（PyMuPDF + pdf2docx），段落和表格会尽量还原，复杂分栏或扫描件可能需要微调。
- **PDF 转 Excel**：先尝试识别表格；识别不到时按行提取成「页码 + 内容」两列，保证信息不丢。
- **清晰度**：图片默认 150 DPI 适合屏幕阅读；如需打印，改成 300 DPI。

## 四、注意事项

- Office 相关转换需要本机装有 **WPS 或 Microsoft Office**；没装也能用 PDF 相关功能（PDF 转 Word/Excel/PPT、合并拆分、导出图片）。
- 转换时请不要关闭本窗口，也不要手动关掉后台弹出的办公软件进程。
- 输出文件名与源文件同名，重名自动加 `(1)(2)`，**不会覆盖**已有文件。
- 失败原因会显示在状态栏；详细日志在 `%LOCALAPPDATA%\格式转换器\error.log`。

## 五、开发者信息

源码结构：

```
converter/
├── main.py              入口（异常写日志）
├── app.py               图形界面（tkinter / ttk）
├── build.spec           PyInstaller 打包配置
├── assets/              图标
├── core/
│   ├── formats.py       格式登记表与可用性判断
│   ├── office_com.py    WPS / Office COM 调用
│   ├── pdf_engine.py    纯 Python 的 PDF/Office 解析
│   └── convert.py       转换调度与策略回退
└── ui/theme.py          配色、字体、控件样式
```

重新打包：

```powershell
python -m PyInstaller --noconfirm --clean build.spec
```

产物在 `dist\格式转换器.exe`。
