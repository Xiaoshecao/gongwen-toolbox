# 第三方组件与许可证声明（THIRD-PARTY NOTICES）

本项目（公文工具箱 / Gongwen Toolbox）以 **GNU Affero 通用公共许可证第 3 版（AGPL-3.0-or-later）** 发布，
完整许可证文本见同目录 `LICENSE`，源代码见 <https://github.com/Xiaoshecao/gongwen-toolbox>。

之所以选择 AGPL-3.0：本项目使用了 **PyMuPDF**，它采用「AGPL-3.0 或 Artifex 商业许可」双许可，
按 AGPL-3.0 使用即要求整个衍生作品以 AGPL-3.0 分发并提供对应源代码。本项目源码已全部公开，
满足该要求。若需以闭源方式分发，须另行向 Artifex 购买 PyMuPDF 商业许可。

---

## 一、随程序分发的第三方组件

| 组件 | 版本 | 许可证 | 项目主页 |
|---|---|---|---|
| PyMuPDF (fitz) | 2.1 | AGPL-3.0 或 Artifex 商业许可（双许可） | https://github.com/pymupdf/PyMuPDF |
| pymupdf4llm | 2.1 | AGPL-3.0（随 PyMuPDF 分发） | https://github.com/pymupdf/pymupdf4llm |
| pymupdf-layout | 2.1 | AGPL-3.0（随 PyMuPDF 分发） | https://github.com/pymupdf/pymupdf-layout |
| pdf2docx | 0.5.13 | MIT | https://github.com/dothinking/pdf2docx |
| python-docx | 1.2.0 | MIT | https://github.com/python-openxml/python-docx |
| python-pptx | 1.0.2 | MIT | https://github.com/scanny/python-pptx |
| openpyxl | 3.1.5 | MIT | https://foss.heptapod.net/openpyxl/openpyxl |
| XlsxWriter | 3.2.9 | BSD-2-Clause | https://github.com/jmcnamara/XlsxWriter |
| lxml | 6.1.3 | BSD-3-Clause | https://lxml.de/ |
| Pillow | 12.3.0 | MIT-CMU (HPND) | https://python-pillow.org/ |
| PyYAML | 6.0.3 | MIT | https://pyyaml.org/ |
| fontTools | 4.68.0 | MIT | https://github.com/fonttools/fonttools |
| onnxruntime | 1.32.0 | MIT | https://github.com/microsoft/onnxruntime |
| RapidOCR (rapidocr-onnxruntime) | 1.4.4 | Apache-2.0 | https://github.com/RapidAI/RapidOCR |
| OpenCV (opencv-python-headless) | 5.0.0.93 | Apache-2.0 | https://opencv.org/ |
| NumPy | 2.3.5 | BSD-3-Clause | https://numpy.org/ |
| protobuf | 7.36.2 | BSD-3-Clause | https://github.com/protocolbuffers/protobuf |
| flatbuffers | 25.12.19 | Apache-2.0 | https://github.com/google/flatbuffers |
| pyclipper | 1.4.0 | MIT | https://github.com/fonttools/pyclipper |
| Shapely | 2.2.0 | BSD-3-Clause | https://github.com/shapely/shapely |
| tkinterdnd2 | 0.7.0 | MIT | https://github.com/Eliav2/tkinterdnd2 |
| pywin32（仅 Windows 版） | 312 | PSF-2.0 | https://github.com/mhammond/pywin32 |
| PyInstaller（构建工具，含引导程序） | 6.22.3 | GPL-2.0-or-later **附打包例外** | https://pyinstaller.org/ |
| pyinstaller-hooks-contrib（构建工具） | 2026.10 | Apache-2.0 | https://github.com/pyinstaller/pyinstaller-hooks-contrib |

> PyInstaller 的许可证带有「打包例外（Bootloader Exception）」：允许将本程序打包成可执行文件并以其他许可证分发，
> 该例外不要求本程序开源；此处致谢其作者。

## 二、运行时环境

| 组件 | 版本 | 许可证 | 说明 |
|---|---|---|---|
| Python | 3.12（Windows 版）/ 3.7（ARM64 版） | PSF License | 解释器随程序打包 |
| Tcl / Tk | 8.6 | BSD 风格（Tcl/Tk License） | 图形界面库随程序打包 |
| zlib / OpenSSL 等系统库 | — | zlib / Apache-2.0 | 随 Python 运行时分发 |

## 三、Linux 版内置的第三方程序与字体

| 组件 | 版本 | 许可证 | 特别说明 |
|---|---|---|---|
| LibreOffice（含 Writer/Calc/Impress） | 24.2 | **MPL-2.0** | 以未修改的官方二进制形式随包分发。MPL-2.0 为文件级 copyleft，**不影响本程序的许可**。LibreOffice 源代码可从 https://www.libreoffice.org/about-us/source-code/ 获取；许可证全文见安装目录 `LICENSE` 与 https://www.mozilla.org/MPL/2.0/ |
| 文泉驿正黑 / 文泉驿微米黑（wqy-zenhei / wqy-microhei） | — | **GPL-2.0 附字体嵌入例外** | 仅为「系统里没有任何中文字体时不出现方块」而随包附带，未做修改。许可证全文见 https://www.gnu.org/licenses/old-licenses/gpl-2.0.html 与字体包内 COPYING |

## 四、源码获取

- 本程序全部源代码：<https://github.com/Xiaoshecao/gongwen-toolbox>
- 上表各开源组件均可通过其项目主页取得对应版本的源代码
- 如需本程序打包时所使用的确切版本与构建脚本，见仓库 `gwtyper/build_win.spec`、`gwtyper/build_linux.spec`、`gwtyper/scripts/`

## 五、声明保留

本程序在分发时保留上述各组件的版权声明与许可证信息；各许可证的完整文本可从对应项目主页
或其随包文件中取得（上表已给出链接）。如发现遗漏，请提交 Issue，我们会尽快补正。
