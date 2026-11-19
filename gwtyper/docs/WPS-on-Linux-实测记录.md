# WPS on Linux 实测记录（2026-11-10）

在 WSL2 / Ubuntu 24.04 里装了官方 WPS Office 12.1.2.29183，实测我们代码对它的处理是否正确。

## 一、装的是什么

| 项 | 值 |
|---|---|
| 包 | `wps-office_12.1.2.29183.AK.preread.sw.Personal_806791_amd64.deb`（562 MB） |
| 来源 | `linux.wps.cn` 官方下载页 |
| 安装 | `dpkg -i` + `apt-get install -f`，装了 `hexdump` 之后 postinst 才过 |
| 可执行 | `/usr/bin/wps`、`/usr/bin/wpp`、`/usr/bin/et`，实体在 `/opt/kingsoft/wps-office/office6/` |

**下载有个坑**：CDN 要签名，直接 `curl` 官方链接一律 **403**。

页面源码里是这么拼的：

```js
var secrityKey = "<WPS-CDN-SIGN-KEY-已隐去>";
var uri = new URL(url).pathname;
var t = Math.floor(Date.now() / 1000);
var k = CryptoJS.MD5(secrityKey + uri + t);
url += '?t=' + t + '&k=' + k;
```

所以能这样下（已验证返回 206）：

```bash
P="/wps/download/ep/Linux2023/29183/wps-office_12.1.2.29183.AK.preread.sw.Personal_806791_amd64.deb"
K="<WPS-CDN-SIGN-KEY-已隐去>"
T=$(date +%s)
S=$(printf '%s' "$K$P$T" | md5sum | cut -d' ' -f1)
curl -fL -o wps.deb "https://wps-linux-personal.wpscdn.cn$P?t=$T&k=$S"
```

## 二、结论：WPS for Linux 没有命令行转换能力

试过这些，**一个都不产出文件**：

| 试法 | 结果 |
|---|---|
| `wps --headless --convert-to pdf --outdir DIR file.docx` | 立刻返回 0，无产物 |
| `wps --to pdf file.docx` | 同上 |
| `wps -convert pdf file.docx` | 同上 |
| `wps --version` / `-h` / `--help` | 全部 0.1 秒返回空，或直接进 GUI 挂住 |
| 查 `/usr/bin/wps` 脚本源码 | 只处理 `-quickstart`、`-t`，**没有任何 convert 相关分支** |
| `grep` office6/wps 二进制 | 找不到 `--convert-to` / `--headless` / `--outdir` 之类的字符串 |

所以：**Linux 上「优先用 WPS」这条路走不通**，不是我们没写对，是 WPS 根本没提供这个接口。
`office.py` 的处理是对的——先试 WPS，0.1 秒内判定不可用，记下来，整个进程内不再重试，直接退 LibreOffice。
界面上现在会写：「LibreOffice（已装 WPS，但这版不支持命令行转换）」，不让人误以为程序没认出 WPS。

> 补充：WPS for Windows 是另一回事。那边走 COM（`KWPS.Application`），
> `Documents.Open` + `SaveAs(FileFormat=17)` 完全可用，Windows 版一直在用这条路。

## 三、真正影响版面的是字体，不是引擎

这是本轮最有价值的发现。

### 装字体之前（只有 Noto / WenQuanYi）

- 版面体检报 **2 条**：第 1 页文字超出右边距 **0.57cm**、末页孤行只用到版心 15%
- 原因：替代字体的字宽和方正小标宋不是一回事，行就被撑出去了

### 装字体之后

把 Windows 的 `simfang.ttf`、`simkai.ttf`、`simhei.ttf`、`simsun.ttc` 拷进
`/usr/share/fonts/truetype/windows-cn/`，再加一个 fontconfig 别名文件
`/etc/fonts/conf.d/99-gongwen-alias.conf`：

```xml
<alias binding="same"><family>仿宋_GB2312</family><accept><family>FangSong</family></accept></alias>
<alias binding="same"><family>楷体_GB2312</family><accept><family>KaiTi</family></accept></alias>
<alias binding="same"><family>方正小标宋简体</family><accept><family>SimSun</family></accept></alias>
<alias binding="same"><family>黑体</family><accept><family>SimHei</family></accept></alias>
```

结果：

| 样本 | 规则集 | Windows | WSL（装字体后） |
|---|---|---|---|
| 重阳节走访慰问 | news | 必改 0，3 页 | **必改 0，3 页** |
| 述职述廉报告 | report | 必改 0，6 页 | **必改 0，6 页** |
| 文明实践机制文件 | gb9704 | 必改 0，12 页 | **必改 0，11 页** |

**版面体检从 2 条变成 0 条**，前两份页数完全一致。第三份差 1 页，是因为它有 1 个表格
（24 行），LibreOffice 和 Word 对表格行高的计算略有出入，属于引擎差异，不是缺陷。

### 顺带修的一个 bug

`fonts.missing()` 原来只查 `fc-list` 的 family 列表，而 fontconfig 别名不改这个列表——
于是明明 `fc-match 仿宋_GB2312` 能解析到 FangSong，界面仍报「缺字体」。

现在 `missing()` 会先用 `fc-match` 问一句「这个名字最终落到哪个文件」，
解析到真实字体文件就算有。Windows 侧不受影响（本来就走 GDI 枚举）。

## 四、怎么复现

```bash
# 在 WSL Ubuntu 里
sudo apt install -y python3 python3-tk python3-docx python3-yaml python3-fitz python3-pil \
                    libreoffice-writer fonts-noto-cjk

# 装字体（可选但强烈建议，否则版面会走形）
sudo mkdir -p /usr/share/fonts/truetype/windows-cn
sudo cp /mnt/c/Windows/Fonts/{simfang.ttf,simkai.ttf,simhei.ttf,simsun.ttc} \
        /usr/share/fonts/truetype/windows-cn/
# 再把上面那段 fontconfig 别名写进 /etc/fonts/conf.d/99-gongwen-alias.conf
sudo fc-cache -f

cd /mnt/d/公文处理
python3 gwtyper/scripts/check_linux.py
```

## 五、留了什么

- WSL 里的 WPS **保留着**（562MB 的 deb 已删，装好的程序还在）。
  跑 `office.engine_hint()` 就能看到它被正确识别为「装了但不支持命令行转换」。
- 四个字体和别名也留着，Linux 侧的版面现在和 Windows 基本一致。
