# 安全说明

## 支持的版本

只维护最新发布版（见 [Releases](https://github.com/Xiaoshecao/gongwen-toolbox/releases)）。

## 报告安全问题

请**不要**在公开 Issue 里贴敏感细节。请用 GitHub 的
[私密漏洞报告](https://github.com/Xiaoshecao/gongwen-toolbox/security/advisories/new)
（Security → Report a vulnerability）说明：

- 受影响的版本与平台；
- 复现步骤或样例文件（请脱敏）；
- 你判断的影响范围。

我会尽快确认。修复后会在这里和 Release 说明里致谢（如果你愿意署名）。

## 这个程序会碰哪些数据

- 全部处理都在**本机**完成，不联网、不上传文件。
- OCR 用的是随程序打包的本地模型（RapidOCR）。
- 程序会在本机数据目录保存：格式标准、界面设置、村（单位）名单与上交状态、预览缓存。
  位置：Windows `%LOCALAPPDATA%\公文排版工具`，Linux `~/.local/share/公文排版工具`。
- 「敏感信息脱敏」功能只是把文中的号码打码后另存一份副本，**不会**删除原文件。
