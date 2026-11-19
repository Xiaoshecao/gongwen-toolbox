# 参与贡献

这是一个给基层办公用的公文处理小工具，欢迎提问题和建议。

## 提问题 / 提需求

- 先搜一下现有 [Issues](https://github.com/Xiaoshecao/gongwen-toolbox/issues)，避免重复。
- 用 Issue 模板填写：说明**你用的是哪个平台**（Windows / 麒麟 V10 等 ARM64 / 其他 Linux）、版本号（程序里「关于」能看到）、以及复现步骤。
- 涉及具体公文时，请先把内容替换成「示例」，不要上传真实的单位材料、人名、电话、身份证号。

## 改代码

1. Fork 本仓库，从 `main` 拉分支。
2. 内核在 `gwtyper/core/`（纯 Python，可独立测试）；Qt 界面在 `gongwen-qt/`。
3. 提交前请跑一遍：
   ```bash
   python -m gwtyper.tests.test_core
   ```
4. 提交信息写清「改了什么、为什么」。
5. 发 Pull Request，说明测试过的平台。

## 关于代码风格

- 中文注释、中文提示语，面向使用者的大白话优先。
- 不引入新的重量级依赖；确需引入请在 PR 里说明理由和包体积影响。
- 界面改动请附截图。

## 许可

本项目以 AGPL-3.0-or-later 发布。你提交的代码默认按同一许可证授权。
