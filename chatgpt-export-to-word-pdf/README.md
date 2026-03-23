# chatgpt export to word/pdf

Export the current ChatGPT conversation or a whole ChatGPT Project to PDF, Word, or PNG images packed into a ZIP file, open a Team checkout panel, and choose whether to export only the current chat or the whole current Project.

将当前 ChatGPT 对话或整个 ChatGPT Project 导出为 PDF、Word，或打包成 ZIP 的 PNG 图片，并支持打开 Team 绑卡面板；现在还可以选择只导出当前聊天，或导出当前 Project 下的全部聊天。

## Features

- Export the current ChatGPT conversation to PDF
- Export the current ChatGPT conversation to PNG images packed into a ZIP file
- Export the current ChatGPT conversation to Word in a formula-friendly mode
- Export the current ChatGPT conversation to Word in a fast LaTeX mode
- Export either the current chat or the whole current Project
- Open an in-page Team checkout panel (ported from the Tampermonkey script)
- Chinese / English popup language switch

## 功能

- 导出当前 ChatGPT 对话为 PDF
- 导出当前 ChatGPT 对话为打包 ZIP 的 PNG 图片
- 导出当前 ChatGPT 对话为 Word（公式保真模式）
- 导出当前 ChatGPT 对话为 Word（快速 LaTeX 模式）
- 支持选择导出当前聊天，或导出当前 Project 下的全部聊天
- 可在当前页面打开 Team 绑卡面板（移植自油猴脚本）
- 弹窗支持中英文切换

## Install

1. Open `chrome://extensions`
2. Enable Developer mode
3. Click `Load unpacked`
4. Select this folder

## 安装

1. 打开 `chrome://extensions`
2. 开启开发者模式
3. 点击 `加载已解压的扩展程序`
4. 选择当前文件夹

## Usage

1. Open a ChatGPT conversation page or a Project page
2. Click the extension icon
3. Choose the export scope: current chat or current Project
4. Choose PDF, PNG image (ZIP), Word (formula-friendly), Word (fast LaTeX), or Team checkout panel

## 使用方式

1. 打开一个 ChatGPT 对话页面或 Project 页面
2. 点击扩展图标
3. 先选择导出范围：当前聊天，或当前 Project
4. 再选择 PDF、PNG 图片（ZIP 打包）、Word（公式保真）、Word（快速 LaTeX）或 Team 绑卡面板

## Notes

- The current-chat mode keeps the previous behavior unchanged.
- The current-Project mode discovers the conversations under the project and then loads each chat page in the background one by one so it can reuse the existing DOM-based exporter.
- Project export usually takes longer than exporting a single chat, and the Project scope can now be selected from either a Project conversation page or the Project home page.
- PNG export keeps the current ChatGPT tab in the foreground while it captures the page, then packages the generated PNG files into a ZIP download. Long captures are automatically paced to stay below Chrome's screenshot quota.
- A larger browser window usually produces a sharper PNG output.
- PDF export no longer injects a `<base>` tag, which avoids the ChatGPT CSP `base-uri 'none'` warning.
- The Team button opens an in-page panel and can generate either the new or old Team checkout link.

## 说明

- “当前聊天”模式保持你原来的导出行为不变。
- “当前 Project”模式会先识别项目里的聊天列表，再在后台逐个打开聊天页面，复用现有的 DOM 导出逻辑。
- Project 导出通常会比单个聊天更慢一些；现在在 Project 聊天页和 Project 主页都可以切换到“当前 Project”。
- PNG 导出时需要保持当前 ChatGPT 标签页处于前台，扩展会在页面上短暂覆盖一个导出预览层后完成截图，并将生成的一张或多张 PNG 自动打包成 ZIP 下载。对于较长内容，扩展会自动放慢截图节奏，以避开 Chrome 的截图频率配额。
- 浏览器窗口越大，导出的 PNG 通常越清晰。
- PDF 导出不再注入 `<base>` 标签，因此不会再触发 ChatGPT 的 `base-uri 'none'` CSP 报错。
- Team 按钮会在当前页面打开一个面板，可生成新旧两种 Team 绑卡链接。
