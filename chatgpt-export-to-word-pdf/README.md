# chatgpt export to word/pdf

Export the current ChatGPT conversation to PDF, Word, or PNG images, and open a Team checkout panel.

将当前 ChatGPT 对话导出为 PDF、Word 或 PNG 图片，并支持打开 Team 绑卡面板。

## Features

- Export the current ChatGPT conversation to PDF
- Export the current ChatGPT conversation to PNG images
- Export the current ChatGPT conversation to Word in a formula-friendly mode
- Export the current ChatGPT conversation to Word in a fast LaTeX mode
- Open an in-page Team checkout panel (ported from the Tampermonkey script)
- Chinese / English popup language switch

## 功能

- 导出当前 ChatGPT 对话为 PDF
- 导出当前 ChatGPT 对话为 PNG 图片
- 导出当前 ChatGPT 对话为 Word（公式保真模式）
- 导出当前 ChatGPT 对话为 Word（快速 LaTeX 模式）
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

1. Open a ChatGPT conversation page
2. Click the extension icon
3. Choose PDF, PNG image, Word (formula-friendly), Word (fast LaTeX), or Team checkout panel

## 使用方式

1. 打开一个 ChatGPT 对话页面
2. 点击扩展图标
3. 选择 PDF、PNG 图片、Word（公式保真）、Word（快速 LaTeX）或 Team 绑卡面板

## Notes

- PNG export keeps the current ChatGPT tab in the foreground while it captures the page.
- A larger browser window usually produces a sharper PNG output.
- PDF export no longer injects a `<base>` tag, which avoids the ChatGPT CSP `base-uri 'none'` warning.
- The Team button opens an in-page panel and can generate either the new or old Team checkout link.

## 说明

- PNG 导出时需要保持当前 ChatGPT 标签页处于前台，扩展会在页面上短暂覆盖一个导出预览层后完成截图。
- 浏览器窗口越大，导出的 PNG 通常越清晰。
- PDF 导出不再注入 `<base>` 标签，因此不会再触发 ChatGPT 的 `base-uri 'none'` CSP 报错。
- Team 按钮会在当前页面打开一个面板，可生成新旧两种 Team 绑卡链接。
