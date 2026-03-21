const statusNode = document.getElementById('status');
const pdfButton = document.getElementById('pdfButton');
const docxRenderedButton = document.getElementById('docxRenderedButton');
const docxFastButton = document.getElementById('docxFastButton');
const langButtons = Array.from(document.querySelectorAll('.lang-btn'));
const brandNode = document.getElementById('brand');
const titleNode = document.getElementById('title');
const leadNode = document.getElementById('lead');
const docxRenderedNoteNode = document.getElementById('docxRenderedNote');
const docxFastNoteNode = document.getElementById('docxFastNote');
const hintNode = document.getElementById('hint');

const LOCALE_STORAGE_KEY = 'cgptExportLocale';

const TEXT = {
  zh: {
    appName: 'chatgpt export to word/pdf',
    brand: 'chatgpt export to word/pdf',
    title: '导出当前对话',
    lead: '将当前 ChatGPT 对话导出为 PDF 或 Word。Word 提供公式保真模式和快速 LaTeX 模式。',
    buttons: {
      pdf: '导出 PDF',
      docxRendered: '导出 Word（公式保真）',
      docxFast: '快速导出 Word（LaTeX）',
    },
    notes: {
      docxRendered: '较慢；优先保留 ChatGPT 页面中的公式外观。',
      docxFast: '较快；公式会保留为 <code>$...$</code> / <code>$$...$$</code> 文本。',
    },
    hint: 'PDF 会打开新标签页并尝试自动打印；若浏览器拦截弹窗，请允许 chatgpt.com 打开新窗口。',
    status: {
      idle: '请先打开 chatgpt.com 的某个对话页面。',
      exportingDocxRendered: '正在提取聊天并生成 Word（公式保真）…',
      exportingDocxFast: '正在提取聊天并生成 Word（快速）…',
      openingPdf: '正在打开 PDF 打印视图…',
      docxRenderedReady: 'Word（公式保真）已开始下载。',
      docxFastReady: 'Word（快速）已开始下载。',
      pdfOpened: 'PDF 视图已打开。若浏览器没有自动弹出打印框，请在新标签页点“打印 / 保存为 PDF”。',
    },
    errors: {
      noActiveTab: '未找到活动标签页。',
      notChatGptTab: '请先打开 ChatGPT 对话页面。',
      extractFailed: '未能从当前页面提取聊天内容。',
      actionFailed: '操作失败。',
      docxLibrary: 'Word 导出库未正确加载。',
      docxFailed: '导出 Word 失败。',
      pdfFailed: '导出 PDF 失败。',
    },
    roles: {
      user: '用户',
      assistant: '助手',
      system: '系统',
      message: '消息',
    },
    doc: {
      fallbackTitle: 'ChatGPT 对话',
      exportedFrom: '导出自 ChatGPT',
      time: '时间：',
      source: '来源：',
      messages: '消息数：',
      wordMode: 'Word 模式：',
      wordModeRendered: '公式保真',
      wordModeFast: '快速 LaTeX',
      mathRendered: '公式：优先导出为渲染后的图像',
      mathFast: '公式：保留为 LaTeX 文本',
    },
  },
  en: {
    appName: 'chatgpt export to word/pdf',
    brand: 'chatgpt export to word/pdf',
    title: 'Export current conversation',
    lead: 'Export the current ChatGPT conversation to PDF or Word. Word includes a formula-friendly mode and a fast LaTeX mode.',
    buttons: {
      pdf: 'Export PDF',
      docxRendered: 'Export Word (formula-friendly)',
      docxFast: 'Export Word (fast LaTeX)',
    },
    notes: {
      docxRendered: 'Slower; tries to preserve the rendered formula appearance from ChatGPT.',
      docxFast: 'Faster; formulas stay as <code>$...$</code> / <code>$$...$$</code> text.',
    },
    hint: 'PDF opens in a new tab and tries to print automatically. If your browser blocks pop-ups, allow chatgpt.com to open a new window.',
    status: {
      idle: 'Open a ChatGPT conversation page first.',
      exportingDocxRendered: 'Extracting the conversation and building Word (formula-friendly)…',
      exportingDocxFast: 'Extracting the conversation and building Word (fast)…',
      openingPdf: 'Opening the PDF print view…',
      docxRenderedReady: 'Word (formula-friendly) download has started.',
      docxFastReady: 'Word (fast) download has started.',
      pdfOpened: 'The PDF view is open. If the print dialog did not appear automatically, use Print / Save as PDF in the new tab.',
    },
    errors: {
      noActiveTab: 'No active tab was found.',
      notChatGptTab: 'Open a ChatGPT conversation page first.',
      extractFailed: 'Failed to extract the conversation from the current page.',
      actionFailed: 'Action failed.',
      docxLibrary: 'The Word export library did not load correctly.',
      docxFailed: 'Failed to export Word.',
      pdfFailed: 'Failed to export PDF.',
    },
    roles: {
      user: 'User',
      assistant: 'Assistant',
      system: 'System',
      message: 'Message',
    },
    doc: {
      fallbackTitle: 'ChatGPT Conversation',
      exportedFrom: 'Exported from ChatGPT',
      time: 'Time:',
      source: 'Source:',
      messages: 'Messages:',
      wordMode: 'Word mode:',
      wordModeRendered: 'Formula-friendly',
      wordModeFast: 'Fast LaTeX',
      mathRendered: 'Math: formulas exported as rendered images when possible',
      mathFast: 'Math: formulas kept as LaTeX text',
    },
  },
};

const DOCX_CSS = `
  body {
    font-family: "Segoe UI", Inter, Arial, sans-serif;
    color: #0f172a;
    line-height: 1.6;
    margin: 0;
    padding: 0;
    background: #ffffff;
  }
  .page {
    max-width: 920px;
    margin: 0 auto;
    padding: 0;
  }
  h1 {
    margin: 0 0 10px;
    font-size: 28px;
    line-height: 1.2;
  }
  .meta {
    margin-bottom: 28px;
    color: #64748b;
    font-size: 12px;
  }
  .message {
    margin: 0 0 26px;
    page-break-inside: avoid;
  }
  .role {
    display: inline-block;
    padding: 5px 10px;
    margin-bottom: 10px;
    border-radius: 999px;
    background: #e2e8f0;
    color: #334155;
    font-size: 12px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.04em;
  }
  .body {
    border: 1px solid #e2e8f0;
    border-radius: 14px;
    padding: 16px 18px;
  }
  .body p,
  .body ul,
  .body ol,
  .body blockquote,
  .body pre,
  .body table,
  .body figure,
  .body h1,
  .body h2,
  .body h3,
  .body h4,
  .body h5,
  .body h6 {
    margin: 0.8em 0;
  }
  .body h1,
  .body h2,
  .body h3,
  .body h4,
  .body h5,
  .body h6 {
    line-height: 1.3;
  }
  .body pre {
    background: #0f172a;
    color: #e2e8f0;
    padding: 14px;
    border-radius: 12px;
    white-space: pre-wrap;
    overflow-wrap: anywhere;
  }
  .body code {
    font-family: Consolas, "SFMono-Regular", Menlo, monospace;
    font-size: 0.92em;
  }
  .body :not(pre) > code {
    padding: 2px 6px;
    border-radius: 6px;
    background: #eff6ff;
    color: #0f172a;
  }
  .body .cgpt-math-inline {
    font-family: Consolas, "SFMono-Regular", Menlo, monospace;
    background: #f8fafc;
  }
  .body .cgpt-math-block {
    font-family: Consolas, "SFMono-Regular", Menlo, monospace;
    white-space: pre-wrap;
    overflow-wrap: anywhere;
    background: #f8fafc;
    color: #0f172a;
    border: 1px solid #e2e8f0;
  }
  .body .cgpt-math-image {
    max-width: 100%;
    height: auto;
  }
  .body .cgpt-math-inline-image {
    display: inline-block;
    vertical-align: middle;
  }
  .body .cgpt-math-block-image-wrap {
    text-align: center;
    margin: 0.8em 0;
  }
  .body .cgpt-math-display-image {
    display: inline-block;
  }
  .body blockquote {
    padding-left: 12px;
    margin-left: 0;
    border-left: 4px solid #cbd5e1;
    color: #475569;
  }
  .body table {
    width: 100%;
    border-collapse: collapse;
    table-layout: fixed;
  }
  .body th,
  .body td {
    border: 1px solid #cbd5e1;
    padding: 8px 10px;
    text-align: left;
    vertical-align: top;
  }
  .body img {
    max-width: 100%;
    height: auto;
  }
  .body hr {
    border: 0;
    border-top: 1px solid #cbd5e1;
    margin: 18px 0;
  }
`;

let currentLocale = detectInitialLocale();
let currentStatus = {
  key: 'status.idle',
  params: {},
  raw: null,
  isError: false,
};

function detectInitialLocale() {
  try {
    const saved = localStorage.getItem(LOCALE_STORAGE_KEY);
    if (saved && TEXT[saved]) return saved;
  } catch (_error) {
    // ignored
  }
  const browserLocale = String(navigator.language || '').toLowerCase();
  return browserLocale.startsWith('zh') ? 'zh' : 'en';
}

function interpolate(template, params = {}) {
  return String(template).replace(/\{\{\s*(\w+)\s*\}\}/g, (_match, key) => {
    return Object.prototype.hasOwnProperty.call(params, key) ? String(params[key]) : '';
  });
}

function t(path, params = {}) {
  const bundle = TEXT[currentLocale] || TEXT.zh;
  const value = String(path || '').split('.').reduce((acc, part) => (acc && acc[part] != null ? acc[part] : undefined), bundle);
  if (typeof value === 'string') {
    return interpolate(value, params);
  }
  return value;
}

function renderStatus() {
  const message = currentStatus.raw != null
    ? currentStatus.raw
    : t(currentStatus.key, currentStatus.params);
  statusNode.textContent = message;
  statusNode.classList.toggle('error', Boolean(currentStatus.isError));
}

function setStatusKey(key, params = {}, isError = false) {
  currentStatus = { key, params, raw: null, isError };
  renderStatus();
}

function setStatusText(message, isError = false) {
  currentStatus = { key: null, params: {}, raw: String(message || ''), isError };
  renderStatus();
}

function setLocale(locale) {
  if (!TEXT[locale]) return;
  currentLocale = locale;
  try {
    localStorage.setItem(LOCALE_STORAGE_KEY, locale);
  } catch (_error) {
    // ignored
  }
  renderUi();
}

function renderUi() {
  document.documentElement.lang = currentLocale === 'zh' ? 'zh-CN' : 'en';
  document.title = t('appName');
  brandNode.textContent = t('brand');
  titleNode.textContent = t('title');
  leadNode.textContent = t('lead');
  pdfButton.textContent = t('buttons.pdf');
  docxRenderedButton.textContent = t('buttons.docxRendered');
  docxFastButton.textContent = t('buttons.docxFast');
  docxRenderedNoteNode.innerHTML = t('notes.docxRendered');
  docxFastNoteNode.innerHTML = t('notes.docxFast');
  hintNode.innerHTML = t('hint');
  langButtons.forEach((button) => {
    button.classList.toggle('active', button.dataset.lang === currentLocale);
  });
  renderStatus();
}

function setBusy(busy) {
  [pdfButton, docxRenderedButton, docxFastButton, ...langButtons].forEach((button) => {
    button.disabled = busy;
  });
}

function cleanTitle(rawTitle) {
  const title = String(rawTitle || '').replace(/\s*[-|·]\s*ChatGPT\s*$/i, '').trim();
  return title || '';
}

function sanitizeFilename(name) {
  return (
    String(name || 'chatgpt-export')
      .replace(/[\\/:*?"<>|]/g, '_')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 100) || 'chatgpt-export'
  );
}

function formatRole(role, locale = currentLocale) {
  const roles = TEXT[locale]?.roles || TEXT.zh.roles;
  if (role === 'user') return roles.user;
  if (role === 'assistant') return roles.assistant;
  if (role === 'system') return roles.system;
  return roles.message;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getTimestamp() {
  const now = new Date();
  const pad = (value) => String(value).padStart(2, '0');
  return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}_${pad(now.getHours())}-${pad(now.getMinutes())}-${pad(now.getSeconds())}`;
}

async function getActiveTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  return tab;
}

function assertChatGptTab(tab) {
  if (!tab?.id || !tab.url) {
    throw new Error(t('errors.noActiveTab'));
  }
  if (!/^https:\/\/(chatgpt\.com|chat\.openai\.com)\//.test(tab.url)) {
    throw new Error(t('errors.notChatGptTab'));
  }
}

async function requestPayload(tabId, mode) {
  const response = await chrome.tabs.sendMessage(tabId, {
    type: 'CGPT_EXPORT_EXTRACT',
    mode,
  });
  if (!response?.ok) {
    throw new Error(response?.error || t('errors.extractFailed'));
  }
  return response.data;
}

async function requestAction(tabId, type) {
  const response = await chrome.tabs.sendMessage(tabId, { type });
  if (!response?.ok) {
    throw new Error(response?.error || t('errors.actionFailed'));
  }
  return response;
}

function getWordModeInfo(locale, mode) {
  const docText = TEXT[locale]?.doc || TEXT.zh.doc;
  if (mode === 'fast') {
    return {
      label: docText.wordModeFast,
      math: docText.mathFast,
      filenameSuffix: 'word-fast',
    };
  }
  return {
    label: docText.wordModeRendered,
    math: docText.mathRendered,
    filenameSuffix: 'word-rendered',
  };
}

function buildDocumentHtml(payload, options = {}) {
  const locale = options.locale === 'en' ? 'en' : 'zh';
  const wordMode = options.wordMode === 'fast' ? 'fast' : 'rendered';
  const textPack = TEXT[locale] || TEXT.zh;
  const title = cleanTitle(payload.title) || textPack.doc.fallbackTitle;
  const exportedTime = new Date(payload.extractedAt || Date.now()).toLocaleString(locale === 'zh' ? 'zh-CN' : 'en-US');
  const wordModeInfo = getWordModeInfo(locale, wordMode);

  const meta = [
    escapeHtml(textPack.doc.exportedFrom),
    `${escapeHtml(textPack.doc.time)} ${escapeHtml(exportedTime)}`,
    `${escapeHtml(textPack.doc.source)} ${escapeHtml(payload.url || '')}`,
    `${escapeHtml(textPack.doc.messages)} ${escapeHtml(payload.messageCount || payload.messages?.length || 0)}`,
    `${escapeHtml(textPack.doc.wordMode)} ${escapeHtml(wordModeInfo.label)}`,
    escapeHtml(wordModeInfo.math),
  ].join('<br>');

  const messagesHtml = (payload.messages || []).map((message) => `
    <section class="message">
      <div class="role">${escapeHtml(formatRole(message.role, locale))}</div>
      <div class="body">${message.html || ''}</div>
    </section>
  `).join('\n');

  return `<!DOCTYPE html>
<html lang="${locale === 'zh' ? 'zh-CN' : 'en'}">
  <head>
    <meta charset="UTF-8">
    <title>${escapeHtml(title)}</title>
    <style>${DOCX_CSS}</style>
  </head>
  <body>
    <div class="page">
      <h1>${escapeHtml(title)}</h1>
      <div class="meta">${meta}</div>
      ${messagesHtml}
    </div>
  </body>
</html>`;
}

async function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  try {
    await chrome.downloads.download({
      url,
      filename,
      saveAs: true,
    });
  } finally {
    setTimeout(() => URL.revokeObjectURL(url), 60_000);
  }
}

async function exportDocx(mode = 'rendered') {
  const requestMode = mode === 'fast' ? 'docx-fast' : 'docx-rendered';
  const statusKey = mode === 'fast' ? 'status.exportingDocxFast' : 'status.exportingDocxRendered';
  const successKey = mode === 'fast' ? 'status.docxFastReady' : 'status.docxRenderedReady';

  setBusy(true);
  setStatusKey(statusKey);
  try {
    const tab = await getActiveTab();
    assertChatGptTab(tab);
    const payload = await requestPayload(tab.id, requestMode);
    const html = buildDocumentHtml(payload, {
      locale: currentLocale,
      wordMode: mode,
    });
    if (!window.htmlDocx?.asBlob) {
      throw new Error(t('errors.docxLibrary'));
    }
    const blob = window.htmlDocx.asBlob(html, {
      orientation: 'portrait',
      margins: {
        top: 720,
        right: 720,
        bottom: 720,
        left: 720,
        header: 360,
        footer: 360,
        gutter: 0,
      },
    });
    const fallbackTitle = TEXT[currentLocale]?.doc?.fallbackTitle || TEXT.zh.doc.fallbackTitle;
    const suffix = getWordModeInfo(currentLocale, mode).filenameSuffix;
    const filename = `${sanitizeFilename(payload.title || fallbackTitle)}_${suffix}_${getTimestamp()}.docx`;
    await downloadBlob(blob, filename);
    setStatusKey(successKey);
  } catch (error) {
    setStatusText(error?.message || t('errors.docxFailed'), true);
  } finally {
    setBusy(false);
  }
}

async function exportPdf() {
  setBusy(true);
  setStatusKey('status.openingPdf');
  try {
    const tab = await getActiveTab();
    assertChatGptTab(tab);
    await requestAction(tab.id, 'CGPT_EXPORT_OPEN_PDF_VIEW');
    setStatusKey('status.pdfOpened');
    window.close();
  } catch (error) {
    setStatusText(error?.message || t('errors.pdfFailed'), true);
    setBusy(false);
  }
}

langButtons.forEach((button) => {
  button.addEventListener('click', () => {
    setLocale(button.dataset.lang);
  });
});

pdfButton.addEventListener('click', exportPdf);
docxRenderedButton.addEventListener('click', () => exportDocx('rendered'));
docxFastButton.addEventListener('click', () => exportDocx('fast'));

renderUi();
setStatusKey('status.idle');
