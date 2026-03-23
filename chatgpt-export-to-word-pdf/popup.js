const statusNode = document.getElementById('status');
const pdfButton = document.getElementById('pdfButton');
const imageButton = document.getElementById('imageButton');
const docxRenderedButton = document.getElementById('docxRenderedButton');
const docxFastButton = document.getElementById('docxFastButton');
const teamPayButton = document.getElementById('teamPayButton');
const scopeLabelNode = document.getElementById('scopeLabel');
const scopeCurrentButton = document.getElementById('scopeCurrentButton');
const scopeProjectButton = document.getElementById('scopeProjectButton');
const scopeNoteNode = document.getElementById('scopeNote');
const scopeButtons = [scopeCurrentButton, scopeProjectButton].filter(Boolean);
const langButtons = Array.from(document.querySelectorAll('.lang-btn'));
const brandNode = document.getElementById('brand');
const titleNode = document.getElementById('title');
const leadNode = document.getElementById('lead');
const docxRenderedNoteNode = document.getElementById('docxRenderedNote');
const docxFastNoteNode = document.getElementById('docxFastNote');
const hintNode = document.getElementById('hint');

const LOCALE_STORAGE_KEY = 'cgptExportLocale';
const SCOPE_STORAGE_KEY = 'cgptExportScope';
const IMAGE_SLICE_HEIGHT = 1800;
const IMAGE_RENDER_FRAME_WIDTH = 1080;
const PROJECT_TAB_LOAD_TIMEOUT_MS = 25_000;
const PROJECT_EXTRACT_TIMEOUT_MS = 35_000;

const TEXT = {
  zh: {
    appName: 'chatgpt export to word/pdf',
    brand: 'chatgpt export to word/pdf',
    title: '导出聊天内容',
    lead: '将当前 ChatGPT 聊天或当前 Project 导出为 PDF、Word 或 PNG 图像，也可打开 Team 绑卡面板。Word 提供公式保真模式和快速 LaTeX 模式。',
    buttons: {
      pdf: '导出 PDF',
      image: '导出图片（PNG）',
      docxRendered: '导出 Word（公式保真）',
      docxFast: '快速导出 Word（LaTeX）',
      teamPay: '打开 Team 绑卡面板',
    },
    notes: {
      docxRendered: '较慢；优先保留 ChatGPT 页面中的公式外观。',
      docxFast: '较快；公式会保留为 <code>$...$</code> / <code>$$...$$</code> 文本。',
    },
    hint: 'PDF 会打开新标签页并尝试自动打印；PNG 会打包为 ZIP 下载；Team 按钮会在当前页面弹出绑卡面板。导出图片时请保持当前 ChatGPT 标签页处于前台；长内容会自动放慢截图速度以避开浏览器配额限制，窗口越大，图片越清晰。',
    status: {
      idle: '请先打开 chatgpt.com 的某个对话或 Project 页面。',
      exportingDocxRendered: '正在提取聊天并生成 Word（公式保真）…',
      exportingDocxFast: '正在提取聊天并生成 Word（快速）…',
      openingPdf: '正在打开 PDF 打印视图…',
      exportingImage: '正在生成 PNG 图片并准备打包，请保持当前 ChatGPT 标签页可见…',
      docxRenderedReady: 'Word（公式保真）已开始下载。',
      docxFastReady: 'Word（快速）已开始下载。',
      pdfOpened: 'PDF 视图已打开。若浏览器没有自动弹出打印框，请在新标签页点“打印 / 保存为 PDF”。',
      imageReadySingle: 'PNG 压缩包已开始下载。',
      imageReadyMultiple: 'PNG 压缩包已开始下载，内含 {{count}} 张图片。',
      openingTeamPay: '正在打开 Team 绑卡面板…',
      teamPayOpened: 'Team 绑卡面板已在当前页面打开。',
    },
    errors: {
      noActiveTab: '未找到活动标签页。',
      notChatGptTab: '请先打开 ChatGPT 对话或 Project 页面。',
      extractFailed: '未能从当前页面提取聊天内容。',
      actionFailed: '操作失败。',
      docxLibrary: 'Word 导出库未正确加载。',
      docxFailed: '导出 Word 失败。',
      pdfFailed: '导出 PDF 失败。',
      imageFailed: '导出图片失败。',
      teamPayFailed: '打开 Team 绑卡面板失败。',
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
    title: 'Export chats',
    lead: 'Export the current ChatGPT chat or the current Project to PDF, Word, or PNG images, or open the Team checkout panel. Word includes a formula-friendly mode and a fast LaTeX mode.',
    buttons: {
      pdf: 'Export PDF',
      image: 'Export image (PNG)',
      docxRendered: 'Export Word (formula-friendly)',
      docxFast: 'Export Word (fast LaTeX)',
      teamPay: 'Open Team checkout panel',
    },
    notes: {
      docxRendered: 'Slower; tries to preserve the rendered formula appearance from ChatGPT.',
      docxFast: 'Faster; formulas stay as <code>$...$</code> / <code>$$...$$</code> text.',
    },
    hint: 'PDF opens in a new tab and tries to print automatically. PNG export packages the captured images into a ZIP file and keeps the current ChatGPT tab in the foreground while capturing. Long exports automatically pace screenshots to stay within the browser quota. The Team button opens an in-page checkout panel. A larger window usually produces sharper images.',
    status: {
      idle: 'Open a ChatGPT conversation or Project page first.',
      exportingDocxRendered: 'Extracting the conversation and building Word (formula-friendly)…',
      exportingDocxFast: 'Extracting the conversation and building Word (fast)…',
      openingPdf: 'Opening the PDF print view…',
      exportingImage: 'Building PNG images and preparing the ZIP package. Keep the current ChatGPT tab visible…',
      docxRenderedReady: 'Word (formula-friendly) download has started.',
      docxFastReady: 'Word (fast) download has started.',
      pdfOpened: 'The PDF view is open. If the print dialog did not appear automatically, use Print / Save as PDF in the new tab.',
      imageReadySingle: 'The PNG ZIP download has started.',
      imageReadyMultiple: 'The PNG ZIP download has started and contains {{count}} images.',
      openingTeamPay: 'Opening the Team checkout panel…',
      teamPayOpened: 'The Team checkout panel is open in the current page.',
    },
    errors: {
      noActiveTab: 'No active tab was found.',
      notChatGptTab: 'Open a ChatGPT conversation or Project page first.',
      extractFailed: 'Failed to extract the conversation from the current page.',
      actionFailed: 'Action failed.',
      docxLibrary: 'The Word export library did not load correctly.',
      docxFailed: 'Failed to export Word.',
      pdfFailed: 'Failed to export PDF.',
      imageFailed: 'Failed to export image.',
      teamPayFailed: 'Failed to open the Team checkout panel.',
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

TEXT.zh.scope = {
  label: '导出范围',
  current: '当前聊天',
  project: '当前 Project',
  noteCurrent: '保持现在的行为，只导出当前正在查看的聊天。',
  noteProjectReady: '会自动汇总当前 Project 下的全部聊天；如果当前打开的是某个聊天，也可以切回“当前聊天”。',
  noteProjectPage: '当前位于 Project 页面，可直接导出整个 Project；打开其中某个聊天后也可切回“当前聊天”。',
  noteProjectUnavailable: '当前页面不在 Project 中，因此只能导出当前聊天。',
};
Object.assign(TEXT.zh.status, {
  discoveringProject: '正在识别当前 Project 并读取聊天列表…',
  collectingProjectConversation: '正在提取 Project 聊天 {{current}}/{{total}}：{{title}}',
  preparingProjectPdf: '正在生成当前 Project 的 PDF 打印视图…',
  packingImageZip: '正在打包 PNG 压缩包…',
});
Object.assign(TEXT.zh.errors, {
  projectUnavailable: '当前页面不在 Project 中。',
  projectDiscoverFailed: '读取当前 Project 失败。',
  projectConversationFailed: '提取 Project 聊天失败。',
  popupBlocked: '浏览器拦截了新窗口。请允许扩展打开新窗口后重试。',
});
Object.assign(TEXT.zh.doc, {
  scope: '导出范围：',
  scopeCurrent: '当前聊天',
  scopeProject: '当前 Project',
  conversations: '会话数：',
  project: 'Project：',
  conversationTime: '会话导出时间：',
  conversationTitlePrefix: '会话',
  projectFallbackTitle: 'ChatGPT Project 导出',
});
TEXT.en.scope = {
  label: 'Export scope',
  current: 'Current chat',
  project: 'Current Project',
  noteCurrent: 'Keeps the existing behavior and exports only the conversation you are viewing now.',
  noteProjectReady: 'Collects all conversations under the current Project. If a conversation is open, you can still switch back to current-chat export.',
  noteProjectPage: 'You are on a Project page, so the whole Project can be exported directly. Open any conversation if you want to switch back to current-chat export.',
  noteProjectUnavailable: 'This page is not inside a Project, so only current-chat export is available.',
};
Object.assign(TEXT.en.status, {
  discoveringProject: 'Detecting the current Project and loading its conversation list…',
  collectingProjectConversation: 'Extracting Project conversation {{current}}/{{total}}: {{title}}',
  preparingProjectPdf: 'Building the current Project PDF print view…',
  packingImageZip: 'Packing the PNG ZIP archive…',
});
Object.assign(TEXT.en.errors, {
  projectUnavailable: 'This page is not inside a Project.',
  projectDiscoverFailed: 'Failed to discover the current Project.',
  projectConversationFailed: 'Failed to extract one of the Project conversations.',
  popupBlocked: 'The browser blocked the new window. Allow the extension to open a window and try again.',
});
Object.assign(TEXT.en.doc, {
  scope: 'Export scope:',
  scopeCurrent: 'Current chat',
  scopeProject: 'Current Project',
  conversations: 'Conversations:',
  project: 'Project:',
  conversationTime: 'Conversation export time:',
  conversationTitlePrefix: 'Conversation',
  projectFallbackTitle: 'ChatGPT Project Export',
});

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
  .body ul,
  .body ol {
    padding-left: 1.6em;
    padding-inline-start: 1.6em;
    list-style-position: outside;
  }
  .body ul {
    list-style-type: disc;
  }
  .body ul ul {
    list-style-type: circle;
  }
  .body ul ul ul {
    list-style-type: square;
  }
  .body ol {
    list-style-type: decimal;
  }
  .body ol ol {
    list-style-type: lower-alpha;
  }
  .body ol ol ol {
    list-style-type: lower-roman;
  }
  .body li {
    display: list-item;
    margin: 0.25em 0;
  }
  .body li::marker {
    color: inherit;
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
  .project-export .conversation {
    margin: 0 0 34px;
    padding-top: 6px;
    page-break-inside: avoid;
  }
  .project-export .conversation + .conversation {
    border-top: 2px solid #e2e8f0;
    padding-top: 28px;
  }
  .conversation-title {
    margin: 0 0 8px;
    font-size: 20px;
    line-height: 1.3;
  }
  .conversation-meta {
    margin: 0 0 18px;
    color: #64748b;
    font-size: 12px;
    line-height: 1.5;
  }
`;

let currentLocale = detectInitialLocale();
let currentScope = detectInitialScope();
let scopeContext = {
  isConversationPage: false,
  isProjectPage: false,
  isProjectConversation: false,
  projectTitle: '',
};
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

function detectInitialScope() {
  try {
    const saved = localStorage.getItem(SCOPE_STORAGE_KEY);
    if (saved === 'project' || saved === 'current') return saved;
  } catch (_error) {
    // ignored
  }
  return 'current';
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


function getScopeLabel(locale, scope) {
  const docText = TEXT[locale]?.doc || TEXT.zh.doc;
  return scope === 'project' ? docText.scopeProject : docText.scopeCurrent;
}

function renderScopeUi() {
  if (!scopeLabelNode || !scopeNoteNode) return;
  const scopeText = TEXT[currentLocale]?.scope || TEXT.zh.scope;
  const canUseCurrent = Boolean(scopeContext?.isConversationPage);
  const canUseProject = Boolean(scopeContext?.isProjectPage);

  scopeLabelNode.textContent = scopeText.label;
  if (scopeCurrentButton) {
    scopeCurrentButton.textContent = scopeText.current;
  }
  if (scopeProjectButton) {
    scopeProjectButton.textContent = scopeText.project;
  }

  if (currentScope === 'project' && !canUseProject) {
    currentScope = 'current';
  }
  if (currentScope === 'current' && !canUseCurrent && canUseProject) {
    currentScope = 'project';
  }

  if (scopeCurrentButton) {
    scopeCurrentButton.classList.toggle('active', currentScope === 'current' && canUseCurrent);
    scopeCurrentButton.disabled = !canUseCurrent;
  }
  if (scopeProjectButton) {
    scopeProjectButton.classList.toggle('active', currentScope === 'project' && canUseProject);
    scopeProjectButton.disabled = !canUseProject;
  }

  let note = scopeText.noteCurrent;
  if (canUseProject) {
    note = canUseCurrent ? scopeText.noteProjectReady : (scopeText.noteProjectPage || scopeText.noteProjectReady);
    if (scopeContext.projectTitle) {
      note += currentLocale === 'zh'
        ? `（${scopeContext.projectTitle}）`
        : ` (${scopeContext.projectTitle})`;
    }
  } else if (canUseCurrent) {
    note = scopeText.noteProjectUnavailable;
  }
  scopeNoteNode.textContent = note;
}

function setScope(scope) {
  const nextScope = scope === 'project' ? 'project' : 'current';
  if (nextScope === 'project' && !scopeContext?.isProjectPage) {
    return;
  }
  if (nextScope === 'current' && !scopeContext?.isConversationPage) {
    return;
  }
  currentScope = nextScope;
  try {
    localStorage.setItem(SCOPE_STORAGE_KEY, currentScope);
  } catch (_error) {
    // ignored
  }
  renderUi();
}

async function updateScopeContext() {
  try {
    const tab = await getActiveTab();
    if (!tab?.id || !tab.url || !/^https:\/\/(chatgpt\.com|chat\.openai\.com)\//.test(tab.url)) {
      scopeContext = {
        isConversationPage: false,
        isProjectPage: false,
        isProjectConversation: false,
        projectTitle: '',
      };
      renderScopeUi();
      return;
    }

    const response = await requestAction(tab.id, 'CGPT_EXPORT_GET_CONTEXT', {}, 'errors.actionFailed');
    scopeContext = {
      isConversationPage: Boolean(response?.data?.isConversationPage),
      isProjectPage: Boolean(response?.data?.isProjectPage || response?.data?.isProjectConversation),
      isProjectConversation: Boolean(response?.data?.isProjectConversation),
      projectTitle: String(response?.data?.projectTitle || ''),
    };
  } catch (_error) {
    scopeContext = {
      isConversationPage: false,
      isProjectPage: false,
      isProjectConversation: false,
      projectTitle: '',
    };
  }

  if (currentScope === 'project' && !scopeContext.isProjectPage) {
    currentScope = 'current';
  }
  if (currentScope === 'current' && !scopeContext.isConversationPage && scopeContext.isProjectPage) {
    currentScope = 'project';
  }
  renderScopeUi();
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
  imageButton.textContent = t('buttons.image');
  docxRenderedButton.textContent = t('buttons.docxRendered');
  docxFastButton.textContent = t('buttons.docxFast');
  teamPayButton.textContent = t('buttons.teamPay');
  docxRenderedNoteNode.innerHTML = t('notes.docxRendered');
  docxFastNoteNode.innerHTML = t('notes.docxFast');
  hintNode.innerHTML = t('hint');
  langButtons.forEach((button) => {
    button.classList.toggle('active', button.dataset.lang === currentLocale);
  });
  renderScopeUi();
  renderStatus();
}

function setBusy(busy) {
  [pdfButton, imageButton, docxRenderedButton, docxFastButton, teamPayButton, ...scopeButtons, ...langButtons].forEach((button) => {
    button.disabled = busy
      || (button === scopeCurrentButton && !scopeContext?.isConversationPage)
      || (button === scopeProjectButton && !scopeContext?.isProjectPage);
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

async function sendTabMessage(tabId, message, fallbackErrorKey = 'errors.actionFailed') {
  let lastError = null;
  for (let attempt = 0; attempt < 4; attempt += 1) {
    try {
      const response = await chrome.tabs.sendMessage(tabId, message);
      if (!response?.ok) {
        throw new Error(response?.error || t(fallbackErrorKey));
      }
      return response;
    } catch (error) {
      lastError = error;
      const messageText = String(error?.message || '');
      const retryable = /Receiving end does not exist|Could not establish connection|The message port closed before a response was received/i.test(messageText);
      if (!retryable || attempt >= 3) {
        break;
      }
      await sleep(250 * (attempt + 1));
    }
  }
  throw new Error(lastError?.message || t(fallbackErrorKey));
}

async function requestPayload(tabId, mode) {
  const response = await sendTabMessage(tabId, {
    type: 'CGPT_EXPORT_EXTRACT',
    mode,
  }, 'errors.extractFailed');
  return response.data;
}

async function requestAction(tabId, type, extra = {}, fallbackErrorKey = 'errors.actionFailed') {
  return sendTabMessage(tabId, {
    type,
    ...extra,
  }, fallbackErrorKey);
}


function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function normalizeSingleConversationPayload(payload, fallback = {}) {
  const messages = Array.isArray(payload?.messages) ? payload.messages : [];
  return {
    id: fallback.id || payload?.id || '',
    title: cleanTitle(fallback.title || payload?.title) || t('doc.fallbackTitle'),
    url: fallback.url || payload?.url || '',
    extractedAt: payload?.extractedAt || fallback.extractedAt || new Date().toISOString(),
    messageCount: Number(payload?.messageCount || messages.length || 0),
    messages,
  };
}

function normalizePayloadConversations(payload) {
  if (Array.isArray(payload?.conversations) && payload.conversations.length) {
    return payload.conversations.map((conversation, index) => normalizeSingleConversationPayload(conversation, {
      title: conversation?.title || `${t('doc.conversationTitlePrefix')} ${index + 1}`,
      url: conversation?.url || payload?.url || '',
    }));
  }
  if (payload) {
    return [normalizeSingleConversationPayload(payload)];
  }
  return [];
}

function maybeTrimStatusTitle(title) {
  const raw = String(title || '').trim();
  if (!raw) return currentLocale === 'zh' ? '未命名聊天' : 'Untitled conversation';
  return raw.length > 46 ? `${raw.slice(0, 43)}…` : raw;
}

function buildAggregatePayload(projectPlan, conversationPayloads) {
  const conversations = conversationPayloads.map((item, index) => ({
    ...normalizeSingleConversationPayload(item, {
      title: item?.title || `${t('doc.conversationTitlePrefix')} ${index + 1}`,
      url: item?.url || projectPlan?.projectUrl || '',
    }),
    index: index + 1,
  }));
  const totalMessages = conversations.reduce((sum, item) => sum + Number(item.messageCount || item.messages?.length || 0), 0);
  const projectTitle = cleanTitle(projectPlan?.projectTitle) || t('doc.projectFallbackTitle');
  return {
    scope: 'project',
    title: projectTitle,
    projectTitle,
    url: projectPlan?.projectUrl || '',
    extractedAt: new Date().toISOString(),
    conversationCount: conversations.length,
    messageCount: totalMessages,
    conversations,
  };
}

async function waitForTabComplete(tabId, timeoutMs = PROJECT_TAB_LOAD_TIMEOUT_MS) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const tab = await chrome.tabs.get(tabId);
    if (tab?.status === 'complete') {
      return tab;
    }
    await sleep(250);
  }
  throw new Error(t('errors.projectConversationFailed'));
}

async function waitForConversationPayload(tabId, mode, timeoutMs = PROJECT_EXTRACT_TIMEOUT_MS) {
  const start = Date.now();
  let lastError = null;
  while (Date.now() - start < timeoutMs) {
    try {
      return await requestPayload(tabId, mode);
    } catch (error) {
      lastError = error;
      await sleep(900);
    }
  }
  throw lastError || new Error(t('errors.projectConversationFailed'));
}

async function extractProjectConversationViaBackgroundTab(baseTab, conversation, mode, index, total) {
  setStatusKey('status.collectingProjectConversation', {
    current: index,
    total,
    title: maybeTrimStatusTitle(conversation.title),
  });

  const tempTab = await chrome.tabs.create({
    url: conversation.url,
    active: false,
    windowId: baseTab.windowId,
  });

  try {
    await waitForTabComplete(tempTab.id);
    const payload = await waitForConversationPayload(tempTab.id, mode);
    return normalizeSingleConversationPayload(payload, conversation);
  } finally {
    if (tempTab?.id) {
      try {
        await chrome.tabs.remove(tempTab.id);
      } catch (_error) {
        // ignored
      }
    }
  }
}

async function collectProjectPayload(tab, mode) {
  setStatusKey('status.discoveringProject');
  const response = await requestAction(tab.id, 'CGPT_EXPORT_LIST_CURRENT_PROJECT_CONVERSATIONS', {}, 'errors.projectDiscoverFailed');
  const projectPlan = response?.data;
  const conversations = Array.isArray(projectPlan?.conversations) ? projectPlan.conversations : [];

  if (!projectPlan?.isProjectPage || !conversations.length) {
    throw new Error(t('errors.projectUnavailable'));
  }

  const currentConversationId = String(projectPlan.currentConversationId || '');
  const payloads = [];

  for (let index = 0; index < conversations.length; index += 1) {
    const conversation = conversations[index];
    const order = index + 1;
    if (conversation.id && conversation.id === currentConversationId) {
      setStatusKey('status.collectingProjectConversation', {
        current: order,
        total: conversations.length,
        title: maybeTrimStatusTitle(conversation.title),
      });
      const payload = await requestPayload(tab.id, mode);
      payloads.push(normalizeSingleConversationPayload(payload, conversation));
      continue;
    }
    const payload = await extractProjectConversationViaBackgroundTab(tab, conversation, mode, order, conversations.length);
    payloads.push(payload);
  }

  return buildAggregatePayload(projectPlan, payloads);
}

async function collectExportPayload(tab, mode) {
  if (currentScope === 'project') {
    return collectProjectPayload(tab, mode);
  }
  return requestPayload(tab.id, mode);
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
  const conversations = normalizePayloadConversations(payload);
  const isProjectExport = payload?.scope === 'project' || conversations.length > 1;
  const title = cleanTitle(payload?.title || payload?.projectTitle) || (isProjectExport ? textPack.doc.projectFallbackTitle : textPack.doc.fallbackTitle);
  const exportedTime = new Date(payload?.extractedAt || Date.now()).toLocaleString(locale === 'zh' ? 'zh-CN' : 'en-US');
  const wordModeInfo = getWordModeInfo(locale, wordMode);
  const totalMessages = Number(payload?.messageCount || conversations.reduce((sum, item) => sum + Number(item.messageCount || item.messages?.length || 0), 0));

  const meta = [
    escapeHtml(textPack.doc.exportedFrom),
    `${escapeHtml(textPack.doc.time)} ${escapeHtml(exportedTime)}`,
    `${escapeHtml(textPack.doc.source)} ${escapeHtml(payload?.url || '')}`,
    `${escapeHtml(textPack.doc.messages)} ${escapeHtml(totalMessages)}`,
    isProjectExport ? `${escapeHtml(textPack.doc.conversations)} ${escapeHtml(payload?.conversationCount || conversations.length)}` : '',
    `${escapeHtml(textPack.doc.scope)} ${escapeHtml(isProjectExport ? textPack.doc.scopeProject : textPack.doc.scopeCurrent)}`,
    isProjectExport && payload?.projectTitle ? `${escapeHtml(textPack.doc.project)} ${escapeHtml(payload.projectTitle)}` : '',
    `${escapeHtml(textPack.doc.wordMode)} ${escapeHtml(wordModeInfo.label)}`,
    escapeHtml(wordModeInfo.math),
  ].filter(Boolean).join('<br>');

  const conversationsHtml = conversations.map((conversation, index) => {
    const messagesHtml = (conversation.messages || []).map((message) => `
      <section class="message">
        <div class="role">${escapeHtml(formatRole(message.role, locale))}</div>
        <div class="body">${message.html || ''}</div>
      </section>
    `).join('\n');

    const conversationMeta = [
      conversation.url ? `${escapeHtml(textPack.doc.source)} ${escapeHtml(conversation.url)}` : '',
      `${escapeHtml(textPack.doc.messages)} ${escapeHtml(conversation.messageCount || conversation.messages?.length || 0)}`,
      `${escapeHtml(textPack.doc.conversationTime)} ${escapeHtml(new Date(conversation.extractedAt || payload?.extractedAt || Date.now()).toLocaleString(locale === 'zh' ? 'zh-CN' : 'en-US'))}`,
    ].filter(Boolean).join('<br>');

    if (!isProjectExport) {
      return messagesHtml;
    }

    const conversationTitle = cleanTitle(conversation.title) || `${textPack.doc.conversationTitlePrefix} ${index + 1}`;
    return `
      <section class="conversation">
        <h2 class="conversation-title">${index + 1}. ${escapeHtml(conversationTitle)}</h2>
        <div class="conversation-meta">${conversationMeta}</div>
        ${messagesHtml}
      </section>
    `;
  }).join('\n');

  return `<!DOCTYPE html>
<html lang="${locale === 'zh' ? 'zh-CN' : 'en'}">
  <head>
    <meta charset="UTF-8">
    <title>${escapeHtml(title)}</title>
    <style>${DOCX_CSS}</style>
  </head>
  <body>
    <div class="page ${isProjectExport ? 'project-export' : ''}">
      <h1>${escapeHtml(title)}</h1>
      <div class="meta">${meta}</div>
      ${conversationsHtml}
    </div>
  </body>
</html>`;
}

function buildPdfWindowHtml(payload, options = {}) {
  const locale = options.locale === 'en' ? 'en' : 'zh';
  const textPack = TEXT[locale] || TEXT.zh;
  const documentHtml = buildDocumentHtml(payload, {
    locale,
    wordMode: options.wordMode === 'fast' ? 'fast' : 'rendered',
  });
  const title = cleanTitle(payload?.title || payload?.projectTitle) || textPack.doc.fallbackTitle;
  const bodyMatch = documentHtml.match(/<body>([\s\S]*)<\/body>/i);
  const innerHtml = bodyMatch ? bodyMatch[1] : documentHtml;
  const scriptUrl = chrome.runtime.getURL('pdf-window.js');
  return `<!DOCTYPE html>
<html lang="${locale === 'zh' ? 'zh-CN' : 'en'}">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${escapeHtml(title)}</title>
    <style>
      ${DOCX_CSS}
      body {
        background: #ffffff;
      }
      .pdf-toolbar {
        position: sticky;
        top: 0;
        z-index: 9999;
        display: flex;
        gap: 10px;
        align-items: center;
        justify-content: space-between;
        padding: 12px 16px;
        border-bottom: 1px solid #cbd5e1;
        background: #ffffff;
      }
      .pdf-toolbar h1 {
        margin: 0;
        font-size: 16px;
        line-height: 1.2;
      }
      .pdf-toolbar .meta {
        margin: 4px 0 0;
        color: #64748b;
        font-size: 12px;
      }
      .pdf-toolbar .actions {
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
      }
      .pdf-toolbar button {
        appearance: none;
        border: 0;
        border-radius: 999px;
        padding: 10px 14px;
        background: #0f172a;
        color: #ffffff;
        font: 700 13px/1 Inter, "Segoe UI", Arial, sans-serif;
        cursor: pointer;
      }
      @page {
        margin: 14mm 12mm;
      }
      @media print {
        .pdf-toolbar {
          display: none !important;
        }
      }
    </style>
  </head>
  <body>
    <div class="pdf-toolbar">
      <div>
        <h1>${escapeHtml(title)}</h1>
        <div class="meta">${escapeHtml(textPack.doc.scope)} ${escapeHtml(payload?.scope === 'project' ? textPack.doc.scopeProject : textPack.doc.scopeCurrent)}</div>
      </div>
      <div class="actions">
        <button type="button" id="cgpt-popup-print-btn">${locale === 'zh' ? '打印 / 保存为 PDF' : 'Print / Save as PDF'}</button>
        <button type="button" id="cgpt-popup-close-btn">${locale === 'zh' ? '关闭' : 'Close'}</button>
      </div>
    </div>
    ${innerHtml}
    <script src="${escapeHtml(scriptUrl)}"></script>
  </body>
</html>`;
}

function openPdfWindow(html) {
  const popup = window.open('', '_blank');
  if (!popup) {
    throw new Error(t('errors.popupBlocked'));
  }
  popup.document.open();
  popup.document.write(html);
  popup.document.close();
  popup.focus();
  return popup;
}

async function downloadBlob(blob, filename, options = {}) {
  const url = URL.createObjectURL(blob);
  try {
    await chrome.downloads.download({
      url,
      filename,
      saveAs: options.saveAs ?? true,
    });
  } finally {
    setTimeout(() => URL.revokeObjectURL(url), 60_000);
  }
}

const ZIP_NAME_ENCODER = new TextEncoder();
const ZIP_CRC32_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let index = 0; index < 256; index += 1) {
    let value = index;
    for (let bit = 0; bit < 8; bit += 1) {
      value = (value & 1) ? (0xEDB88320 ^ (value >>> 1)) : (value >>> 1);
    }
    table[index] = value >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let value = 0xFFFFFFFF;
  for (const byte of bytes) {
    value = ZIP_CRC32_TABLE[(value ^ byte) & 0xFF] ^ (value >>> 8);
  }
  return (value ^ 0xFFFFFFFF) >>> 0;
}

function toDosDateTime(date) {
  const year = Math.max(1980, date.getFullYear());
  const dosTime = ((date.getHours() & 0x1F) << 11) | ((date.getMinutes() & 0x3F) << 5) | Math.floor(date.getSeconds() / 2);
  const dosDate = (((year - 1980) & 0x7F) << 9) | (((date.getMonth() + 1) & 0x0F) << 5) | (date.getDate() & 0x1F);
  return { dosTime, dosDate };
}

async function createStoredZip(entries) {
  const now = new Date();
  const localChunks = [];
  const centralChunks = [];
  let offset = 0;

  for (const entry of entries) {
    const fileName = String(entry?.name || 'file').replace(/\\/g, '/');
    const nameBytes = ZIP_NAME_ENCODER.encode(fileName);
    const dataBytes = entry?.data instanceof Uint8Array ? entry.data : new Uint8Array(entry?.data || []);
    const checksum = crc32(dataBytes);
    const { dosTime, dosDate } = toDosDateTime(entry?.date instanceof Date ? entry.date : now);

    const localHeader = new Uint8Array(30 + nameBytes.length);
    const localView = new DataView(localHeader.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
    localView.setUint16(6, 0x0800, true);
    localView.setUint16(8, 0, true);
    localView.setUint16(10, dosTime, true);
    localView.setUint16(12, dosDate, true);
    localView.setUint32(14, checksum, true);
    localView.setUint32(18, dataBytes.length, true);
    localView.setUint32(22, dataBytes.length, true);
    localView.setUint16(26, nameBytes.length, true);
    localView.setUint16(28, 0, true);
    localHeader.set(nameBytes, 30);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    const centralView = new DataView(centralHeader.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 20, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint16(8, 0x0800, true);
    centralView.setUint16(10, 0, true);
    centralView.setUint16(12, dosTime, true);
    centralView.setUint16(14, dosDate, true);
    centralView.setUint32(16, checksum, true);
    centralView.setUint32(20, dataBytes.length, true);
    centralView.setUint32(24, dataBytes.length, true);
    centralView.setUint16(28, nameBytes.length, true);
    centralView.setUint16(30, 0, true);
    centralView.setUint16(32, 0, true);
    centralView.setUint16(34, 0, true);
    centralView.setUint16(36, 0, true);
    centralView.setUint32(38, 0, true);
    centralView.setUint32(42, offset, true);
    centralHeader.set(nameBytes, 46);

    localChunks.push(localHeader, dataBytes);
    centralChunks.push(centralHeader);
    offset += localHeader.length + dataBytes.length;
  }

  const centralSize = centralChunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const endRecord = new Uint8Array(22);
  const endView = new DataView(endRecord.buffer);
  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(4, 0, true);
  endView.setUint16(6, 0, true);
  endView.setUint16(8, entries.length, true);
  endView.setUint16(10, entries.length, true);
  endView.setUint32(12, centralSize, true);
  endView.setUint32(16, offset, true);
  endView.setUint16(20, 0, true);

  return new Blob([...localChunks, ...centralChunks, endRecord], { type: 'application/zip' });
}

async function createStoredZipFromBlobs(entries) {
  const normalized = [];
  for (const entry of entries) {
    const blob = entry?.blob instanceof Blob ? entry.blob : new Blob([entry?.blob || '']);
    normalized.push({
      name: entry?.name || 'file',
      data: new Uint8Array(await blob.arrayBuffer()),
      date: entry?.date,
    });
  }
  return createStoredZip(normalized);
}


function waitForImageElement(image) {
  return new Promise((resolve) => {
    if (!image || image.complete) {
      resolve();
      return;
    }
    const done = () => resolve();
    image.addEventListener('load', done, { once: true });
    image.addEventListener('error', done, { once: true });
  });
}

async function waitForFrameDocumentReady(iframe) {
  const doc = iframe?.contentDocument;
  if (!doc) return;

  if (doc.readyState !== 'complete') {
    await new Promise((resolve) => {
      let settled = false;
      const finish = () => {
        if (settled) return;
        settled = true;
        resolve();
      };
      iframe.addEventListener('load', finish, { once: true });
      setTimeout(finish, 1500);
    });
  }

  try {
    if (doc.fonts?.ready) {
      await doc.fonts.ready;
    }
  } catch (_error) {
    // ignored
  }

  const images = Array.from(doc.images || []);
  await Promise.all(images.map(waitForImageElement));

  await new Promise((resolve) => {
    const view = doc.defaultView || window;
    view.requestAnimationFrame(() => view.requestAnimationFrame(resolve));
  });
}

async function renderHtmlInHiddenFrame(html) {
  const iframe = document.createElement('iframe');
  iframe.setAttribute('aria-hidden', 'true');
  iframe.style.position = 'fixed';
  iframe.style.left = '-20000px';
  iframe.style.top = '0';
  iframe.style.width = `${IMAGE_RENDER_FRAME_WIDTH}px`;
  iframe.style.height = '1200px';
  iframe.style.opacity = '0';
  iframe.style.pointerEvents = 'none';
  iframe.style.border = '0';
  document.body.appendChild(iframe);

  const doc = iframe.contentDocument;
  doc.open();
  doc.write(html);
  doc.close();
  await waitForFrameDocumentReady(iframe);
  return iframe;
}

function getNodeRenderSize(node) {
  const rect = node.getBoundingClientRect();
  return {
    width: Math.max(1, Math.ceil(rect.width || node.scrollWidth || node.clientWidth || 1)),
    height: Math.max(1, Math.ceil(node.scrollHeight || rect.height || node.clientHeight || 1)),
  };
}

function copyComputedStyles(sourceNode, targetNode) {
  if (sourceNode?.nodeType !== Node.ELEMENT_NODE || targetNode?.nodeType !== Node.ELEMENT_NODE) return;
  const view = sourceNode.ownerDocument?.defaultView || window;
  const computed = view.getComputedStyle(sourceNode);
  if (computed.cssText) {
    targetNode.style.cssText = computed.cssText;
  } else {
    Array.from(computed).forEach((property) => {
      targetNode.style.setProperty(property, computed.getPropertyValue(property), computed.getPropertyPriority(property));
    });
  }
}

function cloneNodeWithInlineStyles(sourceNode) {
  const clone = sourceNode.cloneNode(true);

  const walk = (source, target) => {
    if (source?.nodeType === Node.ELEMENT_NODE && target?.nodeType === Node.ELEMENT_NODE) {
      copyComputedStyles(source, target);
    }
    const sourceChildren = Array.from(source?.childNodes || []);
    const targetChildren = Array.from(target?.childNodes || []);
    for (let index = 0; index < Math.min(sourceChildren.length, targetChildren.length); index += 1) {
      walk(sourceChildren[index], targetChildren[index]);
    }
  };

  walk(sourceNode, clone);
  return clone;
}

function buildNodeSliceSvg(node, startY, sliceHeight) {
  const { width } = getNodeRenderSize(node);
  const clone = cloneNodeWithInlineStyles(node);

  if (clone?.nodeType === Node.ELEMENT_NODE) {
    clone.style.margin = '0';
    clone.style.maxWidth = 'none';
    clone.style.width = `${width}px`;
    clone.style.boxSizing = 'border-box';
    clone.style.position = 'relative';
    clone.style.transform = `translateY(-${startY}px)`;
    clone.style.transformOrigin = 'top left';
  }

  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  svg.setAttribute('width', String(width));
  svg.setAttribute('height', String(sliceHeight));
  svg.setAttribute('viewBox', `0 0 ${width} ${sliceHeight}`);

  const background = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  background.setAttribute('x', '0');
  background.setAttribute('y', '0');
  background.setAttribute('width', String(width));
  background.setAttribute('height', String(sliceHeight));
  background.setAttribute('fill', '#ffffff');
  svg.appendChild(background);

  const foreignObject = document.createElementNS('http://www.w3.org/2000/svg', 'foreignObject');
  foreignObject.setAttribute('x', '0');
  foreignObject.setAttribute('y', '0');
  foreignObject.setAttribute('width', '100%');
  foreignObject.setAttribute('height', '100%');

  const wrapper = document.createElementNS('http://www.w3.org/1999/xhtml', 'div');
  wrapper.setAttribute('xmlns', 'http://www.w3.org/1999/xhtml');
  wrapper.style.margin = '0';
  wrapper.style.padding = '0';
  wrapper.style.width = `${width}px`;
  wrapper.style.height = `${sliceHeight}px`;
  wrapper.style.overflow = 'hidden';
  wrapper.style.boxSizing = 'border-box';
  wrapper.style.background = '#ffffff';

  wrapper.appendChild(clone);
  foreignObject.appendChild(wrapper);
  svg.appendChild(foreignObject);

  return {
    svgText: new XMLSerializer().serializeToString(svg),
    width,
    height: sliceHeight,
  };
}

async function svgTextToBlob(svgText, width, height, mimeType = 'image/png', quality = 0.92) {
  const svgBlob = new Blob([svgText], { type: 'image/svg+xml;charset=utf-8' });
  const url = URL.createObjectURL(svgBlob);
  try {
    const image = await new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = () => reject(new Error(getImageRuntimeError('svgLoad')));
      img.src = url;
    });

    const scale = Math.max(2, Math.min(3, Math.ceil(window.devicePixelRatio || 2)));
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.ceil(width * scale));
    canvas.height = Math.max(1, Math.ceil(height * scale));
    const ctx = canvas.getContext('2d');
    if (!ctx) {
      throw new Error(getImageRuntimeError('canvas'));
    }
    ctx.scale(scale, scale);
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, width, height);
    ctx.drawImage(image, 0, 0, width, height);

    const blob = await new Promise((resolve, reject) => {
      canvas.toBlob((value) => {
        if (value) {
          resolve(value);
        } else {
          reject(new Error(getImageRuntimeError('encode')));
        }
      }, mimeType, quality);
    });

    return blob;
  } finally {
    URL.revokeObjectURL(url);
  }
}

function getImageProgressText(current, total) {
  if (currentLocale === 'zh') {
    return `正在生成图片 ${current}/${total}…`;
  }
  return `Building image ${current}/${total}…`;
}

function getImageRuntimeError(kind) {
  const errors = {
    svgLoad: { zh: 'SVG 图像加载失败。', en: 'Failed to load the SVG snapshot.' },
    canvas: { zh: '无法创建图片画布。', en: 'Failed to create the image canvas.' },
    encode: { zh: '图片编码失败。', en: 'Failed to encode the image.' },
    capture: { zh: '浏览器截图失败。', en: 'Failed to capture the browser tab.' },
  };
  const bundle = errors[kind] || errors.encode;
  return currentLocale === 'zh' ? bundle.zh : bundle.en;
}

let lastVisibleTabCaptureAt = 0;
let visibleTabCaptureQueue = Promise.resolve();

function getVisibleTabCaptureIntervalMs() {
  const maxCallsPerSecond = Number(chrome?.tabs?.MAX_CAPTURE_VISIBLE_TAB_CALLS_PER_SECOND) || 2;
  const baseInterval = Math.ceil(1000 / Math.max(1, maxCallsPerSecond));
  return baseInterval + 120;
}

function isVisibleTabCaptureQuotaError(error) {
  return /MAX_CAPTURE_VISIBLE_TAB_CALLS_PER_SECOND/i.test(String(error?.message || error || ''));
}

async function waitForVisibleTabCaptureSlot() {
  const minInterval = getVisibleTabCaptureIntervalMs();
  const now = Date.now();
  const waitMs = Math.max(0, (lastVisibleTabCaptureAt + minInterval) - now);
  if (waitMs > 0) {
    await sleep(waitMs);
  }
  lastVisibleTabCaptureAt = Date.now();
}

function captureVisibleTabImage(windowId) {
  const runCapture = async () => {
    const maxAttempts = 4;
    let lastError = null;

    for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
      try {
        await waitForVisibleTabCaptureSlot();
        return await chrome.tabs.captureVisibleTab(windowId, { format: 'png' });
      } catch (error) {
        lastError = error;
        if (!isVisibleTabCaptureQuotaError(error) || attempt >= maxAttempts - 1) {
          throw new Error(error?.message || getImageRuntimeError('capture'));
        }
        lastVisibleTabCaptureAt = Date.now();
        await sleep(getVisibleTabCaptureIntervalMs());
      }
    }

    throw new Error(lastError?.message || getImageRuntimeError('capture'));
  };

  visibleTabCaptureQueue = visibleTabCaptureQueue.then(runCapture, runCapture);
  return visibleTabCaptureQueue;
}

async function dataUrlToBlob(dataUrl) {
  const response = await fetch(dataUrl);
  return response.blob();
}

function loadImageFromDataUrl(dataUrl) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error(getImageRuntimeError('capture')));
    image.src = dataUrl;
  });
}

async function cropCapturedImage(dataUrl, viewportHeight, targetHeight) {
  if (targetHeight >= viewportHeight - 1) {
    return dataUrlToBlob(dataUrl);
  }

  const image = await loadImageFromDataUrl(dataUrl);
  const scale = image.naturalHeight / Math.max(1, viewportHeight);
  const cropHeight = Math.max(1, Math.round(targetHeight * scale));
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, image.naturalWidth);
  canvas.height = cropHeight;
  const ctx = canvas.getContext('2d');
  if (!ctx) {
    throw new Error(getImageRuntimeError('canvas'));
  }
  ctx.drawImage(image, 0, 0, image.naturalWidth, cropHeight, 0, 0, image.naturalWidth, cropHeight);

  return new Promise((resolve, reject) => {
    canvas.toBlob((value) => {
      if (value) {
        resolve(value);
      } else {
        reject(new Error(getImageRuntimeError('encode')));
      }
    }, 'image/png');
  });
}

async function exportImage() {
  setBusy(true);
  setStatusKey('status.exportingImage');

  let tab = null;
  try {
    tab = await getActiveTab();
    assertChatGptTab(tab);
    const payload = await collectExportPayload(tab, 'docx-rendered');
    const html = buildDocumentHtml(payload, {
      locale: currentLocale,
      wordMode: 'rendered',
    });

    const prepareResponse = await requestAction(tab.id, 'CGPT_EXPORT_PREPARE_IMAGE_CAPTURE', {
      html,
    }, 'errors.imageFailed');

    const pageHeight = Math.max(1, Math.ceil(prepareResponse?.data?.pageHeight || 0));
    const viewportHeight = Math.max(1, Math.ceil(prepareResponse?.data?.viewportHeight || window.innerHeight || 1));
    const totalSlices = Math.max(1, Math.ceil(pageHeight / viewportHeight));

    const fallbackTitle = TEXT[currentLocale]?.doc?.fallbackTitle || TEXT.zh.doc.fallbackTitle;
    const baseTitle = sanitizeFilename(payload.title || fallbackTitle);
    const timestamp = getTimestamp();
    const imageEntries = [];

    for (let index = 0; index < totalSlices; index += 1) {
      setStatusText(getImageProgressText(index + 1, totalSlices));
      const startY = index * viewportHeight;
      const sliceHeight = Math.min(viewportHeight, Math.max(1, pageHeight - startY));
      await requestAction(tab.id, 'CGPT_EXPORT_SET_IMAGE_CAPTURE_SLICE', {
        startY,
      }, 'errors.imageFailed');
      const dataUrl = await captureVisibleTabImage(tab.windowId);
      const blob = await cropCapturedImage(dataUrl, viewportHeight, sliceHeight);
      const suffix = totalSlices > 1 ? `_${String(index + 1).padStart(2, '0')}` : '';
      const filename = `${baseTitle}_image_${timestamp}${suffix}.png`;
      imageEntries.push({ name: filename, blob });
    }

    setStatusKey('status.packingImageZip');
    const zipBlob = await createStoredZipFromBlobs(imageEntries);
    const zipFilename = `${baseTitle}_images_${timestamp}.zip`;
    await downloadBlob(zipBlob, zipFilename, { saveAs: true });

    setStatusKey(totalSlices === 1 ? 'status.imageReadySingle' : 'status.imageReadyMultiple', {
      count: totalSlices,
    });
  } catch (error) {
    setStatusText(error?.message || t('errors.imageFailed'), true);
  } finally {
    if (tab?.id) {
      try {
        await requestAction(tab.id, 'CGPT_EXPORT_CLEANUP_IMAGE_CAPTURE', {}, 'errors.imageFailed');
      } catch (_error) {
        // ignored
      }
    }
    setBusy(false);
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
    const payload = await collectExportPayload(tab, requestMode);
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
  setStatusKey(currentScope === 'project' ? 'status.preparingProjectPdf' : 'status.openingPdf');
  try {
    const tab = await getActiveTab();
    assertChatGptTab(tab);
    if (currentScope === 'project') {
      const payload = await collectProjectPayload(tab, 'docx-rendered');
      const html = buildPdfWindowHtml(payload, {
        locale: currentLocale,
        wordMode: 'rendered',
      });
      openPdfWindow(html);
      setStatusKey('status.pdfOpened');
      window.close();
      return;
    }
    await requestAction(tab.id, 'CGPT_EXPORT_OPEN_PDF_VIEW');
    setStatusKey('status.pdfOpened');
    window.close();
  } catch (error) {
    setStatusText(error?.message || t('errors.pdfFailed'), true);
    setBusy(false);
  }
}

async function openTeamPayPanel() {
  setBusy(true);
  setStatusKey('status.openingTeamPay');
  try {
    const tab = await getActiveTab();
    assertChatGptTab(tab);
    await requestAction(tab.id, 'CGPT_TEAM_PAY_OPEN_PANEL', {
      locale: currentLocale,
    }, 'errors.teamPayFailed');
    setStatusKey('status.teamPayOpened');
    window.close();
  } catch (error) {
    setStatusText(error?.message || t('errors.teamPayFailed'), true);
    setBusy(false);
  }
}

langButtons.forEach((button) => {
  button.addEventListener('click', () => {
    setLocale(button.dataset.lang);
  });
});

scopeButtons.forEach((button) => {
  button.addEventListener('click', () => {
    setScope(button.dataset.scope);
  });
});

pdfButton.addEventListener('click', exportPdf);
imageButton.addEventListener('click', exportImage);
docxRenderedButton.addEventListener('click', () => exportDocx('rendered'));
docxFastButton.addEventListener('click', () => exportDocx('fast'));
teamPayButton.addEventListener('click', openTeamPayPanel);

async function initializePopup() {
  renderUi();
  setStatusKey('status.idle');
  await updateScopeContext();
}

initializePopup();
