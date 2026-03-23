(() => {
  const MESSAGE_SELECTOR_GROUPS = [
    ['[data-message-author-role]'],
    ['[data-testid^="conversation-turn-"]'],
    ['main article'],
  ];

  const CONTENT_SELECTORS = [
    '[data-message-content]',
    '.markdown',
    '[class*="markdown"]',
    '.prose',
    '[class*="prose"]',
    '.whitespace-pre-wrap',
    'article',
  ];

  const EXPORT_ID_ATTR = 'data-cgpt-export-id';
  const GRAPHIC_SKIP_SELECTOR = 'button, nav, form, textarea, input, select';
  const FORMULA_IMAGE_SCALE = 2;
  const FORMULA_DARK_COLOR = '#111827';
  const FORMULA_LIGHTNESS_THRESHOLD = 0.84;
  const TEAM_PAY_PANEL_ID = 'cgpt-team-pay-panel-root';
  const TEAM_PAY_STYLE_ID = 'cgpt-team-pay-style';
  const TEAM_PAY_STORAGE_KEY = 'cgptTeamPayConfig';
  const TEAM_PAY_DEFAULT_CONFIG = Object.freeze({
    workspace_name: 'MyTeam',
    price_interval: 'month',
    seat_quantity: 5,
    country: 'US',
    currency: 'USD',
    promo_campaign_id: 'team-1-month-free',
    page_mode: 'new',
  });
  const TEAM_PAY_TEXT = {
    zh: {
      title: 'ChatGPT Team 绑卡',
      subtitle: '1 个月免费试用',
      modeLabel: '页面样式',
      modeNew: '新页面',
      modeOld: '老页面',
      promoLabel: '优惠',
      promoValue: 'team-1-month-free（1个月免费）',
      payTodayLabel: '今日应付',
      payTodayValue: '$0.00',
      generate: '生成',
      generating: '生成中…',
      close: '关闭',
      gettingToken: '正在获取 Token…',
      generatingWithWorkspace: '正在生成…（工作区：{{workspace}}）',
      success: '生成成功',
      resultWorkspace: '工作区',
      resultLink: '绑卡链接',
      copy: '复制链接',
      copied: '已复制',
      open: '打开链接',
      tokenMissing: '未找到 accessToken。',
      loginFailed: '获取 Token 失败，请确认当前已登录 ChatGPT。',
      apiFailed: '生成 Team 绑卡链接失败。',
      copyFailed: '复制失败，请手动复制。',
    },
    en: {
      title: 'ChatGPT Team checkout',
      subtitle: '1 month free trial',
      modeLabel: 'Page mode',
      modeNew: 'New page',
      modeOld: 'Old page',
      promoLabel: 'Promotion',
      promoValue: 'team-1-month-free (1 month free)',
      payTodayLabel: 'Due today',
      payTodayValue: '$0.00',
      generate: 'Generate',
      generating: 'Generating…',
      close: 'Close',
      gettingToken: 'Fetching the access token…',
      generatingWithWorkspace: 'Generating… (workspace: {{workspace}})',
      success: 'Link generated',
      resultWorkspace: 'Workspace',
      resultLink: 'Checkout link',
      copy: 'Copy link',
      copied: 'Copied',
      open: 'Open link',
      tokenMissing: 'No accessToken was found.',
      loginFailed: 'Failed to fetch the access token. Make sure you are signed in to ChatGPT.',
      apiFailed: 'Failed to generate the Team checkout link.',
      copyFailed: 'Copy failed. Please copy the link manually.',
    },
  };
  let exportCounter = 0;
  const formulaImageCache = new Map();

  function cleanTitle(rawTitle) {
    const title = String(rawTitle || '').replace(/\s*[-|·]\s*ChatGPT\s*$/i, '').trim();
    return title || '';
  }

  function normalizeText(text) {
    return String(text || '')
      .replace(/\u00a0/g, ' ')
      .replace(/[\t\r]+/g, ' ')
      .replace(/\n{3,}/g, '\n\n')
      .replace(/ {2,}/g, ' ')
      .trim();
  }

  function dedupeContained(nodes) {
    return nodes.filter((node, index) => {
      return !nodes.some((other, otherIndex) => otherIndex !== index && other.contains(node));
    });
  }

  function findContentRoot(node) {
    if (!node) return null;
    for (const selector of CONTENT_SELECTORS) {
      const root = node.matches?.(selector) ? node : node.querySelector(selector);
      if (root && normalizeText(root.textContent || root.innerText || '').length) {
        return root;
      }
    }
    return normalizeText(node.textContent || node.innerText || '').length ? node : null;
  }

  function collectMessageNodes() {
    for (const selectorGroup of MESSAGE_SELECTOR_GROUPS) {
      const candidates = [];
      const seen = new Set();
      for (const selector of selectorGroup) {
        document.querySelectorAll(selector).forEach((node) => {
          if (!seen.has(node)) {
            seen.add(node);
            candidates.push(node);
          }
        });
      }
      const filtered = dedupeContained(candidates).filter((node) => findContentRoot(node));
      if (filtered.length >= 2) {
        return filtered;
      }
    }
    return [];
  }

  function detectRole(node) {
    const direct = node.getAttribute('data-message-author-role');
    if (direct === 'user' || direct === 'assistant' || direct === 'system') {
      return direct;
    }
    const nested = node.querySelector?.('[data-message-author-role]')?.getAttribute('data-message-author-role');
    if (nested === 'user' || nested === 'assistant' || nested === 'system') {
      return nested;
    }
    const hint = `${node.getAttribute('aria-label') || ''} ${String(node.className || '')}`.toLowerCase();
    if (hint.includes('assistant')) return 'assistant';
    if (hint.includes('user')) return 'user';
    return 'assistant';
  }

  function stripJunk(root) {
    if (!root) return;
    root.querySelectorAll('script, style, iframe, button, form, textarea, input, select').forEach((node) => node.remove());
    root.querySelectorAll('[data-testid*="copy"], [data-testid*="message-action"], [aria-label*="Copy"], [aria-label*="复制"]').forEach((node) => node.remove());
    root.querySelectorAll('details').forEach((details) => {
      details.open = true;
    });
  }

  function removeEmptyArtifacts(root) {
    if (!root) return;
    root.querySelectorAll('[data-state="closed"]').forEach((node) => {
      const hasContent = Boolean(
        normalizeText(node.textContent || '').length ||
        node.querySelector('img, svg, math, mjx-container, .katex, .katex-display, canvas, pre, code, table')
      );
      if (!hasContent) {
        node.remove();
      }
    });

    root.querySelectorAll('span.relative.inline-flex.items-center.select-none').forEach((node) => {
      const hasContent = Boolean(
        normalizeText(node.textContent || '').length ||
        node.querySelector('img, svg, math, mjx-container, .katex, .katex-display, canvas')
      );
      if (!hasContent) {
        node.remove();
      }
    });
  }

  function sanitizeAnchors(root) {
    root.querySelectorAll('a').forEach((anchor) => {
      anchor.setAttribute('target', '_blank');
      anchor.setAttribute('rel', 'noreferrer noopener');
    });
  }

  function nextExportId() {
    exportCounter += 1;
    return `cgpt-exp-${Date.now()}-${exportCounter}`;
  }

  function isDisplayFormulaNode(node) {
    if (!node || node.nodeType !== Node.ELEMENT_NODE) return false;
    if (node.matches('.katex-display')) return true;
    if (node.matches('mjx-container[display="true"], mjx-container[display="block"]')) return true;
    const display = getComputedStyle(node).display;
    if (display === 'block' || display === 'flex') return true;
    const rect = node.getBoundingClientRect();
    return rect.width > 0 && rect.width > Math.min(640, window.innerWidth * 0.7);
  }

  function getRenderableFormulaNodes(root) {
    const candidates = Array.from(root.querySelectorAll('.katex-display, .katex, mjx-container, math'));
    return candidates.filter((node) => {
      if (node.closest(GRAPHIC_SKIP_SELECTOR)) return false;
      if (node.matches('.katex-display')) return true;
      if (node.matches('.katex')) return !node.closest('.katex-display') && !node.parentElement?.closest('.katex');
      if (node.matches('mjx-container')) return !node.parentElement?.closest('mjx-container');
      if (node.matches('math')) return !node.closest('.katex, .katex-display, mjx-container');
      return false;
    });
  }

  function extractFormulaText(node) {
    if (!node) return '';
    const annotation = node.querySelector?.('annotation[encoding="application/x-tex"]');
    if (annotation?.textContent) return normalizeText(annotation.textContent);

    const script = node.querySelector?.('script[type="math/tex"], script[type="math/latex"]');
    if (script?.textContent) return normalizeText(script.textContent);

    const aria = node.getAttribute?.('aria-label') || node.closest?.('.katex, .katex-display, mjx-container')?.getAttribute?.('aria-label');
    if (aria) return normalizeText(aria);

    const dataCarrier = node.closest?.('*[data-tex], *[data-latex], *[data-math], *[data-equation], *[data-original], *[data-original-tex], *[data-source]') || node;
    const dataText = dataCarrier?.getAttribute?.('data-tex')
      || dataCarrier?.getAttribute?.('data-latex')
      || dataCarrier?.getAttribute?.('data-math')
      || dataCarrier?.getAttribute?.('data-equation')
      || dataCarrier?.getAttribute?.('data-original-tex')
      || dataCarrier?.getAttribute?.('data-original')
      || dataCarrier?.getAttribute?.('data-source');
    if (dataText) return normalizeText(dataText);

    const math = node.matches?.('math') ? node : node.querySelector?.('math');
    if (math?.textContent) return normalizeText(math.textContent);

    return normalizeText(node.textContent || '').slice(0, 500);
  }

  function markNodes(nodes) {
    const marked = [];
    for (const node of nodes) {
      const id = nextExportId();
      node.setAttribute(EXPORT_ID_ATTR, id);
      marked.push([id, node]);
    }
    return marked;
  }

  function clearMarkedNodes(marked) {
    for (const [, node] of marked) {
      node.removeAttribute(EXPORT_ID_ATTR);
    }
  }

  async function blobToDataUrl(blob) {
    return await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error || new Error('读取 Blob 失败'));
      reader.readAsDataURL(blob);
    });
  }

  async function fetchAsDataUrl(url) {
    if (!url) return null;
    if (/^data:/i.test(url)) return url;
    const attempts = [
      () => fetch(url, { credentials: 'include' }),
      () => fetch(url, { credentials: 'same-origin' }),
      () => fetch(url),
    ];
    let lastError = null;
    for (const attempt of attempts) {
      try {
        const response = await attempt();
        if (response?.ok) {
          const blob = await response.blob();
          if (blob && blob.size > 0) {
            return await blobToDataUrl(blob);
          }
        }
      } catch (error) {
        lastError = error;
      }
    }
    if (lastError) throw lastError;
    return null;
  }

  function getNodeRect(node) {
    const rect = node.getBoundingClientRect();
    const width = Math.max(1, Math.ceil(rect.width || node.scrollWidth || node.clientWidth || 1));
    const height = Math.max(1, Math.ceil(rect.height || node.scrollHeight || node.clientHeight || 1));
    return { width, height };
  }

  function parseRgbColor(value) {
    const match = String(value || '').match(/rgba?\(([^)]+)\)/i);
    if (!match) return null;
    const parts = match[1].split(',').map((item) => Number.parseFloat(item.trim()));
    if (parts.length < 3 || parts.slice(0, 3).some((item) => Number.isNaN(item))) {
      return null;
    }
    return {
      r: parts[0],
      g: parts[1],
      b: parts[2],
      a: Number.isNaN(parts[3]) ? 1 : parts[3],
    };
  }

  function isVeryLightColor(value) {
    const color = parseRgbColor(value);
    if (!color || color.a <= 0.01) return false;
    const lightness = (0.2126 * color.r + 0.7152 * color.g + 0.0722 * color.b) / 255;
    return lightness >= FORMULA_LIGHTNESS_THRESHOLD;
  }

  function normalizeFormulaColor(value) {
    if (!value || value === 'transparent' || value === 'none') return value;
    return isVeryLightColor(value) ? FORMULA_DARK_COLOR : value;
  }

  function copyComputedStyles(sourceNode, targetNode) {
    if (!(sourceNode instanceof Element) || !(targetNode instanceof Element)) return;
    const computed = getComputedStyle(sourceNode);
    if (computed.cssText) {
      targetNode.style.cssText = computed.cssText;
    } else {
      Array.from(computed).forEach((property) => {
        targetNode.style.setProperty(property, computed.getPropertyValue(property), computed.getPropertyPriority(property));
      });
    }

    const colorOverrides = [
      ['color', computed.color],
      ['fill', computed.fill],
      ['stroke', computed.stroke],
      ['border-top-color', computed.getPropertyValue('border-top-color')],
      ['border-right-color', computed.getPropertyValue('border-right-color')],
      ['border-bottom-color', computed.getPropertyValue('border-bottom-color')],
      ['border-left-color', computed.getPropertyValue('border-left-color')],
      ['outline-color', computed.getPropertyValue('outline-color')],
      ['text-decoration-color', computed.getPropertyValue('text-decoration-color')],
      ['-webkit-text-fill-color', computed.getPropertyValue('-webkit-text-fill-color')],
      ['-webkit-text-stroke-color', computed.getPropertyValue('-webkit-text-stroke-color')],
    ];

    colorOverrides.forEach(([property, rawValue]) => {
      const normalized = normalizeFormulaColor(rawValue);
      if (normalized && normalized !== rawValue) {
        targetNode.style.setProperty(property, normalized);
      }
    });

    targetNode.style.setProperty('animation', 'none');
    targetNode.style.setProperty('transition', 'none');
    targetNode.style.setProperty('caret-color', 'transparent');
  }

  function cloneNodeWithInlineStyles(sourceNode) {
    const clone = sourceNode.cloneNode(true);

    const walk = (source, target) => {
      if (source?.nodeType === Node.ELEMENT_NODE && target?.nodeType === Node.ELEMENT_NODE) {
        copyComputedStyles(source, target);
        target.removeAttribute(EXPORT_ID_ATTR);
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

  function getFormulaMeasurementNode(node) {
    if (!node?.matches) return node;
    if (node.matches('.katex-display')) {
      return node.querySelector(':scope > .katex') || node.querySelector('.katex') || node;
    }
    if (node.matches('mjx-container[display="true"], mjx-container[display="block"], mjx-container')) {
      return node.querySelector(':scope > mjx-math, :scope > svg') || node.querySelector('mjx-math, svg') || node;
    }
    return node;
  }

  function buildFormulaSnapshot(node, display) {
    const targetNode = getFormulaMeasurementNode(node) || node;
    const rect = getNodeRect(targetNode);
    if (!rect.width || !rect.height) return null;

    const horizontalPadding = display ? 12 : 6;
    const verticalPadding = display ? 10 : 4;
    const width = rect.width + horizontalPadding * 2;
    const height = rect.height + verticalPadding * 2;
    const styledClone = cloneNodeWithInlineStyles(targetNode);

    if (styledClone instanceof Element) {
      styledClone.removeAttribute(EXPORT_ID_ATTR);
      styledClone.style.margin = '0';
      styledClone.style.maxWidth = 'none';
      styledClone.style.background = 'transparent';
      styledClone.style.transformOrigin = 'top left';
    }

    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
    svg.setAttribute('width', String(width));
    svg.setAttribute('height', String(height));
    svg.setAttribute('viewBox', `0 0 ${width} ${height}`);

    const foreignObject = document.createElementNS('http://www.w3.org/2000/svg', 'foreignObject');
    foreignObject.setAttribute('x', '0');
    foreignObject.setAttribute('y', '0');
    foreignObject.setAttribute('width', '100%');
    foreignObject.setAttribute('height', '100%');

    const wrapper = document.createElementNS('http://www.w3.org/1999/xhtml', 'div');
    wrapper.setAttribute('xmlns', 'http://www.w3.org/1999/xhtml');
    wrapper.style.margin = '0';
    wrapper.style.padding = `${verticalPadding}px ${horizontalPadding}px`;
    wrapper.style.display = display ? 'block' : 'inline-block';
    wrapper.style.boxSizing = 'border-box';
    wrapper.style.background = 'transparent';
    wrapper.style.whiteSpace = display ? 'normal' : 'nowrap';

    wrapper.appendChild(styledClone);
    foreignObject.appendChild(wrapper);
    svg.appendChild(foreignObject);

    return {
      svgText: new XMLSerializer().serializeToString(svg),
      width,
      height,
    };
  }

  async function rasterizeSvgText(svgText, width, height) {
    const blob = new Blob([svgText], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    try {
      const image = await new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => resolve(img);
        img.onerror = () => reject(new Error('SVG 图像加载失败'));
        img.src = url;
      });
      const scale = Math.max(FORMULA_IMAGE_SCALE, Math.min(4, Math.ceil(window.devicePixelRatio || FORMULA_IMAGE_SCALE)));
      const canvas = document.createElement('canvas');
      canvas.width = Math.max(1, Math.ceil(width * scale));
      canvas.height = Math.max(1, Math.ceil(height * scale));
      const ctx = canvas.getContext('2d');
      if (!ctx) throw new Error('无法创建 canvas 上下文');
      ctx.scale(scale, scale);
      ctx.drawImage(image, 0, 0, width, height);
      return canvas.toDataURL('image/png');
    } finally {
      URL.revokeObjectURL(url);
    }
  }

  async function svgNodeToPngDataUrl(svgNode) {
    const { width, height } = getNodeRect(svgNode);
    const clone = svgNode.cloneNode(true);
    if (clone instanceof Element) {
      const computed = getComputedStyle(svgNode);
      const color = normalizeFormulaColor(computed.color);
      if (color && color !== 'transparent') {
        clone.style.color = color;
        clone.style.fill = normalizeFormulaColor(computed.fill) || color;
        clone.style.stroke = normalizeFormulaColor(computed.stroke) || '';
      }
    }
    if (!clone.getAttribute('xmlns')) {
      clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
    }
    if (!clone.getAttribute('width')) {
      clone.setAttribute('width', String(width));
    }
    if (!clone.getAttribute('height')) {
      clone.setAttribute('height', String(height));
    }
    const svgText = new XMLSerializer().serializeToString(clone);
    return await rasterizeSvgText(svgText, width, height);
  }

  function getFastFormulaSvgNode(node) {
    if (!node?.matches) return null;
    if (node.matches('mjx-container')) {
      return node.querySelector('svg');
    }
    return null;
  }

  async function formulaNodeToImage(node) {
    const display = isDisplayFormulaNode(node);
    const measurementNode = getFormulaMeasurementNode(node) || node;
    const rect = getNodeRect(measurementNode);
    const latex = extractFormulaText(node) || extractFormulaText(measurementNode) || '[formula]';
    const cacheKey = JSON.stringify([latex, display ? 1 : 0, rect.width, rect.height]);

    if (formulaImageCache.has(cacheKey)) {
      return formulaImageCache.get(cacheKey);
    }

    const task = (async () => {
      const directSvg = getFastFormulaSvgNode(node);
      if (directSvg) {
        const svgRect = getNodeRect(directSvg);
        const dataUrl = await svgNodeToPngDataUrl(directSvg);
        return dataUrl ? { dataUrl, width: svgRect.width, height: svgRect.height } : null;
      }
      const snapshot = buildFormulaSnapshot(node, display);
      if (!snapshot) return null;
      const dataUrl = await rasterizeSvgText(snapshot.svgText, snapshot.width, snapshot.height);
      return dataUrl ? { dataUrl, width: snapshot.width, height: snapshot.height } : null;
    })().catch(() => null);

    formulaImageCache.set(cacheKey, task);
    return task;
  }

  function createMathImageNode(imageData, source, display) {
    const img = document.createElement('img');
    img.className = display ? 'cgpt-math-image cgpt-math-display-image' : 'cgpt-math-image cgpt-math-inline-image';
    img.src = imageData.dataUrl;
    img.alt = source || 'formula';
    img.setAttribute('data-cgpt-formula', source || 'formula');
    img.style.maxWidth = '100%';
    img.style.height = 'auto';
    if (imageData.width > 0) img.style.width = `${imageData.width}px`;
    if (display) {
      const wrapper = document.createElement('div');
      wrapper.className = 'cgpt-math-block-image-wrap';
      wrapper.style.textAlign = 'center';
      wrapper.style.margin = '0.8em 0';
      img.style.display = 'inline-block';
      wrapper.appendChild(img);
      return wrapper;
    }
    img.style.display = 'inline-block';
    img.style.verticalAlign = 'middle';
    img.style.margin = '0 0.08em';
    return img;
  }

  function createMathFallbackNode(source, display) {
    const node = document.createElement(display ? 'pre' : 'code');
    node.className = display ? 'cgpt-math-block' : 'cgpt-math-inline';
    const payload = normalizeText(source || '[formula]') || '[formula]';
    node.textContent = display ? `$$\n${payload}\n$$` : `$${payload}$`;
    return node;
  }

  function replaceCanvasWithImages(root) {
    root.querySelectorAll('canvas').forEach((canvas) => {
      if (canvas.closest(GRAPHIC_SKIP_SELECTOR)) return;
      try {
        const img = document.createElement('img');
        img.src = canvas.toDataURL('image/png');
        img.alt = canvas.getAttribute('aria-label') || 'canvas';
        img.style.maxWidth = '100%';
        img.style.height = 'auto';
        const rect = getNodeRect(canvas);
        if (rect.width > 0) img.style.width = `${rect.width}px`;
        canvas.replaceWith(img);
      } catch (_error) {
        // ignored
      }
    });
  }

  async function replaceStandaloneSvgWithImages(root) {
    const svgNodes = Array.from(root.querySelectorAll('svg')).filter((node) => {
      if (node.closest(GRAPHIC_SKIP_SELECTOR)) return false;
      if (node.closest('.katex, .katex-display, mjx-container, math')) return false;
      return true;
    });

    for (const svg of svgNodes) {
      try {
        const dataUrl = await svgNodeToPngDataUrl(svg);
        if (!dataUrl) continue;
        const img = document.createElement('img');
        img.src = dataUrl;
        img.alt = svg.getAttribute('aria-label') || 'graphic';
        img.style.maxWidth = '100%';
        img.style.height = 'auto';
        const rect = getNodeRect(svg);
        if (rect.width > 0) img.style.width = `${rect.width}px`;
        svg.replaceWith(img);
      } catch (_error) {
        // ignored
      }
    }
  }

  async function inlineNormalImages(root) {
    const images = Array.from(root.querySelectorAll('img')).filter((img) => {
      if (img.closest(GRAPHIC_SKIP_SELECTOR)) return false;
      const src = img.currentSrc || img.getAttribute('src') || '';
      return Boolean(src);
    });

    await Promise.all(images.map(async (img) => {
      try {
        const src = img.currentSrc || img.getAttribute('src') || '';
        const dataUrl = await fetchAsDataUrl(src);
        if (dataUrl) {
          img.setAttribute('src', dataUrl);
          img.removeAttribute('srcset');
          img.style.maxWidth = img.style.maxWidth || '100%';
          img.style.height = img.style.height || 'auto';
        }
      } catch (_error) {
        // ignored
      }
    }));
  }

  function prepareClone(root) {
    const clone = root.cloneNode(true);
    stripJunk(clone);
    removeEmptyArtifacts(clone);
    sanitizeAnchors(clone);
    return clone;
  }

  function prepareMessageHtmlForPdf(root) {
    const clone = prepareClone(root);
    replaceCanvasWithImages(clone);
    const text = normalizeText(clone.innerText || clone.textContent || '');
    const tagName = (clone.tagName || 'div').toLowerCase();
    const attrs = serializeAttributes(clone);
    const html = `<${tagName}${attrs ? ` ${attrs}` : ''}>${clone.innerHTML}</${tagName}>`;
    return { html, text };
  }

  async function prepareMessageHtmlForDocxFast(root) {
    const sourceText = normalizeText(root.innerText || root.textContent || '');
    const marked = markNodes(getRenderableFormulaNodes(root));
    const clone = prepareClone(root);
    const cloneNodes = new Map();
    clone.querySelectorAll(`[${EXPORT_ID_ATTR}]`).forEach((node) => {
      cloneNodes.set(node.getAttribute(EXPORT_ID_ATTR), node);
    });

    for (const [id, originalNode] of marked) {
      const cloneNode = cloneNodes.get(id);
      if (!cloneNode) continue;
      const latex = extractFormulaText(originalNode) || extractFormulaText(cloneNode) || '[formula]';
      cloneNode.replaceWith(createMathFallbackNode(latex, isDisplayFormulaNode(originalNode)));
    }

    clearMarkedNodes(marked);
    clone.querySelectorAll(`[${EXPORT_ID_ATTR}]`).forEach((node) => node.removeAttribute(EXPORT_ID_ATTR));
    replaceCanvasWithImages(clone);
    await replaceStandaloneSvgWithImages(clone);
    await inlineNormalImages(clone);
    stripJunk(clone);
    removeEmptyArtifacts(clone);
    sanitizeAnchors(clone);
    const text = sourceText || normalizeText(clone.innerText || clone.textContent || '');
    return { html: clone.innerHTML, text };
  }

  async function prepareMessageHtmlForDocxRendered(root) {
    const sourceText = normalizeText(root.innerText || root.textContent || '');
    const marked = markNodes(getRenderableFormulaNodes(root));
    const clone = prepareClone(root);
    const cloneNodes = new Map();
    clone.querySelectorAll(`[${EXPORT_ID_ATTR}]`).forEach((node) => {
      cloneNodes.set(node.getAttribute(EXPORT_ID_ATTR), node);
    });

    const formulaReplacements = await Promise.all(marked.map(async ([id, originalNode]) => {
      const cloneNode = cloneNodes.get(id);
      if (!cloneNode) return null;
      const latex = extractFormulaText(originalNode) || extractFormulaText(cloneNode) || '[formula]';
      const display = isDisplayFormulaNode(originalNode);

      let replacementNode = null;
      try {
        const imageData = await formulaNodeToImage(originalNode);
        if (imageData?.dataUrl) {
          replacementNode = createMathImageNode(imageData, latex, display);
        }
      } catch (_error) {
        // ignored
      }

      return {
        cloneNode,
        replacementNode: replacementNode || createMathFallbackNode(latex, display),
      };
    }));

    formulaReplacements.forEach((item) => {
      if (item?.cloneNode?.isConnected) {
        item.cloneNode.replaceWith(item.replacementNode);
      }
    });

    clearMarkedNodes(marked);
    clone.querySelectorAll(`[${EXPORT_ID_ATTR}]`).forEach((node) => node.removeAttribute(EXPORT_ID_ATTR));
    replaceCanvasWithImages(clone);
    await replaceStandaloneSvgWithImages(clone);
    await inlineNormalImages(clone);
    stripJunk(clone);
    removeEmptyArtifacts(clone);
    sanitizeAnchors(clone);
    const text = sourceText || normalizeText(clone.innerText || clone.textContent || '');
    return { html: clone.innerHTML, text };
  }

  async function prepareMessageHtmlForDocx(root) {
    return prepareMessageHtmlForDocxRendered(root);
  }

  async function buildMessage(node, mode) {
    const root = findContentRoot(node);
    if (!root) return null;

    let prepared = null;
    if (mode === 'docx-fast') {
      prepared = await prepareMessageHtmlForDocxFast(root);
    } else if (mode === 'docx-rendered' || mode === 'docx') {
      prepared = await prepareMessageHtmlForDocxRendered(root);
    } else {
      prepared = prepareMessageHtmlForPdf(root);
    }

    if (!prepared.text) return null;
    return {
      role: detectRole(node),
      text: prepared.text,
      html: prepared.html,
    };
  }

  async function extractConversation(mode = 'pdf') {
    const nodes = collectMessageNodes();
    if (!nodes.length) {
      throw new Error('页面中没有找到可导出的聊天消息。请确认当前对话已经加载完成。');
    }

    const messages = [];
    const seenKeys = new Set();
    for (const node of nodes) {
      const message = await buildMessage(node, mode);
      if (!message) continue;
      const key = `${message.role}::${message.text}`;
      if (seenKeys.has(key)) continue;
      seenKeys.add(key);
      messages.push(message);
    }

    if (!messages.length) {
      throw new Error('提取到的消息为空。');
    }

    const firstUser = messages.find((item) => item.role === 'user');
    const fallbackTitle = firstUser?.text?.slice(0, 60) || 'ChatGPT Conversation';
    return {
      title: cleanTitle(document.title) || fallbackTitle,
      url: location.href,
      extractedAt: new Date().toISOString(),
      messageCount: messages.length,
      messages,
    };
  }

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function formatRole(role) {
    if (role === 'user') return 'User';
    if (role === 'assistant') return 'Assistant';
    if (role === 'system') return 'System';
    return 'Message';
  }

  function serializeAttributes(node) {
    if (!node?.attributes) return '';
    const allowed = [];
    for (const attr of Array.from(node.attributes)) {
      if (
        attr.name === 'class' ||
        attr.name === 'style' ||
        attr.name === 'lang' ||
        attr.name === 'dir' ||
        attr.name.startsWith('data-') ||
        attr.name.startsWith('aria-')
      ) {
        allowed.push(`${attr.name}="${escapeHtml(attr.value)}"`);
      }
    }
    return allowed.join(' ');
  }

  function collectPageStylesHtml() {
    const pieces = [];
    document.querySelectorAll('link[rel="stylesheet"][href], style').forEach((node) => {
      if (node.tagName?.toLowerCase() === 'link') {
        const href = node.getAttribute('href');
        if (!href) return;
        try {
          const absoluteHref = new URL(href, document.baseURI).href;
          pieces.push(`<link rel="stylesheet" href="${escapeHtml(absoluteHref)}">`);
        } catch (_error) {
          pieces.push(node.outerHTML);
        }
        return;
      }
      pieces.push(node.outerHTML);
    });
    return pieces.join('\n');
  }

  function buildPdfWindowHtml(payload) {
    const title = cleanTitle(payload.title) || 'ChatGPT Conversation';
    const pageStyles = collectPageStylesHtml();
    const htmlAttrs = serializeAttributes(document.documentElement);
    const bodyAttrs = serializeAttributes(document.body);
    const messagesHtml = (payload.messages || []).map((message) => `
      <section class="cgpt-export-message">
        <div class="cgpt-export-role">${escapeHtml(formatRole(message.role))}</div>
        <div class="cgpt-export-body">${message.html || ''}</div>
      </section>
    `).join('\n');

    return `<!DOCTYPE html>
<html ${htmlAttrs}>
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${escapeHtml(title)}</title>
    ${pageStyles}
    <style>
      :root {
        color-scheme: light dark;
      }
      body {
        margin: 0;
      }
      .cgpt-export-toolbar {
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
        color: #0f172a;
        font-family: Inter, "Segoe UI", Arial, sans-serif;
      }
      .cgpt-export-toolbar h1 {
        margin: 0;
        font-size: 16px;
        line-height: 1.2;
      }
      .cgpt-export-toolbar .meta {
        color: #64748b;
        font-size: 12px;
        line-height: 1.4;
        margin-top: 4px;
      }
      .cgpt-export-toolbar .actions {
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
      }
      .cgpt-export-toolbar button {
        appearance: none;
        border: 0;
        border-radius: 999px;
        padding: 10px 14px;
        font: 700 13px/1 Inter, "Segoe UI", Arial, sans-serif;
        cursor: pointer;
        background: #0f172a;
        color: #ffffff;
      }
      .cgpt-export-shell {
        max-width: 980px;
        margin: 0 auto;
        padding: 20px 24px 40px;
        box-sizing: border-box;
      }
      .cgpt-export-message {
        break-inside: avoid;
        page-break-inside: avoid;
        margin: 0 0 24px;
      }
      .cgpt-export-role {
        display: inline-block;
        margin-bottom: 10px;
        padding: 5px 10px;
        border-radius: 999px;
        background: #e2e8f0;
        color: #334155;
        font: 700 12px/1.2 Inter, "Segoe UI", Arial, sans-serif;
        text-transform: uppercase;
        letter-spacing: 0.04em;
      }
      .cgpt-export-body {
        font-family: Inter, "Segoe UI", Arial, sans-serif;
        color: #0f172a;
      }
      .cgpt-export-body img,
      .cgpt-export-body svg,
      .cgpt-export-body canvas,
      .cgpt-export-body table {
        max-width: 100%;
      }
      .cgpt-export-body ol,
      .cgpt-export-body ul {
        padding-left: 1.6em;
        padding-inline-start: 1.6em;
        list-style-position: outside;
      }
      .cgpt-export-body ul {
        list-style-type: disc;
      }
      .cgpt-export-body ul ul {
        list-style-type: circle;
      }
      .cgpt-export-body ul ul ul {
        list-style-type: square;
      }
      .cgpt-export-body ol {
        list-style-type: decimal;
      }
      .cgpt-export-body ol ol {
        list-style-type: lower-alpha;
      }
      .cgpt-export-body ol ol ol {
        list-style-type: lower-roman;
      }
      .cgpt-export-body li {
        display: list-item;
        margin: 0.25em 0;
      }
      .cgpt-export-body li::marker {
        color: inherit;
      }
      .cgpt-export-body pre {
        white-space: pre-wrap !important;
        overflow-wrap: anywhere;
      }
      .cgpt-export-body a {
        color: inherit;
      }
      @page {
        margin: 14mm 12mm;
      }
      @media print {
        .cgpt-export-toolbar {
          display: none !important;
        }
        .cgpt-export-shell {
          max-width: none;
          padding: 0;
        }
        body {
          background: #ffffff !important;
        }
      }
    </style>
  </head>
  <body ${bodyAttrs}>
    <div class="cgpt-export-toolbar">
      <div>
        <h1>${escapeHtml(title)}</h1>
        <div class="meta">${payload.messageCount || 0} 条消息 · 导出时间 ${escapeHtml(new Date(payload.extractedAt || Date.now()).toLocaleString())}</div>
      </div>
      <div class="actions">
        <button type="button" id="cgpt-export-print-btn" data-cgpt-export-action="print">打印 / 保存为 PDF</button>
        <button type="button" id="cgpt-export-close-btn" data-cgpt-export-action="close">关闭</button>
      </div>
    </div>
    <main class="cgpt-export-shell">
      ${messagesHtml}
    </main>
  </body>
</html>`;
  }

  function wirePdfPopup(popup) {
    if (!popup || popup.closed) return;

    const bindActions = () => {
      const doc = popup.document;
      if (!doc) return false;

      const printButton = doc.getElementById('cgpt-export-print-btn');
      const closeButton = doc.getElementById('cgpt-export-close-btn');

      if (printButton && !printButton.hasAttribute('data-cgpt-export-bound')) {
        printButton.setAttribute('data-cgpt-export-bound', '1');
        printButton.addEventListener('click', () => {
          try {
            popup.focus();
            setTimeout(() => popup.print(), 60);
          } catch (_error) {
            // ignored
          }
        });
      }

      if (closeButton && !closeButton.hasAttribute('data-cgpt-export-bound')) {
        closeButton.setAttribute('data-cgpt-export-bound', '1');
        closeButton.addEventListener('click', () => {
          try {
            popup.close();
          } catch (_error) {
            // ignored
          }
        });
      }

      return Boolean(printButton && closeButton);
    };

    const tryAutoPrint = () => {
      if (!popup || popup.closed) return;
      const run = () => {
        try {
          popup.focus();
          setTimeout(() => popup.print(), 120);
        } catch (_error) {
          // ignored
        }
      };

      try {
        const fontsReady = popup.document?.fonts?.ready;
        if (fontsReady && typeof fontsReady.then === 'function') {
          fontsReady.then(() => setTimeout(run, 220)).catch(() => setTimeout(run, 220));
        } else {
          setTimeout(run, 320);
        }
      } catch (_error) {
        setTimeout(run, 320);
      }
    };

    let attempts = 0;
    const timer = setInterval(() => {
      attempts += 1;
      if (!popup || popup.closed) {
        clearInterval(timer);
        return;
      }
      const ready = popup.document?.readyState === 'interactive' || popup.document?.readyState === 'complete';
      const bound = bindActions();
      if (ready && bound) {
        clearInterval(timer);
        tryAutoPrint();
      } else if (attempts >= 40) {
        clearInterval(timer);
      }
    }, 120);
  }

  function openPdfView() {
    const nodes = collectMessageNodes();
    if (!nodes.length) {
      throw new Error('页面中没有找到可导出的聊天消息。请确认当前对话已经加载完成。');
    }

    const messages = [];
    const seenKeys = new Set();
    for (const node of nodes) {
      const root = findContentRoot(node);
      if (!root) continue;
      const prepared = prepareMessageHtmlForPdf(root);
      if (!prepared.text) continue;
      const role = detectRole(node);
      const key = `${role}::${prepared.text}`;
      if (seenKeys.has(key)) continue;
      seenKeys.add(key);
      messages.push({ role, text: prepared.text, html: prepared.html });
    }

    if (!messages.length) {
      throw new Error('提取到的消息为空。');
    }

    const firstUser = messages.find((item) => item.role === 'user');
    const payload = {
      title: cleanTitle(document.title) || firstUser?.text?.slice(0, 60) || 'ChatGPT Conversation',
      url: location.href,
      extractedAt: new Date().toISOString(),
      messageCount: messages.length,
      messages,
    };

    const html = buildPdfWindowHtml(payload);
    const popup = window.open('', '_blank');
    if (!popup) {
      throw new Error('浏览器拦截了新窗口。请允许 chatgpt.com 打开弹窗后重试。');
    }
    popup.document.open();
    popup.document.write(html);
    popup.document.close();
    wirePdfPopup(popup);
    popup.focus();
  }

  const IMAGE_CAPTURE_OVERLAY_ID = 'cgpt-export-image-capture-overlay';
  let imageCaptureState = null;

  function waitForImageCaptureImage(image) {
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

  async function waitForImageCaptureFrameReady(iframe) {
    const frameDocument = iframe?.contentDocument;
    const frameWindow = iframe?.contentWindow;
    if (!frameDocument || !frameWindow) return;

    if (frameDocument.readyState !== 'complete') {
      await new Promise((resolve) => {
        let settled = false;
        const finish = () => {
          if (settled) return;
          settled = true;
          resolve();
        };
        iframe.addEventListener('load', finish, { once: true });
        setTimeout(finish, 2000);
      });
    }

    try {
      if (frameDocument.fonts?.ready) {
        await frameDocument.fonts.ready;
      }
    } catch (_error) {
      // ignored
    }

    const images = Array.from(frameDocument.images || []);
    await Promise.all(images.map(waitForImageCaptureImage));

    await new Promise((resolve) => {
      frameWindow.requestAnimationFrame(() => frameWindow.requestAnimationFrame(resolve));
    });
  }

  function getImageCapturePage(frameDocument) {
    if (!frameDocument) return null;
    return frameDocument.querySelector('.page') || frameDocument.body?.firstElementChild || frameDocument.body || null;
  }

  function getImageCapturePageHeight(frameDocument) {
    const page = getImageCapturePage(frameDocument);
    return Math.max(
      1,
      Math.ceil(page?.scrollHeight || 0),
      Math.ceil(frameDocument?.body?.scrollHeight || 0),
      Math.ceil(frameDocument?.documentElement?.scrollHeight || 0)
    );
  }

  function cleanupImageCaptureView() {
    const state = imageCaptureState;
    if (!state) return;

    try {
      state.overlay?.remove();
    } catch (_error) {
      // ignored
    }

    try {
      document.documentElement.style.overflow = state.htmlOverflow;
    } catch (_error) {
      // ignored
    }

    try {
      document.body.style.overflow = state.bodyOverflow;
    } catch (_error) {
      // ignored
    }

    imageCaptureState = null;
  }

  async function prepareImageCaptureView(html) {
    cleanupImageCaptureView();

    if (!html) {
      throw new Error('缺少图片导出的页面内容。');
    }

    const overlay = document.createElement('div');
    overlay.id = IMAGE_CAPTURE_OVERLAY_ID;
    overlay.setAttribute('aria-hidden', 'true');
    overlay.style.position = 'fixed';
    overlay.style.inset = '0';
    overlay.style.zIndex = '2147483647';
    overlay.style.background = '#ffffff';
    overlay.style.overflow = 'hidden';
    overlay.style.pointerEvents = 'none';

    const iframe = document.createElement('iframe');
    iframe.setAttribute('aria-hidden', 'true');
    iframe.style.display = 'block';
    iframe.style.width = '100vw';
    iframe.style.height = '100vh';
    iframe.style.border = '0';
    iframe.style.background = '#ffffff';
    iframe.style.margin = '0';
    iframe.style.padding = '0';

    overlay.appendChild(iframe);
    document.body.appendChild(overlay);

    const htmlOverflow = document.documentElement.style.overflow;
    const bodyOverflow = document.body.style.overflow;
    document.documentElement.style.overflow = 'hidden';
    document.body.style.overflow = 'hidden';

    imageCaptureState = {
      overlay,
      iframe,
      htmlOverflow,
      bodyOverflow,
      page: null,
    };

    const frameDocument = iframe.contentDocument;
    frameDocument.open();
    frameDocument.write(html);
    frameDocument.close();

    await waitForImageCaptureFrameReady(iframe);

    const innerDocument = iframe.contentDocument;
    const page = getImageCapturePage(innerDocument);
    if (!page) {
      throw new Error('没有找到图片导出的页面内容。');
    }

    innerDocument.documentElement.style.margin = '0';
    innerDocument.documentElement.style.overflow = 'hidden';
    innerDocument.documentElement.style.background = '#ffffff';
    innerDocument.body.style.margin = '0';
    innerDocument.body.style.overflow = 'hidden';
    innerDocument.body.style.background = '#ffffff';

    page.style.willChange = 'transform';
    page.style.transformOrigin = 'top center';
    page.style.transform = 'translateY(0)';

    imageCaptureState.page = page;

    return {
      pageHeight: getImageCapturePageHeight(innerDocument),
      viewportHeight: window.innerHeight,
      viewportWidth: window.innerWidth,
    };
  }

  async function setImageCaptureSlice(startY = 0) {
    const state = imageCaptureState;
    if (!state?.iframe?.contentDocument) {
      throw new Error('图片导出预览尚未准备完成。');
    }

    const frameDocument = state.iframe.contentDocument;
    const page = state.page || getImageCapturePage(frameDocument);
    if (!page) {
      throw new Error('没有找到图片导出的页面内容。');
    }

    page.style.transform = `translateY(-${Math.max(0, Math.floor(startY))}px)`;

    const frameWindow = state.iframe.contentWindow || window;
    await new Promise((resolve) => {
      frameWindow.requestAnimationFrame(() => frameWindow.requestAnimationFrame(resolve));
    });

    return {
      viewportHeight: window.innerHeight,
      viewportWidth: window.innerWidth,
    };
  }

  function getTeamPayLocale(preferredLocale) {
    if (preferredLocale === 'zh' || preferredLocale === 'en') {
      return preferredLocale;
    }
    const browserLocale = String(navigator.language || '').toLowerCase();
    return browserLocale.startsWith('zh') ? 'zh' : 'en';
  }

  function teamPayInterpolate(template, params = {}) {
    return String(template).replace(/\{\{\s*(\w+)\s*\}\}/g, (_match, key) => {
      return Object.prototype.hasOwnProperty.call(params, key) ? String(params[key]) : '';
    });
  }

  function teamPayText(key, locale = 'zh', params = {}) {
    const bundle = TEAM_PAY_TEXT[getTeamPayLocale(locale)] || TEAM_PAY_TEXT.zh;
    const value = String(key || '').split('.').reduce((acc, part) => (acc && acc[part] != null ? acc[part] : undefined), bundle);
    if (typeof value === 'string') {
      return teamPayInterpolate(value, params);
    }
    return value;
  }

  function ensureTeamPayStyles() {
    if (document.getElementById(TEAM_PAY_STYLE_ID)) {
      return;
    }
    const style = document.createElement('style');
    style.id = TEAM_PAY_STYLE_ID;
    style.textContent = `
      #${TEAM_PAY_PANEL_ID} {
        position: fixed;
        inset: 0;
        z-index: 2147483646;
        font-family: Inter, "Segoe UI", system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-overlay {
        position: absolute;
        inset: 0;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 24px;
        background: rgba(15, 23, 42, 0.55);
        backdrop-filter: blur(4px);
        box-sizing: border-box;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-dialog {
        width: min(420px, calc(100vw - 32px));
        max-height: min(88vh, 900px);
        overflow: auto;
        background: #ffffff;
        border-radius: 18px;
        box-shadow: 0 30px 80px rgba(15, 23, 42, 0.28);
        color: #0f172a;
        position: relative;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-close {
        position: absolute;
        top: 14px;
        right: 14px;
        width: 34px;
        height: 34px;
        border: 0;
        border-radius: 999px;
        background: rgba(255, 255, 255, 0.18);
        color: #ffffff;
        font-size: 18px;
        cursor: pointer;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-header {
        background: linear-gradient(135deg, #10a37f 0%, #1a7f5a 100%);
        color: #ffffff;
        padding: 24px 24px 20px;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-header h2 {
        margin: 0;
        font-size: 22px;
        line-height: 1.2;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-header p {
        margin: 8px 0 0;
        opacity: 0.92;
        font-size: 14px;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-body {
        padding: 22px 24px 24px;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-field {
        display: block;
        margin-bottom: 16px;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-field span {
        display: block;
        margin-bottom: 6px;
        font-size: 14px;
        font-weight: 600;
        color: #334155;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-field select,
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-result-input {
        width: 100%;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 11px 12px;
        font-size: 14px;
        box-sizing: border-box;
        background: #ffffff;
        color: #0f172a;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-summary {
        margin-bottom: 18px;
        padding: 14px 16px;
        border-radius: 12px;
        background: #f0fdf4;
        color: #166534;
        line-height: 1.65;
        font-size: 13px;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-generate,
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-result-actions button {
        appearance: none;
        border: 0;
        border-radius: 12px;
        padding: 12px 14px;
        font-size: 14px;
        font-weight: 700;
        cursor: pointer;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-generate {
        width: 100%;
        background: linear-gradient(135deg, #10a37f 0%, #1a7f5a 100%);
        color: #ffffff;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-generate:disabled,
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-result-actions button:disabled {
        opacity: 0.7;
        cursor: wait;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-status {
        margin-top: 16px;
        padding: 14px 16px;
        border-radius: 12px;
        font-size: 13px;
        line-height: 1.65;
        border: 1px solid transparent;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-status[data-state="info"] {
        background: #f8fafc;
        color: #334155;
        border-color: #e2e8f0;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-status[data-state="error"] {
        background: #fef2f2;
        color: #b91c1c;
        border-color: #fecaca;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-status[data-state="success"] {
        background: #f0fdf4;
        color: #166534;
        border-color: #bbf7d0;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-result-label,
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-result-meta {
        display: block;
        margin: 10px 0 8px;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-result-meta {
        margin-top: 6px;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-result-actions {
        display: flex;
        gap: 10px;
        margin-top: 10px;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-result-actions button {
        flex: 1;
        background: #0f172a;
        color: #ffffff;
      }
      #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-result-actions .cgpt-team-pay-open {
        background: #334155;
      }
      @media (max-width: 520px) {
        #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-overlay {
          padding: 14px;
        }
        #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-dialog {
          width: min(100vw - 16px, 100%);
        }
        #${TEAM_PAY_PANEL_ID} .cgpt-team-pay-result-actions {
          flex-direction: column;
        }
      }
    `;
    document.documentElement.appendChild(style);
  }

  function loadTeamPayConfig() {
    try {
      const parsed = JSON.parse(localStorage.getItem(TEAM_PAY_STORAGE_KEY) || '{}');
      return {
        ...TEAM_PAY_DEFAULT_CONFIG,
        ...(parsed && typeof parsed === 'object' ? parsed : {}),
      };
    } catch (_error) {
      return { ...TEAM_PAY_DEFAULT_CONFIG };
    }
  }

  function saveTeamPayConfig(config) {
    try {
      localStorage.setItem(TEAM_PAY_STORAGE_KEY, JSON.stringify({
        page_mode: config?.page_mode === 'old' ? 'old' : 'new',
      }));
    } catch (_error) {
      // ignored
    }
  }

  function removeTeamPayPanel() {
    document.getElementById(TEAM_PAY_PANEL_ID)?.remove();
  }

  function setTeamPayStatus(statusNode, html, kind = 'info') {
    if (!statusNode) return;
    statusNode.hidden = false;
    statusNode.dataset.state = kind;
    statusNode.innerHTML = html;
  }

  async function readTeamPayResponse(response) {
    const text = await response.text();
    if (!text) {
      return null;
    }
    try {
      return JSON.parse(text);
    } catch (_error) {
      return { detail: text };
    }
  }

  async function getTeamPayAccessToken(locale) {
    let response;
    try {
      response = await fetch(new URL('/api/auth/session', location.origin).toString(), {
        credentials: 'include',
        headers: {
          Accept: 'application/json',
        },
      });
    } catch (error) {
      throw new Error(error?.message || teamPayText('loginFailed', locale));
    }

    const data = await readTeamPayResponse(response);
    if (!response.ok) {
      throw new Error(data?.detail || data?.message || teamPayText('loginFailed', locale));
    }
    if (data?.accessToken) {
      return data.accessToken;
    }
    throw new Error(teamPayText('tokenMissing', locale));
  }

  function extractTeamWorkspaceName(token) {
    try {
      const parts = String(token || '').split('.');
      if (parts.length < 2) {
        return TEAM_PAY_DEFAULT_CONFIG.workspace_name;
      }
      let payload = parts[1].replace(/-/g, '+').replace(/_/g, '/');
      while (payload.length % 4) {
        payload += '=';
      }
      const decoded = JSON.parse(atob(payload));
      const email = decoded.email
        || decoded['https://api.openai.com/profile']?.email
        || decoded['https://api.openai.com/auth']?.email;
      if (email && email.includes('@')) {
        return email.split('@')[0];
      }
    } catch (_error) {
      // ignored
    }
    return TEAM_PAY_DEFAULT_CONFIG.workspace_name;
  }

  async function generateTeamPayCheckout(token, config, locale) {
    const payload = {
      plan_name: 'chatgptteamplan',
      team_plan_data: {
        workspace_name: config.workspace_name,
        price_interval: config.price_interval,
        seat_quantity: config.seat_quantity,
      },
      billing_details: {
        country: config.country,
        currency: String(config.currency || '').toUpperCase(),
      },
      cancel_url: `${location.origin}/#pricing`,
      promo_campaign: {
        promo_campaign_id: config.promo_campaign_id,
        is_coupon_from_query_param: false,
      },
      checkout_ui_mode: config.page_mode === 'old' ? 'redirect' : 'custom',
    };

    let response;
    try {
      response = await fetch(new URL('/backend-api/payments/checkout', location.origin).toString(), {
        method: 'POST',
        credentials: 'include',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
          Accept: 'application/json',
        },
        body: JSON.stringify(payload),
      });
    } catch (error) {
      throw new Error(error?.message || teamPayText('apiFailed', locale));
    }

    const data = await readTeamPayResponse(response);
    if (!response.ok) {
      throw new Error(data?.detail || data?.message || teamPayText('apiFailed', locale));
    }

    if (config.page_mode === 'new') {
      if (data?.checkout_session_id) {
        const processor = data.processor_entity || 'openai_llc';
        return {
          mode: 'new',
          url: `${location.origin}/checkout/${processor}/${data.checkout_session_id}`,
        };
      }
    } else if (data?.url) {
      return {
        mode: 'old',
        url: data.url,
      };
    }

    if (data?.url) {
      return {
        mode: 'old',
        url: data.url,
      };
    }
    if (data?.checkout_session_id) {
      const processor = data.processor_entity || 'openai_llc';
      return {
        mode: 'new',
        url: `${location.origin}/checkout/${processor}/${data.checkout_session_id}`,
      };
    }

    throw new Error(data?.detail || teamPayText('apiFailed', locale));
  }

  async function copyTeamPayText(text) {
    if (navigator.clipboard?.writeText) {
      try {
        await navigator.clipboard.writeText(text);
        return true;
      } catch (_error) {
        // ignored
      }
    }

    try {
      const textarea = document.createElement('textarea');
      textarea.value = text;
      textarea.setAttribute('readonly', 'readonly');
      textarea.style.position = 'fixed';
      textarea.style.top = '-9999px';
      document.body.appendChild(textarea);
      textarea.focus();
      textarea.select();
      const success = document.execCommand('copy');
      textarea.remove();
      return success;
    } catch (_error) {
      return false;
    }
  }

  function getTeamPayModeLabel(locale, mode) {
    return mode === 'old' ? teamPayText('modeOld', locale) : teamPayText('modeNew', locale);
  }

  function renderTeamPayResult(statusNode, locale, workspaceName, result) {
    if (!statusNode) return;

    const inputId = `cgpt-team-pay-url-${Date.now()}`;
    setTeamPayStatus(statusNode, `
      <div>✅ <strong>${escapeHtml(teamPayText('success', locale))}</strong> (${escapeHtml(getTeamPayModeLabel(locale, result.mode))})</div>
      <div class="cgpt-team-pay-result-meta"><strong>${escapeHtml(teamPayText('resultWorkspace', locale))}:</strong> ${escapeHtml(workspaceName)}</div>
      <label class="cgpt-team-pay-result-label" for="${escapeHtml(inputId)}"><strong>${escapeHtml(teamPayText('resultLink', locale))}:</strong></label>
      <input id="${escapeHtml(inputId)}" class="cgpt-team-pay-result-input" type="text" readonly>
      <div class="cgpt-team-pay-result-actions">
        <button type="button" class="cgpt-team-pay-copy">${escapeHtml(teamPayText('copy', locale))}</button>
        <button type="button" class="cgpt-team-pay-open">${escapeHtml(teamPayText('open', locale))}</button>
      </div>
    `, 'success');

    const input = statusNode.querySelector('.cgpt-team-pay-result-input');
    const copyButton = statusNode.querySelector('.cgpt-team-pay-copy');
    const openButton = statusNode.querySelector('.cgpt-team-pay-open');
    if (input) {
      input.value = result.url;
      input.addEventListener('click', () => input.select());
    }
    if (copyButton) {
      copyButton.addEventListener('click', async () => {
        const copied = await copyTeamPayText(result.url);
        if (copied) {
          copyButton.textContent = teamPayText('copied', locale);
          setTimeout(() => {
            if (copyButton.isConnected) {
              copyButton.textContent = teamPayText('copy', locale);
            }
          }, 1500);
        } else {
          setTeamPayStatus(statusNode, `❌ <strong>${escapeHtml(teamPayText('copyFailed', locale))}</strong>`, 'error');
        }
      });
    }
    if (openButton) {
      openButton.addEventListener('click', () => {
        window.open(result.url, '_blank', 'noopener');
      });
    }
  }

  function showTeamPayPanel(preferredLocale) {
    ensureTeamPayStyles();
    removeTeamPayPanel();

    const locale = getTeamPayLocale(preferredLocale);
    const config = loadTeamPayConfig();
    const panel = document.createElement('div');
    panel.id = TEAM_PAY_PANEL_ID;
    panel.innerHTML = `
      <div class="cgpt-team-pay-overlay">
        <div class="cgpt-team-pay-dialog" role="dialog" aria-modal="true" aria-label="${escapeHtml(teamPayText('title', locale))}">
          <button type="button" class="cgpt-team-pay-close" aria-label="${escapeHtml(teamPayText('close', locale))}">×</button>
          <div class="cgpt-team-pay-header">
            <h2>${escapeHtml(teamPayText('title', locale))}</h2>
            <p>${escapeHtml(teamPayText('subtitle', locale))}</p>
          </div>
          <div class="cgpt-team-pay-body">
            <label class="cgpt-team-pay-field">
              <span>${escapeHtml(teamPayText('modeLabel', locale))}</span>
              <select class="cgpt-team-pay-mode">
                <option value="new"${config.page_mode === 'new' ? ' selected' : ''}>${escapeHtml(teamPayText('modeNew', locale))}</option>
                <option value="old"${config.page_mode === 'old' ? ' selected' : ''}>${escapeHtml(teamPayText('modeOld', locale))}</option>
              </select>
            </label>
            <div class="cgpt-team-pay-summary">
              ✨ <strong>${escapeHtml(teamPayText('promoLabel', locale))}:</strong> ${escapeHtml(teamPayText('promoValue', locale))}<br>
              💰 <strong>${escapeHtml(teamPayText('payTodayLabel', locale))}:</strong> ${escapeHtml(teamPayText('payTodayValue', locale))}
            </div>
            <button type="button" class="cgpt-team-pay-generate">${escapeHtml(teamPayText('generate', locale))}</button>
            <div class="cgpt-team-pay-status" data-state="info" hidden></div>
          </div>
        </div>
      </div>
    `;

    document.body.appendChild(panel);

    const overlay = panel.querySelector('.cgpt-team-pay-overlay');
    const closeButton = panel.querySelector('.cgpt-team-pay-close');
    const generateButton = panel.querySelector('.cgpt-team-pay-generate');
    const statusNode = panel.querySelector('.cgpt-team-pay-status');
    const modeSelect = panel.querySelector('.cgpt-team-pay-mode');

    overlay?.addEventListener('click', (event) => {
      if (event.target === overlay) {
        removeTeamPayPanel();
      }
    });
    closeButton?.addEventListener('click', () => removeTeamPayPanel());
    panel.querySelector('.cgpt-team-pay-dialog')?.addEventListener('click', (event) => event.stopPropagation());

    generateButton?.addEventListener('click', async () => {
      const pageMode = modeSelect?.value === 'old' ? 'old' : 'new';
      const nextConfig = {
        ...config,
        page_mode: pageMode,
      };
      saveTeamPayConfig(nextConfig);

      generateButton.disabled = true;
      generateButton.textContent = teamPayText('generating', locale);
      setTeamPayStatus(statusNode, escapeHtml(teamPayText('gettingToken', locale)), 'info');

      try {
        const token = await getTeamPayAccessToken(locale);
        const workspaceName = extractTeamWorkspaceName(token);
        setTeamPayStatus(statusNode, escapeHtml(teamPayText('generatingWithWorkspace', locale, {
          workspace: workspaceName,
        })), 'info');
        const result = await generateTeamPayCheckout(token, {
          ...nextConfig,
          workspace_name: workspaceName,
        }, locale);
        renderTeamPayResult(statusNode, locale, workspaceName, result);
      } catch (error) {
        setTeamPayStatus(statusNode, `❌ <strong>${escapeHtml(error?.message || teamPayText('apiFailed', locale))}</strong>`, 'error');
      } finally {
        generateButton.disabled = false;
        generateButton.textContent = teamPayText('generate', locale);
      }
    });

    queueMicrotask(() => {
      modeSelect?.focus();
    });
  }

  window.addEventListener('pagehide', cleanupImageCaptureView);

  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    if (message?.type === 'CGPT_EXPORT_EXTRACT') {
      (async () => {
        try {
          const data = await extractConversation(message.mode || 'docx');
          sendResponse({ ok: true, data });
        } catch (error) {
          sendResponse({ ok: false, error: error?.message || '提取聊天失败。' });
        }
      })();
      return true;
    }

    if (message?.type === 'CGPT_EXPORT_OPEN_PDF_VIEW') {
      try {
        openPdfView();
        sendResponse({ ok: true });
      } catch (error) {
        sendResponse({ ok: false, error: error?.message || '打开 PDF 视图失败。' });
      }
      return true;
    }

    if (message?.type === 'CGPT_EXPORT_PREPARE_IMAGE_CAPTURE') {
      (async () => {
        try {
          const data = await prepareImageCaptureView(message.html || '');
          sendResponse({ ok: true, data });
        } catch (error) {
          cleanupImageCaptureView();
          sendResponse({ ok: false, error: error?.message || '准备图片导出预览失败。' });
        }
      })();
      return true;
    }

    if (message?.type === 'CGPT_EXPORT_SET_IMAGE_CAPTURE_SLICE') {
      (async () => {
        try {
          const data = await setImageCaptureSlice(message.startY || 0);
          sendResponse({ ok: true, data });
        } catch (error) {
          sendResponse({ ok: false, error: error?.message || '更新图片导出切片失败。' });
        }
      })();
      return true;
    }

    if (message?.type === 'CGPT_EXPORT_CLEANUP_IMAGE_CAPTURE') {
      try {
        cleanupImageCaptureView();
        sendResponse({ ok: true });
      } catch (error) {
        sendResponse({ ok: false, error: error?.message || '清理图片导出预览失败。' });
      }
      return true;
    }

    if (message?.type === 'CGPT_TEAM_PAY_OPEN_PANEL') {
      try {
        showTeamPayPanel(message.locale);
        sendResponse({ ok: true });
      } catch (error) {
        sendResponse({ ok: false, error: error?.message || '打开 Team 绑卡面板失败。' });
      }
      return true;
    }

    return undefined;
  });
})();
