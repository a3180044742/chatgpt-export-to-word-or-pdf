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
    return { html: clone.innerHTML, text };
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
    <base href="${escapeHtml(document.baseURI)}">
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

    return undefined;
  });
})();
