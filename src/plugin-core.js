/* BetterGravity Office & PDF Preview — Core Plugin Logic v1.1 */

const ROOT_CLASS = "bg-doc-preview-root";
const HIDDEN_CLASS = "bg-doc-preview-hidden-fallback";
const LOCAL_PREVIEW_CONFIG = null;
const PDF_OFFSCREEN_PAGE_LIMIT = 4;
const PDF_OFFSCREEN_BYTE_LIMIT = 64 * 1024 * 1024;

const settings = plugin.settings.define({
  enablePdf: {
    type: "boolean",
    label: "启用 PDF 预览",
    description: "在工作区预览面板中直接渲染 .pdf 文件。",
    default: true
  },
  enableDocx: {
    type: "boolean",
    label: "启用 Word (.doc / .docx) 预览",
    description: "在工作区预览面板中直接显示 Word 文档。",
    default: true
  },
  enablePowerPoint: {
    type: "boolean",
    label: "启用 PowerPoint (.ppt / .pptx) 预览",
    description: "在工作区预览面板中逐页显示幻灯片。",
    default: true
  },
  defaultZoom: {
    type: "select",
    label: "默认缩放比例",
    description: "打开文档预览时的默认显示比例（默认自动适应当前面板宽度）。",
    default: "fit",
    options: [
      { value: "fit", label: "适应宽度（推荐）" },
      { value: "0.6", label: "60%" },
      { value: "0.75", label: "75%" },
      { value: "0.85", label: "85%" },
      { value: "1", label: "100%" },
      { value: "1.25", label: "125%" }
    ]
  }
});

let workerBlobUrl = null;
let disposed = false;
let scanQueued = false;
const mountedMap = new Map();
const activeRoots = new Set();
const observedDocs = new WeakSet();
const observers = [];
const loggedKeys = new Set();

function logOnce(key, level, msg) {
  if (loggedKeys.has(key)) return;
  if (loggedKeys.size > 200) loggedKeys.clear();
  loggedKeys.add(key);
  plugin.log[level](msg);
}

function ensurePdfWorker() {
  if (!workerBlobUrl && typeof PDF_WORKER_SOURCE === "string" && PDF_WORKER_SOURCE.length > 0) {
    const blob = new Blob([PDF_WORKER_SOURCE], { type: "application/javascript" });
    workerBlobUrl = URL.createObjectURL(blob);
    pdfjsLib.GlobalWorkerOptions.workerSrc = workerBlobUrl;
  }
}

function safeDecode(text) {
  if (!text) return "";
  try {
    return decodeURIComponent(text);
  } catch {
    return text;
  }
}

function ensureStylesInDoc(doc) {
  if (!doc || doc === document || !doc.head) return;
  if (doc.head.querySelector('style[data-bg-doc-preview-style="1"]')) return;
  const ownStyles = Array.from(
    document.querySelectorAll('style[data-bettergravity-plugin-style]')
  )
    .filter((s) => (s.textContent || "").includes("bg-doc-preview-root"))
    .map((s) => s.textContent || "")
    .join("\n");
  if (!ownStyles) return;
  const style = doc.createElement("style");
  style.setAttribute("data-bg-doc-preview-style", "1");
  style.textContent = ownStyles;
  doc.head.appendChild(style);
}

function detectExtension(fileName) {
  const lower = (fileName || "").toLowerCase();
  if (lower.endsWith(".pdf")) return "pdf";
  if (lower.endsWith(".docx")) return "docx";
  if (lower.endsWith(".doc")) return "doc";
  if (lower.endsWith(".pptx")) return "pptx";
  if (lower.endsWith(".ppt")) return "ppt";
  return null;
}

function isPresentation(ext) {
  return ext === "ppt" || ext === "pptx";
}

function previewEnabled(ext) {
  if (ext === "pdf") return settings.enablePdf;
  if (isPresentation(ext)) return settings.enablePowerPoint;
  return settings.enableDocx;
}

function collectRoots(doc, out = []) {
  if (!doc || !doc.body) return out;
  out.push({ doc, root: doc.body });
  const iframes = doc.querySelectorAll("iframe");
  for (const frame of iframes) {
    try {
      const childDoc = frame.contentDocument;
      if (childDoc && childDoc.body) {
        observeDocument(childDoc);
        collectRoots(childDoc, out);
      }
    } catch {
      // Cross-origin iframe
    }
  }
  return out;
}

const BINARY_PREVIEW_SELECTOR = '[data-testid="binary-file-no-preview"]';
const PDF_VIEW_SELECTOR = '[data-testid="artifact-pdf-view"]';
const NON_PREVIEW_SELECTOR = [
  "." + ROOT_CLASS,
  '[data-testid="conversation-view"]',
  '[data-testid="agent-input-box"]',
  '[data-testid="sidecars-view"]',
  '[role="article"]',
  '[role="tree"]',
  '.prose'
].join(", ");

function getElementFiber(el) {
  const key = Object.keys(el).find((name) => name.startsWith("__reactFiber"));
  const fiber = plugin.react?.getFiber?.(el) || (key ? el[key] : null);
  if (!fiber?.alternate) return fiber;
  let root = fiber;
  while (root.return) root = root.return;
  const current = root.stateNode?.current;
  return current && current === root.alternate ? fiber.alternate : fiber;
}

function getHostFileSource(box, maxDepth) {
  let fiber = getElementFiber(box);
  for (let depth = 0; fiber && depth < maxDepth; depth++, fiber = fiber.return) {
    const { fileUri, data, fileUrl } = fiber.memoizedProps || {};
    if (typeof fileUri !== "string") continue;
    const path = fileUri.split(/[?#]/, 1)[0];
    const fileName = safeDecode(path.split(/[\\/]/).pop());
    const ext = detectExtension(fileName);
    return ext ? { fileUri, fileName, ext, data, fileUrl } : null;
  }
  return null;
}

function getBinaryFileState(box, fileUri) {
  let fiber = getElementFiber(box);
  for (let depth = 0; fiber && depth < 18; depth++, fiber = fiber.return) {
    const ownerUri = fiber.memoizedProps?.fileUri;
    if (typeof ownerUri !== "string") continue;
    if (ownerUri !== fileUri) break;
    const fileState = fiber.memoizedProps?.fileContents;
    if (fileState?.type === "exists" && fileState.data?.kind === "binary") return fileState;
    for (let hook = fiber.memoizedState, count = 0; hook && count < 120; hook = hook.next, count++) {
      const state = hook.memoizedState;
      if (state?.type === "exists" && state.data?.kind === "binary") return state;
    }
  }
  return null;
}

function getPreviewUrl(value, doc) {
  if (!value || value.startsWith("#")) return "";
  try {
    const url = new URL(value, doc.baseURI);
    if (["blob:", "data:", "http:", "https:"].includes(url.protocol)) return url.href;
  } catch {}
  return "";
}

function collectDocumentPreviews(root, candidates) {
  for (const box of root.querySelectorAll(BINARY_PREVIEW_SELECTOR)) {
    if (box.closest(NON_PREVIEW_SELECTOR)) continue;
    const source = getHostFileSource(box, 12);
    if (!source) continue;
    const downloadEl = box.querySelector("a[href], button");
    const fileState = getBinaryFileState(box, source.fileUri);
    const fileUrl = getPreviewUrl(fileState?.data.blobUrl, box.ownerDocument)
      || getPreviewUrl(downloadEl?.getAttribute("href"), box.ownerDocument);
    candidates.set(box, { ...source, fileUrl, downloadEl });
  }

  for (const box of root.querySelectorAll(PDF_VIEW_SELECTOR)) {
    if (box.closest(NON_PREVIEW_SELECTOR)) continue;
    const source = getHostFileSource(box, 5);
    if (source?.ext === "pdf") candidates.set(box, source);
  }

  for (const label of root.querySelectorAll("p")) {
    if (label.textContent.trim() !== "Failed to read file content" || label.closest(NON_PREVIEW_SELECTOR)) continue;
    const titleRow = label.parentElement;
    const box = titleRow?.parentElement;
    const message = titleRow?.nextElementSibling?.textContent.trim();
    if (!box || !/^(?:File exceeds the \d+\s*MB display limit|File size \(\d+\s*MB\) exceeds \d+\s*MB display limit)\b/i.test(message || "")) continue;
    const source = getHostFileSource(box, 12);
    if (!source) continue;
    let fiber = getElementFiber(box);
    let revision = null;
    for (let depth = 0; fiber && depth < 6; depth++, fiber = fiber.return) {
      const state = fiber.memoizedProps?.fileContents;
      if (state?.type === "error") {
        revision = state;
        break;
      }
    }
    candidates.set(box, { ...source, message, revision, revealButton: box.querySelector("button") });
  }
}

function hasPdfHeader(buffer) {
  const header = new Uint8Array(buffer, 0, Math.min(1024, buffer.byteLength));
  return new TextDecoder().decode(header).includes("%PDF-");
}

function validateDocumentBuffer(buffer, ext) {
  if (ext === "pdf") {
    if (hasPdfHeader(buffer)) return buffer;
    throw new Error("读取的文件不是有效的 PDF");
  }
  const header = new Uint8Array(buffer, 0, Math.min(2, buffer.byteLength));
  if (header[0] === 0x50 && header[1] === 0x4b) return buffer;
  throw new Error("读取的文件不是有效的 DOCX");
}

async function resolveDocumentBuffer(box, source, signal, preferLocal) {
  const { data, ext, fileUri, message } = source;
  const doc = box.ownerDocument;
  const canReadLocal = LOCAL_PREVIEW_CONFIG && /^file:\/\/\//i.test(fileUri);
  const tryLocalFirst = preferLocal || ext === "doc" || isPresentation(ext);
  let sourceError;
  if (tryLocalFirst && canReadLocal) {
    try {
      return await readLargeLocalDocument(fileUri, ext, signal);
    } catch (err) {
      if (signal.aborted || message || (err.status !== 403 && !(err instanceof TypeError))) throw err;
      sourceError = err;
    }
  }
  try {
    if (ArrayBuffer.isView(data)) {
      const { buffer, byteOffset, byteLength } = data;
      const bytes = byteOffset === 0 && byteLength === buffer.byteLength
        ? buffer
        : buffer.slice(byteOffset, byteOffset + byteLength);
      return await prepareDocumentBuffer(bytes, ext, signal);
    }
    if (Object.prototype.toString.call(data) === "[object ArrayBuffer]") {
      return await prepareDocumentBuffer(data, ext, signal);
    }
    if (typeof data?.arrayBuffer === "function") {
      return await prepareDocumentBuffer(await data.arrayBuffer(), ext, signal);
    }
    const fileUrl = getPreviewUrl(source.fileUrl || (typeof data === "string" ? data : ""), doc);
    if (fileUrl) {
      const response = await fetch(fileUrl, { signal });
      if (!response.ok) throw new Error("宿主文件读取失败（HTTP " + response.status + "）");
      return await prepareDocumentBuffer(await response.arrayBuffer(), ext, signal);
    }
  } catch (err) {
    if (signal.aborted) throw err;
    sourceError = err;
  }

  if (canReadLocal && !tryLocalFirst) {
    try {
      return await readLargeLocalDocument(fileUri, ext, signal);
    } catch (err) {
      if (!message && !signal.aborted && err.status === 403) {
        throw sourceError || new Error("宿主尚未提供文件数据");
      }
      throw err;
    }
  }
  throw sourceError || new Error(message || "宿主尚未提供文件数据");
}

function hasSameDocument(previous, next) {
  const sameData = previous.data === next.data ||
    (ArrayBuffer.isView(previous.data) && ArrayBuffer.isView(next.data) &&
      previous.data.buffer === next.data.buffer && previous.data.byteOffset === next.data.byteOffset &&
      previous.data.byteLength === next.data.byteLength);
  return sameData && previous.fileUri === next.fileUri && previous.fileUrl === next.fileUrl &&
    previous.revision === next.revision && previous.message === next.message &&
    previous.downloadEl === next.downloadEl && previous.revealButton === next.revealButton;
}

async function fetchLocalPreview(endpoint, signal, options) {
  const { port, token } = LOCAL_PREVIEW_CONFIG;
  return fetch(new URL(endpoint, "http://127.0.0.1:" + port), {
    ...options,
    headers: { ...options?.headers, Authorization: "Bearer " + token },
    cache: "no-store",
    credentials: "omit",
    redirect: "error",
    signal
  });
}

function conversionError(status, ext) {
  const kind = isPresentation(ext) ? "PowerPoint" : "DOC";
  const messages = {
    413: `此 ${kind} 文件超过 128 MiB 的转换上限`,
    422: `无法解析此 ${kind} 文件，文件可能已加密、损坏或使用了不支持的格式`,
    429: "正在转换其他文档，请稍后重载",
    501: `${kind} 转换组件不可用`
  };
  const error = new Error(messages[status] || `本地预览组件无法转换此 ${kind} 文件`);
  error.status = status;
  return error;
}

async function prepareDocumentBuffer(buffer, ext, signal) {
  const presentation = isPresentation(ext);
  if (ext !== "doc" && !presentation) return validateDocumentBuffer(buffer, ext);
  if (!LOCAL_PREVIEW_CONFIG) throw new Error("请先启用本地文档转换组件");
  if (buffer.byteLength > 128 * 1024 * 1024) throw conversionError(413, ext);
  const endpoint = presentation ? "/presentation?" + new URLSearchParams({ format: ext }) : "/legacy-word";
  const response = await fetchLocalPreview(endpoint, signal, {
    method: "POST",
    headers: { "Content-Type": "application/octet-stream" },
    body: buffer
  });
  if (!response.ok) throw conversionError(response.status, ext);
  return validateDocumentBuffer(await response.arrayBuffer(), presentation ? "pdf" : "docx");
}

async function readLargeLocalDocument(fileUri, ext, signal) {
  const response = await fetchLocalPreview("/document?" + new URLSearchParams({ uri: fileUri }), signal);
  if (response.status === 403) {
    const error = new Error("此文件所在目录未启用大文件直开");
    error.status = response.status;
    throw error;
  }
  if (!response.ok) {
    if (ext === "doc" || isPresentation(ext)) throw conversionError(response.status, ext);
    throw new Error("本地预览组件无法读取此文件");
  }
  const contentType = response.headers.get("Content-Type");
  const format = isPresentation(ext) ? "pdf" : ext === "doc" ? (contentType === "application/pdf" ? "pdf" : "docx") : ext;
  return validateDocumentBuffer(await response.arrayBuffer(), format);
}

async function prepareDocxPreview(buffer, fileUri, signal) {
  if (!LOCAL_PREVIEW_CONFIG) return buffer;
  let hasEmbeddedFonts = false;
  try {
    const archive = await JSZip.loadAsync(buffer);
    hasEmbeddedFonts = archive.file(/^word\/fonts\/.+/i).length > 0;
    const hasMedia = archive.file(/^word\/media\/.+/i).length > 0;
    if ((hasEmbeddedFonts || hasMedia) && /^file:\/\//i.test(fileUri || "")) {
      const params = new URLSearchParams({ uri: fileUri, layout: "word" });
      const response = await fetchLocalPreview("/document?" + params, signal);
      if (response.ok && response.headers.get("Content-Type") === "application/pdf") {
        return validateDocumentBuffer(await response.arrayBuffer(), "pdf");
      }
    }
  } catch (error) {
    if (signal.aborted) throw error;
  }
  if (hasEmbeddedFonts) {
    try {
      const response = await fetchLocalPreview("/word-fonts", signal, {
        method: "POST",
        headers: { "Content-Type": "application/octet-stream" },
        body: buffer
      });
      if (response.ok) return validateDocumentBuffer(await response.arrayBuffer(), "docx");
    } catch (error) {
      if (signal.aborted) throw error;
    }
  }
  return buffer;
}

function createEl(doc, tag, className, text) {
  const el = doc.createElement(tag);
  if (className) el.className = className;
  if (text !== undefined) el.textContent = text;
  return el;
}

function mountPreview(fallbackBox, hostSource) {
  const doc = fallbackBox.ownerDocument;
  const parent = fallbackBox.parentElement;
  if (!parent) return;

  ensureStylesInDoc(doc);

  const { fileName, fileUri, ext, downloadEl } = hostSource;
  const presentation = isPresentation(ext);
  const isWord = ext === "doc" || ext === "docx";
  const badgeType = presentation ? "ppt" : isWord ? "docx" : "pdf";
  logOnce(
    "detect:" + fileUri,
    "info",
    'doc-preview detected file: "' + fileName + '" (' + ext + ")"
  );

  fallbackBox.classList.add(HIDDEN_CLASS);

  const root = createEl(doc, "div", ROOT_CLASS);
  const toolbar = createEl(doc, "div", "bg-doc-preview-toolbar");
  const titleGroup = createEl(doc, "div", "bg-doc-preview-title-group");
  const badge = createEl(
    doc,
    "span",
    `bg-doc-preview-badge bg-doc-preview-badge--${badgeType}`,
    presentation ? "PPT" : isWord ? "WORD" : "PDF"
  );
  const nameEl = createEl(doc, "span", "bg-doc-preview-filename", fileName);
  nameEl.title = fileName;
  const metaEl = createEl(doc, "span", "bg-doc-preview-meta", "加载中...");
  titleGroup.append(badge, nameEl, metaEl);

  const actions = createEl(doc, "div", "bg-doc-preview-actions");
  const zoomOutBtn = createEl(doc, "button", "bg-doc-preview-btn bg-doc-preview-btn--icon", "−");
  zoomOutBtn.type = "button";
  zoomOutBtn.title = "缩小";
  const initialSettingZoom = settings.defaultZoom || "fit";
  let fitMode = initialSettingZoom === "fit";
  let currentZoom = fitMode ? 1 : (Number(initialSettingZoom) || 1);
  const zoomLabel = createEl(
    doc,
    "span",
    "bg-doc-preview-zoom-label",
    fitMode ? "适应" : `${Math.round(currentZoom * 100)}%`
  );
  const zoomInBtn = createEl(doc, "button", "bg-doc-preview-btn bg-doc-preview-btn--icon", "+");
  zoomInBtn.type = "button";
  zoomInBtn.title = "放大";
  const fitBtn = createEl(doc, "button", "bg-doc-preview-btn", "适应");
  fitBtn.type = "button";
  fitBtn.title = "自动适应面板宽度";
  const resetZoomBtn = createEl(doc, "button", "bg-doc-preview-btn", "100%");
  resetZoomBtn.type = "button";
  const reloadBtn = createEl(doc, "button", "bg-doc-preview-btn", "重载");
  reloadBtn.type = "button";
  const downloadBtn = createEl(doc, "button", "bg-doc-preview-btn", "下载");
  downloadBtn.type = "button";

  actions.append(zoomOutBtn, zoomLabel, zoomInBtn, fitBtn, resetZoomBtn, reloadBtn);
  if (downloadEl) actions.append(downloadBtn);

  toolbar.append(titleGroup, actions);

  const viewport = createEl(doc, "div", "bg-doc-preview-viewport");
  root.append(toolbar, viewport);
  fallbackBox.insertAdjacentElement("afterend", root);

  let pdfLayoutZoom = currentZoom;
  let pdfStage = null;
  let pdfReRenderTimer = null;
  let refreshPdfCanvases = null;
  let resizeTimer = null;
  let currentBuffer = null;
  let pdfBuffer = null;
  let pdfLoadingTask = null;
  let pdfPageWidth = null;
  let pageObserver = null;
  let renderGeneration = 0;
  let loadGeneration = 0;
  let activeRead = null;
  let docxStage = null;
  let rootDisposed = false;
  const activePdfPageTasks = new Set();

  const updateZoomLabel = () => {
    zoomLabel.textContent = `${Math.round(currentZoom * 100)}%`;
  };

  const computeFitZoom = (naturalWidth) => {
    const avail = Math.max(1, (viewport.clientWidth || 560) - 28);
    if (!naturalWidth || naturalWidth <= 0) return 0.65;
    const ratio = avail / naturalWidth;
    return Math.min(1.5, Math.floor(ratio * 100) / 100 || ratio);
  };

  const cancelPdfPages = () => {
    pageObserver?.disconnect();
    pageObserver = null;
    for (const task of activePdfPageTasks) task.cancel();
    activePdfPageTasks.clear();
  };

  const cleanupPdf = () => {
    clearTimeout(pdfReRenderTimer);
    pdfReRenderTimer = null;
    refreshPdfCanvases = null;
    cancelPdfPages();
    if (pdfLoadingTask) void pdfLoadingTask.destroy().catch(() => {});
    pdfLoadingTask = null;
    pdfBuffer = null;
    pdfPageWidth = null;
    pdfStage = null;
  };

  const resetPreview = () => {
    renderGeneration++;
    cleanupPdf();
    currentBuffer = null;
    docxStage = null;
  };

  const renderDocx = async (buf) => {
    cleanupPdf();
    viewport.innerHTML = "";
    const stage = createEl(doc, "div", "bg-doc-preview-docx-stage");
    docxStage = stage;
    viewport.append(stage);

    await docx.renderAsync(buf, stage, null, {
      className: "bg-docx",
      inWrapper: true,
      ignoreWidth: false,
      ignoreHeight: false,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: true,
      renderAltChunks: false,
      useBase64URL: true
    });
    if (rootDisposed || stage !== docxStage) return;

    if (fitMode) {
      const pageEl = docxStage.querySelector("section.bg-docx") || docxStage.firstElementChild;
      const naturalW = pageEl?.offsetWidth || 794;
      currentZoom = computeFitZoom(naturalW);
    }
    docxStage.style.zoom = String(currentZoom);
    updateZoomLabel();
    metaEl.textContent = ext === "doc" ? "兼容预览" : "Word 文档";
  };

  const captureAnchor = (stage, clientX, clientY) => {
    if (!stage) return null;
    const vpRect = viewport.getBoundingClientRect();
    const cx = typeof clientX === "number" ? clientX : vpRect.left + vpRect.width / 2;
    const cy = typeof clientY === "number" ? clientY : vpRect.top + vpRect.height / 2;
    const stRect = stage.getBoundingClientRect();
    if (!stRect.width || !stRect.height) return null;
    return {
      clientX: cx,
      clientY: cy,
      relX: (cx - stRect.left) / stRect.width,
      relY: (cy - stRect.top) / stRect.height
    };
  };

  const restoreAnchor = (stage, anchor) => {
    if (!stage || !anchor) return;
    const stRect = stage.getBoundingClientRect();
    if (!stRect.width || !stRect.height) return;
    const targetScreenX = stRect.left + anchor.relX * stRect.width;
    const targetScreenY = stRect.top + anchor.relY * stRect.height;
    viewport.scrollLeft += targetScreenX - anchor.clientX;
    viewport.scrollTop += targetScreenY - anchor.clientY;
  };

  const renderPdf = async (buf) => {
    docxStage = null;
    const generation = ++renderGeneration;
    const isCurrent = () => !rootDisposed && generation === renderGeneration;
    if (pdfBuffer !== buf || !pdfLoadingTask) {
      cleanupPdf();
      ensurePdfWorker();
      pdfLoadingTask = pdfjsLib.getDocument({
        data: new Uint8Array(buf.slice(0)),
        isEvalSupported: false
      });
      pdfBuffer = buf;
    }

    const pdfDoc = await pdfLoadingTask.promise;
    if (!isCurrent()) return;
    const layoutZoom = fitMode ? 1 : currentZoom;
    const stage = createEl(doc, "div", "bg-doc-preview-pdf-stage");

    const cards = [];
    let widestPage = 0;
    for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
      const page = await pdfDoc.getPage(pageNum);
      if (!isCurrent()) return;

      const vp = page.getViewport({ scale: layoutZoom });
      widestPage = Math.max(widestPage, vp.width / layoutZoom);
      const card = createEl(doc, "div", "bg-pdf-page-card");
      card.style.width = `${Math.floor(vp.width)}px`;
      card.style.height = `${Math.floor(vp.height)}px`;
      card.dataset.pageNum = String(pageNum);

      const pageBadge = createEl(doc, "div", "bg-pdf-page-badge", `${pageNum} / ${pdfDoc.numPages}`);
      card.append(pageBadge);
      stage.append(card);
      cards.push({
        pageNum, card, page, vp,
        canvas: null, paintedZoom: 0, painting: false, textLayerStarted: false
      });
    }

    cancelPdfPages();
    pdfPageWidth = widestPage;
    if (fitMode) currentZoom = computeFitZoom(widestPage);
    pdfStage = stage;
    pdfLayoutZoom = layoutZoom;
    stage.style.zoom = String(currentZoom / pdfLayoutZoom);
    updateZoomLabel();
    metaEl.textContent = `共 ${pdfDoc.numPages} 页`;
    viewport.replaceChildren(stage);
    const visiblePages = new Set();
    const paintedPages = new Set();

    const markRecent = (entry) => {
      paintedPages.delete(entry);
      paintedPages.add(entry);
    };

    const trimOffscreenCanvases = () => {
      let bytes = 0;
      let count = 0;
      for (const entry of Array.from(paintedPages).reverse()) {
        if (visiblePages.has(entry)) continue;
        const canvasBytes = entry.canvas.width * entry.canvas.height * 4;
        if (count < PDF_OFFSCREEN_PAGE_LIMIT && bytes + canvasBytes <= PDF_OFFSCREEN_BYTE_LIMIT) {
          count++;
          bytes += canvasBytes;
          continue;
        }
        entry.canvas.remove();
        entry.canvas.width = 0;
        entry.canvas.height = 0;
        entry.canvas = null;
        entry.paintedZoom = 0;
        paintedPages.delete(entry);
      }
    };

    const renderTextLayer = async (entry) => {
      if (typeof pdfjsLib.renderTextLayer !== "function") return;
      try {
        const textContent = await entry.page.getTextContent();
        if (!isCurrent()) return;
        const textLayer = createEl(doc, "div", "bg-pdf-text-layer");
        textLayer.style.setProperty("--scale-factor", String(layoutZoom));
        entry.card.append(textLayer);
        const task = pdfjsLib.renderTextLayer({
          textContentSource: textContent,
          container: textLayer,
          viewport: entry.vp,
          textDivs: []
        });
        activePdfPageTasks.add(task);
        try {
          await task.promise;
        } finally {
          activePdfPageTasks.delete(task);
        }
      } catch (err) {
        if (isCurrent()) plugin.log.warn(`doc-preview pdf page ${entry.pageNum} text layer error: ${err?.message}`);
      }
    };

    const renderPageCard = async (entry) => {
      const zoom = currentZoom;
      if (!isCurrent() || entry.painting || entry.paintedZoom === zoom) return;
      entry.painting = true;
      try {
        const dpr = Math.min(window.devicePixelRatio || 1, 2);
        const renderVp = entry.page.getViewport({ scale: zoom * dpr });
        const canvas = doc.createElement("canvas");
        canvas.width = Math.floor(renderVp.width);
        canvas.height = Math.floor(renderVp.height);

        const task = entry.page.render({ canvasContext: canvas.getContext("2d"), viewport: renderVp });
        activePdfPageTasks.add(task);
        try {
          await task.promise;
        } finally {
          activePdfPageTasks.delete(task);
        }
        if (!isCurrent()) return;

        // Keep the painted canvas in place until the replacement is ready.
        if (entry.canvas) {
          if (zoom !== currentZoom) return;
          entry.canvas.replaceWith(canvas);
        } else {
          entry.card.prepend(canvas);
          if (!entry.textLayerStarted) {
            entry.textLayerStarted = true;
            void renderTextLayer(entry);
          }
        }
        entry.canvas = canvas;
        entry.paintedZoom = zoom;
        markRecent(entry);
        trimOffscreenCanvases();
      } catch (err) {
        if (isCurrent()) plugin.log.warn(`doc-preview pdf page ${entry.pageNum} error: ${err?.message}`);
      } finally {
        entry.painting = false;
        if (isCurrent() && !pdfReRenderTimer && zoom !== currentZoom && visiblePages.has(entry)) {
          void renderPageCard(entry);
        }
      }
    };

    refreshPdfCanvases = () => {
      for (const entry of visiblePages) void renderPageCard(entry);
    };

    const cardMap = new Map(cards.map((c) => [c.card, c]));
    pageObserver = new IntersectionObserver(
      (entries) => {
        for (const item of entries) {
          const entry = cardMap.get(item.target);
          if (item.isIntersecting) {
            visiblePages.add(entry);
            if (entry.canvas) markRecent(entry);
            if (!pdfReRenderTimer || !entry.canvas) void renderPageCard(entry);
          } else {
            visiblePages.delete(entry);
          }
        }
        trimOffscreenCanvases();
      },
      { root: viewport, rootMargin: "500px 0px" }
    );

    for (const entry of cards) pageObserver.observe(entry.card);
    for (const entry of cards.slice(0, 2)) void renderPageCard(entry);
  };

  const renderFromBuffer = async (buf) => {
    if (rootDisposed) return;
    currentBuffer = buf;
    if (isWord && !hasPdfHeader(buf)) {
      await renderDocx(buf);
    } else {
      await renderPdf(buf);
    }
  };

  const showFallbackPicker = (reason) => {
    if (rootDisposed) return;
    resetPreview();
    metaEl.textContent = "需要选择文件";
    viewport.innerHTML = "";

    const box = createEl(doc, "div", "bg-doc-preview-state");
    const title = createEl(doc, "div", "bg-doc-preview-state-title", `无法自动读取「${fileName}」的文件流`);
    const desc = createEl(
      doc,
      "div",
      "bg-doc-preview-state-desc",
      `${reason}。你可以直接把文件拖到这里，或点击下方按钮选择该文件立即预览。`
    );
    const row = createEl(doc, "div", "bg-doc-preview-state-actions");

    const fileInput = doc.createElement("input");
    fileInput.type = "file";
    fileInput.accept = "." + ext;
    fileInput.style.display = "none";

    const pickBtn = createEl(doc, "button", "bg-doc-preview-btn bg-doc-preview-btn--primary", "选择文件预览");
    pickBtn.type = "button";
    pickBtn.addEventListener("click", () => fileInput.click());

    const renderSelectedFile = async (file) => {
      if (!file) return;
      const generation = ++loadGeneration;
      resetPreview();
      activeRead?.abort();
      activeRead = new AbortController();
      const signal = AbortSignal.any([activeRead.signal, AbortSignal.timeout(120_000)]);
      try {
        metaEl.textContent = "解析中...";
        if ((ext === "doc" || presentation) && file.size > 128 * 1024 * 1024) throw conversionError(413, ext);
        let buffer = await prepareDocumentBuffer(await file.arrayBuffer(), ext, signal);
        if (ext === "docx") buffer = await prepareDocxPreview(buffer, null, signal);
        if (generation === loadGeneration) await renderFromBuffer(buffer);
      } catch (err) {
        if (generation === loadGeneration) showFallbackPicker(err?.message || String(err));
      }
    };
    fileInput.addEventListener("change", () => void renderSelectedFile(fileInput.files?.[0]));

    row.append(pickBtn, fileInput);
    if (downloadEl) {
      const dlBtn = createEl(doc, "button", "bg-doc-preview-btn", "触发原下载");
      dlBtn.type = "button";
      dlBtn.addEventListener("click", () => downloadEl.click());
      row.append(dlBtn);
    }
    if (hostSource?.revealButton) {
      const revealBtn = createEl(doc, "button", "bg-doc-preview-btn", "在文件资源管理器中显示");
      revealBtn.type = "button";
      revealBtn.addEventListener("click", () => hostSource.revealButton.click());
      row.append(revealBtn);
    }

    box.addEventListener("dragover", (e) => {
      e.preventDefault();
      box.classList.add("bg-doc-preview-state--dragover");
    });
    box.addEventListener("dragleave", () => {
      box.classList.remove("bg-doc-preview-state--dragover");
    });
    box.addEventListener("drop", (e) => {
      e.preventDefault();
      box.classList.remove("bg-doc-preview-state--dragover");
      void renderSelectedFile(e.dataTransfer?.files?.[0]);
    });

    box.append(title, desc, row);
    viewport.append(box);
  };

  const loadAndRender = async (preferLocal = false, generation = ++loadGeneration, retryCount = 0) => {
    activeRead?.abort();
    const largeFile = Boolean(hostSource.message);
    if (largeFile && (!LOCAL_PREVIEW_CONFIG || !/^file:\/\/\//i.test(fileUri))) {
      showFallbackPicker(hostSource.message);
      return;
    }
    resetPreview();
    metaEl.textContent = "加载中...";
    const loading = createEl(doc, "div", "bg-doc-preview-state");
    loading.append(createEl(doc, "div", "bg-doc-preview-state-title", `正在解析 ${fileName} ...`));
    viewport.replaceChildren(loading);
    activeRead = new AbortController();
    const signal = AbortSignal.any([activeRead.signal, AbortSignal.timeout(120_000)]);
    try {
      let buf = await resolveDocumentBuffer(fallbackBox, hostSource, signal, preferLocal);
      if (rootDisposed || generation !== loadGeneration) return;
      if (ext === "docx") buf = await prepareDocxPreview(buf, fileUri, signal);
      if (rootDisposed || generation !== loadGeneration) return;
      await renderFromBuffer(buf);
      if (rootDisposed || generation !== loadGeneration) return;
      plugin.log.info(`doc-preview rendered ${fileName} (${buf.byteLength} bytes)`);
    } catch (err) {
      if (rootDisposed || generation !== loadGeneration) return;
      const retryLimit = [413, 422, 501].includes(err.status) ? 0
        : largeFile ? (err instanceof TypeError ? 5 : 0) : 3;
      if (retryCount < retryLimit) {
        setTimeout(() => {
          if (!rootDisposed && generation === loadGeneration) void loadAndRender(preferLocal, generation, retryCount + 1);
        }, 450);
        return;
      }
      plugin.log.warn(`doc-preview failed for ${fileName}: ${err?.message}`);
      showFallbackPicker(largeFile && err instanceof TypeError
        ? "暂时无法连接本地预览组件"
        : (err?.message || "未知错误"));
    }
  };

  const applyZoom = (nextZoom, keepFit = false, clientX, clientY) => {
    fitMode = keepFit;
    const clamped = keepFit ? nextZoom : Math.min(3.0, Math.max(0.25, Math.round(nextZoom * 100) / 100));
    if (clamped === currentZoom) return;
    const activeStage = docxStage ?? pdfStage;
    const anchor = captureAnchor(activeStage, clientX, clientY);

    currentZoom = clamped;
    updateZoomLabel();
    if (!currentBuffer) return;

    if (docxStage) {
      docxStage.style.zoom = String(currentZoom);
      if (anchor) restoreAnchor(docxStage, anchor);
    } else if (pdfStage) {
      pdfStage.style.zoom = String(currentZoom / pdfLayoutZoom);
      if (anchor) restoreAnchor(pdfStage, anchor);
      clearTimeout(pdfReRenderTimer);
      pdfReRenderTimer = setTimeout(() => {
        pdfReRenderTimer = null;
        refreshPdfCanvases?.();
      }, 180);
    }
  };

  const fitToWidth = () => {
    fitMode = true;
    if (!currentBuffer) return;
    if (docxStage) {
      const pageEl = docxStage.querySelector("section.bg-docx") || docxStage.firstElementChild;
      const naturalW = pageEl?.offsetWidth || 794;
      applyZoom(computeFitZoom(naturalW), true);
    } else if (pdfPageWidth) {
      applyZoom(computeFitZoom(pdfPageWidth), true);
    }
  };

  const onWheel = (e) => {
    if (!e.ctrlKey && !e.metaKey) return;
    e.preventDefault();
    e.stopPropagation();
    if (!e.deltaY) return;
    const delta = Math.abs(e.deltaY) >= 40
      ? (e.deltaY < 0 ? 0.1 : -0.1)
      : Math.max(-0.2, Math.min(0.2, -e.deltaY * 0.005));
    applyZoom(currentZoom + delta, false, e.clientX, e.clientY);
  };

  const resizeObserver = typeof ResizeObserver === "function"
    ? new ResizeObserver(() => {
        if (!fitMode || rootDisposed || !currentBuffer) return;
        if (resizeTimer) clearTimeout(resizeTimer);
        resizeTimer = setTimeout(() => {
          resizeTimer = null;
          if (fitMode && !rootDisposed && currentBuffer) fitToWidth();
        }, 150);
      })
    : null;
  resizeObserver?.observe(viewport);

  root.addEventListener("wheel", onWheel, { passive: false, capture: true });
  zoomOutBtn.addEventListener("click", () => applyZoom(currentZoom - 0.1));
  zoomInBtn.addEventListener("click", () => applyZoom(currentZoom + 0.1));
  fitBtn.addEventListener("click", () => fitToWidth());
  resetZoomBtn.addEventListener("click", () => applyZoom(1));
  reloadBtn.addEventListener("click", () => void loadAndRender(true));
  downloadBtn.addEventListener("click", () => downloadEl?.click());

  const dispose = () => {
    rootDisposed = true;
    activeRead?.abort();
    resetPreview();
    resizeObserver?.disconnect();
    if (resizeTimer) {
      clearTimeout(resizeTimer);
      resizeTimer = null;
    }
    root.removeEventListener("wheel", onWheel, { capture: true });
    if (root.isConnected) root.remove();
    fallbackBox.classList.remove(HIDDEN_CLASS);
    activeRoots.delete(dispose);
  };

  activeRoots.add(dispose);
  mountedMap.set(fallbackBox, { source: hostSource, root, dispose });
  void loadAndRender();
}

function scan() {
  if (disposed) return;
  const candidates = new Map();
  for (const { root } of collectRoots(document)) collectDocumentPreviews(root, candidates);

  for (const [box, entry] of mountedMap) {
    const source = candidates.get(box);
    if (!source || !hasSameDocument(entry.source, source) ||
        entry.root.parentElement !== box.parentElement || !entry.root.isConnected ||
        !previewEnabled(source.ext)) {
      entry.dispose();
      mountedMap.delete(box);
    }
  }
  for (const [box, source] of candidates) {
    if (mountedMap.has(box) || !box.parentElement) continue;
    if (!previewEnabled(source.ext)) continue;
    mountPreview(box, source);
  }
}

function scheduleScan() {
  if (disposed || scanQueued) return;
  scanQueued = true;
  requestAnimationFrame(() => {
    scanQueued = false;
    scan();
  });
}

function observeDocument(doc) {
  if (!doc || !doc.body || observedDocs.has(doc)) return;
  observedDocs.add(doc);
  const obs = new MutationObserver((mutations) => {
    for (const m of mutations) {
      if (m.target instanceof Element && m.target.closest("." + ROOT_CLASS)) continue;
      scheduleScan();
      break;
    }
  });
  obs.observe(doc.body, {
    childList: true,
    subtree: true,
    characterData: true,
    attributes: true,
    attributeFilter: ["href", "src"]
  });
  observers.push(obs);
}

if (document.body) {
  observeDocument(document);
  scan();
} else {
  window.addEventListener("DOMContentLoaded", () => {
    observeDocument(document);
    scan();
  });
}

// Fallback lightweight poll in case an iframe loads asynchronously
const pollTimer = setInterval(() => {
  if (!disposed) scan();
}, 1500);

plugin.log.info("doc-preview plugin started (v1.1)");

plugin.onDispose(() => {
  disposed = true;
  clearInterval(pollTimer);
  for (const obs of observers) obs.disconnect();
  observers.length = 0;
  for (const fn of [...activeRoots]) fn();
  if (workerBlobUrl) {
    URL.revokeObjectURL(workerBlobUrl);
    workerBlobUrl = null;
  }
  plugin.log.info("doc-preview plugin disposed");
});
