const fs = require('fs');
const path = require('path');

const pkgRoot = path.resolve(__dirname, '..');
const srcDir = path.join(pkgRoot, 'src');
const coreDir = path.join(srcDir, 'core');
const distDir = path.join(pkgRoot, 'dist');
const NL = String.fromCharCode(10);

fs.mkdirSync(coreDir, { recursive: true });
fs.mkdirSync(distDir, { recursive: true });

// Read the core OpenXML engine from src/core/openxml-engine.js
const engineContent = fs.readFileSync(path.join(coreDir, 'openxml-engine.js'), 'utf8');

// Read paginator and markdown modules
const paginatorCode = fs
  .readFileSync(path.join(srcDir, 'renderer/paginator.js'), 'utf8')
  .replace(/module\.exports\s*=\s*\{[\s\S]*?\};\s*$/, '');

const markdownCode = fs
  .readFileSync(path.join(srcDir, 'markdown/md-to-paper.js'), 'utf8')
  .replace(/module\.exports\s*=\s*\{[\s\S]*?\};\s*$/, '');

const highLevelAttachment = `
;(function(root) {
${paginatorCode}

${markdownCode}

  function isMarkdownInput(data, options) {
    options = options || {};
    if (options.format === 'md' || options.format === 'markdown') return true;
    if (options.format === 'docx') return false;
    if (typeof data === 'string') return true;
    if (data instanceof ArrayBuffer || ArrayBuffer.isView(data)) {
      var u8 = data instanceof ArrayBuffer ? new Uint8Array(data) : new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
      if (u8.byteLength >= 2 && u8[0] === 0x50 && u8[1] === 0x4b) return false;
      return true;
    }
    return false;
  }

  async function renderDocx(data, container, options = {}) {
    injectPaperStyles();
    const targetApi = (typeof exports === 'object' && typeof module !== 'undefined') ? exports : (root.docx || root.DocxPaper);
    if (isMarkdownInput(data, options)) {
      await renderMarkdownAsync(data, container, options.styleContainer || null, options.renderOptions || {});
    } else {
      const renderOptions = Object.assign({
        className: 'docx',
        inWrapper: true,
        ignoreWidth: false,
        ignoreHeight: false,
        ignoreFonts: false,
        breakPages: true,
        ignoreLastRenderedPageBreak: true,
        experimental: true,
        trimXmlDeclaration: true,
        renderHeaders: true,
        renderFooters: true,
        renderFootnotes: true,
        renderEndnotes: true
      }, options.renderOptions || {});
      await targetApi.renderAsync(data, container, options.styleContainer || null, renderOptions);
    }

    if (options.paginate !== false) {
      return createPaperPaginator(container, options);
    }
    return {
      get pageCount() { return container.querySelectorAll('section.docx').length; },
      get currentPage() { return 1; },
      get formulaCount() { return container.querySelectorAll('.docx-vml-img, math, mjx-container').length; },
      scrollToPage: function() {},
      nextPage: function() {},
      prevPage: function() {},
      setZoom: function(z) { container.style.zoom = String(z); return z; },
      flushPages: function() {},
      print: function() { if (typeof window !== 'undefined') window.print(); },
      destroy: function() {}
    };
  }

  if (typeof exports === 'object' && typeof module !== 'undefined') {
    exports.renderDocx = renderDocx;
    exports.renderMarkdownAsync = renderMarkdownAsync;
    exports.latexToMathML = latexToMathML;
    exports.createPaperPaginator = createPaperPaginator;
    exports.injectPaperStyles = injectPaperStyles;
  }
  if (root) {
    if (root.docx) {
      root.docx.renderDocx = renderDocx;
      root.docx.renderMarkdownAsync = renderMarkdownAsync;
      root.docx.latexToMathML = latexToMathML;
      root.docx.createPaperPaginator = createPaperPaginator;
      root.docx.injectPaperStyles = injectPaperStyles;
      root.DocxPaper = root.docx;
    }
  }
})(typeof globalThis !== 'undefined' ? globalThis : (typeof self !== 'undefined' ? self : this));
`;

const umdBundle = engineContent + NL + highLevelAttachment;
fs.writeFileSync(path.join(distDir, 'docx-paper.js'), umdBundle, 'utf8');

// Also update the web app's public/vendor/docx-preview.js with the unified bundle
const webAppVendorPath = path.resolve(pkgRoot, '../../public/vendor/docx-preview.js');
if (fs.existsSync(path.dirname(webAppVendorPath))) {
  fs.writeFileSync(webAppVendorPath, umdBundle, 'utf8');
}

// Generate ES Module wrapper
const esmBundle = `import './docx-paper.js';
const api = globalThis.DocxPaper || globalThis.docx;
export const renderDocx = api.renderDocx;
export const renderMarkdownAsync = api.renderMarkdownAsync;
export const latexToMathML = api.latexToMathML;
export const renderAsync = api.renderAsync;
export const parseAsync = api.parseAsync;
export const createPaperPaginator = api.createPaperPaginator;
export const injectPaperStyles = api.injectPaperStyles;
export default api;
`;
fs.writeFileSync(path.join(distDir, 'docx-paper.esm.js'), esmBundle, 'utf8');

// Generate TypeScript declarations
const dtsContent = `export interface DocxPaperPageEvent {
  page: number;
  loadedPages: number;
  total: number;
  formulaCount: number;
}

export interface DocxPaperOptions {
  /** Document format ('docx' | 'md' | 'auto'). Defaults to auto-detection. */
  format?: 'docx' | 'md' | 'markdown' | 'auto';
  /** Scrollable parent element for tracking active page on scroll */
  scrollContainer?: HTMLElement;
  /** Automatically split sections into discrete physical paper cards (default: true) */
  paginate?: boolean;
  /** Render first 2 pages immediately and stream remaining pages in background (default: true) */
  progressive?: boolean;
  /** Number of pages to mount synchronously on first frame (default: 2) */
  immediatePages?: number;
  /** Show page number badge in top-right corner of each paper card (default: true) */
  showPageBadge?: boolean;
  /** Callback fired when active page or loaded page count changes */
  onPageChange?: (event: DocxPaperPageEvent) => void;
  /** Optional container for injected <style> elements */
  styleContainer?: HTMLElement;
  /** Low-level OpenXML rendering options */
  renderOptions?: Record<string, any>;
}

export interface DocxPaperController {
  readonly pageCount: number;
  readonly currentPage: number;
  readonly formulaCount: number;
  scrollToPage(pageNumber: number): void;
  nextPage(): void;
  prevPage(): void;
  setZoom(zoom: number): number;
  flushPages(): void;
  print(): void;
  destroy(): void;
}

export function renderDocx(
  data: Blob | ArrayBuffer | Uint8Array | string,
  container: HTMLElement,
  options?: DocxPaperOptions
): Promise<DocxPaperController>;

export function renderMarkdownAsync(
  data: Blob | ArrayBuffer | Uint8Array | string,
  bodyContainer: HTMLElement,
  styleContainer?: HTMLElement | null,
  options?: Record<string, any>
): Promise<void>;

export function latexToMathML(tex: string, displayMode?: boolean): string;

export function renderAsync(
  data: Blob | ArrayBuffer | Uint8Array,
  bodyContainer: HTMLElement,
  styleContainer?: HTMLElement | null,
  options?: Record<string, any>
): Promise<any>;

export function parseAsync(
  data: Blob | ArrayBuffer | Uint8Array,
  options?: Record<string, any>
): Promise<any>;

export function createPaperPaginator(
  container: HTMLElement,
  options?: DocxPaperOptions
): DocxPaperController;

export function injectPaperStyles(): void;
`;
fs.writeFileSync(path.join(distDir, 'docx-paper.d.ts'), dtsContent, 'utf8');

console.log('Built docx-paper (with Markdown + LaTeX Math support) into packages/docx-paper/dist/');
