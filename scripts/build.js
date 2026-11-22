const fs = require('fs');
const path = require('path');

const pkgRoot = path.resolve(__dirname, '..');
const srcDir = path.join(pkgRoot, 'src');
const coreDir = path.join(srcDir, 'core');
const distDir = path.join(pkgRoot, 'dist');

fs.mkdirSync(coreDir, { recursive: true });
fs.mkdirSync(distDir, { recursive: true });

// Copy our enhanced core engine into src/core/openxml-engine.js
const engineSourcePath = path.resolve(pkgRoot, '../../public/vendor/docx-preview.js');
const engineContent = fs.readFileSync(engineSourcePath, 'utf8');
fs.writeFileSync(path.join(coreDir, 'openxml-engine.js'), engineContent, 'utf8');

// Read paginator to append `renderDocx` and `createPaperPaginator` onto the UMD/ESM exports
const paginatorCode = fs
  .readFileSync(path.join(srcDir, 'renderer/paginator.js'), 'utf8')
  .replace(/module\.exports\s*=\s*\{[\s\S]*?\};\s*$/, '');

const highLevelAttachment = `
;(function(root) {
${paginatorCode}

  async function renderDocx(data, container, options = {}) {
    injectPaperStyles();
    const targetApi = (typeof exports === 'object' && typeof module !== 'undefined') ? exports : (root.docx || root.DocxPaper);
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
    exports.createPaperPaginator = createPaperPaginator;
    exports.injectPaperStyles = injectPaperStyles;
  }
  if (root) {
    if (root.docx) {
      root.docx.renderDocx = renderDocx;
      root.docx.createPaperPaginator = createPaperPaginator;
      root.docx.injectPaperStyles = injectPaperStyles;
      root.DocxPaper = root.docx;
    }
  }
})(typeof globalThis !== 'undefined' ? globalThis : (typeof self !== 'undefined' ? self : this));
`;

const umdBundle = engineContent + String.fromCharCode(10) + highLevelAttachment;
fs.writeFileSync(path.join(distDir, 'docx-paper.js'), umdBundle, 'utf8');

// Generate ES Module wrapper
const esmBundle = `import './docx-paper.js';
const api = globalThis.DocxPaper || globalThis.docx;
export const renderDocx = api.renderDocx;
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
  data: Blob | ArrayBuffer | Uint8Array,
  container: HTMLElement,
  options?: DocxPaperOptions
): Promise<DocxPaperController>;

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

console.log('Built docx-paper successfully into packages/docx-paper/dist/');
