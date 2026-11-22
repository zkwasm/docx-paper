const core = require('./core/openxml-engine');
const { convertWmfToSvg, SYMBOL_CHAR_MAP, MT_EXTRA_CHAR_MAP } = require('./vector/wmf-to-svg');
const { convertEmfToBlob, dibToBmpDataUrl } = require('./vector/emf-to-svg');
const { injectPaperStyles, createPaperPaginator } = require('./renderer/paginator');
const { renderMarkdownAsync, latexToMathML, compileMarkdownToHtmlBlocks } = require('./markdown/md-to-paper');

function isMarkdownInput(data, options = {}) {
  if (options.format === 'md' || options.format === 'markdown') return true;
  if (options.format === 'docx') return false;
  if (typeof data === 'string') return true;
  if (data instanceof ArrayBuffer || ArrayBuffer.isView(data)) {
    const u8 = data instanceof ArrayBuffer ? new Uint8Array(data) : new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
    // ZIP magic bytes for .docx: 0x50 0x4B ('PK')
    if (u8.byteLength >= 2 && u8[0] === 0x50 && u8[1] === 0x4b) return false;
    return true;
  }
  return false;
}

/**
 * High-level API: Render a .docx or .md (Markdown) file into a DOM container
 * with automatic physical page card splitting, WMF/EMF/OMML/LaTeX formula rendering, and print support.
 *
 * @param {Blob|ArrayBuffer|Uint8Array|string} data - The .docx binary or .md text/binary data
 * @param {HTMLElement} container - Target DOM element to render pages into
 * @param {Object} [options] - Rendering & pagination options
 * @returns {Promise<Object>} Interactive viewer controller (nextPage, prevPage, scrollToPage, setZoom, print, destroy)
 */
async function renderDocx(data, container, options = {}) {
  injectPaperStyles();

  if (isMarkdownInput(data, options)) {
    await renderMarkdownAsync(data, container, options.styleContainer || null, options.renderOptions || {});
  } else {
    const renderOptions = {
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
      renderEndnotes: true,
      ...options.renderOptions
    };
    await core.renderAsync(data, container, options.styleContainer || null, renderOptions);
  }

  if (options.paginate !== false) {
    return createPaperPaginator(container, options);
  }

  return {
    get pageCount() {
      return container.querySelectorAll('section.docx').length;
    },
    get currentPage() {
      return 1;
    },
    get formulaCount() {
      return container.querySelectorAll('.docx-vml-img, math, mjx-container').length;
    },
    scrollToPage: () => {},
    nextPage: () => {},
    prevPage: () => {},
    setZoom: (z) => {
      container.style.zoom = String(z);
      return z;
    },
    flushPages: () => {},
    print: () => window.print(),
    destroy: () => {}
  };
}

module.exports = {
  ...core,
  renderDocx,
  renderMarkdownAsync,
  latexToMathML,
  compileMarkdownToHtmlBlocks,
  convertWmfToSvg,
  convertEmfToBlob,
  dibToBmpDataUrl,
  SYMBOL_CHAR_MAP,
  MT_EXTRA_CHAR_MAP,
  injectPaperStyles,
  createPaperPaginator
};
