const core = require('./core/openxml-engine');
const { convertWmfToSvg, SYMBOL_CHAR_MAP, MT_EXTRA_CHAR_MAP } = require('./vector/wmf-to-svg');
const { convertEmfToBlob, dibToBmpDataUrl } = require('./vector/emf-to-svg');
const { injectPaperStyles, createPaperPaginator } = require('./renderer/paginator');

/**
 * High-level API: Render a .docx Blob / ArrayBuffer / Uint8Array into a DOM container
 * with automatic physical page card splitting, WMF/EMF/OMML formula rendering, and print support.
 *
 * @param {Blob|ArrayBuffer|Uint8Array} data - The .docx file binary data
 * @param {HTMLElement} container - Target DOM element to render pages into
 * @param {Object} [options] - Rendering & pagination options
 * @returns {Promise<Object>} Interactive viewer controller (nextPage, prevPage, scrollToPage, setZoom, print, destroy)
 */
async function renderDocx(data, container, options = {}) {
  injectPaperStyles();

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
  convertWmfToSvg,
  convertEmfToBlob,
  dibToBmpDataUrl,
  SYMBOL_CHAR_MAP,
  MT_EXTRA_CHAR_MAP,
  injectPaperStyles,
  createPaperPaginator
};
