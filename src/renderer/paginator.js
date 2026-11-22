/**
 * Progressive Paper Paginator & Viewer Controller for docx-paper
 * Splits continuous sections into discrete physical paper cards (A4/B5/Letter)
 * with gaps between pages, real-time page tracking, and print stylesheet support.
 */

let paperStylesInjected = false;

function injectPaperStyles() {
  if (paperStylesInjected || typeof document === 'undefined') return;
  paperStylesInjected = true;
  const style = document.createElement('style');
  style.id = 'docx-paper-builtin-styles';
  style.textContent = `
.docx-wrapper {
  background: transparent !important;
  padding: 0 !important;
}
.docx-wrapper > section.docx {
  background: #ffffff !important;
  border: 1px solid #cbd5e1 !important;
  box-shadow: 0 8px 24px rgba(15, 23, 42, 0.14), 0 2px 6px rgba(15, 23, 42, 0.08) !important;
  margin-bottom: 32px !important;
  border-radius: 3px;
  position: relative;
}
.docx-wrapper math,
.docx-wrapper mjx-container {
  font-family: "Cambria Math", "Latin Modern Math", "STIX Two Math", "Apple Symbols", serif !important;
  font-size: 1.05em;
  vertical-align: middle;
}
.docx-wrapper .docx-math-block {
  display: block;
  text-align: center;
  margin: 0.45em 0;
}
.docx-wrapper mfrac {
  padding: 0 2px;
}
.docx-vml-picture {
  display: inline;
}
.docx-vml-img {
  vertical-align: middle;
  display: inline-block;
  margin: 2pt 1pt;
}
.docx-page-badge {
  position: absolute;
  right: 10px;
  top: 10px;
  background: rgba(241, 245, 249, 0.92);
  color: #475569;
  border: 1px solid #cbd5e1;
  font-size: 11px;
  font-weight: 600;
  padding: 2px 8px;
  border-radius: 999px;
  pointer-events: none;
  z-index: 6;
  font-variant-numeric: tabular-nums;
}
@media print {
  @page {
    size: auto;
    margin: 14mm 0;
  }
  .docx-page-badge,
  .docx-page-guide,
  .no-print {
    display: none !important;
  }
  .docx-wrapper > section.docx {
    box-shadow: none !important;
    border: none !important;
    margin: 0 auto !important;
    padding-top: 0 !important;
    padding-bottom: 0 !important;
    min-height: auto !important;
    overflow: visible !important;
    page-break-after: always;
    break-after: page;
  }
  .docx-wrapper > section.docx:last-child {
    page-break-after: auto;
    break-after: auto;
  }
  .docx tr,
  .docx-vml-img,
  .docx math,
  .docx img {
    break-inside: avoid;
    page-break-inside: avoid;
  }
}
`;
  document.head.appendChild(style);
}

function parsePtToPx(styleVal, fallbackPx = 1122.5) {
  if (!styleVal) return fallbackPx;
  const num = parseFloat(styleVal);
  if (!num || Number.isNaN(num)) return fallbackPx;
  if (String(styleVal).endsWith('pt')) return num * (96 / 72);
  if (String(styleVal).endsWith('px')) return num;
  return fallbackPx;
}

function isVisibleBlockElement(el) {
  if (!el) return false;
  if (el.tagName === 'TABLE') return true;
  if (el.textContent && el.textContent.trim().length > 0) return true;
  if (el.querySelector && el.querySelector('img, svg, table, math, mjx-container, canvas')) {
    return true;
  }
  return false;
}

function createPageSectionFromBucket(templateSec, templateArticle, headerEl, footerEl, bucket) {
  const newSec = document.createElement('section');
  newSec.className = templateSec.className;
  newSec.setAttribute('style', templateSec.getAttribute('style') || '');

  if (headerEl) {
    newSec.appendChild(headerEl.cloneNode(true));
  }
  const newArticle = document.createElement('article');
  newArticle.className = templateArticle.className;
  newArticle.setAttribute('style', templateArticle.getAttribute('style') || '');
  for (const item of bucket) {
    newArticle.appendChild(item);
  }
  newSec.appendChild(newArticle);

  if (footerEl) {
    newSec.appendChild(footerEl.cloneNode(true));
  }
  return newSec;
}

function createPaperPaginator(container, options = {}) {
  injectPaperStyles();

  const scrollContainer =
    options.scrollContainer || container.parentElement || document.documentElement;
  const showPageBadge = options.showPageBadge !== false;
  const immediatePages = options.immediatePages || 2;
  const onPageChange = typeof options.onPageChange === 'function' ? options.onPageChange : null;

  let pendingPageTasks = [];
  let virtualPages = [];
  let currentPageNum = 1;
  let currentZoom = 1;
  let destroyed = false;

  function refreshMetrics() {
    if (destroyed) return;
    container
      .querySelectorAll('.docx-page-guide, .docx-page-badge')
      .forEach((el) => el.remove());

    const sections = Array.from(container.querySelectorAll('section.docx'));
    const total = sections.length + pendingPageTasks.length;

    virtualPages = sections.map((sec, idx) => ({
      sec,
      topPx: sec.offsetTop,
      pageNum: idx + 1,
      total
    }));

    if (showPageBadge) {
      virtualPages.forEach((vp) => {
        const badge = document.createElement('div');
        badge.className = 'docx-page-badge no-print';
        badge.textContent = `第 ${vp.pageNum} / ${vp.total} 页`;
        vp.sec.appendChild(badge);
      });
    }

    updateActivePage();
  }

  function updateActivePage() {
    if (virtualPages.length === 0) return;
    const scrollTop = scrollContainer.scrollTop || 0;
    const clientHeight = scrollContainer.clientHeight || 800;
    const probeY = (scrollTop + clientHeight * 0.3) / currentZoom;

    let active = virtualPages[0];
    for (const vp of virtualPages) {
      if (vp.topPx <= probeY) {
        active = vp;
      } else {
        break;
      }
    }
    currentPageNum = active.pageNum;
    if (onPageChange) {
      onPageChange({
        page: active.pageNum,
        loadedPages: virtualPages.length,
        total: active.total,
        formulaCount: container.querySelectorAll('.docx-vml-img, math, mjx-container').length
      });
    }
  }

  function paginate() {
    const wrapper = container.querySelector('.docx-wrapper');
    if (!wrapper) {
      refreshMetrics();
      return;
    }
    const sections = Array.from(wrapper.querySelectorAll('section.docx'));

    for (const sec of sections) {
      const pageHeightPx = parsePtToPx(sec.style.minHeight, 1122.5);
      if (sec.offsetHeight <= pageHeightPx + 48) continue;

      const articles = Array.from(sec.querySelectorAll(':scope > article'));
      if (articles.length === 0) continue;
      const primaryArticle = articles[0];
      for (let aIdx = 1; aIdx < articles.length; aIdx++) {
        while (articles[aIdx].firstChild) {
          primaryArticle.appendChild(articles[aIdx].firstChild);
        }
        articles[aIdx].remove();
      }

      const children = Array.from(primaryArticle.children);
      if (children.length <= 1) continue;

      const headerEl = sec.querySelector(':scope > header');
      const footerEl = sec.querySelector(':scope > footer');
      const padTop = parsePtToPx(sec.style.paddingTop, 72);
      const padBottom = parsePtToPx(sec.style.paddingBottom, 72);
      const headerH = headerEl ? headerEl.offsetHeight : 0;
      const footerH = footerEl ? footerEl.offsetHeight : 0;
      const maxContentH = Math.max(240, pageHeightPx - padTop - padBottom - headerH - footerH);

      const buckets = [];
      let currentBucket = [];
      let accH = 0;

      for (const el of children) {
        const h =
          el.offsetHeight +
          (parseFloat(el.style.marginTop) || 0) +
          (parseFloat(el.style.marginBottom) || 0);

        if (currentBucket.length > 0 && accH + h > maxContentH) {
          buckets.push(currentBucket);
          currentBucket = [el];
          accH = h;
        } else {
          currentBucket.push(el);
          accH += h;
        }
      }
      if (currentBucket.length > 0) {
        buckets.push(currentBucket);
      }

      const validBuckets = buckets.filter(
        (b, idx) => idx === 0 || b.some((el) => isVisibleBlockElement(el))
      );

      if (validBuckets.length <= 1) continue;

      for (let i = 1; i < validBuckets.length; i++) {
        for (const item of validBuckets[i]) {
          item.remove();
        }
      }

      let prevSec = sec;
      const limit = options.progressive === false
        ? validBuckets.length
        : Math.min(immediatePages, validBuckets.length);

      for (let i = 1; i < limit; i++) {
        const newSec = createPageSectionFromBucket(
          sec,
          primaryArticle,
          headerEl,
          footerEl,
          validBuckets[i]
        );
        prevSec.after(newSec);
        prevSec = newSec;
      }

      for (let i = limit; i < validBuckets.length; i++) {
        const bucket = validBuckets[i];
        pendingPageTasks.push(() => {
          const newSec = createPageSectionFromBucket(
            sec,
            primaryArticle,
            headerEl,
            footerEl,
            bucket
          );
          prevSec.after(newSec);
          prevSec = newSec;
        });
      }
    }

    refreshMetrics();

    if (pendingPageTasks.length > 0 && typeof requestAnimationFrame === 'function') {
      const pumpPages = () => {
        if (destroyed || pendingPageTasks.length === 0) return;
        const count = Math.min(2, pendingPageTasks.length);
        for (let k = 0; k < count; k++) {
          const fn = pendingPageTasks.shift();
          fn();
        }
        refreshMetrics();
        if (pendingPageTasks.length > 0) {
          requestAnimationFrame(pumpPages);
        }
      };
      requestAnimationFrame(pumpPages);
    }
  }

  function flushPages() {
    while (pendingPageTasks.length > 0) {
      const task = pendingPageTasks.shift();
      task();
    }
    refreshMetrics();
  }

  function scrollToPage(targetPageNum) {
    if (virtualPages.length === 0) return;
    const clamped = Math.max(1, Math.min(virtualPages.length, targetPageNum));
    const target = virtualPages[clamped - 1];
    if (!target || !scrollContainer.scrollTo) return;
    scrollContainer.scrollTo({
      top: Math.max(0, target.topPx * currentZoom - 12),
      behavior: 'smooth'
    });
  }

  function setZoom(zoom) {
    currentZoom = Math.min(2.0, Math.max(0.5, Number(Number(zoom).toFixed(2))));
    container.style.zoom = String(currentZoom);
    updateActivePage();
    return currentZoom;
  }

  function print() {
    flushPages();
    if (typeof window !== 'undefined' && typeof window.print === 'function') {
      window.print();
    }
  }

  if (scrollContainer && scrollContainer.addEventListener) {
    scrollContainer.addEventListener('scroll', updateActivePage, { passive: true });
  }

  paginate();

  return {
    get pageCount() {
      return virtualPages.length + pendingPageTasks.length;
    },
    get currentPage() {
      return currentPageNum;
    },
    get formulaCount() {
      return container.querySelectorAll('.docx-vml-img, math, mjx-container').length;
    },
    scrollToPage,
    nextPage: () => scrollToPage(currentPageNum + 1),
    prevPage: () => scrollToPage(currentPageNum - 1),
    setZoom,
    flushPages,
    print,
    destroy: () => {
      destroyed = true;
      pendingPageTasks = [];
      if (scrollContainer && scrollContainer.removeEventListener) {
        scrollContainer.removeEventListener('scroll', updateActivePage);
      }
    }
  };
}

module.exports = {
  injectPaperStyles,
  createPaperPaginator
};
