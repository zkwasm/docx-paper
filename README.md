# docx-paper

[![Live Demo](https://img.shields.io/badge/Live%20Demo-GitHub%20Pages-2563eb.svg)](https://zkwasm.github.io/docx-paper/demo/)
[![npm version](https://img.shields.io/npm/v/docx-paper.svg)](https://www.npmjs.com/package/docx-paper)
[![npm downloads](https://img.shields.io/npm/dm/docx-paper.svg)](https://www.npmjs.com/package/docx-paper)
[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](./LICENSE)
[![Compatibility](https://img.shields.io/badge/OpenXML%20Benchmark-100%2F100%20Passed-brightgreen.svg)](#compatibility)

**[English](./README.md) | [简体中文](./README.zh-CN.md) | [✨ Live Demo](https://zkwasm.github.io/docx-paper/demo/)**

**`docx-paper`** is a high-fidelity, pure-frontend `.docx` (Office Open XML) page renderer and print engine designed for the AI-native era—where humans read, review, and print Word documents in the browser while AI handles document generation and editing.

Built on top of the `docxjs` foundation and extensively engineered for real-world complex documents, `docx-paper` solves the hardest compatibility problems that break existing browser `.docx` viewers:

- **Built-in WMF & EMF Vector Equation Engine**: Renders embedded MathType / OLE equations (`<w:object>` + `.wmf` / `.emf` Windows Metafiles) directly into crisp vector `<svg>` and `<img type="image/bmp">` in pure JavaScript—no server, LibreOffice, or Microsoft Office required.
- **Spec-Compliant OMML Math (`<m:oMath>`)**: Tokenizes and compiles Word OMML equations into native W3C MathML (`<mi>`, `<mn>`, `<mo>`, `<mfrac>`, `<msqrt>`, `<mroot>`, `<msubsup>`, `<munderover>`, `<mover accent="true">`, `<mtable>`).
- **Full CJK & Chinese Typography Support**: Supports character-unit indentation (`firstLineChars`, `hangingChars`, `leftChars`), line-unit paragraph spacing (`beforeLines`, `afterLines`), distributed justification (`distribute`), emphasis dots (`<w:em>`), phonetic ruby guides (`<w:ruby>`), circled numbering (`①`, `⑴`, `⒈`, `㊀`), and automatic macOS/Windows CJK font fallback chains.
- **Complex Tables & Three-Line Tables (三线表)**: Properly separates outer table borders from inner cell grid borders (`insideH` / `insideV`) and supports `<w:sdt>` content controls inside tables.
- **Progressive Physical Pagination & Native Print**: Automatically splits continuous sections into discrete physical paper cards (A4 / B5 / Letter) with realistic page shadows and gaps, renders the first 2 pages instantaneously while streaming remaining pages and formula images in the background, and includes `@media print` rules that prevent slicing equations or table rows across page breaks.

---

## Quick Start

### 1. Browser (`<script>` via CDN)

```html
<script src="https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/docx-paper@1.0.0/dist/docx-paper.js"></script>

<div id="scroll-viewport" style="height: 100vh; overflow: auto; background: #e2e8f0; padding: 24px;">
  <div id="viewer"></div>
</div>

<script>
  async function openDocument(file) {
    const buffer = await file.arrayBuffer();
    const controller = await DocxPaper.renderDocx(buffer, document.getElementById('viewer'), {
      scrollContainer: document.getElementById('scroll-viewport'),
      paginate: true,
      progressive: true,
      showPageBadge: true,
      onPageChange: ({ page, total, formulaCount }) => {
        console.log(`Page ${page} / ${total} (${formulaCount} formulas)`);
      }
    });

    // Built-in navigation & print controls:
    // controller.nextPage();
    // controller.prevPage();
    // controller.setZoom(1.2);
    // controller.print();
  }
</script>
```

### 2. NPM / ES Module (React, Vue, Next.js)

```bash
npm install docx-paper jszip
```

```js
import { renderDocx } from 'docx-paper';

const controller = await renderDocx(arrayBuffer, containerElement, {
  paginate: true,
  progressive: true
});
```

---

## API Reference

### `renderDocx(data, container, options?)`

| Option | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `scrollContainer` | `HTMLElement` | `container.parentElement` | Scrollable viewport used to compute the active page number on scroll. |
| `paginate` | `boolean` | `true` | Automatically split overflowing sections into separate physical paper cards with gaps between pages. |
| `progressive` | `boolean` | `true` | Render the first `immediatePages` synchronously and stream remaining pages in `requestAnimationFrame`. |
| `immediatePages` | `number` | `2` | Number of initial pages to mount before yielding to the browser paint loop. |
| `showPageBadge` | `boolean` | `true` | Display `Page X / Y` badge in the top-right corner of each paper card (hidden when printing). |
| `onPageChange` | `Function` | `undefined` | Callback receiving `{ page, loadedPages, total, formulaCount }`. |

### Controller Methods

- `controller.pageCount`: Total number of pages.
- `controller.currentPage`: Current active page number (`1..pageCount`).
- `controller.formulaCount`: Total number of rendered formulas and vector objects.
- `controller.nextPage()` / `controller.prevPage()` / `controller.scrollToPage(n)`: Smoothly scroll to a specific page.
- `controller.setZoom(scale)`: Layout-aware zoom scaling (`0.5` to `2.0`).
- `controller.print()`: Flushes any pending background pages and triggers native browser printing (`window.print()`).

---

## Compatibility

Tested against a 100-document benchmark suite comprising:
- **93 Pandoc Official `.docx` Test Suite documents** (equations, complex tables, lists, headers/footers, VML textboxes, EMF/WMF objects, customXML, track changes, footnotes)
- **Calibre Official `demo.docx` Full-Feature Benchmark**
- **Real-world Chinese STEM Exam Papers & Technical Specifications** (1,000+ embedded MathType `.wmf`/`.png` equations, multi-section portrait/landscape tables, mixed CJK/English layouts)

## License

Apache-2.0
