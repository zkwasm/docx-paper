export interface DocxPaperPageEvent {
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
