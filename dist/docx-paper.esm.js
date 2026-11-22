import './docx-paper.js';
const api = globalThis.DocxPaper || globalThis.docx;
export const renderDocx = api.renderDocx;
export const renderAsync = api.renderAsync;
export const parseAsync = api.parseAsync;
export const createPaperPaginator = api.createPaperPaginator;
export const injectPaperStyles = api.injectPaperStyles;
export default api;
