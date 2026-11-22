/**
 * Pure-JS Markdown (.md) + LaTeX Math ($...$ / $$...$$) to A4 Paper DOM Compiler
 * Renders GitHub-Flavored Markdown (GFM tables, code blocks, task lists, blockquotes,
 * headings, and LaTeX equations compiled to native W3C MathML) into standard
 * section.docx paper cards for seamless pagination and printing.
 */

const BS = String.fromCharCode(92);
const NL = String.fromCharCode(10);
const CR = String.fromCharCode(13);
const NUL = String.fromCharCode(0);

const LATEX_SYMBOLS = {
  alpha: '\u03B1', beta: '\u03B2', gamma: '\u03B3', delta: '\u03B4',
  epsilon: '\u03B5', varepsilon: '\u03B5', zeta: '\u03B6', eta: '\u03B7',
  theta: '\u03B8', vartheta: '\u03D1', iota: '\u03B9', kappa: '\u03BA',
  lambda: '\u03BB', mu: '\u03BC', nu: '\u03BD', xi: '\u03BE',
  pi: '\u03C0', varpi: '\u03D6', rho: '\u03C1', varrho: '\u03F1',
  sigma: '\u03C3', varsigma: '\u03C2', tau: '\u03C4', upsilon: '\u03C5',
  phi: '\u03D5', varphi: '\u03C6', chi: '\u03C7', psi: '\u03C8', omega: '\u03C9',
  Gamma: '\u0393', Delta: '\u0394', Theta: '\u0398', Lambda: '\u039B',
  Xi: '\u039E', Pi: '\u03A0', Sigma: '\u03A3', Upsilon: '\u03A5',
  Phi: '\u03A6', Psi: '\u03A8', Omega: '\u03A9',
  infty: '\u221E', partial: '\u2202', nabla: '\u2207', ell: '\u2113', hbar: '\u210F',
  pm: '\u00B1', mp: '\u2213', times: '\u00D7', div: '\u00F7', cdot: '\u22C5',
  ast: '\u2217', star: '\u22C6', circ: '\u2218', bullet: '\u2022',
  le: '\u2264', leq: '\u2264', ge: '\u2265', geq: '\u2265',
  ne: '\u2260', neq: '\u2260', approx: '\u2248', equiv: '\u2261',
  sim: '\u223C', simeq: '\u2243', cong: '\u2245', propto: '\u221D',
  in: '\u2208', notin: '\u2209', ni: '\u220B',
  subset: '\u2282', subseteq: '\u2286', supset: '\u2283', supseteq: '\u2287',
  cup: '\u222A', cap: '\u2229', emptyset: '\u2205', varnothing: '\u2205',
  forall: '\u2200', exists: '\u2203', nexists: '\u2204', neg: '\u00AC', lnot: '\u00AC',
  wedge: '\u2227', land: '\u2227', vee: '\u2228', lor: '\u2228',
  oplus: '\u2295', otimes: '\u2297', perp: '\u22A5', parallel: '\u2225',
  angle: '\u2220', triangle: '\u25B3', degree: '\u00B0', prime: '\u2032',
  to: '\u2192', rightarrow: '\u2192', leftarrow: '\u2190', leftrightarrow: '\u2194',
  Rightarrow: '\u21D2', Leftarrow: '\u21D0', Leftrightarrow: '\u21D4',
  mapsto: '\u21A6', implies: '\u21D2', iff: '\u21D4',
  uparrow: '\u2191', downarrow: '\u2193',
  dots: '\u2026', ldots: '\u2026', cdots: '\u22EF', vdots: '\u22EE', ddots: '\u22F1',
  therefore: '\u2234', because: '\u2235',
  sum: '\u2211', prod: '\u220F', coprod: '\u2210',
  int: '\u222B', iint: '\u222C', iiint: '\u222D', oint: '\u222E',
  bigcup: '\u22C3', bigcap: '\u22C2', bigoplus: '\u2A01', bigotimes: '\u2A02',
  langle: '\u27E8', rangle: '\u27E9', lceil: '\u2308', rceil: '\u2309',
  lfloor: '\u230A', rfloor: '\u230B', vert: '|', Vert: '\u2016'
};

const LARGE_OP_COMMANDS = new Set([
  'sum', 'prod', 'coprod', 'int', 'iint', 'iiint', 'oint',
  'bigcup', 'bigcap', 'bigoplus', 'bigotimes', 'lim', 'sup', 'inf', 'max', 'min'
]);

const FUNCTION_COMMANDS = new Set([
  'sin', 'cos', 'tan', 'cot', 'sec', 'csc',
  'arcsin', 'arccos', 'arctan', 'sinh', 'cosh', 'tanh',
  'ln', 'log', 'exp', 'lim', 'sup', 'inf', 'max', 'min',
  'det', 'dim', 'ker', 'gcd', 'lcm', 'mod', 'deg', 'arg', 'Pr', 'hom'
]);

const ACCENT_COMMANDS = {
  hat: '^', widehat: '^', vec: '\u2192', overrightarrow: '\u2192',
  bar: '\u00AF', overline: '\u00AF', tilde: '~', widetilde: '~',
  dot: '\u02D9', ddot: '\u00A8'
};

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function readBracedGroup(str, pos) {
  while (pos < str.length && /\s/.test(str[pos])) pos++;
  if (str[pos] !== '{') {
    if (str[pos] === BS) {
      let end = pos + 1;
      while (end < str.length && /[A-Za-z]/.test(str[end])) end++;
      if (end === pos + 1 && end < str.length) end++;
      return { content: str.slice(pos, end), nextPos: end };
    }
    return { content: str[pos] || '', nextPos: pos + 1 };
  }
  let depth = 0;
  let i = pos;
  for (; i < str.length; i++) {
    if (str[i] === BS) {
      i++;
      continue;
    }
    if (str[i] === '{') depth++;
    else if (str[i] === '}') {
      depth--;
      if (depth === 0) {
        return { content: str.slice(pos + 1, i), nextPos: i + 1 };
      }
    }
  }
  return { content: str.slice(pos + 1), nextPos: str.length };
}

function parseLatexNodes(tex) {
  const nodes = [];
  let i = 0;

  while (i < tex.length) {
    const ch = tex[i];
    if (/\s/.test(ch)) {
      i++;
      continue;
    }

    if (ch === '{') {
      const grp = readBracedGroup(tex, i);
      nodes.push({ tag: 'mrow', html: `<mrow>${parseLatexToMmlInner(grp.content)}</mrow>` });
      i = grp.nextPos;
      continue;
    }

    if ((ch === '_' || ch === '^') && nodes.length > 0) {
      const prev = nodes.pop();
      let subContent = null;
      let supContent = null;

      while (i < tex.length && (tex[i] === '_' || tex[i] === '^')) {
        const kind = tex[i];
        i++;
        const arg = readBracedGroup(tex, i);
        i = arg.nextPos;
        if (kind === '_') subContent = parseLatexToMmlInner(arg.content);
        else supContent = parseLatexToMmlInner(arg.content);
        while (i < tex.length && /\s/.test(tex[i])) i++;
      }

      const useUnderOver = prev.isLargeOp;
      if (subContent !== null && supContent !== null) {
        const tag = useUnderOver ? 'munderover' : 'msubsup';
        nodes.push({
          tag,
          html: `<${tag}>${prev.html}<mrow>${subContent}</mrow><mrow>${supContent}</mrow></${tag}>`
        });
      } else if (subContent !== null) {
        const tag = useUnderOver ? 'munder' : 'msub';
        nodes.push({
          tag,
          html: `<${tag}>${prev.html}<mrow>${subContent}</mrow></${tag}>`
        });
      } else if (supContent !== null) {
        const tag = useUnderOver ? 'mover' : 'msup';
        nodes.push({
          tag,
          html: `<${tag}>${prev.html}<mrow>${supContent}</mrow></${tag}>`
        });
      }
      continue;
    }

    if (ch === BS) {
      let end = i + 1;
      if (end < tex.length && /[A-Za-z]/.test(tex[end])) {
        while (end < tex.length && /[A-Za-z]/.test(tex[end])) end++;
      } else if (end < tex.length) {
        end++;
      }
      const cmd = tex.slice(i + 1, end);
      i = end;

      if (cmd === 'left' || cmd === 'right') {
        while (i < tex.length && /\s/.test(tex[i])) i++;
        if (tex.startsWith(BS + '{', i) || tex.startsWith(BS + '}', i)) {
          const delim = tex[i + 1];
          nodes.push({ tag: 'mo', html: `<mo stretchy="true">${escapeHtml(delim)}</mo>` });
          i += 2;
        } else if (tex[i] === BS) {
          const dGrp = readBracedGroup(tex, i);
          const dCmd = dGrp.content.startsWith(BS) ? dGrp.content.slice(1) : dGrp.content;
          const sym = LATEX_SYMBOLS[dCmd] || dCmd;
          nodes.push({ tag: 'mo', html: `<mo stretchy="true">${escapeHtml(sym)}</mo>` });
          i = dGrp.nextPos;
        } else if (i < tex.length) {
          const d = tex[i++];
          if (d !== '.') {
            nodes.push({ tag: 'mo', html: `<mo stretchy="true">${escapeHtml(d)}</mo>` });
          }
        }
        continue;
      }

      if (cmd === 'frac' || cmd === 'dfrac' || cmd === 'tfrac') {
        const num = readBracedGroup(tex, i);
        const den = readBracedGroup(tex, num.nextPos);
        i = den.nextPos;
        nodes.push({
          tag: 'mfrac',
          html: `<mfrac><mrow>${parseLatexToMmlInner(num.content)}</mrow><mrow>${parseLatexToMmlInner(den.content)}</mrow></mfrac>`
        });
        continue;
      }

      if (cmd === 'binom') {
        const top = readBracedGroup(tex, i);
        const bot = readBracedGroup(tex, top.nextPos);
        i = bot.nextPos;
        nodes.push({
          tag: 'mrow',
          html: `<mrow><mo>(</mo><mfrac linethickness="0"><mrow>${parseLatexToMmlInner(top.content)}</mrow><mrow>${parseLatexToMmlInner(bot.content)}</mrow></mfrac><mo>)</mo></mrow>`
        });
        continue;
      }

      if (cmd === 'sqrt') {
        while (i < tex.length && /\s/.test(tex[i])) i++;
        let degContent = null;
        if (tex[i] === '[') {
          const closeIdx = tex.indexOf(']', i + 1);
          if (closeIdx > i) {
            degContent = tex.slice(i + 1, closeIdx);
            i = closeIdx + 1;
          }
        }
        const base = readBracedGroup(tex, i);
        i = base.nextPos;
        if (degContent) {
          nodes.push({
            tag: 'mroot',
            html: `<mroot><mrow>${parseLatexToMmlInner(base.content)}</mrow><mrow>${parseLatexToMmlInner(degContent)}</mrow></mroot>`
          });
        } else {
          nodes.push({
            tag: 'msqrt',
            html: `<msqrt><mrow>${parseLatexToMmlInner(base.content)}</mrow></msqrt>`
          });
        }
        continue;
      }

      if (ACCENT_COMMANDS[cmd]) {
        const base = readBracedGroup(tex, i);
        i = base.nextPos;
        const acc = ACCENT_COMMANDS[cmd];
        nodes.push({
          tag: 'mover',
          html: `<mover accent="true"><mrow>${parseLatexToMmlInner(base.content)}</mrow><mo accent="true">${escapeHtml(acc)}</mo></mover>`
        });
        continue;
      }

      if (cmd === 'text' || cmd === 'textrm' || cmd === 'mtext') {
        const grp = readBracedGroup(tex, i);
        i = grp.nextPos;
        nodes.push({ tag: 'mtext', html: `<mtext>${escapeHtml(grp.content)}</mtext>` });
        continue;
      }

      if (cmd === 'mathrm' || cmd === 'mathbf' || cmd === 'mathbb' || cmd === 'mathcal') {
        const grp = readBracedGroup(tex, i);
        i = grp.nextPos;
        const variant = cmd === 'mathbf' ? 'bold' : 'normal';
        nodes.push({
          tag: 'mrow',
          html: `<mrow mathvariant="${variant}">${parseLatexToMmlInner(grp.content)}</mrow>`
        });
        continue;
      }

      if (cmd === 'begin') {
        const envGrp = readBracedGroup(tex, i);
        const envName = envGrp.content.trim();
        const endToken = BS + `end{${envName}}`;
        const endIdx = tex.indexOf(endToken, envGrp.nextPos);
        if (endIdx !== -1) {
          const body = tex.slice(envGrp.nextPos, endIdx);
          i = endIdx + endToken.length;
          const rows = body.split(BS + BS).map((r) => r.trim()).filter(Boolean);
          const mtrHtml = rows
            .map((rowStr) => {
              const cells = rowStr.split('&');
              return `<mtr>${cells.map((c) => `<mtd>${parseLatexToMmlInner(c)}</mtd>`).join('')}</mtr>`;
            })
            .join('');
          const delimMap = {
            pmatrix: ['(', ')'],
            bmatrix: ['[', ']'],
            Bmatrix: ['{', '}'],
            vmatrix: ['|', '|'],
            Vmatrix: ['\u2016', '\u2016'],
            cases: ['{', '']
          };
          const [openD, closeD] = delimMap[envName] || ['', ''];
          const openHtml = openD ? `<mo stretchy="true">${escapeHtml(openD)}</mo>` : '';
          const closeHtml = closeD ? `<mo stretchy="true">${escapeHtml(closeD)}</mo>` : '';
          nodes.push({
            tag: 'mrow',
            html: `<mrow>${openHtml}<mtable>${mtrHtml}</mtable>${closeHtml}</mrow>`
          });
          continue;
        }
      }

      if (FUNCTION_COMMANDS.has(cmd)) {
        nodes.push({
          tag: 'mi',
          isLargeOp: LARGE_OP_COMMANDS.has(cmd),
          html: `<mi mathvariant="normal">${escapeHtml(cmd)}</mi>`
        });
        continue;
      }

      if (LATEX_SYMBOLS[cmd]) {
        const sym = LATEX_SYMBOLS[cmd];
        const isLarge = LARGE_OP_COMMANDS.has(cmd);
        const isGreekLetter = /^[\u0370-\u03FF]$/.test(sym);
        const tag = isGreekLetter ? 'mi' : 'mo';
        const attr = isLarge ? ' largeop="true"' : '';
        nodes.push({
          tag,
          isLargeOp: isLarge,
          html: `<${tag}${attr}>${escapeHtml(sym)}</${tag}>`
        });
        continue;
      }

      if (cmd === ',' || cmd === ';' || cmd === ':' || cmd === 'quad' || cmd === 'qquad') {
        const w = cmd === 'qquad' ? '2em' : cmd === 'quad' ? '1em' : '0.2em';
        nodes.push({ tag: 'mspace', html: `<mspace width="${w}"/>` });
        continue;
      }

      nodes.push({ tag: 'mo', html: `<mo>${escapeHtml(cmd)}</mo>` });
      continue;
    }

    if (/\d/.test(ch)) {
      let end = i + 1;
      while (end < tex.length && /[\d.]/.test(tex[end])) end++;
      const numStr = tex.slice(i, end);
      nodes.push({ tag: 'mn', html: `<mn>${escapeHtml(numStr)}</mn>` });
      i = end;
      continue;
    }

    if (/[A-Za-z]/.test(ch)) {
      nodes.push({ tag: 'mi', html: `<mi>${escapeHtml(ch)}</mi>` });
      i++;
      continue;
    }

    if (/[\u4E00-\u9FFF]/.test(ch)) {
      let end = i + 1;
      while (end < tex.length && /[\u4E00-\u9FFF]/.test(tex[end])) end++;
      nodes.push({ tag: 'mtext', html: `<mtext>${escapeHtml(tex.slice(i, end))}</mtext>` });
      i = end;
      continue;
    }

    const opChar = ch === '-' ? '\u2212' : ch;
    nodes.push({ tag: 'mo', html: `<mo>${escapeHtml(opChar)}</mo>` });
    i++;
  }

  return nodes;
}

function parseLatexToMmlInner(tex) {
  return parseLatexNodes(tex).map((n) => n.html).join('');
}

function latexToMathML(tex, displayMode = false) {
  const inner = parseLatexToMmlInner(tex.trim());
  const displayAttr = displayMode ? ' display="block"' : '';
  const mathHtml = `<math xmlns="http://www.w3.org/1998/Math/MathML"${displayAttr}>${inner}</math>`;
  return displayMode ? `<div class="docx-math-block">${mathHtml}</div>` : mathHtml;
}

function parseMarkdownInline(text) {
  if (!text) return '';
  const placeholders = [];
  const stash = (html) => {
    const id = `${NUL}MD_TOKEN_${placeholders.length}${NUL}`;
    placeholders.push(html);
    return id;
  };

  // 1. Stash inline code `...`
  let s = text.replace(/`([^`]+)`/g, (_, code) =>
    stash(`<code class="docx-md-inline-code">${escapeHtml(code)}</code>`)
  );

  // 2. Stash display math $$...$$ inside a line
  s = s.replace(/\$\$([\s\S]+?)\$\$/g, (_, m1) => stash(latexToMathML(m1, true)));

  // 3. Stash inline math $...$
  s = s.replace(/\$([^\$]+?)\$/g, (_, m1) => stash(latexToMathML(m1, false)));

  // 4. Stash images ![alt](url)
  s = s.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, (_, alt, url) =>
    stash(`<img class="docx-md-img" src="${escapeHtml(url.trim())}" alt="${escapeHtml(alt)}" />`)
  );

  // 5. Stash links [text](url)
  s = s.replace(/\[([^\]]+)\]\(([^)]+)\)/g, (_, label, url) =>
    stash(
      `<a class="docx-md-link" href="${escapeHtml(url.trim())}" target="_blank" rel="noopener noreferrer">${parseMarkdownInline(label)}</a>`
    )
  );

  // Escape remaining raw HTML characters
  s = escapeHtml(s);

  // 6. Bold + Italic, Bold, Italic, Strikethrough
  s = s.replace(/\*\*\*([\s\S]+?)\*\*\*/g, '<strong><em>$1</em></strong>');
  s = s.replace(/\*\*([\s\S]+?)\*\*/g, '<strong>$1</strong>');
  s = s.replace(/__([\s\S]+?)__/g, '<strong>$1</strong>');
  s = s.replace(/\*([^\*]+?)\*/g, '<em>$1</em>');
  s = s.replace(/~~([\s\S]+?)~~/g, '<del>$1</del>');

  // Restore stashed tokens
  const restoreRe = new RegExp(NUL + 'MD_TOKEN_([0-9]+)' + NUL, 'g');
  s = s.replace(restoreRe, (_, idx) => placeholders[Number(idx)] || '');
  return s;
}

function compileMarkdownToHtmlBlocks(markdownText) {
  let src = String(markdownText || '').split(CR + NL).join(NL);

  // Optional YAML frontmatter at top of file
  let frontmatterTitle = '';
  const fmPrefix = '---' + NL;
  const fmDelim = NL + '---' + NL;
  if (src.startsWith(fmPrefix)) {
    const endFm = src.indexOf(fmDelim, 4);
    if (endFm !== -1) {
      const fmBlock = src.slice(4, endFm);
      const titleMatch = fmBlock.match(/^title\s*:\s*(.+)$/m);
      if (titleMatch) {
        frontmatterTitle = titleMatch[1].trim().replace(/^['"]|['"]$/g, '');
      }
      src = src.slice(endFm + 5);
    }
  }

  const lines = src.split(NL);
  const pages = [[]];
  const curPage = () => pages[pages.length - 1];

  if (frontmatterTitle) {
    curPage().push(`<h1 class="docx-md-h1 docx-md-doc-title">${parseMarkdownInline(frontmatterTitle)}</h1>`);
  }

  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    const trimmed = line.trim();

    if (!trimmed) {
      i++;
      continue;
    }

    // Explicit page break directives (pagebreak, newpage, <!-- pagebreak -->)
    if (
      trimmed === BS + 'pagebreak' ||
      trimmed === BS + 'newpage' ||
      /^<!--\s*pagebreak\s*-->$/i.test(trimmed)
    ) {
      if (curPage().length > 0) {
        pages.push([]);
      }
      i++;
      continue;
    }

    // Multi-line display math block $$ ... $$ or \[ ... \]
    if (trimmed.startsWith('$$') || trimmed.startsWith(BS + '[')) {
      const closeToken = trimmed.startsWith('$$') ? '$$' : BS + ']';
      const openLen = 2;
      if (trimmed.length > openLen && trimmed.endsWith(closeToken) && trimmed.length >= 4) {
        const tex = trimmed.slice(openLen, -closeToken.length);
        curPage().push(latexToMathML(tex, true));
        i++;
        continue;
      }
      const mathLines = [trimmed.slice(openLen)];
      i++;
      while (i < lines.length) {
        const l = lines[i];
        const idx = l.indexOf(closeToken);
        if (idx !== -1) {
          mathLines.push(l.slice(0, idx));
          i++;
          break;
        }
        mathLines.push(l);
        i++;
      }
      curPage().push(latexToMathML(mathLines.join(NL), true));
      continue;
    }

    // Fenced code block ```lang ... ```
    const fenceMatch = trimmed.match(/^(`{3,}|~{3,})(.*)$/);
    if (fenceMatch) {
      const fence = fenceMatch[1];
      const lang = (fenceMatch[2] || '').trim();
      const codeLines = [];
      i++;
      while (i < lines.length && !lines[i].trim().startsWith(fence)) {
        codeLines.push(lines[i]);
        i++;
      }
      if (i < lines.length) i++;
      const langBadge = lang
        ? `<div class="docx-md-code-lang">${escapeHtml(lang)}</div>`
        : '';
      curPage().push(
        `<div class="docx-md-codeblock">${langBadge}<pre><code>${escapeHtml(codeLines.join(NL))}</code></pre></div>`
      );
      continue;
    }

    // ATX Headings (# .. ######)
    const hMatch = line.match(/^\s*(#{1,6})\s+(.+?)\s*#*$/);
    if (hMatch) {
      const level = hMatch[1].length;
      curPage().push(
        `<h${level} class="docx-md-h${level}">${parseMarkdownInline(hMatch[2])}</h${level}>`
      );
      i++;
      continue;
    }

    // Horizontal rule (---, ***, ___)
    if (/^(\*\s*\*\s*\*+|-\s*-\s*-+|_\s*_\s*_+)\s*$/.test(trimmed)) {
      curPage().push('<hr class="docx-md-hr" />');
      i++;
      continue;
    }

    // GFM Table
    if (
      trimmed.includes('|') &&
      i + 1 < lines.length &&
      /^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)+\|?\s*$/.test(lines[i + 1])
    ) {
      const parseRow = (r) =>
        r
          .trim()
          .replace(/^\|/, '')
          .replace(/\|$/, '')
          .split('|')
          .map((c) => c.trim());

      const headers = parseRow(lines[i]);
      const aligns = parseRow(lines[i + 1]).map((a) => {
        if (a.startsWith(':') && a.endsWith(':')) return 'center';
        if (a.endsWith(':')) return 'right';
        return 'left';
      });
      i += 2;
      const bodyRows = [];
      while (i < lines.length && lines[i].trim().includes('|') && lines[i].trim() !== '') {
        bodyRows.push(parseRow(lines[i]));
        i++;
      }
      const thHtml = headers
        .map(
          (h, idx) =>
            `<th style="text-align:${aligns[idx] || 'left'}">${parseMarkdownInline(h)}</th>`
        )
        .join('');
      const trHtml = bodyRows
        .map(
          (row) =>
            `<tr>${headers
              .map(
                (_, idx) =>
                  `<td style="text-align:${aligns[idx] || 'left'}">${parseMarkdownInline(row[idx] ?? '')}</td>`
              )
              .join('')}</tr>`
        )
        .join('');
      curPage().push(
        `<table class="docx-md-table"><thead><tr>${thHtml}</tr></thead><tbody>${trHtml}</tbody></table>`
      );
      continue;
    }

    // Blockquote (> ...)
    if (trimmed.startsWith('>')) {
      const qLines = [];
      while (i < lines.length && lines[i].trim().startsWith('>')) {
        qLines.push(lines[i].trim().replace(/^>\s?/, ''));
        i++;
      }
      curPage().push(
        `<blockquote class="docx-md-quote"><p>${parseMarkdownInline(qLines.join('<br/>'))}</p></blockquote>`
      );
      continue;
    }

    // Unordered or Ordered Lists (-, *, +, 1.)
    const ulMatch = line.match(/^(\s*)([-*+])\s+(.*)$/);
    const olMatch = line.match(/^(\s*)(\d+)\.\s+(.*)$/);
    if (ulMatch || olMatch) {
      const isOrdered = Boolean(olMatch);
      const tag = isOrdered ? 'ol' : 'ul';
      const startNum = isOrdered ? Number(olMatch[2]) || 1 : 1;
      const startAttr = isOrdered && startNum !== 1 ? ` start="${startNum}"` : '';
      const itemsHtml = [];

      while (i < lines.length) {
        const l = lines[i];
        const m = isOrdered
          ? l.match(/^(\s*)(\d+)\.\s+(.*)$/)
          : l.match(/^(\s*)([-*+])\s+(.*)$/);
        if (!m) break;
        let itemText = m[3];
        const taskMatch = itemText.match(/^\[([ xX])\]\s+(.*)$/);
        if (taskMatch) {
          const checked = taskMatch[1].toLowerCase() === 'x';
          const box = checked ? '\u2611' : '\u2610';
          itemText = `<span class="docx-md-task-box">${box}</span> ${parseMarkdownInline(taskMatch[2])}`;
        } else {
          itemText = parseMarkdownInline(itemText);
        }
        itemsHtml.push(`<li>${itemText}</li>`);
        i++;
      }
      curPage().push(`<${tag} class="docx-md-list"${startAttr}>${itemsHtml.join('')}</${tag}>`);
      continue;
    }

    // Regular Paragraph
    const pLines = [];
    while (
      i < lines.length &&
      lines[i].trim() !== '' &&
      !lines[i].trim().startsWith('#') &&
      !lines[i].trim().startsWith('```') &&
      !lines[i].trim().startsWith('~~~') &&
      !lines[i].trim().startsWith('$$') &&
      !lines[i].trim().startsWith(BS + '[') &&
      !lines[i].trim().startsWith('>') &&
      !/^(\s*)([-*+]|\d+\.)\s+/.test(lines[i]) &&
      !/^(\*\s*\*\s*\*+|-\s*-\s*-+|_\s*_\s*_+)\s*$/.test(lines[i].trim())
    ) {
      pLines.push(lines[i].trim());
      i++;
    }
    if (pLines.length > 0) {
      curPage().push(`<p class="docx-md-p">${parseMarkdownInline(pLines.join(' '))}</p>`);
    } else {
      i++;
    }
  }

  return pages.filter((p) => p.length > 0);
}

const MD_PAPER_STYLES = `
.docx-md-section {
  font-family: "Times New Roman", "Songti SC", "STSong", "PingFang SC", -apple-system, serif;
  color: #0f172a;
  font-size: 11pt;
  line-height: 1.65;
}
.docx-md-doc-title {
  text-align: center;
  margin-bottom: 18pt !important;
}
.docx-md-h1 {
  font-family: "PingFang SC", "Heiti SC", -apple-system, sans-serif;
  font-size: 20pt;
  font-weight: 700;
  color: #0f172a;
  margin: 14pt 0 8pt;
  padding-bottom: 5pt;
  border-bottom: 1px solid #e2e8f0;
  line-height: 1.3;
}
.docx-md-h2 {
  font-family: "PingFang SC", "Heiti SC", -apple-system, sans-serif;
  font-size: 15.5pt;
  font-weight: 700;
  color: #0f172a;
  margin: 13pt 0 6pt;
  line-height: 1.35;
}
.docx-md-h3 {
  font-family: "PingFang SC", "Heiti SC", -apple-system, sans-serif;
  font-size: 13pt;
  font-weight: 600;
  color: #1e293b;
  margin: 10pt 0 5pt;
  line-height: 1.4;
}
.docx-md-h4, .docx-md-h5, .docx-md-h6 {
  font-family: "PingFang SC", "Heiti SC", -apple-system, sans-serif;
  font-size: 11.5pt;
  font-weight: 600;
  color: #334155;
  margin: 8pt 0 4pt;
}
.docx-md-p {
  margin: 0 0 8pt !important;
  line-height: 1.68 !important;
  text-align: justify;
}
.docx-md-list {
  margin: 0 0 9pt 20pt;
  padding: 0;
}
.docx-md-list li {
  margin-bottom: 4pt;
  line-height: 1.6;
}
.docx-md-task-box {
  font-size: 12pt;
  color: #2563eb;
  margin-right: 3pt;
}
.docx-md-quote {
  margin: 8pt 0 10pt;
  padding: 6pt 12pt;
  border-left: 3pt solid #3b82f6;
  background: #f8fafc;
  color: #334155;
  border-radius: 0 4px 4px 0;
}
.docx-md-quote p {
  margin: 0 !important;
}
.docx-md-inline-code {
  font-family: "SFMono-Regular", Menlo, Monaco, Consolas, monospace;
  font-size: 0.9em;
  background: #f1f5f9;
  color: #be123c;
  padding: 1.5pt 4pt;
  border-radius: 3px;
}
.docx-md-codeblock {
  margin: 8pt 0 10pt;
  background: #f8fafc;
  border: 1px solid #e2e8f0;
  border-radius: 6px;
  overflow: hidden;
  break-inside: avoid;
  page-break-inside: avoid;
}
.docx-md-code-lang {
  background: #f1f5f9;
  color: #64748b;
  font-family: -apple-system, sans-serif;
  font-size: 8.5pt;
  font-weight: 600;
  padding: 3pt 10pt;
  border-bottom: 1px solid #e2e8f0;
  text-transform: uppercase;
}
.docx-md-codeblock pre {
  margin: 0;
  padding: 9pt 11pt;
  overflow-x: auto;
  font-family: "SFMono-Regular", Menlo, Monaco, Consolas, monospace;
  font-size: 9.5pt;
  line-height: 1.5;
  color: #1e293b;
  white-space: pre-wrap;
  word-break: break-word;
}
.docx-md-table {
  width: 100%;
  border-collapse: collapse;
  margin: 10pt 0 12pt;
  font-size: 10.5pt;
  border-top: 1.5pt solid #0f172a;
  border-bottom: 1.5pt solid #0f172a;
}
.docx-md-table th {
  font-weight: 600;
  padding: 6pt 8pt;
  border-bottom: 1pt solid #334155;
  background: #f8fafc;
}
.docx-md-table td {
  padding: 5.5pt 8pt;
  border-bottom: 0.5pt solid #e2e8f0;
}
.docx-md-hr {
  border: none;
  border-top: 1px solid #cbd5e1;
  margin: 14pt 0;
}
.docx-md-img {
  max-width: 100%;
  height: auto;
  display: block;
  margin: 8pt auto;
}
.docx-md-link {
  color: #2563eb !important;
  text-decoration: underline !important;
}
`;

async function decodeMarkdownInput(data) {
  if (typeof data === 'string') return data;
  if (typeof Blob !== 'undefined' && data instanceof Blob) {
    return await data.text();
  }
  if (data instanceof ArrayBuffer || ArrayBuffer.isView(data)) {
    return new TextDecoder('utf-8').decode(data);
  }
  return String(data || '');
}

async function renderMarkdownAsync(data, bodyContainer, styleContainer = null, options = {}) {
  const mdText = await decodeMarkdownInput(data);
  const className = options.className || 'docx';
  const targetStyleContainer = styleContainer || bodyContainer;
  bodyContainer.innerHTML = '';

  const styleEl = document.createElement('style');
  styleEl.textContent = `
.${className}-wrapper { display: flex; flex-flow: column; align-items: center; }
section.${className} { box-sizing: border-box; display: flex; flex-flow: column nowrap; position: relative; overflow: hidden; background: white; }
section.${className}>article { margin-bottom: auto; z-index: 1; }
${MD_PAPER_STYLES}
`;
  targetStyleContainer.appendChild(styleEl);

  const wrapper = document.createElement('div');
  wrapper.className = `${className}-wrapper`;

  const pages = compileMarkdownToHtmlBlocks(mdText);
  const nonEmptyPages = pages.length > 0 ? pages : [['<p class="docx-md-p"></p>']];

  for (const pageBlocks of nonEmptyPages) {
    const sec = document.createElement('section');
    sec.className = `${className} docx-md-section`;
    sec.style.width = '595.3pt';
    sec.style.minHeight = '841.9pt';
    sec.style.paddingTop = '64pt';
    sec.style.paddingBottom = '64pt';
    sec.style.paddingLeft = '68pt';
    sec.style.paddingRight = '68pt';

    const article = document.createElement('article');
    article.innerHTML = pageBlocks.join(NL);
    sec.appendChild(article);
    wrapper.appendChild(sec);
  }

  bodyContainer.appendChild(wrapper);
}

module.exports = {
  latexToMathML,
  parseMarkdownInline,
  compileMarkdownToHtmlBlocks,
  renderMarkdownAsync
};
