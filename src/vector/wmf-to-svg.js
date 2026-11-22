/**
 * Pure-JS Windows Metafile (.wmf) to Vector SVG Converter
 * Specifically engineered for MathType / Equation Editor OLE vector formulas,
 * supporting Placeable WMF headers, META_EXTTEXTOUT with Dx kerning arrays,
 * Symbol & MT Extra mathematical glyphs, over-arrow vector synthesis, and radical polygons.
 */

const SYMBOL_CHAR_MAP = {
  0x22: "\u2200", 0x24: "\u2203", 0x27: "\u220B", 0x2a: "\u2217", 0x2d: "\u2212",
  0x40: "\u2245", 0x41: "\u0391", 0x42: "\u0392", 0x43: "\u03A7", 0x44: "\u0394",
  0x45: "\u0395", 0x46: "\u03A6", 0x47: "\u0393", 0x48: "\u0397", 0x49: "\u0399",
  0x4a: "\u03D1", 0x4b: "\u039A", 0x4c: "\u039B", 0x4d: "\u039C", 0x4e: "\u039D",
  0x4f: "\u039F", 0x50: "\u03A0", 0x51: "\u0398", 0x52: "\u03A1", 0x53: "\u03A3",
  0x54: "\u03A4", 0x55: "\u03A5", 0x56: "\u03C2", 0x57: "\u03A9", 0x58: "\u039E",
  0x59: "\u03A8", 0x5a: "\u0396", 0x5c: "\u2234", 0x5e: "\u22A5", 0x61: "\u03B1",
  0x62: "\u03B2", 0x63: "\u03C7", 0x64: "\u03B4", 0x65: "\u03B5", 0x66: "\u03C6",
  0x67: "\u03B3", 0x68: "\u03B7", 0x69: "\u03B9", 0x6a: "\u03D5", 0x6b: "\u03BA",
  0x6c: "\u03BB", 0x6d: "\u03BC", 0x6e: "\u03BD", 0x6f: "\u03BF", 0x70: "\u03C0",
  0x71: "\u03B8", 0x72: "\u03C1", 0x73: "\u03C3", 0x74: "\u03C4", 0x75: "\u03C5",
  0x76: "\u03D6", 0x77: "\u03C9", 0x78: "\u03BE", 0x79: "\u03C8", 0x7a: "\u03B6",
  0x7e: "\u223C", 0xa2: "\u2032", 0xa3: "\u2264", 0xa5: "\u221E", 0xab: "\u2194",
  0xac: "\u2190", 0xad: "\u2191", 0xae: "\u2192", 0xaf: "\u2193", 0xb0: "\u00B0",
  0xb1: "\u00B1", 0xb2: "\u2033", 0xb3: "\u2265", 0xb4: "\u00D7", 0xb5: "\u221D",
  0xb6: "\u2202", 0xb7: "\u2022", 0xb8: "\u00F7", 0xb9: "\u2260", 0xba: "\u2261",
  0xbb: "\u2248", 0xbc: "\u2026", 0xc4: "\u2297", 0xc5: "\u2295", 0xc6: "\u2205",
  0xc7: "\u2229", 0xc8: "\u222A", 0xc9: "\u2283", 0xca: "\u2287", 0xcb: "\u2284",
  0xcc: "\u2282", 0xcd: "\u2286", 0xce: "\u2208", 0xcf: "\u2209", 0xd0: "\u2220",
  0xd1: "\u2207", 0xd5: "\u220F", 0xd6: "\u221A", 0xd7: "\u22C5", 0xd8: "\u00AC",
  0xd9: "\u2227", 0xda: "\u2228", 0xdb: "\u21D4", 0xdc: "\u21D0", 0xdd: "\u21D1",
  0xde: "\u21D2", 0xdf: "\u21D3", 0xe1: "\u27E8", 0xe5: "\u2211", 0xe6: "\u239B",
  0xe7: "\u239C", 0xe8: "\u239D", 0xe9: "\u23A1", 0xea: "\u23A2", 0xeb: "\u23A3",
  0xec: "\u23A7", 0xed: "\u23A8", 0xee: "\u23A9", 0xef: "\u23AA", 0xf1: "\u27E9",
  0xf2: "\u222B", 0xf3: "\u2320", 0xf4: "\u23AE", 0xf5: "\u2321", 0xf6: "\u239E",
  0xf7: "\u239F", 0xf8: "\u23A0", 0xf9: "\u23A4", 0xfa: "\u23A5", 0xfb: "\u23A6",
  0xfc: "\u23AB", 0xfd: "\u23AC", 0xfe: "\u23AD"
};

const MT_EXTRA_CHAR_MAP = {
  0x4c: "\u22EF", 0x4f: "\u22EE", 0x50: "\u22F1", 0x51: "\u22F0",
  0x55: "\u222A", 0x68: "\u210F", 0x6c: "\u2113"
};

function escapeXmlText(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function convertWmfToSvg(u8) {
  try {
    const view = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let pos = 0;
    let winOrgX = 0, winOrgY = 0, winExtX = 1000, winExtY = 1000;
    if (view.getUint32(0, true) === 0x9AC6CDD7) {
      const left = view.getInt16(6, true);
      const top = view.getInt16(8, true);
      const right = view.getInt16(10, true);
      const bottom = view.getInt16(12, true);
      winOrgX = left;
      winOrgY = top;
      winExtX = Math.max(1, right - left);
      winExtY = Math.max(1, bottom - top);
      pos = 22;
    }
    const mtHeaderSize = view.getUint16(pos + 2, true);
    pos += mtHeaderSize * 2;

    const objs = [];
    const addObj = (o) => {
      const idx = objs.indexOf(null);
      if (idx >= 0) objs[idx] = o;
      else objs.push(o);
    };
    let curFont = { name: "Times New Roman", h: -320, italic: 0, weight: 400, charset: 0 };
    let textAlign = 0;
    let cpX = 0, cpY = 0;
    const svgItems = [];
    const gbkDecoder = typeof TextDecoder !== "undefined" ? new TextDecoder("gbk") : null;

    while (pos + 6 <= u8.byteLength) {
      const sizeWords = view.getUint32(pos, true);
      const func = view.getUint16(pos + 4, true);
      if (func === 0 || sizeWords < 3 || pos + sizeWords * 2 > u8.byteLength) break;

      if (func === 0x020b) {
        winOrgY = view.getInt16(pos + 6, true);
        winOrgX = view.getInt16(pos + 8, true);
      } else if (func === 0x020c) {
        winExtY = view.getInt16(pos + 6, true);
        winExtX = view.getInt16(pos + 8, true);
      } else if (func === 0x012e) {
        textAlign = view.getUint16(pos + 6, true);
      } else if (func === 0x02fa) {
        addObj({ type: "pen", style: view.getUint16(pos + 6, true), width: view.getInt16(pos + 8, true) });
      } else if (func === 0x02fc) {
        addObj({ type: "brush", style: view.getUint16(pos + 6, true) });
      } else if (func === 0x02fb) {
        const h = view.getInt16(pos + 6, true);
        const weight = view.getInt16(pos + 14, true);
        const italic = view.getUint8(pos + 16);
        const charset = view.getUint8(pos + 19);
        const faceBytes = u8.subarray(pos + 24, pos + sizeWords * 2);
        const zeroIdx = faceBytes.indexOf(0);
        const cleanFace = zeroIdx >= 0 ? faceBytes.subarray(0, zeroIdx) : faceBytes;
        const name = gbkDecoder ? gbkDecoder.decode(cleanFace) : String.fromCharCode(...cleanFace);
        addObj({ type: "font", name, h, weight, italic, charset });
      } else if (func === 0x012d) {
        const obj = objs[view.getUint16(pos + 6, true)];
        if (obj?.type === "font") curFont = obj;
      } else if (func === 0x01f0) {
        objs[view.getUint16(pos + 6, true)] = null;
      } else if (func === 0x0214) {
        cpY = view.getInt16(pos + 6, true);
        cpX = view.getInt16(pos + 8, true);
      } else if (func === 0x0213) {
        const toY = view.getInt16(pos + 6, true);
        const toX = view.getInt16(pos + 8, true);
        const strokeW = Math.max(16, Math.round(Math.abs(winExtY) * 0.022));
        svgItems.push(`<line x1="${cpX}" y1="${cpY}" x2="${toX}" y2="${toY}" stroke="#000" stroke-width="${strokeW}" stroke-linecap="butt"/>`);
        cpX = toX;
        cpY = toY;
      } else if (func === 0x0324) {
        const n = view.getInt16(pos + 6, true);
        const pts = [];
        for (let i = 0; i < n; i++) {
          pts.push(`${view.getInt16(pos + 8 + i * 4, true)},${view.getInt16(pos + 10 + i * 4, true)}`);
        }
        svgItems.push(`<polygon points="${pts.join(" ")}" fill="#000" stroke="#000" stroke-width="6" stroke-linejoin="miter"/>`);
      } else if (func === 0x0a32) {
        const y = view.getInt16(pos + 6, true);
        const x = view.getInt16(pos + 8, true);
        const len = view.getInt16(pos + 10, true);
        const opts = view.getUint16(pos + 12, true);
        const strOff = pos + 14 + ((opts & 0x0006) ? 8 : 0);
        const bytes = u8.subarray(strOff, strOff + len);
        const dxOff = strOff + ((len + 1) & ~1);
        const recEnd = pos + sizeWords * 2;
        const dx = [];
        for (let p = dxOff; p + 2 <= recEnd; p += 2) {
          dx.push(view.getInt16(p, true));
        }

        let drawX = (textAlign & 0x01) ? cpX : x;
        let drawY = (textAlign & 0x01) ? cpY : y;
        const fontSize = Math.round(Math.abs(curFont.h || 320) * 0.92);
        const isSymbol = /Symbol/i.test(curFont.name);
        const isMtExtra = /MT Extra/i.test(curFont.name);

        if (isMtExtra) {
          let arrowStartX = null;
          for (let i = 0; i < bytes.length; i++) {
            const b = bytes[i];
            const adv = dx[i] ?? Math.round(fontSize * 0.5);
            if (b === 0x75) {
              if (arrowStartX === null) arrowStartX = drawX;
            } else if (b === 0x72 || b === 0x76) {
              const startX = arrowStartX ?? drawX;
              const endX = drawX + Math.round(fontSize * 0.42);
              const arrowY = drawY - Math.round(fontSize * 0.08);
              const strokeW = Math.max(14, Math.round(fontSize * 0.055));
              const headL = Math.round(fontSize * 0.24);
              const headH = Math.round(fontSize * 0.13);
              svgItems.push(`<line x1="${startX}" y1="${arrowY}" x2="${endX}" y2="${arrowY}" stroke="#000" stroke-width="${strokeW}"/>`);
              svgItems.push(`<polygon points="${endX},${arrowY} ${endX - headL},${arrowY - headH} ${endX - headL},${arrowY + headH}" fill="#000"/>`);
              arrowStartX = null;
            } else {
              const ch = MT_EXTRA_CHAR_MAP[b] ?? String.fromCharCode(b);
              svgItems.push(`<text x="${drawX}" y="${drawY}" font-family="'Cambria Math','Latin Modern Math','Times New Roman',serif" font-size="${fontSize}" fill="#000">${escapeXmlText(ch)}</text>`);
            }
            drawX += adv;
          }
        } else if (curFont.charset === 134 && gbkDecoder) {
          const text = gbkDecoder.decode(bytes);
          svgItems.push(`<text x="${drawX}" y="${drawY}" font-family="'Songti SC','STSong','SimSun',serif" font-size="${fontSize}" fill="#000">${escapeXmlText(text)}</text>`);
        } else {
          const fontStyle = curFont.italic ? ' font-style="italic"' : '';
          const fontWeight = curFont.weight >= 700 ? ' font-weight="bold"' : '';
          const family = isSymbol
            ? "'Cambria Math','STIX Two Math','Symbol','Times New Roman',serif"
            : "'Times New Roman','Cambria Math',serif";
          for (let i = 0; i < bytes.length; i++) {
            const b = bytes[i];
            const ch = isSymbol ? (SYMBOL_CHAR_MAP[b] ?? String.fromCharCode(b)) : String.fromCharCode(b);
            svgItems.push(`<text x="${drawX}" y="${drawY}" font-family="${family}" font-size="${fontSize}"${fontStyle}${fontWeight} fill="#000">${escapeXmlText(ch)}</text>`);
            drawX += dx[i] ?? Math.round(fontSize * 0.5);
          }
        }
        if (textAlign & 0x01) {
          cpX = drawX;
        }
      }
      pos += sizeWords * 2;
    }
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${winOrgX} ${winOrgY} ${winExtX} ${winExtY}" preserveAspectRatio="xMidYMid meet">${svgItems.join("")}</svg>`;
  } catch (e) {
    console.error("WMF to SVG conversion error:", e);
    return null;
  }
}

module.exports = {
  SYMBOL_CHAR_MAP,
  MT_EXTRA_CHAR_MAP,
  escapeXmlText,
  convertWmfToSvg
};
