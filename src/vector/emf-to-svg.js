/**
 * Pure-JS Enhanced Metafile (.emf) to BMP / Vector SVG Converter
 * Supports embedded Device-Independent Bitmaps (EMR_STRETCHDIBITS) and
 * vector drawing records (EMR_POLYGON16, EMR_POLYLINE16, EMR_EXTTEXTOUTW).
 */

const { escapeXmlText } = require('./wmf-to-svg');

function dibToBmpDataUrl(u8, bmiOffset, bmiSize, bitsOffset, bitsSize) {
  const totalSize = 14 + bmiSize + bitsSize;
  const bmp = new Uint8Array(totalSize);
  const dv = new DataView(bmp.buffer);
  bmp[0] = 0x42; // 'B'
  bmp[1] = 0x4d; // 'M'
  dv.setUint32(2, totalSize, true);
  dv.setUint32(6, 0, true);
  dv.setUint32(10, 14 + bmiSize, true);
  bmp.set(u8.subarray(bmiOffset, bmiOffset + bmiSize), 14);
  bmp.set(u8.subarray(bitsOffset, bitsOffset + bitsSize), 14 + bmiSize);
  return bmp;
}

function convertEmfToBlob(u8) {
  try {
    if (u8.byteLength < 88) return null;
    const view = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    if (view.getUint32(0, true) !== 1) return null;
    const left = view.getInt32(8, true);
    const top = view.getInt32(12, true);
    const right = view.getInt32(16, true);
    const bottom = view.getInt32(20, true);
    const width = Math.max(1, right - left);
    const height = Math.max(1, bottom - top);

    let pos = 0;
    const dibBlobs = [];
    const svgItems = [];
    const fonts = {};
    let curFont = { size: 16, family: 'sans-serif' };

    while (pos + 8 <= u8.byteLength) {
      const type = view.getUint32(pos, true);
      const size = view.getUint32(pos + 4, true);
      if (type === 14 || size < 8 || pos + size > u8.byteLength) break;

      if (type === 81 && size >= 80) {
        const offBmi = view.getUint32(pos + 48, true);
        const cbBmi = view.getUint32(pos + 52, true);
        const offBits = view.getUint32(pos + 56, true);
        const cbBits = view.getUint32(pos + 60, true);
        if (cbBmi > 0 && cbBits > 0 && pos + offBits + cbBits <= u8.byteLength) {
          const bmpBytes = dibToBmpDataUrl(u8, pos + offBmi, cbBmi, pos + offBits, cbBits);
          dibBlobs.push(bmpBytes);
        }
      } else if (type === 86 && size >= 28) {
        const count = view.getUint32(pos + 24, true);
        const pts = [];
        for (let i = 0; i < count && 28 + i * 4 + 4 <= size; i++) {
          const px = view.getInt16(pos + 28 + i * 4, true);
          const py = view.getInt16(pos + 30 + i * 4, true);
          pts.push(`${px},${py}`);
        }
        if (pts.length > 0) {
          svgItems.push(`<polygon points="${pts.join(' ')}" fill="none" stroke="#334155" stroke-width="1"/>`);
        }
      } else if (type === 87 && size >= 28) {
        const count = view.getUint32(pos + 24, true);
        const pts = [];
        for (let i = 0; i < count && 28 + i * 4 + 4 <= size; i++) {
          const px = view.getInt16(pos + 28 + i * 4, true);
          const py = view.getInt16(pos + 30 + i * 4, true);
          pts.push(`${px},${py}`);
        }
        if (pts.length > 0) {
          svgItems.push(`<polyline points="${pts.join(' ')}" fill="none" stroke="#334155" stroke-width="1"/>`);
        }
      } else if (type === 82 && size >= 104) {
        const ihFont = view.getUint32(pos + 8, true);
        const lfHeight = Math.abs(view.getInt32(pos + 12, true));
        fonts[ihFont] = { size: Math.max(10, Math.round(lfHeight * 0.85)), family: 'sans-serif' };
      } else if (type === 37 && size >= 12) {
        const objIdx = view.getUint32(pos + 8, true);
        if (fonts[objIdx]) curFont = fonts[objIdx];
      } else if (type === 84 && size >= 76) {
        const refX = view.getInt32(pos + 36, true);
        const refY = view.getInt32(pos + 40, true);
        const nChars = view.getUint32(pos + 44, true);
        const offString = view.getUint32(pos + 48, true);
        if (nChars > 0 && offString > 0 && pos + offString + nChars * 2 <= u8.byteLength) {
          const strBytes = u8.subarray(pos + offString, pos + offString + nChars * 2);
          const text = new TextDecoder('utf-16le').decode(strBytes).replace(/\0/g, '');
          if (text.trim()) {
            svgItems.push(`<text x="${refX}" y="${refY}" font-size="${curFont.size}" font-family="${curFont.family}" fill="#0f172a">${escapeXmlText(text)}</text>`);
          }
        }
      }
      pos += size;
    }

    if (dibBlobs.length === 1 && svgItems.length === 0) {
      return new Blob([dibBlobs[0]], { type: 'image/bmp' });
    }
    if (svgItems.length > 0) {
      const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${left} ${top} ${width} ${height}">${svgItems.join('')}</svg>`;
      return new Blob([svg], { type: 'image/svg+xml;charset=utf-8' });
    }
    if (dibBlobs.length > 0) {
      return new Blob([dibBlobs[0]], { type: 'image/bmp' });
    }
    return null;
  } catch (_) {
    return null;
  }
}

module.exports = {
  dibToBmpDataUrl,
  convertEmfToBlob
};
