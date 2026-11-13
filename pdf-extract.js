/*
 * Lightweight PDF text extractor for selectable-text Ozon QR labels.
 * Runs fully locally. It is not an OCR or a complete PDF implementation.
 * Handles conventional page/content streams and ToUnicode font maps.
 */

function bytesToBinaryString(bytes) {
  const step = 8192;
  const chunks = [];
  for (let i = 0; i < bytes.length; i += step) {
    chunks.push(String.fromCharCode(...bytes.subarray(i, i + step)));
  }
  return chunks.join('');
}

function binaryStringToBytes(s) {
  const result = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) result[i] = s.charCodeAt(i) & 0xff;
  return result;
}

function findObjects(raw) {
  const objects = new Map();
  const pattern = /(?:^|[\r\n])(\d+)\s+(\d+)\s+obj\b/g;
  let match;
  while ((match = pattern.exec(raw))) {
    const start = match.index + match[0].length;
    const end = raw.indexOf('endobj', start);
    if (end < 0) break;
    objects.set(Number(match[1]), raw.slice(start, end));
    pattern.lastIndex = end + 6;
  }
  return objects;
}

async function streamOf(objects, id) {
  const body = objects.get(id);
  if (!body) return '';
  const found = /\bstream(?:\r\n|\n|\r)/.exec(body);
  if (!found) return '';
  const start = found.index + found[0].length;
  const preamble = body.slice(0, found.index);
  const specifiedLength = /\/Length\s+(\d+)\b/.exec(preamble);
  let chunk;
  if (specifiedLength) chunk = body.slice(start, start + Number(specifiedLength[1]));
  else {
    const stop = body.indexOf('endstream', start);
    if (stop < 0) return '';
    chunk = body.slice(start, stop).replace(/[\r\n]+$/, '');
  }
  if (!/\/FlateDecode\b/.test(preamble)) return chunk;
  if (typeof DecompressionStream !== 'function') {
    throw new Error('Для чтения сжатого PDF нужен современный браузер (Chrome или Edge).');
  }
  const compressed = binaryStringToBytes(chunk);
  const uncompressed = await new Response(
    new Blob([compressed]).stream().pipeThrough(new DecompressionStream('deflate'))
  ).arrayBuffer();
  return bytesToBinaryString(new Uint8Array(uncompressed));
}

function buildUnicodeMap(cmap) {
  const map = new Map();
  // Ozon PDFs in the supplied examples use 2-byte CID character codes.
  for (const m of cmap.matchAll(/<([0-9A-Fa-f]{2,8})>\s*<([0-9A-Fa-f]{2,8})>\s*<([0-9A-Fa-f]{4,8})>/g)) {
    const first = parseInt(m[1], 16), last = parseInt(m[2], 16), base = parseInt(m[3], 16);
    if (last < first || last - first > 2048) continue;
    const len = m[1].length;
    for (let c = first; c <= last; c++) map.set(c.toString(16).padStart(len, '0'), String.fromCodePoint(base + c - first));
  }
  for (const block of cmap.matchAll(/\b\d+\s+beginbfchar([\s\S]*?)endbfchar/g)) {
    for (const m of block[1].matchAll(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/g)) {
      try { map.set(m[1].toLowerCase(), String.fromCodePoint(parseInt(m[2], 16))); } catch { /* ignore */ }
    }
  }
  return map;
}

function decodePdfHex(hex, characterMap) {
  const len = characterMap.size ? Math.min(...Array.from(characterMap.keys(), x => x.length)) : 4;
  let result = '';
  for (let i = 0; i < hex.length; i += len) {
    const part = hex.slice(i, i + len).toLowerCase();
    if (!part) continue;
    const mapped = characterMap.get(part);
    if (mapped) result += mapped;
    else {
      const val = parseInt(part, 16);
      if (Number.isFinite(val) && val > 0 && val < 0x110000) result += String.fromCodePoint(val);
    }
  }
  return result;
}

function decodePdfLiteral(token) {
  return token.replace(/^\(/, '').replace(/\)$/, '')
    .replace(/\\([\\()])/g, '$1')
    .replace(/\\([nrt])/g, (_, c) => ({n:'\n',r:'\r',t:'\t'})[c]);
}

function scanTokens(content, maps) {
  const tokens = [];
  let font = ''; // current font is switched with the PDF Tf operator
  const pattern = /\/([A-Za-z0-9]+)\s+[+-]?[\d.]+\s+Tf|<([\da-fA-F]+)>\s*Tj|\(((?:\\.|[^\\)])*)\)\s*Tj|\[([\s\S]*?)\]\s*TJ/g;
  let m;
  while ((m = pattern.exec(content))) {
    if (m[1]) { font = m[1]; continue; }
    const charMap = maps.get(font) || new Map();
    if (m[2]) tokens.push(decodePdfHex(m[2], charMap));
    else if (m[3] !== undefined) tokens.push(decodePdfLiteral('(' + m[3] + ')'));
    else if (m[4]) {
      for (const sub of m[4].matchAll(/<([0-9a-fA-F]+)>|\(((?:\\.|[^\\)])*)\)/g)) {
        tokens.push(sub[1] ? decodePdfHex(sub[1], charMap) : decodePdfLiteral('(' + sub[2] + ')'));
      }
    }
  }
  return tokens;
}

function getPageObjectIds(objects) {
  return [...objects.entries()].filter(([, body]) => /\/Type\s*\/Page\b/.test(body.slice(0, 5000)))
    .map(([id]) => id);
}

function getPdfContentIds(pageBody) {
  const m = /\/Contents\s*(?:\[([\s\S]*?)\]|(\d+)\s+\d+\s+R)/.exec(pageBody);
  if (!m) return [];
  if (m[2]) return [Number(m[2])];
  return [...m[1].matchAll(/(\d+)\s+\d+\s+R/g)].map(v => Number(v[1]));
}

function getPageFontIds(body) {
  const fonts = new Map();
  for (const m of body.matchAll(/\/(F\d+)\s+(\d+)\s+\d+\s+R/g)) fonts.set(m[1], Number(m[2]));
  return fonts;
}

function getUnicodeId(body) {
  const m = /\/ToUnicode\s+(\d+)\s+\d+\s+R/.exec(body || '');
  return m ? Number(m[1]) : null;
}

/** Extracts QR codes in the natural PDF page order. */
export async function extractQrCodes(arrayBuffer, progress = () => {}) {
  const bytes = new Uint8Array(arrayBuffer);
  if (bytes.length < 8 || !bytesToBinaryString(bytes.subarray(0, 8)).startsWith('%PDF-')) {
    throw new Error('Выбранный PDF-файл повреждён или не является PDF.');
  }
  const raw = bytesToBinaryString(bytes);
  const objects = findObjects(raw);
  const pages = getPageObjectIds(objects);
  if (!pages.length) {
    throw new Error('Не удалось найти страницы PDF. Поддерживаются PDF Ozon с выделяемым текстом, без сканирования.');
  }
  const cmapCache = new Map();
  const result = [];
  const failedPages = [];
  for (let pageIndex = 0; pageIndex < pages.length; pageIndex++) {
    const pageBody = objects.get(pages[pageIndex]) || '';
    const fontIds = getPageFontIds(pageBody);
    const maps = new Map();
    for (const [fontName, fontId] of fontIds) {
      const cmapId = getUnicodeId(objects.get(fontId));
      if (cmapId == null) continue;
      if (!cmapCache.has(cmapId)) cmapCache.set(cmapId, buildUnicodeMap(await streamOf(objects, cmapId)));
      maps.set(fontName, cmapCache.get(cmapId));
    }
    const blocks = [];
    for (const id of getPdfContentIds(pageBody)) {
      const content = await streamOf(objects, id);
      blocks.push(...scanTokens(content, maps));
    }
    const joined = blocks.join('');
    const matches = [...joined.matchAll(/ii\d{9,22}/gi)].map(m => m[0]);
    if (matches.length === 0) failedPages.push(pageIndex + 1);
    const seenOnPage = new Set();
    for (const code of matches) {
      if (seenOnPage.has(code)) continue;
      seenOnPage.add(code);
      result.push({ code, page: pageIndex + 1 });
    }
    progress((pageIndex + 1) / pages.length);
  }
  const unique = new Set(result.map(x => x.code));
  const duplicates = result.length - unique.size;
  return { codes: result, totalPages: pages.length, failedPages, duplicates };
}
