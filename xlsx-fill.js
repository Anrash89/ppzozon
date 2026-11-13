/* Modify only the QR-code column of the original .xlsx ZIP package.
 * Other sheets, styles, cell values, and original formatting are retained.
 * JSZip is bundled in vendor/; browser DOMParser/XMLSerializer are native.
 */

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

function children(node, name) {
  return [...node.childNodes].filter(x => x.nodeType === 1 && (!name || x.localName === name));
}
function first(node, name) {
  return node.getElementsByTagNameNS(MAIN_NS, name)[0] || null;
}
function failOnXmlError(doc, label) {
  const error = doc.getElementsByTagName('parsererror')[0];
  if (error) throw new Error(`Ошибка XML в файле Excel (${label}).`);
}
function parseXml(text, label) {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  failOnXmlError(doc, label);
  return doc;
}
function colFromRef(ref) { return (ref.match(/^[A-Z]+/i) || [''])[0].toUpperCase(); }
function rowFromRef(ref) { return Number((ref.match(/\d+$/) || [0])[0]); }
function columnIndex(letters) {
  return [...letters].reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0);
}
function columnLetters(index) {
  let s = '';
  for (; index > 0; index = Math.floor((index - 1) / 26)) s = String.fromCharCode(65 + (index - 1) % 26) + s;
  return s;
}
function normalizedHeader(value) {
  return String(value || '').toLowerCase().replace(/[\s_\-–—]/g, '').replace(/ё/g, 'е');
}
function cellText(cell, strings) {
  if (!cell) return '';
  const v = children(cell, 'v')[0];
  if (cell.getAttribute('t') === 's' && v) return strings[Number(v.textContent)] || '';
  const inline = children(cell, 'is')[0];
  if (inline) return [...inline.getElementsByTagNameNS(MAIN_NS, 't')].map(n => n.textContent).join('');
  return v?.textContent || '';
}
function mapCells(row) {
  const m = new Map();
  for (const cell of children(row, 'c')) m.set(colFromRef(cell.getAttribute('r') || ''), cell);
  return m;
}
function makeStringCell(doc, cell, text) {
  cell.setAttribute('t', 'inlineStr');
  for (const child of [...cell.childNodes]) {
    if (child.nodeType === 1 && ['v', 'f', 'is'].includes(child.localName)) cell.removeChild(child);
  }
  if (text) {
    const is = doc.createElementNS(MAIN_NS, 'is');
    const t = doc.createElementNS(MAIN_NS, 't');
    t.textContent = text;
    is.appendChild(t);
    cell.appendChild(is);
  }
}
function qrCell(doc, row, col, rowNumber) {
  const cells = mapCells(row);
  if (cells.has(col)) return cells.get(col);
  const el = doc.createElementNS(MAIN_NS, 'c');
  el.setAttribute('r', `${col}${rowNumber}`);
  let before = null;
  for (const cell of children(row, 'c')) {
    if (columnIndex(colFromRef(cell.getAttribute('r') || '')) > columnIndex(col)) { before = cell; break; }
  }
  row.insertBefore(el, before);
  return el;
}
function getRowCell(row, col, strings) { return cellText(mapCells(row).get(col), strings); }

async function loadDocument(arrayBuffer) {
  if (typeof JSZip === 'undefined') throw new Error('Не загрузилась встроенная библиотека ZIP.');
  let zip;
  try { zip = await JSZip.loadAsync(arrayBuffer); }
  catch { throw new Error('Не удалось открыть Excel. Нужен файл .xlsx.'); }
  if (!zip.file('xl/workbook.xml') || !zip.file('xl/_rels/workbook.xml.rels')) {
    throw new Error('Файл не похож на стандартный Excel .xlsx.');
  }
  const wbDoc = parseXml(await zip.file('xl/workbook.xml').async('string'), 'workbook');
  const relDoc = parseXml(await zip.file('xl/_rels/workbook.xml.rels').async('string'), 'relationships');
  const relationships = new Map();
  for (const rel of [...relDoc.getElementsByTagNameNS(REL_NS, 'Relationship')]) {
    relationships.set(rel.getAttribute('Id'), rel.getAttribute('Target'));
  }
  const sheetDetails = [];
  for (const el of [...wbDoc.getElementsByTagNameNS(MAIN_NS, 'sheet')]) {
    const relId = el.getAttribute('r:id') || el.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    const target = relationships.get(relId);
    if (!target) continue;
    let path = target.replace(/^\//, '').replace(/^\.\//, '');
    if (!path.startsWith('xl/')) path = 'xl/' + path;
    if (zip.file(path)) sheetDetails.push({ name: el.getAttribute('name') || '', path });
  }
  const strings = [];
  if (zip.file('xl/sharedStrings.xml')) {
    const sharedDoc = parseXml(await zip.file('xl/sharedStrings.xml').async('string'), 'sharedStrings');
    for (const si of [...sharedDoc.getElementsByTagNameNS(MAIN_NS, 'si')]) {
      strings.push([...si.getElementsByTagNameNS(MAIN_NS, 't')].map(n => n.textContent).join(''));
    }
  }
  for (const detail of sheetDetails) {
    const doc = parseXml(await zip.file(detail.path).async('string'), detail.name);
    const data = first(doc, 'sheetData');
    if (!data) continue;
    const rows = children(data, 'row');
    for (const row of rows.slice(0, 25)) {
      const cells = mapCells(row);
      let codeCol = '';
      for (const [col, c] of cells) {
        const h = normalizedHeader(cellText(c, strings));
        if (h === 'qrкод' || h === 'qrcode' || h === 'qr' || h === 'qrкодтовара') { codeCol = col; break; }
      }
      if (!codeCol) continue;
      const headerRow = Number(row.getAttribute('r'));
      const dataRows = rows.filter(r => Number(r.getAttribute('r')) > headerRow &&
        [...mapCells(r)].some(([col, cell]) => col !== codeCol && cellText(cell, strings).trim() !== ''));
      const products = new Set(dataRows.map(r => ['A', 'B', 'D'].map(c => getRowCell(r, c, strings)).join('|')));
      return {
        zip, doc, data, sheetPath: detail.path, sheetName: detail.name,
        codeCol, headerRow, dataRows, strings, distinctProducts: products.size,
        preview: dataRows.map(r => ({
          row: Number(r.getAttribute('r')),
          barcode: getRowCell(r, 'A', strings),
          article: getRowCell(r, 'B', strings),
          place: getRowCell(r, 'D', strings),
          currentQR: getRowCell(r, codeCol, strings)
        }))
      };
    }
  }
  throw new Error('В Excel не найден столбец «QR-код». Проверь исходный шаблон Ozon.');
}

export async function inspectTemplate(arrayBuffer) {
  const document = await loadDocument(arrayBuffer);
  if (!document.dataRows.length) throw new Error('В шаблоне нет строк с товарами.');
  return {
    sheetName: document.sheetName,
    rowCount: document.dataRows.length,
    headerRow: document.headerRow,
    codeColumn: document.codeCol,
    distinctProducts: document.distinctProducts,
    preview: document.preview
  };
}

function clonedProductRow(doc, baseRow, newIndex) {
  const clone = baseRow.cloneNode(true);
  clone.setAttribute('r', String(newIndex));
  for (const cell of children(clone, 'c')) {
    const col = colFromRef(cell.getAttribute('r') || '');
    cell.setAttribute('r', col + newIndex);
  }
  return clone;
}

function extendDimension(doc, newRow, codeCol) {
  const dim = first(doc, 'dimension');
  if (!dim) return;
  const val = dim.getAttribute('ref') || '';
  const parts = val.split(':');
  const last = parts[parts.length - 1];
  const firstRow = parts.length === 1 ? 'A1' : parts[0];
  const lastColumn = colFromRef(last);
  const lastRow = rowFromRef(last);
  const desiredCol = columnIndex(lastColumn) >= columnIndex(codeCol) ? lastColumn : codeCol;
  if (newRow > lastRow) dim.setAttribute('ref', `${firstRow}:${desiredCol}${newRow}`);
}

/** mode = 'existing' (safe default) or 'extend' (clone last product row). */
export async function createFilledWorkbook(originalBuffer, qrCodes, mode = 'existing') {
  if (!['existing', 'extend'].includes(mode)) throw new Error('Неизвестный режим заполнения.');
  const entries = qrCodes.map(x => typeof x === 'string' ? x : x.code);
  if (!entries.length) throw new Error('Не найдены QR-коды для переноса.');
  if (entries.some(x => !/^ii\d{9,22}$/i.test(x))) throw new Error('В PDF обнаружен QR-код неожиданного формата.');
  if (new Set(entries).size !== entries.length) throw new Error('В PDF есть повторяющиеся QR-коды. Исправь PDF перед выгрузкой.');
  const ctx = await loadDocument(originalBuffer);
  if (mode === 'extend' && ctx.distinctProducts > 1 && entries.length > ctx.dataRows.length) {
    throw new Error('В Excel несколько разных товаров. Нельзя автоматически копировать последний товар для новых строк.');
  }
  const targetRows = ctx.dataRows.slice();
  const lastRow = targetRows[targetRows.length - 1];
  if (!lastRow) throw new Error('В Excel нет строк товаров для заполнения.');
  if (mode === 'extend' && entries.length > targetRows.length) {
    let index = Number(lastRow.getAttribute('r'));
    for (let i = targetRows.length; i < entries.length; i++) {
      const duplicate = clonedProductRow(ctx.doc, lastRow, ++index);
      // Preserve row number order if there are later non-product rows.
      const nextRow = children(ctx.data, 'row').find(r => Number(r.getAttribute('r')) > index);
      ctx.data.insertBefore(duplicate, nextRow || null);
      targetRows.push(duplicate);
    }
    extendDimension(ctx.doc, index, ctx.codeCol);
  }
  for (let i = 0; i < targetRows.length; i++) {
    const row = targetRows[i];
    const index = Number(row.getAttribute('r'));
    const cell = qrCell(ctx.doc, row, ctx.codeCol, index);
    makeStringCell(ctx.doc, cell, entries[i] || '');
  }
  const originalSheetXml = new XMLSerializer().serializeToString(ctx.doc);
  ctx.zip.file(ctx.sheetPath, originalSheetXml);
  const array = await ctx.zip.generateAsync({ type:'uint8array', compression:'DEFLATE', compressionOptions:{level:6} });
  return {
    bytes: array, assigned: Math.min(entries.length, targetRows.length),
    unused: Math.max(0, entries.length - targetRows.length),
    blank: Math.max(0, targetRows.length - entries.length),
    extended: Math.max(0, targetRows.length - ctx.dataRows.length)
  };
}
