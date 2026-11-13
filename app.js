import { extractQrCodes } from './pdf-extract.js';
import { inspectTemplate, createFilledWorkbook } from './xlsx-fill.js';

const $ = id => document.getElementById(id);
const state = {
  pdfFile: null, excelFile: null, excelBuffer: null,
  pdf: null, excel: null, mode: 'existing', showAll: false,
  runId: 0, busy: false, ready: false
};

function readableSize(size) {
  return size >= 1048576 ? `${(size / 1048576).toFixed(1)} МБ` : `${Math.max(1, Math.round(size / 1024))} КБ`;
}
function setStatus(kind, title, detail = '', progress = null) {
  $('statusIcon').className = `status-circle ${kind}`;
  $('statusIcon').textContent = ({ success:'✓', error:'!', loading:'↻', neutral:'i' })[kind] || 'i';
  $('statusTitle').textContent = title;
  $('statusDescription').textContent = detail;
  $('progressTrack').hidden = progress == null;
  if (progress != null) $('progressBar').style.width = `${Math.round(progress * 100)}%`;
}
function fileAccepted(file, type) {
  if (!file) return false;
  return type === 'pdf' ? /\.pdf$/i.test(file.name) : /\.xlsx$/i.test(file.name);
}
function setFile(file, type) {
  if (!fileAccepted(file, type)) {
    setStatus('error', 'Неверный формат файла', `Выбери ${type === 'pdf' ? 'документ .pdf' : 'таблицу .xlsx'}.`);
    return;
  }
  state.runId++;
  state.ready = false; state.pdf = null; state.excel = null;
  state.excelBuffer = null; state.mode = 'existing'; state.showAll = false;
  $('resultPanel').hidden = true;
  state[`${type}File`] = file;
  const input = $(type === 'pdf' ? 'pdfInput' : 'excelInput');
  if (input.files?.[0] !== file) input.value = '';
  const name = $(type === 'pdf' ? 'pdfFileName' : 'excelFileName');
  name.textContent = `✓ ${file.name} · ${readableSize(file.size)}`;
  $(type === 'pdf' ? 'pdfDrop' : 'excelDrop').classList.add('has-file');
  $(type === 'pdf' ? 'pdfAction' : 'excelAction').firstChild.nodeValue = 'Изменить файл ';
  if (!state.pdfFile || !state.excelFile) {
    setStatus('neutral', 'Первый файл добавлен', 'Добавь второй файл — и мы сразу начнём обработку.');
    return;
  }
  void analyze();
}
function installDropzone(zoneId, inputId, kind) {
  const zone = $(zoneId), input = $(inputId);
  zone.addEventListener('click', () => input.click());
  zone.addEventListener('keydown', event => {
    if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); input.click(); }
  });
  input.addEventListener('change', () => {
    if (input.files?.length) setFile(input.files[0], kind);
  });
  for (const name of ['dragenter', 'dragover']) zone.addEventListener(name, e => {
    e.preventDefault(); zone.classList.add('drag-over');
  });
  for (const name of ['dragleave', 'drop']) zone.addEventListener(name, e => {
    e.preventDefault(); zone.classList.remove('drag-over');
  });
  zone.addEventListener('drop', e => {
    const files = [...(e.dataTransfer?.files || [])];
    const expected = files.find(f => fileAccepted(f, kind));
    if (expected) setFile(expected, kind);
    else setStatus('error', 'Неподходящий файл', `Нужен формат ${kind === 'pdf' ? '.pdf' : '.xlsx'}.`);
    // Supports dropping PDF and XLSX together on either zone.
    const other = kind === 'pdf' ? 'excel' : 'pdf';
    const second = files.find(f => fileAccepted(f, other));
    if (second) setFile(second, other);
  });
}

async function analyze() {
  const token = ++state.runId;
  state.busy = true; state.ready = false;
  setStatus('loading', 'Обрабатываем документы…', 'Читаем QR-коды по порядку страниц PDF.', 0.04);
  try {
    const [pdfBuffer, excelBuffer] = await Promise.all([
      state.pdfFile.arrayBuffer(), state.excelFile.arrayBuffer()
    ]);
    if (token !== state.runId) return;
    state.excelBuffer = excelBuffer;
    const qr = await extractQrCodes(pdfBuffer, fraction => {
      if (token === state.runId) setStatus('loading', 'Читаем PDF…', `${Math.round(fraction * 100)}% страниц обработано`, .08 + fraction * .70);
    });
    if (token !== state.runId) return;
    setStatus('loading', 'Проверяем Excel…', 'Ищем лист и столбец «QR-код».', .87);
    const excel = await inspectTemplate(excelBuffer);
    if (token !== state.runId) return;
    state.pdf = qr;
    state.excel = excel;
    state.ready = qr.codes.length > 0 && qr.duplicates === 0 && qr.failedPages.length === 0;
    state.mode = 'existing';
    $('resultPanel').hidden = false;
    renderResult();
    if (!qr.codes.length) {
      setStatus('error', 'QR-коды не найдены', 'PDF должен содержать выделяемый текст с кодами вида ii50097774701. Сканированные изображения не поддерживаются.');
    } else if (qr.failedPages.length) {
      setStatus('error', 'Не все страницы прочитаны', `Без кода: страницы ${qr.failedPages.slice(0, 15).join(', ')}${qr.failedPages.length > 15 ? '…' : ''}. Проверь PDF.`);
    } else if (qr.duplicates) {
      setStatus('error', 'Есть повторяющиеся QR-коды', `${qr.duplicates} дублей. Перенос остановлен, чтобы не создать некорректный Excel.`);
    } else {
      setStatus('success', `${qr.codes.length} QR-кодов успешно считаны`, 'Проверь количество строк и скачай готовый Excel.');
    }
  } catch (err) {
    if (token !== state.runId) return;
    $('resultPanel').hidden = true;
    state.ready = false;
    setStatus('error', 'Не удалось обработать файлы', err instanceof Error ? err.message : 'Неизвестная ошибка.');
    console.error(err);
  } finally {
    if (token === state.runId) state.busy = false;
  }
}

function getEffectiveRows() {
  if (!state.excel || !state.pdf) return [];
  const codes = state.pdf.codes;
  const rows = state.excel.preview.map(r => ({ ...r, isNew: false }));
  if (state.mode === 'extend' && codes.length > rows.length && state.excel.distinctProducts <= 1) {
    const last = rows.at(-1);
    for (let i = rows.length; i < codes.length; i++) {
      rows.push({ ...last, row: last.row + i - state.excel.rowCount + 1, isNew: true, currentQR: '' });
    }
  }
  return rows.map((row, i) => ({ ...row, nextQR: codes[i]?.code || '' }));
}
function appendCell(tr, text, klass = '') {
  const td = document.createElement('td');
  td.textContent = String(text || '—');
  if (klass) td.className = klass;
  tr.appendChild(td);
  return td;
}
function renderPreview() {
  const rows = getEffectiveRows();
  const visible = state.showAll ? rows : rows.slice(0, 8);
  const body = $('previewBody');
  body.replaceChildren();
  for (const r of visible) {
    const tr = document.createElement('tr');
    if (r.isNew) tr.className = 'new-row';
    const n = appendCell(tr, r.row);
    if (r.isNew) { const span = document.createElement('span'); span.className = 'new-row-label'; span.textContent = 'новая'; n.appendChild(span); }
    appendCell(tr, r.article);
    appendCell(tr, r.barcode);
    appendCell(tr, r.currentQR, 'code-old');
    appendCell(tr, r.nextQR, 'code-new');
    body.appendChild(tr);
  }
  $('toggleRows').textContent = state.showAll ? 'Свернуть' : `Показать все ${rows.length}`;
  $('toggleRows').hidden = rows.length <= 8;
  $('previewNote').textContent = state.showAll || rows.length <= 8
    ? `Показаны все ${rows.length} строк${rows.length === 1 ? 'а' : ''}.`
    : `Показаны первые 8 из ${rows.length} строк. Порядок QR-кодов соответствует порядку страниц PDF.`;
}
function renderResult() {
  const pdfCount = state.pdf.codes.length;
  const excelCount = state.excel.rowCount;
  const distinct = state.excel.distinctProducts;
  const canExtend = distinct <= 1;
  if (!canExtend && state.mode === 'extend') state.mode = 'existing';
  $('statCodes').textContent = pdfCount;
  $('statPages').textContent = `${state.pdf.totalPages} страниц PDF`;
  $('statRows').textContent = excelCount;
  $('statSheet').textContent = state.excel.sheetName;
  $('statFilled').textContent = Math.min(pdfCount, state.mode === 'extend' ? Math.max(pdfCount, excelCount) : excelCount);
  $('statMode').textContent = state.mode === 'extend' ? `включая ${Math.max(0, pdfCount - excelCount)} новых` : 'строк получат QR';
  $('modeSection').hidden = pdfCount <= excelCount;
  $('modeDescription').textContent = `В PDF ${pdfCount} кодов, а в Excel ${excelCount} строк товаров. Выбери способ заполнения.`;
  $('existingHint').textContent = `Заполнить ${Math.min(excelCount, pdfCount)} строк, не использовать ${Math.max(0, pdfCount - excelCount)} лишних QR`;
  $('extendHint').textContent = canExtend ? `Добавить ${Math.max(0, pdfCount - excelCount)} строк, копируя товар из последней строки` : 'Недоступно: в таблице несколько разных товаров';
  $('extendWrap').querySelector('input').disabled = !canExtend;
  for (const radio of document.querySelectorAll('input[name="fillMode"]')) radio.checked = radio.value === state.mode;
  const warn = $('modeWarning');
  if (pdfCount > excelCount && state.mode === 'extend') warn.textContent = `Будут созданы ${pdfCount - excelCount} дополнительные строки с тем же артикулом и грузоместом, что в последней строке. Проверь количество товаров перед загрузкой в Ozon.`;
  else if (pdfCount > excelCount) warn.textContent = `${pdfCount - excelCount} QR-кодов из PDF не попадут в Excel. Это безопаснее, чем добавлять товары без подтверждения.`;
  else warn.textContent = '';
  if (pdfCount < excelCount) {
    $('downloadHeading').textContent = 'Внимание: не хватает QR-кодов';
    $('downloadNote').textContent = `${excelCount - pdfCount} строк останутся без QR. Перед импортом в Ozon проверь, что количество товаров соответствует PDF.`;
  } else {
    $('downloadHeading').textContent = 'Готово к выгрузке';
    $('downloadNote').textContent = 'Сохраним исходный Excel-шаблон, его стили и лист «Инструкция».';
  }
  $('downloadBtn').disabled = !state.ready || state.busy;
  renderPreview();
}

async function downloadWorkbook() {
  if (!state.ready || !state.excelBuffer || state.busy) return;
  state.busy = true;
  $('downloadBtn').disabled = true;
  $('downloadBtn').querySelector('span').textContent = 'Создаём Excel…';
  try {
    const output = await createFilledWorkbook(state.excelBuffer, state.pdf.codes, state.mode);
    const blob = new Blob([output.bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const href = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = href;
    link.download = `Ozon_Состав_грузомест_QR_${output.assigned}.xlsx`;
    document.body.appendChild(link);
    link.click(); link.remove();
    setTimeout(() => URL.revokeObjectURL(href), 1000);
    setStatus('success', 'Excel сформирован', `Заполнено ${output.assigned} строк${output.extended ? `, добавлено ${output.extended} новых` : ''}. Файл передан браузеру для скачивания.`);
  } catch (error) {
    setStatus('error', 'Ошибка при создании Excel', error instanceof Error ? error.message : 'Неизвестная ошибка');
    console.error(error);
  } finally {
    state.busy = false;
    $('downloadBtn').disabled = !state.ready;
    $('downloadBtn').querySelector('span').textContent = 'Скачать готовый Excel';
  }
}
function resetAll() {
  state.runId++;
  state.pdfFile = null; state.excelFile = null; state.excelBuffer = null;
  state.pdf = null; state.excel = null; state.mode = 'existing'; state.showAll = false;
  state.ready = false; state.busy = false;
  for (const type of ['pdf', 'excel']) {
    $(`${type}Input`).value = '';
    $(`${type}Drop`).classList.remove('has-file');
    $(`${type}FileName`).textContent = type === 'pdf' ? 'Перетащи PDF сюда или выбери файл' : 'Перетащи Excel сюда или выбери файл';
    $(`${type}Action`).firstChild.nodeValue = type === 'pdf' ? 'Выбрать PDF ' : 'Выбрать Excel ';
  }
  $('resultPanel').hidden = true;
  setStatus('neutral', 'Добавь два файла, чтобы начать', 'Обработка начнётся автоматически после загрузки.');
  window.scrollTo({ top:0, behavior:'smooth' });
}

installDropzone('pdfDrop', 'pdfInput', 'pdf');
installDropzone('excelDrop', 'excelInput', 'excel');
for (const radio of document.querySelectorAll('input[name="fillMode"]')) {
  radio.addEventListener('change', e => {
    if (!e.target.checked) return;
    state.mode = e.target.value;
    state.showAll = false;
    renderResult();
  });
}
$('downloadBtn').addEventListener('click', downloadWorkbook);
$('resetBtn').addEventListener('click', resetAll);
$('toggleRows').addEventListener('click', () => { state.showAll = !state.showAll; renderPreview(); });
