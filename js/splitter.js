/*
 * Логика разделителя: чтение книги, поиск шапки, группировка строк
 * и сборка отдельных .xlsx с сохранением оформления.
 * Работает и в браузере (window.Splitter), и в Node (module.exports) — для тестов.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Splitter = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const isNode = typeof window === 'undefined';
  const getExcelJS = () => (isNode ? require('exceljs') : window.ExcelJS);
  const getXLSX = () => (isNode ? require('xlsx') : window.XLSX);

  const tick = () => new Promise((r) => setTimeout(r, 0));

  // Форматы, которые ExcelJS читает напрямую (с полным оформлением)
  const OOXML = ['xlsx', 'xlsm', 'xltx', 'xltm'];
  // Всё, что умеет SheetJS (оформление частично)
  const ACCEPT = ['xlsx', 'xlsm', 'xltx', 'xltm', 'xlsb', 'xls', 'xlt', 'xla', 'ods', 'fods', 'csv', 'tsv', 'txt', 'xml', 'htm', 'html', 'dif', 'sylk', 'slk', 'prn', 'dbf', 'numbers', 'wk1', 'wk3', 'qpw'];

  function extOf(name) {
    const m = /\.([^.]+)$/.exec(name || '');
    return m ? m[1].toLowerCase() : '';
  }
  function baseName(name) {
    return (name || 'файл').replace(/\.[^.]+$/, '');
  }

  function colLetter(n) {
    let s = '';
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }
  function colNumber(letters) {
    let n = 0;
    for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }
  function parseAddr(a) {
    const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(String(a).trim());
    return m ? { col: colNumber(m[1]), row: +m[2] } : null;
  }
  function parseRange(ref) {
    const [a, b] = String(ref).split(':');
    const s = parseAddr(a);
    const e = parseAddr(b || a);
    if (!s || !e) return null;
    return { top: Math.min(s.row, e.row), left: Math.min(s.col, e.col), bottom: Math.max(s.row, e.row), right: Math.max(s.col, e.col) };
  }
  const rangeStr = (r) => `${colLetter(r.left)}${r.top}:${colLetter(r.right)}${r.bottom}`;

  const pad = (n) => String(n).padStart(2, '0');
  function fmtDate(d) {
    let s = `${pad(d.getUTCDate())}.${pad(d.getUTCMonth() + 1)}.${d.getUTCFullYear()}`;
    if (d.getUTCHours() || d.getUTCMinutes()) s += ` ${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}`;
    return s;
  }

  // Значение ячейки -> текст для группировки и предпросмотра
  function valueToText(v) {
    if (v == null) return '';
    if (v instanceof Date) return isNaN(v) ? '' : fmtDate(v);
    if (typeof v === 'number') {
      if (!isFinite(v)) return '';
      const n = Number.isInteger(v) ? v : parseFloat(v.toPrecision(12));
      return String(n).replace('.', ',');
    }
    if (typeof v === 'boolean') return v ? 'ИСТИНА' : 'ЛОЖЬ';
    if (typeof v === 'object') {
      if (Array.isArray(v.richText)) return v.richText.map((t) => t.text || '').join('');
      if ('formula' in v || 'sharedFormula' in v) return valueToText(v.result);
      if (v.text != null) return valueToText(v.text);
      if (v.error) return String(v.error);
      return '';
    }
    return String(v);
  }

  const normKey = (t) => t.replace(/[\s ]+/g, ' ').trim().toLowerCase();
  const tidy = (t) => t.replace(/[\s ]+/g, ' ').trim();

  // Значение для записи в новый файл: формулы заменяем результатом
  // (строки переезжают на новые места, и ссылки формул сломались бы).
  function outValue(v) {
    if (v == null) return null;
    if (v instanceof Date) return new Date(v.getTime());
    if (typeof v === 'object') {
      if ('formula' in v || 'sharedFormula' in v) return outValue(v.result);
      return JSON.parse(JSON.stringify(v));
    }
    return v;
  }
  const clone = (o) => (o == null ? o : JSON.parse(JSON.stringify(o)));

  /* ---------- Чтение файла ---------- */

  function decodeText(buf) {
    const bytes = new Uint8Array(buf);
    try {
      return new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^﻿/, '');
    } catch (e) {
      return new TextDecoder('windows-1251').decode(bytes);
    }
  }

  async function viaSheetJS(buf, ext) {
    const XLSX = getXLSX();
    let wbSrc;
    if (['csv', 'tsv', 'txt', 'prn'].includes(ext)) {
      // Читаем всё как текст, числа распознаём сами: «4,5» — это 4.5, а не 45
      wbSrc = XLSX.read(decodeText(buf), { type: 'string', raw: true });
      for (const name of wbSrc.SheetNames) {
        const sh = wbSrc.Sheets[name];
        for (const addr of Object.keys(sh)) {
          if (addr[0] === '!') continue;
          const cell = sh[addr];
          if (cell.t !== 's') continue;
          const s = String(cell.v).trim();
          if (/^-?(0|[1-9]\d*)([.,]\d+)?$/.test(s) && s.length < 16) {
            cell.t = 'n';
            cell.v = Number(s.replace(',', '.'));
            delete cell.w;
          }
        }
      }
    } else {
      wbSrc = XLSX.read(new Uint8Array(buf), { type: 'array', cellNF: true, cellStyles: true });
    }
    const out = XLSX.write(wbSrc, { bookType: 'xlsx', type: 'array', cellStyles: true });
    const wb = new (getExcelJS().Workbook)();
    await wb.xlsx.load(out);
    return wb;
  }

  /**
   * Загружает файл. Возвращает { workbook, sheets, warning }.
   * warning — текст, если оформление могло потеряться.
   */
  async function loadWorkbook(buf, fileName) {
    const ext = extOf(fileName);
    let workbook;
    let warning = null;
    if (OOXML.includes(ext)) {
      try {
        workbook = new (getExcelJS().Workbook)();
        await workbook.xlsx.load(buf);
      } catch (e) {
        workbook = await viaSheetJS(buf, ext);
        warning = 'Файл прочитан запасным способом — часть оформления могла потеряться.';
      }
      if (ext === 'xlsm' || ext === 'xltm') warning = 'Макросы в новые файлы не переносятся — сохраняем как обычный .xlsx.';
    } else {
      workbook = await viaSheetJS(buf, ext);
      warning = `Формат .${ext || '?'} переводим в .xlsx. Данные сохранятся полностью, оформление — частично (цвета и шрифты старых форматов браузер прочитать не может).`;
    }
    const sheets = workbook.worksheets
      .filter((ws) => ws && ws.state !== 'veryHidden')
      .map((ws) => ({ id: ws.id, name: ws.name, hidden: ws.state === 'hidden' }));
    if (!sheets.length) throw new Error('В файле не нашлось ни одного листа с таблицей.');
    return { workbook, sheets, warning };
  }

  /* ---------- Разбор листа ---------- */

  function isSlave(cell) {
    return cell.isMerged && cell.master && cell.master !== cell;
  }

  /**
   * Находит шапку, колонки и строки с данными.
   */
  async function analyzeSheet(ws) {
    const rowCount = ws.rowCount;
    const texts = new Array(rowCount + 1);
    const ownCount = new Array(rowCount + 1).fill(0);
    let lastRow = 0;
    let lastCol = 0;
    let hasFormulas = false;

    for (let r = 1; r <= rowCount; r++) {
      const row = ws.getRow(r);
      const arr = [];
      const n = row.cellCount;
      for (let c = 1; c <= n; c++) {
        const cell = row.getCell(c);
        const v = cell.value;
        if (v && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v)) hasFormulas = true;
        const t = valueToText(v);
        arr[c] = t;
        if (t.trim() !== '') {
          if (!isSlave(cell)) ownCount[r]++;
          if (c > lastCol) lastCol = c;
          lastRow = r;
        }
      }
      texts[r] = arr;
      if (r % 5000 === 0) await tick();
    }
    if (!lastRow) throw new Error('Лист пустой — в нём нет данных.');

    // Шапка: первая строка, где заполнена хотя бы половина колонок
    // (строка-название таблицы, объединённая в одну ячейку, пропускается).
    let maxCount = 0;
    for (let r = 1; r <= Math.min(lastRow, 50); r++) maxCount = Math.max(maxCount, ownCount[r]);
    const threshold = maxCount >= 2 ? Math.max(2, Math.ceil(maxCount * 0.5)) : 1;
    let headerRow = 1;
    for (let r = 1; r <= Math.min(lastRow, 30); r++) {
      if (ownCount[r] >= threshold) { headerRow = r; break; }
    }

    const cellT = (r, c) => (texts[r] && texts[r][c]) || '';
    const header = [];
    for (let c = 1; c <= lastCol; c++) header[c] = tidy(cellT(headerRow, c));
    const headerKeys = header.map((h) => (h ? normKey(h) : ''));
    const headerFilled = headerKeys.filter(Boolean).length;

    // Строки данных: без пустых и без повторов шапки посреди таблицы
    const dataRows = [];
    let skippedHeaders = 0;
    let skippedEmpty = 0;
    for (let r = headerRow + 1; r <= lastRow; r++) {
      let filled = 0;
      let same = 0;
      for (let c = 1; c <= lastCol; c++) {
        const t = cellT(r, c);
        if (t.trim() === '') continue;
        filled++;
        if (headerKeys[c] && normKey(t) === headerKeys[c]) same++;
      }
      if (!filled) { skippedEmpty++; continue; }
      if (headerFilled >= 2 && same >= Math.max(2, Math.ceil(headerFilled * 0.6))) { skippedHeaders++; continue; }
      dataRows.push(r);
    }

    const columns = [];
    const usedNames = new Set();
    for (let c = 1; c <= lastCol; c++) {
      const seen = new Map();
      let empty = 0;
      for (const r of dataRows) {
        const t = cellT(r, c);
        const k = normKey(t);
        if (!k) { empty++; continue; }
        if (!seen.has(k)) seen.set(k, tidy(t));
      }
      if (!header[c] && !seen.size) continue;
      let name = header[c] || `Колонка ${colLetter(c)}`;
      if (usedNames.has(normKey(name))) name = `${name} (${colLetter(c)})`;
      usedNames.add(normKey(name));
      const labels = [...seen.values()].sort(compareLabels);
      columns.push({
        index: c,
        letter: colLetter(c),
        name,
        distinct: seen.size,
        empty,
        sample: labels.slice(0, 4),
      });
    }

    return { ws, texts, headerRow, lastRow, lastCol, dataRows, columns, skippedHeaders, skippedEmpty, hasFormulas };
  }

  const collator = new Intl.Collator('ru', { numeric: true, sensitivity: 'base' });
  function compareLabels(a, b) {
    const na = Number(a.replace(',', '.'));
    const nb = Number(b.replace(',', '.'));
    if (a !== '' && b !== '' && !isNaN(na) && !isNaN(nb)) return na - nb;
    return collator.compare(a, b);
  }

  /* ---------- Группировка ---------- */

  const BAD_CHARS = /[\\/:*?"<>|\u0000-\u001f]/g;
  function safePart(s, max) {
    let t = String(s).replace(BAD_CHARS, '_').replace(/\s+/g, ' ').trim();
    t = t.replace(/^[.\s]+|[.\s]+$/g, '');
    if (t.length > max) t = t.slice(0, max).trim();
    return t || '_';
  }

  /**
   * Делит строки по выбранной колонке.
   * Возвращает { groups: [{label, rows, isEmpty, fileName, variants}], merged: [...] }
   */
  function groupBy(analysis, colIndex, sourceFileName) {
    const col = analysis.columns.find((c) => c.index === colIndex);
    const map = new Map();
    for (const r of analysis.dataRows) {
      const raw = (analysis.texts[r] && analysis.texts[r][colIndex]) || '';
      const key = normKey(raw);
      let g = map.get(key);
      if (!g) {
        g = { key, label: key ? tidy(raw) : 'пусто', isEmpty: !key, rows: [], variants: new Set() };
        map.set(key, g);
      }
      g.rows.push(r);
      if (key) g.variants.add(raw);
    }
    const groups = [...map.values()].sort((a, b) => {
      if (a.isEmpty !== b.isEmpty) return a.isEmpty ? 1 : -1;
      return compareLabels(a.label, b.label);
    });

    const base = safePart(baseName(sourceFileName), 60);
    const colPart = safePart(col ? col.name : 'Колонка', 40);
    const used = new Set();
    for (const g of groups) {
      let name = `${colPart}_${safePart(g.label, 80)}_${base}`;
      let candidate = name;
      let i = 2;
      while (used.has(candidate.toLowerCase())) candidate = `${name} (${i++})`;
      used.add(candidate.toLowerCase());
      g.fileName = `${candidate}.xlsx`;
    }

    const merged = groups
      .filter((g) => g.variants.size > 1)
      .map((g) => ({ label: g.label, variants: [...g.variants] }));
    const changedName = groups.some((g) => !g.isEmpty && /[\\/:*?"<>|\u0000-\u001f]/.test(g.label));

    return { column: col, groups, merged, changedName };
  }

  /* ---------- Сборка нового файла ---------- */

  function copyCell(src, dst) {
    const v = src.value; // у объединённых ячеек вернёт значение главной ячейки
    const out = outValue(v);
    if (out !== null && out !== undefined) dst.value = out;
    const st = src.style;
    if (st && Object.keys(st).length) dst.style = clone(st);
  }

  function copyRow(srcWs, dstWs, srcR, dstR, lastCol) {
    const s = srcWs.getRow(srcR);
    const d = dstWs.getRow(dstR);
    if (s.height) d.height = s.height;
    if (s.hidden) d.hidden = true;
    if (s.outlineLevel) d.outlineLevel = s.outlineLevel;
    const n = Math.min(s.cellCount, lastCol);
    for (let c = 1; c <= n; c++) {
      const sc = s.getCell(c);
      if (sc.value == null && (!sc.style || !Object.keys(sc.style).length)) continue;
      copyCell(sc, d.getCell(c));
    }
  }

  function getMerges(ws) {
    const out = [];
    const m = ws._merges || {};
    for (const k of Object.keys(m)) {
      const r = m[k];
      const model = r && (r.model || r);
      if (model && model.top) out.push({ top: model.top, left: model.left, bottom: model.bottom, right: model.right });
    }
    return out;
  }

  async function buildGroupWorkbook(analysis, rows, opts) {
    const ExcelJS = getExcelJS();
    const src = analysis.ws;
    const { headerRow, lastCol } = analysis;
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Разделитель Excel';

    const pageSetup = clone(src.pageSetup) || {};
    delete pageSetup.printArea;
    const dst = wb.addWorksheet(src.name, {
      properties: clone(src.properties),
      pageSetup,
      headerFooter: clone(src.headerFooter),
      views: clone(src.views),
    });

    // Ширина и стиль колонок
    const colCount = Math.max(lastCol, src.columnCount || 0);
    for (let c = 1; c <= colCount; c++) {
      const sc = src.getColumn(c);
      const dc = dst.getColumn(c);
      if (sc.width) dc.width = sc.width;
      if (sc.hidden) dc.hidden = true;
      if (sc.outlineLevel) dc.outlineLevel = sc.outlineLevel;
      if (sc.style && Object.keys(sc.style).length) dc.style = clone(sc.style);
    }

    // Шапка (и всё, что над ней) — как есть
    for (let r = 1; r <= headerRow; r++) copyRow(src, dst, r, r, colCount);

    // Строки группы
    const rowMap = new Map();
    let out = headerRow;
    for (let i = 0; i < rows.length; i++) {
      out++;
      rowMap.set(rows[i], out);
      copyRow(src, dst, rows[i], out, colCount);
      if (i % 3000 === 2999) await tick();
    }
    const lastOut = Math.max(out, headerRow);

    // Объединения: в шапке — как были; в строке — переносим; через несколько строк — снимаем
    for (const m of getMerges(src)) {
      try {
        if (m.bottom <= headerRow) dst.mergeCells(m.top, m.left, m.bottom, m.right);
        else if (m.top === m.bottom && rowMap.has(m.top)) {
          const nr = rowMap.get(m.top);
          dst.mergeCells(nr, m.left, nr, m.right);
        }
      } catch (e) { /* пересекающиеся объединения пропускаем */ }
    }

    // Автофильтр
    if (src.autoFilter) {
      let range = null;
      const af = src.autoFilter;
      if (typeof af === 'string') range = parseRange(af);
      else if (af.from && af.to) {
        const f = typeof af.from === 'string' ? parseAddr(af.from) : { col: af.from.column, row: af.from.row };
        const t = typeof af.to === 'string' ? parseAddr(af.to) : { col: af.to.column, row: af.to.row };
        if (f && t) range = { left: f.col, right: t.col };
      }
      if (range) dst.autoFilter = rangeStr({ top: headerRow, bottom: lastOut, left: range.left, right: range.right });
    }

    // Условное форматирование: растягиваем на новые строки
    if (!(opts && opts.noConditional) && Array.isArray(src.conditionalFormattings)) {
      for (const cf of src.conditionalFormattings) {
        const refs = String(cf.ref || '').split(/\s+/).filter(Boolean).map(parseRange).filter(Boolean);
        const fixed = refs.map((r) => {
          if (r.bottom <= headerRow) return r;
          return { left: r.left, right: r.right, top: r.top <= headerRow ? r.top : headerRow + 1, bottom: Math.max(lastOut, headerRow + 1) };
        });
        if (fixed.length) dst.addConditionalFormatting({ ref: fixed.map(rangeStr).join(' '), rules: clone(cf.rules) });
      }
    }
    return wb;
  }

  async function buildGroupFile(analysis, rows) {
    try {
      const wb = await buildGroupWorkbook(analysis, rows);
      return new Uint8Array(await wb.xlsx.writeBuffer());
    } catch (e) {
      // Если условное форматирование не записалось — собираем без него
      const wb = await buildGroupWorkbook(analysis, rows, { noConditional: true });
      return new Uint8Array(await wb.xlsx.writeBuffer());
    }
  }

  /* ---------- Предпросмотр ---------- */

  function preview(analysis, rows, limit) {
    const cols = analysis.columns;
    const lim = limit || 20;
    return {
      header: cols.map((c) => c.name),
      rows: rows.slice(0, lim).map((r) => cols.map((c) => (analysis.texts[r] && analysis.texts[r][c.index]) || '')),
      total: rows.length,
    };
  }

  return {
    ACCEPT, extOf, baseName, colLetter, valueToText,
    loadWorkbook, analyzeSheet, groupBy, buildGroupFile, preview,
  };
});
