/* Интерфейс разделителя: состояния, отрисовка, скачивание. */
(function () {
  'use strict';

  const S = window.Splitter;
  const $ = (id) => document.getElementById(id);
  const MANY = 50; // столько файлов и больше — сначала переспрашиваем
  const PREVIEW_ROWS = 20;
  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  const st = {
    job: 0,
    phase: 'idle', // idle | reading | analyzing | pick | many | splitting | done | error
    file: null,
    workbook: null,
    sheets: [],
    sheetId: null,
    warning: null,
    analysis: null,
    analysisError: null,
    colIndex: null,
    result: null,
    files: new Map(),
    checked: new Set(),
    progress: { done: 0, total: 0 },
    error: null,
    busy: false,
  };

  /* ---------- Помощники ---------- */

  const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fmt = (n) => n.toLocaleString('ru-RU');
  function plural(n, one, few, many) {
    const m10 = n % 10;
    const m100 = n % 100;
    if (m10 === 1 && m100 !== 11) return one;
    if (m10 >= 2 && m10 <= 4 && (m100 < 12 || m100 > 14)) return few;
    return many;
  }
  const rowsWord = (n) => `${fmt(n)} ${plural(n, 'строка', 'строки', 'строк')}`;
  const filesWord = (n) => `${fmt(n)} ${plural(n, 'файл', 'файла', 'файлов')}`;
  // Даём браузеру нарисовать плейсхолдер, прежде чем начать тяжёлую работу
  const paint = () => new Promise((r) => requestAnimationFrame(() => setTimeout(r, 0)));
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  const ICON = {
    file: '<svg aria-hidden="true" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/></svg>',
    upload: '<svg aria-hidden="true" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M12 15V4"/><path d="M7 9l5-5 5 5"/><path d="M5 15v3a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-3"/></svg>',
    x: '<svg aria-hidden="true" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><path d="M6 6l12 12M18 6L6 18"/></svg>',
    info: '<svg aria-hidden="true" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 8v5M12 16.5v.5"/></svg>',
    table: '<svg aria-hidden="true" width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18"/></svg>',
    columns: '<svg aria-hidden="true" width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 3v18M15 3v18"/></svg>',
    eye: '<svg aria-hidden="true" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M2 12s3.6-7 10-7 10 7 10 7-3.6 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/></svg>',
    down: '<svg aria-hidden="true" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 4v11"/><path d="M7 10l5 5 5-5"/><path d="M5 20h14"/></svg>',
  };

  /* ---------- Отрисовка: левая панель ---------- */

  function renderFileZone() {
    const z = $('fileZone');
    if (!st.file) {
      z.innerHTML = `
        <div class="drop" id="drop">
          <span style="color: var(--accent)">${ICON.upload}</span>
          <span class="drop__title">Перетащите Excel-файл сюда</span>
          <span class="drop__hint">.xlsx, .xls, .xlsm, .xlsb, .ods, .csv</span>
          <button type="button" class="btn btn--primary" data-act="pick">Выбрать файл</button>
        </div>`;
      return;
    }
    let meta = '';
    if (st.phase === 'reading') meta = 'Читаем файл…';
    else if (st.phase === 'analyzing') meta = 'Разбираем лист…';
    else if (st.analysis) {
      meta = `${rowsWord(st.analysis.dataRows.length)} · ${fmt(st.sheets.length)} ${plural(st.sheets.length, 'лист', 'листа', 'листов')}`;
    } else meta = 'Перетащите другой файл сюда';
    z.innerHTML = `
      <div class="file" id="drop">
        ${ICON.file}
        <div class="file__info">
          <span class="file__name" title="${esc(st.file.name)}">${esc(st.file.name)}</span>
          <span class="file__meta">${esc(meta)}</span>
        </div>
        <button type="button" class="icon-btn" data-act="clear" aria-label="Убрать файл">${ICON.x}</button>
      </div>`;
  }

  function renderSheets() {
    const block = $('sheetBlock');
    block.hidden = !st.file || st.phase === 'error' && !st.sheets.length;
    if (block.hidden) return;
    if (st.phase === 'reading') {
      $('sheetList').innerHTML = '<div class="sk sk--chip"></div><div class="sk sk--chip"></div><div class="sk sk--chip"></div>';
      return;
    }
    $('sheetList').innerHTML = st.sheets
      .map((s) => `<button type="button" class="chip${s.hidden ? ' chip--hidden' : ''}" data-sheet="${s.id}" aria-pressed="${s.id === st.sheetId}"${s.hidden ? ' title="Скрытый лист"' : ''}>${esc(s.name)}</button>`)
      .join('');
  }

  function colHint(c) {
    if (!c.distinct) return 'пусто';
    if (c.distinct <= 4) return c.sample.join(' · ');
    return `${fmt(c.distinct)} разных`;
  }

  function renderCols() {
    const block = $('colBlock');
    block.hidden = !st.file || (st.phase === 'error' && !st.sheets.length);
    if (block.hidden) return;
    if (st.phase === 'reading' || st.phase === 'analyzing') {
      $('colList').innerHTML = Array.from({ length: 5 }, () => '<div class="sk sk--col"></div>').join('');
      return;
    }
    if (st.analysisError) {
      $('colList').innerHTML = `<p class="muted" style="margin:0">${esc(st.analysisError)}</p>`;
      return;
    }
    if (!st.analysis) { $('colList').innerHTML = ''; return; }
    $('colList').innerHTML = st.analysis.columns
      .map((c) => {
        const hint = colHint(c);
        return `<button type="button" class="col" data-col="${c.index}" aria-pressed="${c.index === st.colIndex}">
          <span class="col__letter">${esc(c.letter)}</span>
          <span class="col__name">${esc(c.name)}</span>
          <span class="col__hint" title="${esc(hint)}">${esc(hint)}</span>
        </button>`;
      })
      .join('');
  }

  /* ---------- Отрисовка: правая часть ---------- */

  function emptyState(icon, title, text) {
    return `<div class="empty">${icon}<h1 class="empty__title">${esc(title)}</h1><p>${esc(text)}</p></div>`;
  }

  function skeletonTable(n) {
    const rows = Array.from({ length: n }, () => `
      <div class="trow">
        <span><span class="sk" style="width:18px;height:18px;border-radius:4px"></span></span>
        <span><span class="sk sk--line" style="display:block;width:60%"></span></span>
        <span class="tnum"><span class="sk sk--line" style="display:inline-block;width:40px"></span></span>
        <span class="tact"><span class="sk" style="width:180px;height:40px"></span></span>
      </div>`).join('');
    return `<div class="table" aria-hidden="true"><div class="trow trow--head"><span></span><span>Файл</span><span class="tnum">Строк</span><span></span></div>${rows}</div>`;
  }

  function colName() {
    const c = st.result && st.result.column;
    return c ? c.name : '';
  }

  function notesHtml() {
    const r = st.result;
    const a = st.analysis;
    const items = [];
    if (st.warning) items.push(esc(st.warning));
    if (r.merged.length) {
      const parts = r.merged.slice(0, 5).map((m) => m.variants.map((v) => `«${esc(v)}»`).join(' и '));
      let t = `Склеены похожие значения: ${parts.join('; ')}`;
      if (r.merged.length > 5) t += ` и ещё ${r.merged.length - 5}`;
      items.push(t + '.');
    }
    const emptyG = r.groups.find((g) => g.isEmpty);
    if (emptyG) {
      const n = emptyG.rows.length;
      items.push(`Строк без значения в колонке «${esc(colName())}»: ${fmt(n)}. Они собраны в файл «${esc(emptyG.fileName)}».`);
    }
    if (a.skippedHeaders) items.push(`Повтор шапки внутри таблицы пропущен (${fmt(a.skippedHeaders)} шт.).`);
    if (a.skippedEmpty) items.push(`Пустые строки пропущены (${fmt(a.skippedEmpty)} шт.).`);
    if (r.changedName) items.push('Символы, которые нельзя использовать в именах файлов (/ \\ : * ? " &lt; &gt; |), заменены на «_».');
    if (a.hasFormulas) items.push('Формулы заменены готовыми значениями, чтобы при переносе строк ничего не сломалось.');
    if (!items.length) return '';
    return `<div class="note">${ICON.info}<ul>${items.map((i) => `<li>${i}</li>`).join('')}</ul></div>`;
  }

  function renderMain() {
    const m = $('main');
    const p = st.phase;

    if (p === 'idle') {
      m.innerHTML = emptyState(ICON.table, 'Загрузите Excel-файл', 'Перетащите файл в окно или выберите его слева. Потом выберите колонку — и строки разложатся по отдельным файлам.');
      return;
    }
    if (p === 'error') {
      m.innerHTML = `
        <div class="head"><h1 class="head__title">Не получилось</h1></div>
        <div class="note note--error">${ICON.info}<span>${esc(st.error || 'Неизвестная ошибка.')}</span></div>
        <div><button type="button" class="btn" data-act="pick">Выбрать другой файл</button></div>`;
      return;
    }
    if (p === 'reading' || p === 'analyzing') {
      m.innerHTML = `
        <div class="head"><h1 class="head__title">${p === 'reading' ? 'Читаем файл…' : 'Разбираем лист…'}</h1></div>
        <p class="muted" style="margin:0">Большие файлы читаются дольше — это нормально.</p>
        ${skeletonTable(4)}`;
      return;
    }
    if (p === 'pick') {
      if (st.analysisError) {
        m.innerHTML = emptyState(ICON.table, 'На этом листе нечего делить', 'Выберите другой лист слева.');
      } else {
        m.innerHTML = emptyState(ICON.columns, 'Выберите колонку слева', 'Строки разложим по её значениям: каждое значение — отдельный файл. Шапка попадёт в каждый файл.');
      }
      return;
    }
    if (p === 'many') {
      const n = st.result.groups.length;
      m.innerHTML = `
        <div class="head"><h1 class="head__title">Получится ${filesWord(n)}</h1></div>
        <div class="note">${ICON.info}<span>В колонке «${esc(colName())}» почти все значения разные. Обычно делят по колонке, где значения повторяются: курс, группа, отдел, город.</span></div>
        <div style="display:flex;gap:12px;flex-wrap:wrap">
          <button type="button" class="btn btn--primary" data-act="force">Всё равно разделить</button>
          <button type="button" class="btn" data-act="backToCols">Выбрать другую колонку</button>
        </div>`;
      return;
    }
    if (p === 'splitting') {
      const { done, total } = st.progress;
      m.innerHTML = `
        <div class="head"><h1 class="head__title">Делим на файлы…</h1><span class="muted" id="progText">${fmt(done)} из ${fmt(total)}</span></div>
        <div class="progress"><div class="bar" role="progressbar" aria-label="Готово файлов" aria-valuemin="0" aria-valuemax="${total}" aria-valuenow="${done}" id="progBar"><div class="bar__fill" id="progFill" style="width:${total ? (done / total) * 100 : 0}%"></div></div></div>
        ${skeletonTable(Math.min(total, 6))}`;
      return;
    }
    if (p === 'done') {
      const r = st.result;
      const total = r.groups.reduce((s, g) => s + g.rows.length, 0);
      const all = st.checked.size === r.groups.length;
      const rows = r.groups.map((g, i) => `
        <div class="trow">
          <span><input type="checkbox" data-check="${i}" aria-label="Отметить ${esc(g.fileName)}"${st.checked.has(g.fileName) ? ' checked' : ''}></span>
          <span class="tname${g.isEmpty ? ' tname--empty' : ''}">${esc(g.fileName)}</span>
          <span class="tnum">${fmt(g.rows.length)}</span>
          <span class="tact">
            <button type="button" class="btn btn--sm" data-preview="${i}">${ICON.eye}Просмотр</button>
            <button type="button" class="btn btn--sm btn--link" data-download="${i}">${ICON.down}Скачать</button>
          </span>
        </div>`).join('');
      m.innerHTML = `
        <div class="head">
          <h1 class="head__title">Результат</h1>
          <span class="muted">${filesWord(r.groups.length)} · ${rowsWord(total)} · по колонке «${esc(colName())}»</span>
        </div>
        <div class="table">
          <div class="trow trow--head">
            <span><input type="checkbox" data-act="checkAll" aria-label="Отметить все"${all ? ' checked' : ''}></span>
            <span>Файл</span><span class="tnum">Строк</span><span></span>
          </div>
          ${rows}
        </div>
        ${notesHtml()}
        <div class="actions">
          <button type="button" class="btn btn--lg" data-act="dlChecked"${st.checked.size && !st.busy ? '' : ' disabled'}>Скачать отмеченные по одному (${fmt(st.checked.size)})</button>
          <button type="button" class="btn btn--lg btn--primary" data-act="dlZip"${st.busy ? ' disabled' : ''}>Скачать всё в ZIP</button>
        </div>`;
      const head = m.querySelector('[data-act="checkAll"]');
      if (head) head.indeterminate = st.checked.size > 0 && !all;
    }
  }

  function renderAll() {
    renderFileZone();
    renderSheets();
    renderCols();
    renderMain();
  }

  function updateProgress() {
    const { done, total } = st.progress;
    const t = $('progText');
    const f = $('progFill');
    const b = $('progBar');
    if (t) t.textContent = `${fmt(done)} из ${fmt(total)}`;
    if (f) f.style.width = `${total ? (done / total) * 100 : 0}%`;
    if (b) b.setAttribute('aria-valuenow', String(done));
  }

  /* ---------- Работа ---------- */

  function fail(msg) {
    st.phase = 'error';
    st.error = msg;
    renderAll();
  }

  function reset() {
    st.job++;
    Object.assign(st, {
      phase: 'idle', file: null, workbook: null, sheets: [], sheetId: null, warning: null,
      analysis: null, analysisError: null, colIndex: null, result: null, error: null, busy: false,
    });
    st.files = new Map();
    st.checked = new Set();
    renderAll();
  }

  async function handleFile(file) {
    if (!file) return;
    if (!window.ExcelJS || !window.XLSX || !window.JSZip || !S) {
      st.file = file;
      fail('Не загрузились нужные части программы. Обновите страницу.');
      return;
    }
    const job = ++st.job;
    Object.assign(st, {
      file, phase: 'reading', workbook: null, sheets: [], sheetId: null, warning: null,
      analysis: null, analysisError: null, colIndex: null, result: null, error: null,
    });
    st.files = new Map();
    st.checked = new Set();
    renderAll();

    const ext = S.extOf(file.name);
    if (!S.ACCEPT.includes(ext)) {
      fail(`Файл «${file.name}» не похож на таблицу Excel. Подойдут .xlsx, .xls, .xlsm, .xlsb, .ods или .csv.`);
      return;
    }
    await paint();
    try {
      const buf = await file.arrayBuffer();
      const res = await S.loadWorkbook(buf, file.name);
      if (job !== st.job) return;
      st.workbook = res.workbook;
      st.sheets = res.sheets;
      st.warning = res.warning;
      const first = res.sheets.find((s) => !s.hidden) || res.sheets[0];
      await selectSheet(first.id);
    } catch (e) {
      if (job !== st.job) return;
      console.error(e);
      const enc = /password|encrypt|CFB/i.test(String(e && e.message));
      fail(enc
        ? 'Файл защищён паролем. Снимите пароль в Excel и загрузите файл ещё раз.'
        : 'Не удалось прочитать файл. Проверьте, что он открывается в Excel, и попробуйте ещё раз.');
    }
  }

  async function selectSheet(id) {
    const job = ++st.job;
    Object.assign(st, { sheetId: id, phase: 'analyzing', analysis: null, analysisError: null, colIndex: null, result: null });
    st.files = new Map();
    renderAll();
    await paint();
    try {
      const ws = st.workbook.getWorksheet(id);
      const a = await S.analyzeSheet(ws);
      if (job !== st.job) return;
      st.analysis = a;
      if (!a.dataRows.length || !a.columns.length) st.analysisError = 'На листе есть только шапка, строк с данными нет.';
    } catch (e) {
      if (job !== st.job) return;
      st.analysisError = e && e.message ? e.message : 'Не удалось разобрать лист.';
    }
    st.phase = 'pick';
    renderAll();
  }

  async function selectColumn(index, force) {
    if (!st.analysis) return;
    const job = ++st.job;
    st.colIndex = index;
    st.result = S.groupBy(st.analysis, index, st.file.name);
    st.files = new Map();
    renderCols();
    if (!force && st.result.groups.length >= MANY) {
      st.phase = 'many';
      renderAll();
      return;
    }
    await runSplit(job);
  }

  async function runSplit(job) {
    const groups = st.result.groups;
    st.phase = 'splitting';
    st.progress = { done: 0, total: groups.length };
    renderAll();
    await paint();
    try {
      for (const g of groups) {
        const bytes = await S.buildGroupFile(st.analysis, g.rows);
        if (job !== st.job) return;
        st.files.set(g.fileName, bytes);
        st.progress.done++;
        updateProgress();
        await paint();
      }
    } catch (e) {
      if (job !== st.job) return;
      console.error(e);
      fail('Не получилось собрать файлы. Попробуйте ещё раз или выберите другую колонку.');
      return;
    }
    st.checked = new Set(groups.map((g) => g.fileName));
    st.phase = 'done';
    renderAll();
  }

  /* ---------- Скачивание ---------- */

  function saveBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 15000);
  }

  function downloadOne(g) {
    const bytes = st.files.get(g.fileName);
    if (bytes) saveBlob(new Blob([bytes], { type: XLSX_MIME }), g.fileName);
  }

  async function downloadChecked() {
    const list = st.result.groups.filter((g) => st.checked.has(g.fileName));
    st.busy = true;
    renderMain();
    for (const g of list) {
      downloadOne(g);
      await sleep(400); // браузеры не любят много скачиваний подряд без паузы
    }
    st.busy = false;
    renderMain();
  }

  async function downloadZip() {
    st.busy = true;
    renderMain();
    const b = document.querySelector('[data-act="dlZip"]');
    if (b) b.textContent = 'Собираем архив…';
    try {
      const zip = new window.JSZip();
      for (const g of st.result.groups) zip.file(g.fileName, st.files.get(g.fileName));
      const blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' });
      const col = colName().replace(/[\\/:*?"<>|]/g, '_');
      saveBlob(blob, `${col}_${S.baseName(st.file.name).replace(/[\\/:*?"<>|]/g, '_')}.zip`);
    } finally {
      st.busy = false;
      renderMain();
    }
  }

  /* ---------- Предпросмотр ---------- */

  let previewGroup = null;
  function openPreview(g) {
    previewGroup = g;
    const p = S.preview(st.analysis, g.rows, PREVIEW_ROWS);
    const keyPos = st.analysis.columns.findIndex((c) => c.index === st.colIndex);
    $('previewTitle').textContent = g.fileName;
    $('previewBody').innerHTML = `
      <table class="ptable">
        <thead><tr><th class="num">№</th>${p.header.map((h, i) => `<th${i === keyPos ? ' class="is-key"' : ''}>${esc(h)}</th>`).join('')}</tr></thead>
        <tbody>${p.rows.map((r, n) => `<tr><td class="num">${n + 1}</td>${r.map((v, i) => `<td${i === keyPos ? ' class="is-key"' : ''} title="${esc(v)}">${esc(v)}</td>`).join('')}</tr>`).join('')}</tbody>
      </table>`;
    $('previewNote').textContent = p.total > p.rows.length
      ? `Показаны первые ${fmt(p.rows.length)} из ${rowsWord(p.total)}. Шапка и оформление в файле — как в исходнике.`
      : `Все ${rowsWord(p.total)}. Шапка и оформление в файле — как в исходнике.`;
    $('previewDlg').showModal();
  }

  /* ---------- События ---------- */

  const input = $('fileInput');
  input.accept = S ? S.ACCEPT.map((e) => '.' + e).join(',') : '';
  input.addEventListener('change', () => {
    const f = input.files && input.files[0];
    input.value = '';
    handleFile(f);
  });

  document.addEventListener('click', (e) => {
    const t = e.target.closest('button, input, #drop');
    if (!t) return;
    const act = t.dataset.act;
    if (act === 'pick' || (t.id === 'drop' && !st.file)) { input.click(); return; }
    if (act === 'clear') { reset(); return; }
    if (t.dataset.sheet) { const id = Number(t.dataset.sheet); if (id !== st.sheetId) selectSheet(id); return; }
    if (t.dataset.col) { const i = Number(t.dataset.col); if (i !== st.colIndex || st.phase === 'error') selectColumn(i); return; }
    if (act === 'force') { selectColumn(st.colIndex, true); return; }
    if (act === 'backToCols') { const b = document.querySelector('.col'); if (b) b.focus(); return; }
    if (act === 'dlChecked') { downloadChecked(); return; }
    if (act === 'dlZip') { downloadZip(); return; }
    if (t.dataset.download != null) { downloadOne(st.result.groups[Number(t.dataset.download)]); return; }
    if (t.dataset.preview != null) { openPreview(st.result.groups[Number(t.dataset.preview)]); }
  });

  document.addEventListener('change', (e) => {
    const t = e.target;
    if (t.dataset.act === 'checkAll') {
      st.checked = t.checked ? new Set(st.result.groups.map((g) => g.fileName)) : new Set();
      renderMain();
    } else if (t.dataset.check != null) {
      const g = st.result.groups[Number(t.dataset.check)];
      if (t.checked) st.checked.add(g.fileName); else st.checked.delete(g.fileName);
      renderMain();
      const again = document.querySelector(`[data-check="${t.dataset.check}"]`);
      if (again) again.focus();
    }
  });

  $('previewClose').addEventListener('click', () => $('previewDlg').close());
  $('previewDownload').addEventListener('click', () => { if (previewGroup) downloadOne(previewGroup); });
  $('previewDlg').addEventListener('click', (e) => { if (e.target === e.currentTarget) e.currentTarget.close(); });

  // Перетаскивание файла в любое место окна
  let dragDepth = 0;
  const hl = (on) => { const d = $('drop'); if (d) d.classList.toggle('is-over', on); };
  window.addEventListener('dragenter', (e) => { e.preventDefault(); dragDepth++; hl(true); });
  window.addEventListener('dragover', (e) => { e.preventDefault(); });
  window.addEventListener('dragleave', () => { dragDepth = Math.max(0, dragDepth - 1); if (!dragDepth) hl(false); });
  window.addEventListener('drop', (e) => {
    e.preventDefault();
    dragDepth = 0;
    hl(false);
    const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) handleFile(f);
  });

  // Для проверки из консоли
  window.__splitterApp = { handleFile, state: st };

  renderAll();
})();
