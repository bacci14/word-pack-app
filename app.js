/* Word Pack — reemplazos masivos en .docx y exportación a PDF.
   Todo corre en el navegador: los archivos nunca se suben a ningún servidor. */
(function () {
  'use strict';

  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const XML_NS = 'http://www.w3.org/XML/1998/namespace';
  const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  const PART_RE = /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/;
  const HEADER_RE = /^word\/header\d*\.xml$/;

  const MONTHS = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto',
    'septiembre', 'octubre', 'noviembre', 'diciembre'];
  const MONTH_ALT = MONTHS.join('|') + '|setiembre';
  const DATE_RE = new RegExp(
    '\\b\\d{1,2}[\\/.\\-]\\d{1,2}[\\/.\\-]\\d{2,4}\\b' +
    '|\\b\\d{1,2}\\s+de\\s+(?:' + MONTH_ALT + ')(?:\\s+(?:de|del)\\s+\\d{4})?\\b' +
    '|\\b(?:' + MONTH_ALT + ')\\s+(?:de\\s+)?\\d{4}\\b', 'gi');

  // ───────────────────────── XML de Word ─────────────────────────

  function isW(el, name) { return el && el.nodeType === 1 && el.namespaceURI === W_NS && el.localName === name; }

  function ownParagraph(node) {
    let n = node.parentNode;
    while (n && !isW(n, 'p')) n = n.parentNode;
    return n;
  }

  // Los <w:t> que pertenecen a este párrafo (no a un cuadro de texto anidado).
  function textNodesOf(p) {
    return Array.from(p.getElementsByTagNameNS(W_NS, 't')).filter(t => ownParagraph(t) === p);
  }

  function paraText(p) { return textNodesOf(p).map(t => t.textContent).join(''); }

  function setT(t, value) {
    t.textContent = value;
    t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  }

  // Word suele partir una misma frase en varios "runs" (p. ej. " 22/10" + "/2026").
  // Buscamos sobre el texto completo del párrafo y repartimos el resultado entre los runs,
  // así se conserva el formato del primer run de cada coincidencia.
  function replaceInParagraph(p, find, repl, caseSensitive) {
    if (!find) return 0;
    const needle = caseSensitive ? find : find.toLowerCase();
    let count = 0, from = 0, guard = 0;
    while (guard++ < 10000) {
      const nodes = textNodesOf(p);
      if (!nodes.length) break;
      const texts = nodes.map(n => n.textContent);
      const full = texts.join('');
      const hay = caseSensitive ? full : full.toLowerCase();
      const idx = hay.indexOf(needle, from);
      if (idx < 0) break;
      const end = idx + find.length;
      let pos = 0, si = -1, so = 0, ei = -1, eo = 0;
      for (let i = 0; i < texts.length; i++) {
        const len = texts[i].length;
        if (si < 0 && idx < pos + len) { si = i; so = idx - pos; }
        if (si >= 0 && end <= pos + len) { ei = i; eo = end - pos; break; }
        pos += len;
      }
      if (si < 0 || ei < 0) break;
      if (si === ei) {
        setT(nodes[si], texts[si].slice(0, so) + repl + texts[si].slice(eo));
      } else {
        setT(nodes[si], texts[si].slice(0, so) + repl);
        for (let k = si + 1; k < ei; k++) setT(nodes[k], '');
        setT(nodes[ei], texts[ei].slice(eo));
      }
      count++;
      from = idx + repl.length;
    }
    return count;
  }

  function allParagraphs(root) { return Array.from(root.getElementsByTagNameNS(W_NS, 'p')); }

  // Portada = el contenido del cuerpo hasta el primer salto de página o de sección.
  function coverParagraphs(doc) {
    const body = doc.getElementsByTagNameNS(W_NS, 'body')[0];
    if (!body) return [];
    const kids = Array.from(body.children);
    const scope = [];
    let closed = false;
    for (let i = 0; i < kids.length && i < 80; i++) {
      const el = kids[i];
      if (isW(el, 'sectPr')) { closed = true; break; }
      if (i > 0 && startsNewPage(el)) { closed = true; break; }
      scope.push(el);
      if (endsPage(el)) { closed = true; break; }
    }
    const used = closed ? scope : scope.slice(0, 15);
    const out = [];
    used.forEach(el => { if (isW(el, 'p')) out.push(el); out.push(...allParagraphs(el)); });
    return out;
  }

  function startsNewPage(el) {
    if (!isW(el, 'p')) return false;
    const pPr = Array.from(el.children).find(c => isW(c, 'pPr'));
    if (pPr && Array.from(pPr.children).some(c => isW(c, 'pageBreakBefore'))) return true;
    // Marca que deja Word donde empezó una página nueva, antes de cualquier texto.
    const walker = el.ownerDocument.createTreeWalker(el, NodeFilter.SHOW_ELEMENT);
    let n;
    while ((n = walker.nextNode())) {
      if (isW(n, 'lastRenderedPageBreak')) return true;
      if (isW(n, 't') && n.textContent.trim()) return false;
    }
    return false;
  }

  function endsPage(el) {
    const brs = el.getElementsByTagNameNS(W_NS, 'br');
    for (const br of brs) if (br.getAttributeNS(W_NS, 'type') === 'page') return true;
    if (isW(el, 'p')) {
      const pPr = Array.from(el.children).find(c => isW(c, 'pPr'));
      const sect = pPr && Array.from(pPr.children).find(c => isW(c, 'sectPr'));
      if (sect) {
        // Un salto de sección "continuo" no empieza una página nueva.
        const type = Array.from(sect.children).find(c => isW(c, 'type'));
        if (!type || type.getAttributeNS(W_NS, 'val') !== 'continuous') return true;
      }
    }
    return false;
  }

  function uniqueLines(paragraphs) {
    const seen = new Set(), out = [];
    paragraphs.forEach(p => {
      const t = paraText(p).replace(/\s+/g, ' ').trim();
      if (t && !seen.has(t)) { seen.add(t); out.push(t); }
    });
    return out;
  }

  function findDates(lines) {
    const seen = new Set(), out = [];
    lines.forEach(l => (l.match(DATE_RE) || []).forEach(d => { if (!seen.has(d)) { seen.add(d); out.push(d); } }));
    return out;
  }

  // ───────────────────────── Fechas ─────────────────────────

  // Escribe la fecha nueva con el mismo formato que la fecha vieja.
  function formatLike(old, date) {
    const d = date.getDate(), m = date.getMonth(), y = date.getFullYear();
    const num = old.match(/^(\d{1,2})([\/.\-])(\d{1,2})\2(\d{2,4})$/);
    if (num) {
      const pad = (v, len) => String(v).padStart(len, '0');
      const yy = num[4].length === 2 ? String(y).slice(-2) : String(y);
      return pad(d, num[1].length) + num[2] + pad(m + 1, num[3].length) + num[2] + yy;
    }
    const lower = old.toLowerCase();
    const mi = MONTHS.findIndex(name => lower.includes(name)) >= 0 || lower.includes('setiembre');
    if (mi) {
      const oldMonth = (old.match(new RegExp(MONTH_ALT, 'i')) || [''])[0];
      let month = MONTHS[m];
      if (oldMonth && oldMonth[0] === oldMonth[0].toUpperCase()) month = month[0].toUpperCase() + month.slice(1);
      if (oldMonth === oldMonth.toUpperCase() && oldMonth.length > 1) month = month.toUpperCase();
      if (/^\d/.test(old)) {
        const dayLen = old.match(/^\d+/)[0].length;
        const day = String(d).padStart(dayLen, '0');
        const yearPart = old.match(/\s+(de|del)\s+\d{4}$/i);
        return day + ' de ' + month + (yearPart ? ' ' + yearPart[1] + ' ' + y : '');
      }
      return month + (/\s+de\s+\d{4}$/i.test(old) ? ' de ' : ' ') + y;
    }
    return old;
  }

  // ───────────────────────── Estado ─────────────────────────

  const state = { docs: [], nextId: 1 };
  const $ = (sel, root = document) => root.querySelector(sel);
  const el = (tag, attrs = {}, ...kids) => {
    const e = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === 'class') e.className = v;
      else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
      else if (v === true) e.setAttribute(k, '');
      else if (v !== false && v != null) e.setAttribute(k, v);
    }
    kids.flat().forEach(c => e.append(c instanceof Node ? c : document.createTextNode(String(c))));
    return e;
  };

  async function analyze(file) {
    const buf = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(buf);
    const parser = new DOMParser();
    const docXml = await zip.file('word/document.xml').async('string');
    const doc = parser.parseFromString(docXml, 'application/xml');
    const coverLines = uniqueLines(coverParagraphs(doc));
    const headerParas = [];
    for (const name of Object.keys(zip.files).filter(n => HEADER_RE.test(n)).sort()) {
      const x = parser.parseFromString(await zip.file(name).async('string'), 'application/xml');
      headerParas.push(...allParagraphs(x));
    }
    const headerLines = uniqueLines(headerParas).filter(l => !/^\[%.*%\]$/.test(l));
    const all = coverLines.concat(headerLines);
    const dates = findDates(all);
    const otherLines = Array.from(new Set(all)).filter(l => l.replace(DATE_RE, '').trim().length > 0);
    return {
      id: state.nextId++, name: file.name, buf, coverLines, headerLines, dates, otherLines,
      editScope: false, dateEdits: {}, lineEdits: {}, result: null
    };
  }

  async function processDoc(d, pairs, caseSensitive) {
    const zip = await JSZip.loadAsync(d.buf);
    const parser = new DOMParser(), ser = new XMLSerializer();
    const stats = {};
    const names = Object.keys(zip.files).filter(n => PART_RE.test(n));
    for (const name of names) {
      const xml = parser.parseFromString(await zip.file(name).async('string'), 'application/xml');
      let changed = false;

      // 1) Portada y encabezado (solo si se tildó para este documento)
      if (d.editScope) {
        let scoped = [];
        if (name === 'word/document.xml') scoped = coverParagraphs(xml);
        else if (HEADER_RE.test(name)) scoped = allParagraphs(xml);
        const lineEdits = Object.entries(d.lineEdits).filter(([a, b]) => b !== '' && a !== b);
        const dateEdits = Object.entries(d.dateEdits).filter(([a, b]) => b !== '' && a !== b);
        scoped.forEach(p => {
          lineEdits.forEach(([a, b]) => { if (replaceInParagraph(p, a, b, true)) changed = true; });
          dateEdits.forEach(([a, b]) => { if (replaceInParagraph(p, a, b, true)) changed = true; });
        });
      }

      // 2) Buscar y reemplazar en todo el documento
      if (pairs.length) {
        const paras = allParagraphs(xml);
        pairs.forEach(({ find, repl }) => {
          let n = 0;
          paras.forEach(p => { n += replaceInParagraph(p, find, repl, caseSensitive); });
          if (n) { stats[find] = (stats[find] || 0) + n; changed = true; }
        });
      }
      if (changed) zip.file(name, ser.serializeToString(xml));
    }
    const blob = await zip.generateAsync({ type: 'blob', mimeType: DOCX_MIME, compression: 'DEFLATE' });
    return { blob, stats };
  }

  // ─────────────── Scripts que convierten a PDF con Microsoft Word ───────────────
  // Van dentro del .zip, al lado de los Word. Doble clic → un PDF por cada .docx de esa carpeta.

  const MAC_SCRIPT = [
    '#!/bin/zsh',
    '# Convierte a PDF, con Microsoft Word, todos los .docx de esta carpeta.',
    'cd "$(dirname "$0")" || exit 1',
    'echo "Convirtiendo a PDF con Microsoft Word…"',
    'osascript - "$(pwd)" <<\'APPLESCRIPT\'',
    'on run argv',
    '  set d to item 1 of argv',
    '  set fileList to paragraphs of (do shell script "cd " & quoted form of d & " && ls -1 | grep -i \'\\\\.docx$\' | grep -v \'^~\' || true")',
    '  set n to 0',
    '  repeat with f in fileList',
    '    set f to f as text',
    '    if f is not "" then',
    '      set src to d & "/" & f',
    '      set dst to d & "/" & (text 1 thru -6 of f) & ".pdf"',
    '      tell application "Microsoft Word"',
    '        open (POSIX file src)',
    '        set theDoc to active document',
    '        save as theDoc file name dst file format format PDF',
    '        close theDoc saving no',
    '      end tell',
    '      set n to n + 1',
    '    end if',
    '  end repeat',
    '  return "Listo: " & n & " PDF creados."',
    'end run',
    'APPLESCRIPT',
    'echo "Podés cerrar esta ventana."',
    ''
  ].join('\n');

  const WIN_SCRIPT = [
    '@echo off',
    'rem Convierte a PDF, con Microsoft Word, todos los .docx de esta carpeta.',
    'echo Convirtiendo a PDF con Microsoft Word...',
    'powershell -NoProfile -ExecutionPolicy Bypass -Command "$w = New-Object -ComObject Word.Application; $w.Visible = $false; Get-ChildItem -LiteralPath \'%~dp0.\' -Filter *.docx | Where-Object { $_.Name -notlike \'~*\' } | ForEach-Object { $d = $w.Documents.Open($_.FullName, $false, $true); $d.ExportAsFixedFormat(($_.FullName -replace \'\\.docx$\', \'.pdf\'), 17); $d.Close(0); Write-Host (\'  OK  \' + $_.Name) }; $w.Quit()"',
    'echo Listo.',
    'pause',
    ''
  ].join('\r\n');

  // ───────────────────────── Interfaz ─────────────────────────

  const STORE_KEY = 'wordpack.pairs.v1';

  function loadPairs() {
    try { const v = JSON.parse(localStorage.getItem(STORE_KEY) || '[]'); return Array.isArray(v) && v.length ? v : [{ find: '', repl: '' }]; }
    catch (e) { return [{ find: '', repl: '' }]; }
  }
  function savePairs() {
    try { localStorage.setItem(STORE_KEY, JSON.stringify(readPairs(true))); } catch (e) { /* sin almacenamiento */ }
  }
  function readPairs(includeEmpty) {
    return Array.from(document.querySelectorAll('#pairs .pair')).map(r => ({
      find: $('.find', r).value, repl: $('.repl', r).value
    })).filter(p => includeEmpty || p.find !== '');
  }

  function addPairRow(p = { find: '', repl: '' }) {
    const row = el('div', { class: 'pair' },
      el('input', { class: 'find', type: 'text', placeholder: 'Buscar…', value: p.find, oninput: savePairs, 'aria-label': 'Texto a buscar' }),
      el('span', { class: 'arrow', 'aria-hidden': 'true' }, '→'),
      el('input', { class: 'repl', type: 'text', placeholder: 'Reemplazar por…', value: p.repl, oninput: savePairs, 'aria-label': 'Reemplazar por' }),
      el('button', { class: 'icon-btn', type: 'button', title: 'Quitar', 'aria-label': 'Quitar fila', onclick: () => { row.remove(); if (!$('#pairs .pair')) addPairRow(); savePairs(); } }, '×'));
    $('#pairs').append(row);
    return row;
  }

  async function addFiles(fileList) {
    const files = Array.from(fileList).filter(f => /\.docx$/i.test(f.name) && !f.name.startsWith('~$'));
    const skipped = fileList.length - files.length;
    setStatus(files.length ? 'Leyendo ' + files.length + ' documento(s)…' : '');
    for (const f of files) {
      try { state.docs.push(await analyze(f)); }
      catch (e) { console.error(e); toast('No pude leer «' + f.name + '». ¿Es un .docx válido?'); }
    }
    if (skipped) toast(skipped + ' archivo(s) ignorado(s): solo se aceptan .docx');
    setStatus('');
    render();
  }

  function render() {
    const list = $('#file-list');
    list.innerHTML = '';
    state.docs.forEach(d => list.append(el('li', {},
      el('span', { class: 'fname' }, d.name),
      el('button', { class: 'icon-btn', type: 'button', title: 'Quitar', 'aria-label': 'Quitar ' + d.name, onclick: () => { state.docs = state.docs.filter(x => x !== d); render(); } }, '×'))));
    $('#file-count').textContent = state.docs.length ? state.docs.length + ' documento(s) cargado(s)' : '';
    document.body.classList.toggle('has-docs', state.docs.length > 0);
    renderScope();
    $('#results').innerHTML = '';
    $('#downloads').hidden = true;
  }

  function renderScope() {
    const box = $('#scope-list');
    box.innerHTML = '';
    state.docs.forEach((d, i) => {
      const body = el('div', { class: 'scope-body' });
      const card = el('article', { class: 'scope-card' + (d.editScope ? ' open' : '') },
        el('label', { class: 'scope-head' },
          el('input', { type: 'checkbox', checked: d.editScope, onchange: e => { d.editScope = e.target.checked; card.classList.toggle('open', d.editScope); } }),
          el('span', { class: 'num' }, String(i + 1)),
          el('span', { class: 'fname' }, d.name),
          el('span', { class: 'hint' }, '¿Cambiar portada y encabezado?')),
        body);

      if (d.dates.length) {
        body.append(el('h4', {}, 'Fechas encontradas'));
        d.dates.forEach(old => {
          const txt = el('input', { type: 'text', class: 'new-val', placeholder: 'Nueva fecha', value: d.dateEdits[old] || '', oninput: e => { d.dateEdits[old] = e.target.value; } });
          const pick = el('input', { type: 'date', class: 'picker', title: 'Elegir del calendario', 'aria-label': 'Elegir fecha del calendario', onchange: e => {
            if (!e.target.value) return;
            const [y, m, dd] = e.target.value.split('-').map(Number);
            txt.value = formatLike(old, new Date(y, m - 1, dd));
            d.dateEdits[old] = txt.value;
          } });
          body.append(el('div', { class: 'edit-row date' }, el('code', { class: 'old' }, old), el('span', { class: 'arrow' }, '→'), txt, pick));
        });
      } else {
        body.append(el('p', { class: 'muted' }, 'No encontré fechas en la portada ni en el encabezado. Podés editar los textos de abajo.'));
      }

      if (d.otherLines.length) {
        const det = el('details', { class: 'other' }, el('summary', {}, 'Otros textos de portada y encabezado (' + d.otherLines.length + ')'));
        d.otherLines.forEach(line => {
          det.append(el('div', { class: 'edit-row' }, el('code', { class: 'old' }, line), el('span', { class: 'arrow' }, '→'),
            el('input', { type: 'text', class: 'new-val', placeholder: 'Dejar igual', value: d.lineEdits[line] || '', oninput: e => { d.lineEdits[line] = e.target.value; } })));
        });
        body.append(det);
      }
      box.append(card);
    });
    $('#scope-tools').hidden = state.docs.length < 2;
  }

  function applyDateToChecked() {
    const v = $('#bulk-date').value;
    if (!v) { toast('Elegí una fecha primero.'); return; }
    const [y, m, dd] = v.split('-').map(Number);
    const date = new Date(y, m - 1, dd);
    let n = 0;
    state.docs.forEach(d => {
      if ($('#bulk-all').checked) d.editScope = true;
      if (!d.editScope) return;
      d.dates.forEach(old => { d.dateEdits[old] = formatLike(old, date); n++; });
    });
    renderScope();
    toast(n ? 'Fecha aplicada en ' + n + ' lugar(es).' : 'No hay documentos tildados con fechas.');
  }

  function safeName(name) { return name.replace(/[\\/:*?"<>|]+/g, '_'); }

  async function run() {
    if (!state.docs.length) { toast('Primero cargá al menos un documento.'); return; }
    const pairs = readPairs(false).map(p => ({ find: p.find, repl: p.repl }));
    const caseSensitive = $('#case').checked;
    const renameFiles = $('#rename').checked;
    const btn = $('#run');
    btn.disabled = true;
    const results = $('#results');
    results.innerHTML = '';
    $('#downloads').hidden = true;
    const out = new JSZip();
    const used = new Set();
    let ok = 0;

    for (let i = 0; i < state.docs.length; i++) {
      const d = state.docs[i];
      let base = d.name.replace(/\.docx$/i, '');
      if (renameFiles) pairs.forEach(({ find, repl }) => { base = base.split(find).join(repl); });
      base = safeName(base);
      let unique = base, k = 2;
      while (used.has(unique.toLowerCase())) unique = base + ' (' + k++ + ')';
      used.add(unique.toLowerCase());
      const row = el('li', {}, el('span', { class: 'fname' }, unique + '.docx'), el('span', { class: 'state' }, 'Aplicando cambios…'));
      results.append(row);
      setStatus('Documento ' + (i + 1) + ' de ' + state.docs.length + '…');
      try {
        const { blob, stats } = await processDoc(d, pairs, caseSensitive);
        out.file(unique + '.docx', blob);
        const parts = Object.entries(stats).map(([f, v]) => '«' + f + '» ×' + v);
        if (d.editScope && (Object.values(d.dateEdits).some(Boolean) || Object.values(d.lineEdits).some(Boolean))) parts.unshift('portada/encabezado ✓');
        $('.state', row).textContent = parts.length ? parts.join(' · ') : 'sin cambios';
        row.classList.add('ok');
        ok++;
      } catch (e) {
        console.error(e);
        $('.state', row).textContent = 'Error: ' + (e && e.message ? e.message : e);
        row.classList.add('err');
      }
    }

    if (ok) {
      setStatus('Armando el .zip…');
      out.file('Convertir a PDF (Mac).command', MAC_SCRIPT, { unixPermissions: 0o100755 });
      out.file('Convertir a PDF (Windows).bat', WIN_SCRIPT);
      const zipBlob = await out.generateAsync({ type: 'blob', platform: 'UNIX', compression: 'DEFLATE' });
      const stamp = new Date().toISOString().slice(0, 10);
      const dl = $('#downloads');
      dl.innerHTML = '';
      const link = downloadLink(zipBlob, 'Pack_' + stamp + '.zip', 'Descargar de nuevo (.zip)');
      dl.append(link);
      dl.hidden = false;
      link.click();
    }
    setStatus('');
    btn.disabled = false;
  }

  function downloadLink(blob, filename, label) {
    const url = URL.createObjectURL(blob);
    return el('a', { href: url, download: filename, class: 'btn' }, label);
  }

  let toastTimer;
  function toast(msg) {
    const t = $('#toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove('show'), 3500);
  }
  function setStatus(msg) { $('#status').textContent = msg; }

  function init() {
    loadPairs().forEach(addPairRow);
    $('#add-pair').addEventListener('click', () => $('.find', addPairRow()).focus());
    $('#file-input').addEventListener('change', e => { addFiles(e.target.files); e.target.value = ''; });
    const drop = $('#drop');
    ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
    drop.addEventListener('drop', e => addFiles(e.dataTransfer.files));
    $('#clear').addEventListener('click', () => { state.docs = []; render(); });
    $('#bulk-apply').addEventListener('click', applyDateToChecked);
    $('#check-all').addEventListener('click', () => { const all = state.docs.every(d => d.editScope); state.docs.forEach(d => { d.editScope = !all; }); renderScope(); });
    $('#run').addEventListener('click', run);
    render();
  }

  // Expuesto para pruebas.
  window.WordPack = { replaceInParagraph, coverParagraphs, formatLike, analyze, processDoc, state, render, MAC_SCRIPT, WIN_SCRIPT };
  document.addEventListener('DOMContentLoaded', init);
})();
