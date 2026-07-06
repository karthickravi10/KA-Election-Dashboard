/**
 * ==========================================================================
 * ELECTION DASHBOARD — APPLICATION LOGIC
 * --------------------------------------------------------------------------
 * Responsibilities:
 *   1. Load constituency data from an .xlsx/.csv file (SheetJS).
 *   2. Normalize raw rows into a predictable JS object shape.
 *   3. Power a searchable "AC Name" combobox.
 *   4. Render the selected constituency's full profile into the card grid.
 *
 * No backend / no build step: everything runs client-side in the browser.
 * ==========================================================================
 */

(() => {
  'use strict';

  /* ---------------------------------------------------------------------
   * CONFIG
   * ------------------------------------------------------------------- */

  // Fallback dataset location, only used if the embedded js/data.js
  // (window.__ELECTION_RECORDS__) isn't present for some reason.
  // Fetching this requires the page to be served over http(s):// — see README.
  const DEFAULT_DATA_URL = 'data/KA_Raw_Data.xlsx';

  // Maps a raw column header (as it appears in the source file) to the
  // internal key used throughout the app. Centralising this here means
  // a column rename in the source file only needs one edit.
  const COLUMN_MAP = {
    'Sl No': 'slNo',
    'AC No': 'acNo',
    'AC Name': 'acName',
    'Zone': 'zone',
    'PC': 'pc',
    'District': 'district',
    'Org District': 'orgDistrict',
    'Dominant Caste 1': 'caste1',
    'Dominant Caste 2': 'caste2',
    'Dominant Caste 3': 'caste3',
    'Dominant Caste 4': 'caste4',
    'Dominant Caste 5': 'caste5',
    'Winner - Name': 'winnerName',
    'Winner - Party': 'winnerParty',
    'Winner - Category': 'winnerCategory',
    'Winner - Caste': 'winnerCaste',
    'Runner Up - Name': 'runnerName',
    'Runner Up - Party': 'runnerParty',
    'Runner Up - Category': 'runnerCategory',
    'Runner Party - Caste': 'runnerCaste',
    '2024 GE Booth Gradation': 'boothGrade',
  };

  // Known party -> accent color, for the small identity dot next to
  // Winner/Runner-up party. Falls back to a neutral slate for anything
  // not in this list (independents, smaller/regional parties, etc.)
  const PARTY_COLORS = {
    'BJP': '#FF9933',
    'INC': '#00AEEF',
    'JD(S)': '#2E7D32',
    'JDS': '#2E7D32',
    'AAP': '#0B5FBA',
    'IND': '#94A3B8',
    'INDEPENDENT': '#94A3B8',
    'JD(U)': '#7A5AF8',
    'KRPP': '#B45309',
    'SDPI': '#059669',
  };

  /* ---------------------------------------------------------------------
   * STATE
   * ------------------------------------------------------------------- */

  let records = [];          // normalized data rows
  let filteredOptions = [];  // currently visible combobox options
  let activeOptionIndex = -1;

  /* ---------------------------------------------------------------------
   * DOM REFS
   * ------------------------------------------------------------------- */

  const $ = (sel) => document.querySelector(sel);

  const el = {
    recordCount: $('#record-count'),
    searchInput: $('#ac-search'),
    clearBtn: $('#combobox-clear'),
    list: $('#combobox-list'),
    fileInput: $('#file-input'),
    loadingState: $('#loading-state'),
    errorState: $('#error-state'),
    errorMessage: $('#error-message'),
    emptyState: $('#empty-state'),
    dashboard: $('#dashboard'),
    casteList: $('#caste-list'),
  };

  /* ---------------------------------------------------------------------
   * DATA LOADING
   * ------------------------------------------------------------------- */

  /**
   * Load the dashboard's starting dataset.
   *
   * Primary path: js/data.js embeds the dataset as a plain JS object
   * (window.__ELECTION_RECORDS__) via a normal <script> tag. That works
   * with zero network calls, including when index.html is opened
   * directly by double-clicking it (file://), and it has no dependency
   * on the SheetJS CDN script being reachable.
   *
   * Fallback path: if the embedded data is missing for some reason,
   * fall back to fetching + parsing the bundled .xlsx (requires the
   * page to be served over http(s):// and SheetJS to have loaded).
   */
  async function loadDefaultData() {
    showState('loading');

    if (Array.isArray(window.__ELECTION_RECORDS__) && window.__ELECTION_RECORDS__.length > 0) {
      ingestRows(window.__ELECTION_RECORDS__);
      return;
    }

    try {
      if (typeof XLSX === 'undefined') {
        throw new Error('SheetJS library not available and no embedded dataset found.');
      }
      const res = await fetch(DEFAULT_DATA_URL);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const buffer = await res.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      ingestWorkbook(workbook);
    } catch (err) {
      console.error('Failed to load default dataset:', err);
      showState(
        'error',
        'Could not load the default dataset. If you opened this file directly ' +
        '(by double-clicking it), please serve the folder with a local web server ' +
        'instead — see README.md — or use "Load a different file" below to pick ' +
        'a .xlsx/.csv manually.'
      );
    }
  }

  /**
   * Handle a user-selected file (xlsx / xls / csv) via the file input.
   * CSV is parsed with a small built-in parser (no external dependency).
   * XLSX/XLS requires the SheetJS library to have loaded successfully.
   */
  function handleFileSelect(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;

    const isCsv = /\.csv$/i.test(file.name);
    showState('loading');
    const reader = new FileReader();

    reader.onerror = () => showState('error', 'The selected file could not be read from disk.');

    if (isCsv) {
      reader.onload = (evt) => {
        try {
          const rows = parseCsv(String(evt.target.result));
          ingestRows(rows);
          resetSelection();
        } catch (err) {
          console.error('Failed to parse uploaded CSV:', err);
          showState('error', 'This CSV could not be read. Please check it matches the expected column layout.');
        }
      };
      reader.readAsText(file);
      return;
    }

    // .xlsx / .xls path — needs SheetJS.
    if (typeof XLSX === 'undefined') {
      showState(
        'error',
        'The Excel-reading library could not be loaded (likely no internet access ' +
        'or it was blocked). Please save your file as .csv and try again — CSV files ' +
        'do not require that library.'
      );
      return;
    }

    reader.onload = (evt) => {
      try {
        const workbook = XLSX.read(evt.target.result, { type: 'array' });
        ingestWorkbook(workbook);
        resetSelection();
      } catch (err) {
        console.error('Failed to parse uploaded file:', err);
        showState('error', 'This file could not be read. Please check it matches the expected column layout.');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  /**
   * Convert a parsed SheetJS workbook into an array of raw row objects,
   * then hand off to ingestRows() for normalization.
   */
  function ingestWorkbook(workbook) {
    const sheetName = workbook.SheetNames.includes('Master') ? 'Master' : workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    // defval keeps missing cells as '' instead of being omitted entirely,
    // which keeps every row shape consistent downstream.
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    ingestRows(rows);
  }

  /**
   * Normalize an array of raw row objects (keyed by original column
   * headers, from either the embedded dataset, a fetched workbook, or
   * an uploaded file) into `records`, then refresh the UI.
   */
  function ingestRows(rows) {
    records = rows
      .map(normalizeRow)
      .filter((r) => r.acName && String(r.acName).trim() !== '');

    if (records.length === 0) {
      showState('error', 'No usable rows were found. Check that the file has an "AC Name" column.');
      return;
    }

    el.recordCount.textContent = `${records.length} constituencies loaded`;
    buildOptions('');
    showState('empty');
  }

  /**
   * Minimal RFC-4180-ish CSV parser: handles quoted fields, escaped
   * quotes (""), and commas/newlines inside quotes. Returns an array
   * of row objects keyed by the header row. Good enough for the
   * well-formed exports this dashboard expects, without pulling in a
   * third-party CSV library.
   */
  function parseCsv(text) {
    const rows = [];
    let row = [];
    let field = '';
    let inQuotes = false;

    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (inQuotes) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; }
          else { inQuotes = false; }
        } else {
          field += c;
        }
      } else if (c === '"') {
        inQuotes = true;
      } else if (c === ',') {
        row.push(field); field = '';
      } else if (c === '\n' || c === '\r') {
        if (c === '\r' && text[i + 1] === '\n') i++;
        row.push(field); field = '';
        if (row.length > 1 || row[0] !== '') rows.push(row);
        row = [];
      } else {
        field += c;
      }
    }
    if (field !== '' || row.length > 0) { row.push(field); rows.push(row); }
    if (rows.length === 0) return [];

    const header = rows[0];
    return rows.slice(1).map((r) => {
      const obj = {};
      header.forEach((h, idx) => { obj[h] = r[idx] !== undefined ? r[idx] : ''; });
      return obj;
    });
  }

  /**
   * Map a raw sheet row (keyed by original column headers) to our
   * internal, code-friendly key names via COLUMN_MAP.
   */
  function normalizeRow(rawRow) {
    const out = {};
    Object.entries(COLUMN_MAP).forEach(([sourceKey, targetKey]) => {
      // Tolerate minor header drift (extra spaces) by matching loosely.
      const matchKey = Object.keys(rawRow).find(
        (k) => k.trim().toLowerCase() === sourceKey.trim().toLowerCase()
      );
      out[targetKey] = matchKey !== undefined ? String(rawRow[matchKey]).trim() : '';
    });
    return out;
  }

  /* ---------------------------------------------------------------------
   * STATE / VISIBILITY HELPERS
   * ------------------------------------------------------------------- */

  function showState(state, message) {
    el.loadingState.hidden = state !== 'loading';
    el.errorState.hidden = state !== 'error';
    el.emptyState.hidden = state !== 'empty';
    el.dashboard.hidden = state !== 'ready';
    if (state === 'error' && message) el.errorMessage.textContent = message;
  }

  function resetSelection() {
    el.searchInput.value = '';
    el.clearBtn.hidden = true;
    closeList();
  }

  /* ---------------------------------------------------------------------
   * SEARCHABLE COMBOBOX
   * ------------------------------------------------------------------- */

  function buildOptions(query) {
    const q = query.trim().toLowerCase();
    // Show every match — the dataset is small enough (a few hundred rows)
    // that a capped list would just hide constituencies the person is
    // looking for. The list itself scrolls (see .combobox-list max-height).
    filteredOptions = !q
      ? records.slice()
      : records.filter((r) => r.acName.toLowerCase().includes(q));
    renderOptions(q);
  }

  function renderOptions(query) {
    el.list.innerHTML = '';
    activeOptionIndex = -1;

    if (filteredOptions.length === 0) {
      const empty = document.createElement('li');
      empty.className = 'combobox-empty';
      empty.textContent = 'No matching constituency found.';
      el.list.appendChild(empty);
      openList();
      return;
    }

    filteredOptions.forEach((record, idx) => {
      const li = document.createElement('li');
      li.className = 'combobox-option';
      li.id = `combobox-option-${idx}`;
      li.setAttribute('role', 'option');
      li.dataset.index = String(idx);

      const nameSpan = document.createElement('span');
      nameSpan.innerHTML = highlightMatch(record.acName, query);

      const meta = document.createElement('small');
      meta.textContent = `AC ${record.acNo}${record.district ? ' · ' + record.district : ''}`;

      li.appendChild(nameSpan);
      li.appendChild(meta);

      li.addEventListener('mousedown', (e) => {
        // mousedown (not click) so it fires before input blur closes the list
        e.preventDefault();
        selectRecord(record);
      });

      el.list.appendChild(li);
    });

    openList();
  }

  function highlightMatch(text, query) {
    if (!query) return escapeHtml(text);
    const idx = text.toLowerCase().indexOf(query);
    if (idx === -1) return escapeHtml(text);
    const before = escapeHtml(text.slice(0, idx));
    const match = escapeHtml(text.slice(idx, idx + query.length));
    const after = escapeHtml(text.slice(idx + query.length));
    return `${before}<mark>${match}</mark>${after}`;
  }

  function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  function openList() {
    el.list.classList.add('is-open');
    el.searchInput.setAttribute('aria-expanded', 'true');
  }

  function closeList() {
    el.list.classList.remove('is-open');
    el.searchInput.setAttribute('aria-expanded', 'false');
  }

  function moveActiveOption(delta) {
    const options = Array.from(el.list.querySelectorAll('.combobox-option'));
    if (options.length === 0) return;
    activeOptionIndex = (activeOptionIndex + delta + options.length) % options.length;
    options.forEach((o, i) => o.classList.toggle('is-active', i === activeOptionIndex));
    options[activeOptionIndex].scrollIntoView({ block: 'nearest' });
  }

  function selectActiveOption() {
    if (activeOptionIndex === -1 || !filteredOptions[activeOptionIndex]) return;
    selectRecord(filteredOptions[activeOptionIndex]);
  }

  function selectRecord(record) {
    el.searchInput.value = record.acName;
    el.clearBtn.hidden = false;
    closeList();
    renderDashboard(record);
    showState('ready');
    // Bring the dashboard into view on small screens after selection.
    if (window.innerWidth < 768) {
      el.dashboard.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }

  /* ---------------------------------------------------------------------
   * DASHBOARD RENDERING
   * ------------------------------------------------------------------- */

  function setText(id, value) {
    const node = document.getElementById(id);
    if (!node) return;
    const isEmpty = value === undefined || value === null || String(value).trim() === '';
    node.textContent = isEmpty ? 'Not available' : value;
    node.classList.toggle('is-empty', isEmpty);
  }

  function renderPartyValue(id, partyName) {
    const node = document.getElementById(id);
    node.innerHTML = '';
    if (!partyName) {
      node.classList.add('is-empty');
      node.textContent = 'Not available';
      return;
    }
    node.classList.remove('is-empty');
    const dot = document.createElement('span');
    dot.className = 'party-dot';
    dot.style.background = PARTY_COLORS[partyName.toUpperCase()] || '#94A3B8';
    const label = document.createElement('span');
    label.textContent = partyName;
    node.appendChild(dot);
    node.appendChild(label);
  }

  function renderDashboard(r) {
    setText('val-acNo', r.acNo);
    setText('val-acName', r.acName);
    setText('val-zone', r.zone);
    setText('val-pc', r.pc);
    setText('val-district', r.district);
    setText('val-orgDistrict', r.orgDistrict);

    setText('val-winnerName', r.winnerName);
    renderPartyValue('val-winnerParty', r.winnerParty);
    setText('val-winnerCaste', r.winnerCaste);
    setText('val-winnerCategory', r.winnerCategory);

    setText('val-runnerName', r.runnerName);
    renderPartyValue('val-runnerParty', r.runnerParty);
    setText('val-runnerCaste', r.runnerCaste);
    setText('val-runnerCategory', r.runnerCategory);

    renderBoothGrade(r.boothGrade);
    renderCasteList([r.caste1, r.caste2, r.caste3, r.caste4, r.caste5]);
  }

  /**
   * Parse a multi-line "booth gradation" cell, e.g.:
   *   "Total no of Booths - 248\n______________\nGrade A - 51\nGrade B - 16\n..."
   * into a total + grade breakdown, then render a proportional bar + legend.
   */
  function renderBoothGrade(raw) {
    const node = document.getElementById('val-boothGrade');
    node.innerHTML = '';

    if (!raw) {
      node.classList.add('is-empty');
      node.textContent = 'Not available';
      return;
    }
    node.classList.remove('is-empty');

    const lines = raw.split('\n').map((l) => l.trim()).filter((l) => l && !/^_+$/.test(l));

    let total = null;
    const grades = {}; // { A: 51, B: 16, ... }

    lines.forEach((line) => {
      const totalMatch = line.match(/total.*?(\d+)/i);
      const gradeMatch = line.match(/grade\s*([A-Da-d])\s*-\s*(\d+)/i);
      if (gradeMatch) {
        grades[gradeMatch[1].toUpperCase()] = parseInt(gradeMatch[2], 10);
      } else if (totalMatch) {
        total = parseInt(totalMatch[1], 10);
      }
    });

    const gradeKeys = ['A', 'B', 'C', 'D'].filter((k) => grades[k] !== undefined);
    const sum = gradeKeys.reduce((acc, k) => acc + grades[k], 0) || total || 0;

    if (total !== null) {
      const totalEl = document.createElement('div');
      totalEl.className = 'booth-total';
      totalEl.textContent = `${total} booths`;
      node.appendChild(totalEl);
    }

    if (gradeKeys.length > 0 && sum > 0) {
      const bar = document.createElement('div');
      bar.className = 'booth-bar';
      gradeKeys.forEach((k) => {
        const seg = document.createElement('span');
        seg.className = `g-${k.toLowerCase()}`;
        seg.style.width = `${(grades[k] / sum) * 100}%`;
        bar.appendChild(seg);
      });
      node.appendChild(bar);

      const legend = document.createElement('div');
      legend.className = 'booth-legend';
      gradeKeys.forEach((k) => {
        const item = document.createElement('span');
        item.innerHTML = `<i class="g-${k.toLowerCase()}"></i> Grade ${k} &middot; ${grades[k]}`;
        legend.appendChild(item);
      });
      node.appendChild(legend);
    } else if (total === null) {
      // Fallback: unparsed / unexpected format — show raw text safely.
      const fallback = document.createElement('div');
      fallback.textContent = raw.replace(/\n/g, ' · ');
      node.appendChild(fallback);
    }
  }

  /**
   * Parse a "Dominant Caste N" cell, e.g.:
   *   "Lingayat - 20%\n_______________\n\nLingayat Chaturth - 12%\nLingayat Panchamasali - 6%"
   * into a headline (name + share) plus optional sub-breakdown lines.
   */
  function parseCasteCell(raw) {
    if (!raw) return null;
    const lines = raw.split('\n').map((l) => l.trim()).filter((l) => l && !/^_+$/.test(l));
    if (lines.length === 0) return null;

    const [headline, ...rest] = lines;
    const headlineMatch = headline.match(/^(.*?)-\s*([\d.]+%?)$/);

    const sub = rest.map((line) => {
      const m = line.match(/^(.*?)-\s*([\d.]+%?)$/);
      return m ? { label: m[1].trim(), value: m[2].trim() } : { label: line, value: '' };
    });

    return headlineMatch
      ? { name: headlineMatch[1].trim(), share: headlineMatch[2].trim(), sub }
      : { name: headline, share: '', sub };
  }

  function renderCasteList(rawValues) {
    el.casteList.innerHTML = '';
    const parsed = rawValues.map(parseCasteCell).filter(Boolean);

    if (parsed.length === 0) {
      const empty = document.createElement('p');
      empty.className = 'caste-empty';
      empty.textContent = 'No dominant caste data available for this constituency.';
      el.casteList.appendChild(empty);
      return;
    }

    parsed.forEach((c, i) => {
      const card = document.createElement('div');
      card.className = 'caste-card';

      const rank = document.createElement('span');
      rank.className = 'caste-rank';
      rank.textContent = `#${i + 1}`;
      card.appendChild(rank);

      const headline = document.createElement('p');
      headline.className = 'caste-headline';
      headline.innerHTML = `${escapeHtml(c.name)}${c.share ? ` <span class="caste-share">${escapeHtml(c.share)}</span>` : ''}`;
      card.appendChild(headline);

      if (c.sub.length > 0) {
        const subList = document.createElement('ul');
        subList.className = 'caste-sub';
        c.sub.forEach((s) => {
          const li = document.createElement('li');
          li.innerHTML = `<span>${escapeHtml(s.label)}</span><b>${escapeHtml(s.value)}</b>`;
          subList.appendChild(li);
        });
        card.appendChild(subList);
      }

      el.casteList.appendChild(card);
    });
  }

  /* ---------------------------------------------------------------------
   * EVENT WIRING
   * ------------------------------------------------------------------- */

  function initEvents() {
    el.searchInput.addEventListener('input', (e) => {
      el.clearBtn.hidden = e.target.value.length === 0;
      buildOptions(e.target.value);
    });

    el.searchInput.addEventListener('focus', () => buildOptions(el.searchInput.value));

    el.searchInput.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowDown') { e.preventDefault(); openList(); moveActiveOption(1); }
      else if (e.key === 'ArrowUp') { e.preventDefault(); openList(); moveActiveOption(-1); }
      else if (e.key === 'Enter') { e.preventDefault(); selectActiveOption(); }
      else if (e.key === 'Escape') { closeList(); }
    });

    el.searchInput.addEventListener('blur', () => {
      // Delay so option mousedown handlers can still fire.
      setTimeout(closeList, 120);
    });

    el.clearBtn.addEventListener('click', () => {
      resetSelection();
      showState('empty');
      el.searchInput.focus();
    });

    document.addEventListener('click', (e) => {
      if (!e.target.closest('#combobox')) closeList();
    });

    el.fileInput.addEventListener('change', handleFileSelect);
  }

  /* ---------------------------------------------------------------------
   * INIT
   * ------------------------------------------------------------------- */

  function init() {
    initEvents();
    loadDefaultData();
  }

  document.addEventListener('DOMContentLoaded', init);
})();
