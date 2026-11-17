(() => {
  'use strict';

  const SAMPLE_CSV_PATH = './Data_Presensi_LSP_Pengadaan.csv';
  const SAMPLE_EXCEL_PATH = './Laporan_Presensi_LSP_Lengkap.xlsx';

  const COLUMN_DEFINITIONS = [
    { key: 'No', label: 'No', technical: false, wrap: false, isStatus: false, isPhoto: false, pdfWidth: 10 },
    { key: 'Tanggal Presensi', label: 'Tanggal Presensi', technical: false, wrap: false, isStatus: false, isPhoto: false, pdfWidth: 22 },
    { key: 'Waktu Masuk (WIB)', label: 'Waktu Masuk (WIB)', technical: false, wrap: false, isStatus: false, isPhoto: false, pdfWidth: 28 },
    { key: 'ID / Username', label: 'ID / Username', technical: false, wrap: false, isStatus: false, isPhoto: false, pdfWidth: 22 },
    { key: 'Nama Lengkap', label: 'Nama Lengkap', technical: false, wrap: false, isStatus: false, isPhoto: false, pdfWidth: 24 },
    { key: 'Status Penugasan', label: 'Status Penugasan', technical: false, wrap: false, isStatus: true, isPhoto: false, pdfWidth: 22 },
    { key: 'Aktivitas / Logbook', label: 'Aktivitas / Logbook', technical: false, wrap: true, isStatus: false, isPhoto: false, pdfWidth: 58 },
    { key: 'Alamat Lokasi Terdeteksi', label: 'Alamat Lokasi Terdeteksi', technical: false, wrap: true, isStatus: false, isPhoto: false, pdfWidth: 62 },
    { key: 'Latitude', label: 'Latitude', technical: true, wrap: false, isStatus: false, isPhoto: false, pdfWidth: 18 },
    { key: 'Longitude', label: 'Longitude', technical: true, wrap: false, isStatus: false, isPhoto: false, pdfWidth: 18 },
    { key: 'Link Foto Ruangan', label: 'Link Foto Ruangan', technical: true, wrap: false, isStatus: false, isPhoto: true, pdfWidth: 26 },
    { key: 'Link Foto Diri', label: 'Link Foto Diri', technical: true, wrap: false, isStatus: false, isPhoto: true, pdfWidth: 26 },
  ];

  // Map raw CSV / Excel header variations to canonical column keys
  const HEADER_ALIASES = {
    'no': 'No',
    'tanggal': 'Tanggal Presensi',
    'tanggal presensi': 'Tanggal Presensi',
    'waktu (wib)': 'Waktu Masuk (WIB)',
    'waktu masuk (wib)': 'Waktu Masuk (WIB)',
    'username': 'ID / Username',
    'id / username': 'ID / Username',
    'nama lengkap': 'Nama Lengkap',
    'status': 'Status Penugasan',
    'status penugasan': 'Status Penugasan',
    'aktivitas / logbook': 'Aktivitas / Logbook',
    'alamat lokasi': 'Alamat Lokasi Terdeteksi',
    'alamat lokasi terdeteksi': 'Alamat Lokasi Terdeteksi',
    'latitude': 'Latitude',
    'longitude': 'Longitude',
    'foto 1: ruangan': 'Link Foto Ruangan',
    'link foto ruangan': 'Link Foto Ruangan',
    'foto 2: selfie / id': 'Link Foto Diri',
    'link foto diri': 'Link Foto Diri',
    'link foto diri (selfie)': 'Link Foto Diri',
  };

  const state = {
    rawRows: [],
    filteredRows: [],
    columns: [],
    visibleColumns: new Set(),
    sort: { key: null, direction: 'asc' },
    fileName: '',
  };

  // DOM Elements
  const els = {
    dropZone: document.getElementById('dropZone'),
    csvFileInput: document.getElementById('csvFileInput'),
    btnLoadSample: document.getElementById('btnLoadSample'),
    btnLoadSampleExcel: document.getElementById('btnLoadSampleExcel'),
    btnExportExcel: document.getElementById('btnExportExcel'),
    btnExportPdf: document.getElementById('btnExportPdf'),
    fileBadge: document.getElementById('fileBadge'),
    fileNameLabel: document.getElementById('fileNameLabel'),
    alertBanner: document.getElementById('alertBanner'),
    alertIcon: document.getElementById('alertIcon'),
    alertMessage: document.getElementById('alertMessage'),
    btnCloseAlert: document.getElementById('btnCloseAlert'),
    dashboardContent: document.getElementById('dashboardContent'),
    statTotal: document.getElementById('statTotal'),
    statFilteredNote: document.getElementById('statFilteredNote'),
    statDiKantor: document.getElementById('statDiKantor'),
    statLuarKantor: document.getElementById('statLuarKantor'),
    statEmployees: document.getElementById('statEmployees'),
    filterEmployee: document.getElementById('filterEmployee'),
    filterStatus: document.getElementById('filterStatus'),
    filterDateFrom: document.getElementById('filterDateFrom'),
    filterDateTo: document.getElementById('filterDateTo'),
    filterKeyword: document.getElementById('filterKeyword'),
    btnResetFilters: document.getElementById('btnResetFilters'),
    columnTogglesContainer: document.getElementById('columnTogglesContainer'),
    btnHideTechnical: document.getElementById('btnHideTechnical'),
    btnShowAllCols: document.getElementById('btnShowAllCols'),
    tableCountShown: document.getElementById('tableCountShown'),
    tableCountTotal: document.getElementById('tableCountTotal'),
    tableHead: document.getElementById('tableHead'),
    tableBody: document.getElementById('tableBody'),
    emptyFilteredState: document.getElementById('emptyFilteredState'),
    imageModal: document.getElementById('imageModal'),
    imageModalTitle: document.getElementById('imageModalTitle'),
    imageModalImg: document.getElementById('imageModalImg'),
    btnCloseImageModal: document.getElementById('btnCloseImageModal'),
  };

  // ---------------------------------------------------------------------------
  // Utilities
  // ---------------------------------------------------------------------------
  function showAlert(message, type = 'error') {
    const styles = {
      error: {
        banner: 'bg-rose-50 border-rose-200 text-rose-800',
        icon: '<path stroke-linecap="round" stroke-linejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />',
      },
      success: {
        banner: 'bg-emerald-50 border-emerald-200 text-emerald-800',
        icon: '<path stroke-linecap="round" stroke-linejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />',
      },
      info: {
        banner: 'bg-indigo-50 border-indigo-200 text-indigo-800',
        icon: '<path stroke-linecap="round" stroke-linejoin="round" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />',
      },
    };
    const cfg = styles[type] || styles.info;
    els.alertBanner.className = `rounded-xl border p-4 text-sm flex items-start justify-between gap-3 ${cfg.banner}`;
    els.alertIcon.innerHTML = cfg.icon;
    els.alertMessage.textContent = message;
    els.alertBanner.classList.remove('hidden');
  }

  function hideAlert() {
    els.alertBanner.classList.add('hidden');
  }

  function normalizeHeader(rawHeader) {
    const cleaned = String(rawHeader || '').trim();
    const lower = cleaned.toLowerCase();
    return HEADER_ALIASES[lower] || cleaned;
  }

  function isValidHttpUrl(value) {
    if (!value) return false;
    const str = String(value).trim();
    return /^https?:\/\/\S+$/i.test(str);
  }

  function isDataImageUrl(value) {
    if (!value) return false;
    return /^data:image\/[a-zA-Z0-9+.-]+;base64,/i.test(String(value).trim());
  }

  function getImageFormatFromDataUrl(dataUrl) {
    const match = String(dataUrl || '').match(/^data:image\/([a-zA-Z0-9+.-]+);base64,/i);
    if (!match) return 'JPEG';
    const ext = match[1].toUpperCase();
    return ext === 'JPG' ? 'JPEG' : ext;
  }

  function openImageModal(src, title) {
    els.imageModalTitle.textContent = title || 'Preview Foto';
    els.imageModalImg.src = src;
    els.imageModal.classList.remove('hidden');
  }

  function closeImageModal() {
    els.imageModal.classList.add('hidden');
    els.imageModalImg.src = '';
  }

  function getTodayStamp() {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  }

  function getFormattedTimestamp() {
    const d = new Date();
    return d.toLocaleString('id-ID', {
      year: 'numeric',
      month: 'long',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
    });
  }

  // ---------------------------------------------------------------------------
  // CSV Parsing & Fallback Unpacking
  // ---------------------------------------------------------------------------
  function unpackSingleColumnRow(cellValue) {
    let raw = String(cellValue || '').trim();
    if (!raw || !raw.includes(',')) {
      return [raw];
    }

    // If wrapped in outer quotes (e.g. Excel re-quoting a packed CSV line)
    if (raw.startsWith('"') && raw.endsWith('"') && raw.length >= 2) {
      const inner = raw.slice(1, -1);
      // Unescape doubled quotes if present
      raw = inner.replace(/""/g, '"');
    }

    const reparsed = Papa.parse(raw, {
      header: false,
      skipEmptyLines: false,
    });

    if (reparsed.data && reparsed.data.length > 0 && reparsed.data[0].length > 1) {
      return reparsed.data[0];
    }

    return [cellValue];
  }

  function parseCsvText(csvText, fileName) {
    hideAlert();
    const cleanText = String(csvText || '').replace(/^\uFEFF/, '');

    const parsed = Papa.parse(cleanText, {
      header: false,
      skipEmptyLines: 'greedy',
    });

    if (!parsed.data || parsed.data.length < 2) {
      showAlert('File CSV kosong atau tidak memiliki baris data yang valid.', 'error');
      return;
    }

    // Normalize each row, applying fallback unpacker if a row is packed into 1 column
    const normalizedMatrix = parsed.data.map((row) => {
      if (Array.isArray(row) && row.length === 1 && String(row[0]).includes(',')) {
        return unpackSingleColumnRow(row[0]);
      }
      return row;
    });

    populateStateFromMatrix(normalizedMatrix, fileName);
  }

  // ---------------------------------------------------------------------------
  // Excel (.xlsx / .xls) Parsing & Embedded Image Extraction
  // ---------------------------------------------------------------------------
  async function extractExcelImages(arrayBuffer) {
    const imageMap = new Map();
    if (typeof window.JSZip === 'undefined') {
      return imageMap;
    }

    try {
      const zip = await window.JSZip.loadAsync(arrayBuffer);
      const drawingFiles = Object.keys(zip.files).filter((name) =>
        /^xl\/drawings\/drawing\d*\.xml$/i.test(name)
      );

      const parser = new DOMParser();

      for (const drawingPath of drawingFiles) {
        const fileName = drawingPath.split('/').pop();
        const relsPath = `xl/drawings/_rels/${fileName}.rels`;
        const relsFile = zip.file(relsPath);
        if (!relsFile) continue;

        const relsXml = await relsFile.async('text');
        const relsDoc = parser.parseFromString(relsXml, 'application/xml');
        const relMap = new Map();
        const relationships = relsDoc.getElementsByTagName('Relationship');

        for (const rel of relationships) {
          const id = rel.getAttribute('Id');
          let target = rel.getAttribute('Target') || '';
          if (target.startsWith('/')) {
            target = target.slice(1);
          } else if (target.startsWith('../')) {
            target = `xl/${target.slice(3)}`;
          } else if (!target.startsWith('xl/')) {
            target = `xl/drawings/${target}`;
          }
          if (id && target) {
            relMap.set(id, target);
          }
        }

        const drawingXml = await zip.file(drawingPath).async('text');
        const drawingDoc = parser.parseFromString(drawingXml, 'application/xml');
        const anchors = [
          ...drawingDoc.getElementsByTagNameNS('*', 'oneCellAnchor'),
          ...drawingDoc.getElementsByTagNameNS('*', 'twoCellAnchor'),
        ];

        for (const anchor of anchors) {
          const fromEl = anchor.getElementsByTagNameNS('*', 'from')[0];
          const blipEl = anchor.getElementsByTagNameNS('*', 'blip')[0];
          if (!fromEl || !blipEl) continue;

          const colEl = fromEl.getElementsByTagNameNS('*', 'col')[0];
          const rowEl = fromEl.getElementsByTagNameNS('*', 'row')[0];
          const rId =
            blipEl.getAttribute('r:embed') ||
            blipEl.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed');

          if (!colEl || !rowEl || !rId) continue;

          const colIdx = Number(colEl.textContent);
          const rowIdx = Number(rowEl.textContent);
          const mediaPath = relMap.get(rId);
          const mediaFile = mediaPath ? zip.file(mediaPath) : null;
          if (!mediaFile) continue;

          const base64 = await mediaFile.async('base64');
          const ext = (mediaPath.split('.').pop() || 'jpeg').toLowerCase();
          const mime = ext === 'png' ? 'image/png' : ext === 'webp' ? 'image/webp' : 'image/jpeg';
          imageMap.set(`${rowIdx},${colIdx}`, `data:${mime};base64,${base64}`);
        }
      }
    } catch (err) {
      console.warn('Gagal mengekstrak gambar tertanam dari Excel:', err);
    }

    return imageMap;
  }

  async function parseExcelBuffer(arrayBuffer, fileName) {
    hideAlert();
    try {
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      if (!sheetName) {
        showAlert('File Excel tidak memiliki sheet yang valid.', 'error');
        return;
      }

      const worksheet = workbook.Sheets[sheetName];
      const matrix = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: '',
        raw: false,
      });

      const embeddedImages = await extractExcelImages(arrayBuffer);
      populateStateFromMatrix(matrix, fileName, embeddedImages);
    } catch (err) {
      showAlert(`Gagal memproses file Excel: ${err.message}`, 'error');
    }
  }

  function populateStateFromMatrix(matrix, fileName, embeddedImages = new Map()) {
    if (!matrix || matrix.length < 2) {
      showAlert('File kosong atau tidak memiliki baris data yang valid.', 'error');
      return;
    }

    const rawHeaders = matrix[0].map((h) => normalizeHeader(h));
    const indexedRows = [];
    for (let r = 1; r < matrix.length; r++) {
      const rowCells = matrix[r] || [];
      const hasText = rowCells.some((cell) => String(cell || '').trim() !== '');
      const hasImage = rawHeaders.some((_, cIdx) => embeddedImages.has(`${r},${cIdx}`));
      if (hasText || hasImage) {
        indexedRows.push({ rowIdx: r, cells: rowCells });
      }
    }

    if (indexedRows.length === 0) {
      showAlert('Header ditemukan, tetapi tidak ada baris data presensi.', 'error');
      return;
    }

    // Build column definitions preserving order from file, enriched with known metadata
    const defMap = new Map(COLUMN_DEFINITIONS.map((c) => [c.key, c]));
    state.columns = rawHeaders.map((key) => {
      const known = defMap.get(key);
      if (known) return { ...known };
      return {
        key,
        label: key,
        technical: false,
        wrap: false,
        isStatus: false,
        isPhoto: false,
        pdfWidth: 25,
      };
    });

    // Convert matrix rows into keyed objects (attaching embedded images if cell is empty)
    state.rawRows = indexedRows.map(({ rowIdx, cells }) => {
      const obj = {};
      state.columns.forEach((col, colIdx) => {
        const textVal = cells[colIdx] !== undefined && cells[colIdx] !== null ? String(cells[colIdx]).trim() : '';
        const imgVal = embeddedImages.get(`${rowIdx},${colIdx}`);
        obj[col.key] = textVal || imgVal || '';
      });
      return obj;
    });

    state.visibleColumns = new Set(state.columns.map((c) => c.key));
    state.sort = { key: null, direction: 'asc' };
    state.fileName = fileName;

    // Update UI
    els.fileNameLabel.textContent = fileName;
    els.fileBadge.classList.remove('hidden');
    els.fileBadge.classList.add('inline-flex');
    els.dashboardContent.classList.remove('hidden');
    els.btnExportExcel.disabled = false;
    els.btnExportPdf.disabled = false;

    populateEmployeeFilter();
    resetFilterInputs();
    renderColumnToggles();
    applyFiltersAndRender();

    const imgNote = embeddedImages.size > 0 ? ` (${embeddedImages.size} foto tertanam diekstrak)` : '';
    showAlert(`Berhasil memuat ${state.rawRows.length} baris data dari "${fileName}"${imgNote}.`, 'success');
  }

  // ---------------------------------------------------------------------------
  // File Loading Handlers
  // ---------------------------------------------------------------------------
  function handleFileObject(file) {
    if (!file) return;
    const lowerName = file.name.toLowerCase();
    const isCsv = lowerName.endsWith('.csv') || file.type === 'text/csv';
    const isExcel = lowerName.endsWith('.xlsx') || lowerName.endsWith('.xls');

    if (!isCsv && !isExcel) {
      showAlert('Harap pilih file dengan format .csv atau .xlsx.', 'error');
      return;
    }

    const reader = new FileReader();
    reader.onerror = () => {
      showAlert('Gagal membaca file dari perangkat Anda.', 'error');
    };

    if (isExcel) {
      reader.onload = (e) => {
        parseExcelBuffer(e.target.result, file.name);
      };
      reader.readAsArrayBuffer(file);
    } else {
      reader.onload = (e) => {
        parseCsvText(e.target.result, file.name);
      };
      reader.readAsText(file, 'UTF-8');
    }
  }

  async function loadSampleCsv() {
    els.btnLoadSample.disabled = true;
    const originalHtml = els.btnLoadSample.innerHTML;
    els.btnLoadSample.innerHTML = '<span>Memuat sampel...</span>';

    try {
      const response = await fetch(SAMPLE_CSV_PATH);
      if (!response.ok) {
        throw new Error(`HTTP ${response.status} ${response.statusText}`);
      }
      const text = await response.text();
      parseCsvText(text, 'Data_Presensi_LSP_Pengadaan.csv');
    } catch (err) {
      const isFileProtocol = window.location.protocol === 'file:';
      const hint = isFileProtocol
        ? ' Browser memblokir fetch file lokal pada protokol file://. Silakan seret & lepas file langsung ke kotak unggah, atau jalankan lewat local server (mis. "npx serve" / "python -m http.server").'
        : ` Detail: ${err.message}`;
      showAlert(`Gagal memuat sampel "${SAMPLE_CSV_PATH}".${hint}`, 'error');
    } finally {
      els.btnLoadSample.disabled = false;
      els.btnLoadSample.innerHTML = originalHtml;
    }
  }

  async function loadSampleExcel() {
    els.btnLoadSampleExcel.disabled = true;
    const originalHtml = els.btnLoadSampleExcel.innerHTML;
    els.btnLoadSampleExcel.innerHTML = '<span>Memuat Excel...</span>';

    try {
      const response = await fetch(SAMPLE_EXCEL_PATH);
      if (!response.ok) {
        throw new Error(`HTTP ${response.status} ${response.statusText}`);
      }
      const buffer = await response.arrayBuffer();
      await parseExcelBuffer(buffer, 'Laporan_Presensi_LSP_Lengkap.xlsx');
    } catch (err) {
      const isFileProtocol = window.location.protocol === 'file:';
      const hint = isFileProtocol
        ? ' Browser memblokir fetch file lokal pada protokol file://. Silakan seret & lepas file langsung ke kotak unggah, atau jalankan lewat local server (mis. "npx serve" / "python -m http.server").'
        : ` Detail: ${err.message}`;
      showAlert(`Gagal memuat sampel "${SAMPLE_EXCEL_PATH}".${hint}`, 'error');
    } finally {
      els.btnLoadSampleExcel.disabled = false;
      els.btnLoadSampleExcel.innerHTML = originalHtml;
    }
  }

  // ---------------------------------------------------------------------------
  // Filters, Sorting & Stats
  // ---------------------------------------------------------------------------
  function populateEmployeeFilter() {
    const names = Array.from(
      new Set(state.rawRows.map((r) => r['Nama Lengkap']).filter(Boolean))
    ).sort((a, b) => a.localeCompare(b, 'id'));

    els.filterEmployee.innerHTML = '<option value="">Semua Pegawai</option>';
    names.forEach((name) => {
      const opt = document.createElement('option');
      opt.value = name;
      opt.textContent = name;
      els.filterEmployee.appendChild(opt);
    });
  }

  function resetFilterInputs() {
    els.filterEmployee.value = '';
    els.filterStatus.value = '';
    els.filterDateFrom.value = '';
    els.filterDateTo.value = '';
    els.filterKeyword.value = '';
  }

  function applyFiltersAndRender() {
    const emp = els.filterEmployee.value;
    const status = els.filterStatus.value;
    const dateFrom = els.filterDateFrom.value;
    const dateTo = els.filterDateTo.value;
    const keyword = els.filterKeyword.value.trim().toLowerCase();

    let rows = state.rawRows.filter((row) => {
      if (emp && row['Nama Lengkap'] !== emp) return false;
      if (status && row['Status Penugasan'] !== status) return false;

      const rowDate = (row['Tanggal Presensi'] || '').slice(0, 10);
      if (dateFrom && rowDate < dateFrom) return false;
      if (dateTo && rowDate > dateTo) return false;

      if (keyword) {
        const logbook = (row['Aktivitas / Logbook'] || '').toLowerCase();
        if (!logbook.includes(keyword)) return false;
      }

      return true;
    });

    if (state.sort.key) {
      const { key, direction } = state.sort;
      const mult = direction === 'asc' ? 1 : -1;
      rows = [...rows].sort((a, b) => {
        const va = a[key] || '';
        const vb = b[key] || '';
        if (key === 'No') {
          const na = Number(va);
          const nb = Number(vb);
          if (!Number.isNaN(na) && !Number.isNaN(nb)) {
            return (na - nb) * mult;
          }
        }
        return va.localeCompare(vb, 'id', { numeric: true }) * mult;
      });
    }

    state.filteredRows = rows;
    updateSummaryCards();
    renderTable();
  }

  function updateSummaryCards() {
    const rows = state.filteredRows;
    const totalAll = state.rawRows.length;
    const diKantorCount = rows.filter((r) => (r['Status Penugasan'] || '').toLowerCase() === 'di kantor').length;
    const luarKantorCount = rows.filter((r) => (r['Status Penugasan'] || '').toLowerCase() === 'luar kantor').length;
    const uniqueEmployees = new Set(rows.map((r) => r['Nama Lengkap']).filter(Boolean)).size;

    els.statTotal.textContent = rows.length;
    els.statFilteredNote.textContent = rows.length === totalAll ? 'Semua data ditampilkan' : `Dari total ${totalAll} data`;
    els.statDiKantor.textContent = diKantorCount;
    els.statLuarKantor.textContent = luarKantorCount;
    els.statEmployees.textContent = uniqueEmployees;

    els.tableCountShown.textContent = rows.length;
    els.tableCountTotal.textContent = totalAll;
  }

  // ---------------------------------------------------------------------------
  // Column Visibility Toggles
  // ---------------------------------------------------------------------------
  function renderColumnToggles() {
    els.columnTogglesContainer.innerHTML = '';
    state.columns.forEach((col) => {
      const isChecked = state.visibleColumns.has(col.key);
      const label = document.createElement('label');
      label.className = `inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-medium cursor-pointer border transition select-none ${
        isChecked
          ? 'bg-indigo-50 text-indigo-700 border-indigo-200'
          : 'bg-slate-50 text-slate-500 border-slate-200 hover:bg-slate-100'
      }`;

      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.checked = isChecked;
      checkbox.className = 'rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 w-3.5 h-3.5';
      checkbox.addEventListener('change', () => {
        if (checkbox.checked) {
          state.visibleColumns.add(col.key);
        } else {
          if (state.visibleColumns.size <= 1) {
            checkbox.checked = true;
            showAlert('Minimal satu kolom harus tetap ditampilkan.', 'info');
            return;
          }
          state.visibleColumns.delete(col.key);
        }
        renderColumnToggles();
        renderTable();
      });

      const textSpan = document.createElement('span');
      textSpan.textContent = col.label;

      label.appendChild(checkbox);
      label.appendChild(textSpan);
      els.columnTogglesContainer.appendChild(label);
    });
  }

  function getActiveColumns() {
    return state.columns.filter((c) => state.visibleColumns.has(c.key));
  }

  // ---------------------------------------------------------------------------
  // Table Rendering
  // ---------------------------------------------------------------------------
  function renderTable() {
    const activeCols = getActiveColumns();

    // Render Head
    els.tableHead.innerHTML = '';
    const headRow = document.createElement('tr');
    activeCols.forEach((col) => {
      const th = document.createElement('th');
      th.className =
        'px-3.5 py-3 text-xs font-semibold uppercase tracking-wider text-slate-600 whitespace-nowrap cursor-pointer hover:bg-slate-200/70 transition select-none';
      th.addEventListener('click', () => {
        if (state.sort.key === col.key) {
          state.sort.direction = state.sort.direction === 'asc' ? 'desc' : 'asc';
        } else {
          state.sort.key = col.key;
          state.sort.direction = 'asc';
        }
        applyFiltersAndRender();
      });

      const wrap = document.createElement('div');
      wrap.className = 'flex items-center gap-1';
      const titleSpan = document.createElement('span');
      titleSpan.textContent = col.label;
      wrap.appendChild(titleSpan);

      if (state.sort.key === col.key) {
        const arrow = document.createElement('span');
        arrow.className = 'text-indigo-600 font-bold';
        arrow.textContent = state.sort.direction === 'asc' ? '▲' : '▼';
        wrap.appendChild(arrow);
      }

      th.appendChild(wrap);
      headRow.appendChild(th);
    });
    els.tableHead.appendChild(headRow);

    // Render Body
    els.tableBody.innerHTML = '';
    if (state.filteredRows.length === 0) {
      els.emptyFilteredState.classList.remove('hidden');
      return;
    }
    els.emptyFilteredState.classList.add('hidden');

    const fragment = document.createDocumentFragment();
    state.filteredRows.forEach((row) => {
      const tr = document.createElement('tr');
      tr.className = 'hover:bg-slate-50/90 transition-colors align-top';

      activeCols.forEach((col) => {
        const td = document.createElement('td');
        const val = row[col.key] || '';

        if (col.isStatus) {
          td.className = 'px-3.5 py-3 whitespace-nowrap';
          const badge = document.createElement('span');
          const lower = val.toLowerCase();
          if (lower === 'di kantor') {
            badge.className =
              'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200';
            badge.innerHTML = '<span class="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>';
          } else if (lower === 'luar kantor') {
            badge.className =
              'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold bg-amber-50 text-amber-700 border border-amber-200';
            badge.innerHTML = '<span class="w-1.5 h-1.5 rounded-full bg-amber-500"></span>';
          } else {
            badge.className =
              'inline-flex items-center px-2.5 py-1 rounded-full text-xs font-medium bg-slate-100 text-slate-700';
          }
          const txt = document.createElement('span');
          txt.textContent = val || '-';
          badge.appendChild(txt);
          td.appendChild(badge);
        } else if (col.isPhoto) {
          td.className = 'px-3.5 py-3 whitespace-nowrap';
          if (isDataImageUrl(val)) {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className =
              'group relative inline-flex items-center justify-center rounded-lg overflow-hidden border border-slate-200 hover:ring-2 hover:ring-indigo-500 transition shadow-sm bg-slate-100';
            btn.title = 'Klik untuk memperbesar foto';
            btn.addEventListener('click', () => {
              openImageModal(val, `${col.label} — ${row['Nama Lengkap'] || ''} (${row['Tanggal Presensi'] || ''})`);
            });

            const img = document.createElement('img');
            img.src = val;
            img.alt = col.label;
            img.className = 'w-12 h-12 object-cover';
            btn.appendChild(img);
            td.appendChild(btn);
          } else if (isValidHttpUrl(val)) {
            const link = document.createElement('a');
            link.href = val;
            link.target = '_blank';
            link.rel = 'noopener noreferrer';
            link.className =
              'inline-flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-medium bg-indigo-50 text-indigo-700 hover:bg-indigo-100 border border-indigo-200 transition';
            link.innerHTML = `
              <span>View Photo</span>
              <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14" />
              </svg>
            `;
            td.appendChild(link);
          } else {
            td.className += ' text-slate-400 text-xs';
            td.textContent = val || '-';
          }
        } else if (col.wrap) {
          td.className = 'px-3.5 py-3 text-slate-700 text-xs sm:text-sm min-w-[240px] max-w-[360px] whitespace-normal break-words leading-relaxed';
          td.textContent = val || '-';
        } else {
          td.className = 'px-3.5 py-3 text-slate-700 text-xs sm:text-sm whitespace-nowrap';
          if (col.key === 'No') {
            td.className += ' font-medium text-slate-500 text-center';
          } else if (col.key === 'Nama Lengkap') {
            td.className += ' font-semibold text-slate-900';
          }
          td.textContent = val || '-';
        }

        tr.appendChild(td);
      });

      fragment.appendChild(tr);
    });

    els.tableBody.appendChild(fragment);
  }

  // ---------------------------------------------------------------------------
  // Export to Excel (SheetJS / xlsx)
  // ---------------------------------------------------------------------------
  function exportToExcel() {
    if (state.filteredRows.length === 0) {
      showAlert('Tidak ada data untuk diekspor ke Excel.', 'info');
      return;
    }

    const activeCols = getActiveColumns();
    const headers = activeCols.map((c) => c.label);
    const body = state.filteredRows.map((row) =>
      activeCols.map((c) => {
        const v = row[c.key] || '';
        return isDataImageUrl(v) ? '[Foto Terlampir]' : v;
      })
    );
    const aoa = [headers, ...body];

    const worksheet = XLSX.utils.aoa_to_sheet(aoa);

    // Auto-adjust column widths with sensible min/max bounds
    worksheet['!cols'] = activeCols.map((col, idx) => {
      let maxLen = col.label.length;
      for (const r of body) {
        const cellText = String(r[idx] || '');
        if (cellText.length > maxLen) {
          maxLen = cellText.length;
        }
      }
      const width = Math.min(Math.max(maxLen + 3, 10), col.wrap ? 55 : 35);
      return { wch: width };
    });

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Presensi & Logbook');

    const outName = `Laporan_Presensi_LSP_Pengadaan_${getTodayStamp()}.xlsx`;
    XLSX.writeFile(workbook, outName);
  }

  // ---------------------------------------------------------------------------
  // Export to Landscape PDF (jsPDF + jspdf-autotable)
  // ---------------------------------------------------------------------------
  function buildFilterSummaryText() {
    const parts = [];
    if (els.filterEmployee.value) parts.push(`Pegawai: ${els.filterEmployee.value}`);
    if (els.filterStatus.value) parts.push(`Status: ${els.filterStatus.value}`);
    if (els.filterDateFrom.value || els.filterDateTo.value) {
      parts.push(`Tanggal: ${els.filterDateFrom.value || '...'} s/d ${els.filterDateTo.value || '...'}`);
    }
    if (els.filterKeyword.value.trim()) {
      parts.push(`Kata Kunci: "${els.filterKeyword.value.trim()}"`);
    }
    const filterDesc = parts.length > 0 ? parts.join('  |  ') : 'Semua Data (Tanpa Filter)';
    return `${filterDesc}  |  Total Baris: ${state.filteredRows.length}`;
  }

  function exportToPdf() {
    if (state.filteredRows.length === 0) {
      showAlert('Tidak ada data untuk diekspor ke PDF.', 'info');
      return;
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a4',
    });

    const totalPagesExp = '{total_pages_count_string}';
    const pageWidth = doc.internal.pageSize.getWidth();
    const marginX = 12;

    // Header Title
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(14);
    doc.setTextColor(15, 23, 42);
    doc.text('Laporan Presensi & Logbook LSP Pengadaan', marginX, 14);

    // Timestamp & Filter Summary
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8.5);
    doc.setTextColor(71, 85, 105);
    doc.text(`Dibuat pada: ${getFormattedTimestamp()}`, marginX, 19.5);
    doc.text(`Filter Aktif: ${buildFilterSummaryText()}`, marginX, 24);

    const activeCols = getActiveColumns();
    const head = [activeCols.map((c) => c.label)];
    const imageCellMap = new Map();
    let hasEmbeddedImages = false;

    const body = state.filteredRows.map((row, rowIdx) =>
      activeCols.map((col, colIdx) => {
        const val = row[col.key] || '';
        if (col.isPhoto && isDataImageUrl(val)) {
          imageCellMap.set(`${rowIdx},${colIdx}`, val);
          hasEmbeddedImages = true;
          return '';
        }
        return val || '-';
      })
    );

    // Scale column widths proportionally to fit A4 landscape printable width (~273mm)
    const printableWidth = pageWidth - marginX * 2;
    const totalBaseWidth = activeCols.reduce((sum, c) => sum + (c.pdfWidth || 24), 0);
    const scale = totalBaseWidth > 0 ? printableWidth / totalBaseWidth : 1;

    const columnStyles = {};
    activeCols.forEach((col, idx) => {
      columnStyles[idx] = {
        cellWidth: Number(((col.pdfWidth || 24) * scale).toFixed(2)),
      };
      if (col.key === 'No') {
        columnStyles[idx].halign = 'center';
      }
    });

    doc.autoTable({
      head,
      body,
      startY: 28,
      margin: { left: marginX, right: marginX, bottom: 14 },
      theme: 'grid',
      styles: {
        font: 'helvetica',
        fontSize: 7.5,
        cellPadding: 2,
        overflow: 'linebreak',
        valign: 'top',
        textColor: [30, 41, 59],
        lineColor: [226, 232, 240],
        lineWidth: 0.2,
      },
      bodyStyles: hasEmbeddedImages ? { minCellHeight: 18 } : {},
      headStyles: {
        fillColor: [30, 41, 59],
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        valign: 'middle',
      },
      alternateRowStyles: {
        fillColor: [248, 250, 252],
      },
      columnStyles,
      didDrawCell: (data) => {
        if (data.section !== 'body') return;
        const imgDataUrl = imageCellMap.get(`${data.row.index},${data.column.index}`);
        if (!imgDataUrl) return;

        const pad = 1.5;
        const availW = data.cell.width - pad * 2;
        const availH = data.cell.height - pad * 2;
        const boxSize = Math.min(availW, availH, 16);
        if (boxSize <= 0) return;

        const x = data.cell.x + (data.cell.width - boxSize) / 2;
        const y = data.cell.y + (data.cell.height - boxSize) / 2;
        const fmt = getImageFormatFromDataUrl(imgDataUrl);
        try {
          doc.addImage(imgDataUrl, fmt, x, y, boxSize, boxSize);
        } catch (err) {
          console.warn('Gagal menggambar foto di sel PDF:', err);
        }
      },
      didDrawPage: (data) => {
        const pageHeight = doc.internal.pageSize.getHeight();
        const pageNumber = doc.internal.getCurrentPageInfo().pageNumber;
        const footerStr = `Halaman ${pageNumber} dari ${totalPagesExp}`;

        doc.setFont('helvetica', 'normal');
        doc.setFontSize(8);
        doc.setTextColor(100, 116, 139);
        doc.text('LSP Pengadaan — Attendance Report', marginX, pageHeight - 6);
        doc.text(footerStr, pageWidth - marginX, pageHeight - 6, { align: 'right' });
      },
    });

    if (typeof doc.putTotalPages === 'function') {
      doc.putTotalPages(totalPagesExp);
    }

    const outName = `Laporan_Presensi_LSP_Pengadaan_${getTodayStamp()}.pdf`;
    doc.save(outName);
  }

  // ---------------------------------------------------------------------------
  // Event Listeners
  // ---------------------------------------------------------------------------
  function initEvents() {
    // Click to open file dialog
    els.dropZone.addEventListener('click', () => els.csvFileInput.click());
    els.dropZone.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        els.csvFileInput.click();
      }
    });

    els.csvFileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files[0]) {
        handleFileObject(e.target.files[0]);
        e.target.value = '';
      }
    });

    // Drag and drop events
    ['dragenter', 'dragover'].forEach((evtName) => {
      els.dropZone.addEventListener(evtName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        els.dropZone.classList.add('border-indigo-500', 'bg-indigo-50/70');
      });
    });

    ['dragleave', 'drop'].forEach((evtName) => {
      els.dropZone.addEventListener(evtName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        els.dropZone.classList.remove('border-indigo-500', 'bg-indigo-50/70');
      });
    });

    els.dropZone.addEventListener('drop', (e) => {
      const files = e.dataTransfer?.files;
      if (files && files[0]) {
        handleFileObject(files[0]);
      }
    });

    // Load Samples
    els.btnLoadSample.addEventListener('click', loadSampleCsv);
    els.btnLoadSampleExcel.addEventListener('click', loadSampleExcel);

    // Image preview modal
    els.btnCloseImageModal.addEventListener('click', closeImageModal);
    els.imageModal.addEventListener('click', (e) => {
      if (e.target === els.imageModal) closeImageModal();
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && !els.imageModal.classList.contains('hidden')) {
        closeImageModal();
      }
    });

    // Close alert
    els.btnCloseAlert.addEventListener('click', hideAlert);

    // Filters
    [els.filterEmployee, els.filterStatus, els.filterDateFrom, els.filterDateTo].forEach((el) => {
      el.addEventListener('change', applyFiltersAndRender);
    });
    els.filterKeyword.addEventListener('input', applyFiltersAndRender);

    els.btnResetFilters.addEventListener('click', () => {
      resetFilterInputs();
      applyFiltersAndRender();
    });

    // Column Visibility Quick Actions
    els.btnHideTechnical.addEventListener('click', () => {
      state.columns.forEach((col) => {
        if (col.technical) {
          state.visibleColumns.delete(col.key);
        } else {
          state.visibleColumns.add(col.key);
        }
      });
      renderColumnToggles();
      renderTable();
    });

    els.btnShowAllCols.addEventListener('click', () => {
      state.columns.forEach((col) => state.visibleColumns.add(col.key));
      renderColumnToggles();
      renderTable();
    });

    // Exports
    els.btnExportExcel.addEventListener('click', exportToExcel);
    els.btnExportPdf.addEventListener('click', exportToPdf);
  }

  initEvents();
})();
