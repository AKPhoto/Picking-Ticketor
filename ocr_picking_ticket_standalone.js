(function () {
  const state = {
    pdfDoc: null,
    page: null,
    viewport: null,
    scale: 1.6,
    selection: null,
    ocrAreaSelection: null,
    columnSelections: {
      pointNumber: null,
      itemCode: null,
      size: null,
      quantity: null
    },
    selectionMode: 'ocrArea',
    activeColumnKey: null,
    isPanning: false,
    panStart: null,
    panScrollStart: null,
    isDragging: false,
    dragStart: null,
    learningStore: null,
    ticketNumbersReserved: false,
    currentRenderedTicketCount: 0,
    lastPreviewPayload: null,
    lastSavePickerHandle: null,
    logoDataUrl: '',
    pdfQueueFiles: [],
    pdfQueueIndex: -1,
    activePdfFileName: '',
    selectedTableRow: null,
    itemCodeDescriptionCatalog: {},
    itemCodeDescriptionCatalogLoaded: false,
    userItemCodeCatalog: {},
    craftByItemCodeStore: {},
    deletedItemCodeSet: {},
    pendingTicketCommit: null,
    hasGeneratedCurrentDrawing: false
  };
  const OCR_LEARNING_STORAGE_KEY = 'matcon_ocr_learning_v1';
  const USER_ITEM_CODE_STORAGE_KEY = 'matcon_item_code_user_catalog_v1';
  const BASE_ITEM_CODE_CACHE_STORAGE_KEY = 'matcon_item_code_base_catalog_v1';
  const CRAFT_BY_ITEM_CODE_STORAGE_KEY = 'matcon_craft_by_item_code_v1';
  const DELETED_ITEM_CODE_STORAGE_KEY = 'matcon_item_code_deleted_catalog_v1';
  const LEARNING_FIELDS = ['pointNumber', 'itemCode', 'size', 'quantity'];
  const PLACEHOLDER_ITEM_CODE = 'PLACEHOLDER-CODE';
  const PLACEHOLDER_DESCRIPTION_TEXT = 'UNRECOGNIZED ITEM CODE - DESCRIPTION REQUIRED';
  const ITEM_CODE_DESCRIPTION_FILE = 'item_code_descriptions.json';

  const SHEET_ORDER = ['Pipe', 'Fabrication', 'Erection', 'Supports'];

  const elements = {
    projectNo: document.getElementById('projectNo'),
    ticketStartNo: document.getElementById('ticketStartNo'),
    workpackNo: document.getElementById('workpackNo'),
    drawingNumber: document.getElementById('drawingNumber'),
    operatingTemperature: document.getElementById('operatingTemperature'),
    insulationType: document.getElementById('insulationType'),
    materialType: document.getElementById('materialType'),
    sheetNo: document.getElementById('sheetNo'),
    revision: document.getElementById('revision'),
    ocrLanguage: document.getElementById('ocrLanguage'),
    pdfInput: document.getElementById('pdfInput'),
    pdfFolderInput: document.getElementById('pdfFolderInput'),
    renderBtn: document.getElementById('renderBtn'),
    loadFolderBtn: document.getElementById('loadFolderBtn'),
    queueStatusBadge: document.getElementById('queueStatusBadge'),
    runOcrBtn: document.getElementById('runOcrBtn'),
    itemCount: document.getElementById('itemCount'),
    lockCanvasScroll: document.getElementById('lockCanvasScroll'),
    selectOcrAreaBtn: document.getElementById('selectOcrAreaBtn'),
    resetOcrAreaBtn: document.getElementById('resetOcrAreaBtn'),
    selectPointColumnBtn: document.getElementById('selectPointColumnBtn'),
    selectItemCodeColumnBtn: document.getElementById('selectItemCodeColumnBtn'),
    selectSizeColumnBtn: document.getElementById('selectSizeColumnBtn'),
    selectQuantityColumnBtn: document.getElementById('selectQuantityColumnBtn'),
    pointColumnStatus: document.getElementById('pointColumnStatus'),
    ocrAreaStatus: document.getElementById('ocrAreaStatus'),
    itemCodeColumnStatus: document.getElementById('itemCodeColumnStatus'),
    sizeColumnStatus: document.getElementById('sizeColumnStatus'),
    quantityColumnStatus: document.getElementById('quantityColumnStatus'),
    addRowBtn: document.getElementById('addRowBtn'),
    deleteRowBtn: document.getElementById('deleteRowBtn'),
    generateBtn: document.getElementById('generateBtn'),
    drawingProcessedBtn: document.getElementById('drawingProcessedBtn'),
    exportPdfBtn: document.getElementById('exportPdfBtn'),
    reprintBtn: document.getElementById('reprintBtn'),
    manageItemCodesBtn: document.getElementById('manageItemCodesBtn'),
    itemCodeFlyout: document.getElementById('itemCodeFlyout'),
    itemCodeFlyoutOverlay: document.getElementById('itemCodeFlyoutOverlay'),
    closeItemCodeFlyoutBtn: document.getElementById('closeItemCodeFlyoutBtn'),
    itemCodeInput: document.getElementById('itemCodeInput'),
    itemDescriptionInput: document.getElementById('itemDescriptionInput'),
    addItemCodeBtn: document.getElementById('addItemCodeBtn'),
    itemCodeSearchInput: document.getElementById('itemCodeSearchInput'),
    itemCodeCountLabel: document.getElementById('itemCodeCountLabel'),
    itemCodeCatalogTableBody: document.getElementById('itemCodeCatalogTableBody'),
    previewCraftSelect: document.getElementById('previewCraftSelect'),
    openPreviewWindowBtn: document.getElementById('openPreviewWindowBtn'),
    ticketPreviewContainer: document.getElementById('ticketPreviewContainer'),
    status: document.getElementById('status'),
    viewerWrap: document.getElementById('viewerWrap'),
    pdfCanvas: document.getElementById('pdfCanvas'),
    selectionOverlay: document.getElementById('selectionOverlay'),
    ocrTableBody: document.querySelector('#ocrTable tbody')
  };

  const ctx = elements.pdfCanvas.getContext('2d');
  if (window.pdfjsLib) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('./pdf.worker.min.js?v=20260304', window.location.href).toString();
  }

  function setStatus(message, isError) {
    elements.status.textContent = message;
    elements.status.style.color = isError ? '#b91c1c' : '#334155';
  }

  function canvasToBlob(canvas, type, quality) {
    return new Promise((resolve, reject) => {
      if (!canvas || typeof canvas.toBlob !== 'function') {
        reject(new Error('Canvas toBlob is unavailable.'));
        return;
      }

      canvas.toBlob((blob) => {
        if (!blob) {
          reject(new Error('Canvas export failed.'));
          return;
        }
        resolve(blob);
      }, type || 'image/png', quality);
    });
  }

  async function loadLogoDataUrl() {
    if (state.logoDataUrl) return state.logoDataUrl;
    if (typeof window.AUREX_LOGO_DATA_URL === 'string' && window.AUREX_LOGO_DATA_URL.startsWith('data:image/')) {
      state.logoDataUrl = window.AUREX_LOGO_DATA_URL;
      return state.logoDataUrl;
    }

    try {
      const response = await fetch('./Aurex%20Logo.jpg', { cache: 'no-store' });
      if (!response.ok) return '';
      const blob = await response.blob();
      const dataUrl = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result || ''));
        reader.onerror = () => reject(new Error('Failed to read logo file.'));
        reader.readAsDataURL(blob);
      });

      state.logoDataUrl = String(dataUrl || '');
      return state.logoDataUrl;
    } catch (error) {
      console.warn('Logo preload failed, using file path fallback.', error);
      return '';
    }
  }

  async function waitForImagesToLoad(container, timeoutMs) {
    const root = container instanceof HTMLElement ? container : null;
    if (!root) return;

    const images = Array.from(root.querySelectorAll('img'));
    if (!images.length) return;

    const timeout = Math.max(500, Number(timeoutMs) || 5000);

    await Promise.all(images.map((img) => new Promise((resolve) => {
      if (img.complete && img.naturalWidth > 0) {
        resolve();
        return;
      }

      let settled = false;
      const finish = () => {
        if (settled) return;
        settled = true;
        img.removeEventListener('load', onLoad);
        img.removeEventListener('error', onError);
        clearTimeout(timer);
        resolve();
      };

      const onLoad = () => finish();
      const onError = () => finish();
      const timer = setTimeout(finish, timeout);

      img.addEventListener('load', onLoad, { once: true });
      img.addEventListener('error', onError, { once: true });
    })));
  }

  function normalizeCatalogCodes(rawCodes) {
    const source = rawCodes && typeof rawCodes === 'object' ? rawCodes : {};
    const normalized = {};

    Object.entries(source).forEach(([key, value]) => {
      const normalizedKey = String(key || '').trim().toUpperCase();
      const description = String(value || '').trim();
      if (!normalizedKey || !description) return;
      normalized[normalizedKey] = description;
    });

    return normalized;
  }

  function loadDeletedItemCodeSet() {
    try {
      const raw = localStorage.getItem(DELETED_ITEM_CODE_STORAGE_KEY);
      if (!raw) return {};
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== 'object') return {};
      const normalized = {};
      Object.entries(parsed).forEach(([key, value]) => {
        const code = String(key || '').trim().toUpperCase();
        if (!code || !value) return;
        normalized[code] = true;
      });
      return normalized;
    } catch (error) {
      console.warn('Failed to load deleted item codes.', error);
      return {};
    }
  }

  function saveDeletedItemCodeSet() {
    localStorage.setItem(DELETED_ITEM_CODE_STORAGE_KEY, JSON.stringify(state.deletedItemCodeSet || {}));
  }

  function applyDeletedCodesFilter(catalog) {
    const source = normalizeCatalogCodes(catalog);
    const deletedSet = state.deletedItemCodeSet || {};
    const filtered = {};
    Object.entries(source).forEach(([code, description]) => {
      if (deletedSet[code]) return;
      filtered[code] = description;
    });
    return filtered;
  }

  function loadBaseCatalogFromCache() {
    try {
      const raw = localStorage.getItem(BASE_ITEM_CODE_CACHE_STORAGE_KEY);
      if (!raw) return {};
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== 'object') return {};
      return normalizeCatalogCodes(parsed);
    } catch (error) {
      console.warn('Failed to load cached base item catalog.', error);
      return {};
    }
  }

  function saveBaseCatalogToCache(catalog) {
    try {
      const normalized = normalizeCatalogCodes(catalog);
      localStorage.setItem(BASE_ITEM_CODE_CACHE_STORAGE_KEY, JSON.stringify(normalized));
    } catch (error) {
      console.warn('Failed to cache base item catalog.', error);
    }
  }

  function promptForCatalogFileSelection() {
    return new Promise((resolve, reject) => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = 'application/json,.json';
      input.style.display = 'none';
      document.body.appendChild(input);

      const cleanup = () => {
        input.removeEventListener('change', onChange);
        if (input.parentNode) {
          input.parentNode.removeChild(input);
        }
      };

      const onChange = async () => {
        try {
          const file = input.files && input.files[0];
          if (!file) {
            resolve(null);
            return;
          }

          const text = await file.text();
          const parsed = JSON.parse(text);
          const codes = parsed?.codes && typeof parsed.codes === 'object' ? parsed.codes : parsed;
          const normalized = normalizeCatalogCodes(codes);
          resolve(normalized);
        } catch (error) {
          reject(error);
        } finally {
          cleanup();
        }
      };

      input.addEventListener('change', onChange, { once: true });
      input.click();
    });
  }

  async function ensureCatalogAvailableForGeneration() {
    if (Object.keys(state.itemCodeDescriptionCatalog || {}).length > 0) {
      return true;
    }

    if (window.location.protocol !== 'file:') {
      return true;
    }

    const shouldImport = window.confirm(
      'Running from file:// blocks auto-loading item_code_descriptions.json.\n\nSelect item_code_descriptions.json now?'
    );
    if (!shouldImport) {
      setStatus('Catalog not loaded. Use the launcher script (http://127.0.0.1) or import the JSON catalog to continue.', true);
      return false;
    }

    try {
      const importedCatalog = await promptForCatalogFileSelection();
      if (!importedCatalog || Object.keys(importedCatalog).length === 0) {
        setStatus('No valid catalog selected. Export canceled.', true);
        return false;
      }

      saveBaseCatalogToCache(importedCatalog);
      state.itemCodeDescriptionCatalog = applyDeletedCodesFilter({
        ...importedCatalog,
        ...normalizeCatalogCodes(state.userItemCodeCatalog || {})
      });
      state.itemCodeDescriptionCatalogLoaded = true;
      setStatus(`Catalog imported (${Object.keys(importedCatalog).length} codes).`, false);
      return true;
    } catch (error) {
      console.warn('Catalog import failed.', error);
      setStatus('Catalog import failed. Please select a valid JSON file.', true);
      return false;
    }
  }

  async function loadItemCodeDescriptionCatalog() {
    if (state.itemCodeDescriptionCatalogLoaded) {
      return state.itemCodeDescriptionCatalog;
    }

    try {
      state.deletedItemCodeSet = loadDeletedItemCodeSet();
      const userCatalog = loadUserItemCodeCatalog();
      state.userItemCodeCatalog = { ...userCatalog };

      let baseCatalog = loadBaseCatalogFromCache();

      if (window.location.protocol !== 'file:') {
        const response = await fetch(ITEM_CODE_DESCRIPTION_FILE, { cache: 'no-store' });
        if (!response.ok) {
          throw new Error(`HTTP ${response.status}`);
        }

        const data = await response.json();
        baseCatalog = normalizeCatalogCodes(data?.codes || {});
        if (Object.keys(baseCatalog).length > 0) {
          saveBaseCatalogToCache(baseCatalog);
        }
      }

      state.itemCodeDescriptionCatalog = applyDeletedCodesFilter({
        ...baseCatalog,
        ...normalizeCatalogCodes(userCatalog)
      });
      state.itemCodeDescriptionCatalogLoaded = true;
      return state.itemCodeDescriptionCatalog;
    } catch (error) {
      console.warn('Failed to load item code description catalog.', error);
      state.itemCodeDescriptionCatalog = {};
      state.deletedItemCodeSet = loadDeletedItemCodeSet();
      state.userItemCodeCatalog = loadUserItemCodeCatalog();
      Object.entries(state.userItemCodeCatalog).forEach(([key, value]) => {
        const normalizedKey = String(key || '').trim().toUpperCase();
        const description = String(value || '').trim();
        if (!normalizedKey || !description) return;
        if (state.deletedItemCodeSet[normalizedKey]) return;
        state.itemCodeDescriptionCatalog[normalizedKey] = description;
      });
      state.itemCodeDescriptionCatalogLoaded = true;
      return state.itemCodeDescriptionCatalog;
    }
  }

  function lookupDescriptionInCatalog(itemCode) {
    const catalog = state.itemCodeDescriptionCatalog || {};
    const candidates = getItemCodeCandidates(itemCode);
    for (const candidate of candidates) {
      if (catalog[candidate]) return catalog[candidate];
    }
    return '';
  }

  function loadUserItemCodeCatalog() {
    try {
      const raw = localStorage.getItem(USER_ITEM_CODE_STORAGE_KEY);
      if (!raw) return {};
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== 'object') return {};

      const normalized = {};
      Object.entries(parsed).forEach(([key, value]) => {
        const code = String(key || '').trim().toUpperCase();
        const description = String(value || '').trim();
        if (!code || !description) return;
        normalized[code] = description;
      });
      return normalized;
    } catch (error) {
      console.warn('Failed to load user item code catalog.', error);
      return {};
    }
  }

  function saveUserItemCodeCatalog() {
    localStorage.setItem(USER_ITEM_CODE_STORAGE_KEY, JSON.stringify(state.userItemCodeCatalog || {}));
  }

  function addCodeToCatalog(code, description) {
    const cleanDescription = String(description || '').trim();
    const cleanCode = String(code || '').trim().toUpperCase();
    if (!cleanCode || !cleanDescription) return;

    state.userItemCodeCatalog[cleanCode] = cleanDescription;
    if (state.deletedItemCodeSet[cleanCode]) {
      delete state.deletedItemCodeSet[cleanCode];
      saveDeletedItemCodeSet();
    }

    const candidates = getItemCodeCandidates(cleanCode);
    candidates.forEach((candidate) => {
      state.itemCodeDescriptionCatalog[candidate] = cleanDescription;
    });

    saveUserItemCodeCatalog();
  }

  function deleteCodeFromCatalog(code) {
    const cleanCode = String(code || '').trim().toUpperCase();
    if (!cleanCode) return;
    delete state.userItemCodeCatalog[cleanCode];
    delete state.itemCodeDescriptionCatalog[cleanCode];
    state.deletedItemCodeSet[cleanCode] = true;
    saveUserItemCodeCatalog();
    saveDeletedItemCodeSet();
  }

  function getCatalogRowsForFlyout() {
    const entries = Object.entries(state.itemCodeDescriptionCatalog || {});
    return entries
      .map(([code, description]) => ({ code, description }))
      .sort((a, b) => a.code.localeCompare(b.code, undefined, { sensitivity: 'base', numeric: true }));
  }

  function renderItemCodeFlyoutRows() {
    if (!elements.itemCodeCatalogTableBody) return;

    const term = String(elements.itemCodeSearchInput?.value || '').trim().toUpperCase();
    const allRows = getCatalogRowsForFlyout();
    const rows = term
      ? allRows.filter((item) => item.code.includes(term) || String(item.description || '').toUpperCase().includes(term))
      : allRows;

    elements.itemCodeCatalogTableBody.innerHTML = rows.map((item) => `
      <tr data-code="${escapeHtml(item.code)}">
        <td>${escapeHtml(item.code)}</td>
        <td>${escapeHtml(item.description || '')}</td>
        <td>
          <button class="secondary" data-action="edit" data-code="${escapeHtml(item.code)}">Edit</button>
          <button class="secondary" data-action="delete" data-code="${escapeHtml(item.code)}">Delete</button>
        </td>
      </tr>
    `).join('');

    if (elements.itemCodeCountLabel) {
      elements.itemCodeCountLabel.textContent = `Codes: ${rows.length}`;
    }
  }

  function openItemCodeFlyout() {
    if (!elements.itemCodeFlyout || !elements.itemCodeFlyoutOverlay) return;
    renderItemCodeFlyoutRows();
    elements.itemCodeFlyout.classList.add('open');
    elements.itemCodeFlyoutOverlay.classList.add('open');
    elements.itemCodeFlyout.setAttribute('aria-hidden', 'false');
  }

  function closeItemCodeFlyout() {
    if (!elements.itemCodeFlyout || !elements.itemCodeFlyoutOverlay) return;
    elements.itemCodeFlyout.classList.remove('open');
    elements.itemCodeFlyoutOverlay.classList.remove('open');
    elements.itemCodeFlyout.setAttribute('aria-hidden', 'true');
  }

  function upsertItemCodeFromFlyout() {
    const code = String(elements.itemCodeInput?.value || '').trim().toUpperCase();
    const description = String(elements.itemDescriptionInput?.value || '').trim();
    if (!code || !description) {
      setStatus('Enter both item code and description before adding/updating.', true);
      return;
    }

    addCodeToCatalog(code, description);
    state.itemCodeDescriptionCatalog[code] = description;
    if (state.deletedItemCodeSet[code]) {
      delete state.deletedItemCodeSet[code];
      saveDeletedItemCodeSet();
    }
    renderItemCodeFlyoutRows();
    elements.itemCodeInput.value = '';
    elements.itemDescriptionInput.value = '';
    Array.from(elements.ocrTableBody.querySelectorAll('tr')).forEach((row) => updateMaterialDescriptionForTableRow(row));
    setStatus(`Item code ${code} saved in local catalog.`, false);
  }

  function handleItemCodeFlyoutTableClick(target) {
    if (!(target instanceof HTMLElement)) return;
    const action = target.dataset.action;
    const code = String(target.dataset.code || '').trim().toUpperCase();
    if (!action || !code) return;

    if (action === 'edit') {
      const currentDescription = state.itemCodeDescriptionCatalog[code] || '';
      if (elements.itemCodeInput) elements.itemCodeInput.value = code;
      if (elements.itemDescriptionInput) elements.itemDescriptionInput.value = currentDescription;
      elements.itemDescriptionInput?.focus();
      return;
    }

    if (action === 'delete') {
      const confirmDelete = window.confirm(`Delete item code ${code} from this browser catalog?`);
      if (!confirmDelete) return;
      deleteCodeFromCatalog(code);
      renderItemCodeFlyoutRows();
      Array.from(elements.ocrTableBody.querySelectorAll('tr')).forEach((row) => updateMaterialDescriptionForTableRow(row));
      setStatus(`Item code ${code} deleted from local catalog.`, false);
    }
  }

  function showDrawingProcessedButton(show) {
    if (!elements.drawingProcessedBtn) return;
    elements.drawingProcessedBtn.style.display = show ? 'inline-flex' : 'none';
    elements.drawingProcessedBtn.disabled = !show;
  }

  function promptToAddMissingItemCodes(missingCodes) {
    const codes = Array.isArray(missingCodes) ? missingCodes : [];
    if (!codes.length) return { added: 0, skipped: 0 };

    let added = 0;
    let skipped = 0;

    codes.forEach((code) => {
      const shouldAdd = window.confirm(`Item code ${code} is not in the catalog. Do you want to add it now?`);
      if (!shouldAdd) {
        skipped += 1;
        return;
      }

      const description = window.prompt(`Enter material description for ${code}:`, '');
      const cleanDescription = String(description || '').trim();
      if (!cleanDescription) {
        skipped += 1;
        return;
      }

      addCodeToCatalog(code, cleanDescription);
      added += 1;
    });

    return { added, skipped };
  }

  function hasCatalogMatch(itemCode) {
    const catalog = state.itemCodeDescriptionCatalog || {};
    const candidates = getItemCodeCandidates(itemCode);
    return candidates.some((candidate) => Boolean(catalog[candidate]));
  }

  function findMissingCatalogItemCodes(rows) {
    const missing = new Set();
    rows.forEach((row) => {
      const code = String(row.itemCode || '').trim();
      if (!code || isPlaceholderItemCode(code)) return;
      if (!hasCatalogMatch(code)) {
        missing.add(code.toUpperCase());
      }
    });
    return Array.from(missing).sort();
  }

  function createEmptyLearningStore() {
    return {
      pointNumber: {},
      itemCode: {},
      size: {},
      quantity: {}
    };
  }

  function normalizeLearningKey(value) {
    return String(value || '').trim().toUpperCase();
  }

  function loadLearningStore() {
    try {
      const raw = localStorage.getItem(OCR_LEARNING_STORAGE_KEY);
      if (!raw) return createEmptyLearningStore();
      const parsed = JSON.parse(raw);
      return {
        pointNumber: parsed.pointNumber || {},
        itemCode: parsed.itemCode || {},
        size: parsed.size || {},
        quantity: parsed.quantity || {}
      };
    } catch (error) {
      console.warn('Failed to load OCR learning store. Resetting.', error);
      return createEmptyLearningStore();
    }
  }

  function saveLearningStore() {
    localStorage.setItem(OCR_LEARNING_STORAGE_KEY, JSON.stringify(state.learningStore || createEmptyLearningStore()));
  }

  function applyLearnedCorrections(row) {
    let applied = 0;
    LEARNING_FIELDS.forEach((field) => {
      const rawValue = String(row[field] || '').trim();
      if (!rawValue) return;

      const key = normalizeLearningKey(rawValue);
      const learned = state.learningStore?.[field]?.[key];
      if (!learned || typeof learned.value !== 'string') return;
      if (field === 'itemCode' && isPlaceholderItemCode(learned.value)) return;

      if (learned.value !== rawValue) {
        row[field] = learned.value;
        applied += 1;
      }
    });
    return applied;
  }

  const COLUMN_KEYS = ['pointNumber', 'itemCode', 'size', 'quantity'];
  const COLUMN_LABELS = {
    pointNumber: 'Point Number',
    itemCode: 'Item Code',
    size: 'Size',
    quantity: 'Quantity'
  };
  const SELECTION_COLORS = {
    ocrArea: '#2563eb',
    pointNumber: '#16a34a',
    itemCode: '#ea580c',
    size: '#7c3aed',
    quantity: '#dc2626'
  };

  function normalizePointNumberValue(value) {
    return String(value || '').replace(/\D+/g, '');
  }

  function getSelectionForKey(key) {
    if (key === 'ocrArea') return state.ocrAreaSelection;
    return state.columnSelections[key] || null;
  }

  function removeSelectionHighlights() {
    elements.viewerWrap.querySelectorAll('.selection-highlight').forEach((item) => item.remove());
  }

  function renderSelectionHighlights() {
    removeSelectionHighlights();
    if (!state.viewport) return;

    const keys = ['ocrArea', ...COLUMN_KEYS];
    keys.forEach((key) => {
      const selection = getSelectionForKey(key);
      if (!selection) return;

      const highlight = document.createElement('div');
      highlight.className = 'selection-highlight';
      highlight.style.left = `${selection.x}px`;
      highlight.style.top = `${selection.y}px`;
      highlight.style.width = `${selection.width}px`;
      highlight.style.height = `${selection.height}px`;
      highlight.style.borderColor = SELECTION_COLORS[key] || '#2563eb';
      highlight.style.backgroundColor = `${(SELECTION_COLORS[key] || '#2563eb')}22`;
      highlight.title = key === 'ocrArea' ? 'OCR Area' : COLUMN_LABELS[key];
      elements.viewerWrap.appendChild(highlight);
    });
  }

  function scaleSelection(selection, ratio) {
    if (!selection) return null;
    return {
      x: Math.round(selection.x * ratio),
      y: Math.round(selection.y * ratio),
      width: Math.round(selection.width * ratio),
      height: Math.round(selection.height * ratio)
    };
  }

  function getSelectedColumnsCount() {
    return COLUMN_KEYS.filter((key) => state.columnSelections[key]).length;
  }

  function resetColumnSelections() {
    state.columnSelections = {
      pointNumber: null,
      itemCode: null,
      size: null,
      quantity: null
    };
    state.activeColumnKey = null;
  }

  function setColumnButtonsEnabled(enabled) {
    [elements.selectPointColumnBtn, elements.selectItemCodeColumnBtn, elements.selectSizeColumnBtn, elements.selectQuantityColumnBtn]
      .forEach((button) => {
        if (!button) return;
        button.disabled = !enabled;
      });
  }

  function updateRunOcrAvailability() {
    const itemCount = parseInt(elements.itemCount?.value || '0', 10);
    const hasAllColumns = getSelectedColumnsCount() === COLUMN_KEYS.length;
    elements.runOcrBtn.disabled = !(state.page && state.ocrAreaSelection && hasAllColumns && itemCount > 0);
  }

  function updateColumnStatus() {
    const areaSelected = Boolean(state.ocrAreaSelection);
    elements.ocrAreaStatus.textContent = `OCR Area: ${areaSelected ? 'Selected' : 'Not selected'}`;
    elements.ocrAreaStatus.style.color = areaSelected ? SELECTION_COLORS.ocrArea : '#334155';

    const map = [
      ['pointNumber', elements.pointColumnStatus, 'Point'],
      ['itemCode', elements.itemCodeColumnStatus, 'Item Code'],
      ['size', elements.sizeColumnStatus, 'Size'],
      ['quantity', elements.quantityColumnStatus, 'Quantity']
    ];

    map.forEach(([key, el, label]) => {
      const selected = Boolean(state.columnSelections[key]);
      el.textContent = `${label}: ${selected ? 'Selected' : 'Not selected'}`;
      el.style.color = selected ? (SELECTION_COLORS[key] || '#166534') : '#334155';
    });

    updateRunOcrAvailability();
  }

  function setActiveColumnSelection(columnKey) {
    if (!state.ocrAreaSelection) {
      setStatus('Select OCR Area first, then choose column buttons.', true);
      return;
    }

    state.selectionMode = 'column';
    state.activeColumnKey = columnKey;

    const buttonMap = {
      ocrArea: elements.selectOcrAreaBtn,
      pointNumber: elements.selectPointColumnBtn,
      itemCode: elements.selectItemCodeColumnBtn,
      size: elements.selectSizeColumnBtn,
      quantity: elements.selectQuantityColumnBtn
    };

    Object.entries(buttonMap).forEach(([key, button]) => {
      if (!button) return;
      const active = key === columnKey;
      button.classList.toggle('active', active);
      button.classList.toggle('secondary', !active);
    });

    setStatus(`Drag on the PDF to select the ${COLUMN_LABELS[columnKey]} column.`, false);
  }

  function activateOcrAreaSelectionMode() {
    state.selectionMode = 'ocrArea';
    state.activeColumnKey = null;

    const buttonMap = [
      elements.selectOcrAreaBtn,
      elements.selectPointColumnBtn,
      elements.selectItemCodeColumnBtn,
      elements.selectSizeColumnBtn,
      elements.selectQuantityColumnBtn
    ];

    buttonMap.forEach((button, index) => {
      if (!button) return;
      const active = index === 0;
      button.classList.toggle('active', active);
      button.classList.toggle('secondary', !active);
    });

    setStatus('Drag on the PDF to select the full OCR area. The view will auto-zoom to this area.', false);
  }

  async function renderPageAtScale(scale) {
    if (!state.page) return;
    const previousScale = state.scale || scale;
    state.scale = scale;
    state.viewport = state.page.getViewport({ scale: state.scale });
    elements.pdfCanvas.width = state.viewport.width;
    elements.pdfCanvas.height = state.viewport.height;
    await state.page.render({ canvasContext: ctx, viewport: state.viewport }).promise;

    const ratio = previousScale > 0 ? (state.scale / previousScale) : 1;
    if (ratio !== 1) {
      state.ocrAreaSelection = scaleSelection(state.ocrAreaSelection, ratio);
      COLUMN_KEYS.forEach((key) => {
        state.columnSelections[key] = scaleSelection(state.columnSelections[key], ratio);
      });
      state.selection = scaleSelection(state.selection, ratio);
    }

    renderSelectionHighlights();
  }

  async function resetOcrAreaAndZoom() {
    if (!state.page) {
      setStatus('Load a PDF first.', true);
      return;
    }

    await renderPageAtScale(1.6);
    state.selection = null;
    state.ocrAreaSelection = null;
    resetColumnSelections();
    setColumnButtonsEnabled(false);
    activateOcrAreaSelectionMode();
    elements.selectionOverlay.style.display = 'none';
    removeSelectionHighlights();
    elements.viewerWrap.scrollLeft = 0;
    elements.viewerWrap.scrollTop = 0;
    updateColumnStatus();
    setStatus('Reset complete. Select OCR Area again and drag the correct region.', false);
  }

  async function zoomToSelection(selection) {
    const currentScale = state.scale;
    const viewerWidth = Math.max(1, elements.viewerWrap.clientWidth || 1);
    const viewerHeight = Math.max(1, elements.viewerWrap.clientHeight || 1);
    const padding = 24;
    const fitWidth = Math.max(1, viewerWidth - (padding * 2));
    const fitHeight = Math.max(1, viewerHeight - (padding * 2));

    const scaleByWidth = fitWidth / Math.max(1, selection.width);
    const scaleByHeight = fitHeight / Math.max(1, selection.height);
    const fitScaleOnCurrentCanvas = Math.min(scaleByWidth, scaleByHeight);

    const targetScale = Math.min(6, Math.max(0.4, currentScale * fitScaleOnCurrentCanvas));

    if (Math.abs(targetScale - currentScale) > 0.05) {
      await renderPageAtScale(targetScale);
    }

    const targetX = selection.x;
    const targetY = selection.y;
    const targetW = selection.width;
    const targetH = selection.height;

    const centerX = targetX + (targetW / 2);
    const centerY = targetY + (targetH / 2);

    const desiredScrollLeft = centerX - (viewerWidth / 2);
    const desiredScrollTop = centerY - (viewerHeight / 2);

    const maxScrollLeft = Math.max(0, elements.pdfCanvas.width - viewerWidth);
    const maxScrollTop = Math.max(0, elements.pdfCanvas.height - viewerHeight);

    await new Promise((resolve) => {
      window.requestAnimationFrame(() => {
        window.requestAnimationFrame(resolve);
      });
    });

    elements.viewerWrap.scrollLeft = Math.max(0, Math.min(maxScrollLeft, desiredScrollLeft));
    elements.viewerWrap.scrollTop = Math.max(0, Math.min(maxScrollTop, desiredScrollTop));
  }

  function getDrawingNumber() {
    return String(elements.drawingNumber?.value || '').trim().toUpperCase();
  }

  function createTypeCellHtml(nameBase, selected) {
    const types = [
      { value: 'P', initial: 'P' },
      { value: 'WF', initial: 'WF' },
      { value: 'SF', initial: 'SF' },
      { value: 'E', initial: 'E' },
      { value: 'S', initial: 'S' },
      { value: 'EI', initial: 'EI' }
    ];

    const options = types
      .map((type) => {
        const checked = selected === type.value ? 'checked' : '';
        return `<label class="typeOption"><span>${type.initial}</span><input type="radio" name="${nameBase}" value="${type.value}" ${checked}></label>`;
      })
      .join('');

    return `<div class="typeOptions">${options}</div>`;
  }

  function ensureTypeNames() {
    const rows = Array.from(elements.ocrTableBody.querySelectorAll('tr'));
    rows.forEach((row, index) => {
      row.dataset.rowIndex = String(index);
      const radios = row.querySelectorAll('input[type="radio"]');
      radios.forEach((radio) => {
        radio.name = `type-${index}`;
      });
    });
  }

  function setSelectedTableRow(row) {
    elements.ocrTableBody.querySelectorAll('tr.row-selected').forEach((item) => {
      item.classList.remove('row-selected');
    });

    if (!row) {
      state.selectedTableRow = null;
      if (elements.deleteRowBtn) {
        elements.deleteRowBtn.disabled = true;
      }
      return;
    }

    state.selectedTableRow = row;
    row.classList.add('row-selected');
    if (elements.deleteRowBtn) {
      elements.deleteRowBtn.disabled = false;
    }
  }

  function deleteSelectedTableRow() {
    const row = state.selectedTableRow;
    if (!row || !row.parentNode) {
      setStatus('Select a row in the table, then click Delete Selected Row.', true);
      setSelectedTableRow(null);
      return;
    }

    row.parentNode.removeChild(row);
    setSelectedTableRow(null);
    ensureTypeNames();

    const hasRows = elements.ocrTableBody.querySelectorAll('tr').length > 0;
    elements.generateBtn.disabled = !hasRows;
    if (elements.exportPdfBtn) {
      elements.exportPdfBtn.disabled = true;
    }
    if (elements.reprintBtn) {
      elements.reprintBtn.disabled = true;
    }

    state.pendingTicketCommit = null;
    state.hasGeneratedCurrentDrawing = false;
    showDrawingProcessedButton(false);

    setStatus('Row deleted.', false);
  }

  function addRow(data) {
    const rowIndex = elements.ocrTableBody.querySelectorAll('tr').length;
    const selectedType = data.type || getRememberedCraftTypeForItemCode(data.itemCode || '') || '';
    const isManualType = Boolean(data.type);
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td contenteditable="true">${data.pointNumber || ''}</td>
      <td contenteditable="true">${data.itemCode || ''}</td>
      <td contenteditable="true">${data.size || ''}</td>
      <td contenteditable="true">${data.quantity || ''}</td>
      <td contenteditable="true">${data.materialDescription || data.description || ''}</td>
      <td contenteditable="true">${data.projectNo || elements.projectNo.value.trim()}</td>
      <td class="typeCell">${createTypeCellHtml(`type-${rowIndex}`, selectedType)}</td>
    `;
    elements.ocrTableBody.appendChild(tr);
    tr.dataset.typeManual = isManualType ? 'true' : 'false';
    tr.__ocrOriginal = data._original || null;
    ensureTypeNames();
    elements.generateBtn.disabled = elements.ocrTableBody.querySelectorAll('tr').length === 0;
    if (elements.exportPdfBtn) {
      elements.exportPdfBtn.disabled = true;
    }
  }

  function setDescriptionNeedsAttention(cell, needsAttention) {
    if (!cell) return;
    cell.classList.toggle('desc-needs-attention', Boolean(needsAttention));
  }

  function updateMaterialDescriptionForTableRow(row) {
    const cells = row.querySelectorAll('td');
    const itemCode = cells[1]?.textContent.trim() || '';
    const descriptionCell = cells[4];
    if (!descriptionCell) return;

    const matchedDescription = lookupDescriptionInCatalog(itemCode);
    if (matchedDescription) {
      descriptionCell.textContent = matchedDescription;
      setDescriptionNeedsAttention(descriptionCell, false);
      return;
    }

    const currentDescription = String(descriptionCell.textContent || '').trim();
    if (!itemCode) {
      setDescriptionNeedsAttention(descriptionCell, false);
      return;
    }

    if (!currentDescription) {
      descriptionCell.textContent = PLACEHOLDER_DESCRIPTION_TEXT;
      setDescriptionNeedsAttention(descriptionCell, true);
      return;
    }

    setDescriptionNeedsAttention(descriptionCell, currentDescription === PLACEHOLDER_DESCRIPTION_TEXT);
  }

  function clearRowTypeSelection(row) {
    if (!row) return;
    row.querySelectorAll('input[type="radio"]').forEach((radio) => {
      radio.checked = false;
    });
  }

  function getRowCellValue(row, columnIndex) {
    const cells = row.querySelectorAll('td');
    return cells[columnIndex]?.textContent.trim() || '';
  }

  function isPlaceholderItemCode(value) {
    return String(value || '').trim().toUpperCase().startsWith(PLACEHOLDER_ITEM_CODE);
  }

  function isInvalidItemCode(value) {
    const text = String(value || '').trim();
    if (!text) return true;
    if (/^-+$/.test(text)) return true;
    return false;
  }

  function normalizeCraftType(value) {
    const type = String(value || '').trim().toUpperCase();
    if (['P', 'WF', 'SF', 'E', 'S', 'EI'].includes(type)) return type;
    return '';
  }

  function loadCraftByItemCodeStore() {
    try {
      const raw = localStorage.getItem(CRAFT_BY_ITEM_CODE_STORAGE_KEY);
      if (!raw) return {};
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== 'object') return {};

      const normalized = {};
      Object.entries(parsed).forEach(([code, craftType]) => {
        const key = String(code || '').trim().toUpperCase();
        const type = normalizeCraftType(craftType);
        if (!key || !type) return;
        normalized[key] = type;
      });

      return normalized;
    } catch (error) {
      console.warn('Failed to load item-code craft memory.', error);
      return {};
    }
  }

  function saveCraftByItemCodeStore() {
    localStorage.setItem(CRAFT_BY_ITEM_CODE_STORAGE_KEY, JSON.stringify(state.craftByItemCodeStore || {}));
  }

  function getRowSelectedType(row) {
    return normalizeCraftType(row?.querySelector('input[type="radio"]:checked')?.value || '');
  }

  function rememberCraftTypeForItemCode(itemCode, craftType) {
    const code = String(itemCode || '').trim();
    const normalizedType = normalizeCraftType(craftType);
    if (!code || !normalizedType || isInvalidItemCode(code) || isPlaceholderItemCode(code)) {
      return false;
    }

    const candidates = getItemCodeCandidates(code);
    let changed = false;
    candidates.forEach((candidate) => {
      if (!candidate) return;
      if (state.craftByItemCodeStore[candidate] === normalizedType) return;
      state.craftByItemCodeStore[candidate] = normalizedType;
      changed = true;
    });

    if (changed) saveCraftByItemCodeStore();
    return changed;
  }

  function getRememberedCraftTypeForItemCode(itemCode) {
    const code = String(itemCode || '').trim();
    if (!code || isInvalidItemCode(code) || isPlaceholderItemCode(code)) return '';

    const candidates = getItemCodeCandidates(code);
    for (const candidate of candidates) {
      const remembered = normalizeCraftType(state.craftByItemCodeStore[candidate]);
      if (remembered) return remembered;
    }

    return '';
  }

  function applyRememberedCraftTypeToRow(row, options) {
    const settings = options || {};
    if (!row) return false;
    const preserveExisting = settings.preserveExisting !== false;
    if (preserveExisting && getRowSelectedType(row)) return false;

    const cells = row.querySelectorAll('td');
    const itemCode = cells[1]?.textContent.trim() || '';
    if (!itemCode) return false;

    if (!preserveExisting) {
      clearRowTypeSelection(row);
    }

    const rememberedType = getRememberedCraftTypeForItemCode(itemCode);
    if (!rememberedType) return false;

    const radio = row.querySelector(`input[type="radio"][value="${rememberedType}"]`);
    if (!radio) return false;
    radio.checked = true;
    row.dataset.typeManual = 'false';
    return true;
  }

  function rememberCraftMappingsFromTable() {
    let rememberedCount = 0;
    const rows = Array.from(elements.ocrTableBody.querySelectorAll('tr'));
    rows.forEach((row) => {
      const cells = row.querySelectorAll('td');
      const itemCode = cells[1]?.textContent.trim() || '';
      const selectedType = getRowSelectedType(row);
      if (!selectedType || !itemCode) return;
      if (rememberCraftTypeForItemCode(itemCode, selectedType)) {
        rememberedCount += 1;
      }
    });
    return rememberedCount;
  }

  function ensureUsableItemCodesInTable() {
    const rows = Array.from(elements.ocrTableBody.querySelectorAll('tr'));
    let promptedCount = 0;
    let placeholderCount = 0;

    rows.forEach((row, index) => {
      const cells = row.querySelectorAll('td');
      const itemCodeCell = cells[1];
      if (!itemCodeCell) return;

      const currentValue = itemCodeCell.textContent.trim();
      if (!isInvalidItemCode(currentValue)) return;

      promptedCount += 1;
      const entered = window.prompt(
        `Row ${index + 1}: Item Code is blank or dashes. Enter a usable code. Leave blank to use ${PLACEHOLDER_ITEM_CODE}.`,
        ''
      );

      const normalizedEntered = String(entered || '').trim();
      if (normalizedEntered && !isInvalidItemCode(normalizedEntered)) {
        itemCodeCell.textContent = normalizedEntered;
        const matchedDescription = lookupDescriptionInCatalog(normalizedEntered);
        if (cells[4]) {
          cells[4].textContent = matchedDescription || PLACEHOLDER_DESCRIPTION_TEXT;
          setDescriptionNeedsAttention(cells[4], !matchedDescription);
        }
      } else {
        const placeholder = `${PLACEHOLDER_ITEM_CODE}-${String(index + 1).padStart(2, '0')}`;
        itemCodeCell.textContent = placeholder;
        if (cells[4]) {
          cells[4].textContent = PLACEHOLDER_DESCRIPTION_TEXT;
          setDescriptionNeedsAttention(cells[4], true);
        }
        placeholderCount += 1;
      }

      row.__skipLearning = true;
    });

    return { promptedCount, placeholderCount };
  }

  function learnCorrectionsFromTable() {
    if (!state.learningStore) {
      state.learningStore = createEmptyLearningStore();
    }

    let learnedCount = 0;
    const rows = Array.from(elements.ocrTableBody.querySelectorAll('tr'));

    rows.forEach((row) => {
      const original = row.__ocrOriginal;
      if (!original) return;
      if (row.__skipLearning) return;

      const currentByField = {
        pointNumber: getRowCellValue(row, 0),
        itemCode: getRowCellValue(row, 1),
        size: getRowCellValue(row, 2),
        quantity: getRowCellValue(row, 3)
      };

      LEARNING_FIELDS.forEach((field) => {
        const originalValue = String(original[field] || '').trim();
        const correctedValue = String(currentByField[field] || '').trim();
        if (!originalValue || !correctedValue || originalValue === correctedValue) return;
        if (field === 'itemCode' && (isPlaceholderItemCode(correctedValue) || isInvalidItemCode(correctedValue))) return;

        const key = normalizeLearningKey(originalValue);
        const existing = state.learningStore[field][key];

        if (!existing || existing.value !== correctedValue) {
          state.learningStore[field][key] = {
            value: correctedValue,
            count: 1,
            updatedAt: Date.now()
          };
          learnedCount += 1;
          return;
        }

        existing.count = (existing.count || 0) + 1;
        existing.updatedAt = Date.now();
      });
    });

    if (learnedCount > 0) {
      saveLearningStore();
    }

    return learnedCount;
  }

  function collectRows() {
    const rows = Array.from(elements.ocrTableBody.querySelectorAll('tr'));
    return rows.map((row, rowIndex) => {
      const cells = row.querySelectorAll('td');
      const selected = row.querySelector('input[type="radio"]:checked');
      const materialDescription = cells[4]?.textContent.trim() || '';
      return {
        pointNumber: normalizePointNumberValue(cells[0]?.textContent.trim() || ''),
        itemCode: cells[1]?.textContent.trim() || '',
        size: cells[2]?.textContent.trim() || '',
        description: materialDescription,
        materialDescription,
        quantity: cells[3]?.textContent.trim() || '',
        projectNo: cells[5]?.textContent.trim() || elements.projectNo.value.trim(),
        type: selected ? selected.value : ''
      };
    }).filter((row) => row.itemCode || row.quantity || row.size || row.pointNumber);
  }

  function enforceExactCatalogDescriptionsInTable() {
    const rows = Array.from(elements.ocrTableBody.querySelectorAll('tr'));
    let correctedCount = 0;

    rows.forEach((row) => {
      const cells = row.querySelectorAll('td');
      const itemCode = String(cells[1]?.textContent || '').trim();
      const descriptionCell = cells[4];
      if (!itemCode || !descriptionCell) return;
      if (isPlaceholderItemCode(itemCode)) return;

      const matchedDescription = lookupDescriptionInCatalog(itemCode);
      if (!matchedDescription) return;

      const currentDescription = String(descriptionCell.textContent || '').trim();
      if (currentDescription === matchedDescription) {
        setDescriptionNeedsAttention(descriptionCell, false);
        return;
      }

      descriptionCell.textContent = matchedDescription;
      setDescriptionNeedsAttention(descriptionCell, false);
      correctedCount += 1;
    });

    return correctedCount;
  }

  function hasMeaningfulRowData(row) {
    const cells = row.querySelectorAll('td');
    const pointNumber = cells[0]?.textContent.trim() || '';
    const itemCode = cells[1]?.textContent.trim() || '';
    const size = cells[2]?.textContent.trim() || '';
    const quantity = cells[3]?.textContent.trim() || '';
    return Boolean(pointNumber || itemCode || size || quantity);
  }

  function clearCraftMissingHighlights() {
    elements.ocrTableBody.querySelectorAll('tr').forEach((row) => row.classList.remove('craft-missing'));
  }

  function validateCraftSelectionBeforeGeneration() {
    clearCraftMissingHighlights();

    const missingRowNumbers = [];
    const rows = Array.from(elements.ocrTableBody.querySelectorAll('tr'));
    rows.forEach((row, index) => {
      if (!hasMeaningfulRowData(row)) return;
      if (getRowSelectedType(row)) return;
      row.classList.add('craft-missing');
      missingRowNumbers.push(index + 1);
    });

    if (missingRowNumbers.length > 0) {
      setStatus(`Generation blocked: select craft type for row(s): ${missingRowNumbers.join(', ')}.`, true);
      return false;
    }

    return true;
  }

  function determineCraft(type) {
    const value = (type || '').toUpperCase();
    if (value === 'P' || value.includes('PIP')) return 'Piping';
    if (value === 'WF' || value.includes('WORK')) return 'Workshop Fabrication';
    if (value === 'SF' || value.includes('SITE')) return 'Site Fabrication';
    if (value === 'E' || value.includes('ERECT')) return 'Erection';
    if (value === 'S' || value.includes('SUPPORT')) return 'Supports';
    if (value === 'EI' || value.includes('ELECTR')) return 'Electrical & Instrumentation';
    return 'Fabrication';
  }

  function mapTypeValueToSheet(typeValue) {
    const value = normalizeCraftType(typeValue);
    if (value === 'P') return 'Pipe';
    if (value === 'E') return 'Erection';
    if (value === 'S') return 'Supports';
    if (value === 'WF' || value === 'SF' || value === 'EI') return 'Fabrication';
    return 'Fabrication';
  }

  function mapCraftToSheet(craft) {
    if (craft === 'Piping') return 'Pipe';
    if (craft === 'Erection') return 'Erection';
    if (craft === 'Supports') return 'Supports';
    return 'Fabrication';
  }

  function craftLabelFromSheet(sheetName) {
    if (sheetName === 'Pipe') return 'Piping';
    if (sheetName === 'Fabrication') return 'Fabrication';
    if (sheetName === 'Erection') return 'Erection';
    if (sheetName === 'Supports') return 'Supports';
    return sheetName;
  }

  function buildRequestedNumber(projectNo) {
    const text = String(projectNo || '').trim();
    const match = text.match(/(\d+)/);
    if (!match) return '-';
    return `${match[1]}-`;
  }

  function escapeHtml(value) {
    return String(value || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function sanitizeFilenamePart(value, fallback) {
    const clean = String(value || '').trim().replace(/[\\/:*?"<>|]+/g, '-').replace(/\s+/g, '_');
    return clean || fallback;
  }

  function normalizeInsulationText(value) {
    const text = String(value || '').trim();
    if (!text) return 'None';
    return text;
  }

  function normalizeInsulationCode(value) {
    const text = String(value || '').trim();
    if (!text) return '';
    if (isNoInsulation(text)) return '';
    return text.replace(/\s+/g, '').toUpperCase();
  }

  function buildDrawingNumberWithInsulation(drawingNumber, insulationType) {
    const baseDrawing = String(drawingNumber || '').trim().toUpperCase();
    const insulationText = String(insulationType || '').trim();
    if (!insulationText || isNoInsulation(insulationText)) return baseDrawing;
    return `${baseDrawing}-${insulationText}`;
  }

  function isNoInsulation(insulationText) {
    const normalized = String(insulationText || '').trim().toLowerCase();
    return !normalized || normalized === 'none' || normalized === 'n/a' || normalized === 'na';
  }

  function computePaintSpec(sheetName, temperatureValue, insulationText, materialType) {
    const normalizedMaterial = String(materialType || '').trim().toUpperCase();
    if (normalizedMaterial === 'SS' || normalizedMaterial === 'GS') {
      return 'N/A';
    }

    if (sheetName === 'Supports') {
      return 'N/A';
    }

    const temperature = Number(temperatureValue);
    if (!Number.isFinite(temperature)) return '';

    const noInsulation = isNoInsulation(insulationText);
    let firstCode = '';
    if (noInsulation) {
      if (temperature <= 90) firstCode = '2A';
      else if (temperature <= 120) firstCode = '2B';
      else if (temperature <= 200) firstCode = '2C';
      else if (temperature <= 400) firstCode = '2D';
      else if (temperature <= 650) firstCode = '2I';
    } else {
      if (temperature <= 150) firstCode = '2E';
      else if (temperature <= 650) firstCode = '2I';
    }

    if (sheetName === 'Pipe') {
      return firstCode;
    }

    let secondCode = '';
    if (noInsulation) {
      if (temperature <= 90) secondCode = '5A';
      else if (temperature <= 200) secondCode = '5B';
      else if (temperature <= 650) secondCode = '2I';
    } else {
      if (temperature <= 150) secondCode = '5C';
      else if (temperature <= 650) secondCode = '2I';
    }

    if (!firstCode && !secondCode) return '';
    if (!secondCode) return firstCode;
    return `${firstCode}/${secondCode}`;
  }

  function getCraftSuffix(sheetName) {
    if (sheetName === 'Pipe') return 'pipe';
    if (sheetName === 'Fabrication') return 'fabrication';
    if (sheetName === 'Erection') return 'erection';
    if (sheetName === 'Supports') return 'supports';
    return sanitizeFilenamePart(sheetName, 'craft').toLowerCase();
  }

  function getTicketCounterKey(projectNo) {
    return `matcon_standalone_ticket_counter_${projectNo || 'default'}`;
  }

  function getStoredLastTicketNumber(projectNo) {
    const key = getTicketCounterKey(projectNo);
    const parsed = parseInt(localStorage.getItem(key) || '0', 10);
    return Number.isFinite(parsed) && parsed > 0 ? parsed : 0;
  }

  function setStoredLastTicketNumber(projectNo, value) {
    const key = getTicketCounterKey(projectNo);
    localStorage.setItem(key, String(Math.max(0, parseInt(String(value || 0), 10) || 0)));
  }

  function formatTicketNumber(numberValue) {
    return String(numberValue).padStart(6, '0');
  }

  function getTicketStartFromInput(projectNo) {
    const entered = parseInt(String(elements.ticketStartNo?.value || '').trim(), 10);
    if (Number.isFinite(entered) && entered > 0) {
      return entered;
    }

    return getStoredLastTicketNumber(projectNo) + 1;
  }

  function syncTicketStartField(projectNo) {
    const nextTicket = getStoredLastTicketNumber(projectNo) + 1;
    if (elements.ticketStartNo) {
      elements.ticketStartNo.value = String(nextTicket);
    }
  }

  function buildTicketNumbers(startNumber, ticketCount) {
    const safeCount = Math.max(0, ticketCount || 0);
    return Array.from({ length: safeCount }, (_, index) => formatTicketNumber(startNumber + index));
  }

  function commitTicketNumbers(projectNo, finalNumber) {
    const parsedFinal = parseInt(String(finalNumber || 0), 10);
    if (!Number.isFinite(parsedFinal) || parsedFinal < 1) return;
    setStoredLastTicketNumber(projectNo, parsedFinal);
    syncTicketStartField(projectNo);
  }

  function buildTicketPayload(rows) {
    applyHeaderFieldFormatting();
    const learnedCount = learnCorrectionsFromTable();
    rememberCraftMappingsFromTable();

    const missingTypes = rows.filter((row) => !row.type);
    if (missingTypes.length > 0) {
      alert('All rows need a craft type before ticket generation.');
      return null;
    }

    const drawingNumber = getDrawingNumber();
    if (!drawingNumber) {
      alert('Fill Drawing Number before generating tickets.');
      return null;
    }

    const operatingTemperature = String(elements.operatingTemperature?.value || '').trim();
    const insulationType = String(elements.insulationType?.value || '').trim();
      const drawingNumberWithInsulation = buildDrawingNumberWithInsulation(drawingNumber, insulationType);

    if (!operatingTemperature) {
      alert('Fill Operating Temperature before generating tickets.');
      return null;
    }
    if (!insulationType) {
      alert('Fill Insulation Type before generating tickets.');
      return null;
    }

    const projectNo = elements.projectNo.value.trim() || rows[0]?.projectNo || 'UNKNOWN';
    const workpackNo = String(elements.workpackNo?.value || '').trim();
    if (!workpackNo) {
      alert('Fill Workpack before generating tickets.');
      return null;
    }
    const grouped = {
      Pipe: [],
      Fabrication: [],
      Erection: [],
      Supports: []
    };

    rows.forEach((row) => {
      const sheetName = mapTypeValueToSheet(row.type);
      grouped[sheetName].push(row);
    });

    const activeSheetNames = SHEET_ORDER.filter((sheetName) => grouped[sheetName].length > 0);
    const ticketStartNumber = getTicketStartFromInput(projectNo);
    const reservedTicketNumbers = buildTicketNumbers(ticketStartNumber, activeSheetNames.length);
    const ticketFinalNumber = activeSheetNames.length > 0
      ? ticketStartNumber + activeSheetNames.length - 1
      : ticketStartNumber - 1;

    const tickets = activeSheetNames.map((sheetName, index) => ({
      sheetName,
      craftLabel: craftLabelFromSheet(sheetName),
      rows: grouped[sheetName],
      ticketNo: reservedTicketNumbers[index]
    }));

    return {
      drawingNumber,
      drawingNumberWithInsulation,
      projectNo,
      workpackNo,
      learnedCount,
      operatingTemperature,
      insulationType: normalizeInsulationText(insulationType),
      materialType: String(elements.materialType?.value || 'CS').trim().toUpperCase() || 'CS',
      sheetNo: elements.sheetNo.value.trim(),
      revision: elements.revision.value.trim(),
      ticketStartNumber,
      ticketFinalNumber,
      grouped,
      tickets
    };
  }

  const ROWS_PER_TICKET_PAGE = 24;

  function splitRowsIntoPages(rows, pageSize) {
    const source = Array.isArray(rows) ? rows : [];
    const size = Math.max(1, pageSize || ROWS_PER_TICKET_PAGE);
    if (!source.length) return [[]];

    const pages = [];
    for (let i = 0; i < source.length; i += size) {
      pages.push(source.slice(i, i + size));
    }
    return pages.length ? pages : [[]];
  }

  function buildExcelStylePreviewHtml(payload, ticket, options) {
    const settings = options || {};
    const pageNumber = Math.max(1, parseInt(settings.pageNumber || '1', 10) || 1);
    const pageCount = Math.max(1, parseInt(settings.pageCount || '1', 10) || 1);
    const isReprint = Boolean(settings.isReprint);
    const requestedNo = buildRequestedNumber(payload?.projectNo);
    const customTicketNumber = String(settings.ticketNumberOverride || '').trim();
    const ticketNumberText = customTicketNumber || ticket.ticketNo;
    const ticketNoLabel = isReprint ? `Reprint: ${ticketNumberText}` : `${ticketNumberText}`;
    const ticketNoColor = '#b91c1c';
    const paintSpec = computePaintSpec(ticket?.sheetName, payload?.operatingTemperature, payload?.insulationType, payload?.materialType);
    const logoSrc = state.logoDataUrl || './Aurex%20Logo.jpg';
    const logoCellHtml = `<span class="logo-frame"><img src="${logoSrc}" alt="Aurex" crossorigin="anonymous" class="aurex-logo" /></span>`;

    const rows = Array.isArray(settings.rows)
      ? settings.rows.slice(0, ROWS_PER_TICKET_PAGE)
      : (ticket?.rows || []).slice(0, ROWS_PER_TICKET_PAGE);

    while (rows.length < ROWS_PER_TICKET_PAGE) {
      rows.push({ pointNumber: '', itemCode: '', size: '', quantity: '', materialDescription: '' });
    }

    const dataRowsHtml = rows.map((item) => `
      <tr style="height: 26px;">
        <td class="input-cell">${escapeHtml(item.pointNumber)}</td>
        <td class="input-cell">${escapeHtml(item.itemCode)}</td>
        <td class="input-cell">${escapeHtml(item.size)}</td>
        <td class="input-cell">${escapeHtml(item.quantity)}</td>
        <td class="input-cell">${escapeHtml(item.materialDescription || '')}</td>
        <td class="input-cell"></td>
        <td class="input-cell"></td>
      </tr>
    `).join('');

    return `
      <div class="ticket-root">
        <style>
          .ticket-root { font-family: Calibri, Arial, sans-serif; color: #111827; }
          .ticket-root .sheet-wrap { width: 616px; margin: 0 auto; background: #fff; }
          .ticket-root table.sheet { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 11px; line-height: 1.05; border: 1.5px solid #111827; }
          .ticket-root .sheet td { border: 1px solid #111827; padding: 1px 3px; vertical-align: middle; white-space: nowrap; }
          .ticket-root .hdr { font-weight: 700; text-align: center; vertical-align: middle; }
          .ticket-root .major-row td { border-top-width: 1.5px; border-bottom-width: 1.5px; }
          .ticket-root .table-head td { border-top-width: 1.5px; border-bottom-width: 1.5px; background: #B8CCE4; }
          .ticket-root .wrap-head { white-space: normal !important; line-height: 1.1; text-align: center; overflow-wrap: anywhere; word-break: normal; }
          .ticket-root .top-band td { border: none; background: #fff; }
          .ticket-root .picking-title { background: #B8CCE4 !important; border: none !important; font-weight: 700; font-family: Arial, sans-serif; font-size: 10pt; text-align: center; }
          .ticket-root .page-title { font-family: Arial, sans-serif; font-size: 10pt; font-weight: 700; text-align: center; }
          .ticket-root .craft-row td { font-family: Arial, sans-serif; font-size: 8pt; font-weight: 700; text-align: center; border-bottom: none; }
          .ticket-root .craft-input-row td { border-top: none; }
          .ticket-root .project-row td { border-bottom: none; }
          .ticket-root .project-input-row td { border-top: none; }
          .ticket-root .signature-title td { background: #B8CCE4; }
          .ticket-root .sig-label-row td { border-bottom: none; }
          .ticket-root .sig-input-row td { border-top: none; }
          .ticket-root .sig-split { width: 100%; height: 100%; border-collapse: collapse; table-layout: fixed; }
          .ticket-root .sig-split td { border: none; text-align: center; vertical-align: middle; padding: 0 3px; background: transparent; }
          .ticket-root .sig-left-group { position: relative; }
          .ticket-root .sig-left-group::after { content: ''; position: absolute; top: 0; bottom: 0; left: 50%; width: 1px; background: #111827; transform: translateX(-0.5px); }
          .ticket-root .logo-cell { padding: 2px 4px; text-align: left; vertical-align: middle; position: relative; overflow: visible !important; z-index: 20; background: #fff; }
          .ticket-root .logo-frame { display: inline-flex; align-items: center; justify-content: flex-start; background: #fff; position: relative; z-index: 21; padding: 0; min-height: 42px; }
          .ticket-root .aurex-logo { display: block; width: 106px; height: 40px; object-fit: contain; max-width: none; position: relative; z-index: 22; opacity: 1 !important; filter: none !important; mix-blend-mode: normal; }
          .ticket-root .aurex-logo-fallback { display: inline-block; position: relative; top: -6px; left: 2px; font-family: Arial, sans-serif; font-size: 18px; font-weight: 700; letter-spacing: 0.2px; color: #111827; }
          .ticket-root .input-cell { background: #fbfdff; text-align: center; vertical-align: middle; padding-top: 0; padding-bottom: 0; font-weight: 700; }
          .ticket-root .left-red { text-align: left; color: #b91c1c; }
          .ticket-root .center-red { color: #b91c1c; }
          .ticket-root .ticket-no { text-align: center; color: ${ticketNoColor}; font-weight: 700; }
          .ticket-root .date-heading { color: #111827 !important; font-weight: 700 !important; text-align: center !important; }
        </style>
        <div class="sheet-wrap">
          <table class="sheet" aria-label="Picking Ticket">
            <colgroup>
              <col style="width: 56px;" />
              <col style="width: 74px;" />
              <col style="width: 82px;" />
              <col style="width: 68px;" />
              <col style="width: 192px;" />
              <col style="width: 60px;" />
              <col style="width: 84px;" />
            </colgroup>

            <tr style="height: 20px;" class="top-band"><td colspan="7"></td></tr>
            <tr style="height: 20px;" class="top-band"><td colspan="7"></td></tr>
            <tr style="height: 46px;" class="top-band">
              <td colspan="2" class="logo-cell">${logoCellHtml}</td>
              <td colspan="3" class="picking-title">Picking Ticket</td>
              <td colspan="2" class="page-title">Page ${pageNumber}&nbsp;&nbsp;of&nbsp;&nbsp;${pageCount}</td>
            </tr>
            <tr style="height: 11px;" class="top-band"><td colspan="7"></td></tr>

            <tr style="height: 15px;" class="major-row craft-row">
              <td colspan="2" class="hdr">Craft</td>
              <td class="hdr">Requested No.</td>
              <td class="hdr">Paint Spec:</td>
              <td class="hdr">Picking Ticket No.</td>
              <td colspan="2" class="hdr date-heading">Date:</td>
            </tr>
            <tr style="height: 26px;" class="craft-input-row">
              <td colspan="2" class="input-cell">${escapeHtml(ticket.craftLabel)}</td>
              <td class="input-cell left-red">${escapeHtml(requestedNo)}</td>
              <td class="input-cell">${escapeHtml(paintSpec)}</td>
              <td class="input-cell ticket-no">${escapeHtml(ticketNoLabel)}</td>
              <td colspan="2" class="input-cell"></td>
            </tr>

            <tr style="height: 16px;" class="major-row project-row">
              <td colspan="3" class="hdr">Project No. &amp; Contract No.</td>
              <td class="hdr">Work Pack</td>
              <td class="hdr">Drawing No.</td>
              <td class="hdr">Sht</td>
              <td class="hdr">Rev</td>
            </tr>
            <tr style="height: 26px;" class="project-input-row">
              <td colspan="3" class="input-cell">${escapeHtml(payload.projectNo)}</td>
              <td class="input-cell center-red">${escapeHtml(payload.workpackNo)}</td>
              <td class="input-cell">${escapeHtml(payload.drawingNumberWithInsulation || payload.drawingNumber)}</td>
              <td class="input-cell">${escapeHtml(payload.sheetNo || '-')}</td>
              <td class="input-cell">${escapeHtml(payload.revision || '-')}</td>
            </tr>

            <tr style="height: 42px;" class="table-head">
              <td class="hdr wrap-head">Drawing Item<br>No.</td>
              <td class="hdr wrap-head">Item Code or<br>Tag No,</td>
              <td class="hdr">Size</td>
              <td class="hdr wrap-head">Quantity<br>Required</td>
              <td class="hdr">Material Description</td>
              <td class="hdr wrap-head">Quantity<br>Issued</td>
              <td class="hdr">Heat Number</td>
            </tr>

            ${dataRowsHtml}

            <tr style="height: 17px;" class="major-row signature-title">
              <td colspan="7" class="hdr">Signature</td>
            </tr>
            <tr style="height: 16px;" class="sig-label-row">
              <td colspan="5" class="sig-left-group" style="padding: 0;">
                <table class="sig-split" aria-hidden="true"><tr><td class="hdr">Approved:</td><td class="hdr">Issued By:</td></tr></table>
              </td>
              <td colspan="2" class="hdr">Received By:</td>
            </tr>
            <tr style="height: 46px;" class="sig-input-row">
              <td colspan="5" class="sig-left-group" style="padding: 0;">
                <table class="sig-split" aria-hidden="true"><tr><td class="input-cell"></td><td class="input-cell"></td></tr></table>
              </td>
              <td colspan="2" class="input-cell"></td>
            </tr>
          </table>
        </div>
      </div>
    `;
  }

  function renderTicketPreview(payload, selectedSheetName) {
    const tickets = payload?.tickets || [];
    if (!tickets.length) {
      elements.ticketPreviewContainer.innerHTML = 'No ticket data available for preview.';
      elements.previewCraftSelect.innerHTML = '';
      elements.previewCraftSelect.disabled = true;
      elements.openPreviewWindowBtn.disabled = true;
      return;
    }

    elements.previewCraftSelect.innerHTML = tickets
      .map((ticket) => `<option value="${escapeHtml(ticket.sheetName)}">${escapeHtml(ticket.craftLabel)} (${escapeHtml(ticket.ticketNo)})</option>`)
      .join('');
    elements.previewCraftSelect.disabled = false;
    elements.openPreviewWindowBtn.disabled = false;

    const selected = tickets.find((ticket) => ticket.sheetName === selectedSheetName) || tickets[0];
    const pages = splitRowsIntoPages(selected.rows || [], ROWS_PER_TICKET_PAGE);
    const firstPageRows = pages[0] || [];
    elements.previewCraftSelect.value = selected.sheetName;
    elements.ticketPreviewContainer.innerHTML = buildExcelStylePreviewHtml(payload, selected, {
      rows: firstPageRows,
      pageNumber: 1,
      pageCount: pages.length
    });
    state.lastPreviewPayload = payload;
  }

  function openTicketPreviewWindow() {
    if (!state.lastPreviewPayload || !state.lastPreviewPayload.tickets?.length) {
      setStatus('Generate ticket data first, then open print preview.', true);
      return;
    }

    const selectedSheetName = elements.previewCraftSelect.value;
    const selectedTicket = state.lastPreviewPayload.tickets.find((ticket) => ticket.sheetName === selectedSheetName) || state.lastPreviewPayload.tickets[0];
    const pages = splitRowsIntoPages(selectedTicket.rows || [], ROWS_PER_TICKET_PAGE);
    const previewHtml = pages.map((rows, index) => `
      <div class="preview-page">${buildExcelStylePreviewHtml(state.lastPreviewPayload, selectedTicket, {
        rows,
        pageNumber: index + 1,
        pageCount: pages.length
      })}</div>
    `).join('');

    const popup = window.open('', '_blank', 'width=1000,height=900');
    if (!popup) {
      setStatus('Pop-up blocked. Allow pop-ups to open print preview.', true);
      return;
    }

    popup.document.write(`
      <!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8">
        <title>Picking Ticket Preview</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 12px; background: #fff; }
          .excel-ticket { border: 1px solid #000; padding: 10px; width: 740px; margin: 0 auto 12px; }
          .preview-page { page-break-after: always; }
          .preview-page:last-child { page-break-after: auto; }
          .excel-ticket .sheet-title { font-weight: 700; text-align: center; font-size: 16px; margin-bottom: 4px; }
          .excel-ticket .sheet-subtitle { text-align: center; font-size: 11px; margin-bottom: 8px; }
          .excel-ticket table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 10px; }
          .excel-ticket td, .excel-ticket th { border: 1px solid #000; padding: 2px 4px; vertical-align: middle; word-wrap: break-word; }
          .excel-ticket .center { text-align: center; }
          .excel-ticket .bold { font-weight: 700; }
          .excel-ticket .h28 td { height: 28px; }
          .excel-ticket .h22 td { height: 22px; }
          .excel-ticket .h18 td { height: 18px; }
          @media print {
            @page { size: A4 portrait; margin: 9mm; }
            body { margin: 0; }
          }
        </style>
      </head>
      <body>
        ${previewHtml}
      </body>
      </html>
    `);
    popup.document.close();
  }

  function normalizeRevisionForFilename(value) {
    const raw = String(value || '').trim();
    if (!raw) return '00';
    return raw.replace(/^[Rr]+/, '') || '00';
  }

  function buildTicketPdfFilename(payload, ticket) {
    const drawingPart = sanitizeFilenamePart(payload?.drawingNumber, 'DRAWING');
    const insulationPart = sanitizeFilenamePart(normalizeInsulationCode(payload?.insulationType), '');
    const sheetPart = sanitizeFilenamePart(payload?.sheetNo, 'SHEET');
    const revisionPart = normalizeRevisionForFilename(payload?.revision);
    const craftSuffix = getCraftSuffix(ticket?.sheetName);
    if (insulationPart) {
      return `${ticket.ticketNo}-${drawingPart}-${insulationPart}-${sheetPart}-R${revisionPart}-${craftSuffix}.pdf`;
    }
    return `${ticket.ticketNo}-${drawingPart}-${sheetPart}-R${revisionPart}-${craftSuffix}.pdf`;
  }

  async function pickOutputFolder() {
    if (typeof window.showDirectoryPicker !== 'function') {
      return null;
    }

    alert('Select the default folder for PDF Save As dialogs. You can still change location/file name each time.');
    return await window.showDirectoryPicker({ mode: 'readwrite' });
  }

  async function saveBlobToDirectory(handle, filename, blob) {
    const fileHandle = await handle.getFileHandle(filename, { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(blob);
    await writable.close();
  }

  async function saveBlobWithSaveDialog(blob, filename, startInHandle) {
    if (typeof window.showSaveFilePicker !== 'function') {
      return null;
    }

    const pickerOptions = {
      suggestedName: filename,
      types: [
        {
          description: 'PDF Document',
          accept: { 'application/pdf': ['.pdf'] }
        }
      ]
    };

    if (startInHandle) {
      pickerOptions.startIn = startInHandle;
    }

    const fileHandle = await window.showSaveFilePicker(pickerOptions);
    const writable = await fileHandle.createWritable();
    await writable.write(blob);
    await writable.close();
    return fileHandle;
  }

  async function renderTicketPageCanvas(payload, ticket, pageRows, pageNumber, pageCount, options) {
    const settings = options || {};
    await loadLogoDataUrl();
    const wrapper = document.createElement('div');
    wrapper.style.position = 'fixed';
    wrapper.style.left = '-10000px';
    wrapper.style.top = '0';
    wrapper.style.width = 'fit-content';
    wrapper.style.padding = '8px 12px';
    wrapper.style.boxSizing = 'border-box';
    wrapper.style.overflow = 'visible';
    wrapper.style.opacity = '1';
    wrapper.style.pointerEvents = 'none';
    wrapper.style.zIndex = '-1';
    wrapper.style.background = '#fff';
    wrapper.innerHTML = buildExcelStylePreviewHtml(payload, ticket, {
      rows: pageRows,
      pageNumber,
      pageCount,
      isReprint: Boolean(settings.isReprint),
      ticketNumberOverride: settings.ticketNumberOverride || ''
    });
    document.body.appendChild(wrapper);

    try {
      await new Promise((resolve) => window.requestAnimationFrame(() => window.requestAnimationFrame(resolve)));
      await waitForImagesToLoad(wrapper, 6000);
      return await window.html2canvas(wrapper, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#ffffff'
      });
    } finally {
      document.body.removeChild(wrapper);
    }
  }

  async function buildPdfBlobFromTicket(payload, ticket, options) {
    const settings = options || {};
    const pages = splitRowsIntoPages(ticket.rows || [], ROWS_PER_TICKET_PAGE);
    const doc = new window.jspdf.jsPDF({
      orientation: 'portrait',
      unit: 'pt',
      format: 'a4'
    });

    for (let index = 0; index < pages.length; index += 1) {
      const pageRows = pages[index];
      const canvas = await renderTicketPageCanvas(payload, ticket, pageRows, index + 1, pages.length, {
        isReprint: Boolean(settings.isReprint),
        ticketNumberOverride: settings.ticketNumberOverride || ''
      });
      const pngBlob = await canvasToBlob(canvas, 'image/png');
      const imageData = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(new Error('Failed to read rendered page image.'));
        reader.readAsDataURL(pngBlob);
      });

      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();
      const margin = 18;
      const maxWidth = pageWidth - (margin * 2);
      const maxHeight = pageHeight - (margin * 2);

      const widthRatio = maxWidth / canvas.width;
      const heightRatio = maxHeight / canvas.height;
      const scale = Math.min(widthRatio, heightRatio);
      const renderWidth = canvas.width * scale;
      const renderHeight = canvas.height * scale;
      const offsetX = (pageWidth - renderWidth) / 2;
      const offsetY = (pageHeight - renderHeight) / 2;

      if (index > 0) {
        doc.addPage();
      }

      doc.addImage(imageData, 'PNG', offsetX, offsetY, renderWidth, renderHeight);
    }

    const pdfArrayBuffer = doc.output('arraybuffer');
    return new Blob([pdfArrayBuffer], { type: 'application/pdf' });
  }

  async function exportCraftPdfsFromHtml(payload, options) {
    const settings = options || {};
    if (!payload?.tickets?.length) {
      setStatus('Generate picking tickets before exporting PDFs.', true);
      return false;
    }

    if (!window.jspdf || !window.jspdf.jsPDF || !window.html2canvas) {
      setStatus('PDF export library not loaded (jsPDF/html2canvas).', true);
      return false;
    }

    const exportLabel = settings.isReprint ? 'reprint PDF(s)' : 'craft PDF(s)';
    const exportedCraftList = (payload.tickets || []).map((ticket) => ticket.craftLabel).join(', ');
    setStatus(`Exporting ${payload.tickets.length} ${exportLabel} from HTML...`, false);

    let outputFolderHandle = null;
    let saveDialogStartHandle = state.lastSavePickerHandle || null;
    if (settings.preferDirectoryPicker) {
      try {
        if (typeof window.showSaveFilePicker === 'function') {
          if (!saveDialogStartHandle) {
            saveDialogStartHandle = await pickOutputFolder();
            state.lastSavePickerHandle = saveDialogStartHandle;
          }
        } else {
          outputFolderHandle = await pickOutputFolder();
        }
      } catch (error) {
        if (error?.name === 'AbortError') {
          setStatus('PDF export canceled while selecting folders.', true);
          return false;
        }
        console.warn('Folder picker unavailable, using download fallback.', error);
        outputFolderHandle = null;
      }
    }

    for (let ticketIndex = 0; ticketIndex < payload.tickets.length; ticketIndex += 1) {
      const ticket = payload.tickets[ticketIndex];
      setStatus(
        `Exporting ${ticketIndex + 1}/${payload.tickets.length}: ${ticket.craftLabel} (${ticket.ticketNo})...`,
        false
      );

      const blob = await buildPdfBlobFromTicket(payload, ticket, {
        isReprint: Boolean(settings.isReprint),
        ticketNumberOverride: settings.ticketNumberOverride || ''
      });
      const filename = buildTicketPdfFilename(payload, ticket);

      if (settings.preferDirectoryPicker && typeof window.showSaveFilePicker === 'function') {
        try {
          const chosenFileHandle = await saveBlobWithSaveDialog(blob, filename, saveDialogStartHandle);
          if (chosenFileHandle) {
            saveDialogStartHandle = chosenFileHandle;
            state.lastSavePickerHandle = chosenFileHandle;
          }
        } catch (error) {
          if (error?.name === 'AbortError') {
            setStatus('PDF export canceled in Save As dialog.', true);
            return false;
          }
          throw error;
        }
      } else if (outputFolderHandle) {
        await saveBlobToDirectory(outputFolderHandle, filename, blob);
      } else {
        triggerDownload(blob, filename);
        await new Promise((resolve) => setTimeout(resolve, 120));
      }
    }

    if (outputFolderHandle) {
      const doneText = settings.isReprint
        ? `Reprint complete: PDFs saved to the Picking Tickets folder. Crafts: ${exportedCraftList}.`
        : `Export complete: PDFs saved to the Picking Tickets folder. Crafts: ${exportedCraftList}.`;
      setStatus(doneText, false);
      return true;
    }

    const fallbackText = settings.isReprint
      ? `Reprint complete: browser fallback used (filenames include craft suffix). Crafts: ${exportedCraftList}.`
      : `Export complete: browser fallback used (filenames include craft suffix). Crafts: ${exportedCraftList}.`;
    setStatus(fallbackText, false);
    return true;
  }

  function getItemCodeCandidates(value) {
    const raw = String(value || '').trim().toUpperCase();
    if (!raw) return [];

    const compact = raw.replace(/[^A-Z0-9]/g, '');
    const compactO0 = compact.replace(/O/g, '0');

    const candidates = new Set([raw, compact, compactO0]);

    [compact, compactO0].forEach((candidate) => {
      if (!candidate) return;
      if (/^\d+$/.test(candidate)) {
        candidates.add(candidate.replace(/^0+/, '') || '0');
      }
    });

    return Array.from(candidates).filter(Boolean);
  }
  function applyMaterialDescriptions(rows) {
    return rows.map((row) => {
      const catalogDescription = lookupDescriptionInCatalog(row.itemCode);
      const existingDescription = String(row.materialDescription || row.description || '').trim();
      const description = existingDescription || catalogDescription;
      return {
        ...row,
        description,
        materialDescription: description
      };
    });
  }

  function triggerDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }

  function getQueueProgressLabel() {
    const total = state.pdfQueueFiles.length;
    const index = state.pdfQueueIndex;
    if (!total || index < 0) return '';
    return ` [Drawing ${index + 1}/${total}]`;
  }

  function updateQueueStatusBadge() {
    if (!elements.queueStatusBadge) return;
    const total = state.pdfQueueFiles.length;
    const index = state.pdfQueueIndex;
    if (!total || index < 0) {
      elements.queueStatusBadge.style.display = 'none';
      return;
    }

    elements.queueStatusBadge.textContent = `Queue: Drawing ${index + 1}/${total}`;
    elements.queueStatusBadge.style.display = 'inline-flex';
  }

  function clearTicketWorkspaceForNextDrawing() {
    elements.ocrTableBody.innerHTML = '';
    setSelectedTableRow(null);
    clearCraftMissingHighlights();
    elements.generateBtn.disabled = true;
    if (elements.exportPdfBtn) elements.exportPdfBtn.disabled = true;
    if (elements.reprintBtn) elements.reprintBtn.disabled = true;
    elements.previewCraftSelect.innerHTML = '';
    elements.previewCraftSelect.disabled = true;
    elements.openPreviewWindowBtn.disabled = true;
    elements.ticketPreviewContainer.innerHTML = 'Generate picking tickets to load preview.';
    state.lastPreviewPayload = null;
    state.ticketNumbersReserved = false;
    state.currentRenderedTicketCount = 0;
    state.pendingTicketCommit = null;
    state.hasGeneratedCurrentDrawing = false;
    showDrawingProcessedButton(false);
    removeSelectionHighlights();
  }

  async function loadPdfFile(file, options) {
    const settings = options || {};
    if (!file) {
      setStatus('Select a PDF file first.', true);
      return false;
    }

    if (!window.pdfjsLib) {
      setStatus('PDF.js not loaded.', true);
      return false;
    }

    try {
      setStatus(`Loading PDF: ${file.name}${getQueueProgressLabel()}...`, false);
      clearTicketWorkspaceForNextDrawing();
      const buffer = await file.arrayBuffer();
      let loadingTask = window.pdfjsLib.getDocument({ data: buffer });
      try {
        state.pdfDoc = await loadingTask.promise;
      } catch (workerError) {
        const message = String(workerError?.message || workerError || '');
        const isWorkerSetupIssue = /worker|fake worker|cannot load script/i.test(message);
        if (!isWorkerSetupIssue) {
          throw workerError;
        }

        console.warn('PDF worker setup failed, retrying with worker disabled.', workerError);
        loadingTask = window.pdfjsLib.getDocument({ data: buffer, disableWorker: true });
        state.pdfDoc = await loadingTask.promise;
      }

      state.page = await state.pdfDoc.getPage(1);
      state.activePdfFileName = file.name || '';
      state.scale = 1.6;
      await renderPageAtScale(state.scale);
      state.selection = null;
      state.ocrAreaSelection = null;
      resetColumnSelections();
      elements.selectionOverlay.style.display = 'none';
      setColumnButtonsEnabled(false);
      activateOcrAreaSelectionMode();
      updateColumnStatus();

      if (settings.fromQueue) {
        setStatus(`Loaded ${file.name}${getQueueProgressLabel()}. Select OCR area/columns, run OCR, generate, then export to continue queue.`, false);
      } else {
        setStatus('PDF loaded. First select OCR Area and drag a block around the table. The app will auto-zoom; then capture each required column.', false);
      }
      return true;
    } catch (error) {
      console.error(error);
      setStatus(`Failed to load PDF: ${error.message}`, true);
      return false;
    }
  }

  function cancelPdfQueue() {
    state.pdfQueueFiles = [];
    state.pdfQueueIndex = -1;
    updateQueueStatusBadge();
  }

  async function startPdfQueueFromFolder(fileList) {
    const source = Array.from(fileList || []);
    const pdfFiles = source
      .filter((file) => file && /\.pdf$/i.test(file.name || ''))
      .sort((a, b) => (a.webkitRelativePath || a.name || '').localeCompare(b.webkitRelativePath || b.name || '', undefined, { numeric: true, sensitivity: 'base' }));

    if (!pdfFiles.length) {
      setStatus('No PDF files found in the selected folder.', true);
      return;
    }

    state.pdfQueueFiles = pdfFiles;
    state.pdfQueueIndex = 0;
    updateQueueStatusBadge();
    const loaded = await loadPdfFile(pdfFiles[0], { fromQueue: true });
    if (!loaded) {
      cancelPdfQueue();
    }
  }

  async function loadNextPdfInQueueIfAvailable() {
    const total = state.pdfQueueFiles.length;
    if (!total || state.pdfQueueIndex < 0) {
      return false;
    }

    const nextIndex = state.pdfQueueIndex + 1;
    if (nextIndex >= total) {
      setStatus('Queue complete: all selected folder PDFs have been processed.', false);
      cancelPdfQueue();
      return false;
    }

    state.pdfQueueIndex = nextIndex;
    updateQueueStatusBadge();
    const nextFile = state.pdfQueueFiles[nextIndex];
    await loadPdfFile(nextFile, { fromQueue: true });
    return true;
  }


  async function renderFirstPage() {
    const file = elements.pdfInput.files[0];
    cancelPdfQueue();
    await loadPdfFile(file, { fromQueue: false });
  }

  function formatDrawingNumberValue(value) {
    return String(value || '').toUpperCase();
  }

  function formatSheetNumberValue(value) {
    const raw = String(value || '').trim();
    if (!raw) return '';
    if (/^\d+$/.test(raw)) return raw.padStart(3, '0');
    return raw;
  }

  function formatRevisionValue(value) {
    const raw = String(value || '').trim().toUpperCase();
    if (!raw) return '';
    const match = raw.match(/^(\d+)([A-Z]*)$/);
    if (!match) return raw;
    const numeric = match[1].padStart(2, '0');
    const suffix = match[2] || '';
    return `${numeric}${suffix}`;
  }

  function applyHeaderFieldFormatting() {
    if (elements.drawingNumber) {
      elements.drawingNumber.value = formatDrawingNumberValue(elements.drawingNumber.value);
    }
    if (elements.sheetNo) {
      elements.sheetNo.value = formatSheetNumberValue(elements.sheetNo.value);
    }
    if (elements.revision) {
      elements.revision.value = formatRevisionValue(elements.revision.value);
    }
  }

  function viewerPoint(clientX, clientY) {
    const rect = elements.pdfCanvas.getBoundingClientRect();
    return {
      x: Math.max(0, Math.min(rect.width, clientX - rect.left)),
      y: Math.max(0, Math.min(rect.height, clientY - rect.top))
    };
  }

  function drawSelection(start, end) {
    const x = Math.min(start.x, end.x);
    const y = Math.min(start.y, end.y);
    const width = Math.abs(end.x - start.x);
    const height = Math.abs(end.y - start.y);

    elements.selectionOverlay.style.left = `${x}px`;
    elements.selectionOverlay.style.top = `${y}px`;
    elements.selectionOverlay.style.width = `${width}px`;
    elements.selectionOverlay.style.height = `${height}px`;
    elements.selectionOverlay.style.display = width > 2 && height > 2 ? 'block' : 'none';
  }

  async function storeSelection(start, end) {
    const canvasRect = elements.pdfCanvas.getBoundingClientRect();
    const scaleX = elements.pdfCanvas.width / canvasRect.width;
    const scaleY = elements.pdfCanvas.height / canvasRect.height;

    const left = Math.min(start.x, end.x);
    const top = Math.min(start.y, end.y);
    const width = Math.abs(end.x - start.x);
    const height = Math.abs(end.y - start.y);

    const selection = {
      x: Math.round(left * scaleX),
      y: Math.round(top * scaleY),
      width: Math.round(width * scaleX),
      height: Math.round(height * scaleY)
    };

    const valid = selection.width > 8 && selection.height > 8;
    if (!valid) return;

    state.selection = selection;
    if (state.selectionMode === 'ocrArea') {
      state.ocrAreaSelection = selection;
      resetColumnSelections();
      setColumnButtonsEnabled(true);
      updateColumnStatus();
      renderSelectionHighlights();
      await zoomToSelection(selection);
      activateOcrAreaSelectionMode();
      elements.selectionOverlay.style.display = 'none';
      setStatus('OCR area captured and zoomed. Now select each column button and drag each column.', false);
      return;
    }

    if (state.activeColumnKey) {
      state.columnSelections[state.activeColumnKey] = selection;
      updateColumnStatus();
      renderSelectionHighlights();
      setStatus(`${COLUMN_LABELS[state.activeColumnKey]} column captured (${getSelectedColumnsCount()}/4).`, false);
    } else {
      setStatus('Choose a column button first, then drag to capture that column.', true);
    }
  }

  function parseColumnValues(text) {
    return String(text || '')
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean);
  }

  function parseColumnValuesFromResult(columnKey, result, expectedCount) {
    const lineValues = Array.isArray(result?.lines)
      ? result.lines.map((line) => String(line || '').trim()).filter(Boolean)
      : [];

    let values = lineValues.length > 0 ? lineValues : parseColumnValues(result?.text || '');

    if (values.length <= 1 && expectedCount > 1) {
      const source = String(result?.text || '').trim();
      if (source) {
        const byWideSpacing = source
          .split(/\t+|\s{2,}/)
          .map((part) => part.trim())
          .filter(Boolean);

        if (byWideSpacing.length > values.length) {
          values = byWideSpacing;
        } else if (columnKey !== 'size') {
          const bySingleSpacing = source
            .split(/\s+/)
            .map((part) => part.trim())
            .filter(Boolean);
          if (bySingleSpacing.length > values.length) {
            values = bySingleSpacing;
          }
        }
      }
    }

    return values;
  }

  function normalizeColumnValues(values, expectedCount) {
    const normalized = values.slice(0, expectedCount);
    while (normalized.length < expectedCount) {
      normalized.push('');
    }
    return normalized;
  }

  async function recognizeColumn(columnKey, selection, lang) {
    const { x, y, width, height } = selection;
    const cropCanvas = document.createElement('canvas');
    cropCanvas.width = width;
    cropCanvas.height = height;
    const cropCtx = cropCanvas.getContext('2d');
    cropCtx.drawImage(elements.pdfCanvas, x, y, width, height, 0, 0, width, height);

    setStatus(`Running OCR for ${COLUMN_LABELS[columnKey]}...`, false);
    const ocrConfig = {
      tessedit_pageseg_mode: '6',
      preserve_interword_spaces: '1'
    };
    if (columnKey === 'pointNumber') {
      ocrConfig.tessedit_char_whitelist = '0123456789';
    } else if (columnKey === 'itemCode') {
      ocrConfig.tessedit_char_whitelist = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_/.';
    } else if (columnKey === 'size') {
      ocrConfig.tessedit_char_whitelist = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789./xX-';
    } else if (columnKey === 'quantity') {
      ocrConfig.tessedit_char_whitelist = '0123456789.,';
    }

    const result = await window.Tesseract.recognize(cropCanvas, lang, {
      ...ocrConfig,
      logger: (m) => {
        if (m.status === 'recognizing text' && typeof m.progress === 'number') {
          setStatus(`OCR ${COLUMN_LABELS[columnKey]}: ${Math.round(m.progress * 100)}%`, false);
        }
      }
    });

    return {
      text: result?.data?.text || '',
      lines: (result?.data?.lines || []).map((line) => line?.text || '')
    };
  }

  async function runOcrOnSelection() {
    if (!window.Tesseract) {
      setStatus('Tesseract is not available.', true);
      return;
    }

    const missingColumns = COLUMN_KEYS.filter((key) => !state.columnSelections[key]);
    if (missingColumns.length > 0) {
      setStatus(`Missing column selections: ${missingColumns.map((key) => COLUMN_LABELS[key]).join(', ')}`, true);
      return;
    }

    const expectedFromInput = parseInt(elements.itemCount.value || '0', 10);
    if (!expectedFromInput || expectedFromInput < 1) {
      setStatus('Enter a valid expected item count (minimum 1).', true);
      return;
    }

    try {
      const lang = (elements.ocrLanguage.value || 'eng').trim() || 'eng';
      const columnTextByKey = {};
      for (const key of COLUMN_KEYS) {
        columnTextByKey[key] = await recognizeColumn(key, state.columnSelections[key], lang);
      }

      const rawPointValues = parseColumnValuesFromResult('pointNumber', columnTextByKey.pointNumber, expectedFromInput);
      const rawItemCodeValues = parseColumnValuesFromResult('itemCode', columnTextByKey.itemCode, expectedFromInput);
      const rawSizeValues = parseColumnValuesFromResult('size', columnTextByKey.size, expectedFromInput);
      const rawQuantityValues = parseColumnValuesFromResult('quantity', columnTextByKey.quantity, expectedFromInput);

      const detectedMax = Math.max(
        rawPointValues.length,
        rawItemCodeValues.length,
        rawSizeValues.length,
        rawQuantityValues.length
      );
      const expectedCount = Math.max(expectedFromInput, detectedMax || 0);

      if (expectedCount < 1) {
        setStatus('Could not detect OCR rows. Re-select columns and retry.', true);
        return;
      }

      const pointValues = normalizeColumnValues(rawPointValues, expectedCount);
      const itemCodeValues = normalizeColumnValues(rawItemCodeValues, expectedCount);
      const sizeValues = normalizeColumnValues(rawSizeValues, expectedCount);
      const quantityValues = normalizeColumnValues(rawQuantityValues, expectedCount);

      const parsed = Array.from({ length: expectedCount }, (_, index) => ({
        pointNumber: normalizePointNumberValue(pointValues[index] || ''),
        itemCode: itemCodeValues[index] || '',
        size: sizeValues[index] || '',
        quantity: quantityValues[index] || '',
        materialDescription: '',
        projectNo: elements.projectNo.value.trim(),
        type: ''
      }));

      let appliedCorrections = 0;
      const corrected = parsed.map((row) => {
        const original = {
          pointNumber: row.pointNumber,
          itemCode: row.itemCode,
          size: row.size,
          quantity: row.quantity
        };
        const nextRow = { ...row };
        appliedCorrections += applyLearnedCorrections(nextRow);
        nextRow.pointNumber = normalizePointNumberValue(nextRow.pointNumber);
        return {
          ...nextRow,
          _original: original
        };
      });

      const correctedWithDescriptions = applyMaterialDescriptions(corrected);

      elements.ocrTableBody.innerHTML = '';
      setSelectedTableRow(null);
      correctedWithDescriptions.forEach((row) => addRow(row));
      elements.generateBtn.disabled = false;
      if (elements.reprintBtn) {
        elements.reprintBtn.disabled = true;
      }
      state.pendingTicketCommit = null;
      state.hasGeneratedCurrentDrawing = false;
      showDrawingProcessedButton(false);

      const mismatchMessages = [];
      const foundByColumn = {
        point: rawPointValues.length,
        itemCode: rawItemCodeValues.length,
        size: rawSizeValues.length,
        quantity: rawQuantityValues.length
      };
      if (foundByColumn.point !== expectedCount) mismatchMessages.push(`Point ${foundByColumn.point}/${expectedCount}`);
      if (foundByColumn.itemCode !== expectedCount) mismatchMessages.push(`Item Code ${foundByColumn.itemCode}/${expectedCount}`);
      if (foundByColumn.size !== expectedCount) mismatchMessages.push(`Size ${foundByColumn.size}/${expectedCount}`);
      if (foundByColumn.quantity !== expectedCount) mismatchMessages.push(`Quantity ${foundByColumn.quantity}/${expectedCount}`);

      if (mismatchMessages.length > 0) {
        const correctionText = appliedCorrections > 0 ? ` Auto-corrected ${appliedCorrections} value(s) from local learning.` : '';
        setStatus(`OCR completed with partial results. Loaded ${expectedCount} editable row(s); review blanks. (${mismatchMessages.join(', ')})${correctionText}`, true);
      } else {
        const correctionText = appliedCorrections > 0 ? ` Auto-corrected ${appliedCorrections} value(s) from local learning.` : '';
        const autoDetectedText = expectedFromInput !== expectedCount ? ` Auto-detected ${expectedCount} row(s).` : '';
        setStatus(`Column OCR completed: ${corrected.length} row(s) loaded into editable table.${autoDetectedText}${correctionText}`, false);
      }
    } catch (error) {
      console.error(error);
      setStatus(`OCR failed: ${error.message}`, true);
    }
  }

  function invalidateGeneratedTicketState() {
    state.pendingTicketCommit = null;
    state.hasGeneratedCurrentDrawing = false;
    showDrawingProcessedButton(false);
    if (elements.reprintBtn) {
      elements.reprintBtn.disabled = true;
    }
  }

  function fillCellFromAbove(row, columnIndex) {
    if (!row || columnIndex < 0) return false;
    const previousRow = row.previousElementSibling;
    if (!previousRow) return false;
    const previousCells = previousRow.querySelectorAll('td');
    const currentCells = row.querySelectorAll('td');
    const sourceCell = previousCells[columnIndex];
    const targetCell = currentCells[columnIndex];
    if (!sourceCell || !targetCell) return false;
    targetCell.textContent = sourceCell.textContent || '';
    targetCell.dispatchEvent(new Event('input', { bubbles: true }));
    return true;
  }

  async function processDrawingAndMoveNext() {
    if (!state.pendingTicketCommit) {
      setStatus('No generated ticket batch is waiting for processing.', true);
      return;
    }

    commitTicketNumbers(state.pendingTicketCommit.projectNo, state.pendingTicketCommit.ticketFinalNumber);
    clearTicketWorkspaceForNextDrawing();
    elements.viewerWrap.scrollTop = 0;
    window.scrollTo({ top: 0, behavior: 'smooth' });

    const moved = await loadNextPdfInQueueIfAvailable();
    if (!moved) {
      if (state.page) {
        await resetOcrAreaAndZoom();
      }
      setStatus('Drawing processed. Ready to start the next drawing.', false);
    }
  }

  function getPromptedReprintTicketNumber() {
    const entered = window.prompt('Enter the picking ticket number to reprint:', '');
    if (entered === null) return '';
    const digits = String(entered).replace(/\D+/g, '');
    if (!digits) return '';
    return digits.padStart(6, '0');
  }

  elements.renderBtn?.addEventListener('click', renderFirstPage);
  elements.pdfInput?.addEventListener('change', async () => {
    if (!elements.pdfInput.files || !elements.pdfInput.files[0]) return;
    await renderFirstPage();
  });
  elements.loadFolderBtn?.addEventListener('click', () => {
    if (!elements.pdfFolderInput) return;
    elements.pdfFolderInput.value = '';
    elements.pdfFolderInput.click();
  });
  elements.pdfFolderInput?.addEventListener('change', async () => {
    const files = elements.pdfFolderInput.files;
    await startPdfQueueFromFolder(files);
  });
  elements.selectOcrAreaBtn.addEventListener('click', activateOcrAreaSelectionMode);
  elements.resetOcrAreaBtn.addEventListener('click', () => {
    resetOcrAreaAndZoom().catch((error) => {
      console.error(error);
      setStatus(`Reset failed: ${error.message}`, true);
    });
  });
  elements.selectPointColumnBtn.addEventListener('click', () => setActiveColumnSelection('pointNumber'));
  elements.selectItemCodeColumnBtn.addEventListener('click', () => setActiveColumnSelection('itemCode'));
  elements.selectSizeColumnBtn.addEventListener('click', () => setActiveColumnSelection('size'));
  elements.selectQuantityColumnBtn.addEventListener('click', () => setActiveColumnSelection('quantity'));
  elements.itemCount.addEventListener('input', updateRunOcrAvailability);
  elements.viewerWrap.addEventListener('wheel', (event) => {
    if (!elements.lockCanvasScroll?.checked) return;
    event.preventDefault();
  }, { passive: false });

  elements.runOcrBtn.addEventListener('click', runOcrOnSelection);
  elements.manageItemCodesBtn?.addEventListener('click', openItemCodeFlyout);
  elements.closeItemCodeFlyoutBtn?.addEventListener('click', closeItemCodeFlyout);
  elements.itemCodeFlyoutOverlay?.addEventListener('click', closeItemCodeFlyout);
  elements.addItemCodeBtn?.addEventListener('click', upsertItemCodeFromFlyout);
  elements.itemCodeSearchInput?.addEventListener('input', renderItemCodeFlyoutRows);
  [elements.itemCodeInput, elements.itemDescriptionInput].forEach((input) => {
    input?.addEventListener('keydown', (event) => {
      if (event.key !== 'Enter') return;
      event.preventDefault();
      upsertItemCodeFromFlyout();
    });
  });
  elements.itemCodeCatalogTableBody?.addEventListener('click', (event) => {
    handleItemCodeFlyoutTableClick(event.target);
  });
  elements.addRowBtn.addEventListener('click', () => {
    const rowsBefore = elements.ocrTableBody.querySelectorAll('tr').length;
    addRow({});
    const addedRow = elements.ocrTableBody.querySelectorAll('tr')[rowsBefore];
    if (addedRow) {
      setSelectedTableRow(addedRow);
    }
    invalidateGeneratedTicketState();
  });
  elements.deleteRowBtn?.addEventListener('click', deleteSelectedTableRow);
  elements.ocrTableBody.addEventListener('click', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const row = target.closest('tr');
    if (!row) return;
    setSelectedTableRow(row);
  });
  elements.ocrTableBody.addEventListener('keydown', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement) || target.tagName !== 'TD') return;
    if (!(event.ctrlKey || event.metaKey) || event.key.toLowerCase() !== 'd') return;

    event.preventDefault();
    const row = target.closest('tr');
    if (!row) return;
    const cells = Array.from(row.querySelectorAll('td'));
    const columnIndex = cells.indexOf(target);
    if (columnIndex < 0 || columnIndex > 5) return;

    const filled = fillCellFromAbove(row, columnIndex);
    if (!filled) {
      setStatus('Ctrl+D needs a row above to copy from.', true);
      return;
    }

    setStatus('Ctrl+D applied: copied value from row above.', false);
  });
  elements.ocrTableBody.addEventListener('input', (event) => {
    const cell = event.target;
    if (!(cell instanceof HTMLElement) || cell.tagName !== 'TD') return;

    const row = cell.closest('tr');
    if (!row) return;
    setSelectedTableRow(row);

    const cells = Array.from(row.querySelectorAll('td'));
    const columnIndex = cells.indexOf(cell);
    if (columnIndex === 0) {
      const normalized = normalizePointNumberValue(cell.textContent || '');
      if ((cell.textContent || '') !== normalized) {
        cell.textContent = normalized;
      }
    } else if (columnIndex === 1) {
      updateMaterialDescriptionForTableRow(row);
      const isManualType = row.dataset.typeManual === 'true';
      if (!isManualType) {
        applyRememberedCraftTypeToRow(row, { preserveExisting: false });
      }

      const selectedType = getRowSelectedType(row);
      if (selectedType) {
        rememberCraftTypeForItemCode(cell.textContent.trim(), selectedType);
      }
    } else if (columnIndex === 4) {
      const cleanDescription = String(cell.textContent || '').trim();
      setDescriptionNeedsAttention(cell, cleanDescription === PLACEHOLDER_DESCRIPTION_TEXT || !cleanDescription);
    }

    if (elements.exportPdfBtn) {
      elements.exportPdfBtn.disabled = true;
    }
    if (elements.reprintBtn) {
      elements.reprintBtn.disabled = true;
    }
    invalidateGeneratedTicketState();
  });
  elements.ocrTableBody.addEventListener('change', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) return;
    if (target.type !== 'radio') return;

    const row = target.closest('tr');
    if (!row) return;
    setSelectedTableRow(row);

    const itemCode = row.querySelectorAll('td')[1]?.textContent.trim() || '';
    if (itemCode) {
      rememberCraftTypeForItemCode(itemCode, target.value);
    }
    row.dataset.typeManual = 'true';

    row.classList.remove('craft-missing');

    if (elements.exportPdfBtn) {
      elements.exportPdfBtn.disabled = true;
    }
    if (elements.reprintBtn) {
      elements.reprintBtn.disabled = true;
    }
    invalidateGeneratedTicketState();
  });
  elements.previewCraftSelect.addEventListener('change', () => {
    if (!state.lastPreviewPayload) return;
    renderTicketPreview(state.lastPreviewPayload, elements.previewCraftSelect.value);
  });
  elements.openPreviewWindowBtn.addEventListener('click', openTicketPreviewWindow);
  elements.generateBtn.addEventListener('click', async () => {
    if (!validateCraftSelectionBeforeGeneration()) {
      alert('Each populated row must have a craft selected before generating tickets.');
      return;
    }

    const codePromptResult = ensureUsableItemCodesInTable();
    if (codePromptResult.placeholderCount > 0) {
      alert(`${codePromptResult.placeholderCount} row(s) were assigned ${PLACEHOLDER_ITEM_CODE} placeholders. These placeholders are not learned and must be replaced with valid codes before successful export.`);
    }

    const rows = collectRows();
    if (rows.length === 0) {
      alert('No rows available to generate tickets.');
      return;
    }

    await loadItemCodeDescriptionCatalog();
    if (!await ensureCatalogAvailableForGeneration()) {
      return;
    }
    let rowsWithDescriptions = applyMaterialDescriptions(rows);

    let missingCatalogCodes = findMissingCatalogItemCodes(rowsWithDescriptions);
    if (missingCatalogCodes.length > 0) {
      const result = promptToAddMissingItemCodes(missingCatalogCodes);
      if (result.added > 0) {
        rowsWithDescriptions = applyMaterialDescriptions(rowsWithDescriptions);
        const tableRows = Array.from(elements.ocrTableBody.querySelectorAll('tr'));
        tableRows.forEach((tableRow) => updateMaterialDescriptionForTableRow(tableRow));
      }

      missingCatalogCodes = findMissingCatalogItemCodes(rowsWithDescriptions);
      if (missingCatalogCodes.length > 0) {
        alert(
          'Some item code(s) are still unrecognized. Add descriptions for these codes to continue:\n\n' +
          missingCatalogCodes.join(', ')
        );
        setStatus('Generation blocked: unresolved item code(s) in local catalog.', true);
        return;
      }
    }

    const correctedDescriptions = enforceExactCatalogDescriptionsInTable();
    if (correctedDescriptions > 0) {
      rowsWithDescriptions = applyMaterialDescriptions(collectRows());
      setStatus(`Corrected ${correctedDescriptions} material description(s) to exact catalog values.`, false);
    }

    const payload = buildTicketPayload(rowsWithDescriptions);
    if (!payload) return;

    renderTicketPreview(payload);

    state.ticketNumbersReserved = payload.tickets.length > 0;
    state.currentRenderedTicketCount = payload.tickets.length;

    if (elements.exportPdfBtn) {
      elements.exportPdfBtn.disabled = payload.tickets.length === 0;
    }
    if (elements.reprintBtn) {
      elements.reprintBtn.disabled = payload.tickets.length === 0;
    }

    try {
      const success = await exportCraftPdfsFromHtml(payload, { isReprint: false, preferDirectoryPicker: true });
      if (success) {
        state.pendingTicketCommit = {
          projectNo: payload.projectNo,
          ticketFinalNumber: payload.ticketFinalNumber
        };
        state.hasGeneratedCurrentDrawing = true;
        showDrawingProcessedButton(true);
        if (elements.reprintBtn) {
          elements.reprintBtn.disabled = false;
        }
        setStatus('PDFs generated. Click Drawing Processed to reset this drawing and move to the next one.', false);
      }
    } catch (error) {
      console.error(error);
      setStatus(`Automatic PDF export failed: ${error.message}`, true);
    }
  });
  elements.exportPdfBtn?.addEventListener('click', async () => {
    if (!state.lastPreviewPayload) {
      setStatus('Generate ticket data first.', true);
      return;
    }

    try {
      const success = await exportCraftPdfsFromHtml(state.lastPreviewPayload, { isReprint: false, preferDirectoryPicker: true });
      if (success) {
        state.pendingTicketCommit = {
          projectNo: state.lastPreviewPayload.projectNo,
          ticketFinalNumber: state.lastPreviewPayload.ticketFinalNumber
        };
        state.hasGeneratedCurrentDrawing = true;
        showDrawingProcessedButton(true);
        if (elements.reprintBtn) {
          elements.reprintBtn.disabled = false;
        }
      }
    } catch (error) {
      console.error(error);
      setStatus(`HTML PDF export failed: ${error.message}`, true);
    }
  });
  elements.reprintBtn?.addEventListener('click', async () => {
    if (!state.lastPreviewPayload) {
      setStatus('Generate ticket data first.', true);
      return;
    }

    const reprintTicketNumber = getPromptedReprintTicketNumber();
    if (!reprintTicketNumber) {
      setStatus('Reprint canceled: valid picking ticket number not provided.', true);
      return;
    }

    try {
      await exportCraftPdfsFromHtml(state.lastPreviewPayload, {
        isReprint: true,
        preferDirectoryPicker: true,
        ticketNumberOverride: reprintTicketNumber
      });
    } catch (error) {
      console.error(error);
      setStatus(`Reprint failed: ${error.message}`, true);
    }
  });
  elements.drawingProcessedBtn?.addEventListener('click', () => {
    processDrawingAndMoveNext().catch((error) => {
      console.error(error);
      setStatus(`Drawing reset failed: ${error.message}`, true);
    });
  });
  elements.projectNo.addEventListener('change', () => {
    const projectNo = String(elements.projectNo?.value || '').trim() || 'default';
    syncTicketStartField(projectNo);
  });

  elements.ticketStartNo.addEventListener('change', () => {
    const projectNo = String(elements.projectNo?.value || '').trim() || 'default';
    const entered = parseInt(String(elements.ticketStartNo?.value || '').trim(), 10);
    if (!Number.isFinite(entered) || entered < 1) {
      syncTicketStartField(projectNo);
      return;
    }

    setStoredLastTicketNumber(projectNo, entered - 1);
    syncTicketStartField(projectNo);
  });

  elements.drawingNumber?.addEventListener('input', () => {
    elements.drawingNumber.value = formatDrawingNumberValue(elements.drawingNumber.value);
  });
  elements.sheetNo?.addEventListener('change', () => {
    elements.sheetNo.value = formatSheetNumberValue(elements.sheetNo.value);
  });
  elements.revision?.addEventListener('change', () => {
    elements.revision.value = formatRevisionValue(elements.revision.value);
  });

  elements.pdfCanvas.addEventListener('mousedown', (event) => {
    if (!state.page) return;

    if (event.button === 0 && event.shiftKey) {
      state.isPanning = true;
      state.panStart = { x: event.clientX, y: event.clientY };
      state.panScrollStart = {
        left: elements.viewerWrap.scrollLeft,
        top: elements.viewerWrap.scrollTop
      };
      elements.viewerWrap.style.cursor = 'grabbing';
      return;
    }

    state.isDragging = true;
    state.dragStart = viewerPoint(event.clientX, event.clientY);
    drawSelection(state.dragStart, state.dragStart);
  });

  window.addEventListener('mousemove', (event) => {
    if (state.isPanning && state.panStart && state.panScrollStart) {
      const deltaX = event.clientX - state.panStart.x;
      const deltaY = event.clientY - state.panStart.y;
      elements.viewerWrap.scrollLeft = state.panScrollStart.left - deltaX;
      elements.viewerWrap.scrollTop = state.panScrollStart.top - deltaY;
      return;
    }

    if (!state.isDragging || !state.dragStart) return;
    const current = viewerPoint(event.clientX, event.clientY);
    drawSelection(state.dragStart, current);
  });

  window.addEventListener('mouseup', (event) => {
    if (state.isPanning) {
      state.isPanning = false;
      state.panStart = null;
      state.panScrollStart = null;
      elements.viewerWrap.style.cursor = 'default';
      return;
    }

    if (!state.isDragging || !state.dragStart) return;
    state.isDragging = false;
    const end = viewerPoint(event.clientX, event.clientY);
    drawSelection(state.dragStart, end);
    storeSelection(state.dragStart, end).catch((error) => {
      console.error(error);
      setStatus(`Selection failed: ${error.message}`, true);
    });
    state.dragStart = null;
  });

  setColumnButtonsEnabled(false);
  state.learningStore = loadLearningStore();
  state.craftByItemCodeStore = loadCraftByItemCodeStore();
  if (elements.exportPdfBtn) {
    elements.exportPdfBtn.disabled = true;
  }
  if (elements.reprintBtn) {
    elements.reprintBtn.disabled = true;
  }
  loadItemCodeDescriptionCatalog().then((catalog) => {
    const count = Object.keys(catalog || {}).length;
    renderItemCodeFlyoutRows();
    if (count > 0) {
      setStatus(`Item code catalog loaded (${count} codes). Load a PDF and run OCR.`, false);
    }
  }).catch((error) => {
    console.warn('Catalog initialization failed.', error);
  });
  applyHeaderFieldFormatting();
  syncTicketStartField(String(elements.projectNo?.value || '').trim() || 'default');
  updateColumnStatus();
})();
