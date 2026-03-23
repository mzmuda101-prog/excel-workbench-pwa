const rootEl = document.documentElement;
const logEl = document.getElementById("log");
const statusEl = document.getElementById("status");
const tableEl = document.getElementById("dataTable");
const theadEl = tableEl.querySelector("thead");
const tbodyEl = tableEl.querySelector("tbody");
const tableWrapEl = document.getElementById("tableWrap");
const tableScrollbarEl = document.getElementById("tableScrollbar");
const tableScrollbarInnerEl = document.getElementById("tableScrollbarInner");
const emptyStateEl = document.getElementById("emptyState");
const emptyTitleEl = document.getElementById("emptyTitle");
const emptySubEl = document.getElementById("emptySub");
const DEFAULT_EMPTY_TITLE = emptyTitleEl.textContent;
const DEFAULT_EMPTY_SUB = emptySubEl.textContent;

const fileInput = document.getElementById("fileInput");
const dropZone = document.getElementById("dropZone");
const fileNameEl = document.getElementById("fileName");
const fileNameTextEl = document.getElementById("fileNameText");
const sheetSelect = document.getElementById("sheetSelect");
const headerRowEl = document.getElementById("headerRow");
const autoHeaderRowEl = document.getElementById("autoHeaderRow");
const displayModeEl = document.getElementById("displayMode");
const maxRowsEl = document.getElementById("maxRows");
const excelLayoutToggleEl = document.getElementById("excelLayoutToggle");
const loadBtn = document.getElementById("loadBtn");

const searchQueryEl = document.getElementById("searchQuery");
const searchQuery2El = document.getElementById("searchQuery2");
const filterModeEl = document.getElementById("filterMode");
const filterMode2El = document.getElementById("filterMode2");
const filter1ColumnsEl = document.getElementById("filter1Columns");
const filter2ColumnsEl = document.getElementById("filter2Columns");
const filter1PickBtn = document.getElementById("filter1Pick");
const filter2PickBtn = document.getElementById("filter2Pick");
const onlyNonEmptyEl = document.getElementById("onlyNonEmpty");
const applyFilterBtn = document.getElementById("applyFilterBtn");
const filterBadgeEl = document.getElementById("filterBadge");

const dateModeEl = document.getElementById("dateMode");
const dateFromEl = document.getElementById("dateFrom");
const dateToEl = document.getElementById("dateTo");
const lastDaysEl = document.getElementById("lastDays");
const dateColumnsEl = document.getElementById("dateColumns");
const datePickBtn = document.getElementById("datePick");

const resetFiltersBtn = document.getElementById("resetFiltersBtn");
const exportCsvBtn = document.getElementById("exportCsvBtn");
const saveBtn = document.getElementById("saveBtn");
const saveAsBtn = document.getElementById("saveAsBtn");
const resetWidthsBtn = document.getElementById("resetWidthsBtn");
const resetSortBtn = document.getElementById("resetSortBtn");
const readingToggle = document.getElementById("readingToggle");
const quickSearchWrap = document.getElementById("quickSearchWrap");
const quickSearchEl = document.getElementById("quickSearch");
const quickSearchColumnsBtn = document.getElementById("quickSearchColumnsBtn");
const quickSearchBtn = document.getElementById("quickSearchBtn");

const columnPickerEl = document.getElementById("columnPicker");
const columnPickerTitleEl = document.getElementById("columnPickerTitle");
const columnListEl = document.getElementById("columnList");
const columnSearchEl = document.getElementById("columnSearch");
const selectAllBtn = document.getElementById("selectAllBtn");
const clearAllBtn = document.getElementById("clearAllBtn");
const applyPickBtn = document.getElementById("applyPickBtn");
const closePickerBtn = document.getElementById("closePicker");

const themeToggle = document.getElementById("themeToggle");
const panelToggle = document.getElementById("panelToggle");
const panelHandle = document.getElementById("panelHandle");
const sidebarEl = document.querySelector(".sidebar");
const sidebarScrim = document.getElementById("sidebarScrim");
const brandRefreshBtn = document.getElementById("brandRefresh");
const networkBadgeEl = document.getElementById("networkBadge");
const heroRightEl = document.getElementById("heroRight");
const loadingOverlayEl = document.getElementById("loadingOverlay");
const loadingTextEl = document.getElementById("loadingText");
const toastContainerEl = document.getElementById("toastContainer");
const cellTooltipEl = document.getElementById("cellTooltip");
const quickSearchPopupEl = document.getElementById("quickSearchPopup");
const quickSearchPopupInput = document.getElementById("quickSearchPopupInput");
const quickSearchPopupColumnsBtn = document.getElementById("quickSearchPopupColumnsBtn");
const quickSearchPopupBtn = document.getElementById("quickSearchPopupBtn");

const quickRangeButtons = Array.from(document.querySelectorAll(".chip[data-range]"));

let workbook = null;
let currentHeaders = [];
let baseRows = [];
let viewRows = [];
let currentFileName = "";
let currentSheetName = "";
let currentHeaderRow = 1;
let currentStartCol = 0;
let currentMerges = [];
let currentHeaderStyles = [];
let currentSheetColWidths = [];
let currentSheetRowHeights = {};
const columnSelections = {
  filter1: new Set(),
  filter2: new Set(),
  date: new Set(),
};
let activePickerKey = null;
let lastPickerTriggerEl = null;
let sortState = { col: "", dir: "asc" };
let manualColumnWidths = {};
let hasUnsavedChanges = false;
let syncingHorizontalScroll = false;
let tooltipHideTimer = null;

const THEME_KEY = "excel-workbench-theme";
const MAX_ROWS_KEY = "excel-workbench-max-rows";
const EXCEL_LAYOUT_KEY = "excel-workbench-excel-layout";
const INTRO_PLAYED_KEY = "introPlayed";
const BASE_TITLE = document.title || "Excel Workbench";

function log(msg, type = "info") {
  const line = document.createElement("div");
  line.className = `log-line log-${type}`;
  line.textContent = `${new Date().toLocaleTimeString()} ${msg}`;
  logEl.prepend(line);
}

function toast(msg, type = "info") {
  const toastEl = document.createElement("div");
  toastEl.className = `toast ${type}`;

  const icon = document.createElement("div");
  icon.className = "toast-icon";
  icon.textContent = type === "success" ? "✓" : type === "error" ? "!" : type === "warning" ? "!" : "i";

  const label = document.createElement("div");
  label.textContent = msg;

  toastEl.appendChild(icon);
  toastEl.appendChild(label);
  toastContainerEl.appendChild(toastEl);

  setTimeout(() => {
    toastEl.classList.add("out");
    setTimeout(() => toastEl.remove(), 200);
  }, 2800);
}

function setLoading(isLoading, text) {
  if (text) loadingTextEl.textContent = text;
  loadingOverlayEl.classList.toggle("hidden", !isLoading);
}

function setStatus(msg) {
  statusEl.textContent = msg;
}

function setDirtyState(isDirty) {
  hasUnsavedChanges = !!isDirty;
  statusEl.classList.toggle("unsaved", hasUnsavedChanges);
  document.title = hasUnsavedChanges ? `* ${BASE_TITLE}` : BASE_TITLE;
}

function valuesEqual(a, b) {
  if (a === b) return true;
  if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
  if ((a === null || a === undefined) && (b === null || b === undefined)) return true;
  return false;
}

function isXlsxAvailable(showFeedback = false) {
  const available = typeof window !== "undefined" && !!window.XLSX;
  if (!available && showFeedback) {
    setStatus("Brak biblioteki XLSX");
    toast("Brak biblioteki XLSX. Odśwież stronę lub sprawdź połączenie.", "error");
    log("Brak biblioteki XLSX (window.XLSX).", "error");
  }
  return available;
}

function setRuntimeAvailability(isAvailable) {
  fileInput.disabled = !isAvailable;
  loadBtn.disabled = !isAvailable;
  saveAsBtn.disabled = !isAvailable;
  if (excelLayoutToggleEl) {
    excelLayoutToggleEl.disabled = !isAvailable;
    excelLayoutToggleEl.setAttribute("aria-disabled", isAvailable ? "false" : "true");
  }
  saveAsBtn.setAttribute("aria-disabled", isAvailable ? "false" : "true");
  dropZone.classList.toggle("disabled", !isAvailable);
  dropZone.setAttribute("aria-disabled", isAvailable ? "false" : "true");
}

function loadMaxRowsPreference() {
  const saved = localStorage.getItem(MAX_ROWS_KEY);
  const value = saved ? parseInt(saved, 10) : null;
  if (value && Number.isFinite(value) && value > 0) {
    maxRowsEl.value = String(value);
  }
}

function saveMaxRowsPreference() {
  const value = Math.max(1, parseInt(maxRowsEl.value || "200", 10));
  localStorage.setItem(MAX_ROWS_KEY, String(value));
}

function loadExcelLayoutPreference() {
  if (!excelLayoutToggleEl) return;
  setExcelLayoutEnabled(localStorage.getItem(EXCEL_LAYOUT_KEY) === "1");
}

function saveExcelLayoutPreference() {
  if (!excelLayoutToggleEl) return;
  localStorage.setItem(EXCEL_LAYOUT_KEY, isExcelLayoutEnabled() ? "1" : "0");
}

function isExcelLayoutEnabled() {
  if (!excelLayoutToggleEl) return false;
  return excelLayoutToggleEl.getAttribute("aria-pressed") === "true";
}

function setExcelLayoutEnabled(enabled) {
  if (!excelLayoutToggleEl) return;
  const next = !!enabled;
  excelLayoutToggleEl.setAttribute("aria-pressed", next ? "true" : "false");
  excelLayoutToggleEl.classList.toggle("active", next);
  excelLayoutToggleEl.textContent = next ? "Widok Excel: ON" : "Widok Excel";
}

function setEmptyState(title, subtitle) {
  emptyTitleEl.textContent = title;
  emptySubEl.textContent = subtitle;
  emptyStateEl.classList.remove("hidden");
  tableWrapEl.classList.add("hidden");
}

function showTable() {
  emptyStateEl.classList.add("hidden");
  tableWrapEl.classList.remove("hidden");
  if (tableScrollbarEl) tableScrollbarEl.classList.remove("hidden");
}

function hideCellTooltip() {
  if (!cellTooltipEl) return;
  if (tooltipHideTimer) {
    clearTimeout(tooltipHideTimer);
    tooltipHideTimer = null;
  }
  cellTooltipEl.classList.add("hidden");
  cellTooltipEl.textContent = "";
}

function getTooltipText(cell) {
  if (!cell) return "";
  return (cell.dataset.fullText || cell.textContent || "").trim();
}

function isCellTextTruncated(cell) {
  if (!cell) return false;
  return cell.scrollWidth - cell.clientWidth > 1;
}

function positionCellTooltip(cell) {
  if (!cellTooltipEl || !cell) return;
  const rect = cell.getBoundingClientRect();
  const tooltipRect = cellTooltipEl.getBoundingClientRect();
  const margin = 12;
  let left = rect.left + rect.width / 2 - tooltipRect.width / 2;
  left = Math.max(margin, Math.min(left, window.innerWidth - tooltipRect.width - margin));
  let top = rect.top - tooltipRect.height - 10;
  if (top < margin) top = Math.min(window.innerHeight - tooltipRect.height - margin, rect.bottom + 10);
  cellTooltipEl.style.left = `${left}px`;
  cellTooltipEl.style.top = `${top}px`;
}

function showCellTooltip(cell, persistent = false) {
  if (!cellTooltipEl || !isCellTextTruncated(cell)) return;
  const text = getTooltipText(cell);
  if (!text) return;
  hideCellTooltip();
  cellTooltipEl.textContent = text;
  cellTooltipEl.classList.remove("hidden");
  positionCellTooltip(cell);
  if (!persistent) return;
  tooltipHideTimer = window.setTimeout(() => {
    hideCellTooltip();
  }, 2200);
}

function syncHorizontalScrollbar() {
  if (!tableWrapEl || !tableScrollbarEl || !tableScrollbarInnerEl) return;
  const active = !tableWrapEl.classList.contains("hidden") && tableWrapEl.scrollWidth > tableWrapEl.clientWidth + 1;
  tableScrollbarEl.classList.toggle("hidden", !active);
  if (!active) return;
  tableScrollbarInnerEl.style.width = `${tableWrapEl.scrollWidth}px`;
  if (Math.abs(tableScrollbarEl.scrollLeft - tableWrapEl.scrollLeft) > 1) {
    tableScrollbarEl.scrollLeft = tableWrapEl.scrollLeft;
  }
}

function toDisplay(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) {
    const dd = String(value.getDate()).padStart(2, "0");
    const mm = String(value.getMonth() + 1).padStart(2, "0");
    const yy = String(value.getFullYear()).slice(-2);
    return `${dd}-${mm}-${yy}`;
  }
  return String(value);
}

function getDisplayValue(row, index) {
  if (row && Array.isArray(row.display) && index < row.display.length) {
    return row.display[index];
  }
  if (row && Array.isArray(row.values) && index < row.values.length) {
    return toDisplay(row.values[index]);
  }
  return "";
}

function parseDateFlexible(value) {
  if (value instanceof Date) return value;
  if (typeof value === "number" && Number.isFinite(value)) {
    const ms = (value - 25569) * 86400000;
    const d = new Date(ms);
    return Number.isNaN(d.getTime()) ? null : d;
  }
  if (typeof value !== "string") return null;
  let v = value.trim();
  if (!v) return null;

  if (/^\d+(\.\d+)?$/.test(v)) {
    const numeric = Number(v);
    if (Number.isFinite(numeric)) {
      const ms = (numeric - 25569) * 86400000;
      const d = new Date(ms);
      return Number.isNaN(d.getTime()) ? null : d;
    }
  }

  v = v.replace(/T.*$/, "");
  v = v.replace(/\s+\d{1,2}:\d{2}(:\d{2})?.*$/, "");
  const normalized = v.replace(/[.\/]/g, "-");

  const monthMap = {
    // PL
    "sty": 1, "stycz": 1, "stycznia": 1,
    "lut": 2, "lutego": 2,
    "mar": 3, "marca": 3,
    "kwi": 4, "kwie": 4, "kwietnia": 4,
    "maj": 5, "maja": 5,
    "cze": 6, "czer": 6, "czerwca": 6,
    "lip": 7, "lipca": 7,
    "sie": 8, "sier": 8, "sierpnia": 8,
    "wrz": 9, "wrzes": 9, "wrzesnia": 9,
    "paź": 10, "paz": 10, "paźdz": 10, "pazdz": 10, "października": 10, "pazdziernika": 10,
    "lis": 11, "list": 11, "listopada": 11,
    "gru": 12, "grud": 12, "grudnia": 12,
    // EN
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
  };

  let m = normalized.match(/^(\d{1,2})-(\d{1,2})-(\d{2}|\d{4})$/);
  if (m) {
    const y = m[3].length === 2 ? Number(`20${m[3]}`) : Number(m[3]);
    return new Date(y, Number(m[2]) - 1, Number(m[1]));
  }

  m = normalized.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) {
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }

  const words = v.toLowerCase()
    .replace(/,/g, "")
    .replace(/\s+/g, " ")
    .trim();
  let wm = words.match(/^(\d{1,2})\s+([a-ząćęłńóśźż\.]+)\s+(\d{4}|\d{2})$/i);
  if (wm) {
    const day = Number(wm[1]);
    const monthKey = wm[2].replace(/\.$/, "");
    const month = monthMap[monthKey];
    const year = wm[3].length === 2 ? Number(`20${wm[3]}`) : Number(wm[3]);
    if (month) return new Date(year, month - 1, day);
  }
  wm = words.match(/^([a-ząćęłńóśźż\.]+)\s+(\d{1,2})\s+(\d{4}|\d{2})$/i);
  if (wm) {
    const monthKey = wm[1].replace(/\.$/, "");
    const month = monthMap[monthKey];
    const day = Number(wm[2]);
    const year = wm[3].length === 2 ? Number(`20${wm[3]}`) : Number(wm[3]);
    if (month) return new Date(year, month - 1, day);
  }

  const parsed = Date.parse(value);
  if (!Number.isNaN(parsed)) return new Date(parsed);
  return null;
}

function parseInputValue(raw) {
  const text = String(raw ?? "").trim();
  if (!text) return null;
  if (text.startsWith("=")) return { value: text, type: "formula" };
  const asDate = parseDateFlexible(text);
  if (asDate) return { value: asDate, type: "date" };
  if (/^-?\d+(\.\d+)?$/.test(text)) return { value: Number(text), type: "number" };
  return { value: text, type: "string" };
}

function updateSheetCell(rowIndex0, colIndex0, parsed) {
  if (!workbook || !currentSheetName) return;
  const sheet = workbook.Sheets[currentSheetName];
  if (!sheet) return;
  const absoluteCol = currentStartCol + colIndex0;
  const cellRef = XLSX.utils.encode_cell({ r: rowIndex0, c: absoluteCol });
  if (!parsed || parsed.value === null) {
    delete sheet[cellRef];
    return;
  }
  if (parsed.type === "formula") {
    toast("Edycja formul jest zablokowana", "warning");
    return;
  }
  if (parsed.type === "date") {
    sheet[cellRef] = { v: parsed.value, t: "d" };
    return;
  }
  if (parsed.type === "number") {
    sheet[cellRef] = { v: parsed.value, t: "n" };
    return;
  }
  sheet[cellRef] = { v: parsed.value, t: "s" };
}

function getDateRange() {
  const mode = dateModeEl.value;
  if (mode === "last_n_days") {
    const days = Math.max(1, parseInt(lastDaysEl.value || "30", 10));
    const now = new Date();
    const from = new Date(now.getFullYear(), now.getMonth(), now.getDate() - days);
    return { from, to: now };
  }
  const from = parseDateFlexible(dateFromEl.value);
  const to = parseDateFlexible(dateToEl.value);
  if (mode === "before") return { from: null, to };
  if (mode === "after") return { from, to: null };
  return { from, to };
}

function rowMatchesTextFilter(row, criteria, onlyNonEmpty) {
  const values = row.values;
  let usedIndexes = new Set();
  criteria.forEach((c) => c.indexes.forEach((i) => usedIndexes.add(i)));
  const indexes = usedIndexes.size ? Array.from(usedIndexes) : values.map((_, i) => i);

  if (onlyNonEmpty) {
    const anyNonEmpty = indexes.some((i) => {
      const txt = getDisplayValue(row, i).trim();
      return txt.length > 0;
    });
    if (!anyNonEmpty) return false;
  }

  for (const criterion of criteria) {
    const query = criterion.query;
    if (!query) continue;
    let matched = false;
    for (const i of criterion.indexes) {
      if (i >= values.length) continue;
      const text = getDisplayValue(row, i).toLowerCase();
      const altDate = parseDateFlexible(values[i]);
      const candidates = [text];
      if (altDate instanceof Date) {
        const dd = String(altDate.getDate()).padStart(2, "0");
        const mm = String(altDate.getMonth() + 1).padStart(2, "0");
        const yyyy = String(altDate.getFullYear());
        const yy = yyyy.slice(-2);
        candidates.push(`${dd}-${mm}-${yy}`);
        candidates.push(`${dd}-${mm}-${yyyy}`);
      }
      if (criterion.mode === "Równa się" && candidates.some((c) => c === query)) matched = true;
      if (criterion.mode === "Zaczyna się" && candidates.some((c) => c.startsWith(query))) matched = true;
      if (criterion.mode === "Zawiera" && candidates.some((c) => c.includes(query))) matched = true;
      if (matched) break;
    }
    if (!matched) return false;
  }

  return true;
}

function rowMatchesDateFilter(row, indexes, dateRange) {
  if (!dateRange.from && !dateRange.to) return true;
  for (const idx of indexes) {
    if (idx >= row.values.length) continue;
    const raw = row.rawValues ? row.rawValues[idx] : row.values[idx];
    const d = parseDateFlexible(raw ?? getDisplayValue(row, idx));
    if (!d) continue;
    if (dateRange.from && d < dateRange.from) continue;
    if (dateRange.to && d > dateRange.to) continue;
    return true;
  }
  return false;
}

function resolveIndexes(headers, selected) {
  if (!selected.size) return headers.map((_, i) => i);
  return headers.map((h, i) => (selected.has(h) ? i : -1)).filter((i) => i >= 0);
}

function applyFilters() {
  const criteria = [
    {
      query: (searchQueryEl.value || "").trim().toLowerCase(),
      mode: filterModeEl.value,
      indexes: resolveIndexes(currentHeaders, columnSelections.filter1),
    },
    {
      query: (searchQuery2El.value || "").trim().toLowerCase(),
      mode: filterMode2El.value,
      indexes: resolveIndexes(currentHeaders, columnSelections.filter2),
    },
  ];

  const dateIndexes = resolveIndexes(currentHeaders, columnSelections.date);
  const dateRange = getDateRange();
  const onlyNonEmpty = onlyNonEmptyEl.checked;

  viewRows = baseRows.filter((row) => {
    if (!rowMatchesTextFilter(row, criteria, onlyNonEmpty)) return false;
    if (!rowMatchesDateFilter(row, dateIndexes, dateRange)) return false;
    return true;
  });
}

function sortRows() {
  const sortCol = sortState.col;
  const ascending = sortState.dir === "asc";
  if (!sortCol) return;
  const idx = currentHeaders.indexOf(sortCol);
  if (idx < 0) return;
  viewRows.sort((a, b) => {
    const av = a.rawValues ? a.rawValues[idx] : a.values[idx];
    const bv = b.rawValues ? b.rawValues[idx] : b.values[idx];
    const ad = parseDateFlexible(av);
    const bd = parseDateFlexible(bv);
    if (ad && bd) return ad - bd;
    if (typeof av === "number" && typeof bv === "number") return av - bv;
    return String(av || "").localeCompare(String(bv || ""));
  });
  if (!ascending) viewRows.reverse();
}

function updateSortControls() {
  if (!resetSortBtn) return;
  const active = !!sortState.col;
  resetSortBtn.classList.toggle("hidden", !active);
  resetSortBtn.setAttribute("aria-hidden", active ? "false" : "true");
}

function toPixelWidth(meta) {
  if (!meta || typeof meta !== "object") return null;
  if (Number.isFinite(meta.wpx)) return Math.max(40, Math.round(meta.wpx));
  if (Number.isFinite(meta.wch)) return Math.max(40, Math.round(meta.wch * 8 + 16));
  if (Number.isFinite(meta.width)) return Math.max(40, Math.round(meta.width * 7 + 8));
  return null;
}

function toPixelHeight(meta) {
  if (!meta || typeof meta !== "object") return null;
  if (Number.isFinite(meta.hpx)) return Math.max(18, Math.round(meta.hpx));
  if (Number.isFinite(meta.hpt)) return Math.max(18, Math.round((meta.hpt * 96) / 72));
  return null;
}

function normalizeHexColor(input) {
  if (!input) return null;
  const raw = String(input).replace(/^#/, "").trim();
  if (/^[A-Fa-f0-9]{8}$/.test(raw)) return `#${raw.slice(2)}`;
  if (/^[A-Fa-f0-9]{6}$/.test(raw)) return `#${raw}`;
  if (/^[A-Fa-f0-9]{3}$/.test(raw)) return `#${raw[0]}${raw[0]}${raw[1]}${raw[1]}${raw[2]}${raw[2]}`;
  return null;
}

function colorFromStyleNode(node) {
  if (!node || typeof node !== "object") return null;
  const rgb = node.rgb ?? node.RGB;
  const direct = normalizeHexColor(rgb);
  if (direct) return direct;
  const auto = normalizeHexColor(node.auto);
  if (auto) return auto;
  return null;
}

function isDefaultLikeFill(fill, fillColor) {
  if (!fill || typeof fill !== "object") return true;
  const patternType = String(fill.patternType || fill.PatternType || "none").toLowerCase();
  if (!patternType || patternType === "none") return true;
  if (!fillColor) return true;
  const fg = fill.fgColor || fill.FgColor || null;
  const hasExplicitFgColor = !!(
    fg
    && typeof fg === "object"
    && (
      fg.rgb != null
      || fg.RGB != null
      || fg.indexed != null
      || fg.Indexed != null
      || fg.theme != null
      || fg.Theme != null
      || fg.tint != null
      || fg.Tint != null
    )
  );
  const normalized = String(fillColor).toUpperCase();
  // White-ish fill can be intentionally chosen by the user (especially solid fill),
  // so treat it as custom when the fg color is explicitly present in style.
  if (normalized === "#FFFFFF" || normalized === "#FFFFFE") {
    return !(patternType === "solid" && hasExplicitFgColor);
  }
  if (normalized === "#000000") return true;
  return false;
}

function isDefaultLikeFontColor(fontColor) {
  if (!fontColor) return true;
  const normalized = String(fontColor).toUpperCase();
  return normalized === "#000000" || normalized === "#FFFFFF";
}

function isCustomAlignment(alignment) {
  if (!alignment || typeof alignment !== "object") return false;
  const horizontal = String(alignment.horizontal || alignment.Horizontal || "").toLowerCase();
  const vertical = String(alignment.vertical || alignment.Vertical || "").toLowerCase();
  const wrapText = !!(alignment.wrapText || alignment.wrap || alignment.WrapText);
  const isDefaultHorizontal = !horizontal || horizontal === "general";
  const isDefaultVertical = !vertical || vertical === "bottom";
  return !isDefaultHorizontal || !isDefaultVertical || wrapText;
}

function hasCustomBorder(border) {
  if (!border || typeof border !== "object") return false;
  const edges = [
    border.top || border.Top,
    border.right || border.Right,
    border.bottom || border.Bottom,
    border.left || border.Left,
    border.diagonal || border.Diagonal,
  ];
  return edges.some((edge) => {
    const style = String(edge?.style || edge?.Style || "").toLowerCase();
    return !!style && style !== "none";
  });
}

function resolveXfStyle(styleIndex, wb) {
  if (!Number.isFinite(styleIndex) || !wb || !wb.Styles) return null;
  const styles = wb.Styles;
  const xfs = styles.CellXfs || styles.cellXfs;
  const xf = Array.isArray(xfs) ? xfs[styleIndex] : null;
  if (!xf) return null;
  const fontId = xf.fontId ?? xf.FontId;
  const fillId = xf.fillId ?? xf.FillId;
  const borderId = xf.borderId ?? xf.BorderId;
  const alignment = xf.alignment || xf.Alignment || null;
  const numFmtId = xf.numFmtId ?? xf.NumFmtId;
  const fonts = styles.Fonts || styles.fonts || [];
  const fills = styles.Fills || styles.fills || [];
  const borders = styles.Borders || styles.borders || [];
  return {
    font: Number.isFinite(fontId) ? fonts[fontId] : null,
    fill: Number.isFinite(fillId) ? fills[fillId] : null,
    border: Number.isFinite(borderId) ? borders[borderId] : null,
    alignment,
    numFmtId,
  };
}

function extractCellStyle(cell, wb) {
  if (!cell) return null;
  let style = null;
  if (cell.s && typeof cell.s === "object") {
    style = cell.s;
  } else if (Number.isFinite(cell.s)) {
    style = resolveXfStyle(cell.s, wb);
  }
  if (!style || typeof style !== "object") return null;

  const fill = style.fill || style.Fill || null;
  const font = style.font || style.Font || null;
  const border = style.border || style.Border || null;
  const alignment = style.alignment || style.Alignment || null;

  const fillColor = colorFromStyleNode(fill?.fgColor || fill?.FgColor || fill?.bgColor || fill?.BgColor);
  const fontColor = colorFromStyleNode(font?.color || font?.Color);
  const hasCustomFill = !isDefaultLikeFill(fill, fillColor);
  const hasCustomFontColor = !isDefaultLikeFontColor(fontColor);
  const hasCustomAlign = isCustomAlignment(alignment);
  const hasBorder = hasCustomBorder(border);

  const styleOut = {
    fillColor,
    hasCustomFill,
    fontColor,
    hasCustomFontColor,
    bold: !!(font && (font.bold || font.b || font.Bold)),
    italic: !!(font && (font.italic || font.i || font.Italic)),
    underline: !!(font && (font.underline || font.u || font.Underline)),
    horizontal: hasCustomAlign ? (alignment?.horizontal || alignment?.Horizontal || "") : "",
    vertical: hasCustomAlign ? (alignment?.vertical || alignment?.Vertical || "") : "",
    wrapText: hasCustomAlign ? !!(alignment && (alignment.wrapText || alignment.wrap || alignment.WrapText)) : false,
    hasBorder,
    border,
  };

  return styleOut;
}

function applyEdgeBorder(td, edge) {
  if (!edge) return;
  const borderStyle = edge.style || edge.Style || "";
  if (!borderStyle || borderStyle === "none") return;
  const color = colorFromStyleNode(edge.color || edge.Color) || "rgba(0,0,0,0.32)";
  return `1px solid ${color}`;
}

function applyCellStyle(td, style) {
  if (!style) return;
  if (style.hasCustomFill && style.fillColor) {
    td.classList.add("cell-has-fill");
    td.style.background = hexToRgba(style.fillColor, 0.28) || td.style.background;
  }
  if (style.hasCustomFontColor && style.fontColor) td.style.color = style.fontColor;
  if (style.bold) td.style.fontWeight = "700";
  if (style.italic) td.style.fontStyle = "italic";
  if (style.underline) td.style.textDecoration = "underline";
  if (style.horizontal) td.style.textAlign = style.horizontal;
  if (style.vertical) td.style.verticalAlign = style.vertical;
  if (style.wrapText) td.style.whiteSpace = "normal";

  if (style.hasBorder && style.border && typeof style.border === "object") {
    const t = applyEdgeBorder(td, style.border.top || style.border.Top);
    const r = applyEdgeBorder(td, style.border.right || style.border.Right);
    const b = applyEdgeBorder(td, style.border.bottom || style.border.Bottom);
    const l = applyEdgeBorder(td, style.border.left || style.border.Left);
    if (t) td.style.borderTop = t;
    if (r) td.style.borderRight = r;
    if (b) td.style.borderBottom = b;
    if (l) td.style.borderLeft = l;
  }
}

function computeMergeLayout(rowsShown, colCount) {
  if (!Array.isArray(currentMerges) || !currentMerges.length || !rowsShown.length) return null;
  const rowPosByAbs = new Map();
  rowsShown.forEach((row, pos) => rowPosByAbs.set(row.rowIndex0, pos));
  const anchors = new Map();
  const covered = new Set();

  currentMerges.forEach((merge) => {
    if (!merge || !merge.s || !merge.e) return;
    if (merge.s.c < currentStartCol || merge.e.c >= currentStartCol + colCount) return;
    const topPos = rowPosByAbs.get(merge.s.r);
    if (topPos == null) return;
    for (let r = merge.s.r; r <= merge.e.r; r++) {
      const p = rowPosByAbs.get(r);
      if (p == null || p !== topPos + (r - merge.s.r)) return;
    }
    const startCol = merge.s.c - currentStartCol;
    const endCol = merge.e.c - currentStartCol;
    const rowspan = merge.e.r - merge.s.r + 1;
    const colspan = endCol - startCol + 1;
    if (rowspan < 2 && colspan < 2) return;

    const anchorKey = `${topPos}:${startCol}`;
    anchors.set(anchorKey, {
      rowspan,
      colspan,
      ref: XLSX.utils.encode_range({
        s: { r: merge.s.r, c: merge.s.c },
        e: { r: merge.e.r, c: merge.e.c },
      }),
    });
    for (let rp = topPos; rp < topPos + rowspan; rp++) {
      for (let c = startCol; c <= endCol; c++) {
        if (rp === topPos && c === startCol) continue;
        covered.add(`${rp}:${c}`);
      }
    }
  });

  return { anchors, covered };
}

function computeHeaderMergeLayout(colCount) {
  if (!Array.isArray(currentMerges) || !currentMerges.length) return null;
  const headerAbsRow = currentHeaderRow - 1;
  const anchors = new Map();
  const covered = new Set();

  currentMerges.forEach((merge) => {
    if (!merge || !merge.s || !merge.e) return;
    if (merge.s.r !== headerAbsRow || merge.e.r !== headerAbsRow) return;
    if (merge.s.c < currentStartCol || merge.e.c >= currentStartCol + colCount) return;
    const startCol = merge.s.c - currentStartCol;
    const endCol = merge.e.c - currentStartCol;
    const colspan = endCol - startCol + 1;
    if (colspan < 2) return;
    anchors.set(startCol, {
      colspan,
      ref: XLSX.utils.encode_range({
        s: { r: merge.s.r, c: merge.s.c },
        e: { r: merge.e.r, c: merge.e.c },
      }),
    });
    for (let c = startCol + 1; c <= endCol; c++) covered.add(c);
  });

  return { anchors, covered };
}

function computeColumnWidths(headers, rows, useExcelLayout) {
  const widths = headers.map(() => 0);
  const min = 80;
  const max = 520;

  if (useExcelLayout && Array.isArray(currentSheetColWidths) && currentSheetColWidths.length) {
    return widths.map((_, i) => {
      const manual = manualColumnWidths[i];
      if (manual) return Math.max(min, Math.min(max, manual));
      const fromSheet = toPixelWidth(currentSheetColWidths[i]);
      if (fromSheet) return Math.max(min, Math.min(max, fromSheet));
      return 140;
    });
  }

  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");
  const tableFont = getComputedStyle(tableEl).font;
  ctx.font = tableFont;

  headers.forEach((h, i) => {
    widths[i] = Math.max(widths[i], ctx.measureText(h).width);
  });
  const limit = Math.min(rows.length, 300);
  const samples = headers.map(() => []);
  for (let r = 0; r < limit; r++) {
    rows[r].values.forEach((v, i) => {
      const text = getDisplayValue(rows[r], i);
      const w = ctx.measureText(text).width;
      samples[i].push(w);
    });
  }
  const padding = 24;
  return widths.map((base, i) => {
    const colSamples = samples[i].sort((a, b) => a - b);
    const idx = Math.floor(colSamples.length * 0.9);
    const p90 = colSamples.length ? colSamples[Math.min(idx, colSamples.length - 1)] : base;
    const raw = Math.max(base, p90) + padding;
    const manual = manualColumnWidths[i];
    if (manual) return Math.max(min, Math.min(max, manual));
    return Math.max(min, Math.min(max, Math.ceil(raw)));
  });
}

function renderTable(headers, rows) {
  updateSortControls();
  if (!headers.length) {
    setStatus("Brak danych");
    if (tableScrollbarEl) tableScrollbarEl.classList.add("hidden");
    setEmptyState(DEFAULT_EMPTY_TITLE, DEFAULT_EMPTY_SUB);
    return;
  }
  if (!rows.length) {
    setStatus("Wierszy: 0");
    if (tableScrollbarEl) tableScrollbarEl.classList.add("hidden");
    setEmptyState("Brak wynikow", "Zmien filtry albo wybierz inny arkusz.");
    return;
  }

  showTable();
  theadEl.innerHTML = "";
  tbodyEl.innerHTML = "";

  const useExcelLayout = isExcelLayoutEnabled();
  const widths = computeColumnWidths(headers, rows, useExcelLayout);
  const rowHeaderDigits = String(rows.length + currentHeaderRow).length;
  const rowHeaderWidth = Math.max(42, rowHeaderDigits * 8 + 18);

  const colgroup = document.createElement("colgroup");
  const rowHeadCol = document.createElement("col");
  rowHeadCol.style.width = `${rowHeaderWidth}px`;
  colgroup.appendChild(rowHeadCol);
  widths.forEach((w) => {
    const col = document.createElement("col");
    col.style.width = `${w}px`;
    colgroup.appendChild(col);
  });
  tableEl.innerHTML = "";
  tableEl.appendChild(colgroup);
  tableEl.appendChild(theadEl);
  tableEl.appendChild(tbodyEl);

  const guideRow = document.createElement("tr");
  guideRow.className = "guide-row";
  const corner = document.createElement("th");
  corner.className = "corner-cell";
  corner.textContent = "";
  guideRow.appendChild(corner);
  headers.forEach((_, i) => {
    const th = document.createElement("th");
    th.className = "guide-cell";
    th.setAttribute("scope", "col");
    th.textContent = XLSX.utils.encode_col(i + currentStartCol);
    const resizer = document.createElement("div");
    resizer.className = "col-resizer";
    resizer.dataset.colIndex = String(i);
    th.appendChild(resizer);
    guideRow.appendChild(th);
  });
  theadEl.appendChild(guideRow);

  const headRow = document.createElement("tr");
  headRow.className = "header-row";
  const rowHead = document.createElement("th");
  rowHead.className = "row-head";
  rowHead.setAttribute("scope", "row");
  rowHead.textContent = String(currentHeaderRow);
  headRow.appendChild(rowHead);
  const headerMergeLayout = computeHeaderMergeLayout(headers.length);
  for (let i = 0; i < headers.length; i++) {
    if (headerMergeLayout && headerMergeLayout.covered.has(i)) continue;
    const h = headers[i];
    const th = document.createElement("th");
    th.setAttribute("scope", "col");
    th.textContent = h;
    if (currentHeaderStyles[i]) applyCellStyle(th, currentHeaderStyles[i]);

    if (headerMergeLayout) {
      const merge = headerMergeLayout.anchors.get(i);
      if (merge) {
        th.colSpan = merge.colspan;
        th.classList.add("cell-merged");
        if (merge.ref) th.title = `Scalona komórka: ${merge.ref}`;
      }
    }

    th.addEventListener("click", () => {
      if (sortState.col === h) {
        sortState.dir = sortState.dir === "asc" ? "desc" : "asc";
      } else {
        sortState.col = h;
        sortState.dir = "asc";
      }
      sortRows();
      updateSortControls();
      renderTable(currentHeaders, viewRows);
    });

    if (sortState.col === h) {
      const arrow = document.createElement("span");
      arrow.className = "sort-arrow";
      arrow.textContent = sortState.dir === "asc" ? "▲" : "▼";
      th.appendChild(arrow);
    }

    headRow.appendChild(th);
  }
  theadEl.appendChild(headRow);

  const limit = Math.max(1, parseInt(maxRowsEl.value || "200", 10));
  const rowsShown = rows.slice(0, limit);
  const mergeLayout = computeMergeLayout(rowsShown, headers.length);

  rowsShown.forEach((row, rowPos) => {
    const tr = document.createElement("tr");
    if (row.isSubheader) tr.classList.add("row-subheader");
    if (typeof row.rowIndex0 === "number") {
      tr.dataset.rowIndex = String(row.rowIndex0);
    }
    if (useExcelLayout) {
      const rowMeta = currentSheetRowHeights[row.rowIndex0];
      const h = toPixelHeight(rowMeta);
      if (h) tr.style.height = `${h}px`;
    }
    const rowHead = document.createElement("td");
    rowHead.className = "row-head";
    rowHead.textContent = String(row.rowIndex0 + 1);
    tr.appendChild(rowHead);
    row.values.forEach((v, i) => {
      const mergeKey = `${rowPos}:${i}`;
      if (mergeLayout && mergeLayout.covered.has(mergeKey)) return;
      const td = document.createElement("td");
      const displayValue = getDisplayValue(row, i);
      td.textContent = displayValue;
      td.dataset.fullText = displayValue;
      td.dataset.colIndex = String(i);

      if (mergeLayout) {
        const anchor = mergeLayout.anchors.get(mergeKey);
        if (anchor) {
          if (anchor.rowspan > 1) td.rowSpan = anchor.rowspan;
          if (anchor.colspan > 1) td.colSpan = anchor.colspan;
          td.classList.add("cell-merged");
          if (anchor.ref) td.title = `Scalona komórka: ${anchor.ref}`;
        }
      }

      if (row.cellStyles && row.cellStyles[i]) applyCellStyle(td, row.cellStyles[i]);
      tr.appendChild(td);
    });
    tbodyEl.appendChild(tr);
  });

  setStatus(`Wierszy: ${rows.length} (pokazano: ${Math.min(rows.length, limit)})`);
  syncHorizontalScrollbar();
}

function buildRows(sheet, headerRow, wb) {
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const colMeta = sheet["!cols"] || [];
  const rowMeta = sheet["!rows"] || [];
  const merges = Array.isArray(sheet["!merges"]) ? sheet["!merges"] : [];
  const headers = [];
  const headerStyles = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: headerRow - 1, c })];
    const v = cell ? cell.v : null;
    headers.push(v ? String(v).trim() : XLSX.utils.encode_col(c));
    headerStyles.push(wb ? extractCellStyle(cell, wb) : null);
  }
  const rows = [];
  for (let r = headerRow; r <= range.e.r; r++) {
    const values = [];
    const display = [];
    const cellStyles = [];
    let any = false;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      let v = cell ? cell.v : null;
      let shown = cell && cell.w ? String(cell.w) : toDisplay(v);
      if (displayModeEl.value === "Formuły" && cell && cell.f) {
        v = "=" + cell.f;
        shown = v;
      }
      values.push(v);
      display.push(shown);
      cellStyles.push(wb ? extractCellStyle(cell, wb) : null);
      if (v !== null && v !== "") any = true;
    }
    if (!any) continue;
    rows.push({ values, display, rawValues: values, rowIndex0: r, cellStyles });
  }
  return {
    headers,
    headerStyles,
    rows,
    startCol: range.s.c,
    merges,
    colWidths: headers.map((_, idx) => colMeta[range.s.c + idx] || null),
    rowHeights: rowMeta,
  };
}

function hexToRgba(hex, alpha = 0.35) {
  if (!hex || typeof hex !== "string") return null;
  const m = hex.replace(/^#/, "").replace(/^([A-Fa-f0-9]{6})[A-Fa-f0-9]*$/, "$1").match(/([A-Fa-f0-9]{2})([A-Fa-f0-9]{2})([A-Fa-f0-9]{2})/);
  if (!m) return null;
  return `rgba(${parseInt(m[1], 16)}, ${parseInt(m[2], 16)}, ${parseInt(m[3], 16)}, ${alpha})`;
}

// [EN] Mark rows that look like subheaders (e.g. second header row or important info) in first N data rows
function markSubheaderRows(rows, maxCheck = 10) {
  const toCheck = Math.min(maxCheck, rows.length);
  for (let i = 0; i < toCheck; i++) {
    const row = rows[i];
    let nonEmpty = 0;
    let textLike = 0;
    row.values.forEach((v) => {
      if (v != null && String(v).trim() !== "") {
        nonEmpty += 1;
        if (typeof v === "string") textLike += 1;
        else if (!(v instanceof Date) && typeof v !== "number") textLike += 1;
      }
    });
    const n = row.values.length;
    if (n === 0) continue;
    if (nonEmpty >= n * 0.5 && textLike >= n * 0.5) row.isSubheader = true;
  }
  return rows;
}

function detectHeaderRowSimple(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const maxRow = Math.min(range.e.r, range.s.r + 100);
  for (let r = range.s.r; r <= maxRow; r++) {
    let filled = 0;
    let stringCount = 0;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      if (!cell) continue;
      const v = cell.v;
      if (v === null || v === "") continue;
      filled += 1;
      if (typeof v === "string") stringCount += 1;
      if (filled >= 2 || (filled >= 1 && stringCount >= 1)) {
        return r + 1;
      }
    }
  }
  return 1;
}

function applyAutoHeaderRowIfEnabled() {
  if (!autoHeaderRowEl || !autoHeaderRowEl.checked) return false;
  if (!workbook) return false;
  const sheetName = sheetSelect.value;
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return false;
  const detected = detectHeaderRowSimple(sheet);
  headerRowEl.value = String(detected);
  return true;
}

function columnSummary(set) {
  if (!set.size) return "Wszystkie kolumny";
  if (set.size === 1) return Array.from(set)[0];
  return `${set.size} kolumn`;
}

function updateColumnSummary() {
  filter1ColumnsEl.value = columnSummary(columnSelections.filter1);
  filter2ColumnsEl.value = columnSummary(columnSelections.filter2);
  dateColumnsEl.value = columnSummary(columnSelections.date);
  updateQuickSearchColumnButtons();
}

function updateFilterBadge() {
  let count = 0;
  if (searchQueryEl.value.trim()) count += 1;
  if (searchQuery2El.value.trim()) count += 1;
  if (onlyNonEmptyEl.checked) count += 1;
  if (dateModeEl.value === "last_n_days") count += 1;
  if (dateFromEl.value.trim() || dateToEl.value.trim()) count += 1;
  if (columnSelections.filter1.size) count += 1;
  if (columnSelections.filter2.size) count += 1;
  if (columnSelections.date.size) count += 1;

  filterBadgeEl.textContent = String(count);
  filterBadgeEl.classList.toggle("hidden", count === 0);
}

function updateDateChipsActive() {
  const isLastN = dateModeEl.value === "last_n_days";
  const days = lastDaysEl.value.trim() ? String(lastDaysEl.value) : "30";
  quickRangeButtons.forEach((btn) => {
    const active = isLastN && btn.dataset.range === days;
    btn.classList.toggle("active", !!active);
  });
}

function isSidebarOpen() {
  return rootEl.classList.contains("sidebar-open");
}

function syncQuickSearchInputs() {
  if (quickSearchEl) quickSearchEl.value = searchQueryEl.value;
  if (quickSearchPopupInput) quickSearchPopupInput.value = searchQueryEl.value;
}

function updateQuickSearchColumnButtons() {
  const summary = columnSummary(columnSelections.filter1);
  const count = columnSelections.filter1.size;
  const label = count ? `Kolumny (${count})` : "Kolumny";
  [quickSearchColumnsBtn, quickSearchPopupColumnsBtn].forEach((btn) => {
    if (!btn) return;
    btn.textContent = label;
    btn.title = `Szybkie szukanie: ${summary}`;
    btn.setAttribute("aria-label", `Kolumny szybkiego szukania. ${summary}.`);
  });
}

function resetFilterInputs() {
  searchQueryEl.value = "";
  searchQuery2El.value = "";
  filterModeEl.value = "Zawiera";
  filterMode2El.value = "Zawiera";
  onlyNonEmptyEl.checked = false;
  dateModeEl.value = "between";
  dateFromEl.value = "";
  dateToEl.value = "";
  lastDaysEl.value = "";
  columnSelections.filter1.clear();
  columnSelections.filter2.clear();
  columnSelections.date.clear();
  syncQuickSearchInputs();
  updateColumnSummary();
  updateDateChipsActive();
  updateFilterBadge();
}

function setSidebarOpen(open) {
  const shouldOpen = !!open;
  rootEl.classList.toggle("sidebar-open", shouldOpen);
  if (sidebarScrim) sidebarScrim.classList.toggle("hidden", !shouldOpen);
  if (panelToggle) {
    panelToggle.setAttribute("aria-expanded", shouldOpen ? "true" : "false");
    panelToggle.textContent = shouldOpen ? "Zamknij filtry" : "Filtry";
  }
}

function openColumnPicker(key) {
  if (!currentHeaders.length) {
    toast("Wczytaj arkusz, żeby wybrac kolumny", "info");
    return;
  }
  activePickerKey = key;
  if (columnPickerTitleEl) {
    columnPickerTitleEl.textContent = key === "filter1"
      ? "Kolumny szybkiego szukania"
      : key === "filter2"
        ? "Kolumny filtru tekstowego 2"
        : "Kolumny filtru dat";
  }
  columnListEl.innerHTML = "";
  columnSearchEl.value = "";
  const currentSet = columnSelections[key];
  const isAll = currentSet.size === 0;
  currentHeaders.forEach((h, idx) => {
    const row = document.createElement("div");
    row.className = "field checkbox";
    const input = document.createElement("input");
    input.type = "checkbox";
    input.id = `colpick-${key}-${idx}`;
    input.value = h;
    input.checked = isAll ? true : currentSet.has(h);
    const label = document.createElement("label");
    label.htmlFor = input.id;
    label.textContent = h;
    row.appendChild(input);
    row.appendChild(label);
    columnListEl.appendChild(row);
  });
  columnPickerEl.classList.remove("hidden");
  columnSearchEl.focus();
}

function closeColumnPicker() {
  columnPickerEl.classList.add("hidden");
  if (lastPickerTriggerEl) {
    lastPickerTriggerEl.focus();
    lastPickerTriggerEl = null;
  }
}

function getModalFocusables() {
  const modalContent = columnPickerEl.querySelector(".modal-content");
  if (!modalContent) return [];
  const all = Array.from(modalContent.querySelectorAll("button, input:not([type=hidden]), [tabindex]:not([tabindex^='-'])"));
  return all.filter((el) => {
    const row = el.closest(".field.checkbox");
    return !row || !row.classList.contains("hidden");
  });
}

function handlePickerKeydown(e) {
  if (columnPickerEl.classList.contains("hidden")) return;
  if (e.key === "Tab") {
    const focusables = getModalFocusables();
    if (focusables.length === 0) return;
    const idx = focusables.indexOf(document.activeElement);
    if (idx === -1) return;
    if (e.shiftKey && idx === 0) {
      e.preventDefault();
      focusables[focusables.length - 1].focus();
    } else if (!e.shiftKey && idx === focusables.length - 1) {
      e.preventDefault();
      focusables[0].focus();
    }
  }
}

function filterColumnList() {
  const q = columnSearchEl.value.trim().toLowerCase();
  columnListEl.querySelectorAll(".field.checkbox").forEach((row) => {
    const text = row.textContent.toLowerCase();
    row.classList.toggle("hidden", q && !text.includes(q));
  });
}

columnListEl.addEventListener("click", (e) => {
  const row = e.target.closest(".field.checkbox");
  if (!row) return;
  const cb = row.querySelector("input[type=checkbox]");
  if (!cb) return;
  if (e.target !== cb) cb.checked = !cb.checked;
});

columnListEl.addEventListener("mousedown", (e) => {
  const row = e.target.closest(".field.checkbox");
  if (!row) return;
  const cb = row.querySelector("input[type=checkbox]");
  if (!cb) return;
  if (e.target !== cb) {
    e.preventDefault();
    cb.checked = !cb.checked;
  }
});


function attachResizeHandlers() {
  let active = null;
  let startX = 0;
  let startW = 0;

  const start = (e) => {
    const handle = e.target.closest(".col-resizer");
    if (!handle) return;
    e.preventDefault();
    const colIndex = parseInt(handle.dataset.colIndex, 10);
    const th = handle.parentElement;
    active = { colIndex, th };
    startX = e.clientX || (e.touches && e.touches[0].clientX) || 0;
    startW = th.getBoundingClientRect().width;
    document.body.classList.add("resizing");
  };

  const move = (e) => {
    if (!active) return;
    const x = e.clientX || (e.touches && e.touches[0].clientX) || 0;
    const delta = x - startX;
    const next = Math.max(80, Math.min(520, Math.round(startW + delta)));
    manualColumnWidths[active.colIndex] = next;
    renderTable(currentHeaders, viewRows);
  };

  const stop = () => {
    if (!active) return;
    active = null;
    document.body.classList.remove("resizing");
  };

  tableEl.addEventListener("mousedown", start);
  tableEl.addEventListener("touchstart", start, { passive: true });
  window.addEventListener("mousemove", move);
  window.addEventListener("touchmove", move, { passive: true });
  window.addEventListener("mouseup", stop);
  window.addEventListener("touchend", stop);
}

function initTheme() {
  const saved = localStorage.getItem(THEME_KEY);
  const prefersDark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
  const theme = saved || (prefersDark ? "dark" : "light");
  setTheme(theme, false);
}

function initIntroSplash() {
  const splash = document.getElementById("heroSplash");
  const vid = document.getElementById("introVideo");
  if (!splash) return;

  if (sessionStorage.getItem(INTRO_PLAYED_KEY)) {
    splash.style.display = "none";
    document.body.classList.remove("splashing");
    return;
  }

  document.body.classList.add("splashing");

  const hideSplash = () => {
    if (!splash || splash.classList.contains("hide")) return;
    splash.classList.add("hide");
    sessionStorage.setItem(INTRO_PLAYED_KEY, "true");
    setTimeout(() => {
      splash.style.display = "none";
      document.body.classList.remove("splashing");
    }, 700);
  };

  if (vid) {
    try {
      vid.currentTime = 0;
      vid.muted = true;
      const playPromise = vid.play();
      if (playPromise && typeof playPromise.catch === "function") {
        playPromise.catch(() => hideSplash());
      }
    } catch {
      hideSplash();
    }

    const fallback = setTimeout(hideSplash, 15000);
    vid.addEventListener("ended", () => {
      clearTimeout(fallback);
      hideSplash();
    });
  } else {
    setTimeout(hideSplash, 8000);
  }
}

function setTheme(theme, persist = true) {
  rootEl.setAttribute("data-theme", theme);
  themeToggle.setAttribute("aria-pressed", theme === "dark" ? "true" : "false");
  if (persist) localStorage.setItem(THEME_KEY, theme);
  themeToggle.innerHTML =
    theme === "dark"
      ? "<svg width=\"18\" height=\"18\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><circle cx=\"12\" cy=\"12\" r=\"5\"/><line x1=\"12\" y1=\"1\" x2=\"12\" y2=\"3\"/><line x1=\"12\" y1=\"21\" x2=\"12\" y2=\"23\"/><line x1=\"4.22\" y1=\"4.22\" x2=\"5.64\" y2=\"5.64\"/><line x1=\"18.36\" y1=\"18.36\" x2=\"19.78\" y2=\"19.78\"/><line x1=\"1\" y1=\"12\" x2=\"3\" y2=\"12\"/><line x1=\"21\" y1=\"12\" x2=\"23\" y2=\"12\"/><line x1=\"4.22\" y1=\"19.78\" x2=\"5.64\" y2=\"18.36\"/><line x1=\"18.36\" y1=\"5.64\" x2=\"19.78\" y2=\"4.22\"/></svg>"
      : "<svg width=\"18\" height=\"18\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><path d=\"M21 12.79A9 9 0 1 1 11.21 3a7 7 0 0 0 9.79 9.79z\"/></svg>";
}

function updateNetworkBadge() {
  if (!networkBadgeEl) return;
  const isOnline = navigator.onLine;
  networkBadgeEl.textContent = isOnline ? "Online" : "Offline";
  networkBadgeEl.classList.toggle("offline", !isOnline);
  const safetyNote = "Pliki Excel są wczytywane i przetwarzane lokalnie na Twoim urządzeniu.";
  networkBadgeEl.setAttribute(
    "title",
    isOnline
      ? `Połączenie aktywne. ${safetyNote}`
      : `Brak połączenia sieciowego. ${safetyNote}`
  );
}

async function handleFile(file) {
  if (!file) return;
  if (!isXlsxAvailable(true)) return;
  try {
    setLoading(true, "Wczytywanie pliku...");
    const data = await file.arrayBuffer();
    try {
      workbook = XLSX.read(data, { cellDates: true, cellStyles: true });
    } catch {
      workbook = XLSX.read(data, { cellDates: true });
    }
    sheetSelect.innerHTML = "";
    workbook.SheetNames.forEach((s) => {
      const opt = document.createElement("option");
      opt.value = s;
      opt.textContent = s;
      sheetSelect.appendChild(opt);
    });
    currentFileName = file.name;
    currentStartCol = 0;
    currentMerges = [];
    currentHeaderStyles = [];
    currentSheetColWidths = [];
    currentSheetRowHeights = {};
    fileNameTextEl.textContent = file.name;
    fileNameEl.classList.remove("hidden");
    dropZone.classList.add("has-file");
    setDirtyState(false);
    setStatus("Plik wczytany");
    toast("Plik wczytany", "success");
    log(`Wczytano plik: ${file.name}`, "success");
  } catch (err) {
    toast("Nie udalo sie wczytac pliku", "error");
    log("Blad przy wczytywaniu pliku.", "error");
  } finally {
    setLoading(false);
  }
}

function escapeCsv(value) {
  const raw = String(value ?? "");
  if (raw.includes("\"") || raw.includes(",") || raw.includes("\n")) {
    return `"${raw.replace(/\"/g, '""')}"`;
  }
  return raw;
}

function exportCsv() {
  if (!currentHeaders.length || !viewRows.length) {
    toast("Brak danych do eksportu", "warning");
    return;
  }
  const rows = [
    currentHeaders,
    ...viewRows.map((row) => row.values.map((v, i) => getDisplayValue(row, i))),
  ];
  const csv = rows.map((row) => row.map(escapeCsv).join(",")).join("\n");
  const base = currentFileName ? currentFileName.replace(/\.[^.]+$/, "") : "excel-workbench";
  const sheet = sheetSelect.value ? sheetSelect.value.replace(/\s+/g, "_") : "arkusz";
  const filename = `${base}_${sheet}.csv`;
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  toast("Wyeksportowano CSV", "success");
}

function saveWorkbook() {
  if (!isXlsxAvailable(true)) return;
  if (!workbook) {
    toast("Brak pliku do zapisu", "warning");
    return;
  }
  const base = currentFileName ? currentFileName.replace(/\.[^.]+$/, "") : "excel-workbench";
  const ext = currentFileName && currentFileName.toLowerCase().endsWith(".xlsm") ? "xlsm" : "xlsx";
  if (ext === "xlsm") {
    const ok = window.confirm("Plik .xlsm moze utracic makra. Kontynuowac zapis?");
    if (!ok) return;
  }
  const filename = `${base}_edited.${ext}`;
  XLSX.writeFile(workbook, filename, { bookType: ext });
  setDirtyState(false);
  toast("Zapisano plik", "success");
  log(`Zapisano plik: ${filename}`, "success");
}

function saveWorkbookAs() {
  if (!isXlsxAvailable(true)) return;
  if (!workbook) {
    toast("Brak pliku do zapisu", "warning");
    return;
  }
  const base = currentFileName ? currentFileName.replace(/\.[^.]+$/, "") : "excel-workbench";
  const suggested = `${base}_edited.xlsx`;
  const nameRaw = window.prompt("Podaj nazwe pliku (xlsx lub xlsm):", suggested);
  if (!nameRaw) return;
  let name = nameRaw.trim();
  if (!name) return;
  if (!/\.(xlsx|xlsm)$/i.test(name)) {
    name = `${name}.xlsx`;
  }
  const ext = name.toLowerCase().endsWith(".xlsm") ? "xlsm" : "xlsx";
  if (ext === "xlsm") {
    const ok = window.confirm("Plik .xlsm moze utracic makra. Kontynuowac zapis?");
    if (!ok) return;
  }
  XLSX.writeFile(workbook, name, { bookType: ext });
  setDirtyState(false);
  toast("Zapisano plik", "success");
  log(`Zapisano plik: ${name}`, "success");
}

fileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  handleFile(file);
});

loadBtn.addEventListener("click", () => {
  if (!isXlsxAvailable(true)) return;
  if (!workbook) {
    toast("Najpierw wybierz plik", "warning");
    log("Najpierw wybierz plik.", "warn");
    return;
  }
  setLoading(true, "Budowanie tabeli...");
  setTimeout(() => {
    try {
      const sheetName = sheetSelect.value;
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) {
        toast("Brak arkusza", "error");
        log("Brak arkusza.", "error");
        return;
      }
      applyAutoHeaderRowIfEnabled();
      const headerRow = Math.max(1, parseInt(headerRowEl.value || "1", 10));
      currentHeaderRow = headerRow;
      currentSheetName = sheetName;
      const data = buildRows(sheet, headerRow, workbook);
      currentHeaders = data.headers;
      currentStartCol = data.startCol || 0;
      currentMerges = Array.isArray(data.merges) ? data.merges : [];
      currentHeaderStyles = Array.isArray(data.headerStyles) ? data.headerStyles : [];
      currentSheetColWidths = Array.isArray(data.colWidths) ? data.colWidths : [];
      currentSheetRowHeights = data.rowHeights || {};
      baseRows = markSubheaderRows(data.rows);
      viewRows = baseRows.slice();
      sortState = { col: "", dir: "asc" };
      manualColumnWidths = {};
      columnSelections.filter1.clear();
      columnSelections.filter2.clear();
      columnSelections.date.clear();
      updateColumnSummary();
      updateFilterBadge();
      renderTable(currentHeaders, viewRows);
      setDirtyState(false);
      toast("Arkusz wczytany", "success");
      log(`Wczytano arkusz: ${sheetName}`, "success");
    } finally {
      setLoading(false);
    }
  }, 50);
});

applyFilterBtn.addEventListener("click", () => {
  if (!currentHeaders.length) return;
  applyFilters();
  sortRows();
  renderTable(currentHeaders, viewRows);
  updateFilterBadge();
  toast("Zastosowano filtry", "info");
});

function applyQuickSearch() {
  if (!currentHeaders.length) return;
  let value = "";
  if (quickSearchPopupEl && !quickSearchPopupEl.classList.contains("hidden") && quickSearchPopupInput) value = quickSearchPopupInput.value;
  else if (quickSearchEl) value = quickSearchEl.value;
  else value = searchQueryEl.value || "";
  if (quickSearchPopupInput) quickSearchPopupInput.value = value;
  if (quickSearchEl) quickSearchEl.value = value;
  searchQueryEl.value = value;
  applyFilters();
  sortRows();
  renderTable(currentHeaders, viewRows);
  updateFilterBadge();
  if (quickSearchPopupEl && !quickSearchPopupEl.classList.contains("hidden")) {
    quickSearchPopupEl.classList.add("hidden");
  }
}

if (tableWrapEl && tableScrollbarEl) {
  tableWrapEl.addEventListener("scroll", () => {
    hideCellTooltip();
    if (syncingHorizontalScroll) return;
    syncingHorizontalScroll = true;
    tableScrollbarEl.scrollLeft = tableWrapEl.scrollLeft;
    requestAnimationFrame(() => {
      syncingHorizontalScroll = false;
    });
  }, { passive: true });

  tableScrollbarEl.addEventListener("scroll", () => {
    if (syncingHorizontalScroll) return;
    syncingHorizontalScroll = true;
    tableWrapEl.scrollLeft = tableScrollbarEl.scrollLeft;
    requestAnimationFrame(() => {
      syncingHorizontalScroll = false;
    });
  }, { passive: true });
}

tbodyEl.addEventListener("pointerenter", (e) => {
  const td = e.target.closest("td");
  if (!td || td.classList.contains("row-head")) return;
  showCellTooltip(td);
}, true);

tbodyEl.addEventListener("pointerleave", (e) => {
  const td = e.target.closest("td");
  if (!td) return;
  hideCellTooltip();
}, true);

tbodyEl.addEventListener("touchstart", (e) => {
  const td = e.target.closest("td");
  if (!td || td.classList.contains("row-head")) return;
  showCellTooltip(td, true);
}, { passive: true });

window.addEventListener("resize", () => {
  syncHorizontalScrollbar();
  hideCellTooltip();
});

if (quickSearchBtn) {
  quickSearchBtn.addEventListener("click", applyQuickSearch);
}

if (quickSearchColumnsBtn) {
  quickSearchColumnsBtn.addEventListener("click", () => {
    lastPickerTriggerEl = quickSearchColumnsBtn;
    openColumnPicker("filter1");
  });
}

if (quickSearchEl) {
  quickSearchEl.addEventListener("keydown", (e) => {
    if (e.key === "Enter") applyQuickSearch();
  });
}

if (quickSearchPopupInput) {
  quickSearchPopupInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); applyQuickSearch(); }
  });
}
if (quickSearchPopupBtn) {
  quickSearchPopupBtn.addEventListener("click", applyQuickSearch);
}
if (quickSearchPopupColumnsBtn) {
  quickSearchPopupColumnsBtn.addEventListener("click", () => {
    lastPickerTriggerEl = quickSearchPopupColumnsBtn;
    openColumnPicker("filter1");
  });
}
if (quickSearchPopupEl) {
  quickSearchPopupEl.addEventListener("click", (e) => {
    if (e.target === quickSearchPopupEl) quickSearchPopupEl.classList.add("hidden");
  });
}


resetFiltersBtn.addEventListener("click", () => {
  resetFilterInputs();
  viewRows = baseRows.slice();
  sortRows();
  renderTable(currentHeaders, viewRows);
  toast("Reset filtrow", "info");
});

filter1PickBtn.addEventListener("click", () => {
  lastPickerTriggerEl = filter1PickBtn;
  openColumnPicker("filter1");
});
filter2PickBtn.addEventListener("click", () => {
  lastPickerTriggerEl = filter2PickBtn;
  openColumnPicker("filter2");
});
datePickBtn.addEventListener("click", () => {
  lastPickerTriggerEl = datePickBtn;
  openColumnPicker("date");
});

quickRangeButtons.forEach((btn) => {
  btn.addEventListener("click", () => {
    const days = parseInt(btn.dataset.range || "30", 10);
    dateModeEl.value = "last_n_days";
    lastDaysEl.value = String(days);
    updateDateChipsActive();
    applyFilters();
    sortRows();
    renderTable(currentHeaders, viewRows);
    updateFilterBadge();
  });
});

selectAllBtn.addEventListener("click", () => {
  columnListEl.querySelectorAll("input[type=checkbox]").forEach((cb) => {
    cb.checked = true;
  });
});

clearAllBtn.addEventListener("click", () => {
  columnListEl.querySelectorAll("input[type=checkbox]").forEach((cb) => {
    cb.checked = false;
  });
});

applyPickBtn.addEventListener("click", () => {
  if (!activePickerKey) return;
  const checked = Array.from(columnListEl.querySelectorAll("input[type=checkbox]"))
    .filter((cb) => cb.checked)
    .map((cb) => cb.value);
  if (checked.length === currentHeaders.length) {
    columnSelections[activePickerKey].clear();
  } else {
    columnSelections[activePickerKey] = new Set(checked);
  }
  updateColumnSummary();
  updateFilterBadge();
  closeColumnPicker();
});

columnPickerEl.addEventListener("click", (e) => {
  if (e.target === columnPickerEl) closeColumnPicker();
});

columnPickerEl.addEventListener("keydown", handlePickerKeydown);

closePickerBtn.addEventListener("click", closeColumnPicker);
columnSearchEl.addEventListener("input", filterColumnList);

exportCsvBtn.addEventListener("click", exportCsv);
if (resetSortBtn) {
  resetSortBtn.addEventListener("click", () => {
    sortState = { col: "", dir: "asc" };
    applyFilters();
    renderTable(currentHeaders, viewRows);
    updateSortControls();
    toast("Przywrocono domyslne sortowanie", "info");
  });
}
saveBtn.addEventListener("click", () => {
  toast("Wersja webowa nie nadpisuje pliku. Użyj „Zapisz jako…”", "info");
});
saveAsBtn.addEventListener("click", saveWorkbookAs);
resetWidthsBtn.addEventListener("click", () => {
  manualColumnWidths = {};
  renderTable(currentHeaders, viewRows);
  toast("Przywrocono automatyczne szerokosci", "info");
});

tbodyEl.addEventListener("dblclick", (e) => {
  const td = e.target.closest("td");
  if (!td) return;
  const tr = td.parentElement;
  const rowIndex0 = tr.dataset.rowIndex ? parseInt(tr.dataset.rowIndex, 10) : null;
  const colIndex0 = td.dataset.colIndex ? parseInt(td.dataset.colIndex, 10) : null;
  if (rowIndex0 === null || colIndex0 === null) return;

  if (!workbook || !currentSheetName) return;
  const sheet = workbook.Sheets[currentSheetName];
  const absoluteCol = currentStartCol + colIndex0;
  const cellRef = XLSX.utils.encode_cell({ r: rowIndex0, c: absoluteCol });
  const cell = sheet ? sheet[cellRef] : null;
  if (cell && cell.f) {
    toast("Edycja formul jest zablokowana", "warning");
    return;
  }

  const rowObj = viewRows.find((r) => r.rowIndex0 === rowIndex0);
  if (!rowObj) return;

  const oldValue = rowObj.values[colIndex0];
  const input = document.createElement("input");
  input.className = "cell-editor";
  input.value = oldValue == null ? "" : String(oldValue);
  td.innerHTML = "";
  td.appendChild(input);
  input.focus();
  input.select();

  const commit = () => {
    const parsed = parseInputValue(input.value);
    if (parsed && parsed.type === "formula") {
      toast("Edycja formul jest zablokowana", "warning");
      renderTable(currentHeaders, viewRows);
      return;
    }
    if (!parsed) {
      rowObj.values[colIndex0] = null;
      rowObj.display[colIndex0] = "";
      updateSheetCell(rowIndex0, colIndex0, null);
      if (!valuesEqual(oldValue, null)) setDirtyState(true);
    } else {
      rowObj.values[colIndex0] = parsed.value;
      rowObj.display[colIndex0] = toDisplay(parsed.value);
      updateSheetCell(rowIndex0, colIndex0, parsed);
      if (!valuesEqual(oldValue, parsed.value)) setDirtyState(true);
    }
    renderTable(currentHeaders, viewRows);
  };

  const cancel = () => {
    renderTable(currentHeaders, viewRows);
  };

  input.addEventListener("keydown", (evt) => {
    if (evt.key === "Enter") commit();
    if (evt.key === "Escape") cancel();
  });
  input.addEventListener("blur", commit);
});

[searchQueryEl, searchQuery2El, onlyNonEmptyEl, dateModeEl, dateFromEl, dateToEl, lastDaysEl].forEach((el) => {
  el.addEventListener("input", updateFilterBadge);
  el.addEventListener("change", updateFilterBadge);
});
[dateModeEl, lastDaysEl].forEach((el) => {
  el.addEventListener("change", updateDateChipsActive);
  el.addEventListener("input", updateDateChipsActive);
});

searchQueryEl.addEventListener("input", syncQuickSearchInputs);

maxRowsEl.addEventListener("change", () => {
  saveMaxRowsPreference();
  renderTable(currentHeaders, viewRows);
});

if (excelLayoutToggleEl) {
  excelLayoutToggleEl.addEventListener("click", () => {
    setExcelLayoutEnabled(!isExcelLayoutEnabled());
    saveExcelLayoutPreference();
    renderTable(currentHeaders, viewRows);
  });
}

initIntroSplash();
initTheme();
loadMaxRowsPreference();
loadExcelLayoutPreference();
attachResizeHandlers();
updateNetworkBadge();
window.addEventListener("online", updateNetworkBadge);
window.addEventListener("offline", updateNetworkBadge);

themeToggle.addEventListener("click", () => {
  const next = rootEl.getAttribute("data-theme") === "dark" ? "light" : "dark";
  setTheme(next);
});

if (brandRefreshBtn) {
  brandRefreshBtn.addEventListener("click", () => {
    window.location.reload();
  });

  const expandLogo = () => {
    brandRefreshBtn.classList.add("expanded");
    if (heroRightEl) heroRightEl.classList.add("expanded");
  };
  const collapseLogo = () => {
    brandRefreshBtn.classList.remove("expanded");
    if (heroRightEl) heroRightEl.classList.remove("expanded");
  };

  brandRefreshBtn.addEventListener("mouseenter", expandLogo);
  brandRefreshBtn.addEventListener("mouseleave", collapseLogo);
  brandRefreshBtn.addEventListener("pointerenter", expandLogo);
  brandRefreshBtn.addEventListener("pointerleave", collapseLogo);
  brandRefreshBtn.addEventListener("focus", expandLogo);
  brandRefreshBtn.addEventListener("blur", collapseLogo);
  brandRefreshBtn.addEventListener("touchstart", expandLogo, { passive: true });

  window.addEventListener("pageshow", collapseLogo);
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden") collapseLogo();
  });
}

function toggleSidebar() {
  setSidebarOpen(!isSidebarOpen());
  syncSidebarHandle();
}

function syncSidebarHandle() {
  if (panelToggle) {
    panelToggle.setAttribute("aria-expanded", isSidebarOpen() ? "true" : "false");
    panelToggle.textContent = isSidebarOpen() ? "Zamknij filtry" : "Filtry";
  }
  if (panelHandle) {
    panelHandle.textContent = "";
    panelHandle.setAttribute("aria-expanded", isSidebarOpen() ? "true" : "false");
    panelHandle.setAttribute("aria-label", isSidebarOpen() ? "Zamknij panel filtrow" : "Otworz panel filtrow");
    panelHandle.setAttribute("title", isSidebarOpen() ? "Schowaj filtry" : "Pokaz filtry");
  }
}

function setReadingMode(enabled) {
  rootEl.classList.toggle("reading", enabled);
  if (enabled) {
    if (quickSearchWrap) quickSearchWrap.classList.remove("hidden");
    if (readingToggle) readingToggle.textContent = "Tryb standardowy";
  } else {
    if (quickSearchWrap) quickSearchWrap.classList.add("hidden");
    if (readingToggle) readingToggle.textContent = "Tryb szybkie szukanie";
  }
  syncSidebarHandle();
}

panelToggle.addEventListener("click", toggleSidebar);
if (panelHandle) panelHandle.addEventListener("click", toggleSidebar);
if (sidebarScrim) sidebarScrim.addEventListener("click", () => setSidebarOpen(false));
if (readingToggle) {
  readingToggle.addEventListener("click", () => {
    const enabled = !rootEl.classList.contains("reading");
    setReadingMode(enabled);
  });
}

document.addEventListener("click", (e) => {
  if (!isSidebarOpen()) return;
  if (sidebarEl && sidebarEl.contains(e.target)) return;
  if (panelToggle && panelToggle.contains(e.target)) return;
  if (panelHandle && panelHandle.contains(e.target)) return;
  if (columnPickerEl && !columnPickerEl.classList.contains("hidden") && columnPickerEl.contains(e.target)) return;
  if (quickSearchPopupEl && !quickSearchPopupEl.classList.contains("hidden") && quickSearchPopupEl.contains(e.target)) return;
  setSidebarOpen(false);
});


dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.classList.add("dragover");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("dragover");
});

dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropZone.classList.remove("dragover");
  const file = e.dataTransfer.files[0];
  handleFile(file);
});

dropZone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") {
    e.preventDefault();
    fileInput.click();
  }
});

sheetSelect.addEventListener("change", () => {
  if (!workbook) return;
  setStatus("Wybrano arkusz");
  applyAutoHeaderRowIfEnabled();
});

if (autoHeaderRowEl) {
  autoHeaderRowEl.addEventListener("change", () => {
    if (applyAutoHeaderRowIfEnabled()) {
      toast("Wykryto wiersz nagłówka", "info");
    }
  });
}

document.addEventListener("keydown", (e) => {
  const meta = e.ctrlKey || e.metaKey;
  if (meta && e.key === "Enter") {
    e.preventDefault();
    applyFilterBtn.click();
  }
  if (meta && e.shiftKey && e.key.toLowerCase() === "s") {
    e.preventDefault();
    saveAsBtn.click();
  }
  if (meta && e.shiftKey && e.key.toLowerCase() === "e") {
    e.preventDefault();
    exportCsvBtn.click();
  }
  if (meta && e.shiftKey && e.key.toLowerCase() === "f") {
    e.preventDefault();
    resetFiltersBtn.click();
  }
  if (meta && e.shiftKey && e.key.toLowerCase() === "w") {
    e.preventDefault();
    resetWidthsBtn.click();
  }
  if (meta && e.key.toLowerCase() === "k") {
    e.preventDefault();
    lastPickerTriggerEl = filter1PickBtn;
    openColumnPicker("filter1");
  }
  if (meta && e.key === "/") {
    e.preventDefault();
    themeToggle.click();
  }
  if (meta && e.key === "f") {
    e.preventDefault();
    if (quickSearchPopupEl && !quickSearchPopupEl.classList.contains("hidden")) {
      quickSearchPopupEl.classList.add("hidden");
    } else if (currentHeaders.length && quickSearchPopupEl && quickSearchPopupInput) {
      quickSearchPopupInput.value = searchQueryEl.value || "";
      quickSearchPopupEl.classList.remove("hidden");
      quickSearchPopupInput.focus();
    } else if (!currentHeaders.length) {
      toast("Wczytaj arkusz, żeby szukać", "info");
    }
  }
  if (e.key === "Escape" && !columnPickerEl.classList.contains("hidden")) {
    closeColumnPicker();
  }
  if (e.key === "Escape" && quickSearchPopupEl && !quickSearchPopupEl.classList.contains("hidden")) {
    quickSearchPopupEl.classList.add("hidden");
  }
  if (e.key === "Escape" && isSidebarOpen()) {
    setSidebarOpen(false);
  }
});

setEmptyState(DEFAULT_EMPTY_TITLE, DEFAULT_EMPTY_SUB);
updateDateChipsActive();
updateQuickSearchColumnButtons();
updateSortControls();
setDirtyState(false);
syncQuickSearchInputs();
setSidebarOpen(true);
syncSidebarHandle();

const xlsxReady = isXlsxAvailable(false);
setRuntimeAvailability(xlsxReady);
if (!xlsxReady) {
  setEmptyState(
    "Brak biblioteki XLSX",
    "Aplikacja nie zaladowala silnika arkuszy. Odswiez strone i sprawdz polaczenie z internetem."
  );
  setStatus("Brak biblioteki XLSX");
}

window.addEventListener("beforeunload", (e) => {
  if (!hasUnsavedChanges) return;
  e.preventDefault();
  e.returnValue = "";
});

if ("serviceWorker" in navigator) {
  navigator.serviceWorker.register("sw.js?v=20260323-5").then((registration) => {
    registration.update().catch(() => {});
  }).catch(() => {});
}
