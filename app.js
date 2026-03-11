const rootEl = document.documentElement;
const logEl = document.getElementById("log");
const statusEl = document.getElementById("status");
const tableEl = document.getElementById("dataTable");
const theadEl = tableEl.querySelector("thead");
const tbodyEl = tableEl.querySelector("tbody");
const tableWrapEl = document.getElementById("tableWrap");
const emptyStateEl = document.getElementById("emptyState");
const emptyTitleEl = document.getElementById("emptyTitle");
const emptySubEl = document.getElementById("emptySub");

const fileInput = document.getElementById("fileInput");
const dropZone = document.getElementById("dropZone");
const fileNameEl = document.getElementById("fileName");
const fileNameTextEl = document.getElementById("fileNameText");
const sheetSelect = document.getElementById("sheetSelect");
const headerRowEl = document.getElementById("headerRow");
const autoDetectBtn = document.getElementById("autoDetectBtn");
const displayModeEl = document.getElementById("displayMode");
const maxRowsEl = document.getElementById("maxRows");
const loadBtn = document.getElementById("loadBtn");
const hideEmptyColsEl = document.getElementById("hideEmptyCols");

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
const resetHeightsBtn = document.getElementById("resetHeightsBtn");
const toggleCellStylesEl = document.getElementById("toggleCellStyles");

const columnPickerEl = document.getElementById("columnPicker");
const columnListEl = document.getElementById("columnList");
const columnSearchEl = document.getElementById("columnSearch");
const selectAllBtn = document.getElementById("selectAllBtn");
const clearAllBtn = document.getElementById("clearAllBtn");
const applyPickBtn = document.getElementById("applyPickBtn");
const closePickerBtn = document.getElementById("closePicker");

const themeToggle = document.getElementById("themeToggle");
const panelToggle = document.getElementById("panelToggle");
const readingToggle = document.getElementById("readingToggle");
const panelHandle = document.getElementById("panelHandle");
const loadingOverlayEl = document.getElementById("loadingOverlay");
const loadingTextEl = document.getElementById("loadingText");
const toastContainerEl = document.getElementById("toastContainer");
const readingSearchEl = document.getElementById("readingSearch");

const quickRangeButtons = Array.from(document.querySelectorAll(".chip[data-range]"));

let workbook = null;
let currentHeaders = [];
let currentHeaderStyles = [];
let currentColIndices = [];
let baseRows = [];
let viewRows = [];
let currentFileName = "";
let currentSheetName = "";
let currentHeaderRow1 = 1;
let currentMerges = [];
let currentDataStartRow0 = 0;
let currentDataEndRow0 = 0;
let currentColIndexMap = {};
let currentDataStartRow1 = 1;
let currentColumnTypes = [];
let readingMode = false;
let prevSidebarCollapsed = false;
const columnSelections = {
  filter1: new Set(),
  filter2: new Set(),
  date: new Set(),
};
let activePickerKey = null;
let sortState = { col: "", dir: "asc" };
let manualColumnWidths = {};
let manualRowHeights = {};
let autoRangeEndRow = null;

const THEME_KEY = "excel-workbench-theme";
const MAX_ROWS_KEY = "excel-workbench-max-rows";
const INTRO_PLAYED_KEY = "introPlayed";

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

function setEmptyState(title, subtitle) {
  emptyTitleEl.textContent = title;
  emptySubEl.textContent = subtitle;
  emptyStateEl.classList.remove("hidden");
  tableWrapEl.classList.add("hidden");
}

function showTable() {
  emptyStateEl.classList.add("hidden");
  tableWrapEl.classList.remove("hidden");
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

function normalizeRgb(rgb) {
  if (!rgb || typeof rgb !== "string") return null;
  const cleaned = rgb.trim();
  if (cleaned.length === 8) return `#${cleaned.slice(2).toLowerCase()}`;
  if (cleaned.length === 6) return `#${cleaned.toLowerCase()}`;
  return null;
}

function readableTextColor(bg) {
  if (!bg || bg[0] !== "#" || bg.length !== 7) return null;
  const r = parseInt(bg.slice(1, 3), 16);
  const g = parseInt(bg.slice(3, 5), 16);
  const b = parseInt(bg.slice(5, 7), 16);
  const luminance = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
  return luminance < 0.5 ? "#f7f7f7" : "#1f2a24";
}

function extractHeaderStyle(cell) {
  if (!cell || !cell.s) return null;
  const fill = cell.s.fill;
  const font = cell.s.font;
  const bg = fill && fill.fgColor ? normalizeRgb(fill.fgColor.rgb) : null;
  const color = font && font.color ? normalizeRgb(font.color.rgb) : null;
  const bold = font && !!font.bold;
  if (!bg && !color && !bold) return null;
  return { bg, color, bold };
}

function extractCellStyle(cell) {
  if (!cell || !cell.s) return null;
  const fill = cell.s.fill;
  const font = cell.s.font;
  const bg = fill && fill.fgColor ? normalizeRgb(fill.fgColor.rgb) : null;
  const color = font && font.color ? normalizeRgb(font.color.rgb) : null;
  const bold = font && !!font.bold;
  if (!bg && !color && !bold) return null;
  return { bg, color, bold };
}

function isLikelyDateString(value) {
  if (typeof value !== "string") return false;
  const v = value.trim();
  if (!v) return false;
  if (/[-/.]/.test(v)) return true;
  if (/[A-Za-ząćęłńóśźż]{3,}/i.test(v)) return true;
  return false;
}

function inferColumnTypes(rows, limit = 200) {
  const counts = currentHeaders.map(() => ({ total: 0, num: 0, date: 0, text: 0 }));
  const sample = rows.slice(0, limit);
  sample.forEach((row) => {
    row.values.forEach((v, i) => {
      if (v === null || v === "") return;
      counts[i].total += 1;
      if (v instanceof Date) {
        counts[i].date += 1;
        return;
      }
      if (typeof v === "number") {
        counts[i].num += 1;
        return;
      }
      if (typeof v === "string") {
        if (isLikelyDateString(v) && parseDateFlexible(v)) {
          counts[i].date += 1;
        } else if (!Number.isNaN(Number(v))) {
          counts[i].num += 1;
        } else {
          counts[i].text += 1;
        }
        return;
      }
      counts[i].text += 1;
    });
  });

  return counts.map((c) => {
    if (!c.total) return "";
    const dateRatio = c.date / c.total;
    const numRatio = c.num / c.total;
    const textRatio = c.text / c.total;
    if (dateRatio >= 0.6) return "Data";
    if (numRatio >= 0.6) return "Liczba";
    if (textRatio >= 0.6) return "Tekst";
    return "";
  });
}

function detectHeaderRow(sheet) {
  if (!sheet || !sheet["!ref"]) return 1;
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const merges = sheet["!merges"] || [];
  const maxScan = Math.min(range.e.r, range.s.r + 50);
  let best = { row: range.s.r, score: -Infinity };

  for (let r = range.s.r; r <= maxScan; r++) {
    let text = 0;
    let num = 0;
    let any = false;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      if (!cell || cell.v === null || cell.v === "") continue;
      any = true;
      if (typeof cell.v === "string") text += 1;
      else if (typeof cell.v === "number") num += 1;
      else text += 1;
    }
    if (!any || text === 0) continue;
    let score = text * 2 - num;
    const hasMerge = merges.some((m) => m.s.r <= r && m.e.r >= r && m.e.c > m.s.c);
    if (hasMerge) score -= 1;
    if (score > best.score) best = { row: r, score };
  }

  return best.row + 1;
}

function detectDataEndRow(sheet, dataStartRow1) {
  if (!sheet || !sheet["!ref"]) return dataStartRow1;
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const start0 = Math.max(range.s.r, dataStartRow1 - 1);
  let lastDataRow0 = start0;
  let emptyStreak = 0;
  for (let r = start0; r <= range.e.r; r++) {
    let any = false;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.v !== null && cell.v !== "") {
        any = true;
        break;
      }
    }
    if (any) {
      lastDataRow0 = r;
      emptyStreak = 0;
    } else {
      emptyStreak += 1;
      if (emptyStreak >= 2) break;
    }
  }
  return lastDataRow0 + 1;
}

function getDisplayValue(row, index) {
  if (row && Array.isArray(row.display) && index < row.display.length) {
    return row.display[index];
  }
  if (row && Array.isArray(row.values) && index < row.values.length) {
  return toDisplay(row.values[index]);
}

function computeRowHeaderWidth(rows) {
  let maxRow = currentHeaderRow1 || 1;
  rows.forEach((row) => {
    if (typeof row.rowIndex0 === "number") {
      maxRow = Math.max(maxRow, row.rowIndex0 + 1);
    }
  });
  const digits = String(maxRow).length;
  return Math.min(64, Math.max(32, 12 + digits * 8));
}

function buildMergeMaps(rows) {
  const merges = Array.isArray(currentMerges) ? currentMerges : [];
  if (!merges.length) return { skip: new Set(), topLeft: new Map() };
  const visibleRows = new Set(rows.map((r) => r.rowIndex0));
  const skip = new Set();
  const topLeft = new Map();

  merges.forEach((m) => {
    const s = m.s;
    const e = m.e;
    if (s.r < currentDataStartRow0 || e.r > currentDataEndRow0) return;
    for (let c = s.c; c <= e.c; c++) {
      if (currentColIndexMap[c] === undefined) return;
    }
    for (let r = s.r; r <= e.r; r++) {
      if (!visibleRows.has(r)) return;
    }
    const colStart = currentColIndexMap[s.c];
    const colEnd = currentColIndexMap[e.c];
    if (colStart === undefined || colEnd === undefined) return;
    const colSpan = colEnd - colStart + 1;
    const rowSpan = e.r - s.r + 1;
    if (colSpan <= 1 && rowSpan <= 1) return;
    topLeft.set(`${s.r}:${s.c}`, { colspan: colSpan, rowspan: rowSpan });
    for (let r = s.r; r <= e.r; r++) {
      for (let c = s.c; c <= e.c; c++) {
        if (r === s.r && c === s.c) continue;
        skip.add(`${r}:${c}`);
      }
    }
  });

  return { skip, topLeft };
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
  const cellRef = XLSX.utils.encode_cell({ r: rowIndex0, c: colIndex0 });
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
      if (criterion.mode === "Rowna sie" && candidates.some((c) => c === query)) matched = true;
      if (criterion.mode === "Zaczyna sie" && candidates.some((c) => c.startsWith(query))) matched = true;
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

function computeColumnWidths(headers, rows) {
  const widths = headers.map(() => 0);
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
    const min = 80;
    const max = 520;
    const manual = manualColumnWidths[i];
    if (manual) return Math.max(min, Math.min(max, manual));
    return Math.max(min, Math.min(max, Math.ceil(raw)));
  });
}

function renderTable(headers, rows) {
  if (!headers.length) {
    setStatus("Brak danych");
    setEmptyState("Wczytaj plik Excel", "Przeciagnij plik lub wybierz go z dysku, aby zaczac prace.");
    return;
  }
  if (!rows.length) {
    setStatus("Wierszy: 0");
    setEmptyState("Brak wynikow", "Zmien filtry albo wybierz inny arkusz.");
    return;
  }

  showTable();
  theadEl.innerHTML = "";
  tbodyEl.innerHTML = "";

  const widths = computeColumnWidths(headers, rows);
  const rowHeaderWidth =
    typeof computeRowHeaderWidth === "function" ? computeRowHeaderWidth(rows) : 36;
  const mergeMaps =
    typeof buildMergeMaps === "function"
      ? buildMergeMaps(rows)
      : { skip: new Set(), topLeft: new Map() };
  const colgroup = document.createElement("colgroup");
  const rowCol = document.createElement("col");
  rowCol.style.width = `${rowHeaderWidth}px`;
  colgroup.appendChild(rowCol);
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
  corner.className = "row-head col-head corner";
  corner.textContent = "";
  guideRow.appendChild(corner);
  headers.forEach((h, i) => {
    const th = document.createElement("th");
    th.className = "col-head";
    const colIndex = currentColIndices[i] ?? i;
    th.textContent = XLSX.utils.encode_col(colIndex);
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
  rowHead.className = "row-head data-head";
  rowHead.textContent = currentHeaderRow1 ? String(currentHeaderRow1) : "1";
  headRow.appendChild(rowHead);
  headers.forEach((h, i) => {
    const th = document.createElement("th");
    th.className = "data-head";
    const label = document.createElement("span");
    label.className = "header-label";
    label.textContent = h;
    th.appendChild(label);
    const typeLabel = currentColumnTypes[i];
    if (typeLabel) {
      const badge = document.createElement("span");
      badge.className = "type-badge";
      badge.textContent = typeLabel;
      th.appendChild(badge);
    }
    const style = currentHeaderStyles && currentHeaderStyles[i] ? currentHeaderStyles[i] : null;
    if (style) {
      if (style.bg) {
        th.style.backgroundColor = style.bg;
        th.style.color = style.color || readableTextColor(style.bg) || "";
      }
      if (style.bold) th.style.fontWeight = "700";
      if (style.color && !style.bg) th.style.color = style.color;
    }
    th.addEventListener("click", () => {
      if (sortState.col === h) {
        sortState.dir = sortState.dir === "asc" ? "desc" : "asc";
      } else {
        sortState.col = h;
        sortState.dir = "asc";
      }
      sortRows();
      renderTable(currentHeaders, viewRows);
    });

    if (sortState.col === h) {
      const arrow = document.createElement("span");
      arrow.className = "sort-arrow";
      arrow.textContent = sortState.dir === "asc" ? "▲" : "▼";
      th.appendChild(arrow);
    }

    headRow.appendChild(th);
  });
  theadEl.appendChild(headRow);

  const limit = Math.max(1, parseInt(maxRowsEl.value || "200", 10));
  rows.slice(0, limit).forEach((row) => {
    const tr = document.createElement("tr");
    if (typeof row.rowIndex0 === "number") {
      tr.dataset.rowIndex = String(row.rowIndex0);
    }
    const rowTh = document.createElement("th");
    rowTh.className = "row-head";
    const rowNumber = typeof row.rowIndex0 === "number" ? row.rowIndex0 + 1 : "";
    rowTh.textContent = rowNumber ? String(rowNumber) : "";
    if (typeof row.rowIndex0 === "number" && manualRowHeights[row.rowIndex0]) {
      const h = manualRowHeights[row.rowIndex0];
      tr.style.height = `${h}px`;
      rowTh.style.height = `${h}px`;
      rowTh.style.lineHeight = `${h - 2}px`;
    }
    const rowResizer = document.createElement("div");
    rowResizer.className = "row-resizer";
    if (typeof row.rowIndex0 === "number") {
      rowResizer.dataset.rowIndex = String(row.rowIndex0);
    }
    rowTh.appendChild(rowResizer);
    tr.appendChild(rowTh);
    row.values.forEach((v, i) => {
      const sheetCol = currentColIndices[i] ?? i;
      if (mergeMaps.skip.has(`${row.rowIndex0}:${sheetCol}`)) {
        return;
      }
      const td = document.createElement("td");
      td.textContent = getDisplayValue(row, i);
      td.dataset.colIndex = String(i);
      if (toggleCellStylesEl && toggleCellStylesEl.checked && row.styles && row.styles[i]) {
        const style = row.styles[i];
        if (style.bg) {
          td.style.backgroundColor = style.bg;
          td.style.color = style.color || readableTextColor(style.bg) || "";
        }
        if (style.bold) td.style.fontWeight = "600";
        if (style.color && !style.bg) td.style.color = style.color;
      }
      const key = `${row.rowIndex0}:${sheetCol}`;
      const merge = mergeMaps.topLeft.get(key);
      if (merge) {
        td.colSpan = merge.colspan;
        td.rowSpan = merge.rowspan;
      }
      if (typeof row.rowIndex0 === "number" && manualRowHeights[row.rowIndex0]) {
        const h = manualRowHeights[row.rowIndex0];
        td.style.height = `${h}px`;
        td.style.lineHeight = `${h - 2}px`;
      }
      tr.appendChild(td);
    });
    tbodyEl.appendChild(tr);
  });

  let status = `Wierszy: ${rows.length} (pokazano: ${Math.min(rows.length, limit)})`;
  if (currentHeaderRow1 && currentDataStartRow1) {
    status += ` • Naglowek: ${currentHeaderRow1} • Dane od: ${currentDataStartRow1}`;
  }
  setStatus(status);
}

function buildRows(sheet, dataStartRow1, options = {}) {
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const headerRow0 = Math.max(range.s.r, dataStartRow1 - 2);
  const dataStartRow0 = Math.max(headerRow0 + 1, dataStartRow1 - 1);
  const endRow0 = options.endRow1 ? Math.min(range.e.r, options.endRow1 - 1) : range.e.r;

  const colIndices = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    if (options.hideEmptyCols) {
      let hasData = false;
      for (let r = dataStartRow0; r <= endRow0; r++) {
        const cell = sheet[XLSX.utils.encode_cell({ r, c })];
        if (cell && cell.v !== null && cell.v !== "") {
          hasData = true;
          break;
        }
      }
      if (!hasData) continue;
    }
    colIndices.push(c);
  }
  const colIndexMap = {};
  colIndices.forEach((c, idx) => {
    colIndexMap[c] = idx;
  });

  const headers = [];
  const headerStyles = [];
  for (const c of colIndices) {
    let cell = sheet[XLSX.utils.encode_cell({ r: headerRow0, c })];
    let v = cell ? cell.v : null;
    if (v === null || v === "") {
      for (let rr = headerRow0 - 1; rr >= Math.max(range.s.r, headerRow0 - 3); rr--) {
        const above = sheet[XLSX.utils.encode_cell({ r: rr, c })];
        if (above && above.v !== null && above.v !== "") {
          v = above.v;
          cell = above;
          break;
        }
      }
    }
    headers.push(v ? String(v).trim() : XLSX.utils.encode_col(c));
    headerStyles.push(extractHeaderStyle(cell));
  }

  const rows = [];
  for (let r = dataStartRow0; r <= endRow0; r++) {
    const values = [];
    const display = [];
    const styles = [];
    let any = false;
    for (const c of colIndices) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      let v = cell ? cell.v : null;
      let shown = cell && cell.w ? String(cell.w) : toDisplay(v);
      if (displayModeEl.value === "Formuly" && cell && cell.f) {
        v = "=" + cell.f;
        shown = v;
      }
      values.push(v);
      display.push(shown);
      styles.push(extractCellStyle(cell));
      if (v !== null && v !== "") any = true;
    }
    if (!any) continue;
    rows.push({ values, display, rawValues: values, styles, rowIndex0: r });
  }
  const merges = Array.isArray(sheet["!merges"]) ? sheet["!merges"] : [];
  return {
    headers,
    rows,
    headerStyles,
    colIndices,
    colIndexMap,
    dataStartRow0,
    dataEndRow0: endRow0,
    merges,
  };
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

function openColumnPicker(key) {
  if (!currentHeaders.length) return;
  activePickerKey = key;
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
  let startY = 0;
  let startH = 0;

  const start = (e) => {
    const handle = e.target.closest(".col-resizer");
    const rowHandle = e.target.closest(".row-resizer");
    if (!handle && !rowHandle) return;
    e.preventDefault();
    if (handle) {
      const colIndex = parseInt(handle.dataset.colIndex, 10);
      const th = handle.parentElement;
      active = { type: "col", colIndex, th };
      startX = e.clientX || (e.touches && e.touches[0].clientX) || 0;
      startW = th.getBoundingClientRect().width;
    } else if (rowHandle) {
      const rowIndex = parseInt(rowHandle.dataset.rowIndex, 10);
      const tr = rowHandle.closest("tr");
      active = { type: "row", rowIndex, tr };
      startY = e.clientY || (e.touches && e.touches[0].clientY) || 0;
      startH = tr.getBoundingClientRect().height;
    }
    document.body.classList.add("resizing");
  };

  const move = (e) => {
    if (!active) return;
    if (active.type === "col") {
      const x = e.clientX || (e.touches && e.touches[0].clientX) || 0;
      const delta = x - startX;
      const next = Math.max(80, Math.min(520, Math.round(startW + delta)));
      manualColumnWidths[active.colIndex] = next;
      renderTable(currentHeaders, viewRows);
    } else if (active.type === "row") {
      const y = e.clientY || (e.touches && e.touches[0].clientY) || 0;
      const delta = y - startY;
      const next = Math.max(26, Math.min(200, Math.round(startH + delta)));
      manualRowHeights[active.rowIndex] = next;
      renderTable(currentHeaders, viewRows);
    }
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

async function handleFile(file) {
  if (!file) return;
  try {
    setLoading(true, "Wczytywanie pliku...");
    const data = await file.arrayBuffer();
    workbook = XLSX.read(data, { cellDates: true, cellStyles: true });
    sheetSelect.innerHTML = "";
    autoRangeEndRow = null;
    workbook.SheetNames.forEach((s) => {
      const opt = document.createElement("option");
      opt.value = s;
      opt.textContent = s;
      sheetSelect.appendChild(opt);
    });
    currentFileName = file.name;
    fileNameTextEl.textContent = file.name;
    fileNameEl.classList.remove("hidden");
    dropZone.classList.add("has-file");
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
  toast("Zapisano plik", "success");
  log(`Zapisano plik: ${filename}`, "success");
}

function saveWorkbookAs() {
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
  toast("Zapisano plik", "success");
  log(`Zapisano plik: ${name}`, "success");
}

fileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  handleFile(file);
});

loadBtn.addEventListener("click", () => {
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
      const headerRow = Math.max(1, parseInt(headerRowEl.value || "1", 10));
      const dataStartRow1 = headerRow + 1;
      currentSheetName = sheetName;
      currentHeaderRow1 = headerRow;
      currentDataStartRow1 = dataStartRow1;
      const endRow1 = autoRangeEndRow && autoRangeEndRow >= dataStartRow1 ? autoRangeEndRow : null;
      const data = buildRows(sheet, dataStartRow1, {
        endRow1,
        hideEmptyCols: hideEmptyColsEl ? hideEmptyColsEl.checked : false,
      });
      currentHeaders = data.headers;
      currentHeaderStyles = data.headerStyles || [];
      currentColIndices = data.colIndices || [];
      currentColIndexMap = data.colIndexMap || {};
      currentDataStartRow0 = data.dataStartRow0 ?? 0;
      currentDataEndRow0 = data.dataEndRow0 ?? 0;
      currentMerges = data.merges || [];
      currentColumnTypes = inferColumnTypes(data.rows || []);
      baseRows = data.rows;
      viewRows = data.rows.slice();
      sortState = { col: currentHeaders[0] || "", dir: "asc" };
      manualColumnWidths = {};
      columnSelections.filter1.clear();
      columnSelections.filter2.clear();
      columnSelections.date.clear();
      updateColumnSummary();
      updateFilterBadge();
      renderTable(currentHeaders, viewRows);
      toast("Arkusz wczytany", "success");
      log(`Wczytano arkusz: ${sheetName}`, "success");
    } finally {
      setLoading(false);
    }
  }, 50);
});

autoDetectBtn.addEventListener("click", () => {
  if (!workbook) {
    toast("Najpierw wybierz plik", "warning");
    return;
  }
  const sheetName = sheetSelect.value;
  const sheet = workbook.Sheets[sheetName];
  if (!sheet || !sheet["!ref"]) {
    toast("Brak arkusza", "error");
    return;
  }
  const headerRow1 = detectHeaderRow(sheet);
  const dataStartRow1 = Math.max(1, headerRow1 + 1);
  headerRowEl.value = String(headerRow1);
  autoRangeEndRow = detectDataEndRow(sheet, dataStartRow1);
  currentHeaderRow1 = headerRow1;
  currentDataStartRow1 = dataStartRow1;
  setStatus(`Wykryto naglowek: ${headerRow1} • Dane od: ${dataStartRow1}`);
  toast(`Wykryto nagłówek: wiersz ${headerRow1}`, "info");
});

applyFilterBtn.addEventListener("click", () => {
  if (!currentHeaders.length) return;
  applyFilters();
  sortRows();
  renderTable(currentHeaders, viewRows);
  updateFilterBadge();
  toast("Zastosowano filtry", "info");
});


resetFiltersBtn.addEventListener("click", () => {
  searchQueryEl.value = "";
  searchQuery2El.value = "";
  filterModeEl.value = "Zawiera";
  filterMode2El.value = "Zawiera";
  onlyNonEmptyEl.checked = false;
  dateModeEl.value = "between";
  dateFromEl.value = "";
  dateToEl.value = "";
  lastDaysEl.value = "30";
  columnSelections.filter1.clear();
  columnSelections.filter2.clear();
  columnSelections.date.clear();
  updateColumnSummary();
  updateFilterBadge();
  viewRows = baseRows.slice();
  sortRows();
  renderTable(currentHeaders, viewRows);
  toast("Reset filtrow", "info");
});

filter1PickBtn.addEventListener("click", () => openColumnPicker("filter1"));
filter2PickBtn.addEventListener("click", () => openColumnPicker("filter2"));
datePickBtn.addEventListener("click", () => openColumnPicker("date"));

quickRangeButtons.forEach((btn) => {
  btn.addEventListener("click", () => {
    const days = parseInt(btn.dataset.range || "30", 10);
    dateModeEl.value = "last_n_days";
    lastDaysEl.value = String(days);
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

closePickerBtn.addEventListener("click", closeColumnPicker);
columnSearchEl.addEventListener("input", filterColumnList);

exportCsvBtn.addEventListener("click", exportCsv);
saveBtn.addEventListener("click", saveWorkbook);
saveAsBtn.addEventListener("click", saveWorkbookAs);
resetWidthsBtn.addEventListener("click", () => {
  manualColumnWidths = {};
  renderTable(currentHeaders, viewRows);
  toast("Przywrocono automatyczne szerokosci", "info");
});
resetHeightsBtn.addEventListener("click", () => {
  manualRowHeights = {};
  renderTable(currentHeaders, viewRows);
  toast("Przywrocono automatyczne wysokosci", "info");
});
toggleCellStylesEl.addEventListener("change", () => {
  renderTable(currentHeaders, viewRows);
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
  const cellRef = XLSX.utils.encode_cell({ r: rowIndex0, c: colIndex0 });
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
    } else {
      rowObj.values[colIndex0] = parsed.value;
      rowObj.display[colIndex0] = toDisplay(parsed.value);
      updateSheetCell(rowIndex0, colIndex0, parsed);
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

let readingSearchTimer = null;
if (readingSearchEl) {
  readingSearchEl.addEventListener("input", () => {
    searchQueryEl.value = readingSearchEl.value;
    updateFilterBadge();
    if (readingSearchTimer) clearTimeout(readingSearchTimer);
    readingSearchTimer = setTimeout(() => {
      applyFilters();
      sortRows();
      renderTable(currentHeaders, viewRows);
    }, 250);
  });
  readingSearchEl.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      applyFilterBtn.click();
    }
  });
}

searchQueryEl.addEventListener("input", () => {
  if (readingSearchEl && readingSearchEl.value !== searchQueryEl.value) {
    readingSearchEl.value = searchQueryEl.value;
  }
});

maxRowsEl.addEventListener("change", () => {
  saveMaxRowsPreference();
  renderTable(currentHeaders, viewRows);
});

headerRowEl.addEventListener("input", () => {
  autoRangeEndRow = null;
});

initIntroSplash();
initTheme();
loadMaxRowsPreference();
attachResizeHandlers();

themeToggle.addEventListener("click", () => {
  const next = rootEl.getAttribute("data-theme") === "dark" ? "light" : "dark";
  setTheme(next);
});

function toggleSidebar() {
  rootEl.classList.toggle("sidebar-collapsed");
  const collapsed = rootEl.classList.contains("sidebar-collapsed");
  if (panelHandle) panelHandle.textContent = collapsed ? "›" : "‹";
}

function toggleReadingMode() {
  readingMode = !rootEl.classList.contains("reading-mode");
  if (readingMode) {
    prevSidebarCollapsed = rootEl.classList.contains("sidebar-collapsed");
    rootEl.classList.add("reading-mode", "sidebar-collapsed");
    if (readingToggle) readingToggle.textContent = "Powrot";
    if (readingSearchEl) {
      readingSearchEl.value = searchQueryEl.value || "";
      readingSearchEl.focus();
    }
  } else {
    rootEl.classList.remove("reading-mode");
    if (!prevSidebarCollapsed) rootEl.classList.remove("sidebar-collapsed");
    if (readingToggle) readingToggle.textContent = "Czytanie";
  }
}

panelToggle.addEventListener("click", toggleSidebar);
if (panelHandle) panelHandle.addEventListener("click", toggleSidebar);
if (readingToggle) readingToggle.addEventListener("click", toggleReadingMode);


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
  autoRangeEndRow = null;
});

document.addEventListener("keydown", (e) => {
  const meta = e.ctrlKey || e.metaKey;
  const isEditable = e.target && (e.target.tagName === "INPUT" || e.target.tagName === "TEXTAREA" || e.target.isContentEditable);
  if (meta && e.key === "Enter") {
    e.preventDefault();
    applyFilterBtn.click();
  }
  if (meta && e.key.toLowerCase() === "f" && !isEditable) {
    e.preventDefault();
    const target = rootEl.classList.contains("reading-mode") && readingSearchEl ? readingSearchEl : searchQueryEl;
    target.focus();
    target.select();
  }
  if (meta && e.key.toLowerCase() === "s") {
    e.preventDefault();
    saveBtn.click();
  }
  if (meta && e.shiftKey && e.key.toLowerCase() === "e") {
    e.preventDefault();
    exportCsvBtn.click();
  }
  if (meta && e.key.toLowerCase() === "k") {
    e.preventDefault();
    openColumnPicker("filter1");
  }
  if (meta && e.key === "/") {
    e.preventDefault();
    themeToggle.click();
  }
  if (e.key === "Escape" && !columnPickerEl.classList.contains("hidden")) {
    closeColumnPicker();
  }
});

setEmptyState("Wczytaj plik Excel", "Przeciagnij plik lub wybierz go z dysku, aby zaczac prace.");

if ("serviceWorker" in navigator) {
  navigator.serviceWorker.register("sw.js?v=20260311").then((registration) => {
    registration.update().catch(() => {});
  }).catch(() => {});
}
