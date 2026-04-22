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
const zoomLevelEl = document.getElementById("zoomLevel");
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
const filterEmptyModeEl = document.getElementById("filterEmptyMode");
const filterNegateEl = document.getElementById("filterNegate");
const filterEmptyMode2El = document.getElementById("filterEmptyMode2");
const filterNegate2El = document.getElementById("filterNegate2");
const onlyNonEmptyEl = document.getElementById("onlyNonEmpty");
const applyFilterBtn = document.getElementById("applyFilterBtn");
const filterBadgeEl = document.getElementById("filterBadge");
const sortColumnSelectEl = document.getElementById("sortColumnSelect");
const sortDirectionSelectEl = document.getElementById("sortDirectionSelect");
const addSortRuleBtn = document.getElementById("addSortRuleBtn");
const sortRulesListEl = document.getElementById("sortRulesList");
const sortPresetSelectEl = document.getElementById("sortPresetSelect");
const saveSortPresetBtn = document.getElementById("saveSortPresetBtn");
const applySortPresetBtn = document.getElementById("applySortPresetBtn");
const deleteSortPresetBtn = document.getElementById("deleteSortPresetBtn");

const dateModeEl = document.getElementById("dateMode");
const dateFromEl = document.getElementById("dateFrom");
const dateToEl = document.getElementById("dateTo");
const lastDaysEl = document.getElementById("lastDays");
const dateEmptyModeEl = document.getElementById("dateEmptyMode");
const dateNegateEl = document.getElementById("dateNegate");
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
const quickSearchModeEl = document.getElementById("quickSearchMode");
const quickSearchColumnsBtn = document.getElementById("quickSearchColumnsBtn");
const quickSearchBtn = document.getElementById("quickSearchBtn");
const wideLongToggleEl = document.getElementById("wideLongToggle");

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
const quickSearchPopupModeEl = document.getElementById("quickSearchPopupMode");
const quickSearchPopupColumnsBtn = document.getElementById("quickSearchPopupColumnsBtn");
const quickSearchPopupBtn = document.getElementById("quickSearchPopupBtn");
const workbookInsightsEl = document.getElementById("workbookInsights");
const sheetInsightsEl = document.getElementById("sheetInsights");
const insightFlagsEl = document.getElementById("insightFlags");
const kpiSummaryEl = document.getElementById("kpiSummary");
const kpiListEl = document.getElementById("kpiList");
const sheetInspectorSummaryEl = document.getElementById("sheetInspectorSummary");
const columnProfilerEl = document.getElementById("columnProfiler");
const sectionNavigatorEl = document.getElementById("sectionNavigator");
const repeatBlockDetectorEl = document.getElementById("repeatBlockDetector");
const durationAnalysisSummaryEl = document.getElementById("durationAnalysisSummary");
const durationAnalysisListEl = document.getElementById("durationAnalysisList");
const aggregationWorkbenchSummaryEl = document.getElementById("aggregationWorkbenchSummary");
const aggregationWorkbenchListEl = document.getElementById("aggregationWorkbenchList");
const formulaSearchEl = document.getElementById("formulaSearch");
const formulaFilterEl = document.getElementById("formulaFilter");
const formulaFunctionFilterEl = document.getElementById("formulaFunctionFilter");
const formulaWorkbenchSummaryEl = document.getElementById("formulaWorkbenchSummary");
const formulaWorkbenchListEl = document.getElementById("formulaWorkbenchList");

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
let currentWorkbookStats = null;
let currentSheetStats = null;
let currentKpiEntries = [];
let currentKpiAnchorRow = 1;
let currentColumnProfiles = [];
let currentSections = [];
let currentRepeatingBlocks = [];
let currentFormulaEntries = [];
let currentDisplayModel = null;
let tableViewMode = "wide";

const columnSelections = {
  filter1: new Set(),
  filter2: new Set(),
  date: new Set(),
};
let activePickerKey = null;
let lastPickerTriggerEl = null;
let sortState = { col: "", dir: "asc" };
let multiSortState = [];
let manualColumnWidths = {};
let hasUnsavedChanges = false;
let focusedCellState = null;
let selectedCellState = null;
let syncingHorizontalScroll = false;
let tooltipHideTimer = null;
let durationAnalysisState = {
  statusFilter: "all",
  sortMetric: "avg",
  showCount: 14,
};
let aggregationWorkbenchState = {
  sourceMode: "auto",
  scopeMode: "filtered",
  headerRowChoice: "auto",
  customHeaderRow: 1,
  groupBy: "",
  measure: "count_rows",
  aggregation: "count",
  matchMode: "contains",
  showCount: 20,
  havingMode: "all",
  havingValue: 10,
  measureFilterMode: "all",
  measureFilterValue: "",
  resultSearch: "",
};
const APP_BUILD_VERSION = "20260421-12";

const THEME_KEY = "excel-workbench-theme";
const MAX_ROWS_KEY = "excel-workbench-max-rows";
const EXCEL_LAYOUT_KEY = "excel-workbench-excel-layout";
const SORT_PRESETS_KEY = "excel-workbench-sort-presets";
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

function applyZoom() {
  if (!tableEl || !zoomLevelEl) return;
  const zoom = parseFloat(zoomLevelEl.value) || 1;
  const baseSize = 12;
  
  if (zoom === 1) {
    tableEl.style.zoom = "";
    tableEl.style.transform = "";
    tableEl.style.fontSize = "";
    tableEl.style.marginRight = "";
    tableEl.style.marginBottom = "";
    return;
  }
  
  if ("zoom" in tableEl.style) {
    tableEl.style.zoom = String(zoom);
  } else {
    tableEl.style.fontSize = `${baseSize * zoom}px`;
  }
}

function valuesEqual(a, b) {
  if (a === b) return true;
  if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
  if ((a === null || a === undefined) && (b === null || b === undefined)) return true;
  return false;
}

function makeHeadersUnique(headers) {
  const seen = new Map();
  return headers.map((header, index) => {
    const base = String(header || `Kolumna ${index + 1}`).trim() || `Kolumna ${index + 1}`;
    const count = seen.get(base) || 0;
    seen.set(base, count + 1);
    return count ? `${base} (${count + 1})` : base;
  });
}

function formatPercent(part, total) {
  if (!total) return "0%";
  return `${Math.round((part / total) * 100)}%`;
}

function createEmptyInsight(text) {
  const el = document.createElement("div");
  el.className = "insight-empty";
  el.textContent = text;
  return el;
}

function renderInsightList(container, items, emptyText) {
  if (!container) return;
  container.replaceChildren();
  if (!items || !items.length) {
    container.appendChild(createEmptyInsight(emptyText));
    return;
  }
  items.forEach((item) => {
    const row = document.createElement("div");
    row.className = `insight-item${item.tone ? ` ${item.tone}` : ""}`;

    const label = document.createElement("div");
    label.className = "insight-label";
    label.textContent = item.label;

    const value = document.createElement("div");
    value.className = "insight-value";
    value.textContent = item.value;

    row.appendChild(label);
    row.appendChild(value);
    container.appendChild(row);
  });
}

function renderInsightFlags(items) {
  if (!insightFlagsEl) return;
  insightFlagsEl.replaceChildren();
  if (!items || !items.length) {
    insightFlagsEl.appendChild(createEmptyInsight("Brak istotnych flag dla aktualnego pliku."));
    return;
  }
  items.forEach((item) => {
    const badge = document.createElement("div");
    badge.className = `insight-flag${item.tone ? ` ${item.tone}` : ""}`;
    badge.textContent = item.label;
    insightFlagsEl.appendChild(badge);
  });
}

function cleanSectionLabel(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function formatColRange(startColAbs, endColAbs = startColAbs) {
  const start = XLSX.utils.encode_col(startColAbs);
  const end = XLSX.utils.encode_col(endColAbs);
  return start === end ? start : `${start}:${end}`;
}

function getCellDisplayText(sheet, rowAbs, colAbs) {
  if (!sheet) return "";
  const ref = XLSX.utils.encode_cell({ r: rowAbs, c: colAbs });
  const cell = sheet[ref];
  if (!cell) return "";
  return cleanSectionLabel(cell.w ?? cell.v ?? "");
}

function inferSectionKindLabel(kind) {
  if (kind === "table") return "Tabela";
  if (kind === "group") return "Blok";
  if (kind === "candidate") return "Nagłówek";
  if (kind === "subheader") return "Sekcja";
  return "Układ";
}

function addSection(sections, seen, entry) {
  if (!entry || !entry.label) return;
  const key = `${entry.kind}|${entry.label}|${entry.rowIndex0 ?? ""}|${entry.headerRow ?? ""}|${entry.colIndex ?? ""}`;
  if (seen.has(key)) return;
  seen.add(key);
  sections.push(entry);
}

function detectSections(sheet, headerRow, data) {
  if (!sheet || !data || !data.headers || !data.headers.length) return [];
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const sections = [];
  const seen = new Set();
  const headerAbsRow = headerRow - 1;

  addSection(sections, seen, {
    kind: "table",
    label: "Tabela danych",
    meta: `Nagłówek: wiersz ${headerRow} • kolumny ${formatColRange(data.startCol || 0, (data.startCol || 0) + data.headers.length - 1)}`,
    tone: "info",
    action: "scroll-top",
  });

  const scanHeaderMax = Math.min(range.e.r, range.s.r + 7);
  for (let r = range.s.r; r <= scanHeaderMax; r++) {
    const texts = [];
    let filled = 0;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const txt = getCellDisplayText(sheet, r, c);
      if (!txt) continue;
      filled += 1;
      if (texts.length < 3) texts.push(txt);
    }
    if (filled < 2) continue;
    addSection(sections, seen, {
      kind: "candidate",
      label: r + 1 === headerRow ? `Aktualny wiersz nagłówka: ${r + 1}` : `Możliwy wiersz nagłówka: ${r + 1}`,
      meta: texts.join(" • "),
      tone: r + 1 === headerRow ? "info" : "",
      action: r + 1 === headerRow ? "scroll-top" : "set-header",
      headerRow: r + 1,
    });
  }

  const merges = Array.isArray(data.merges) ? data.merges : [];
  merges
    .filter((merge) => merge && merge.s && merge.e)
    .sort((a, b) => (a.s.r - b.s.r) || (a.s.c - b.s.c))
    .forEach((merge) => {
      const colspan = merge.e.c - merge.s.c + 1;
      if (colspan < 2) return;
      const label = getCellDisplayText(sheet, merge.s.r, merge.s.c);
      if (!label) return;
      const isAboveHeader = merge.s.r < headerAbsRow;
      const overlapsTable = merge.e.c >= (data.startCol || 0) && merge.s.c <= (data.startCol || 0) + data.headers.length - 1;
      if (!isAboveHeader && !overlapsTable) return;
      addSection(sections, seen, {
        kind: "group",
        label,
        meta: `Wiersz ${merge.s.r + 1} • kolumny ${formatColRange(merge.s.c, merge.e.c)}`,
        tone: isAboveHeader ? "" : "info",
        action: overlapsTable ? "scroll-col" : "set-header",
        colIndex: Math.max(0, merge.s.c - (data.startCol || 0)),
        headerRow: merge.s.r + 1,
      });
    });

  baseRows
    .filter((row) => row && row.isSubheader)
    .slice(0, 8)
    .forEach((row) => {
      const firstText = row.values.find((value) => typeof value === "string" && cleanSectionLabel(value));
      if (!firstText) return;
      addSection(sections, seen, {
        kind: "subheader",
        label: cleanSectionLabel(firstText),
        meta: `Wiersz danych ${row.rowIndex0 + 1}`,
        tone: "",
        action: "scroll-row",
        rowIndex0: row.rowIndex0,
      });
    });

  return sections.slice(0, 14);
}

function renderSections() {
  if (!sectionNavigatorEl) return;
  sectionNavigatorEl.replaceChildren();
  if (!currentSections.length) {
    sectionNavigatorEl.appendChild(createEmptyInsight("Wczytaj arkusz, aby wykryc sekcje i bloki layoutu."));
    return;
  }

  currentSections.forEach((section, index) => {
    const item = document.createElement("div");
    item.className = `section-nav-item${section.tone ? ` ${section.tone}` : ""}`;

    const top = document.createElement("div");
    top.className = "section-nav-top";

    const title = document.createElement("div");
    title.className = "section-nav-title";
    title.textContent = section.label;

    const kind = document.createElement("div");
    kind.className = "section-nav-kind";
    kind.textContent = inferSectionKindLabel(section.kind);

    top.appendChild(title);
    top.appendChild(kind);

    const meta = document.createElement("div");
    meta.className = "section-nav-meta";
    meta.textContent = section.meta || "Sekcja arkusza";

    const actions = document.createElement("div");
    actions.className = "section-nav-actions";

    const primary = document.createElement("button");
    primary.className = "btn ghost btn-sm";
    primary.type = "button";
    primary.dataset.sectionIndex = String(index);
    primary.dataset.sectionAction = section.action || "scroll-top";
    primary.textContent = section.action === "set-header" ? "Ustaw nagłówek" : "Skocz";
    actions.appendChild(primary);

    item.appendChild(top);
    item.appendChild(meta);
    item.appendChild(actions);
    sectionNavigatorEl.appendChild(item);
  });
}

function renderSheetInspectorSummary() {
  if (!sheetInspectorSummaryEl) return;
  sheetInspectorSummaryEl.replaceChildren();

  if (!currentHeaders.length || !baseRows.length) {
    sheetInspectorSummaryEl.appendChild(createEmptyInsight("Wczytaj arkusz, aby zobaczyc szybkie podsumowanie struktury i najwazniejszych sygnalow."));
    return;
  }

  const blockCount = currentRepeatingBlocks.reduce((sum, group) => sum + (Array.isArray(group.blocks) ? group.blocks.length : 0), 0);
  const flaggedProfiles = currentColumnProfiles.filter((profile) => Array.isArray(profile.flags) && profile.flags.length).length;
  const chips = [
    { label: "Kolumny", value: String(currentHeaders.length) },
    { label: "Sekcje", value: String(currentSections.length), tone: currentSections.length ? "" : "info" },
    { label: "Bloki", value: String(blockCount), tone: blockCount ? "info" : "" },
    { label: "Kolumny z flagami", value: String(flaggedProfiles), tone: flaggedProfiles ? "warning" : "" },
  ];

  const topProfile = currentColumnProfiles[0];
  if (topProfile) {
    chips.push({
      label: "Top sygnal",
      value: topProfile.flags.length ? `${topProfile.header} • ${topProfile.flags[0]}` : `${topProfile.header} • ${topProfile.type}`,
      tone: topProfile.flags.length ? "warning" : "info",
      wide: true,
    });
  }

  chips.forEach((chip) => {
    const item = document.createElement("div");
    item.className = `sheet-inspector-chip${chip.tone ? ` ${chip.tone}` : ""}${chip.wide ? " wide" : ""}`;

    const label = document.createElement("div");
    label.className = "sheet-inspector-chip-label";
    label.textContent = chip.label;

    const value = document.createElement("div");
    value.className = "sheet-inspector-chip-value";
    value.textContent = chip.value;

    item.appendChild(label);
    item.appendChild(value);
    sheetInspectorSummaryEl.appendChild(item);
  });

  const actions = document.createElement("div");
  actions.className = "sheet-inspector-actions";

  const suggestedHeader = currentSections.find((section) => section.action === "set-header" && section.headerRow && section.headerRow !== currentHeaderRow);
  if (suggestedHeader) {
    const btn = document.createElement("button");
    btn.className = "btn ghost btn-sm";
    btn.type = "button";
    btn.dataset.inspectorAction = "set-header";
    btn.dataset.inspectorHeaderRow = String(suggestedHeader.headerRow);
    btn.textContent = `Ustaw naglowek: ${suggestedHeader.headerRow}`;
    actions.appendChild(btn);
  }

  if (canUseLongView()) {
    const btn = document.createElement("button");
    btn.className = "btn ghost btn-sm";
    btn.type = "button";
    btn.dataset.inspectorAction = "toggle-long";
    btn.textContent = tableViewMode === "long" ? "Wroc do widoku klasycznego" : "Przelacz na Wide-to-Long";
    actions.appendChild(btn);
  }

  if (topProfile) {
    const btn = document.createElement("button");
    btn.className = "btn ghost btn-sm";
    btn.type = "button";
    btn.dataset.inspectorAction = "focus-col";
    btn.dataset.profileColIndex = String(topProfile.colIdx);
    btn.textContent = `Skocz do kolumny: ${topProfile.header}`;
    actions.appendChild(btn);
  }

  if (actions.childNodes.length) {
    sheetInspectorSummaryEl.appendChild(actions);
  }
}

function focusSection(section) {
  if (!section) return;
  if (section.action === "set-header" && section.headerRow) {
    if (autoHeaderRowEl) autoHeaderRowEl.checked = false;
    headerRowEl.value = String(section.headerRow);
    toast(`Ustawiono wiersz nagłówka ${section.headerRow}`, "info");
    loadBtn.click();
    return;
  }

  if (section.action === "scroll-row" && Number.isFinite(section.rowIndex0)) {
    const rowEl = tbodyEl.querySelector(`tr[data-row-index="${section.rowIndex0}"]`);
    if (rowEl) {
      rowEl.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" });
      return;
    }
    toast("Ta sekcja nie miesci sie w aktualnym limicie wierszy", "info");
    return;
  }

  if (section.action === "scroll-col" && Number.isFinite(section.colIndex)) {
    const cells = theadEl.querySelectorAll(".guide-row .guide-cell");
    const cell = cells[section.colIndex];
    if (cell && tableWrapEl) {
      const targetLeft = Math.max(0, cell.offsetLeft - 64);
      tableWrapEl.scrollTo({ left: targetLeft, behavior: "smooth" });
      syncHorizontalScrollbar();
      return;
    }
  }

  if (tableWrapEl) {
    tableWrapEl.scrollTo({ top: 0, left: 0, behavior: "smooth" });
    syncHorizontalScrollbar();
  }
}

function parseRepeatedHeader(header) {
  const raw = cleanSectionLabel(header);
  if (!raw) return null;
  const match = raw.match(/^(.*?)(\d+)$/);
  if (!match) return { base: raw, order: 1 };
  const base = cleanSectionLabel(match[1]);
  const order = Number(match[2]);
  if (!base || !Number.isFinite(order)) return { base: raw, order: 1 };
  return { base, order };
}

function normalizeAnalysisKey(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function pluralizeDays(days) {
  const n = Math.abs(days);
  if (n === 1) return "dzien";
  const mod10 = n % 10;
  const mod100 = n % 100;
  if (mod10 >= 2 && mod10 <= 4 && !(mod100 >= 12 && mod100 <= 14)) return "dni";
  return "dni";
}

function formatDurationDays(days) {
  if (!Number.isFinite(days)) return "brak";
  const rounded = Math.max(0, Math.round(days));
  const months = Math.floor(rounded / 30);
  const restDays = rounded % 30;
  const parts = [];
  if (months > 0) parts.push(`${months} mies.`);
  if (restDays > 0 || !parts.length) parts.push(`${restDays} ${pluralizeDays(restDays)}`);
  return parts.join(" ");
}

function pluralizeEntityLabel(label) {
  if (label === "Osoba") return "Osoby";
  if (label === "Pracownik") return "Pracownicy";
  if (label === "Wlasciciel") return "Wlasciciele";
  return `${label}y`;
}

function computeMedian(values) {
  const nums = values.filter((value) => Number.isFinite(value)).slice().sort((a, b) => a - b);
  if (!nums.length) return null;
  const mid = Math.floor(nums.length / 2);
  if (nums.length % 2 === 1) return nums[mid];
  return (nums[mid - 1] + nums[mid]) / 2;
}

function diffDays(start, end) {
  if (!(start instanceof Date) || !(end instanceof Date)) return null;
  const a = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const b = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  const ms = b.getTime() - a.getTime();
  if (!Number.isFinite(ms)) return null;
  const days = Math.round(ms / 86400000);
  return days >= 0 ? days : null;
}

function parseDurationDaysFlexible(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value >= 0 ? value : null;
  }
  if (typeof value !== "string") return null;
  const text = normalizeAnalysisKey(value);
  if (!text) return null;
  const monthMatch = text.match(/(\d+(?:[.,]\d+)?)\s*(m|mies|miesiac|miesiace|miesiecy|month|months)\b/);
  const dayMatch = text.match(/(\d+(?:[.,]\d+)?)\s*(d|dzien|dni|day|days)\b/);
  if (!monthMatch && !dayMatch) {
    if (/^\d+(?:[.,]\d+)?$/.test(text)) {
      const numeric = Number(text.replace(",", "."));
      return Number.isFinite(numeric) && numeric >= 0 ? numeric : null;
    }
    return null;
  }
  const months = monthMatch ? Number(monthMatch[1].replace(",", ".")) : 0;
  const days = dayMatch ? Number(dayMatch[1].replace(",", ".")) : 0;
  const total = (months * 30) + days;
  return Number.isFinite(total) && total >= 0 ? total : null;
}

function findAnalysisColumnIndex(candidates, matchers) {
  for (const matcher of matchers) {
    const hit = candidates.find((candidate) => matcher(candidate.norm, candidate.base));
    if (hit) return hit.idx;
  }
  return -1;
}

function collectDurationBlockStats(group) {
  const firstBlock = group && Array.isArray(group.blocks) ? group.blocks[0] : null;
  if (!firstBlock || !Array.isArray(firstBlock.headers) || !firstBlock.headers.length) return [];

  const stats = firstBlock.headers.map((header, idx) => ({
    idx,
    header,
    nonEmptyCount: 0,
    dateCount: 0,
    durationCount: 0,
    textCount: 0,
    uniqueText: new Set(),
  }));

  const rowSample = viewRows.slice(0, 400);
  rowSample.forEach((row) => {
    group.blocks.forEach((block) => {
      stats.forEach((entry) => {
        const absIdx = block.startIndex + entry.idx;
        const raw = row.values[absIdx] ?? getDisplayValue(row, absIdx);
        const text = String(raw ?? "").trim();
        if (!text) return;
        entry.nonEmptyCount += 1;

        const asDate = parseDateFlexible(raw);
        if (asDate instanceof Date) {
          entry.dateCount += 1;
          return;
        }

        const asDuration = parseDurationDaysFlexible(raw);
        if (asDuration !== null) {
          entry.durationCount += 1;
          return;
        }

        entry.textCount += 1;
        entry.uniqueText.add(normalizeAnalysisKey(text));
      });
    });
  });

  return stats.map((entry) => ({
    ...entry,
    uniqueTextCount: entry.uniqueText.size,
  }));
}

function inferDurationAnalysisConfigFromData(group) {
  const stats = collectDurationBlockStats(group);
  if (!stats.length) return null;

  const entityCandidate = stats
    .filter((entry) => entry.textCount > 0)
    .sort((a, b) => {
      const scoreA = (a.textCount * 5) + Math.min(a.uniqueTextCount, 25);
      const scoreB = (b.textCount * 5) + Math.min(b.uniqueTextCount, 25);
      return scoreB - scoreA || a.idx - b.idx;
    })[0];

  const dateCandidates = stats
    .filter((entry) => entry.dateCount > 0)
    .sort((a, b) => b.dateCount - a.dateCount || a.idx - b.idx);

  const durationCandidate = stats
    .filter((entry) => entry.durationCount > 0)
    .sort((a, b) => b.durationCount - a.durationCount || a.idx - b.idx)[0];

  if (!entityCandidate) return null;
  if (!dateCandidates.length && !durationCandidate) return null;

  const orderedDateCandidates = dateCandidates.slice().sort((a, b) => a.idx - b.idx);
  const startCandidate = orderedDateCandidates[0] || null;
  const endCandidate = orderedDateCandidates[1] || null;

  return {
    entityIdx: entityCandidate.idx,
    startIdx: startCandidate ? startCandidate.idx : -1,
    endIdx: endCandidate ? endCandidate.idx : -1,
    durationIdx: durationCandidate ? durationCandidate.idx : -1,
    entityLabel: "Osoba",
    entityHeader: entityCandidate.header || "Osoba",
    inferred: true,
  };
}

function detectDurationAnalysisConfig(group) {
  const firstBlock = group && Array.isArray(group.blocks) ? group.blocks[0] : null;
  if (!firstBlock || !Array.isArray(firstBlock.headers) || !firstBlock.headers.length) return null;

  const candidates = firstBlock.headers.map((header, idx) => {
    const base = parseRepeatedHeader(header)?.base || cleanSectionLabel(header) || String(header || "");
    return {
      idx,
      header,
      base,
      norm: normalizeAnalysisKey(base),
    };
  });

  const entityIdx = findAnalysisColumnIndex(candidates, [
    (norm) => /\b(imie|nazwisko|osoba|pracownik|opiekun|wlasciciel|owner|assignee|user|agent|operator)\b/.test(norm),
    (norm) => norm.includes("imie") || norm.includes("nazwisk"),
  ]);
  const startIdx = findAnalysisColumnIndex(candidates, [
    (norm) => norm === "od" || norm === "data od",
    (norm) => /\b(start|from|poczatek|rozpoczecie|rozpoczecia)\b/.test(norm),
  ]);
  const endIdx = findAnalysisColumnIndex(candidates, [
    (norm) => norm === "do" || norm === "data do",
    (norm) => /\b(koniec|zakonczenie|end|to|until)\b/.test(norm),
  ]);
  const durationIdx = findAnalysisColumnIndex(candidates, [
    (norm) => norm.includes("dlugosc") || norm.includes("czas"),
    (norm) => /\b(duration|age|days)\b/.test(norm),
  ]);

  const inferred = inferDurationAnalysisConfigFromData(group);

  const resolvedEntityIdx = entityIdx >= 0 ? entityIdx : (inferred?.entityIdx ?? -1);
  const resolvedStartIdx = startIdx >= 0 ? startIdx : (inferred?.startIdx ?? -1);
  const resolvedEndIdx = endIdx >= 0 ? endIdx : (inferred?.endIdx ?? -1);
  const resolvedDurationIdx = durationIdx >= 0 ? durationIdx : (inferred?.durationIdx ?? -1);

  if (resolvedEntityIdx < 0 || (resolvedStartIdx < 0 && resolvedDurationIdx < 0)) return null;

  const entityBase = candidates[resolvedEntityIdx]?.base || inferred?.entityHeader || "Wartosc";
  const normEntity = normalizeAnalysisKey(entityBase);
  let entityLabel = "Wartosc";
  if (normEntity.includes("imie") || normEntity.includes("nazwisk") || normEntity.includes("osoba")) entityLabel = "Osoba";
  else if (normEntity.includes("pracownik")) entityLabel = "Pracownik";
  else if (normEntity.includes("owner") || normEntity.includes("wlasciciel")) entityLabel = "Wlasciciel";
  else if (inferred?.inferred) entityLabel = "Osoba";
  else if (entityBase) entityLabel = entityBase;

  return {
    entityIdx: resolvedEntityIdx,
    startIdx: resolvedStartIdx,
    endIdx: resolvedEndIdx,
    durationIdx: resolvedDurationIdx,
    entityLabel,
    entityHeader: entityBase,
    inferred: !!(inferred && (entityIdx < 0 || startIdx < 0 || endIdx < 0 || durationIdx < 0)),
  };
}

function buildDurationAnalysisFromRows(group, rows, meta = {}) {
  const config = detectDurationAnalysisConfig(group);
  if (!config) {
    return { status: "no-config", group, ...meta };
  }

  const today = new Date();
  const records = [];
  const aggregate = new Map();

  rows.forEach((row) => {
    group.blocks.forEach((block, blockIndex) => {
      const entityCol = block.startIndex + config.entityIdx;
      const startCol = config.startIdx >= 0 ? block.startIndex + config.startIdx : -1;
      const endCol = config.endIdx >= 0 ? block.startIndex + config.endIdx : -1;
      const durationCol = config.durationIdx >= 0 ? block.startIndex + config.durationIdx : -1;

      const entityValue = String(row.values[entityCol] ?? "").trim();
      if (!entityValue) return;

      const startDate = startCol >= 0 ? parseDateFlexible(row.values[startCol] ?? getDisplayValue(row, startCol)) : null;
      const endDate = endCol >= 0 ? parseDateFlexible(row.values[endCol] ?? getDisplayValue(row, endCol)) : null;
      let durationDays = null;
      let isOpen = false;

      if (startDate instanceof Date) {
        if (endDate instanceof Date) {
          durationDays = diffDays(startDate, endDate);
        } else {
          durationDays = diffDays(startDate, today);
          isOpen = durationDays !== null;
        }
      }

      if (durationDays === null && durationCol >= 0) {
        durationDays = parseDurationDaysFlexible(row.values[durationCol] ?? getDisplayValue(row, durationCol));
      }

      records.push({
        entity: entityValue,
        durationDays,
        isOpen,
        isClosed: durationDays !== null && !isOpen,
        blockLabel: block.label,
        blockIndex: blockIndex + 1,
        rowIndex0: row.rowIndex0,
      });
    });
  });

  const filteredRecords = records.filter((record) => {
    if (!Number.isFinite(record.durationDays)) return false;
    if (durationAnalysisState.statusFilter === "open") return record.isOpen;
    if (durationAnalysisState.statusFilter === "closed") return !record.isOpen;
    return true;
  });

  filteredRecords.forEach((record) => {
    const key = normalizeAnalysisKey(record.entity);
    const entry = aggregate.get(key) || {
      entity: record.entity,
      durations: [],
      openCount: 0,
      minDays: null,
      maxDays: null,
      blocks: new Set(),
      rowIndexes: new Set(),
    };
    entry.durations.push(record.durationDays);
    if (record.isOpen) entry.openCount += 1;
    entry.minDays = entry.minDays === null ? record.durationDays : Math.min(entry.minDays, record.durationDays);
    entry.maxDays = entry.maxDays === null ? record.durationDays : Math.max(entry.maxDays, record.durationDays);
    entry.blocks.add(record.blockLabel);
    entry.rowIndexes.add(record.rowIndex0);
    aggregate.set(key, entry);
  });

  const entries = Array.from(aggregate.values())
    .map((entry) => ({
      entity: entry.entity,
      averageDays: entry.durations.length ? entry.durations.reduce((sum, value) => sum + value, 0) / entry.durations.length : null,
      medianDays: computeMedian(entry.durations),
      count: entry.durations.length,
      openCount: entry.openCount,
      minDays: entry.minDays,
      maxDays: entry.maxDays,
      blockCount: entry.blocks.size,
      rowCount: entry.rowIndexes.size,
    }))
    .sort((a, b) => {
      const metricMap = {
        avg: "averageDays",
        median: "medianDays",
        count: "count",
        min: "minDays",
        max: "maxDays",
      };
      const metric = metricMap[durationAnalysisState.sortMetric] || "averageDays";
      const left = Number(a[metric] || 0);
      const right = Number(b[metric] || 0);
      const diff = right - left;
      if (Math.abs(diff) > 0.001) return diff;
      const countDiff = b.count - a.count;
      if (countDiff) return countDiff;
      return a.entity.localeCompare(b.entity, "pl");
    });

  if (!entries.length) {
    return { status: "no-records", config, group, records, filteredRecords, ...meta };
  }

  const totalDurationRecords = filteredRecords.length;
  const totalOpen = filteredRecords.filter((record) => record.isOpen).length;
  const totalClosed = filteredRecords.filter((record) => !record.isOpen).length;
  const allDurations = filteredRecords.map((record) => record.durationDays).filter((value) => Number.isFinite(value));
  const totalDays = allDurations.reduce((sum, value) => sum + value, 0);

  return {
    status: "ok",
    config,
    group,
    entries,
    records,
    filteredRecords,
    ...meta,
    summary: {
      uniqueEntities: entries.length,
      totalDurationRecords,
      totalOpen,
      totalClosed,
      averageDays: totalDurationRecords ? totalDays / totalDurationRecords : null,
      medianDays: computeMedian(allDurations),
      minDays: allDurations.length ? Math.min(...allDurations) : null,
      maxDays: allDurations.length ? Math.max(...allDurations) : null,
      visibleRows: rows.length,
      sourceRows: rows.length,
    },
  };
}

function tryBuildDurationAnalysisFromAlternateHeaders() {
  if (!workbook || !currentSheetName) return null;
  const sheet = workbook.Sheets[currentSheetName];
  if (!sheet) return null;

  const candidateRows = [];
  const seen = new Set();
  const minHeader = 1;
  const maxHeader = Math.max(minHeader, currentHeaderRow + 4);

  for (let row = Math.max(minHeader, currentHeaderRow - 3); row <= maxHeader; row++) {
    if (row === currentHeaderRow) continue;
    if (seen.has(row)) continue;
    seen.add(row);
    candidateRows.push(row);
  }

  let best = null;

  candidateRows.forEach((headerRow) => {
    try {
      const data = buildRows(sheet, headerRow, workbook);
      const groups = detectRepeatingBlocks(sheet, headerRow, data);
      const group = Array.isArray(groups) && groups.length ? groups[0] : null;
      if (!group || !Array.isArray(group.blocks) || group.blocks.length < 2) return;

      const shadowRows = markSubheaderRows(data.rows.slice());
      const result = buildDurationAnalysisFromRows(group, shadowRows, {
        helperHeaderRow: headerRow,
        helperMode: true,
      });
      if (!result || result.status !== "ok") return;

      const score = (result.summary.totalDurationRecords * 10) + result.summary.uniqueEntities;
      if (!best || score > best.score) {
        best = { ...result, score };
      }
    } catch {
      // Ignore helper header candidates that fail to parse well.
    }
  });

  return best;
}

function buildDurationAnalysis() {
  const group = getActiveRepeatingGroup();
  if (!group || !Array.isArray(group.blocks) || group.blocks.length < 2) {
    const fallback = tryBuildDurationAnalysisFromAlternateHeaders();
    return fallback || { status: "no-group" };
  }

  const currentResult = buildDurationAnalysisFromRows(group, viewRows, {
    helperHeaderRow: currentHeaderRow,
    helperMode: false,
  });

  if (currentResult.status === "ok") {
    return currentResult;
  }

  const fallback = tryBuildDurationAnalysisFromAlternateHeaders();
  return fallback || currentResult;
}

function renderDurationAnalysis() {
  if (!durationAnalysisSummaryEl || !durationAnalysisListEl) return;
  durationAnalysisSummaryEl.replaceChildren();
  durationAnalysisListEl.replaceChildren();

  const analysis = buildDurationAnalysis();

  if (analysis.status === "no-group") {
    durationAnalysisSummaryEl.appendChild(createEmptyInsight("Wykryj najpierw powtarzalne bloki kolumn. Ten panel najlepiej dziala na arkuszach z cyklami albo seriami podobnych pol."));
    return;
  }

  if (analysis.status === "no-config") {
    durationAnalysisSummaryEl.appendChild(createEmptyInsight("Wykryto bloki, ale nie udalo sie znalezc pary typu osoba + od/do albo osoba + dlugosc. Jesli naglowek jest nietypowy, modul probuje tez zgadywac po danych, ale tu to wciaz za malo."));
    return;
  }

  if (analysis.status === "no-records") {
    durationAnalysisSummaryEl.appendChild(createEmptyInsight("Bloki zostaly rozpoznane, ale w aktualnym widoku nie ma rekordow z pelnymi danymi czasu dla tej samej wartosci."));
    return;
  }

  const summaryGrid = document.createElement("div");
  summaryGrid.className = "sheet-inspector-summary";
  [
    { label: pluralizeEntityLabel(analysis.config.entityLabel), value: String(analysis.summary.uniqueEntities) },
    { label: "Rekordy czasu", value: String(analysis.summary.totalDurationRecords) },
    { label: "Sredni czas", value: formatDurationDays(analysis.summary.averageDays) },
    { label: "Mediana", value: formatDurationDays(analysis.summary.medianDays) },
    { label: "Min", value: formatDurationDays(analysis.summary.minDays) },
    { label: "Max", value: formatDurationDays(analysis.summary.maxDays) },
    { label: "W toku", value: String(analysis.summary.totalOpen), tone: analysis.summary.totalOpen ? "info" : "" },
    { label: "Zamkniete", value: String(analysis.summary.totalClosed) },
  ].forEach((item) => {
    const chip = document.createElement("div");
    chip.className = `sheet-inspector-chip${item.tone ? ` ${item.tone}` : ""}`;

    const label = document.createElement("div");
    label.className = "sheet-inspector-chip-label";
    label.textContent = item.label;

    const value = document.createElement("div");
    value.className = "sheet-inspector-chip-value";
    value.textContent = item.value;

    chip.appendChild(label);
    chip.appendChild(value);
    summaryGrid.appendChild(chip);
  });
  durationAnalysisSummaryEl.appendChild(summaryGrid);

  const note = document.createElement("div");
  note.className = "duration-analysis-note";
  const filtered = analysis.summary.visibleRows !== analysis.summary.sourceRows;
  note.textContent = filtered
    ? `Analiza dotyczy aktualnie przefiltrowanego widoku (${analysis.summary.visibleRows} z ${analysis.summary.sourceRows} wierszy). Otwarte rekordy bez daty "do" sa liczone do dzisiaj.`
    : 'Analiza dotyczy calego aktualnego widoku arkusza. Otwarte rekordy bez daty "do" sa liczone do dzisiaj.';
  if (analysis.config.inferred) {
    note.textContent += " Uklad kolumn zostal czesciowo odgadniety na podstawie danych, bo naglowek nie byl idealny.";
  }
  if (analysis.helperMode && Number.isFinite(analysis.helperHeaderRow) && analysis.helperHeaderRow !== currentHeaderRow) {
    note.textContent += ` Do tej analizy uzyto pomocniczo wiersza naglowka ${analysis.helperHeaderRow}, bo lepiej pasowal niz aktualnie wybrany ${currentHeaderRow}.`;
  }
  durationAnalysisSummaryEl.appendChild(note);

  const controls = document.createElement("div");
  controls.className = "duration-analysis-controls";

  const statusField = document.createElement("label");
  statusField.className = "field";
  statusField.append("Status");
  const statusSelect = document.createElement("select");
  statusSelect.dataset.durationControl = "status";
  [
    { value: "all", label: "Wszystkie" },
    { value: "closed", label: "Tylko zamkniete" },
    { value: "open", label: "Tylko otwarte" },
  ].forEach((item) => {
    const option = document.createElement("option");
    option.value = item.value;
    option.textContent = item.label;
    statusSelect.appendChild(option);
  });
  statusSelect.value = durationAnalysisState.statusFilter;
  statusField.appendChild(statusSelect);

  const sortField = document.createElement("label");
  sortField.className = "field";
  sortField.append("Sortuj po");
  const sortSelect = document.createElement("select");
  sortSelect.dataset.durationControl = "sort";
  [
    { value: "avg", label: "Sredniej" },
    { value: "median", label: "Medianie" },
    { value: "count", label: "Liczbie rekordow" },
    { value: "max", label: "Maksimum" },
    { value: "min", label: "Minimum" },
  ].forEach((item) => {
    const option = document.createElement("option");
    option.value = item.value;
    option.textContent = item.label;
    sortSelect.appendChild(option);
  });
  sortSelect.value = durationAnalysisState.sortMetric;
  sortField.appendChild(sortSelect);

  const countField = document.createElement("label");
  countField.className = "field";
  countField.append("Pokaz rekordow");
  const countSelect = document.createElement("select");
  countSelect.dataset.durationControl = "count";
  ["14", "24", "40", "80", "999"].forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value === "999" ? "Wszystkie" : value;
    countSelect.appendChild(option);
  });
  countSelect.value = String(durationAnalysisState.showCount);
  countField.appendChild(countSelect);

  controls.appendChild(statusField);
  controls.appendChild(sortField);
  controls.appendChild(countField);
  durationAnalysisSummaryEl.appendChild(controls);

  const actions = document.createElement("div");
  actions.className = "section-nav-actions";
  if (canUseLongView()) {
    const toggleBtn = document.createElement("button");
    toggleBtn.className = "btn ghost btn-sm";
    toggleBtn.type = "button";
    toggleBtn.dataset.durationAction = "toggle-long";
    toggleBtn.textContent = tableViewMode === "long" ? "Widok klasyczny" : "Wide-to-Long";
    actions.appendChild(toggleBtn);
  }
  if (filtered) {
    const resetBtn = document.createElement("button");
    resetBtn.className = "btn ghost btn-sm";
    resetBtn.type = "button";
    resetBtn.dataset.durationAction = "reset-filters";
    resetBtn.textContent = "Pokaz calosc";
    actions.appendChild(resetBtn);
  }
  durationAnalysisSummaryEl.appendChild(actions);

  const listNote = document.createElement("div");
  listNote.className = "duration-analysis-note";
  const visibleCount = Math.min(durationAnalysisState.showCount, analysis.entries.length);
  listNote.textContent = analysis.entries.length > visibleCount
    ? `Pokazano ${visibleCount} z ${analysis.entries.length} wynikow.`
    : `Pokazano wszystkie wyniki: ${analysis.entries.length}.`;
  durationAnalysisListEl.appendChild(listNote);

  analysis.entries.slice(0, durationAnalysisState.showCount).forEach((entry, index) => {
    const item = document.createElement("div");
    item.className = "duration-person-item";

    const top = document.createElement("div");
    top.className = "duration-person-top";

    const titleWrap = document.createElement("div");
    titleWrap.className = "duration-person-title-wrap";

    const rank = document.createElement("div");
    rank.className = "duration-person-rank";
    rank.textContent = String(index + 1);

    const title = document.createElement("div");
    title.className = "duration-person-title";
    title.textContent = entry.entity;

    const value = document.createElement("div");
    value.className = "duration-person-value";
    value.textContent = formatDurationDays(entry.averageDays);

    titleWrap.appendChild(rank);
    titleWrap.appendChild(title);
    top.appendChild(titleWrap);
    top.appendChild(value);

    const meta = document.createElement("div");
    meta.className = "duration-person-meta";
    const avgDaysText = entry.averageDays !== null ? `${Math.round(entry.averageDays * 10) / 10} dni` : "brak";
    const medianDaysText = entry.medianDays !== null ? `${Math.round(entry.medianDays * 10) / 10} dni` : "brak";
    meta.textContent = `Srednio ${avgDaysText} • mediana ${medianDaysText} • rekordy ${entry.count} • w toku ${entry.openCount} • zakres ${formatDurationDays(entry.minDays)} -> ${formatDurationDays(entry.maxDays)}`;

    const actionsRow = document.createElement("div");
    actionsRow.className = "section-nav-actions";

    const filterBtn = document.createElement("button");
    filterBtn.className = "btn ghost btn-sm";
    filterBtn.type = "button";
    filterBtn.dataset.durationAction = "filter-entity";
    filterBtn.dataset.durationEntity = entry.entity;
    filterBtn.textContent = "Pokaz w tabeli";
    actionsRow.appendChild(filterBtn);

    item.appendChild(top);
    item.appendChild(meta);
    item.appendChild(actionsRow);
    durationAnalysisListEl.appendChild(item);
  });
}

function inferAggregationValueKind(header, profile) {
  const norm = normalizeAnalysisKey(header);
  if (profile?.measureType === "date_range") return "duration";
  if (profile?.durationCount > 0 || norm.includes("dlugosc") || norm.includes("czas")) return "duration";
  if (profile?.numericCount > 0) return "number";
  return "text";
}

function formatAggregationMetricValue(value, kind = "number") {
  if (!Number.isFinite(value)) return "brak";
  if (kind === "duration") return formatDurationDays(value);
  const rounded = Math.round(value * 100) / 100;
  return String(rounded).replace(".", ",");
}

function collectAggregationProfiles(model) {
  if (!model || !Array.isArray(model.headers) || !Array.isArray(model.rows)) return [];
  return model.headers.map((header, idx) => {
    const profile = {
      header,
      idx,
      nonEmptyCount: 0,
      numericCount: 0,
      durationCount: 0,
      dateCount: 0,
      textCount: 0,
      uniqueValues: new Set(),
    };

    model.rows.forEach((row) => {
      const raw = row.values?.[idx];
      const display = getDisplayValue(row, idx);
      const text = String(display ?? raw ?? "").trim();
      if (!text) return;
      profile.nonEmptyCount += 1;
      profile.uniqueValues.add(normalizeAnalysisKey(text));

      if (typeof raw === "number" && Number.isFinite(raw)) {
        profile.numericCount += 1;
        return;
      }

      const duration = parseDurationDaysFlexible(raw ?? display);
      if (duration !== null) {
        profile.durationCount += 1;
        return;
      }

      const asDate = parseDateFlexible(raw ?? display);
      if (asDate instanceof Date) {
        profile.dateCount += 1;
        return;
      }

      profile.textCount += 1;
    });

    profile.uniqueCount = profile.uniqueValues.size;
    return profile;
  });
}

function detectAggregationDateRangeCandidates(model, profiles) {
  const candidates = [];
  const startRegex = /\b(od|start|data od|from|poczatek|rozpoczecie)\b/;
  const endRegex = /\b(do|end|data do|to|until|koniec|zakonczenie)\b/;

  profiles.forEach((profile, idx) => {
    if (profile.dateCount <= 0) return;
    const base = parseRepeatedHeader(model.headers[idx])?.base || cleanSectionLabel(model.headers[idx]) || model.headers[idx];
    const norm = normalizeAnalysisKey(base);
    if (!startRegex.test(norm)) return;

    let endIdx = -1;
    for (let next = idx + 1; next < profiles.length; next++) {
      if (profiles[next].dateCount <= 0) continue;
      const nextBase = parseRepeatedHeader(model.headers[next])?.base || cleanSectionLabel(model.headers[next]) || model.headers[next];
      const nextNorm = normalizeAnalysisKey(nextBase);
      if (endRegex.test(nextNorm)) {
        endIdx = next;
        break;
      }
      if (next > idx + 2) break;
    }
    if (endIdx < 0) return;

    candidates.push({
      key: `date_range:${idx}:${endIdx}`,
      label: `${model.headers[idx]} -> ${model.headers[endIdx]}`,
      kind: "duration",
      measureType: "date_range",
      startIdx: idx,
      endIdx,
      getValue: (row) => {
        const start = parseDateFlexible(row.values?.[idx] ?? getDisplayValue(row, idx));
        const end = parseDateFlexible(row.values?.[endIdx] ?? getDisplayValue(row, endIdx));
        if (!(start instanceof Date) || !(end instanceof Date)) return null;
        return diffDays(start, end);
      },
    });
  });

  return candidates;
}

function detectAggregationMeasureCandidates(model, profiles) {
  const candidates = [{
    key: "count_rows",
    label: "Liczba wierszy",
    kind: "count",
    measureType: "count_rows",
    getValue: () => 1,
  }];

  detectAggregationDateRangeCandidates(model, profiles).forEach((candidate) => {
    candidates.push(candidate);
  });

  profiles.forEach((profile) => {
    if (profile.nonEmptyCount <= 0) return;
    if (profile.numericCount <= 0 && profile.durationCount <= 0) return;
    const kind = inferAggregationValueKind(profile.header, profile);
    candidates.push({
      key: `column:${profile.idx}`,
      label: profile.header,
      kind,
      measureType: "column",
      colIdx: profile.idx,
      getValue: (row) => {
        const raw = row.values?.[profile.idx];
        if (typeof raw === "number" && Number.isFinite(raw)) return raw;
        return parseDurationDaysFlexible(raw ?? getDisplayValue(row, profile.idx));
      },
      getRawText: (row) => {
        const raw = row.values?.[profile.idx];
        return raw != null ? String(raw).trim() : "";
      },
    });
  });

  return candidates;
}

function resolveAggregationGroupOptions(profiles) {
  return profiles
    .filter((profile) => profile.nonEmptyCount > 0)
    .sort((a, b) => {
      const aTextScore = a.textCount > 0 ? 1 : 0;
      const bTextScore = b.textCount > 0 ? 1 : 0;
      if (aTextScore !== bTextScore) return bTextScore - aTextScore;
      return a.idx - b.idx;
    })
    .map((profile) => ({
      value: profile.header,
      label: profile.header,
      idx: profile.idx,
    }));
}

function chooseDefaultAggregationGroup(groupOptions) {
  if (!groupOptions.length) return "";
  const preferred = groupOptions.find((option) => /\b(imie|nazwisko|osoba|pracownik|owner|assignee|blok)\b/.test(normalizeAnalysisKey(option.label)));
  return preferred ? preferred.value : groupOptions[0].value;
}

function chooseDefaultAggregationMeasure(measures) {
  if (!measures.length) return "count_rows";
  const dateRange = measures.find((candidate) => candidate.measureType === "date_range");
  if (dateRange) return dateRange.key;
  const duration = measures.find((candidate) => candidate.kind === "duration");
  if (duration) return duration.key;
  const numeric = measures.find((candidate) => candidate.measureType === "column");
  return numeric ? numeric.key : "count_rows";
}

function chooseDefaultAggregationMethod(measure) {
  if (!measure || measure.measureType === "count_rows") return "count";
  return "avg";
}

function getNormalizedAggregationWorkbenchContext() {
  const headerCandidates = getAggregationHeaderCandidateRows();
  let resolvedHeaderRow = currentHeaderRow;
  let context = null;

  if (aggregationWorkbenchState.headerRowChoice === "auto") {
    headerCandidates.forEach((candidateRow) => {
      const candidateContext = collectAggregationContextForHeaderRow(
        candidateRow,
        aggregationWorkbenchState.sourceMode,
        aggregationWorkbenchState.scopeMode
      );
      const score = scoreAggregationContext(candidateContext);
      if (!context || score > context.score) {
        context = { ...candidateContext, score };
        resolvedHeaderRow = candidateRow;
      }
    });
  } else {
    const explicitRow = Number.isFinite(aggregationWorkbenchState.customHeaderRow)
      ? aggregationWorkbenchState.customHeaderRow
      : currentHeaderRow;
    resolvedHeaderRow = explicitRow > 0 ? explicitRow : currentHeaderRow;
    context = collectAggregationContextForHeaderRow(
      resolvedHeaderRow,
      aggregationWorkbenchState.sourceMode,
      aggregationWorkbenchState.scopeMode
    );
  }

  if (!context) {
    context = collectAggregationContextForHeaderRow(
      currentHeaderRow,
      aggregationWorkbenchState.sourceMode,
      aggregationWorkbenchState.scopeMode
    );
    resolvedHeaderRow = currentHeaderRow;
  }

  const { model, profiles, groupOptions, measures, longAvailable } = context;

  const nextGroupBy = groupOptions.some((option) => option.value === aggregationWorkbenchState.groupBy)
    ? aggregationWorkbenchState.groupBy
    : chooseDefaultAggregationGroup(groupOptions);
  const nextMeasure = measures.some((candidate) => candidate.key === aggregationWorkbenchState.measure)
    ? aggregationWorkbenchState.measure
    : chooseDefaultAggregationMeasure(measures);
  const measure = measures.find((candidate) => candidate.key === nextMeasure) || measures[0] || null;
  const allowedAggregations = measure?.measureType === "count_rows"
    ? ["count"]
    : ["avg", "median", "min", "max", "sum", "count", "distinct"];
  const nextAggregation = allowedAggregations.includes(aggregationWorkbenchState.aggregation)
    ? aggregationWorkbenchState.aggregation
    : chooseDefaultAggregationMethod(measure);

  aggregationWorkbenchState.groupBy = nextGroupBy;
  aggregationWorkbenchState.measure = nextMeasure;
  aggregationWorkbenchState.aggregation = nextAggregation;
  if (aggregationWorkbenchState.sourceMode === "long" && !longAvailable) {
    aggregationWorkbenchState.sourceMode = "auto";
  }

  return {
    ...context,
    model,
    profiles,
    groupOptions,
    measures,
    measure,
    longAvailable,
    allowedAggregations,
    headerCandidates,
    resolvedHeaderRow,
  };
}

function computeAggregateMetric(values, aggregation) {
  if (aggregation === "distinct") {
    const unique = new Set(values.map((v) => String(v).trim()).filter((v) => v));
    return unique.size;
  }
  const nums = values.filter((value) => Number.isFinite(value));
  if (aggregation === "count") return nums.length;
  if (!nums.length) return null;
  if (aggregation === "sum") return nums.reduce((sum, value) => sum + value, 0);
  if (aggregation === "avg") return nums.reduce((sum, value) => sum + value, 0) / nums.length;
  if (aggregation === "median") return computeMedian(nums);
  if (aggregation === "min") return Math.min(...nums);
  if (aggregation === "max") return Math.max(...nums);
  return null;
}

function buildAggregationWorkbenchResult() {
  const context = getNormalizedAggregationWorkbenchContext();
  const { model, groupOptions, measures, measure } = context;
  if (!model?.headers?.length || !model?.rows?.length) {
    return { status: "empty", ...context };
  }
  if (!groupOptions.length || !measure) {
    return { status: "no-options", ...context };
  }

  const groupIdx = model.headers.indexOf(aggregationWorkbenchState.groupBy);
  if (groupIdx < 0) {
    return { status: "no-options", ...context };
  }

  const isDistinct = aggregationWorkbenchState.aggregation === "distinct";
  const measureFilterMode = aggregationWorkbenchState.measureFilterMode || "all";
  const measureFilterValue = aggregationWorkbenchState.measureFilterValue || "";
  const buckets = new Map();
  model.rows.forEach((row) => {
    if (measureFilterMode !== "all" && measureFilterValue) {
      const rowMeasureText = measure.getRawText
        ? measure.getRawText(row)
        : getDisplayValue(row, measure.colIdx) || "";
      const rowMeasureLower = rowMeasureText.toLowerCase();
      const filterLower = measureFilterValue.toLowerCase();
      if (measureFilterMode === "contains" && !rowMeasureLower.includes(filterLower)) return;
      if (measureFilterMode === "exact" && rowMeasureLower !== filterLower) return;
    }
    const rawGroup = row.values?.[groupIdx];
    const groupLabel = String(getDisplayValue(row, groupIdx) || rawGroup || "(puste)").trim() || "(puste)";
    const key = normalizeAnalysisKey(groupLabel) || "(puste)";
    const entry = buckets.get(key) || {
      label: groupLabel,
      values: [],
      rowIndexes: new Set(),
    };
    if (isDistinct) {
      const rawValue = measure.getRawText
        ? measure.getRawText(row)
        : measure.colIdx != null
          ? getDisplayValue(row, measure.colIdx) || ""
          : "";
      entry.values.push(rawValue);
    } else {
      const measureValue = measure.getValue(row);
      if (measure.measureType === "count_rows") {
        entry.values.push(1);
      } else if (Number.isFinite(measureValue)) {
        entry.values.push(measureValue);
      }
    }
    entry.rowIndexes.add(row.rowIndex0);
    buckets.set(key, entry);
  });

  const rawEntries = Array.from(buckets.values())
    .map((entry) => {
      const count = entry.values.filter((value) => Number.isFinite(value)).length;
      return {
        label: entry.label,
        count,
        average: computeAggregateMetric(entry.values, "avg"),
        median: computeAggregateMetric(entry.values, "median"),
        min: computeAggregateMetric(entry.values, "min"),
        max: computeAggregateMetric(entry.values, "max"),
        sum: computeAggregateMetric(entry.values, "sum"),
        primary: computeAggregateMetric(entry.values, aggregationWorkbenchState.aggregation),
      };
    })
    .filter((entry) => {
      if (aggregationWorkbenchState.aggregation === "distinct") return true;
      return entry.count > 0 || aggregationWorkbenchState.aggregation === "count";
    });

  const totalPrimary = rawEntries.reduce((sum, e) => sum + (e.primary || 0), 0);
  const maxPrimary = rawEntries.length > 0 ? Math.max(...rawEntries.map(e => e.primary || 0)) : 0;

  const entries = rawEntries
    .filter((entry) => {
      if (aggregationWorkbenchState.havingMode === "all") return true;
      const primary = entry.primary || 0;
      const value = aggregationWorkbenchState.havingValue;
      if (aggregationWorkbenchState.havingMode === "above_value") return primary > value;
      if (aggregationWorkbenchState.havingMode === "above_percent") return totalPrimary > 0 && (primary / totalPrimary) * 100 > value;
      if (aggregationWorkbenchState.havingMode === "above_max_percent") return maxPrimary > 0 && (primary / maxPrimary) * 100 > value;
      return true;
    })
    .sort((a, b) => {
      const diff = Number(b.primary || 0) - Number(a.primary || 0);
      if (Math.abs(diff) > 0.001) return diff;
      const countDiff = b.count - a.count;
      if (countDiff) return countDiff;
      return a.label.localeCompare(b.label, "pl");
    });

  if (!entries.length) {
    return { status: "no-results", ...context };
  }

  return {
    status: "ok",
    ...context,
    entries,
    summary: {
      groups: entries.length,
      sourceRows: model.rows.length,
      measuredRows: entries.reduce((sum, entry) => sum + entry.count, 0),
    },
  };
}

function renderAggregationWorkbench() {
  if (!aggregationWorkbenchSummaryEl || !aggregationWorkbenchListEl) return;
  aggregationWorkbenchSummaryEl.replaceChildren();
  aggregationWorkbenchListEl.replaceChildren();

  const result = buildAggregationWorkbenchResult();
  if (result.status === "empty") {
    aggregationWorkbenchSummaryEl.appendChild(createEmptyInsight("Wczytaj arkusz, aby uruchomic agregacje."));
    return;
  }
  if (result.status === "no-options") {
    aggregationWorkbenchSummaryEl.appendChild(createEmptyInsight("Brak sensownych opcji grupowania lub mierzenia dla aktualnego zrodla danych."));
    return;
  }
  if (result.status === "no-results") {
    aggregationWorkbenchSummaryEl.appendChild(createEmptyInsight("Aktualna kombinacja grupowania i mierzenia nie zwrocila zadnych wynikow."));
    return;
  }

  const measure = result.measure;
  const isDistinctMode = aggregationWorkbenchState.aggregation === "distinct";
  const primaryKind = isDistinctMode ? "number" : (measure?.kind === "duration" ? "duration" : "number");
  const aggregationLabels = {
    count: "Liczebnosc",
    avg: "Srednia",
    median: "Mediana",
    min: "Minimum",
    max: "Maksimum",
    sum: "Suma",
    distinct: "Unikalnych",
  };

  const summaryGrid = document.createElement("div");
  summaryGrid.className = "sheet-inspector-summary";
  [
    { label: "Grupy", value: String(result.summary.groups) },
    { label: "Wiersze zrodla", value: String(result.summary.sourceRows) },
    { label: "Zmierzonych rekordow", value: String(result.summary.measuredRows) },
    { label: aggregationLabels[aggregationWorkbenchState.aggregation] || "Wynik", value: formatAggregationMetricValue(result.entries[0]?.primary, primaryKind), tone: "info" },
  ].forEach((item) => {
    const chip = document.createElement("div");
    chip.className = `sheet-inspector-chip${item.tone ? ` ${item.tone}` : ""}`;
    const label = document.createElement("div");
    label.className = "sheet-inspector-chip-label";
    label.textContent = item.label;
    const value = document.createElement("div");
    value.className = "sheet-inspector-chip-value";
    value.textContent = item.value;
    chip.appendChild(label);
    chip.appendChild(value);
    summaryGrid.appendChild(chip);
  });
  aggregationWorkbenchSummaryEl.appendChild(summaryGrid);

  const controls = document.createElement("div");
  controls.className = "aggregation-controls";

  const sourceField = document.createElement("label");
  sourceField.className = "field";
  sourceField.append("Zrodlo");
  const sourceSelect = document.createElement("select");
  sourceSelect.dataset.aggregationControl = "source";
  [
    { value: "auto", label: "Auto" },
    { value: "wide", label: "Widok klasyczny" },
    ...(result.longAvailable ? [{ value: "long", label: "Wide-to-Long" }] : []),
  ].forEach((item) => {
    const option = document.createElement("option");
    option.value = item.value;
    option.textContent = item.label;
    sourceSelect.appendChild(option);
  });
  sourceSelect.value = aggregationWorkbenchState.sourceMode;
  sourceField.appendChild(sourceSelect);

  const scopeField = document.createElement("label");
  scopeField.className = "field";
  scopeField.append("Zakres");
  const scopeSelect = document.createElement("select");
  scopeSelect.dataset.aggregationControl = "scope";
  [
    { value: "filtered", label: "Aktualny widok" },
    { value: "all", label: "Caly arkusz" },
  ].forEach((item) => {
    const option = document.createElement("option");
    option.value = item.value;
    option.textContent = item.label;
    scopeSelect.appendChild(option);
  });
  scopeSelect.value = aggregationWorkbenchState.scopeMode;
  scopeField.appendChild(scopeSelect);

  const headerField = document.createElement("label");
  headerField.className = "field";
  headerField.append("Naglowek agregacji");
  const headerRowWrap = document.createElement("div");
  headerRowWrap.className = "aggregation-header-row";
  const headerSelect = document.createElement("select");
  headerSelect.dataset.aggregationControl = "header";
  const headerNumberInput = document.createElement("input");
  headerNumberInput.dataset.aggregationControl = "header-number";
  headerNumberInput.type = "number";
  headerNumberInput.min = "1";
  headerNumberInput.step = "1";
  headerNumberInput.inputMode = "numeric";
  headerNumberInput.placeholder = "nr wiersza";
  headerRowWrap.appendChild(headerSelect);
  headerRowWrap.appendChild(headerNumberInput);
  headerField.appendChild(headerRowWrap);
  const autoOpt = document.createElement("option");
  autoOpt.value = "auto";
  autoOpt.textContent = `Auto (najlepszy: ${result.resolvedHeaderRow})`;
  headerSelect.appendChild(autoOpt);
  const manualOpt = document.createElement("option");
  manualOpt.value = "manual";
  manualOpt.textContent = "Wlasny numer";
  headerSelect.appendChild(manualOpt);
  headerSelect.value = aggregationWorkbenchState.headerRowChoice;
  headerNumberInput.value = String(aggregationWorkbenchState.customHeaderRow || currentHeaderRow);
  headerNumberInput.disabled = aggregationWorkbenchState.headerRowChoice !== "manual";

  const groupField = document.createElement("label");
  groupField.className = "field";
  groupField.append("Grupuj po");
  const groupSelect = document.createElement("select");
  groupSelect.dataset.aggregationControl = "group";
  groupField.appendChild(groupSelect);
  result.groupOptions.forEach((option) => {
    const opt = document.createElement("option");
    opt.value = option.value;
    opt.textContent = option.label;
    groupSelect.appendChild(opt);
  });
  groupSelect.value = aggregationWorkbenchState.groupBy;

  const measureField = document.createElement("label");
  measureField.className = "field";
  measureField.append("Mierz");
  const measureSelect = document.createElement("select");
  measureSelect.dataset.aggregationControl = "measure";
  measureField.appendChild(measureSelect);
  result.measures.forEach((candidate) => {
    const opt = document.createElement("option");
    opt.value = candidate.key;
    opt.textContent = candidate.label;
    measureSelect.appendChild(opt);
  });
  measureSelect.value = aggregationWorkbenchState.measure;

  const aggregationField = document.createElement("label");
  aggregationField.className = "field";
  aggregationField.append("Agregacja");
  const aggregationSelect = document.createElement("select");
  aggregationSelect.dataset.aggregationControl = "aggregation";
  aggregationField.appendChild(aggregationSelect);
  const aggregationTooltips = {
    count: "Ile wierszy jest w kazdej grupie (np. 15 wierszy w Krakowie)",
    avg: "Srednia wartosc (np. sredni czas lub srednia kwota)",
    median: "Srodkowa wartosc (polowa wartosci jest mniejsza, polowa wieksza)",
    min: "Najmniejsza wartosc w grupie",
    max: "Najwieksza wartosc w grupie",
    sum: "Laczna suma wszystkich wartosci w grupie",
    distinct: "Ile roznych, niepowtarzalnych wartosci (np. 5 roznych klientow)",
  };
  result.allowedAggregations.forEach((key) => {
    const opt = document.createElement("option");
    opt.value = key;
    opt.textContent = aggregationLabels[key];
    opt.title = aggregationTooltips[key] || "";
    aggregationSelect.appendChild(opt);
  });
  aggregationSelect.value = aggregationWorkbenchState.aggregation;

  const matchField = document.createElement("label");
  matchField.className = "field";
  matchField.append("Dopasowanie tekstu");
  const matchSelect = document.createElement("select");
  matchSelect.dataset.aggregationControl = "match";
  [
    { value: "contains", label: "Zawiera" },
    { value: "exact", label: "Dokladnie" },
  ].forEach((item) => {
    const option = document.createElement("option");
    option.value = item.value;
    option.textContent = item.label;
    matchSelect.appendChild(option);
  });
  matchSelect.value = aggregationWorkbenchState.matchMode;
  matchField.appendChild(matchSelect);

const measureFilterField = document.createElement("label");
  measureFilterField.className = "field";
  measureFilterField.append("Zawiera w mierze");
  const measureFilterSelect = document.createElement("select");
  measureFilterSelect.dataset.aggregationControl = "measurefilter";
  [
    { value: "all", label: "Wszystkie", title: "Pokaz wszystkie wiersze bez filtrowania" },
    { value: "contains", label: "Zawiera", title: "Znajdz wiersze zawierajace szukany tekst (np. czesc imienia)" },
    { value: "exact", label: "Dokladnie", title: "Znajdz wiersze dokladnie rowne szukanej wartosci" },
  ].forEach((item) => {
    const option = document.createElement("option");
    option.value = item.value;
    option.textContent = item.label;
    option.title = item.title || "";
    measureFilterSelect.appendChild(option);
  });
  measureFilterSelect.value = aggregationWorkbenchState.measureFilterMode || "all";
  measureFilterField.appendChild(measureFilterSelect);

  const measureFilterInput = document.createElement("input");
  measureFilterInput.type = "text";
  measureFilterInput.className = "aggregation-measurefilter-value";
  measureFilterInput.dataset.aggregationControl = "measurefilter-value";
  measureFilterInput.value = aggregationWorkbenchState.measureFilterValue || "";
  measureFilterInput.placeholder = "szukaj...";
  measureFilterInput.title = "Szukana wartosc w kolumnie mierzonej";
  measureFilterInput.style.display = aggregationWorkbenchState.measureFilterMode === "all" ? "none" : "inline-block";
  measureFilterField.appendChild(measureFilterInput);

  const showCountField = document.createElement("label");
  showCountField.className = "field";
  showCountField.append("Pokaz wynikow");
  const showCountSelect = document.createElement("select");
  showCountSelect.dataset.aggregationControl = "count";
  ["10", "20", "40", "80", "999"].forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value === "999" ? "Wszystkie" : value;
    showCountSelect.appendChild(option);
  });
  showCountSelect.value = String(aggregationWorkbenchState.showCount);
  showCountField.appendChild(showCountSelect);

  const havingField = document.createElement("label");
  havingField.className = "field";
  havingField.append("Filtrowanie grup");
  const havingSelect = document.createElement("select");
  havingSelect.dataset.aggregationControl = "having";
  [
    { value: "all", label: "Wszystkie", title: "Pokaz wszystkie grupy bez ograniczen" },
    { value: "above_value", label: " Wartosc >", title: "Pokaz tylko grupy z wartoscia wieksza niz podana liczba" },
    { value: "above_percent", label: "% sumy >", title: "Pokaz grupy ktore stanowia wiecej niz X% lacznej sumy wszystkich grup" },
    { value: "above_max_percent", label: "% max >", title: "Pokaz grupy wieksze niz polowa najsilniejszej grupy" },
  ].forEach((item) => {
    const option = document.createElement("option");
    option.value = item.value;
    option.textContent = item.label;
    option.title = item.title;
    havingSelect.appendChild(option);
  });
  havingSelect.value = aggregationWorkbenchState.havingMode;
  havingField.appendChild(havingSelect);

  const havingValueInput = document.createElement("input");
  havingValueInput.type = "number";
  havingValueInput.className = "aggregation-having-value";
  havingValueInput.dataset.aggregationControl = "having-value";
  havingValueInput.value = String(aggregationWorkbenchState.havingValue);
  havingValueInput.min = "0";
  havingValueInput.step = "1";
  havingValueInput.title = "Podaj wartosc progową (np. 10 oznacza >10)";
  havingValueInput.style.display = aggregationWorkbenchState.havingMode === "all" ? "none" : "inline-block";
  havingField.appendChild(havingValueInput);

  [sourceField, scopeField, headerField, groupField, measureField, aggregationField, matchField, measureFilterField, showCountField, havingField].forEach((field) => controls.appendChild(field));
  aggregationWorkbenchSummaryEl.appendChild(controls);

  const note = document.createElement("div");
  note.className = "duration-analysis-note";
  const headerModeText = aggregationWorkbenchState.headerRowChoice === "auto"
    ? `auto -> wiersz ${result.resolvedHeaderRow}`
    : `wiersz ${result.resolvedHeaderRow}`;
  const havingText = aggregationWorkbenchState.havingMode === "all"
    ? ""
    : aggregationWorkbenchState.havingMode === "above_value"
      ? ` • filtr: > ${aggregationWorkbenchState.havingValue}`
      : ` • filtr: > ${aggregationWorkbenchState.havingValue}%`;
  note.textContent = `Aktualne zrodlo: ${result.model.mode === "long" ? "Wide-to-Long" : "widok klasyczny"} • zakres: ${aggregationWorkbenchState.scopeMode === "all" ? "caly arkusz" : "aktualny widok"} • naglowek: ${headerModeText}${result.helperMode ? " (pomocniczy)" : ""} • dopasowanie: ${aggregationWorkbenchState.matchMode === "exact" ? "dokladnie" : "zawiera"}${havingText}.`;
  aggregationWorkbenchSummaryEl.appendChild(note);

  const currentSearch = aggregationWorkbenchState.resultSearch || "";
  const filteredEntries = currentSearch
    ? result.entries.filter((e) => e.label.toLowerCase().includes(currentSearch.toLowerCase()))
    : result.entries;
  const searchWrap = document.createElement("div");
  searchWrap.className = "aggregation-result-search-wrap";

  const searchLabel = document.createElement("span");
  searchLabel.textContent = "Szukaj:";
  searchLabel.style.fontSize = "12px";
  searchLabel.style.color = "var(--muted)";
  searchWrap.appendChild(searchLabel);

  const resultSearchInput = document.createElement("input");
  resultSearchInput.type = "text";
  resultSearchInput.className = "aggregation-result-search";
  resultSearchInput.placeholder = "np. Julian...";
  resultSearchInput.title = "Wpisz tekst ktory ma sie zawierac w nazwie grupy (np. czesc imienia)";
  resultSearchInput.style.flex = "1";
  resultSearchInput.style.minWidth = "100px";
  resultSearchInput.style.padding = "4px 8px";
  resultSearchInput.style.borderRadius = "var(--r-sm)";
  resultSearchInput.style.border = "1px solid var(--border)";
  resultSearchInput.style.fontSize = "13px";
  resultSearchInput.value = currentSearch;
  searchWrap.appendChild(resultSearchInput);

  const searchCount = document.createElement("span");
  searchCount.style.fontSize = "12px";
  searchCount.style.color = "var(--muted)";
  searchCount.style.marginLeft = "8px";
  searchCount.style.whiteSpace = "nowrap";
  searchCount.textContent = currentSearch ? `${filteredEntries.length} z ${result.entries.length}` : "";
  searchWrap.appendChild(searchCount);

  aggregationWorkbenchListEl.appendChild(searchWrap);

  const showCount = Math.min(aggregationWorkbenchState.showCount, filteredEntries.length);
  filteredEntries.slice(0, showCount).forEach((entry, index) => {
    const item = document.createElement("div");
    item.className = "aggregation-item";

    const top = document.createElement("div");
    top.className = "duration-person-top";

    const titleWrap = document.createElement("div");
    titleWrap.className = "duration-person-title-wrap";

    const rank = document.createElement("div");
    rank.className = "duration-person-rank";
    rank.textContent = String(index + 1);

    const title = document.createElement("div");
    title.className = "duration-person-title";
    title.textContent = entry.label;

    const value = document.createElement("div");
    value.className = "duration-person-value";
    value.textContent = formatAggregationMetricValue(entry.primary, primaryKind);

    titleWrap.appendChild(rank);
    titleWrap.appendChild(title);
    top.appendChild(titleWrap);
    top.appendChild(value);

    const meta = document.createElement("div");
    meta.className = "duration-person-meta";
    meta.textContent = `Liczba ${entry.count} • srednia ${formatAggregationMetricValue(entry.average, primaryKind)} • mediana ${formatAggregationMetricValue(entry.median, primaryKind)} • zakres ${formatAggregationMetricValue(entry.min, primaryKind)} -> ${formatAggregationMetricValue(entry.max, primaryKind)}`;

    const actions = document.createElement("div");
    actions.className = "section-nav-actions";
    const btn = document.createElement("button");
    btn.className = "btn ghost btn-sm";
    btn.type = "button";
    btn.dataset.aggregationAction = "filter-group";
    btn.dataset.aggregationValue = entry.label;
    btn.textContent = "Szukaj w tabeli";
    actions.appendChild(btn);

    item.appendChild(top);
    item.appendChild(meta);
    item.appendChild(actions);
    aggregationWorkbenchListEl.appendChild(item);
  });
}

function buildMergedLabelMap(merges, sheet, tableStartCol, tableEndCol, headerAbsRow) {
  const labels = new Map();
  merges
    .filter((merge) => merge && merge.s && merge.e && merge.s.r < headerAbsRow)
    .forEach((merge) => {
      if (merge.e.c < tableStartCol || merge.s.c > tableEndCol) return;
      const label = getCellDisplayText(sheet, merge.s.r, merge.s.c);
      if (!label) return;
      const startIndex = Math.max(0, merge.s.c - tableStartCol);
      labels.set(startIndex, label);
    });
  return labels;
}

function isMeaningfulSheetCell(cell) {
  if (!cell || typeof cell !== "object") return false;
  if (cell.f) return true;
  if (cell.l && (cell.l.Target || cell.l.target)) return true;
  if (Array.isArray(cell.c) && cell.c.length) return true;
  if (cell.v instanceof Date) return true;
  if (typeof cell.v === "number" && Number.isFinite(cell.v)) return true;
  if (typeof cell.v === "boolean") return true;
  if (typeof cell.v === "string" && cell.v.trim() !== "") return true;
  if (typeof cell.w === "string" && cell.w.trim() !== "") return true;
  return false;
}

function computeEffectiveSheetRange(sheet, headerRow) {
  const fallback = XLSX.utils.decode_range(sheet["!ref"]);
  const headerIndex0 = Math.max(0, (headerRow || 1) - 1);
  let minCol = fallback.e.c;
  let maxCol = fallback.s.c;
  let maxRow = headerIndex0;
  let found = false;

  Object.keys(sheet).forEach((key) => {
    if (!key || key[0] === "!") return;
    const cell = sheet[key];
    if (!isMeaningfulSheetCell(cell)) return;
    const ref = XLSX.utils.decode_cell(key);
    if (ref.r < headerIndex0) {
      minCol = Math.min(minCol, ref.c);
      maxCol = Math.max(maxCol, ref.c);
      found = true;
      return;
    }
    minCol = Math.min(minCol, ref.c);
    maxCol = Math.max(maxCol, ref.c);
    maxRow = Math.max(maxRow, ref.r);
    found = true;
  });

  for (let c = fallback.s.c; c <= fallback.e.c; c++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: headerIndex0, c })];
    if (!isMeaningfulSheetCell(cell)) continue;
    minCol = Math.min(minCol, c);
    maxCol = Math.max(maxCol, c);
    found = true;
  }

  if (!found) {
    return fallback;
  }

  const merges = Array.isArray(sheet["!merges"]) ? sheet["!merges"] : [];
  merges.forEach((merge) => {
    if (!merge || !merge.s || !merge.e) return;
    const anchor = sheet[XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c })];
    if (!isMeaningfulSheetCell(anchor)) return;
    minCol = Math.min(minCol, merge.s.c);
    maxCol = Math.max(maxCol, merge.e.c);
    if (merge.e.r >= headerIndex0) {
      maxRow = Math.max(maxRow, merge.e.r);
    }
  });

  return {
    s: { r: fallback.s.r, c: Math.max(fallback.s.c, minCol) },
    e: { r: Math.max(headerIndex0, maxRow), c: Math.max(Math.max(fallback.s.c, minCol), maxCol) },
  };
}

function buildGroupFromSignature(headers, startIndex, span, repeatCount, tableStartCol, mergedLabels) {
  const bases = headers
    .slice(startIndex, startIndex + span)
    .map((header) => parseRepeatedHeader(header)?.base || cleanSectionLabel(header))
    .filter(Boolean);
  const uniqueBases = Array.from(new Set(bases));
  const blocks = [];

  for (let i = 0; i < repeatCount; i++) {
    const blockStart = startIndex + (i * span);
    const blockEnd = blockStart + span - 1;
    blocks.push({
      label: mergedLabels.get(blockStart) || `Blok ${i + 1}`,
      span,
      startIndex: blockStart,
      endIndex: blockEnd,
      startAbs: tableStartCol + blockStart,
      endAbs: tableStartCol + blockEnd,
      headers: headers.slice(blockStart, blockEnd + 1),
    });
  }

  return {
    label: repeatCount >= 2 ? `Powtarzalny układ: ${repeatCount} bloków` : "Powtarzalny układ kolumn",
    kind: "repeating-signature",
    meta: `${repeatCount} bloków po ${span} kolumny`,
    prefixCount: startIndex,
    prefixLabel: startIndex > 0 ? formatColRange(tableStartCol, tableStartCol + startIndex - 1) : "",
    families: uniqueBases.slice(0, 8).map((label) => ({ label, count: repeatCount })),
    blocks,
  };
}

function detectSignatureRepeatingBlocks(headers, tableStartCol, mergedLabels) {
  if (!Array.isArray(headers) || headers.length < 4) return [];

  let best = null;
  const maxSpan = Math.min(12, Math.floor(headers.length / 2));

  for (let startIndex = 0; startIndex < headers.length - 3; startIndex++) {
    for (let span = 2; span <= maxSpan; span++) {
      if (startIndex + (span * 2) > headers.length) break;

      const template = headers.slice(startIndex, startIndex + span).map((header) => parseRepeatedHeader(header)?.base || cleanSectionLabel(header));
      if (!template.some(Boolean)) continue;

      let repeatCount = 1;
      let nextStart = startIndex + span;

      while (nextStart + span <= headers.length) {
        const candidate = headers.slice(nextStart, nextStart + span).map((header) => parseRepeatedHeader(header)?.base || cleanSectionLabel(header));
        if (candidate.length !== template.length) break;
        if (!candidate.every((value, idx) => value === template[idx])) break;
        repeatCount += 1;
        nextStart += span;
      }

      if (repeatCount < 2) continue;

      const score = (repeatCount * span * 100) - startIndex;
      if (!best || score > best.score) {
        best = { score, startIndex, span, repeatCount };
      }
    }
  }

  if (!best) return [];
  return [buildGroupFromSignature(headers, best.startIndex, best.span, best.repeatCount, tableStartCol, mergedLabels)];
}

function detectRepeatingBlocks(sheet, headerRow, data) {
  if (!sheet || !data || !Array.isArray(data.headers) || !data.headers.length) return [];
  const merges = Array.isArray(data.merges) ? data.merges : [];
  const headerAbsRow = headerRow - 1;
  const tableStartCol = data.startCol || 0;
  const tableEndCol = tableStartCol + data.headers.length - 1;
  const mergedLabels = buildMergedLabelMap(merges, sheet, tableStartCol, tableEndCol, headerAbsRow);

  const signatureGroups = detectSignatureRepeatingBlocks(data.headers, tableStartCol, mergedLabels);
  if (signatureGroups.length) {
    return signatureGroups;
  }

  const groups = [];

  const mergeBlocks = merges
    .filter((merge) => merge && merge.s && merge.e && merge.s.r < headerAbsRow && merge.e.c >= tableStartCol && merge.s.c <= tableEndCol)
    .sort((a, b) => a.s.c - b.s.c);

  if (mergeBlocks.length >= 2) {
    const bySpan = new Map();
    mergeBlocks.forEach((merge) => {
      const span = merge.e.c - merge.s.c + 1;
      if (span < 2) return;
      const label = getCellDisplayText(sheet, merge.s.r, merge.s.c);
      if (!label) return;
      const list = bySpan.get(span) || [];
      const startIndex = Math.max(0, merge.s.c - tableStartCol);
      const endIndex = Math.min(data.headers.length - 1, merge.e.c - tableStartCol);
      list.push({
        label,
        span,
        startIndex,
        endIndex,
        startAbs: merge.s.c,
        endAbs: merge.e.c,
        headers: data.headers.slice(startIndex, endIndex + 1),
      });
      bySpan.set(span, list);
    });

    bySpan.forEach((blocks, span) => {
      if (blocks.length < 2) return;
      const familyMap = new Map();
      blocks.forEach((block) => {
        block.headers.forEach((header) => {
          const parsed = parseRepeatedHeader(header);
          const key = parsed ? parsed.base : header;
          familyMap.set(key, (familyMap.get(key) || 0) + 1);
        });
      });
      const families = Array.from(familyMap.entries())
        .filter(([, count]) => count >= 2)
        .map(([label, count]) => ({ label, count }))
        .slice(0, 8);

      groups.push({
        label: `Powtarzalny układ: ${blocks.length} bloków`,
        kind: "merged",
        meta: `${blocks.length} bloków po ${span} kolumny • ${blocks[0].label} -> ${blocks[blocks.length - 1].label}`,
        prefixCount: blocks[0].startIndex,
        prefixLabel: blocks[0].startIndex > 0 ? formatColRange(tableStartCol, tableStartCol + blocks[0].startIndex - 1) : "",
        families,
        blocks,
      });
    });
  }

  if (groups.length) return groups.slice(0, 4);

  const familyMap = new Map();
  data.headers.forEach((header, index) => {
    const parsed = parseRepeatedHeader(header);
    if (!parsed) return;
    const entry = familyMap.get(parsed.base) || { label: parsed.base, indexes: [], orders: [] };
    entry.indexes.push(index);
    entry.orders.push(parsed.order);
    familyMap.set(parsed.base, entry);
  });
  const families = Array.from(familyMap.values())
    .filter((entry) => entry.indexes.length >= 3)
    .sort((a, b) => b.indexes.length - a.indexes.length);

  if (!families.length) return [];

  return [{
    label: "Powtarzalne rodziny kolumn",
    kind: "family",
    meta: `${families.length} rodzin powtarzalnych kolumn`,
    prefixCount: 0,
    prefixLabel: "",
    families: families.slice(0, 10).map((entry) => ({ label: entry.label, count: entry.indexes.length })),
    blocks: families.slice(0, 10).map((entry) => ({
      label: entry.label,
      span: 1,
      startIndex: entry.indexes[0],
      endIndex: entry.indexes[entry.indexes.length - 1],
      startAbs: tableStartCol + entry.indexes[0],
      endAbs: tableStartCol + entry.indexes[entry.indexes.length - 1],
      headers: entry.indexes.map((idx) => data.headers[idx]),
    })),
  }];
}

function renderRepeatingBlocks() {
  if (!repeatBlockDetectorEl) return;
  repeatBlockDetectorEl.replaceChildren();
  if (!currentRepeatingBlocks.length) {
    repeatBlockDetectorEl.appendChild(createEmptyInsight("Brak wyraznych powtarzalnych blokow dla aktualnego arkusza. Najlepiej dziala na szerokich tabelach z cyklami, etapami albo seriami podobnych kolumn."));
    return;
  }

  currentRepeatingBlocks.forEach((group, groupIndex) => {
    const summary = document.createElement("div");
    summary.className = "repeat-summary";
    const prefixNote = group.prefixCount ? ` • stałe kolumny przed blokami: ${group.prefixLabel}` : "";
    summary.textContent = `${group.meta || group.label}${prefixNote}`;
    repeatBlockDetectorEl.appendChild(summary);

    group.blocks.slice(0, 10).forEach((block, blockIndex) => {
      const item = document.createElement("div");
      item.className = "repeat-block-item";

      const top = document.createElement("div");
      top.className = "repeat-block-top";

      const title = document.createElement("div");
      title.className = "repeat-block-title";
      title.textContent = block.label;

      const badge = document.createElement("div");
      badge.className = "repeat-block-badge";
      badge.textContent = `${block.span} kol.`;

      top.appendChild(title);
      top.appendChild(badge);

      const meta = document.createElement("div");
      meta.className = "repeat-block-meta";
      const headerPreview = block.headers.slice(0, 4).join(" • ");
      meta.textContent = `Kolumny ${formatColRange(block.startAbs, block.endAbs)}${headerPreview ? ` • ${headerPreview}` : ""}`;

      const actions = document.createElement("div");
      actions.className = "section-nav-actions";

      const btn = document.createElement("button");
      btn.className = "btn ghost btn-sm";
      btn.type = "button";
      btn.dataset.repeatGroupIndex = String(groupIndex);
      btn.dataset.repeatBlockIndex = String(blockIndex);
      btn.textContent = "Skocz do bloku";
      actions.appendChild(btn);

      item.appendChild(top);
      item.appendChild(meta);
      item.appendChild(actions);

      if (group.families && group.families.length) {
        const familyList = document.createElement("div");
        familyList.className = "repeat-family-list";
        group.families.slice(0, 6).forEach((family) => {
          const chip = document.createElement("div");
          chip.className = "repeat-family-chip";
          chip.textContent = `${family.label} ×${family.count}`;
          familyList.appendChild(chip);
        });
        item.appendChild(familyList);
      }

      repeatBlockDetectorEl.appendChild(item);
    });
  });
}

function focusRepeatingBlock(groupIndex, blockIndex) {
  const group = currentRepeatingBlocks[groupIndex];
  const block = group && group.blocks ? group.blocks[blockIndex] : null;
  if (!block || !tableWrapEl) return;
  const cells = theadEl.querySelectorAll(".guide-row .guide-cell");
  const cell = cells[block.startIndex];
  if (!cell) {
    toast("Tego bloku nie widac jeszcze w aktualnym widoku arkusza", "info");
    return;
  }
  const targetLeft = Math.max(0, cell.offsetLeft - 64);
  tableWrapEl.scrollTo({ left: targetLeft, behavior: "smooth" });
  syncHorizontalScrollbar();
}

function collectWorkbookStats(wb, fileName) {
  const book = wb && wb.Workbook ? wb.Workbook : {};
  const sheetsMeta = Array.isArray(book.Sheets) ? book.Sheets : [];
  let hiddenSheets = 0;
  let veryHiddenSheets = 0;
  sheetsMeta.forEach((sheetMeta) => {
    const hidden = Number(sheetMeta && sheetMeta.Hidden);
    if (hidden === 1) hiddenSheets += 1;
    if (hidden === 2) veryHiddenSheets += 1;
  });

  const definedNames = Array.isArray(book.Names) ? book.Names.length : 0;
  const ext = String(fileName || "").toLowerCase();
  const hasMacros = !!wb?.vbaraw || ext.endsWith(".xlsm");

  return {
    sheets: Array.isArray(wb?.SheetNames) ? wb.SheetNames.length : 0,
    hiddenSheets,
    veryHiddenSheets,
    definedNames,
    hasMacros,
  };
}

function collectSheetInsights() {
  const workbookItems = currentWorkbookStats ? [
    { label: "Arkusze", value: String(currentWorkbookStats.sheets) },
    { label: "Ukryte arkusze", value: String(currentWorkbookStats.hiddenSheets), tone: currentWorkbookStats.hiddenSheets ? "warning" : "" },
    { label: "Very hidden", value: String(currentWorkbookStats.veryHiddenSheets), tone: currentWorkbookStats.veryHiddenSheets ? "warning" : "" },
    { label: "Nazwane zakresy", value: String(currentWorkbookStats.definedNames), tone: currentWorkbookStats.definedNames ? "info" : "" },
  ] : [];

  if (!currentHeaders.length || !baseRows.length) {
    return {
      workbookRows: workbookItems,
      rows: [],
      flags: currentWorkbookStats?.hasMacros ? [{ label: "Plik makr .xlsm", tone: "warning" }] : [],
    };
  }

  const totalRows = baseRows.length;
  const visibleRows = viewRows.length;
  const totalCols = currentHeaders.length;
  const duplicateHeaders = currentSheetStats?.duplicateHeaderCount || 0;
  const duplicateRows = (() => {
    const keys = baseRows.map((row) => JSON.stringify(row.values.map((value) => value instanceof Date ? value.toISOString() : value ?? "")));
    return keys.length - new Set(keys).size;
  })();

  let numericColumns = 0;
  let dateColumns = 0;
  let longTextColumns = 0;
  let sparseColumns = 0;

  currentHeaders.forEach((_, colIdx) => {
    let nonEmpty = 0;
    let numeric = 0;
    let dates = 0;
    let maxLen = 0;
    baseRows.forEach((row) => {
      const value = row.values[colIdx];
      if (value === null || value === undefined || String(value).trim() === "") return;
      nonEmpty += 1;
      if (typeof value === "number") numeric += 1;
      if (parseDateFlexible(value) instanceof Date) dates += 1;
      maxLen = Math.max(maxLen, String(getDisplayValue(row, colIdx)).length);
    });
    if (nonEmpty && numeric / nonEmpty >= 0.8) numericColumns += 1;
    if (nonEmpty && dates / nonEmpty >= 0.8) dateColumns += 1;
    if (maxLen > 150) longTextColumns += 1;
    if (totalRows && nonEmpty / totalRows <= 0.4) sparseColumns += 1;
  });

  const sheetItems = [
    { label: "Widoczne / wszystkie wiersze", value: `${visibleRows} / ${totalRows}`, tone: visibleRows !== totalRows ? "info" : "" },
    { label: "Kolumny", value: String(totalCols) },
    { label: "Formuły", value: String(currentSheetStats?.formulaCount || 0), tone: (currentSheetStats?.formulaCount || 0) ? "info" : "" },
    {
      label: "Scalenia (zakresy / komorki)",
      value: `${currentSheetStats?.mergeRegions || 0} / ${currentSheetStats?.mergedCells || 0}`,
      tone: (currentSheetStats?.mergeRegions || 0) ? "info" : "",
    },
    { label: "Ukryte kolumny / wiersze", value: `${currentSheetStats?.hiddenColumns || 0} / ${currentSheetStats?.hiddenRows || 0}`, tone: ((currentSheetStats?.hiddenColumns || 0) || (currentSheetStats?.hiddenRows || 0)) ? "warning" : "" },
    { label: "Kolumny liczbowe / datowe", value: `${numericColumns} / ${dateColumns}` },
    { label: "Rzadkie kolumny", value: `${sparseColumns} (${formatPercent(sparseColumns, totalCols)})`, tone: sparseColumns ? "warning" : "" },
    { label: "Długie teksty", value: String(longTextColumns), tone: longTextColumns ? "info" : "" },
  ];

  const flags = [];
  if (currentWorkbookStats?.hasMacros) flags.push({ label: "Plik makr .xlsm", tone: "warning" });
  if (duplicateHeaders) flags.push({ label: `Zdublowane nagłówki: ${duplicateHeaders}`, tone: "warning" });
  if (duplicateRows) flags.push({ label: `Duplikaty wierszy: ${duplicateRows}`, tone: duplicateRows > 0 ? "warning" : "" });
  if ((currentSheetStats?.formulaMissingResultCount || 0) > 0) {
    flags.push({ label: `Formuły bez wyniku: ${currentSheetStats.formulaMissingResultCount}`, tone: "warning" });
  }
  if ((currentSheetStats?.commentCount || 0) > 0) flags.push({ label: `Komentarze: ${currentSheetStats.commentCount}`, tone: "info" });
  if ((currentSheetStats?.hyperlinkCount || 0) > 0) flags.push({ label: `Linki: ${currentSheetStats.hyperlinkCount}`, tone: "info" });
  if (currentWorkbookStats?.veryHiddenSheets) flags.push({ label: "Są arkusze very hidden", tone: "warning" });

  return {
    workbookRows: workbookItems,
    rows: sheetItems,
    flags,
  };
}

function isKpiLabelCandidate(text) {
  const label = cleanSectionLabel(text);
  if (!label || label.length < 3 || label.length > 48) return false;
  if (/^\d+$/.test(label)) return false;
  const lowered = label.toLowerCase();
  const keywords = [
    "suma",
    "razem",
    "wartosc",
    "wartość",
    "koszt",
    "budzet",
    "budżet",
    "roznica",
    "różnica",
    "saldo",
    "marza",
    "marża",
    "przychod",
    "przychód",
    "wynik",
    "netto",
    "brutto",
    "plan",
    "wykonanie",
    "liczba",
    "ilosc",
    "ilość",
    "procent",
    "udzial",
    "udział",
    "kpi",
  ];
  return keywords.some((keyword) => lowered.includes(keyword)) || /[:\-]$/.test(label);
}

function isKpiValueCandidate(cell, displayText) {
  if (!cell) return false;
  const display = cleanSectionLabel(displayText);
  if (!display) return false;
  if (typeof cell.v === "number" && Number.isFinite(cell.v)) return true;
  if (cell.v instanceof Date) return true;
  if (/%|zl|zł|pln|eur|usd/i.test(display)) return true;
  if (/^-?\d[\d\s,.\u00A0]*$/.test(display)) return true;
  return false;
}

function normalizeKpiLabel(label) {
  const raw = cleanSectionLabel(label).toLowerCase();
  if (!raw) return "";
  return raw
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\b(na|do|od|i|oraz|jeszcze|samochod|samochodu|wartosc|wartosc)\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function inferKpiSemanticBucket(label) {
  const normalized = normalizeKpiLabel(label);
  if (!normalized) return "";
  if (/budzet/.test(normalized)) return "budget";
  if (/koszt|suma|razem|subtotal/.test(normalized)) return "total";
  if (/roznic|saldo|wynik/.test(normalized)) return "difference";
  if (/marz|procent|udzial|wykonanie/.test(normalized)) return "ratio";
  return normalized;
}

function scoreKpiLabelQuality(label) {
  const clean = cleanSectionLabel(label);
  const normalized = normalizeKpiLabel(label);
  if (!clean || !normalized) return 0;
  let score = normalized.length;
  if (normalized.includes(" ")) score += 4;
  if (clean.length >= 16) score += 4;
  if (/^(suma|razem|subtotal|wynik)$/i.test(normalized)) score -= 8;
  if (/koszt|budzet|roznic|saldo|marza|wykonanie|przychod|brutto|netto/i.test(normalized)) score += 6;
  return score;
}

function dedupeKpiEntries(entries) {
  const bestByKey = new Map();
  entries.forEach((entry) => {
    const semantic = inferKpiSemanticBucket(entry.label);
    const valueKey = cleanSectionLabel(entry.value);
    const key = `${semantic}|${valueKey}`;
    const existing = bestByKey.get(key);
    if (!existing) {
      bestByKey.set(key, {
        ...entry,
        aliases: [],
      });
      return;
    }
    if (entry.label !== existing.label && !existing.aliases.includes(entry.label)) {
      existing.aliases.push(entry.label);
    }
    const existingCompositeScore = existing.score + scoreKpiLabelQuality(existing.label);
    const entryCompositeScore = entry.score + scoreKpiLabelQuality(entry.label);
    if (entryCompositeScore > existingCompositeScore) {
      bestByKey.set(key, {
        ...entry,
        aliases: Array.from(new Set([existing.label, ...existing.aliases])),
      });
    }
  });
  return Array.from(bestByKey.values());
}

function pushKpiEntry(entries, seen, labelText, valueText, valueRef, labelRef, valueCell, scoreExtras = 0) {
  const cleanLabel = cleanSectionLabel(labelText).replace(/[:\-]\s*$/, "");
  const cleanValue = cleanSectionLabel(valueText);
  if (!cleanLabel || !cleanValue) return;
  const seenKey = `${normalizeKpiLabel(cleanLabel)}|${valueRef}`;
  if (seen.has(seenKey)) return;
  seen.add(seenKey);

  const decoded = XLSX.utils.decode_cell(valueRef);
  let score = scoreExtras;
  if (typeof valueCell?.v === "number") score += 3;
  if (valueCell?.f) score += 2;
  if (/%|zl|zł|pln|eur|usd/i.test(cleanValue)) score += 2;
  if (isKpiLabelCandidate(cleanLabel)) score += 2;

  entries.push({
    label: cleanLabel,
    value: cleanValue,
    address: valueRef,
    labelAddress: labelRef,
    rowIndex0: decoded.r,
    colAbs: decoded.c,
    score,
  });
}

function findDistantSameRowKpiTarget(sheet, rowIndex0, labelColAbs, endColAbs) {
  if (!sheet) return null;
  const farCandidates = [];
  const maxOffset = Math.min(6, endColAbs - labelColAbs);
  for (let offset = 3; offset <= maxOffset; offset++) {
    const candidateCol = labelColAbs + offset;
    let gapHasContent = false;
    for (let mid = labelColAbs + 1; mid < candidateCol; mid++) {
      if (getCellDisplayText(sheet, rowIndex0, mid)) {
        gapHasContent = true;
        break;
      }
    }
    if (gapHasContent) continue;
    const valueRef = XLSX.utils.encode_cell({ r: rowIndex0, c: candidateCol });
    const valueCell = sheet[valueRef];
    const valueDisplay = cleanSectionLabel(getCellDisplayText(sheet, rowIndex0, candidateCol));
    if (!isKpiValueCandidate(valueCell, valueDisplay)) continue;
    farCandidates.push({
      valueRef,
      valueCell,
      valueDisplay,
      colAbs: candidateCol,
    });
  }
  return farCandidates.length === 1 ? farCandidates[0] : null;
}

function scanKpiZone(sheet, entries, seen, rowStart, rowEnd, colStart, colEnd) {
  for (let r = rowStart; r <= rowEnd; r++) {
    for (let c = colStart; c <= colEnd; c++) {
      const labelText = getCellDisplayText(sheet, r, c);
      if (!isKpiLabelCandidate(labelText)) continue;

      const candidates = [
        { row: r, col: c + 1 },
        { row: r, col: c + 2 },
        { row: r + 1, col: c },
      ];

      candidates.forEach((candidate) => {
        if (candidate.row > rowEnd || candidate.col > colEnd) return;
        const valueRef = XLSX.utils.encode_cell({ r: candidate.row, c: candidate.col });
        const valueCell = sheet[valueRef];
        const valueDisplay = cleanSectionLabel(getCellDisplayText(sheet, candidate.row, candidate.col));
        if (!isKpiValueCandidate(valueCell, valueDisplay)) return;
        let score = 0;
        if (candidate.row === r) score += 2;
        if (candidate.col === c + 1) score += 1;
        pushKpiEntry(
          entries,
          seen,
          labelText,
          valueDisplay,
          valueRef,
          XLSX.utils.encode_cell({ r, c }),
          valueCell,
          score
        );
      });

      const distantTarget = findDistantSameRowKpiTarget(sheet, r, c, colEnd);
      if (distantTarget) {
        pushKpiEntry(
          entries,
          seen,
          labelText,
          distantTarget.valueDisplay,
          distantTarget.valueRef,
          XLSX.utils.encode_cell({ r, c }),
          distantTarget.valueCell,
          4
        );
      }
    }
  }

  for (let r = rowStart; r <= rowEnd; r++) {
    for (let c = colStart; c <= colEnd; c++) {
      const valueRef = XLSX.utils.encode_cell({ r, c });
      const valueCell = sheet[valueRef];
      const valueDisplay = cleanSectionLabel(getCellDisplayText(sheet, r, c));
      if (!isKpiValueCandidate(valueCell, valueDisplay)) continue;

      const leftLabel = c > colStart ? getCellDisplayText(sheet, r, c - 1) : "";
      const twoLeftLabel = c - 2 >= colStart ? getCellDisplayText(sheet, r, c - 2) : "";
      const aboveLabel = r > rowStart ? getCellDisplayText(sheet, r - 1, c) : "";
      const diagLabel = (r > rowStart && c > colStart) ? getCellDisplayText(sheet, r - 1, c - 1) : "";
      const candidates = [
        { label: leftLabel, ref: c > colStart ? XLSX.utils.encode_cell({ r, c: c - 1 }) : valueRef, bonus: 2 },
        { label: twoLeftLabel, ref: c - 2 >= colStart ? XLSX.utils.encode_cell({ r, c: c - 2 }) : valueRef, bonus: 1 },
        { label: aboveLabel, ref: r > rowStart ? XLSX.utils.encode_cell({ r: r - 1, c }) : valueRef, bonus: 1 },
        { label: diagLabel, ref: (r > rowStart && c > colStart) ? XLSX.utils.encode_cell({ r: r - 1, c: c - 1 }) : valueRef, bonus: 0 },
      ];

      candidates.forEach((candidate) => {
        if (!isKpiLabelCandidate(candidate.label)) return;
        pushKpiEntry(
          entries,
          seen,
          candidate.label,
          valueDisplay,
          valueRef,
          candidate.ref,
          valueCell,
          candidate.bonus
        );
      });
    }
  }
}

function collectKpiEntries(sheet, headerRow) {
  if (!sheet) return { entries: [], anchorRow: headerRow || 1 };
  const inferredAnchorRow = Math.max(1, detectHeaderRowSimple(sheet));
  const anchorRow = Math.max(1, headerRow || 1, inferredAnchorRow);
  if (anchorRow <= 1) return { entries: [], anchorRow };
  const entries = [];
  const seen = new Set();
  const effectiveRange = computeEffectiveSheetRange(sheet, anchorRow);
  const startRow = Math.max(effectiveRange.s.r, anchorRow - 8);
  const endRow = Math.max(startRow, anchorRow - 1);
  const endCol = effectiveRange.e.c;

  scanKpiZone(sheet, entries, seen, startRow, endRow, effectiveRange.s.c, endCol);

  const bottomStartRow = Math.max(anchorRow, effectiveRange.e.r - 4);
  if (bottomStartRow <= effectiveRange.e.r) {
    scanKpiZone(sheet, entries, seen, bottomStartRow, effectiveRange.e.r, effectiveRange.s.c, endCol);
  }

  return {
    anchorRow,
    entries: dedupeKpiEntries(entries).sort((a, b) => {
      if (b.score !== a.score) return b.score - a.score;
      return a.address.localeCompare(b.address, "pl");
    })
    .slice(0, 8),
  };
}

function renderKpiExtractor() {
  if (!kpiSummaryEl || !kpiListEl) return;
  kpiSummaryEl.replaceChildren();
  kpiListEl.replaceChildren();

  if (!currentHeaders.length || !currentKpiEntries.length) {
    renderInsightList(kpiSummaryEl, [], "Brak wykrytych KPI lub podsumowan dla aktualnego arkusza.");
    kpiListEl.appendChild(createEmptyInsight("Nie wykryto wiarygodnych KPI nad aktualna tabela danych."));
    return;
  }

  renderInsightList(kpiSummaryEl, [
    { label: "Kandydaci KPI", value: String(currentKpiEntries.length), tone: "info" },
    {
      label: "Źródło",
      value: currentKpiAnchorRow === currentHeaderRow
        ? `Wiersze nad nagłówkiem ${currentHeaderRow}`
        : `Wiersze nad wykrytym nagłówkiem ${currentKpiAnchorRow}`,
      tone: currentKpiAnchorRow === currentHeaderRow ? "" : "info",
    },
  ], "Brak podsumowania KPI.");

  currentKpiEntries.forEach((entry) => {
    const item = document.createElement("div");
    item.className = "kpi-card";

    const label = document.createElement("div");
    label.className = "kpi-label";
    label.textContent = entry.label;

    const value = document.createElement("div");
    value.className = "kpi-value";
    value.textContent = entry.value;
    value.title = entry.aliases?.length
      ? `${entry.label}: ${entry.value}\nRowniez jako: ${entry.aliases.join(", ")}`
      : `${entry.label}: ${entry.value}`;

    const meta = document.createElement("div");
    meta.className = "kpi-meta";
    meta.textContent = entry.aliases?.length
      ? `${entry.address} • etykieta ${entry.labelAddress} • również jako: ${entry.aliases.slice(0, 2).join(", ")}${entry.aliases.length > 2 ? ` +${entry.aliases.length - 2}` : ""}`
      : `${entry.address} • etykieta ${entry.labelAddress}`;

    const actions = document.createElement("div");
    actions.className = "section-nav-actions";

    const btn = document.createElement("button");
    btn.className = "btn ghost btn-sm";
    btn.type = "button";
    btn.dataset.kpiAddress = entry.address;
    btn.textContent = "Pokaż źródło";

    actions.appendChild(btn);
    item.appendChild(label);
    item.appendChild(value);
    item.appendChild(meta);
    item.appendChild(actions);
    kpiListEl.appendChild(item);
  });
}

function focusKpiEntry(address) {
  const entry = currentKpiEntries.find((item) => item.address === address);
  if (!entry) return;
  const rowEl = tbodyEl.querySelector(`tr[data-row-index="${entry.rowIndex0}"]`);
  if (rowEl) {
    rowEl.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" });
  } else {
    toast(`Źródło KPI jest nad aktualną tabelą: wiersz ${entry.rowIndex0 + 1}`, "info");
    if (tableWrapEl) {
      tableWrapEl.scrollTo({ top: 0, behavior: "smooth" });
    }
  }
  const relativeColIdx = entry.colAbs - currentStartCol;
  if (Number.isFinite(relativeColIdx) && relativeColIdx >= 0) {
    focusColumnProfile(relativeColIdx);
  }
}

function inferColumnProfileType(stats) {
  if (!stats || !stats.nonEmpty) return "pusta";
  const ratio = (count) => (stats.nonEmpty ? count / stats.nonEmpty : 0);
  const numberRatio = ratio(stats.numericCount);
  const dateRatio = ratio(stats.dateCount);
  const formulaRatio = ratio(stats.formulaCount);
  const textRatio = ratio(stats.textCount);

  if (formulaRatio >= 0.8) return "formuly";
  if (dateRatio >= 0.8) return "daty";
  if (numberRatio >= 0.8) return "liczby";
  if (textRatio >= 0.8) return "tekst";
  return "mixed";
}

function formatColumnProfileRange(profile) {
  if (!profile) return "";
  if (profile.type === "liczby" && Number.isFinite(profile.minValue) && Number.isFinite(profile.maxValue)) {
    return `${profile.minValue} -> ${profile.maxValue}`;
  }
  if (profile.type === "daty" && profile.minDate instanceof Date && profile.maxDate instanceof Date) {
    return `${toDisplay(profile.minDate)} -> ${toDisplay(profile.maxDate)}`;
  }
  return "";
}

function collectColumnProfiles() {
  if (!currentHeaders.length || !baseRows.length) return [];
  const totalRows = baseRows.length;
  const profiles = currentHeaders.map((header, colIdx) => {
    const stats = {
      nonEmpty: 0,
      numericCount: 0,
      dateCount: 0,
      textCount: 0,
      formulaCount: 0,
      longTextCount: 0,
      minValue: null,
      maxValue: null,
      minDate: null,
      maxDate: null,
      unique: new Map(),
    };

    baseRows.forEach((row) => {
      const value = row.values[colIdx];
      const displayValue = getDisplayValue(row, colIdx);
      const text = String(displayValue ?? "").trim();
      if (text === "") return;

      stats.nonEmpty += 1;
      stats.unique.set(text, (stats.unique.get(text) || 0) + 1);
      if (text.length > 60) stats.longTextCount += 1;

      if (typeof value === "string" && value.startsWith("=")) stats.formulaCount += 1;
      if (typeof value === "number") {
        stats.numericCount += 1;
        stats.minValue = stats.minValue == null ? value : Math.min(stats.minValue, value);
        stats.maxValue = stats.maxValue == null ? value : Math.max(stats.maxValue, value);
      }

      const asDate = parseDateFlexible(value);
      if (asDate instanceof Date) {
        stats.dateCount += 1;
        stats.minDate = !stats.minDate || asDate < stats.minDate ? asDate : stats.minDate;
        stats.maxDate = !stats.maxDate || asDate > stats.maxDate ? asDate : stats.maxDate;
      }

      if (typeof value === "string" && !(parseDateFlexible(value) instanceof Date) && !value.startsWith("=")) {
        stats.textCount += 1;
      } else if (!(value instanceof Date) && typeof value !== "number" && typeof value !== "string") {
        stats.textCount += 1;
      }
    });

    const emptyCount = totalRows - stats.nonEmpty;
    const uniqueCount = stats.unique.size;
    const topValues = Array.from(stats.unique.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(([label, count]) => ({ label, count }));
    const type = inferColumnProfileType(stats);
    const flags = [];
    let score = 0;

    if (stats.nonEmpty && stats.nonEmpty / totalRows <= 0.4) {
      flags.push("rzadka");
      score += 2;
    }
    if (type === "mixed") {
      flags.push("mixed");
      score += 3;
    }
    if (stats.longTextCount > 0) {
      flags.push("dlugie teksty");
      score += 1;
    }
    if (uniqueCount > Math.max(20, totalRows * 0.9) && type === "tekst") {
      flags.push("prawie same unikalne");
      score += 1;
    }
    if (stats.formulaCount > 0 && stats.formulaCount / stats.nonEmpty >= 0.8) {
      flags.push("kolumna formul");
      score += 1;
    }
    if (emptyCount === totalRows) {
      flags.push("pusta");
      score += 4;
    }

    return {
      header,
      colIdx,
      colAbs: currentStartCol + colIdx,
      nonEmpty: stats.nonEmpty,
      emptyCount,
      emptyPct: totalRows ? Math.round((emptyCount / totalRows) * 100) : 0,
      uniqueCount,
      type,
      topValues,
      rangeLabel: formatColumnProfileRange({
        type,
        minValue: stats.minValue,
        maxValue: stats.maxValue,
        minDate: stats.minDate,
        maxDate: stats.maxDate,
      }),
      flags,
      score,
    };
  });

  return profiles.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    if (a.emptyPct !== b.emptyPct) return b.emptyPct - a.emptyPct;
    return a.header.localeCompare(b.header, "pl");
  });
}

function renderColumnProfiles() {
  if (!columnProfilerEl) return;
  columnProfilerEl.replaceChildren();
  if (!currentColumnProfiles.length) {
    columnProfilerEl.appendChild(createEmptyInsight("Wczytaj arkusz, aby zobaczyc profil kolumn i szybkie sygnaly problemowosci."));
    return;
  }

  currentColumnProfiles.slice(0, 14).forEach((profile, index) => {
    const item = document.createElement("div");
    item.className = "column-profile-item";

    const top = document.createElement("div");
    top.className = "column-profile-top";

    const title = document.createElement("div");
    title.className = "column-profile-title";
    title.textContent = profile.header;

    const kind = document.createElement("div");
    kind.className = "column-profile-kind";
    kind.textContent = profile.type;

    top.appendChild(title);
    top.appendChild(kind);

    const meta = document.createElement("div");
    meta.className = "column-profile-meta";
    meta.textContent = `Kolumna ${XLSX.utils.encode_col(profile.colAbs)} • puste ${profile.emptyPct}% • unikalne ${profile.uniqueCount}`;

    const stats = document.createElement("div");
    stats.className = "column-profile-stats";
    if (profile.rangeLabel) {
      const rangeChip = document.createElement("div");
      rangeChip.className = "column-profile-chip";
      rangeChip.textContent = profile.rangeLabel;
      stats.appendChild(rangeChip);
    }
    profile.topValues.forEach((entry) => {
      const chip = document.createElement("div");
      chip.className = "column-profile-chip";
      chip.textContent = `${entry.label.slice(0, 24)}${entry.label.length > 24 ? "..." : ""} ×${entry.count}`;
      stats.appendChild(chip);
    });

    if (profile.flags.length) {
      const flags = document.createElement("div");
      flags.className = "column-profile-flags";
      profile.flags.forEach((flag) => {
        const badge = document.createElement("div");
        badge.className = "column-profile-flag";
        badge.textContent = flag;
        flags.appendChild(badge);
      });
      item.appendChild(top);
      item.appendChild(meta);
      if (stats.childNodes.length) item.appendChild(stats);
      item.appendChild(flags);
    } else {
      item.appendChild(top);
      item.appendChild(meta);
      if (stats.childNodes.length) item.appendChild(stats);
    }

    const actions = document.createElement("div");
    actions.className = "section-nav-actions";
    const btn = document.createElement("button");
    btn.className = "btn ghost btn-sm";
    btn.type = "button";
    btn.dataset.profileColIndex = String(profile.colIdx);
    btn.textContent = "Skocz do kolumny";
    actions.appendChild(btn);
    item.appendChild(actions);

    columnProfilerEl.appendChild(item);
  });
}

function focusColumnProfile(colIdx) {
  if (!Number.isFinite(colIdx)) return;
  const cells = theadEl.querySelectorAll(".guide-row .guide-cell");
  const cell = cells[colIdx];
  if (cell && tableWrapEl) {
    const targetLeft = Math.max(0, cell.offsetLeft - 64);
    tableWrapEl.scrollTo({ left: targetLeft, behavior: "smooth" });
    syncHorizontalScrollbar();
    return;
  }
  toast("Ta kolumna nie miesci sie jeszcze w aktualnym widoku tabeli", "info");
}

function renderInsights() {
  const data = collectSheetInsights();
  renderInsightList(
    workbookInsightsEl,
    data.workbookRows || [],
    "Wczytaj plik, aby zobaczyc metadane skoroszytu."
  );
  renderInsightList(
    sheetInsightsEl,
    data.rows || [],
    "Wczytaj arkusz, aby zobaczyc sygnaly jakosci danych i struktury."
  );
  renderInsightFlags(data.flags || []);
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

function rowMatchesEmptyMode(row, indexes, emptyMode) {
  if (!emptyMode || emptyMode === "all") return true;
  const resolvedIndexes = indexes && indexes.length ? indexes : row.values.map((_, i) => i);
  const emptyStates = resolvedIndexes.map((i) => {
    if (i >= row.values.length) return true;
    return getDisplayValue(row, i).trim().length === 0;
  });
  if (!emptyStates.length) return emptyMode === "any_empty";
  if (emptyMode === "any_empty") return emptyStates.some(Boolean);
  if (emptyMode === "all_empty") return emptyStates.every(Boolean);
  if (emptyMode === "any_non_empty") return emptyStates.some((isEmpty) => !isEmpty);
  if (emptyMode === "all_non_empty") return emptyStates.every((isEmpty) => !isEmpty);
  return true;
}

function combinePrimaryAndEmptyMatch(primaryMatched, emptyMatched, negated, hasPrimaryRule, hasEmptyRule) {
  if (!hasPrimaryRule && !hasEmptyRule) return true;
  if (hasPrimaryRule && negated) {
    if (hasEmptyRule) return !primaryMatched && emptyMatched;
    return !primaryMatched;
  }
  if (hasPrimaryRule && !negated) {
    if (hasEmptyRule) return primaryMatched && emptyMatched;
    return primaryMatched;
  }
  if (hasEmptyRule && negated) return !emptyMatched;
  if (hasEmptyRule) return emptyMatched;
  return true;
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
    const emptyMode = criterion.emptyMode || "all";
    const hasQuery = !!query;
    const hasEmptyRule = emptyMode !== "all";
    if (!hasQuery && !hasEmptyRule) continue;

    let textMatched = !hasQuery;
    for (const i of criterion.indexes) {
      if (!hasQuery) break;
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
      if (criterion.mode === "Równa się" && candidates.some((c) => c === query)) textMatched = true;
      if (criterion.mode === "Zaczyna się" && candidates.some((c) => c.startsWith(query))) textMatched = true;
      if (criterion.mode === "Zawiera" && candidates.some((c) => c.includes(query))) textMatched = true;
      if (textMatched) break;
    }
    const emptyMatched = rowMatchesEmptyMode(row, criterion.indexes, emptyMode);
    const matched = combinePrimaryAndEmptyMatch(textMatched, emptyMatched, criterion.negated, hasQuery, hasEmptyRule);
    if (!matched) return false;
  }

  return true;
}

function rowMatchesDateFilter(row, filter) {
  const indexes = filter.indexes || [];
  const dateRange = filter.range || { from: null, to: null };
  const hasRange = !!(dateRange.from || dateRange.to);
  const emptyMode = filter.emptyMode || "all";
  const hasEmptyRule = emptyMode !== "all";
  if (!hasRange && !hasEmptyRule) return true;

  let rangeMatched = !hasRange;
  if (hasRange) {
    rangeMatched = false;
    for (const idx of indexes) {
      if (idx >= row.values.length) continue;
      const raw = row.rawValues ? row.rawValues[idx] : row.values[idx];
      const d = parseDateFlexible(raw ?? getDisplayValue(row, idx));
      if (!d) continue;
      if (dateRange.from && d < dateRange.from) continue;
      if (dateRange.to && d > dateRange.to) continue;
      rangeMatched = true;
      break;
    }
  }

  const emptyMatched = rowMatchesEmptyMode(row, indexes, emptyMode);
  return combinePrimaryAndEmptyMatch(rangeMatched, emptyMatched, filter.negated, hasRange, hasEmptyRule);
}

function resolveIndexes(headers, selected) {
  if (!selected.size) return headers.map((_, i) => i);
  return headers.map((h, i) => (selected.has(h) ? i : -1)).filter((i) => i >= 0);
}

function compareSortValues(av, bv) {
  const ad = parseDateFlexible(av);
  const bd = parseDateFlexible(bv);
  if (ad && bd) return ad - bd;
  if (typeof av === "number" && typeof bv === "number") return av - bv;
  return String(av || "").localeCompare(String(bv || ""), "pl");
}

function normalizeSortState() {
  multiSortState = multiSortState
    .filter((rule) => rule && rule.col && currentHeaders.includes(rule.col))
    .map((rule) => ({ col: rule.col, dir: rule.dir === "desc" ? "desc" : "asc" }));
  const primary = multiSortState[0] || null;
  sortState = primary ? { col: primary.col, dir: primary.dir } : { col: "", dir: "asc" };
}

function setPrimarySort(col, dir = "asc") {
  if (!col) {
    multiSortState = [];
    normalizeSortState();
    return;
  }
  const next = [{ col, dir: dir === "desc" ? "desc" : "asc" }];
  multiSortState.forEach((rule) => {
    if (rule.col === col) return;
    next.push(rule);
  });
  multiSortState = next;
  normalizeSortState();
}

function populateSortColumnSelect() {
  if (!sortColumnSelectEl) return;
  sortColumnSelectEl.replaceChildren();
  if (!currentHeaders.length) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = "Najpierw wczytaj arkusz";
    sortColumnSelectEl.appendChild(opt);
    sortColumnSelectEl.disabled = true;
    if (addSortRuleBtn) addSortRuleBtn.disabled = true;
    return;
  }
  sortColumnSelectEl.disabled = false;
  if (addSortRuleBtn) addSortRuleBtn.disabled = false;
  currentHeaders.forEach((header) => {
    const opt = document.createElement("option");
    opt.value = header;
    opt.textContent = header;
    sortColumnSelectEl.appendChild(opt);
  });
}

function renderSortRules() {
  if (!sortRulesListEl) return;
  sortRulesListEl.replaceChildren();
  if (!multiSortState.length) {
    sortRulesListEl.appendChild(createEmptyInsight("Brak aktywnych sortowan. Kliknij naglowek tabeli albo dodaj regule tutaj."));
    return;
  }
  multiSortState.forEach((rule, index) => {
    const item = document.createElement("div");
    item.className = "sort-rule-item";

    const label = document.createElement("div");
    label.className = "sort-rule-label";
    label.textContent = `${index + 1}. ${rule.col}`;

    const dir = document.createElement("div");
    dir.className = "sort-rule-dir";
    dir.textContent = rule.dir === "asc" ? "Rosnąco" : "Malejąco";

    const actions = document.createElement("div");
    actions.className = "sort-rule-actions";

    const upBtn = document.createElement("button");
    upBtn.className = "btn ghost btn-sm";
    upBtn.type = "button";
    upBtn.dataset.sortAction = "up";
    upBtn.dataset.sortIndex = String(index);
    upBtn.textContent = "Góra";
    upBtn.disabled = index === 0;

    const downBtn = document.createElement("button");
    downBtn.className = "btn ghost btn-sm";
    downBtn.type = "button";
    downBtn.dataset.sortAction = "down";
    downBtn.dataset.sortIndex = String(index);
    downBtn.textContent = "Dół";
    downBtn.disabled = index === multiSortState.length - 1;

    const toggleBtn = document.createElement("button");
    toggleBtn.className = "btn ghost btn-sm";
    toggleBtn.type = "button";
    toggleBtn.dataset.sortAction = "toggle";
    toggleBtn.dataset.sortIndex = String(index);
    toggleBtn.textContent = "Zmień kierunek";

    const removeBtn = document.createElement("button");
    removeBtn.className = "btn ghost btn-sm";
    removeBtn.type = "button";
    removeBtn.dataset.sortAction = "remove";
    removeBtn.dataset.sortIndex = String(index);
    removeBtn.textContent = "Usuń";

    actions.appendChild(upBtn);
    actions.appendChild(downBtn);
    actions.appendChild(toggleBtn);
    actions.appendChild(removeBtn);

    item.appendChild(label);
    item.appendChild(dir);
    item.appendChild(actions);
    sortRulesListEl.appendChild(item);
  });
}

function loadSortPresets() {
  try {
    const raw = localStorage.getItem(SORT_PRESETS_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveSortPresets(presets) {
  localStorage.setItem(SORT_PRESETS_KEY, JSON.stringify(presets));
}

function renderSortPresets() {
  if (!sortPresetSelectEl) return;
  const presets = loadSortPresets();
  sortPresetSelectEl.replaceChildren();
  if (!presets.length) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = "Brak zapisanych presetów";
    sortPresetSelectEl.appendChild(opt);
    return;
  }
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Wybierz preset";
  sortPresetSelectEl.appendChild(placeholder);
  presets.forEach((preset) => {
    const opt = document.createElement("option");
    opt.value = preset.name;
    opt.textContent = preset.name;
    sortPresetSelectEl.appendChild(opt);
  });
}

function applyCurrentSort() {
  applyFilters();
  sortRows();
  renderActiveTable();
  renderInsights();
  renderSheetInspectorSummary();
  renderColumnProfiles();
  renderSections();
  renderRepeatingBlocks();
  renderDurationAnalysis();
  renderAggregationWorkbench();
  renderFormulaWorkbench();
}

function applyFilters() {
  const criteria = [
    {
      query: (searchQueryEl.value || "").trim().toLowerCase(),
      mode: filterModeEl.value,
      indexes: resolveIndexes(currentHeaders, columnSelections.filter1),
      emptyMode: filterEmptyModeEl.value,
      negated: filterNegateEl.checked,
    },
    {
      query: (searchQuery2El.value || "").trim().toLowerCase(),
      mode: filterMode2El.value,
      indexes: resolveIndexes(currentHeaders, columnSelections.filter2),
      emptyMode: filterEmptyMode2El.value,
      negated: filterNegate2El.checked,
    },
  ];

  const dateFilter = {
    indexes: resolveIndexes(currentHeaders, columnSelections.date),
    range: getDateRange(),
    emptyMode: dateEmptyModeEl.value,
    negated: dateNegateEl.checked,
  };
  const onlyNonEmpty = onlyNonEmptyEl.checked;

  viewRows = baseRows.filter((row) => {
    if (!rowMatchesTextFilter(row, criteria, onlyNonEmpty)) return false;
    if (!rowMatchesDateFilter(row, dateFilter)) return false;
    return true;
  });
}

function sortRows() {
  normalizeSortState();
  if (!multiSortState.length) return;
  viewRows.sort((a, b) => {
    for (const rule of multiSortState) {
      const idx = currentHeaders.indexOf(rule.col);
      if (idx < 0) continue;
      const av = a.rawValues ? a.rawValues[idx] : a.values[idx];
      const bv = b.rawValues ? b.rawValues[idx] : b.values[idx];
      const cmp = compareSortValues(av, bv);
      if (cmp !== 0) return rule.dir === "desc" ? -cmp : cmp;
    }
    return 0;
  });
}

function sortRowsForHeaders(rows, headers) {
  normalizeSortState();
  if (!multiSortState.length || !Array.isArray(rows) || !Array.isArray(headers)) return;
  rows.sort((a, b) => {
    for (const rule of multiSortState) {
      const idx = headers.indexOf(rule.col);
      if (idx < 0) continue;
      const av = a.rawValues ? a.rawValues[idx] : a.values[idx];
      const bv = b.rawValues ? b.rawValues[idx] : b.values[idx];
      const cmp = compareSortValues(av, bv);
      if (cmp !== 0) return rule.dir === "desc" ? -cmp : cmp;
    }
    return 0;
  });
}

function getActiveRepeatingGroup() {
  return Array.isArray(currentRepeatingBlocks) && currentRepeatingBlocks.length ? currentRepeatingBlocks[0] : null;
}

function canUseLongView() {
  const group = getActiveRepeatingGroup();
  return !!(group && Array.isArray(group.blocks) && group.blocks.length >= 2);
}

function updateWideLongToggle() {
  if (!wideLongToggleEl) return;
  const available = canUseLongView();
  if (!available) {
    tableViewMode = "wide";
  }
  wideLongToggleEl.classList.toggle("hidden", !available);
  wideLongToggleEl.setAttribute("aria-hidden", available ? "false" : "true");
  wideLongToggleEl.setAttribute("aria-pressed", tableViewMode === "long" ? "true" : "false");
  wideLongToggleEl.textContent = tableViewMode === "long" ? "Widok klasyczny" : "Wide-to-Long";
  wideLongToggleEl.title = tableViewMode === "long"
    ? "Wroc do klasycznego ukladu arkusza"
    : "Przelacz wykryte bloki kolumn na dlugi widok analityczny";
}

function buildWideDisplayModelFromRows(rows, options = {}) {
  const headers = Array.isArray(options.headers) ? options.headers.slice() : currentHeaders.slice();
  const startCol = Number.isFinite(options.startCol) ? options.startCol : currentStartCol;
  const headerRow = Number.isFinite(options.headerRow) ? options.headerRow : currentHeaderRow;
  return {
    mode: "wide",
    headers,
    rows: rows.slice(),
    guideLabels: headers.map((_, i) => XLSX.utils.encode_col(i + startCol)),
    headerRowLabel: String(headerRow),
    rowHeadFormatter: (row) => String((row?.rowIndex0 ?? 0) + 1),
    editable: true,
  };
}

function buildWideDisplayModel() {
  return buildWideDisplayModelFromRows(viewRows);
}

function buildLongViewModelFromRows(rows, group = getActiveRepeatingGroup(), options = {}) {
  if (!group || !Array.isArray(group.blocks) || !group.blocks.length) return buildWideDisplayModelFromRows(rows);

  const firstBlock = group.blocks[0];
  const prefixCount = Math.max(0, Number(group.prefixCount) || 0);
  const sourceHeaders = Array.isArray(options.headers) ? options.headers.slice() : currentHeaders.slice();
  const headerRow = Number.isFinite(options.headerRow) ? options.headerRow : currentHeaderRow;
  const prefixHeaders = sourceHeaders.slice(0, prefixCount);
  const repeatedHeaders = firstBlock.headers.map((header) => parseRepeatedHeader(header)?.base || cleanSectionLabel(header) || header);
  const headers = [...prefixHeaders, "Nr bloku", "Blok", ...repeatedHeaders];
  const nextRows = [];

  rows.forEach((row) => {
    group.blocks.forEach((block, blockIndex) => {
      const blockValues = row.values.slice(block.startIndex, block.endIndex + 1);
      const blockDisplay = blockValues.map((_, idx) => getDisplayValue(row, block.startIndex + idx));
      const hasMeaningfulValue = blockDisplay.some((value) => String(value ?? "").trim() !== "");
      if (!hasMeaningfulValue) return;

      const prefixValues = row.values.slice(0, prefixCount);
      const prefixDisplay = prefixValues.map((_, idx) => getDisplayValue(row, idx));
      const values = [...prefixValues, blockIndex + 1, block.label, ...blockValues];
      const display = [...prefixDisplay, String(blockIndex + 1), block.label, ...blockDisplay];

      nextRows.push({
        values,
        rawValues: values.slice(),
        display,
        rowIndex0: row.rowIndex0,
        sourceRowIndex0: row.rowIndex0,
        sourceBlockIndex: blockIndex,
        sourceBlockLabel: block.label,
        isLongViewRow: true,
      });
    });
  });

  return {
    mode: "long",
    headers,
    rows: nextRows,
    guideLabels: headers.map((_, idx) => `${idx + 1}`),
    headerRowLabel: `${headerRow} -> long`,
    rowHeadFormatter: (row) => `${(row?.sourceRowIndex0 ?? row?.rowIndex0 ?? 0) + 1}.${(row?.sourceBlockIndex ?? 0) + 1}`,
    editable: false,
  };
}

function buildLongViewModel() {
  return buildLongViewModelFromRows(viewRows);
}

function getRowSelectionKey(row) {
  if (!row) return "";
  if (row.isLongViewRow) {
    return `long:${row.sourceRowIndex0 ?? row.rowIndex0}:${row.sourceBlockIndex ?? 0}`;
  }
  return `wide:${row.rowIndex0 ?? ""}`;
}

function findCellElement(cellState) {
  if (!cellState) return null;
  return tbodyEl.querySelector(
    `tr[data-row-key="${CSS.escape(cellState.rowKey)}"] td[data-col-index="${cellState.colIndex0}"]`
  );
}

function findFocusedRowElement() {
  if (!focusedCellState) return null;
  return tbodyEl.querySelector(`tr[data-row-key="${CSS.escape(focusedCellState.rowKey)}"]`);
}

function syncFocusedCellInDom(options = {}) {
  tbodyEl.querySelectorAll("tr.row-focused").forEach((row) => row.classList.remove("row-focused"));
  const rowEl = findFocusedRowElement();
  if (!rowEl) {
    if (options.clearMissing !== false) focusedCellState = null;
    return null;
  }
  rowEl.classList.add("row-focused");
  const cell = findCellElement(focusedCellState);
  if (options.scroll) {
    (cell || rowEl).scrollIntoView({ block: "nearest", inline: "nearest" });
  }
  return rowEl;
}

function syncSelectedCellInDom(options = {}) {
  tbodyEl.querySelectorAll("td.cell-selected").forEach((cell) => cell.classList.remove("cell-selected"));
  const cell = findCellElement(selectedCellState);
  if (!cell) {
    if (options.clearMissing !== false) selectedCellState = null;
    return null;
  }
  cell.classList.add("cell-selected");
  if (options.scroll) {
    cell.scrollIntoView({ block: "nearest", inline: "nearest" });
  }
  return cell;
}

function setFocusedCell(rowKey, colIndex0, options = {}) {
  if (!rowKey || !Number.isFinite(colIndex0) || colIndex0 < 0) {
    focusedCellState = null;
    syncFocusedCellInDom({ clearMissing: false });
    return;
  }
  focusedCellState = { rowKey, colIndex0 };
  syncFocusedCellInDom(options);
}

function setSelectedCell(rowKey, colIndex0, options = {}) {
  if (!rowKey || !Number.isFinite(colIndex0) || colIndex0 < 0) {
    selectedCellState = null;
    syncSelectedCellInDom({ clearMissing: false });
    return;
  }
  selectedCellState = { rowKey, colIndex0 };
  syncSelectedCellInDom(options);
}

function moveFocusedCell(rowDelta, colDelta) {
  if (!focusedCellState || !currentDisplayModel?.rows?.length || !currentDisplayModel?.headers?.length) return false;
  const rowIndex = currentDisplayModel.rows.findIndex((row) => getRowSelectionKey(row) === focusedCellState.rowKey);
  if (rowIndex < 0) {
    focusedCellState = null;
    return false;
  }
  const nextRowIndex = Math.max(0, Math.min(currentDisplayModel.rows.length - 1, rowIndex + rowDelta));
  const nextColIndex = Math.max(0, Math.min(currentDisplayModel.headers.length - 1, focusedCellState.colIndex0 + colDelta));
  const nextRow = currentDisplayModel.rows[nextRowIndex];
  setFocusedCell(getRowSelectionKey(nextRow), nextColIndex, { scroll: true });
  return true;
}

function moveSelectedCell(rowDelta, colDelta) {
  if (!selectedCellState || !currentDisplayModel?.rows?.length || !currentDisplayModel?.headers?.length) return false;
  const rowIndex = currentDisplayModel.rows.findIndex((row) => getRowSelectionKey(row) === selectedCellState.rowKey);
  if (rowIndex < 0) {
    selectedCellState = null;
    return false;
  }
  const nextRowIndex = Math.max(0, Math.min(currentDisplayModel.rows.length - 1, rowIndex + rowDelta));
  const nextColIndex = Math.max(0, Math.min(currentDisplayModel.headers.length - 1, selectedCellState.colIndex0 + colDelta));
  const nextRow = currentDisplayModel.rows[nextRowIndex];
  setSelectedCell(getRowSelectionKey(nextRow), nextColIndex, { scroll: true });
  return true;
}

function shouldIgnoreTableArrowNavigation() {
  const active = document.activeElement;
  if (!active) return false;
  const tag = String(active.tagName || "").toLowerCase();
  return active.isContentEditable || ["input", "textarea", "select", "button"].includes(tag);
}

function getAggregationSourceRows(scopeMode) {
  return scopeMode === "all" ? baseRows.slice() : viewRows.slice();
}

function getAggregationHeaderCandidateRows() {
  const candidates = new Set([currentHeaderRow]);
  currentSections.forEach((section) => {
    if (section?.action === "set-header" && Number.isFinite(section.headerRow)) {
      candidates.add(section.headerRow);
    }
  });
  for (let row = Math.max(1, currentHeaderRow - 3); row <= currentHeaderRow + 4; row++) {
    candidates.add(row);
  }
  return Array.from(candidates)
    .filter((row) => Number.isFinite(row) && row > 0)
    .sort((a, b) => a - b);
}

function getAggregationHeaderSourceData(headerRow = currentHeaderRow, scopeMode = aggregationWorkbenchState.scopeMode) {
  if (!workbook || !currentSheetName || headerRow === currentHeaderRow) {
    return {
      headerRow: currentHeaderRow,
      rows: getAggregationSourceRows(scopeMode),
      headers: currentHeaders.slice(),
      startCol: currentStartCol,
      group: getActiveRepeatingGroup(),
      longAvailable: canUseLongView(),
      helperMode: false,
    };
  }

  const sheet = workbook.Sheets[currentSheetName];
  if (!sheet) {
    return {
      headerRow: currentHeaderRow,
      rows: getAggregationSourceRows(scopeMode),
      headers: currentHeaders.slice(),
      startCol: currentStartCol,
      group: getActiveRepeatingGroup(),
      longAvailable: canUseLongView(),
      helperMode: false,
    };
  }

  try {
    const data = buildRows(sheet, headerRow, workbook);
    const rows = markSubheaderRows(data.rows.slice());
    const visibleRowIndexes = scopeMode === "filtered"
      ? new Set(viewRows.map((row) => row.rowIndex0))
      : null;
    const scopedRows = visibleRowIndexes
      ? rows.filter((row) => visibleRowIndexes.has(row.rowIndex0))
      : rows;
    const groups = detectRepeatingBlocks(sheet, headerRow, data);
    const group = Array.isArray(groups) && groups.length ? groups[0] : null;
    return {
      headerRow,
      rows: scopedRows,
      headers: data.headers.slice(),
      startCol: data.startCol || 0,
      group,
      longAvailable: !!(group && Array.isArray(group.blocks) && group.blocks.length >= 2),
      helperMode: headerRow !== currentHeaderRow,
    };
  } catch {
    return {
      headerRow: currentHeaderRow,
      rows: getAggregationSourceRows(scopeMode),
      headers: currentHeaders.slice(),
      startCol: currentStartCol,
      group: getActiveRepeatingGroup(),
      longAvailable: canUseLongView(),
      helperMode: false,
    };
  }
}

function collectAggregationContextForHeaderRow(headerRow, sourceMode = aggregationWorkbenchState.sourceMode, scopeMode = aggregationWorkbenchState.scopeMode) {
  const source = getAggregationHeaderSourceData(headerRow, scopeMode);
  const normalizedSource = sourceMode === "auto"
    ? (source.longAvailable ? "long" : "wide")
    : sourceMode;
  const model = normalizedSource === "long" && source.longAvailable
    ? buildLongViewModelFromRows(source.rows, source.group, {
      headers: source.headers,
      headerRow: source.headerRow,
      startCol: source.startCol,
    })
    : buildWideDisplayModelFromRows(source.rows, {
      headers: source.headers,
      headerRow: source.headerRow,
      startCol: source.startCol,
    });
  const profiles = collectAggregationProfiles(model);
  const groupOptions = resolveAggregationGroupOptions(profiles);
  const measures = detectAggregationMeasureCandidates(model, profiles);
  return {
    ...source,
    model,
    profiles,
    groupOptions,
    measures,
    resolvedSourceMode: normalizedSource,
  };
}

function scoreAggregationContext(context) {
  if (!context?.model?.rows?.length) return -1;
  const dateRangeBonus = context.measures.some((candidate) => candidate.measureType === "date_range") ? 30 : 0;
  const textGroupBonus = context.groupOptions.some((option) => /\b(imie|nazwisko|osoba|pracownik|owner|assignee)\b/.test(normalizeAnalysisKey(option.label))) ? 12 : 0;
  const currentHeaderBonus = context.headerRow === currentHeaderRow ? 4 : 0;
  return (context.groupOptions.length * 8)
    + (context.measures.length * 10)
    + dateRangeBonus
    + textGroupBonus
    + currentHeaderBonus
    + Math.min(context.model.rows.length, 500) * 0.02;
}

function isValidAggregationHeaderRow(headerRow) {
  if (!Number.isFinite(headerRow) || headerRow < 1) return false;
  if (!workbook || !currentSheetName) return false;
  const sheet = workbook.Sheets[currentSheetName];
  if (!sheet) return false;
  try {
    const data = buildRows(sheet, headerRow, workbook);
    return Array.isArray(data?.headers) && data.headers.length > 0 && Array.isArray(data?.rows) && data.rows.length > 0;
  } catch {
    return false;
  }
}

function getAggregationSourceModel(sourceMode = aggregationWorkbenchState.sourceMode, scopeMode = aggregationWorkbenchState.scopeMode) {
  return collectAggregationContextForHeaderRow(currentHeaderRow, sourceMode, scopeMode).model;
}

function getDisplayModel() {
  if (tableViewMode === "long" && canUseLongView()) {
    return buildLongViewModel();
  }
  return buildWideDisplayModel();
}

function renderActiveTable() {
  currentDisplayModel = getDisplayModel();
  sortRowsForHeaders(currentDisplayModel.rows, currentDisplayModel.headers);
  renderTable(currentDisplayModel);
  updateWideLongToggle();
}

function updateSortControls() {
  if (!resetSortBtn) return;
  normalizeSortState();
  const active = multiSortState.length > 0;
  resetSortBtn.classList.toggle("hidden", !active);
  resetSortBtn.setAttribute("aria-hidden", active ? "false" : "true");
  renderSortRules();
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

function renderTable(modelOrHeaders, maybeRows) {
  const model = Array.isArray(modelOrHeaders)
    ? {
        mode: "wide",
        headers: modelOrHeaders,
        rows: Array.isArray(maybeRows) ? maybeRows : [],
        guideLabels: modelOrHeaders.map((_, i) => XLSX.utils.encode_col(i + currentStartCol)),
        headerRowLabel: String(currentHeaderRow),
        rowHeadFormatter: (row) => String((row?.rowIndex0 ?? 0) + 1),
        editable: true,
      }
    : (modelOrHeaders || { headers: [], rows: [] });
  const headers = Array.isArray(model.headers) ? model.headers : [];
  const rows = Array.isArray(model.rows) ? model.rows : [];

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
  theadEl.replaceChildren();
  tbodyEl.replaceChildren();

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
  tableEl.replaceChildren();
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
    th.textContent = Array.isArray(model.guideLabels) && model.guideLabels[i] ? model.guideLabels[i] : XLSX.utils.encode_col(i + currentStartCol);
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
  rowHead.textContent = model.headerRowLabel || String(currentHeaderRow);
  headRow.appendChild(rowHead);
  const headerMergeLayout = model.mode === "wide" ? computeHeaderMergeLayout(headers.length) : null;
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
        setPrimarySort(h, sortState.dir === "asc" ? "desc" : "asc");
      } else {
        setPrimarySort(h, "asc");
      }
      if (model.mode === "wide") {
        sortRows();
      }
      updateSortControls();
      renderActiveTable();
    });

    const primarySort = multiSortState[0];
    if (primarySort && primarySort.col === h) {
      const arrow = document.createElement("span");
      arrow.className = "sort-arrow";
      arrow.textContent = primarySort.dir === "asc" ? "▲" : "▼";
      th.appendChild(arrow);
    }

    headRow.appendChild(th);
  }
  theadEl.appendChild(headRow);

  const limit = Math.max(1, parseInt(maxRowsEl.value || "200", 10));
  const rowsShown = rows.slice(0, limit);
  const mergeLayout = model.mode === "wide" ? computeMergeLayout(rowsShown, headers.length) : null;

  rowsShown.forEach((row, rowPos) => {
    const tr = document.createElement("tr");
    tr.dataset.rowKey = getRowSelectionKey(row);
    if (focusedCellState && focusedCellState.rowKey === tr.dataset.rowKey) tr.classList.add("row-focused");
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
    rowHead.textContent = model.rowHeadFormatter ? model.rowHeadFormatter(row, rowPos) : String(row.rowIndex0 + 1);
    tr.appendChild(rowHead);
    row.values.forEach((v, i) => {
      const mergeKey = `${rowPos}:${i}`;
      if (mergeLayout && mergeLayout.covered.has(mergeKey)) return;
      const td = document.createElement("td");
      const displayValue = getDisplayValue(row, i);
      td.textContent = displayValue;
      td.dataset.fullText = displayValue;
      td.dataset.colIndex = String(i);
      if (selectedCellState && selectedCellState.rowKey === tr.dataset.rowKey && selectedCellState.colIndex0 === i) {
        td.classList.add("cell-selected");
      }

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

  const modeLabel = model.mode === "long" ? " • tryb long" : "";
  setStatus(`Wierszy: ${rows.length} (pokazano: ${Math.min(rows.length, limit)})${modeLabel}`);
  syncFocusedCellInDom({ clearMissing: true });
  syncSelectedCellInDom({ clearMissing: true });
  syncHorizontalScrollbar();
  applyZoom();
}

function buildRows(sheet, headerRow, wb) {
  const originalRange = XLSX.utils.decode_range(sheet["!ref"]);
  const range = computeEffectiveSheetRange(sheet, headerRow);
  const colMeta = sheet["!cols"] || [];
  const rowMeta = sheet["!rows"] || [];
  const merges = Array.isArray(sheet["!merges"]) ? sheet["!merges"] : [];
  const rawHeaders = [];
  const headerStyles = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: headerRow - 1, c })];
    const v = cell ? cell.v : null;
    rawHeaders.push(v ? String(v).trim() : XLSX.utils.encode_col(c));
    headerStyles.push(wb ? extractCellStyle(cell, wb) : null);
  }
  const headers = makeHeadersUnique(rawHeaders);
  const duplicateHeaderCount = rawHeaders.length - new Set(rawHeaders).size;
  const rows = [];
  let formulaCount = 0;
  let formulaMissingResultCount = 0;
  let commentCount = 0;
  let hyperlinkCount = 0;
  for (let r = headerRow; r <= range.e.r; r++) {
    const values = [];
    const display = [];
    const cellStyles = [];
    let any = false;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      let v = cell ? cell.v : null;
      let shown = cell && cell.w ? String(cell.w) : toDisplay(v);
      if (cell && cell.f) {
        formulaCount += 1;
        if (cell.v == null && cell.w == null) formulaMissingResultCount += 1;
      }
      if (cell && Array.isArray(cell.c) && cell.c.length) commentCount += 1;
      if (cell && cell.l && (cell.l.Target || cell.l.target)) hyperlinkCount += 1;
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
    stats: {
      duplicateHeaderCount,
      formulaCount,
      formulaMissingResultCount,
      commentCount,
      hyperlinkCount,
      mergeRegions: merges.length,
      mergedCells: merges.reduce((sum, merge) => sum + ((merge.e.r - merge.s.r + 1) * (merge.e.c - merge.s.c + 1)), 0),
      hiddenColumns: colMeta.filter((meta) => meta && meta.hidden).length,
      hiddenRows: rowMeta.filter((meta) => meta && meta.hidden).length,
      trimmedColumns: Math.max(0, (originalRange.e.c - originalRange.s.c + 1) - (range.e.c - range.s.c + 1)),
      trimmedRows: Math.max(0, (originalRange.e.r - originalRange.s.r + 1) - (range.e.r - range.s.r + 1)),
    },
    colWidths: headers.map((_, idx) => colMeta[range.s.c + idx] || null),
    rowHeights: rowMeta,
  };
}

function extractFormulaFunctionName(formulaText) {
  const text = String(formulaText || "").replace(/^=/, "").trim();
  const match = text.match(/^([A-Z_][A-Z0-9\._]*)\s*\(/i);
  return match ? match[1].toUpperCase() : "INNE";
}

function collectFormulaEntries(sheet, data, headerRow) {
  if (!sheet || !data || !Array.isArray(data.headers) || !data.headers.length) return [];
  const entries = [];
  const range = computeEffectiveSheetRange(sheet, headerRow);

  Object.keys(sheet).forEach((key) => {
    if (!key || key[0] === "!") return;
    const cell = sheet[key];
    if (!cell || !cell.f) return;
    const ref = XLSX.utils.decode_cell(key);
    if (ref.r < range.s.r || ref.r > range.e.r || ref.c < range.s.c || ref.c > range.e.c) return;

    const formulaText = `=${cell.f}`;
    const functionName = extractFormulaFunctionName(formulaText);
    const resultText = cell.w != null ? String(cell.w) : toDisplay(cell.v);
    const missingResult = cell.v == null && cell.w == null;
    const hasError = String(resultText || "").trim().startsWith("#");
    const colIdx = ref.c - data.startCol;
    const inTable = ref.r >= headerRow && colIdx >= 0 && colIdx < data.headers.length;

    entries.push({
      address: key,
      formulaText,
      functionName,
      resultText,
      missingResult,
      hasError,
      rowIndex0: ref.r,
      colAbs: ref.c,
      colIdx,
      inTable,
      header: inTable ? data.headers[colIdx] : XLSX.utils.encode_col(ref.c),
    });
  });

  return entries.sort((a, b) => {
    if (a.missingResult !== b.missingResult) return a.missingResult ? -1 : 1;
    if (a.hasError !== b.hasError) return a.hasError ? -1 : 1;
    if (a.functionName !== b.functionName) return a.functionName.localeCompare(b.functionName, "pl");
    return a.address.localeCompare(b.address, "pl");
  });
}

function getFilteredFormulaEntries() {
  const search = String(formulaSearchEl?.value || "").trim().toLowerCase();
  const filter = formulaFilterEl?.value || "all";
  const functionFilter = String(formulaFunctionFilterEl?.value || "").trim().toUpperCase();
  return currentFormulaEntries.filter((entry) => {
    if (filter === "missing" && !entry.missingResult) return false;
    if (filter === "error" && !entry.hasError) return false;
    if (functionFilter && entry.functionName !== functionFilter) return false;
    if (!search) return true;
    const haystack = [
      entry.address,
      entry.header,
      entry.functionName,
      entry.formulaText,
      entry.resultText,
    ].join(" ").toLowerCase();
    return haystack.includes(search);
  });
}

function renderFormulaFunctionFilter() {
  if (!formulaFunctionFilterEl) return;
  const previous = formulaFunctionFilterEl.value;
  const names = Array.from(new Set(currentFormulaEntries.map((entry) => entry.functionName))).sort((a, b) => a.localeCompare(b, "pl"));
  formulaFunctionFilterEl.replaceChildren();

  const allOpt = document.createElement("option");
  allOpt.value = "";
  allOpt.textContent = "Wszystkie funkcje";
  formulaFunctionFilterEl.appendChild(allOpt);

  names.forEach((name) => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    formulaFunctionFilterEl.appendChild(opt);
  });

  formulaFunctionFilterEl.value = names.includes(previous) ? previous : "";
}

function truncateFormulaPreview(text, maxLength = 120) {
  const raw = String(text || "").trim();
  if (raw.length <= maxLength) return raw;
  const head = Math.max(36, Math.floor(maxLength * 0.55));
  const tail = Math.max(18, maxLength - head - 3);
  return `${raw.slice(0, head)}...${raw.slice(-tail)}`;
}

function formatFormulaAddressSample(entries, limit = 4) {
  const sample = entries.slice(0, limit).map((entry) => entry.address);
  if (entries.length <= limit) return sample.join(", ");
  return `${sample.join(", ")} +${entries.length - limit}`;
}

function aggregateFormulaEntries(entries) {
  const groups = new Map();
  entries.forEach((entry) => {
    const key = [
      entry.functionName,
      entry.formulaText,
      entry.header,
      entry.missingResult ? "1" : "0",
      entry.hasError ? "1" : "0",
      entry.inTable ? "1" : "0",
    ].join("||");
    const existing = groups.get(key);
    if (existing) {
      existing.entries.push(entry);
      return;
    }
    groups.set(key, {
      key,
      functionName: entry.functionName,
      formulaText: entry.formulaText,
      header: entry.header,
      missingResult: entry.missingResult,
      hasError: entry.hasError,
      inTable: entry.inTable,
      resultText: entry.resultText,
      entries: [entry],
      firstEntry: entry,
    });
  });
  return Array.from(groups.values()).sort((a, b) => {
    if (a.missingResult !== b.missingResult) return a.missingResult ? -1 : 1;
    if (a.hasError !== b.hasError) return a.hasError ? -1 : 1;
    if (b.entries.length !== a.entries.length) return b.entries.length - a.entries.length;
    if (a.functionName !== b.functionName) return a.functionName.localeCompare(b.functionName, "pl");
    return a.firstEntry.address.localeCompare(b.firstEntry.address, "pl");
  });
}

function renderFormulaWorkbench() {
  if (!formulaWorkbenchSummaryEl || !formulaWorkbenchListEl) return;
  formulaWorkbenchSummaryEl.replaceChildren();
  formulaWorkbenchListEl.replaceChildren();
  renderFormulaFunctionFilter();

  if (!currentHeaders.length || !currentFormulaEntries.length) {
    renderInsightList(
      formulaWorkbenchSummaryEl,
      [],
      "Aktualny arkusz nie ma wykrytych formuł albo nie został jeszcze wczytany."
    );
    formulaWorkbenchListEl.appendChild(createEmptyInsight("Brak formuł do pokazania dla aktualnego arkusza."));
    return;
  }

  const filtered = getFilteredFormulaEntries();
  const grouped = aggregateFormulaEntries(filtered);
  const functionCounts = new Map();
  currentFormulaEntries.forEach((entry) => {
    functionCounts.set(entry.functionName, (functionCounts.get(entry.functionName) || 0) + 1);
  });
  const topFunction = Array.from(functionCounts.entries()).sort((a, b) => b[1] - a[1])[0];
  const summaryItems = [
    { label: "Formuły", value: String(currentFormulaEntries.length) },
    {
      label: "Bez wyniku",
      value: String(currentFormulaEntries.filter((entry) => entry.missingResult).length),
      tone: currentFormulaEntries.some((entry) => entry.missingResult) ? "warning" : "",
    },
    {
      label: "Z błędem",
      value: String(currentFormulaEntries.filter((entry) => entry.hasError).length),
      tone: currentFormulaEntries.some((entry) => entry.hasError) ? "warning" : "",
    },
    {
      label: "Top funkcja",
      value: topFunction ? `${topFunction[0]} ×${topFunction[1]}` : "Brak",
      tone: topFunction ? "info" : "",
    },
    {
      label: "Widoczne po filtrze",
      value: String(filtered.length),
      tone: filtered.length !== currentFormulaEntries.length ? "info" : "",
    },
    {
      label: "Grupy",
      value: String(grouped.length),
      tone: grouped.length < filtered.length ? "info" : "",
    },
  ];

  renderInsightList(formulaWorkbenchSummaryEl, summaryItems, "Brak podsumowania formuł.");

  if (!filtered.length) {
    formulaWorkbenchListEl.appendChild(createEmptyInsight("Brak formuł pasujących do bieżącego filtru."));
    return;
  }

  grouped.slice(0, 60).forEach((group) => {
    const item = document.createElement("div");
    item.className = "formula-item";

    const top = document.createElement("div");
    top.className = "formula-item-top";

    const title = document.createElement("div");
    title.className = "formula-item-title";
    title.textContent = group.entries.length > 1
      ? `${group.header} • ${group.entries.length} takich samych`
      : `${group.firstEntry.address} • ${group.header}`;

    const kind = document.createElement("div");
    kind.className = `formula-item-kind${group.missingResult || group.hasError ? " warning" : ""}`;
    kind.textContent = group.functionName;

    top.appendChild(title);
    top.appendChild(kind);

    const formula = document.createElement("div");
    formula.className = "formula-item-formula";
    formula.textContent = truncateFormulaPreview(group.formulaText);
    formula.title = group.formulaText;

    const meta = document.createElement("div");
    meta.className = "formula-item-meta";
    const resultLabel = group.missingResult ? "brak wyniku" : (group.resultText || "pusty wynik");
    const addressLabel = formatFormulaAddressSample(group.entries);
    const outsideTable = group.inTable ? "" : " • poza tabela";
    meta.textContent = `Adresy: ${addressLabel} • wynik: ${resultLabel}${outsideTable}`;

    const actions = document.createElement("div");
    actions.className = "section-nav-actions";

    const btn = document.createElement("button");
    btn.className = "btn ghost btn-sm";
    btn.type = "button";
    btn.dataset.formulaAddress = group.firstEntry.address;
    btn.textContent = group.entries.length > 1 ? "Skocz do pierwszej komórki" : "Skocz do komórki";

    actions.appendChild(btn);
    item.appendChild(top);
    item.appendChild(formula);
    item.appendChild(meta);
    item.appendChild(actions);
    formulaWorkbenchListEl.appendChild(item);
  });
}

function focusFormulaEntry(address) {
  const entry = currentFormulaEntries.find((item) => item.address === address);
  if (!entry) return;
  if (!entry.inTable) {
    toast("Ta formuła jest poza głównym zakresem aktualnej tabeli", "info");
    return;
  }
  const rowEl = tbodyEl.querySelector(`tr[data-row-index="${entry.rowIndex0}"]`);
  if (rowEl) {
    rowEl.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" });
  }
  focusColumnProfile(entry.colIdx);
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
    let numericLike = 0;
    row.values.forEach((v) => {
      if (v != null && String(v).trim() !== "") {
        nonEmpty += 1;
        if (typeof v === "string") textLike += 1;
        else if (typeof v === "number" || v instanceof Date) numericLike += 1;
        else if (!(v instanceof Date) && typeof v !== "number") textLike += 1;
      }
    });
    const n = row.values.length;
    if (n === 0) continue;
    const textRatio = nonEmpty ? textLike / nonEmpty : 0;
    const numericRatio = nonEmpty ? numericLike / nonEmpty : 0;
    if (
      nonEmpty >= 2
      && textRatio >= 0.8
      && numericRatio === 0
      && nonEmpty <= Math.max(6, Math.ceil(n * 0.75))
    ) {
      row.isSubheader = true;
    }
  }
  return rows;
}

function detectHeaderRowSimple(sheet) {
  const range = computeEffectiveSheetRange(sheet, 1);
  const maxRow = Math.min(range.e.r, range.s.r + 100);
  let bestRow = range.s.r;
  let bestScore = -Infinity;
  for (let r = range.s.r; r <= maxRow; r++) {
    let filled = 0;
    let stringCount = 0;
    let numericCount = 0;
    let formulaCount = 0;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      if (!cell) continue;
      const v = cell.v;
      if (v === null || v === "") continue;
      filled += 1;
      if (typeof v === "string") stringCount += 1;
      if (typeof v === "number" || v instanceof Date) numericCount += 1;
      if (cell.f) formulaCount += 1;
    }
    if (!filled) continue;
    const textRatio = stringCount / filled;
    const numericRatio = numericCount / filled;
    let score = (filled * 5) + (stringCount * 4) - (numericCount * 3) - formulaCount;
    if (filled >= 4) score += 10;
    if (textRatio >= 0.7) score += 10;
    if (numericRatio === 0) score += 4;
    if (r > range.s.r) score += Math.min(4, r - range.s.r);
    if (score > bestScore) {
      bestScore = score;
      bestRow = r;
    }
  }
  return bestRow + 1;
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
  if (filterEmptyModeEl.value !== "all") count += 1;
  if (filterEmptyMode2El.value !== "all") count += 1;
  if (filterNegateEl.checked) count += 1;
  if (filterNegate2El.checked) count += 1;
  if (onlyNonEmptyEl.checked) count += 1;
  if (dateModeEl.value === "last_n_days") count += 1;
  if (dateFromEl.value.trim() || dateToEl.value.trim()) count += 1;
  if (dateEmptyModeEl.value !== "all") count += 1;
  if (dateNegateEl.checked) count += 1;
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

function getQuickSearchModeValue() {
  return filterModeEl && filterModeEl.value === "Równa się" ? "exact" : "contains";
}

function syncQuickSearchModeControls() {
  const mode = getQuickSearchModeValue();
  if (quickSearchModeEl) quickSearchModeEl.value = mode;
  if (quickSearchPopupModeEl) quickSearchPopupModeEl.value = mode;
}

function applyQuickSearchMode(mode) {
  const normalized = mode === "exact" ? "Równa się" : "Zawiera";
  if (filterModeEl) filterModeEl.value = normalized;
  syncQuickSearchModeControls();
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
  filterEmptyModeEl.value = "all";
  filterEmptyMode2El.value = "all";
  filterNegateEl.checked = false;
  filterNegate2El.checked = false;
  onlyNonEmptyEl.checked = false;
  dateModeEl.value = "between";
  dateFromEl.value = "";
  dateToEl.value = "";
  lastDaysEl.value = "";
  dateEmptyModeEl.value = "all";
  dateNegateEl.checked = false;
  columnSelections.filter1.clear();
  columnSelections.filter2.clear();
  columnSelections.date.clear();
  syncQuickSearchInputs();
  syncQuickSearchModeControls();
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
  requestAnimationFrame(() => syncSidebarHandle());
  window.setTimeout(() => syncSidebarHandle(), 180);
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
  columnListEl.replaceChildren();
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
    renderActiveTable();
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

function createThemeIcon(isDark) {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("width", "18");
  svg.setAttribute("height", "18");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("fill", "none");
  svg.setAttribute("stroke", "currentColor");
  svg.setAttribute("stroke-width", "2");
  svg.setAttribute("stroke-linecap", "round");
  svg.setAttribute("stroke-linejoin", "round");

  if (isDark) {
    // Sun icon for dark mode active (click to switch to light)
    const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
    circle.setAttribute("cx", "12");
    circle.setAttribute("cy", "12");
    circle.setAttribute("r", "5");
    svg.appendChild(circle);

    const rays = [
      { x1: "12", y1: "1", x2: "12", y2: "3" },
      { x1: "12", y1: "21", x2: "12", y2: "23" },
      { x1: "4.22", y1: "4.22", x2: "5.64", y2: "5.64" },
      { x1: "18.36", y1: "18.36", x2: "19.78", y2: "19.78" },
      { x1: "1", y1: "12", x2: "3", y2: "12" },
      { x1: "21", y1: "12", x2: "23", y2: "12" },
      { x1: "4.22", y1: "19.78", x2: "5.64", y2: "18.36" },
      { x1: "18.36", y1: "5.64", x2: "19.78", y2: "4.22" },
    ];

    rays.forEach(ray => {
      const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
      Object.entries(ray).forEach(([attr, val]) => line.setAttribute(attr, val));
      svg.appendChild(line);
    });
  } else {
    // Moon icon for light mode active (click to switch to dark)
    const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("d", "M21 12.79A9 9 0 1 1 11.21 3a7 7 0 0 0 9.79 9.79z");
    svg.appendChild(path);
  }

  return svg;
}

function setTheme(theme, persist = true) {
  rootEl.setAttribute("data-theme", theme);
  themeToggle.setAttribute("aria-pressed", theme === "dark" ? "true" : "false");
  if (persist) localStorage.setItem(THEME_KEY, theme);
  
  // Clear existing content and append safe SVG element
  themeToggle.textContent = "";
  themeToggle.appendChild(createThemeIcon(theme === "dark"));
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

async function hardRefreshApp() {
  try {
    if ("serviceWorker" in navigator) {
      const registrations = await navigator.serviceWorker.getRegistrations();
      await Promise.all(registrations.map((registration) => registration.update().catch(() => {})));
    }

    if ("caches" in window) {
      const keys = await caches.keys();
      const appKeys = keys.filter((key) => key.startsWith("excel-wb-"));
      await Promise.all(appKeys.map((key) => caches.delete(key).catch(() => false)));
    }

    toast("Czyszcze cache i odswiezam aplikacje...", "info");
  } catch {
    toast("Odswiezam aplikacje...", "info");
  }

  window.location.reload();
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
    sheetSelect.replaceChildren();
    workbook.SheetNames.forEach((s) => {
      const opt = document.createElement("option");
      opt.value = s;
      opt.textContent = s;
    sheetSelect.appendChild(opt);
  });
    currentWorkbookStats = collectWorkbookStats(workbook, file.name);
    currentSheetStats = null;
    currentKpiEntries = [];
    currentKpiAnchorRow = 1;
    currentColumnProfiles = [];
    currentSections = [];
    currentRepeatingBlocks = [];
    currentDisplayModel = null;
    tableViewMode = "wide";
    multiSortState = [];
    sortState = { col: "", dir: "asc" };
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
    renderInsights();
    renderKpiExtractor();
    renderColumnProfiles();
    renderSections();
    renderRepeatingBlocks();
    renderDurationAnalysis();
    populateSortColumnSelect();
    renderSortPresets();
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
  const model = currentDisplayModel || getDisplayModel();
  if (!model.headers.length || !model.rows.length) {
    toast("Brak danych do eksportu", "warning");
    return;
  }
  const rows = [
    model.headers,
    ...model.rows.map((row) => row.values.map((v, i) => getDisplayValue(row, i))),
  ];
  const csv = rows.map((row) => row.map(escapeCsv).join(",")).join("\n");
  const base = currentFileName ? currentFileName.replace(/\.[^.]+$/, "") : "excel-workbench";
  const sheet = sheetSelect.value ? sheetSelect.value.replace(/\s+/g, "_") : "arkusz";
  const suffix = model.mode === "long" ? "long" : "wide";
  const filename = `${base}_${sheet}_${suffix}.csv`;
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
      if (!Number.isFinite(aggregationWorkbenchState.customHeaderRow) || aggregationWorkbenchState.customHeaderRow < 1 || aggregationWorkbenchState.headerRowChoice === "auto") {
        aggregationWorkbenchState.customHeaderRow = headerRow;
      }
      currentSheetName = sheetName;
      const data = buildRows(sheet, headerRow, workbook);
      currentHeaders = data.headers;
      currentStartCol = data.startCol || 0;
      currentMerges = Array.isArray(data.merges) ? data.merges : [];
      currentHeaderStyles = Array.isArray(data.headerStyles) ? data.headerStyles : [];
      currentSheetColWidths = Array.isArray(data.colWidths) ? data.colWidths : [];
      currentSheetRowHeights = data.rowHeights || {};
      currentSheetStats = data.stats || null;
      baseRows = markSubheaderRows(data.rows);
      const kpiData = collectKpiEntries(sheet, headerRow);
      currentKpiEntries = Array.isArray(kpiData?.entries) ? kpiData.entries : [];
      currentKpiAnchorRow = Number(kpiData?.anchorRow) || headerRow;
      currentColumnProfiles = collectColumnProfiles();
      currentSections = detectSections(sheet, headerRow, data);
      currentRepeatingBlocks = detectRepeatingBlocks(sheet, headerRow, data);
      currentFormulaEntries = collectFormulaEntries(sheet, data, headerRow);
      if (!canUseLongView()) tableViewMode = "wide";
      viewRows = baseRows.slice();
      multiSortState = [];
      sortState = { col: "", dir: "asc" };
      manualColumnWidths = {};
      columnSelections.filter1.clear();
      columnSelections.filter2.clear();
      columnSelections.date.clear();
      updateColumnSummary();
      updateFilterBadge();
      populateSortColumnSelect();
      renderActiveTable();
      renderInsights();
      renderKpiExtractor();
      renderSheetInspectorSummary();
      renderColumnProfiles();
      renderSections();
      renderRepeatingBlocks();
      renderDurationAnalysis();
      renderAggregationWorkbench();
      renderFormulaWorkbench();
      setDirtyState(false);
      if ((currentSheetStats?.trimmedColumns || 0) > 0) {
        log(`Przycięto puste kolumny poza realnym zakresem danych: ${currentSheetStats.trimmedColumns}`, "info");
      }
      if (currentSheetStats?.duplicateHeaderCount) {
        toast(`Zdublowane naglowki rozrozniono (${currentSheetStats.duplicateHeaderCount})`, "warning");
      }
      toast("Arkusz wczytany", "success");
      log(`Wczytano arkusz: ${sheetName}`, "success");
      setTimeout(() => {
        const panelFileSheet = document.getElementById("panel-file-sheet");
        if (panelFileSheet) panelFileSheet.removeAttribute("open");
      }, 100);
    } finally {
      setLoading(false);
    }
  }, 50);
});

applyFilterBtn.addEventListener("click", () => {
  if (!currentHeaders.length) return;
  applyFilters();
  sortRows();
  renderActiveTable();
  renderInsights();
  renderKpiExtractor();
  renderSheetInspectorSummary();
  renderColumnProfiles();
  renderSections();
  renderRepeatingBlocks();
  renderDurationAnalysis();
  renderAggregationWorkbench();
  updateFilterBadge();
  toast("Zastosowano filtry", "info");
});

function applyQuickSearch() {
  if (!currentHeaders.length) return;
  let value = "";
  if (quickSearchPopupEl && !quickSearchPopupEl.classList.contains("hidden") && quickSearchPopupInput) value = quickSearchPopupInput.value;
  else if (quickSearchEl) value = quickSearchEl.value;
  else value = searchQueryEl.value || "";
  const popupActive = quickSearchPopupEl && !quickSearchPopupEl.classList.contains("hidden");
  if (popupActive && quickSearchPopupModeEl) applyQuickSearchMode(quickSearchPopupModeEl.value);
  else if (quickSearchModeEl) applyQuickSearchMode(quickSearchModeEl.value);
  if (quickSearchPopupInput) quickSearchPopupInput.value = value;
  if (quickSearchEl) quickSearchEl.value = value;
  searchQueryEl.value = value;
  applyFilters();
  sortRows();
  renderActiveTable();
  renderInsights();
  renderSheetInspectorSummary();
  renderColumnProfiles();
  renderSections();
  renderRepeatingBlocks();
  renderDurationAnalysis();
  renderAggregationWorkbench();
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
  syncSidebarHandle();
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

if (quickSearchModeEl) {
  quickSearchModeEl.addEventListener("change", () => {
    applyQuickSearchMode(quickSearchModeEl.value);
  });
}

if (quickSearchPopupInput) {
  quickSearchPopupInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); applyQuickSearch(); }
  });
}
if (quickSearchPopupModeEl) {
  quickSearchPopupModeEl.addEventListener("change", () => {
    applyQuickSearchMode(quickSearchPopupModeEl.value);
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
  renderActiveTable();
  renderInsights();
  renderColumnProfiles();
  renderSections();
  renderRepeatingBlocks();
  renderDurationAnalysis();
  renderAggregationWorkbench();
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
    renderActiveTable();
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

if (addSortRuleBtn) {
  addSortRuleBtn.addEventListener("click", () => {
    if (!currentHeaders.length) {
      toast("Najpierw wczytaj arkusz", "info");
      return;
    }
    const col = sortColumnSelectEl?.value;
    const dir = sortDirectionSelectEl?.value === "desc" ? "desc" : "asc";
    if (!col) return;
    multiSortState = multiSortState.filter((rule) => rule.col !== col);
    multiSortState.push({ col, dir });
    normalizeSortState();
    applyCurrentSort();
    toast("Dodano sortowanie do kolejki", "info");
  });
}

if (sortRulesListEl) {
  sortRulesListEl.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-sort-action]");
    if (!btn) return;
    const action = btn.dataset.sortAction;
    const index = parseInt(btn.dataset.sortIndex || "", 10);
    if (!Number.isFinite(index) || index < 0 || index >= multiSortState.length) return;

    if (action === "remove") {
      multiSortState.splice(index, 1);
    } else if (action === "toggle") {
      multiSortState[index].dir = multiSortState[index].dir === "asc" ? "desc" : "asc";
    } else if (action === "up" && index > 0) {
      [multiSortState[index - 1], multiSortState[index]] = [multiSortState[index], multiSortState[index - 1]];
    } else if (action === "down" && index < multiSortState.length - 1) {
      [multiSortState[index + 1], multiSortState[index]] = [multiSortState[index], multiSortState[index + 1]];
    }

    normalizeSortState();
    applyCurrentSort();
  });
}

if (saveSortPresetBtn) {
  saveSortPresetBtn.addEventListener("click", () => {
    normalizeSortState();
    if (!multiSortState.length) {
      toast("Brak sortowan do zapisania", "warning");
      return;
    }
    const name = window.prompt("Nazwa presetu sortowania:", "");
    if (!name || !name.trim()) return;
    const trimmed = name.trim();
    const presets = loadSortPresets().filter((preset) => preset.name !== trimmed);
    presets.push({ name: trimmed, rules: multiSortState.map((rule) => ({ ...rule })) });
    presets.sort((a, b) => a.name.localeCompare(b.name, "pl"));
    saveSortPresets(presets);
    renderSortPresets();
    if (sortPresetSelectEl) sortPresetSelectEl.value = trimmed;
    toast("Zapisano preset sortowania", "success");
  });
}

if (applySortPresetBtn) {
  applySortPresetBtn.addEventListener("click", () => {
    const name = sortPresetSelectEl?.value;
    if (!name) {
      toast("Wybierz preset", "info");
      return;
    }
    const preset = loadSortPresets().find((item) => item.name === name);
    if (!preset) {
      toast("Nie znaleziono presetu", "warning");
      renderSortPresets();
      return;
    }
    multiSortState = Array.isArray(preset.rules) ? preset.rules.map((rule) => ({ col: rule.col, dir: rule.dir })) : [];
    normalizeSortState();
    applyCurrentSort();
    toast("Wczytano preset sortowania", "success");
  });
}

if (deleteSortPresetBtn) {
  deleteSortPresetBtn.addEventListener("click", () => {
    const name = sortPresetSelectEl?.value;
    if (!name) {
      toast("Wybierz preset do usuniecia", "info");
      return;
    }
    const presets = loadSortPresets().filter((preset) => preset.name !== name);
    saveSortPresets(presets);
    renderSortPresets();
    toast("Usunieto preset sortowania", "info");
  });
}

columnPickerEl.addEventListener("click", (e) => {
  if (e.target === columnPickerEl) closeColumnPicker();
});

columnPickerEl.addEventListener("keydown", handlePickerKeydown);

closePickerBtn.addEventListener("click", closeColumnPicker);
columnSearchEl.addEventListener("input", filterColumnList);

exportCsvBtn.addEventListener("click", exportCsv);
if (resetSortBtn) {
  resetSortBtn.addEventListener("click", () => {
    multiSortState = [];
    normalizeSortState();
    applyCurrentSort();
    toast("Przywrocono domyslne sortowanie", "info");
  });
}
saveBtn.addEventListener("click", () => {
  toast("Wersja webowa nie nadpisuje pliku. Użyj „Zapisz jako…”", "info");
});
saveAsBtn.addEventListener("click", saveWorkbookAs);
resetWidthsBtn.addEventListener("click", () => {
  manualColumnWidths = {};
  renderActiveTable();
  toast("Przywrocono automatyczne szerokosci", "info");
});

tbodyEl.addEventListener("click", (e) => {
  const td = e.target.closest("td");
  if (!td || td.classList.contains("row-head")) return;
  const tr = td.parentElement;
  const rowKey = tr?.dataset.rowKey || "";
  const colIndex0 = parseInt(td.dataset.colIndex || "", 10);
  if (!rowKey || !Number.isFinite(colIndex0)) return;
  setFocusedCell(rowKey, colIndex0, { scroll: false });
});

tbodyEl.addEventListener("dblclick", (e) => {
  const td = e.target.closest("td");
  if (!td || td.classList.contains("row-head")) return;
  toast("Edycja komorek jest tymczasowo zablokowana, dopoki lepiej nie dopracujemy zapisu stylow i zgodnosci pliku.", "info");
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
filterModeEl.addEventListener("change", syncQuickSearchModeControls);

maxRowsEl.addEventListener("change", () => {
  saveMaxRowsPreference();
  renderActiveTable();
});

zoomLevelEl.addEventListener("change", () => {
  setTimeout(applyZoom, 50);
});

if (excelLayoutToggleEl) {
  excelLayoutToggleEl.addEventListener("click", () => {
    setExcelLayoutEnabled(!isExcelLayoutEnabled());
    saveExcelLayoutPreference();
    renderActiveTable();
  });
}

initIntroSplash();
initTheme();
loadMaxRowsPreference();
loadExcelLayoutPreference();
attachResizeHandlers();
applyZoom();
updateNetworkBadge();
window.addEventListener("online", updateNetworkBadge);
window.addEventListener("offline", updateNetworkBadge);

themeToggle.addEventListener("click", () => {
  const next = rootEl.getAttribute("data-theme") === "dark" ? "light" : "dark";
  setTheme(next);
});

if (brandRefreshBtn) {
  brandRefreshBtn.addEventListener("click", () => {
    hardRefreshApp();
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
    if (isSidebarOpen() && sidebarEl) {
      const rect = sidebarEl.getBoundingClientRect();
      const overlap = 8;
      const nextLeft = Math.max(8, Math.round(rect.right - overlap)); // [EN] Allow tighter edge on narrow viewports; CSS handles closed state
      panelHandle.style.left = `${nextLeft}px`;
    } else {
      panelHandle.style.removeProperty("left"); // [EN] Let .sidebar-handle use fluid clamp() when closed
    }
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
document.querySelectorAll("details.panel").forEach((det) => {
  det.addEventListener("toggle", () => {
    if (!isSidebarOpen()) return;
    requestAnimationFrame(() => syncSidebarHandle()); // [EN] :has() width changes — no resize event; keep handle aligned
    window.setTimeout(() => syncSidebarHandle(), 260);
  });
});
if (sectionNavigatorEl) {
  sectionNavigatorEl.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-section-index]");
    if (!btn) return;
    const idx = parseInt(btn.dataset.sectionIndex || "", 10);
    if (!Number.isFinite(idx) || idx < 0 || idx >= currentSections.length) return;
    focusSection(currentSections[idx]);
  });
}
if (repeatBlockDetectorEl) {
  repeatBlockDetectorEl.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-repeat-group-index]");
    if (!btn) return;
    const groupIndex = parseInt(btn.dataset.repeatGroupIndex || "", 10);
    const blockIndex = parseInt(btn.dataset.repeatBlockIndex || "", 10);
    if (!Number.isFinite(groupIndex) || !Number.isFinite(blockIndex)) return;
    focusRepeatingBlock(groupIndex, blockIndex);
  });
}
if (durationAnalysisSummaryEl) {
  durationAnalysisSummaryEl.addEventListener("click", (e) => {
    e.stopPropagation();
    const btn = e.target.closest("button[data-duration-action]");
    if (!btn) return;
    const action = btn.dataset.durationAction;

    if (action === "toggle-long" && canUseLongView()) {
      tableViewMode = tableViewMode === "long" ? "wide" : "long";
      manualColumnWidths = {};
      renderActiveTable();
      renderSheetInspectorSummary();
      renderDurationAnalysis();
      renderAggregationWorkbench();
      toast(tableViewMode === "long" ? "Wlaczono Wide-to-Long" : "Wrocono do klasycznego widoku", "info");
      return;
    }

    if (action === "reset-filters") {
      resetFiltersBtn.click();
    }
  });
  durationAnalysisSummaryEl.addEventListener("change", (e) => {
    e.stopPropagation();
    const control = e.target.closest("[data-duration-control]");
    if (!control) return;
    const kind = control.dataset.durationControl;
    if (kind === "status") {
      durationAnalysisState.statusFilter = control.value || "all";
    } else if (kind === "sort") {
      durationAnalysisState.sortMetric = control.value || "avg";
    } else if (kind === "count") {
      const next = parseInt(control.value || "14", 10);
      durationAnalysisState.showCount = Number.isFinite(next) && next > 0 ? next : 14;
    }
    renderDurationAnalysis();
    renderAggregationWorkbench();
  });
}
if (durationAnalysisListEl) {
  durationAnalysisListEl.addEventListener("click", (e) => {
    e.stopPropagation();
    const btn = e.target.closest("button[data-duration-action='filter-entity']");
    if (!btn) return;
    const entity = (btn.dataset.durationEntity || "").trim();
    if (!entity) return;
    searchQueryEl.value = entity;
    applyFilters();
    sortRows();
    renderActiveTable();
    renderInsights();
    renderKpiExtractor();
    renderSheetInspectorSummary();
    renderColumnProfiles();
    renderSections();
    renderRepeatingBlocks();
    renderDurationAnalysis();
    renderAggregationWorkbench();
    renderFormulaWorkbench();
    updateFilterBadge();
    toast(`Przefiltrowano widok dla: ${entity}`, "info");
  });
}
if (aggregationWorkbenchSummaryEl) {
  aggregationWorkbenchSummaryEl.addEventListener("change", (e) => {
    e.stopPropagation();
    const sidebarEl = document.querySelector(".sidebar");
    const savedSidebarScroll = sidebarEl ? sidebarEl.scrollTop : 0;
    const control = e.target.closest("[data-aggregation-control]");
    if (!control) return;
    const kind = control.dataset.aggregationControl;
    if (kind === "source") aggregationWorkbenchState.sourceMode = control.value || "auto";
    if (kind === "scope") aggregationWorkbenchState.scopeMode = control.value || "filtered";
    if (kind === "header") {
      aggregationWorkbenchState.headerRowChoice = control.value === "manual" ? "manual" : "auto";
      if (aggregationWorkbenchState.headerRowChoice === "manual") {
        const fallbackRow = Number.isFinite(aggregationWorkbenchState.customHeaderRow) && aggregationWorkbenchState.customHeaderRow > 0
          ? aggregationWorkbenchState.customHeaderRow
          : currentHeaderRow;
        aggregationWorkbenchState.customHeaderRow = fallbackRow;
      }
    }
    if (kind === "header-number") {
      const next = parseInt(control.value || "", 10);
      if (!Number.isFinite(next) || next < 1) {
        toast("Podaj dodatni numer wiersza naglowka.", "warning");
        control.value = String(aggregationWorkbenchState.customHeaderRow || currentHeaderRow);
        return;
      }
      if (!isValidAggregationHeaderRow(next)) {
        toast(`Wiersz ${next} nie wyglada na poprawny naglowek dla tego arkusza.`, "error");
        control.value = String(aggregationWorkbenchState.customHeaderRow || currentHeaderRow);
        return;
      }
      aggregationWorkbenchState.customHeaderRow = next;
      aggregationWorkbenchState.headerRowChoice = "manual";
    }
    if (kind === "group") aggregationWorkbenchState.groupBy = control.value || "";
    if (kind === "measure") aggregationWorkbenchState.measure = control.value || "count_rows";
    if (kind === "aggregation") aggregationWorkbenchState.aggregation = control.value || "count";
    if (kind === "match") aggregationWorkbenchState.matchMode = control.value || "contains";
    if (kind === "measurefilter") {
      aggregationWorkbenchState.measureFilterMode = control.value || "all";
      const valueInput = aggregationWorkbenchSummaryEl.querySelector("[data-aggregation-control=\"measurefilter-value\"]");
      if (valueInput) {
        valueInput.style.display = aggregationWorkbenchState.measureFilterMode === "all" ? "none" : "inline-block";
      }
    }
    if (kind === "measurefilter-value") {
      aggregationWorkbenchState.measureFilterValue = control.value || "";
    }
    if (kind === "count") {
      const next = parseInt(control.value || "20", 10);
      aggregationWorkbenchState.showCount = Number.isFinite(next) && next > 0 ? next : 20;
    }
    if (kind === "having") {
      aggregationWorkbenchState.havingMode = control.value || "all";
      const valueInput = aggregationWorkbenchSummaryEl.querySelector("[data-aggregation-control=\"having-value\"]");
      if (valueInput) {
        valueInput.style.display = aggregationWorkbenchState.havingMode === "all" ? "none" : "inline-block";
      }
    }
    if (kind === "having-value") {
      const next = parseFloat(control.value || "0", 10);
      aggregationWorkbenchState.havingValue = Number.isFinite(next) && next >= 0 ? next : 10;
    }
    renderAggregationWorkbench();
    if (sidebarEl) sidebarEl.scrollTop = savedSidebarScroll;
  });
}
if (aggregationWorkbenchListEl) {
  aggregationWorkbenchListEl.addEventListener("keydown", (e) => {
    if (e.target.classList.contains("aggregation-result-search") && e.key === "Enter") {
      e.preventDefault();
      aggregationWorkbenchState.resultSearch = e.target.value || "";
      renderAggregationWorkbench();
    }
  });
  aggregationWorkbenchListEl.addEventListener("click", (e) => {
    e.stopPropagation();
    const btn = e.target.closest("button[data-aggregation-action='filter-group']");
    if (!btn) return;
    const value = (btn.dataset.aggregationValue || "").trim();
    if (!value) return;
    searchQueryEl.value = value;
    if (filterModeEl) {
      filterModeEl.value = aggregationWorkbenchState.matchMode === "exact" ? "Równa się" : "Zawiera";
    }
    applyFilters();
    sortRows();
    renderActiveTable();
    renderInsights();
    renderKpiExtractor();
    renderSheetInspectorSummary();
    renderColumnProfiles();
    renderSections();
    renderRepeatingBlocks();
    renderDurationAnalysis();
    renderAggregationWorkbench();
    renderFormulaWorkbench();
    updateFilterBadge();
    toast(`Przefiltrowano widok dla: ${value}`, "info");
  });
}
if (columnProfilerEl) {
  columnProfilerEl.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-profile-col-index]");
    if (!btn) return;
    const colIdx = parseInt(btn.dataset.profileColIndex || "", 10);
    if (!Number.isFinite(colIdx)) return;
    focusColumnProfile(colIdx);
  });
}
if (sheetInspectorSummaryEl) {
  sheetInspectorSummaryEl.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-inspector-action]");
    if (!btn) return;
    const action = btn.dataset.inspectorAction;

    if (action === "set-header") {
      const headerRow = parseInt(btn.dataset.inspectorHeaderRow || "", 10);
      if (!Number.isFinite(headerRow)) return;
      if (autoHeaderRowEl) autoHeaderRowEl.checked = false;
      headerRowEl.value = String(headerRow);
      loadBtn.click();
      return;
    }

    if (action === "toggle-long" && canUseLongView()) {
      tableViewMode = tableViewMode === "long" ? "wide" : "long";
      manualColumnWidths = {};
      renderActiveTable();
      renderSheetInspectorSummary();
      renderDurationAnalysis();
      renderAggregationWorkbench();
      toast(tableViewMode === "long" ? "Wlaczono Wide-to-Long" : "Wrocono do klasycznego widoku", "info");
      return;
    }

    if (action === "focus-col") {
      const colIdx = parseInt(btn.dataset.profileColIndex || "", 10);
      if (!Number.isFinite(colIdx)) return;
      focusColumnProfile(colIdx);
    }
  });
}
if (formulaWorkbenchListEl) {
  formulaWorkbenchListEl.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-formula-address]");
    if (!btn) return;
    focusFormulaEntry(btn.dataset.formulaAddress || "");
  });
}
if (kpiListEl) {
  kpiListEl.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-kpi-address]");
    if (!btn) return;
    focusKpiEntry(btn.dataset.kpiAddress || "");
  });
}
if (wideLongToggleEl) {
  wideLongToggleEl.addEventListener("click", () => {
    if (!canUseLongView()) return;
    tableViewMode = tableViewMode === "long" ? "wide" : "long";
    manualColumnWidths = {};
    renderActiveTable();
    renderDurationAnalysis();
    renderAggregationWorkbench();
    toast(tableViewMode === "long" ? "Wlaczono Wide-to-Long" : "Wrocono do klasycznego widoku", "info");
  });
}
if (readingToggle) {
  readingToggle.addEventListener("click", () => {
    const enabled = !rootEl.classList.contains("reading");
    setReadingMode(enabled);
  });
}
[formulaSearchEl, formulaFilterEl, formulaFunctionFilterEl].forEach((el) => {
  if (!el) return;
  el.addEventListener("input", renderFormulaWorkbench);
  el.addEventListener("change", renderFormulaWorkbench);
});

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
  if (!meta && !e.altKey && !shouldIgnoreTableArrowNavigation()) {
    let handled = false;
    if (e.shiftKey) {
      if (!selectedCellState && focusedCellState) {
        setSelectedCell(focusedCellState.rowKey, focusedCellState.colIndex0, { scroll: false });
      }
      if (e.key === "ArrowUp") handled = moveSelectedCell(-1, 0);
      if (e.key === "ArrowDown") handled = moveSelectedCell(1, 0);
      if (e.key === "ArrowLeft") handled = moveSelectedCell(0, -1);
      if (e.key === "ArrowRight") handled = moveSelectedCell(0, 1);
    } else {
      if (e.key === "ArrowUp") handled = moveFocusedCell(-1, 0);
      if (e.key === "ArrowDown") handled = moveFocusedCell(1, 0);
      if (e.key === "ArrowLeft") handled = moveFocusedCell(0, -1);
      if (e.key === "ArrowRight") handled = moveFocusedCell(0, 1);
    }
    if (handled) {
      e.preventDefault();
      return;
    }
  }
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
  // Dodatkowo umożliwiam użycie klawisza "Q" zamiast Escape (np. dla klawiatur bez klawisza Escape)

  // Wersja z nowym const (pierwotna)
 /*  const isEscapeOrQ = (key) => key === "Escape" || key.toLowerCase() === "q";

  if (e.shiftKey && isEscapeOrQ(e.key) && selectedCellState) {
    e.preventDefault();
    setSelectedCell("", -1);
    return;
  }
  if (!e.shiftKey && isEscapeOrQ(e.key) && focusedCellState) {
    e.preventDefault();
    setFocusedCell("", -1);
    return;
  }
 */
  if (e.shiftKey && (e.key === "Escape" || e.key.toLowerCase() === "q") && selectedCellState) {
    e.preventDefault();
    setSelectedCell("", -1);
    return;
  }
  if (!e.shiftKey && (e.key === "Escape" || e.key.toLowerCase() === "q") && focusedCellState) {
    e.preventDefault();
    setFocusedCell("", -1);
    return;
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
renderInsights();
renderKpiExtractor();
renderSheetInspectorSummary();
renderColumnProfiles();
renderSections();
renderRepeatingBlocks();
renderDurationAnalysis();
renderAggregationWorkbench();
renderFormulaWorkbench();
populateSortColumnSelect();
renderSortPresets();
updateWideLongToggle();

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
  navigator.serviceWorker.register(`sw.js?v=${APP_BUILD_VERSION}`).then((registration) => {
    registration.update().catch(() => {});
  }).catch(() => {});
}
