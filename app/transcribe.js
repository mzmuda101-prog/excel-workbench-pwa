// Tryb spisywania — ekran do PRZEPISYWANIA danych ręcznie na papier.
//
// Po co osobny tryb: eksport (CSV / Druk-PDF) obsługuje wyjście MASZYNOWE.
// Ręczne spisywanie ma inne wąskie gardła: gubienie wiersza po podniesieniu wzroku
// znad kartki, brak śladu „gdzie skończyłem”, kolejność kolumn na ekranie inna niż
// kolejność rubryk na formularzu, gaszący się tablet i dłoń przewijająca tabelę.
//
// Model: JEDEN wiersz naraz jako duża karta pól, wielki przycisk „Spisane i dalej”,
// odhaczanie zapisywane lokalnie (wznowienie po przerwie), Wake Lock i blokada dotyku.
//
// Źródło danych: currentDisplayModel (ten sam, co eksport), więc tryb automatycznie
// respektuje filtry, sortowanie, Wide-to-Long i kolumny wyliczane. Wiersze są
// SNAPSHOTOWANE przy otwarciu — widok pod spodem nie przesunie się pod ręką.

const TR_STORE_KEY = "excel-workbench-transcribe";
const TR_MAX_SCOPES = 12;      // ile plików/arkuszy pamiętamy (reszta wypada, najstarsze pierwsze)
const TR_MAX_DONE = 20000;     // bezpiecznik na rozmiar localStorage
const TR_DEFAULT_FIELDS = 10;  // ile pól proponujemy przy pierwszym otwarciu arkusza
const TR_FONT_STEPS = [1, 2, 3];

const trOverlayEl = document.getElementById("transcribeOverlay");
const trBtn = document.getElementById("transcribeBtn");
const trCardEl = document.getElementById("trCard");
const trEmptyEl = document.getElementById("trEmpty");
const trCounterEl = document.getElementById("trCounter");
const trDoneCountEl = document.getElementById("trDoneCount");
const trProgressBarEl = document.getElementById("trProgressBar");
const trSourceEl = document.getElementById("trSource");
const trPrevBtn = document.getElementById("trPrevBtn");
const trNextBtn = document.getElementById("trNextBtn");
const trMarkBtn = document.getElementById("trMarkBtn");
const trMarkChipEl = document.getElementById("trMarkChip");
const trHideDoneEl = document.getElementById("trHideDone");
const trCloseBtn = document.getElementById("trCloseBtn");
const trFontBtn = document.getElementById("trFontBtn");
const trLockBtn = document.getElementById("trLockBtn");
const trFieldsBtn = document.getElementById("trFieldsBtn");
const trFieldsPanelEl = document.getElementById("trFieldsPanel");
const trFieldsListEl = document.getElementById("trFieldsList");
const trFieldsAllBtn = document.getElementById("trFieldsAllBtn");
const trFieldsNoneBtn = document.getElementById("trFieldsNoneBtn");
const trFieldsDoneBtn = document.getElementById("trFieldsDoneBtn");
const trResetBtn = document.getElementById("trResetBtn");
const trAutoFieldsEl = document.getElementById("trAutoFields");
const trAutoNoteEl = document.getElementById("trAutoNote");
const trTouchShieldEl = document.getElementById("trTouchShield");
const trLiveEl = document.getElementById("trLive");

let trIsOpen = false;
let trRows = [];             // snapshot wierszy z modelu widoku
let trHeaders = [];
let trRowHeadFormatter = null;
let trFieldOrder = [];       // WSZYSTKIE indeksy kolumn w kolejności ustawionej przez użytkownika
let trSelected = new Set();  // które z nich trafiają na kartę
let trDone = new Set();      // klucze wierszy już spisanych
let trOrder = [];            // indeksy do trRows widoczne w bieżącym trybie (hideDone)
let trPos = 0;
let trHideDone = false;
let trAutoFields = false;
let trFont = 2;
let trLocked = false;
let trScope = "";
let trWakeLock = null;
let trReturnFocusEl = null;
let trResetArmed = false;
let trResetTimer = 0;

// ── Trwałość ────────────────────────────────────────────────────────────────
// Jeden klucz, w środku mapa „zakresów” (plik + arkusz + tryb widoku). Dzięki temu
// wracasz do tego samego pliku i zastajesz swoje ✓ oraz swój układ pól.

function trLoadStore() {
  try {
    const parsed = JSON.parse(localStorage.getItem(TR_STORE_KEY) || "{}");
    if (!parsed || typeof parsed !== "object") return { scopes: {} };
    if (!parsed.scopes || typeof parsed.scopes !== "object") parsed.scopes = {};
    return parsed;
  } catch {
    return { scopes: {} };
  }
}

function trSaveStore(store) {
  try {
    const scopes = store.scopes || {};
    const keys = Object.keys(scopes);
    if (keys.length > TR_MAX_SCOPES) {
      keys
        .sort((a, b) => (scopes[a]?.ts || 0) - (scopes[b]?.ts || 0))
        .slice(0, keys.length - TR_MAX_SCOPES)
        .forEach((k) => { delete scopes[k]; });
    }
    localStorage.setItem(TR_STORE_KEY, JSON.stringify(store));
  } catch {
    /* prywatne okno / brak miejsca — tryb działa dalej, tylko bez pamięci */
  }
}

function trPersist() {
  if (!trScope) return;
  const store = trLoadStore();
  store.font = trFont;
  store.hideDone = trHideDone;
  store.scopes[trScope] = {
    order: trFieldOrder.slice(),
    sel: Array.from(trSelected),
    auto: trAutoFields,
    done: Array.from(trDone).slice(0, TR_MAX_DONE),
    cursor: trCurrentKey(),
    ts: Date.now(),
  };
  trSaveStore(store);
}

function trScopeKey(model) {
  const file = currentFileName || "?";
  const sheet = (typeof sheetSelect !== "undefined" && sheetSelect?.value) || "?";
  return `${file}::${sheet}::${model.mode || "wide"}`;
}

// ── Model / wiersze ─────────────────────────────────────────────────────────

function trKeyOf(row) {
  return typeof getRowSelectionKey === "function" ? getRowSelectionKey(row) : String(row?.rowIndex0 ?? "");
}

function trCurrentRow() {
  const idx = trOrder[trPos];
  return Number.isInteger(idx) ? trRows[idx] : null;
}

function trCurrentKey() {
  const row = trCurrentRow();
  return row ? trKeyOf(row) : "";
}

function trFieldLabel(idx) {
  return typeof exportColLabel === "function"
    ? exportColLabel(trHeaders[idx], idx)
    : String(trHeaders[idx] ?? idx + 1);
}

// Dwa tryby doboru pól:
//   ręczny  — dokładnie to, co zaznaczone (stałe rubryki formularza),
//   auto    — pola z wartością W TYM wierszu, brane ze WSZYSTKICH kolumn.
// Auto jest odpowiedzią na arkusze z powtarzanymi blokami (Kw1_*, Kw2_*, Kw3_*)
// i na „przesunięte” wiersze: raz dane siedzą jedną kolumnę w prawo, raz dwie.
// Zaznaczenia w trybie auto nie mają wpływu, ale KOLEJNOŚĆ owszem.
function trHasValue(row, idx) {
  return String(getDisplayValue(row, idx) ?? "").trim() !== "";
}

function trVisibleCols(row) {
  if (trAutoFields) {
    if (!row) return [];
    return trFieldOrder.filter((idx) => trHasValue(row, idx));
  }
  return trFieldOrder.filter((idx) => trSelected.has(idx));
}

// Ile pól tryb auto przemilczał — bez tego licznika znikające rubryki wyglądają
// jak zgubione dane, a nie jak świadome pominięcie pustych.
function trSkippedCount(row) {
  if (!trAutoFields || !row) return 0;
  return trFieldOrder.length - trVisibleCols(row).length;
}

// Domyślny zestaw pól przy pierwszym otwarciu arkusza: kolumny, które faktycznie
// coś zawierają (próbka wierszy), przycięte do TR_DEFAULT_FIELDS. Lepszy start niż
// 40 pustych rubryk — resztę użytkownik dokłada w panelu „Pola”.
function trDefaultSelection() {
  const sample = trRows.slice(0, 300);
  const filled = [];
  trHeaders.forEach((_, idx) => {
    const has = sample.some((row) => String(getDisplayValue(row, idx) ?? "").trim() !== "");
    if (has) filled.push(idx);
  });
  const base = filled.length ? filled : trHeaders.map((_, i) => i);
  return new Set(base.slice(0, TR_DEFAULT_FIELDS));
}

function trRebuildOrder(preserveKey) {
  trOrder = [];
  trRows.forEach((row, i) => {
    if (!trHideDone || !trDone.has(trKeyOf(row))) trOrder.push(i);
  });
  if (preserveKey) {
    const at = trOrder.findIndex((i) => trKeyOf(trRows[i]) === preserveKey);
    if (at >= 0) trPos = at;
  }
  trPos = Math.max(0, Math.min(trPos, Math.max(0, trOrder.length - 1)));
}

// ── Render ──────────────────────────────────────────────────────────────────

function trRenderCard() {
  if (!trCardEl) return;
  const row = trCurrentRow();
  const cols = trVisibleCols(row);
  const total = trOrder.length;

  trCardEl.replaceChildren();
  const noRows = !row;
  const noFields = !cols.length;
  trCardEl.classList.toggle("hidden", noRows || noFields);
  if (trEmptyEl) {
    trEmptyEl.classList.toggle("hidden", !(noRows || noFields));
    if (noRows || noFields) {
      trEmptyEl.replaceChildren();
      const title = document.createElement("div");
      title.className = "tr-empty-title";
      const sub = document.createElement("div");
      sub.className = "tr-empty-sub";
      if (noFields && trAutoFields) {
        // W trybie auto „brak pól” znaczy: cały wiersz jest pusty. To inny komunikat
        // niż „nic nie zaznaczyłeś” — inaczej wygląda na awarię.
        title.textContent = t("trRowEmpty");
        sub.textContent = t("trRowEmptySub");
      } else if (noFields) {
        title.textContent = t("trNoFields");
        sub.textContent = t("trNoFieldsSub");
      } else if (trDone.size >= trRows.length && trRows.length) {
        title.textContent = t("trAllDone");
        sub.textContent = t("trAllDoneSub");
      } else {
        title.textContent = t("trNothingToShow");
        sub.textContent = t("trAllDoneSub");
      }
      trEmptyEl.append(title, sub);
    }
  }

  if (!noRows && !noFields) {
    const head = document.createElement("div");
    head.className = "tr-card-head";
    const srcRow = document.createElement("span");
    srcRow.className = "tr-card-rownum";
    const rowLabel = trRowHeadFormatter ? trRowHeadFormatter(row) : String((row.rowIndex0 ?? 0) + 1);
    srcRow.textContent = t("trSheetRow", { n: rowLabel });
    const skipped = trSkippedCount(row);
    if (skipped > 0) {
      const skippedEl = document.createElement("span");
      skippedEl.className = "tr-card-skipped";
      skippedEl.textContent = t("trSkipped", { n: skipped });
      head.appendChild(skippedEl);
    }
    head.appendChild(srcRow);
    trCardEl.appendChild(head);

    cols.forEach((ci) => {
      const field = document.createElement("div");
      field.className = "tr-field";
      const label = document.createElement("div");
      label.className = "tr-field-label";
      label.textContent = trFieldLabel(ci);
      const value = document.createElement("div");
      value.className = "tr-field-value";
      const raw = String(getDisplayValue(row, ci) ?? "").trim();
      if (raw) {
        value.textContent = raw;
      } else {
        value.textContent = "—";
        value.classList.add("is-empty");
      }
      field.append(label, value);
      trCardEl.appendChild(field);
    });
  }

  // Licznik, pasek postępu, stan ✓ bieżącego wiersza
  const pos = total ? trPos + 1 : 0;
  if (trCounterEl) trCounterEl.textContent = t("trCounter", { pos, total });
  if (trDoneCountEl) trDoneCountEl.textContent = t("trDoneCount", { done: trDone.size, all: trRows.length });
  if (trProgressBarEl) {
    const pct = trRows.length ? Math.round((trDone.size / trRows.length) * 100) : 0;
    trProgressBarEl.style.width = `${pct}%`;
  }
  const isDone = !!row && trDone.has(trKeyOf(row));
  if (trMarkChipEl) {
    trMarkChipEl.classList.toggle("is-done", isDone);
    trMarkChipEl.textContent = isDone ? t("trChipDone") : t("trChipPending");
    trMarkChipEl.setAttribute("aria-pressed", isDone ? "true" : "false");
    trMarkChipEl.disabled = !row;
  }
  if (trPrevBtn) trPrevBtn.disabled = !total || trPos <= 0;
  if (trNextBtn) trNextBtn.disabled = !total || trPos >= total - 1;
  if (trMarkBtn) trMarkBtn.disabled = !row;

  if (trLiveEl && row) trLiveEl.textContent = t("trLiveRow", { pos, total });
  trPersist();
}

// ── Nawigacja ───────────────────────────────────────────────────────────────

function trGo(delta) {
  if (!trOrder.length) return;
  const next = trPos + delta;
  if (next < 0 || next >= trOrder.length) return;
  trPos = next;
  trRenderCard();
}

function trGoEdge(which) {
  if (!trOrder.length) return;
  trPos = which < 0 ? 0 : trOrder.length - 1;
  trRenderCard();
}

function trToggleDone() {
  const row = trCurrentRow();
  if (!row) return;
  const key = trKeyOf(row);
  if (trDone.has(key)) trDone.delete(key);
  else trDone.add(key);
  if (trHideDone) {
    const keep = trPos;
    trRebuildOrder(null);
    trPos = Math.max(0, Math.min(keep, Math.max(0, trOrder.length - 1)));
  }
  trRenderCard();
}

// Główna akcja: odhacz i przejdź dalej. Przy „ukryj spisane” bieżący wiersz znika,
// więc pozycja ZOSTAJE na miejscu i sama pokazuje następny — bez przeskoku o dwa.
function trMarkAndNext() {
  const row = trCurrentRow();
  if (!row) return;
  trDone.add(trKeyOf(row));
  if (trHideDone) {
    const keep = trPos;
    trRebuildOrder(null);
    trPos = Math.max(0, Math.min(keep, Math.max(0, trOrder.length - 1)));
  } else if (trPos < trOrder.length - 1) {
    trPos += 1;
  }
  trRenderCard();
  if (trDone.size >= trRows.length && trRows.length) toast(t("trAllDone"), "success");
}

function trSetHideDone(on) {
  trHideDone = !!on;
  if (trHideDoneEl) trHideDoneEl.checked = trHideDone;
  trRebuildOrder(trHideDone ? null : trCurrentKey());
  trRenderCard();
}

function trResetProgress() {
  trDone.clear();
  trRebuildOrder(null);
  trPos = 0;
  trRenderCard();
  toast(t("trResetDone"), "success");
}

// Kasowanie ✓ to operacja nieodwracalna — dwustopniowy przycisk zamiast confirm(),
// żeby nie wyrywać użytkownika z pełnoekranowego trybu systemowym oknem.
function trArmReset() {
  if (!trResetBtn) return;
  if (trResetArmed) {
    clearTimeout(trResetTimer);
    trResetArmed = false;
    trResetBtn.classList.remove("is-armed");
    trResetBtn.textContent = t("trReset");
    trResetProgress();
    return;
  }
  trResetArmed = true;
  trResetBtn.classList.add("is-armed");
  trResetBtn.textContent = t("trResetConfirm");
  trResetTimer = setTimeout(() => {
    trResetArmed = false;
    trResetBtn.classList.remove("is-armed");
    trResetBtn.textContent = t("trReset");
  }, 3500);
}

function trSetAutoFields(on) {
  trAutoFields = !!on;
  if (trAutoFieldsEl) trAutoFieldsEl.checked = trAutoFields;
  if (trAutoNoteEl) trAutoNoteEl.classList.toggle("hidden", !trAutoFields);
  if (trFieldsBtn) trFieldsBtn.textContent = trAutoFields ? t("trFieldsAuto") : t("trFields");
  if (trFieldsListEl) trFieldsListEl.classList.toggle("is-auto", trAutoFields);
  trRenderCard();
}

// ── Rozmiar tekstu / blokada dotyku / Wake Lock ─────────────────────────────

function trApplyFont() {
  if (trOverlayEl) trOverlayEl.dataset.font = String(trFont);
  if (trFontBtn) trFontBtn.textContent = `A${"+".repeat(Math.max(0, trFont - 1))}`;
}

function trCycleFont() {
  const at = TR_FONT_STEPS.indexOf(trFont);
  trFont = TR_FONT_STEPS[(at + 1) % TR_FONT_STEPS.length];
  trApplyFont();
  trPersist();
}

// Blokada dotyku: tarcza przykrywa kartę (dłoń oparta o tablet nie przewinie ani nie
// zaznaczy), natomiast dolny pasek i sam przycisk blokady zostają nad nią klikalne.
function trSetLocked(on) {
  trLocked = !!on;
  if (trOverlayEl) trOverlayEl.classList.toggle("is-locked", trLocked);
  if (trTouchShieldEl) trTouchShieldEl.classList.toggle("hidden", !trLocked);
  if (trLockBtn) {
    trLockBtn.setAttribute("aria-pressed", trLocked ? "true" : "false");
    trLockBtn.textContent = trLocked ? t("trUnlock") : t("trLock");
  }
  if (trLocked && trFieldsPanelEl) trCloseFields();
}

async function trRequestWakeLock() {
  try {
    if (!("wakeLock" in navigator)) return;
    trWakeLock = await navigator.wakeLock.request("screen");
    trWakeLock.addEventListener("release", () => { trWakeLock = null; });
    if (trOverlayEl) trOverlayEl.classList.add("has-wakelock");
  } catch {
    /* brak zgody / nieobsługiwane — tryb działa, ekran po prostu może zgasnąć */
  }
}

function trReleaseWakeLock() {
  try { trWakeLock?.release?.(); } catch { /* ignore */ }
  trWakeLock = null;
  if (trOverlayEl) trOverlayEl.classList.remove("has-wakelock");
}

document.addEventListener("visibilitychange", () => {
  if (trIsOpen && document.visibilityState === "visible" && !trWakeLock) trRequestWakeLock();
});

// ── Panel „Pola” ────────────────────────────────────────────────────────────

function trRenderFields() {
  if (!trFieldsListEl) return;
  trFieldsListEl.replaceChildren();
  trFieldOrder.forEach((colIdx, pos) => {
    const item = document.createElement("div");
    item.className = "tr-field-row";

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.id = `trfield-${colIdx}`;
    cb.checked = trSelected.has(colIdx);
    cb.addEventListener("change", () => {
      if (cb.checked) trSelected.add(colIdx);
      else trSelected.delete(colIdx);
      trRenderCard();
    });

    const label = document.createElement("label");
    label.htmlFor = cb.id;
    label.className = "tr-field-name";
    label.textContent = trFieldLabel(colIdx);

    const actions = document.createElement("div");
    actions.className = "tr-field-actions";
    const up = document.createElement("button");
    up.type = "button";
    up.className = "btn btn-xs ghost";
    up.textContent = "▲";
    up.setAttribute("aria-label", `${t("moveUp")}: ${trFieldLabel(colIdx)}`);
    up.disabled = pos === 0;
    up.addEventListener("click", () => trMoveField(pos, -1));
    const down = document.createElement("button");
    down.type = "button";
    down.className = "btn btn-xs ghost";
    down.textContent = "▼";
    down.setAttribute("aria-label", `${t("moveDown")}: ${trFieldLabel(colIdx)}`);
    down.disabled = pos === trFieldOrder.length - 1;
    down.addEventListener("click", () => trMoveField(pos, 1));
    actions.append(up, down);

    item.append(cb, label, actions);
    trFieldsListEl.appendChild(item);
  });
  if (typeof ensureKeyboardReachable === "function") ensureKeyboardReachable(trFieldsListEl);
}

function trMoveField(pos, delta) {
  const next = pos + delta;
  if (next < 0 || next >= trFieldOrder.length) return;
  const [moved] = trFieldOrder.splice(pos, 1);
  trFieldOrder.splice(next, 0, moved);
  trRenderFields();
  trRenderCard();
  const list = trFieldsListEl.querySelectorAll(".tr-field-row");
  const btn = list[next]?.querySelector(`.tr-field-actions button:nth-child(${delta < 0 ? 1 : 2})`);
  if (btn && !btn.disabled) btn.focus();
}

function trOpenFields() {
  if (!trFieldsPanelEl) return;
  trRenderFields();
  trFieldsPanelEl.classList.remove("hidden");
  if (trFieldsBtn) trFieldsBtn.setAttribute("aria-expanded", "true");
  const first = trFieldsPanelEl.querySelector("input, button");
  if (first) first.focus();
}

function trCloseFields() {
  if (!trFieldsPanelEl) return;
  trFieldsPanelEl.classList.add("hidden");
  if (trFieldsBtn) {
    trFieldsBtn.setAttribute("aria-expanded", "false");
    trFieldsBtn.focus();
  }
  trRenderCard();
}

// ── Otwarcie / zamknięcie ───────────────────────────────────────────────────

function trBackgroundInert(on) {
  document.querySelectorAll(".app, .hero-overlay").forEach((el) => {
    if (on) el.setAttribute("inert", "");
    else el.removeAttribute("inert");
  });
}

function openTranscribe() {
  if (!trOverlayEl) return;
  const model = (typeof currentDisplayModel !== "undefined" && currentDisplayModel) || getDisplayModel();
  if (!model?.headers?.length || !model?.rows?.length) {
    toast(t("noDataForExport"), "warning");
    return;
  }

  trRows = model.rows.slice();
  trHeaders = model.headers.slice();
  trRowHeadFormatter = typeof model.rowHeadFormatter === "function" ? model.rowHeadFormatter : null;
  trScope = trScopeKey(model);

  const store = trLoadStore();
  const saved = store.scopes?.[trScope] || null;
  trFont = TR_FONT_STEPS.includes(store.font) ? store.font : 2;
  trHideDone = !!store.hideDone;

  const validCol = (i) => Number.isInteger(i) && i >= 0 && i < trHeaders.length;
  if (saved && Array.isArray(saved.order) && saved.order.length) {
    // Układ z poprzedniej sesji, ale arkusz mógł zmienić liczbę kolumn — dokładamy brakujące.
    trFieldOrder = saved.order.filter(validCol);
    trHeaders.forEach((_, i) => { if (!trFieldOrder.includes(i)) trFieldOrder.push(i); });
    trSelected = new Set((saved.sel || []).filter(validCol));
    if (!trSelected.size) trSelected = trDefaultSelection();
  } else {
    trFieldOrder = trHeaders.map((_, i) => i);
    trSelected = trDefaultSelection();
  }
  trDone = new Set(Array.isArray(saved?.done) ? saved.done : []);
  trAutoFields = !!saved?.auto;

  trRebuildOrder(null);
  // Wznowienie: wracamy na zapamiętany wiersz, a jak go nie ma — na pierwszy nieodhaczony.
  const resumeKey = saved?.cursor;
  let at = resumeKey ? trOrder.findIndex((i) => trKeyOf(trRows[i]) === resumeKey) : -1;
  if (at < 0) at = trOrder.findIndex((i) => !trDone.has(trKeyOf(trRows[i])));
  trPos = at >= 0 ? at : 0;

  if (trSourceEl) {
    const sheet = (typeof sheetSelect !== "undefined" && sheetSelect?.value) || "";
    trSourceEl.textContent = [currentFileName || "", sheet].filter(Boolean).join(" · ");
  }
  if (trHideDoneEl) trHideDoneEl.checked = trHideDone;
  trSetAutoFields(trAutoFields);
  trApplyFont();
  trSetLocked(false);
  if (trFieldsPanelEl) trFieldsPanelEl.classList.add("hidden");

  trReturnFocusEl = document.activeElement;
  trOverlayEl.classList.remove("hidden");
  document.body.classList.add("tr-active");
  trBackgroundInert(true);
  trIsOpen = true;
  trRenderCard();
  trRequestWakeLock();
  if (trMarkBtn) trMarkBtn.focus();
  if (trDone.size) toast(t("trResumed", { done: trDone.size }), "info");
}

function closeTranscribe() {
  if (!trOverlayEl || !trIsOpen) return;
  trPersist();
  trIsOpen = false;
  trSetLocked(false);
  trOverlayEl.classList.add("hidden");
  document.body.classList.remove("tr-active");
  trBackgroundInert(false);
  trReleaseWakeLock();
  const back = trReturnFocusEl;
  trReturnFocusEl = null;
  if (back && document.contains(back) && !back.closest("[inert]")) back.focus();
  else if (trBtn) trBtn.focus();
}

// ── Klawiatura ──────────────────────────────────────────────────────────────
// Capture na document: dopóki nakładka jest otwarta, klawisze NIE docierają do
// globalnego handlera w bootstrap.js (inaczej strzałki przesuwałyby zaznaczenie
// w tabeli pod spodem, a Cmd+Shift+F otwierałby okno szukania).

function trFocusables() {
  if (!trOverlayEl) return [];
  const sel = 'button:not([disabled]), input:not([disabled]), [tabindex]:not([tabindex="-1"])';
  return Array.from(trOverlayEl.querySelectorAll(sel)).filter((el) => {
    if (el.closest(".hidden")) return false;
    return el.offsetParent !== null || el === document.activeElement;
  });
}

function trTrapTab(e) {
  const list = trFocusables();
  if (!list.length) return;
  const first = list[0];
  const last = list[list.length - 1];
  if (e.shiftKey && document.activeElement === first) {
    e.preventDefault();
    last.focus();
  } else if (!e.shiftKey && document.activeElement === last) {
    e.preventDefault();
    first.focus();
  }
}

document.addEventListener("keydown", (e) => {
  if (!trIsOpen) return;
  e.stopPropagation();

  if (e.key === "Escape") {
    e.preventDefault();
    if (trFieldsPanelEl && !trFieldsPanelEl.classList.contains("hidden")) trCloseFields();
    else if (trLocked) trSetLocked(false);
    else closeTranscribe();
    return;
  }
  if (e.key === "Tab") {
    trTrapTab(e);
    return;
  }
  if (trFieldsPanelEl && trFieldsPanelEl.contains(e.target)) return;

  const tag = String(e.target?.tagName || "").toLowerCase();
  if (tag === "input" || tag === "select" || tag === "textarea") return;

  switch (e.key) {
    case "ArrowRight":
    case "PageDown":
      e.preventDefault();
      trGo(1);
      break;
    case "ArrowLeft":
    case "PageUp":
      e.preventDefault();
      trGo(-1);
      break;
    case "Home":
      e.preventDefault();
      trGoEdge(-1);
      break;
    case "End":
      e.preventDefault();
      trGoEdge(1);
      break;
    case " ":
    case "Enter":
      if (tag === "button") return; // niech ofokusowany przycisk zadziała sam
      e.preventDefault();
      trMarkAndNext();
      break;
    default:
      break;
  }
}, true);

// ── Podpięcie ───────────────────────────────────────────────────────────────

if (trBtn) trBtn.addEventListener("click", openTranscribe);
if (trCloseBtn) trCloseBtn.addEventListener("click", closeTranscribe);
if (trMarkBtn) trMarkBtn.addEventListener("click", trMarkAndNext);
if (trMarkChipEl) trMarkChipEl.addEventListener("click", trToggleDone);
if (trPrevBtn) trPrevBtn.addEventListener("click", () => trGo(-1));
if (trNextBtn) trNextBtn.addEventListener("click", () => trGo(1));
if (trHideDoneEl) trHideDoneEl.addEventListener("change", () => trSetHideDone(trHideDoneEl.checked));
if (trFontBtn) trFontBtn.addEventListener("click", trCycleFont);
if (trLockBtn) trLockBtn.addEventListener("click", () => trSetLocked(!trLocked));
if (trResetBtn) trResetBtn.addEventListener("click", trArmReset);
if (trAutoFieldsEl) trAutoFieldsEl.addEventListener("change", () => trSetAutoFields(trAutoFieldsEl.checked));
if (trFieldsBtn) {
  trFieldsBtn.addEventListener("click", () => {
    if (trFieldsPanelEl && trFieldsPanelEl.classList.contains("hidden")) trOpenFields();
    else trCloseFields();
  });
}
if (trFieldsDoneBtn) trFieldsDoneBtn.addEventListener("click", trCloseFields);
if (trFieldsAllBtn) {
  trFieldsAllBtn.addEventListener("click", () => {
    trFieldOrder.forEach((i) => trSelected.add(i));
    trRenderFields();
    trRenderCard();
  });
}
if (trFieldsNoneBtn) {
  trFieldsNoneBtn.addEventListener("click", () => {
    trSelected.clear();
    trRenderFields();
    trRenderCard();
  });
}
if (trTouchShieldEl) {
  // Tarcza połyka dotyk i klik — ale nie „na ślepo”: podwójny tap odblokowuje,
  // żeby nie dało się zamknąć w trybie bez wyjścia, gdy przycisk zniknie z pola widzenia.
  ["touchstart", "touchmove", "pointerdown", "click", "wheel"].forEach((evt) => {
    trTouchShieldEl.addEventListener(evt, (e) => { e.preventDefault(); e.stopPropagation(); }, { passive: false });
  });
  trTouchShieldEl.addEventListener("dblclick", () => trSetLocked(false));
}

// Hook testowy — Playwright steruje trybem bez klikania po pikselach.
window.__transcribe = {
  open: openTranscribe,
  close: closeTranscribe,
  state: () => ({
    open: trIsOpen,
    pos: trPos,
    total: trOrder.length,
    rows: trRows.length,
    done: trDone.size,
    cols: trVisibleCols(trCurrentRow()),
    auto: trAutoFields,
    skipped: trSkippedCount(trCurrentRow()),
    locked: trLocked,
    font: trFont,
    hideDone: trHideDone,
    values: (() => {
      const row = trCurrentRow();
      return row ? trVisibleCols(row).map((ci) => String(getDisplayValue(row, ci) ?? "")) : [];
    })(),
  }),
  mark: trMarkAndNext,
  go: trGo,
  setHideDone: trSetHideDone,
  setAutoFields: trSetAutoFields,
  reset: trResetProgress,
  setFields: (cols) => {
    trSelected = new Set(cols);
    trRenderCard();
  },
  moveField: trMoveField,
};
