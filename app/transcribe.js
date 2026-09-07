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
const trInheritEl = document.getElementById("trInherit");
const trInheritNoteEl = document.getElementById("trInheritNote");
const trTouchShieldEl = document.getElementById("trTouchShield");
const trStageEl = document.getElementById("trStage");
const trStageWrapEl = document.getElementById("trStageWrap");
const trScrollRailEl = document.getElementById("trScrollRail");
const trScrollThumbEl = document.getElementById("trScrollThumb");
const trScrollMoreEl = document.getElementById("trScrollMore");
const trScrollMoreTextEl = document.getElementById("trScrollMoreText");
const trUndoBtn = document.getElementById("trUndoBtn");
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
let trInheritOn = false;          // „dziedzicz z góry" — master switch
let trInheritCols = new Set();    // kolumny, które przenoszą wartość w dół
let trMergeCols = new Set();      // kolumny objęte PIONOWYM scaleniem (auto-podpowiedź + znacznik)
let trMergeRanges = new Map();    // kolumna -> zakresy scaleń = GRANICE przenoszenia
let trInheritMap = new Map();     // `col:rowIndex0` -> { text, from }
let trLongMode = false;           // Wide-to-Long: dziedziczenie nie ma tam sensu
let trFont = 2;
let trLocked = false;
let trScope = "";
let trWakeLock = null;
let trReturnFocusEl = null;
let trResetArmed = false;
let trResetTimer = 0;
let trBulkMode = false;       // seria szybkiego odhaczania — wstrzymuje zapis do localStorage
let trScrollRaf = 0;
let trHoldTimer = 0;          // odliczanie do startu turbo (przytrzymanie)
let trHoldProgressTimer = 0;  // animacja paska „ładowania" przytrzymania
let trTurboTimer = 0;         // pętla szybkiego odhaczania
let trTurboCount = 0;
let trTurboSource = "";       // "key" albo "pointer" — kto trzyma, ten puszcza
let trBurstKeys = [];         // klucze odhaczone w ostatniej serii (do cofnięcia)
let trUndoTimer = 0;

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
  if (!trScope || trBulkMode) return; // w trakcie serii zapisujemy RAZ, na końcu
  const store = trLoadStore();
  store.font = trFont;
  store.hideDone = trHideDone;
  store.scopes[trScope] = {
    order: trFieldOrder.slice(),
    sel: Array.from(trSelected),
    auto: trAutoFields,
    inherit: trInheritOn,
    inheritCols: Array.from(trInheritCols),
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
// ── Dziedziczenie z góry (scalone komórki / „wartość tylko w pierwszym wierszu") ──
//
// Bardzo częsty układ: nazwisko scalone przez 5 wierszy, pod spodem pozycje. W danych
// wiersze-kontynuacje są PUSTE (build-rows-core czyta `!merges` tylko po to, żeby
// wyznaczyć zakres arkusza — wartości nie rozlewa), więc bez tego mechanizmu tryb
// „dobieraj pola z wiersza" ukrywałby pole tożsamości akurat tam, gdzie jest najbardziej
// potrzebne.
//
// DWIE DECYZJE PROJEKTOWE, które trzymają to w ryzach:
//  1. Liczymy po `baseRows` — pełnym, NIEPRZEFILTROWANYM zestawie w kolejności arkusza.
//     Liczenie po widoku dawałoby inny wynik po filtrze albo sortowaniu, a „wiersz wyżej"
//     to własność PLIKU, nie bieżącego widoku.
//  2. Kolumny wskazuje użytkownik (scalenia tylko je podpowiadają). Zgadywanie „ta kolumna
//     chyba się przenosi" mogłoby wpisać cudzą wartość do rubryki na papierze, a tego się
//     nie cofa gumką.
// Pionowe scalenia arkusza, pogrupowane po kolumnie MODELU. Trzymamy pełne zakresy,
// nie same numery kolumn, bo zakres jest jednocześnie GRANICĄ przenoszenia.
function trDetectMergeRanges() {
  const byCol = new Map();
  if (trLongMode) return byCol;
  try {
    const sheet = workbook?.Sheets?.[currentSheetName];
    const merges = Array.isArray(sheet?.["!merges"]) ? sheet["!merges"] : [];
    const startCol = Number.isFinite(currentStartCol) ? currentStartCol : 0;
    merges.forEach((m) => {
      if (!m?.s || !m?.e || m.e.r <= m.s.r) return; // interesują nas tylko scalenia PIONOWE
      for (let c = m.s.c; c <= m.e.c; c++) {
        const col = c - startCol;
        if (col < 0 || col >= trHeaders.length) continue;
        if (!byCol.has(col)) byCol.set(col, []);
        byCol.get(col).push({ start: m.s.r, end: m.e.r });
      }
    });
  } catch {
    /* brak dostępu do arkusza — zostaje ręczny wybór kolumn */
  }
  return byCol;
}

// DWIE REGUŁY, świadomie różne — bo różna jest pewność, skąd bierze się pustka:
//
//  • kolumna ZE SCALENIAMI → przenosimy DOKŁADNIE w granicach scalenia. Plik sam mówi,
//    dokąd sięga rekord, więc nie ma miejsca na wyciek wartości do następnego rekordu.
//  • kolumna wskazana RĘCZNIE (bez scaleń) → najbliższa wartość powyżej. Tu granicy
//    rekordu nikt nie zapisał, więc to świadomy wybór użytkownika; wartość zawsze
//    dostaje na karcie numer wiersza źródłowego, żeby dało się ją sprawdzić.
function trBuildInheritance() {
  trInheritMap.clear();
  if (!trInheritOn || trLongMode || !trInheritCols.size) return;
  const source = Array.isArray(baseRows) ? baseRows : [];
  if (!source.length) return;

  const byRowIndex = new Map();
  source.forEach((row) => { byRowIndex.set(row.rowIndex0, row); });

  const mergeCols = [];
  const carryCols = [];
  trInheritCols.forEach((col) => {
    if (trMergeRanges.has(col)) mergeCols.push(col);
    else carryCols.push(col);
  });

  // 1) Kolumny ze scaleniami — zakres po zakresie, kotwicą jest pierwszy wiersz scalenia.
  mergeCols.forEach((col) => {
    (trMergeRanges.get(col) || []).forEach((range) => {
      const anchorRow = byRowIndex.get(range.start);
      if (!anchorRow) return; // kotwica nad wierszem nagłówka albo poza zakresem danych
      const text = String(getDisplayValue(anchorRow, col) ?? "").trim();
      if (!text) return;
      const from = (range.start ?? 0) + 1;
      for (let r = range.start + 1; r <= range.end; r++) {
        const row = byRowIndex.get(r);
        if (!row) continue;
        if (String(getDisplayValue(row, col) ?? "").trim()) continue; // własna wartość wygrywa
        trInheritMap.set(`${col}:${r}`, { text, from });
      }
    });
  });

  // 2) Kolumny wskazane ręcznie — jeden przebieg po wierszach dla wszystkich naraz.
  if (!carryCols.length) return;
  const carry = new Array(carryCols.length).fill(null);
  for (const row of source) {
    // Podnagłówek = granica sekcji, więc zrywa przenoszenie. UWAGA: markSubheaderRows
    // sprawdza tylko pierwsze wiersze arkusza, więc to zabezpieczenie łapie nagłówki
    // sekcji u góry pliku, a nie jest pełnym wykrywaniem rekordów.
    if (row.isSubheader) {
      carry.fill(null);
      continue;
    }
    const rowIdx = row.rowIndex0;
    for (let k = 0; k < carryCols.length; k++) {
      const col = carryCols[k];
      const txt = String(getDisplayValue(row, col) ?? "").trim();
      if (txt) carry[k] = { text: txt, from: (rowIdx ?? 0) + 1 };
      else if (carry[k]) trInheritMap.set(`${col}:${rowIdx}`, carry[k]);
    }
  }
}

// Jedno miejsce, które odpowiada „co ma stanąć w tej rubryce”: wartość własna wiersza,
// a jak jej nie ma — odziedziczona (z numerem wiersza źródłowego, żeby dało się sprawdzić).
function trResolveField(row, idx) {
  const raw = String(getDisplayValue(row, idx) ?? "").trim();
  if (raw) return { text: raw, from: 0 };
  if (!trInheritOn || trLongMode || !trInheritCols.has(idx)) return { text: "", from: 0 };
  const hit = trInheritMap.get(`${idx}:${row?.rowIndex0}`);
  return hit ? { text: hit.text, from: hit.from } : { text: "", from: 0 };
}

function trHasValue(row, idx) {
  return trResolveField(row, idx).text !== "";
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

// ── Belka przewijania / „jest tego więcej" ──────────────────────────────────
//
// Problem z tabletu: karta z kilkunastoma polami nie mieści się na ekranie, a jedyną
// informacją o tym jest natywny pasek, który na iPadOS pojawia się DOPIERO w trakcie
// przewijania. Przy przepisywaniu na papier to realna pomyłka — pole 12 zostaje
// nieprzepisane, bo nikt nie wiedział, że istnieje.
//
// Stąd trzy sygnały naraz: własna belka z boku (jest tu w ogóle co przewijać?),
// cienie-krawędzie (treść jest ucięta w tę stronę) i pigułka z LICZBĄ pól poniżej
// (ile dokładnie zostało) — pigułka jest jednocześnie przyciskiem „przewiń o ekran".

function trFieldsBelowFold() {
  if (!trStageEl) return 0;
  const foldY = trStageEl.getBoundingClientRect().bottom;
  let n = 0;
  trCardEl?.querySelectorAll(".tr-field").forEach((el) => {
    // liczymy pole jako „poniżej”, gdy jego etykieta i wartość nie są w całości widoczne
    if (el.getBoundingClientRect().bottom > foldY + 2) n += 1;
  });
  return n;
}

function trUpdateScrollUi() {
  if (!trStageEl || !trStageWrapEl) return;
  const max = trStageEl.scrollHeight - trStageEl.clientHeight;
  const overflow = max > 4;
  const top = trStageEl.scrollTop;
  trStageWrapEl.classList.toggle("has-overflow", overflow);
  trStageWrapEl.classList.toggle("at-top", top <= 2);
  trStageWrapEl.classList.toggle("at-bottom", !overflow || top >= max - 2);

  if (trScrollThumbEl && trScrollRailEl) {
    const railH = trScrollRailEl.clientHeight;
    const ratio = trStageEl.scrollHeight ? trStageEl.clientHeight / trStageEl.scrollHeight : 1;
    const thumbH = Math.max(22, Math.round(railH * Math.min(1, ratio)));
    const travel = Math.max(0, railH - thumbH);
    const progress = max > 0 ? Math.min(1, Math.max(0, top / max)) : 0;
    trScrollThumbEl.style.height = `${thumbH}px`;
    trScrollThumbEl.style.transform = `translateY(${Math.round(travel * progress)}px)`;
  }

  if (trScrollMoreTextEl) {
    const below = overflow ? trFieldsBelowFold() : 0;
    // Gdy pole jest jedno, ale bardzo wysokie, licznik pokazałby „0” — wtedy sam napis.
    trScrollMoreTextEl.textContent = below > 0 ? t("trScrollMore", { n: below }) : t("trScrollMoreMore");
  }
}

function trScheduleScrollUi() {
  if (trScrollRaf) return;
  trScrollRaf = requestAnimationFrame(() => {
    trScrollRaf = 0;
    trUpdateScrollUi();
  });
}

// Nowy wiersz = nowa kartka: zawsze zaczynamy od GÓRY. Bez tego po przewinięciu
// długiego wiersza następny otwierał się w połowie i pierwsze pola uciekały nad ekran.
function trResetScroll() {
  if (!trStageEl) return;
  trStageEl.scrollTop = 0;
  trScheduleScrollUi();
}

if (trStageEl) trStageEl.addEventListener("scroll", trScheduleScrollUi, { passive: true });
if (trScrollMoreEl) {
  trScrollMoreEl.addEventListener("click", () => {
    if (!trStageEl) return;
    const step = Math.max(120, Math.round(trStageEl.clientHeight * 0.82));
    trStageEl.scrollBy({ top: step, behavior: "smooth" });
  });
}
window.addEventListener("resize", () => { if (trIsOpen) trScheduleScrollUi(); });

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
      const labelText = document.createElement("span");
      labelText.textContent = trFieldLabel(ci);
      label.appendChild(labelText);
      const resolved = trResolveField(row, ci);
      if (resolved.from) {
        // Wartość odziedziczona MUSI być rozpoznawalna — inaczej przepisze się ją
        // jak własną i nie da się już wychwycić pomyłki.
        field.classList.add("is-inherited");
        const badge = document.createElement("span");
        badge.className = "tr-inherited-from";
        badge.textContent = t("trInheritedFrom", { n: resolved.from });
        label.appendChild(badge);
      }
      const value = document.createElement("div");
      value.className = "tr-field-value";
      if (resolved.text) {
        value.textContent = resolved.text;
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
  // Pomiar po wstawieniu pól do DOM — inaczej scrollHeight jest jeszcze sprzed renderu.
  trScheduleScrollUi();
  trPersist();
}

// ── Nawigacja ───────────────────────────────────────────────────────────────

function trGo(delta) {
  if (!trOrder.length) return;
  const next = trPos + delta;
  if (next < 0 || next >= trOrder.length) return;
  trPos = next;
  trResetScroll();
  trRenderCard();
}

function trGoEdge(which) {
  if (!trOrder.length) return;
  trPos = which < 0 ? 0 : trOrder.length - 1;
  trResetScroll();
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
// Zwraca klucz odhaczonego wiersza (albo "" gdy nie było czego odhaczyć) — potrzebne
// szybkiemu odhaczaniu, żeby dało się całą serię cofnąć jednym ruchem.
function trMarkAndNext(options = {}) {
  const row = trCurrentRow();
  if (!row) return "";
  const key = trKeyOf(row);
  const wasDone = trDone.has(key);
  trDone.add(key);
  const before = trPos;
  if (trHideDone) {
    const keep = trPos;
    trRebuildOrder(null);
    trPos = Math.max(0, Math.min(keep, Math.max(0, trOrder.length - 1)));
  } else if (trPos < trOrder.length - 1) {
    trPos += 1;
  }
  const moved = trHideDone ? true : trPos !== before;
  trResetScroll();
  trRenderCard();
  if (!options.quiet && trDone.size >= trRows.length && trRows.length) toast(t("trAllDone"), "success");
  return { key, wasDone, moved };
}

// ── Szybkie odhaczanie (przytrzymanie) ──────────────────────────────────────
//
// Po co: apka w tle na tablecie potrafi zostać ubita, a po powrocie trzeba dojść do
// miejsca sprzed przerwy. Klikanie „Spisane i dalej” 200 razy to nie jest plan.
// PRZYTRZYMANIE (palcem na przycisku albo spacją, gdy przycisk ma fokus) rozpędza
// odhaczanie: po ~0,55 s startuje pętla, która przyspiesza z 200 ms do 60 ms na wiersz.
//
// Trzy bezpieczniki, bo to operacja masowa:
//  1. próg czasu — zwykły tap NIGDY nie wejdzie w tryb szybki,
//  2. pasek na przycisku pokazuje, ile zostało do startu (nic nie dzieje się „nagle”),
//  3. cała seria cofa się jednym przyciskiem w pasku meta (i wraca na wiersz startowy).
const TR_HOLD_MS = 550;        // ile trzymać, zanim ruszy tryb szybki
const TR_TURBO_START_MS = 200; // pierwszy krok
const TR_TURBO_MIN_MS = 45;    // najszybszy krok
const TR_TURBO_ACCEL = 0.86;   // mnożnik między krokami

let trBurstStartKey = "";
let trSuppressNextClick = false;

// Licznik „+N" na przycisku żyje w custom property, bo tekst przycisku podmienia i18n.
function trSetTurboLabel(text) {
  if (!trMarkBtn) return;
  if (text) trMarkBtn.style.setProperty("--tr-turbo-label", JSON.stringify(text));
  else trMarkBtn.style.removeProperty("--tr-turbo-label");
}

function trSetHoldProgress(pct) {
  if (trMarkBtn) trMarkBtn.style.setProperty("--tr-hold", `${Math.round(pct)}%`);
}

function trHoldStart(source) {
  if (!trIsOpen || trTurboTimer || trHoldTimer) return;
  if (!trCurrentRow()) return;
  trTurboSource = source;
  const began = Date.now();
  trHoldProgressTimer = setInterval(() => {
    trSetHoldProgress(Math.min(100, ((Date.now() - began) / TR_HOLD_MS) * 100));
  }, 60);
  trHoldTimer = setTimeout(() => {
    trHoldTimer = 0;
    trTurboStart();
  }, TR_HOLD_MS);
}

// Zatrzymanie odliczania. `fired` = czy zdążył wystartować tryb szybki — od tego zależy,
// czy puszczenie klawisza/palca ma jeszcze odhaczyć pojedynczy wiersz.
function trHoldCancel() {
  const fired = !!trTurboTimer;
  if (trHoldTimer) { clearTimeout(trHoldTimer); trHoldTimer = 0; }
  if (trHoldProgressTimer) { clearInterval(trHoldProgressTimer); trHoldProgressTimer = 0; }
  trSetHoldProgress(0);
  if (fired) trTurboStop();
  trTurboSource = "";
  return fired;
}

function trTurboStart() {
  if (trHoldProgressTimer) { clearInterval(trHoldProgressTimer); trHoldProgressTimer = 0; }
  trSetHoldProgress(100);
  trTurboCount = 0;
  trBurstKeys = [];
  trBurstStartKey = trCurrentKey();
  // Zapis stanu to serializacja całego zbioru ✓ (do 20 tys. kluczy). Przy 20 wierszach
  // na sekundę robiłoby to z tabletu podkładkę pod kawę — zapisujemy raz, po serii.
  trBulkMode = true;
  if (trMarkBtn) trMarkBtn.classList.add("is-turbo");
  if (typeof navigator !== "undefined" && navigator.vibrate) { try { navigator.vibrate(12); } catch { /* brak wsparcia */ } }
  toast(t("trTurboStarted"), "info");

  let delay = TR_TURBO_START_MS;
  const step = () => {
    const res = trMarkAndNext({ quiet: true });
    if (!res || !res.key) { trTurboStop(); return; }
    if (!res.wasDone) trBurstKeys.push(res.key);
    trTurboCount += 1;
    trSetTurboLabel(`+${trTurboCount}`);
    if (!res.moved) { // koniec listy — dalej nie ma dokąd
      toast(t("trTurboEnd"), "info");
      trTurboStop();
      return;
    }
    delay = Math.max(TR_TURBO_MIN_MS, Math.round(delay * TR_TURBO_ACCEL));
    trTurboTimer = setTimeout(step, delay);
  };
  trTurboTimer = setTimeout(step, 0);
}

function trTurboStop() {
  if (trTurboTimer) { clearTimeout(trTurboTimer); trTurboTimer = 0; }
  const wasBulk = trBulkMode;
  trBulkMode = false;
  if (wasBulk) trPersist();
  if (trMarkBtn) trMarkBtn.classList.remove("is-turbo");
  trSetTurboLabel("");
  trSetHoldProgress(0);
  if (trTurboCount > 0) {
    toast(t("trTurboDone", { n: trTurboCount }), "success");
    trShowUndo(trBurstKeys.length);
  }
  trTurboCount = 0;
}

// Cofnięcie serii: zdejmujemy ✓ tylko z wierszy odhaczonych W TEJ serii (te, które
// były odhaczone wcześniej, zostają) i wracamy kursorem na wiersz startowy.
function trShowUndo(n) {
  if (!trUndoBtn) return;
  if (!n) { trHideUndo(); return; }
  trUndoBtn.textContent = t("trUndoBurst", { n });
  trUndoBtn.classList.remove("hidden");
  if (trUndoTimer) clearTimeout(trUndoTimer);
  trUndoTimer = setTimeout(trHideUndo, 12000);
}

function trHideUndo() {
  if (trUndoTimer) { clearTimeout(trUndoTimer); trUndoTimer = 0; }
  if (trUndoBtn) trUndoBtn.classList.add("hidden");
  trBurstKeys = [];
  trBurstStartKey = "";
}

function trUndoBurst() {
  if (!trBurstKeys.length) { trHideUndo(); return; }
  const n = trBurstKeys.length;
  const back = trBurstStartKey;
  trBurstKeys.forEach((key) => trDone.delete(key));
  trHideUndo();
  trRebuildOrder(back || null);
  trResetScroll();
  trRenderCard();
  toast(t("trTurboUndone", { n }), "success");
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

function trUpdateInheritNote() {
  if (trInheritNoteEl) {
    let note;
    if (trLongMode) note = t("trInheritLong");
    else if (trMergeCols.size) note = t("trInheritMergeInfo", { cols: trMergeCols.size });
    else note = t("trInheritNoMerges");
    trInheritNoteEl.textContent = note;
  }
  if (trInheritEl) trInheritEl.disabled = trLongMode;
}

function trSetInherit(on, options = {}) {
  trInheritOn = !!on && !trLongMode;
  // Pierwsze włączenie na arkuszu ze scaleniami samo proponuje kolumny — użytkownik
  // i tak może je dowolnie zmienić, ale nie musi zaczynać od pustej listy.
  if (trInheritOn && !trInheritCols.size && trMergeCols.size && !options.keepCols) {
    trInheritCols = new Set(trMergeCols);
  }
  if (trInheritEl) trInheritEl.checked = trInheritOn;
  trUpdateInheritNote();
  trBuildInheritance();
  if (trFieldsListEl && !trFieldsPanelEl?.classList.contains("hidden")) trRenderFields();
  if (options.silent) return;
  trRenderCard();
}

function trToggleInheritCol(colIdx) {
  if (!trInheritOn || trLongMode) return;
  if (trInheritCols.has(colIdx)) trInheritCols.delete(colIdx);
  else trInheritCols.add(colIdx);
  trBuildInheritance();
  trRenderFields();
  trRenderCard();
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
  trScheduleScrollUi(); // większa czcionka = karta może przestać się mieścić
  trPersist();
}

// Blokada dotyku: tarcza przykrywa kartę (dłoń oparta o tablet nie przewinie ani nie
// zaznaczy), natomiast dolny pasek i sam przycisk blokady zostają nad nią klikalne.
function trSetLocked(on) {
  if (on && trTurboSource) trHoldCancel(); // blokada dotyku w trakcie serii = stop
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

    const inh = document.createElement("button");
    inh.type = "button";
    inh.className = "btn btn-xs ghost tr-inherit-btn";
    inh.textContent = "⤓";
    inh.disabled = !trInheritOn || trLongMode;
    inh.classList.toggle("is-on", trInheritCols.has(colIdx));
    inh.classList.toggle("is-merge", trMergeCols.has(colIdx));
    inh.setAttribute("aria-pressed", trInheritCols.has(colIdx) ? "true" : "false");
    inh.setAttribute("aria-label", `${t("trInheritColAria")}: ${trFieldLabel(colIdx)}`);
    inh.addEventListener("click", () => trToggleInheritCol(colIdx));
    actions.appendChild(inh);

    const up = document.createElement("button");
    up.type = "button";
    up.className = "btn btn-xs ghost tr-move-up";
    up.textContent = "▲";
    up.setAttribute("aria-label", `${t("moveUp")}: ${trFieldLabel(colIdx)}`);
    up.disabled = pos === 0;
    up.addEventListener("click", () => trMoveField(pos, -1));
    const down = document.createElement("button");
    down.type = "button";
    down.className = "btn btn-xs ghost tr-move-down";
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
  const btn = list[next]?.querySelector(`.tr-field-actions .tr-move-${delta < 0 ? "up" : "down"}`);
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
  trLongMode = model.mode === "long";
  trMergeRanges = trDetectMergeRanges();
  trMergeCols = new Set(trMergeRanges.keys());
  const validCol2 = (i) => Number.isInteger(i) && i >= 0 && i < trHeaders.length;
  trInheritCols = new Set(Array.isArray(saved?.inheritCols) ? saved.inheritCols.filter(validCol2) : []);
  trInheritOn = !!saved?.inherit && !trLongMode;

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
  trSetInherit(trInheritOn, { keepCols: true, silent: true });
  trApplyFont();
  trSetLocked(false);
  if (trFieldsPanelEl) trFieldsPanelEl.classList.add("hidden");

  trReturnFocusEl = document.activeElement;
  trOverlayEl.classList.remove("hidden");
  document.body.classList.add("tr-active");
  trBackgroundInert(true);
  trIsOpen = true;
  trHideUndo();
  trResetScroll();
  trRenderCard();
  trRequestWakeLock();
  if (trMarkBtn) trMarkBtn.focus();
  if (trDone.size) toast(t("trResumed", { done: trDone.size }), "info");
}

function closeTranscribe() {
  if (!trOverlayEl || !trIsOpen) return;
  trHoldCancel();
  trBulkMode = false;
  trHideUndo();
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
    if (el.getAttribute("tabindex") === "-1") return false; // np. pigułka „więcej ↓" — klikalna, ale poza Tabem
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
    case "Enter": {
      // Na „Spisane i dalej" spacja jest PRZYTRZYMYWALNA (szybkie odhaczanie), więc
      // blokujemy natywną aktywację przycisku i sami decydujemy przy puszczeniu klawisza.
      const onMark = e.target === trMarkBtn;
      if (tag === "button" && !onMark) return; // inny przycisk niech zadziała sam
      e.preventDefault();
      if (!e.repeat) trHoldStart("key");
      break;
    }
    default:
      break;
  }
}, true);

// Puszczenie klawisza: albo kończy tryb szybki, albo — gdy nie zdążył wystartować —
// wykonuje zwykłe „Spisane i dalej". Capture, bo keydown wyżej też jest w capture.
document.addEventListener("keyup", (e) => {
  if (!trIsOpen || trTurboSource !== "key") return;
  if (e.key !== " " && e.key !== "Enter") return;
  e.stopPropagation();
  e.preventDefault();
  const fired = trHoldCancel();
  if (!fired) trMarkAndNext();
}, true);

// Utrata fokusu / przejście w tło w trakcie trzymania — nie zostawiamy pętli w biegu.
window.addEventListener("blur", () => { if (trTurboSource) trHoldCancel(); });
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState !== "visible" && trTurboSource) trHoldCancel();
});

// ── Podpięcie ───────────────────────────────────────────────────────────────

if (trBtn) trBtn.addEventListener("click", openTranscribe);
if (trCloseBtn) trCloseBtn.addEventListener("click", closeTranscribe);
if (trMarkBtn) {
  trMarkBtn.addEventListener("click", () => {
    // Klik doleci też po zakończeniu przytrzymania — wtedy go połykamy, żeby seria
    // nie dostała jednego wiersza w bonusie.
    if (trSuppressNextClick) { trSuppressNextClick = false; return; }
    trMarkAndNext();
  });
  trMarkBtn.addEventListener("pointerdown", (e) => {
    if (e.button != null && e.button > 0) return; // tylko lewy / dotyk / pióro
    trHoldStart("pointer");
  });
  const endPointerHold = () => {
    if (trTurboSource !== "pointer") return;
    if (trHoldCancel()) trSuppressNextClick = true;
  };
  ["pointerup", "pointercancel", "pointerleave"].forEach((evt) => trMarkBtn.addEventListener(evt, endPointerHold));
  // Palec puszczony poza przyciskiem (albo mysz zwolniona gdzie indziej) też kończy serię.
  window.addEventListener("pointerup", endPointerHold);
}
if (trUndoBtn) trUndoBtn.addEventListener("click", trUndoBurst);
if (trMarkChipEl) trMarkChipEl.addEventListener("click", trToggleDone);
if (trPrevBtn) trPrevBtn.addEventListener("click", () => trGo(-1));
if (trNextBtn) trNextBtn.addEventListener("click", () => trGo(1));
if (trHideDoneEl) trHideDoneEl.addEventListener("change", () => trSetHideDone(trHideDoneEl.checked));
if (trFontBtn) trFontBtn.addEventListener("click", trCycleFont);
if (trLockBtn) trLockBtn.addEventListener("click", () => trSetLocked(!trLocked));
if (trResetBtn) trResetBtn.addEventListener("click", trArmReset);
if (trAutoFieldsEl) trAutoFieldsEl.addEventListener("change", () => trSetAutoFields(trAutoFieldsEl.checked));
if (trInheritEl) trInheritEl.addEventListener("change", () => trSetInherit(trInheritEl.checked));
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
    inherit: trInheritOn,
    inheritCols: Array.from(trInheritCols),
    mergeCols: Array.from(trMergeCols),
    longMode: trLongMode,
    fields: (() => {
      const row = trCurrentRow();
      if (!row) return [];
      return trVisibleCols(row).map((ci) => {
        const r = trResolveField(row, ci);
        return { col: ci, label: trFieldLabel(ci), text: r.text, from: r.from };
      });
    })(),
    locked: trLocked,
    font: trFont,
    hideDone: trHideDone,
    scrollTop: trStageEl ? trStageEl.scrollTop : 0,
    canScroll: !!(trStageEl && trStageEl.scrollHeight - trStageEl.clientHeight > 4),
    overflowUi: !!trStageWrapEl?.classList.contains("has-overflow"),
    atBottom: !!trStageWrapEl?.classList.contains("at-bottom"),
    turbo: !!trTurboTimer,
    undoVisible: !!(trUndoBtn && !trUndoBtn.classList.contains("hidden")),
    burst: trBurstKeys.length,
    values: (() => {
      const row = trCurrentRow();
      return row ? trVisibleCols(row).map((ci) => String(getDisplayValue(row, ci) ?? "")) : [];
    })(),
  }),
  mark: trMarkAndNext,
  holdStart: (source = "key") => trHoldStart(source),
  holdCancel: () => trHoldCancel(),
  undoBurst: trUndoBurst,
  scrollBy: (px) => { if (trStageEl) { trStageEl.scrollTop += px; trUpdateScrollUi(); } },
  go: trGo,
  setHideDone: trSetHideDone,
  setAutoFields: trSetAutoFields,
  setInherit: (on) => trSetInherit(on),
  toggleInheritCol: trToggleInheritCol,
  reset: trResetProgress,
  setFields: (cols) => {
    trSelected = new Set(cols);
    trRenderCard();
  },
  moveField: trMoveField,
};
