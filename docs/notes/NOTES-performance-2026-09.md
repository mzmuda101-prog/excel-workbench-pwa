# Przejście wydajnościowe — wrzesień 2026 (urządzenia słabe / mobilne / obciążone)

> Kontekst: temat „przyspieszyć apkę na starych i mobilnych urządzeniach, bez psucia funkcji”.
> Uzupełnia `NOTES-performance-plan-2026-06.md` — tamten plan opisywał Fazy 1–4, ten dokument
> opisuje pomiary i zmiany zrobione **po** tamtych decyzjach.

## Jak mierzone

- Playwright + CDP: dławienie CPU (`Emulation.setCPUThrottlingRate`) 1× / 4× / 6×.
- Profile: `iPad Pro 11` (WebKit-like viewport + dotyk), `Pixel 5`, desktop 1440×900.
- Pliki: `scripts/stress-test-workbench.xlsx` (~2000×40), `RODO Obieg terenów 2026 V5.xlsx` (610 KB),
  `Kosztorys Renault Laguna 3 .xlsx`.
- Mierzone: czas do 2× rAF po akcji („paint”), suma `longtask` w oknie akcji, klatki przy scrollu,
  heap po wymuszonym GC (`HeapProfiler.collectGarbage`) — nie surowy `usedJSHeapSize`,
  bo tuż po ciężkiej pracy pokazuje głównie nieposprzątane śmieci.

## Wyniki (iPad @4× CPU, stress-test-workbench.xlsx)

| Akcja | Przed | Po | Zmiana |
|---|---|---|---|
| Otwarcie panelu agregacji | 3654 ms | 1338 ms | **−63%** |
| „Filtruj” przy OTWARTYCH panelach analiz | 4287 ms | 923 ms | **−78%** |
| „Filtruj” przy zamkniętych panelach | 562 ms | 494 ms | −12% |
| „Wyczyść filtry” | 449 ms | 402 ms | −10% |
| Wybór pliku → lista arkuszy | 1575 ms | 1277 ms | −19% |
| Heap zatrzymany po GC | 30 MB | 30 MB | bez zmian |

Wczytanie pliku RODO (@4× CPU), rozbicie fazy „plik → lista arkuszy”: 3151 ms → 2395 ms (−24%).

## Co było nie tak (przyczyny, nie objawy)

1. **`classifyAggregationHeader()` wołane w pętli po wierszach** (`collectAggregationProfiles`),
   choć zależy wyłącznie od nagłówka kolumny → **440 532 wywołania** na jedno otwarcie panelu.
2. **`parseDateFlexible()` budowało ~50-kluczowy `monthMap` przy każdym wywołaniu** —
   a wołane jest raz na komórkę, setki tysięcy razy.
3. **Autodetekcja wiersza nagłówka robiła pełny `buildRows()` + `detectRepeatingBlocks()`
   dla ~8 kandydatów — przy KAŻDYM renderze panelu**, czyli także po każdym filtrze i sorcie.
   Stąd absurd: filtrowanie było wolniejsze niż pierwsze otwarcie panelu.
4. **Ten sam plik rozpakowywany 3× niezależnie** przy wczytywaniu (mapa stylów, formatowanie
   warunkowe, data validation) — każda z tych funkcji dekompresowała XML **każdego arkusza** osobno.
5. **Data liczona zachłannie** w profilowaniu kolumn — dla zwykłych liczb w kolumnie nie-czasowej
   wynik i tak był odrzucany.

## Co zmienione

| # | Zmiana | Plik |
|---|--------|------|
| 1 | Wyniesienie `classifyAggregationHeader()` z pętli po wierszach (raz na kolumnę) | `app/analysis.js` |
| 2 | Memoizacja `classifyAggregationHeader()` i `normalizeAnalysisKey()` (limit + czyszczenie) | `app/analysis.js` |
| 3 | `monthMap` → stała modułu `DATE_MONTH_MAP` + memoizacja `parseDateFlexible` (cache znacznika czasu, zwracany świeży `Date`) | `app/workbook.js` |
| 4 | Leniwe liczenie daty w `collectAggregationProfiles` | `app/analysis.js` |
| 5 | Cache skanu nagłówków (`buildRows` + `detectRepeatingBlocks`) per (arkusz, wiersz nagłówka), unieważniany znacznikiem `sheetDataStamp` i językiem | `app/table.js` |
| 6 | `sheetDataStamp` + `bumpSheetDataStamp()` — bump przy wczytaniu pliku i każdej edycji komórki | `app/core.js`, `app/workbook.js`, `app/ui-controls.js` |
| 7 | Jedna instancja JSZip na plik (`getSharedWorkbookZip`) + cache rozpakowanych wpisów wiązany z instancją (`WeakMap`), zwalniany po wczytaniu | `app/core.js` + 3 konsumentów |

### Dlaczego memoizacja `parseDateFlexible` jest bezpieczna

Sprawdzone: w całej aplikacji **nie ma ani jednego wywołania** `setHours/setDate/setMonth/setFullYear/…`,
więc nic nie mutuje zwróconego `Date`. Mimo to cache trzyma **znacznik czasu**, a nie instancję —
każde wywołanie dostaje świeży obiekt. Klucz cache brany jest przed normalizacją `v`
(niżej obcinane są godziny/sufiks `T`), żeby zapis i odczyt trafiały w to samo miejsce.

## Weryfikacja braku regresji

- `npm test` — 176 asercji, exit 0 (chromium + webkit + webkit/dotyk).
- Dedykowane porównanie „odcisku palca”: 3 pliki × 5 stanów (po wczytaniu, po filtrze, po sorcie,
  po przełączeniu na EN, po edycji komórki) × 29 pól — teksty wszystkich paneli analiz, nagłówki,
  pierwsze/ostatnie wiersze tabeli, a także **wewnętrzny stan agregacji**: wybrany wiersz nagłówka,
  opcje grupowania, miary i pełne profile kolumn (liczniki typów + unikaty).
  **Wynik: 0 różnic** względem kodu sprzed zmian.
- Stany „po EN” i „po edycji komórki” są w teście celowo — sprawdzają unieważnianie cache
  zależnego od języka i od zmiany danych.

## Co ZOSTAJE wolne (świadomie, nie zrobione)

1. **Pierwsze otwarcie panelu agregacji: ~1,3 s @4×.** Zostało 7× `buildRows` przy zimnym cache.
   Lek: pociąć autodetekcję na kawałki z oddawaniem wątku (odłożony pomysł „chunking”), żeby
   UI nie zamarzał — zysk w odczuciu, nie w sumie CPU. Wymaga decyzji, bo zmienia timing na async.
2. **Pełny re-render tabeli przy przełącznikach: 280–490 ms @2000 wierszy.** Profil pokazuje,
   że to **nie JavaScript** (JS ≈ 30 ms), tylko layout/style/paint przy przebudowie ~40 000 komórek.
   Jedyny realny lek to odłożony lewar **DOM-reuse / CSS-gating** — bez niego tu nie ma czego optymalizować.
3. **`XLSX.read` ~1,3 s @4× na pliku 610 KB** — synchroniczne, na main thread. To Faza 4 (Worker).
   Uwaga: worker (`buildRowsAsync`) istnieje, ale obsługuje tylko `buildRows`, nie samo parsowanie pliku.
4. **Scroll tabeli jest OK** — p50 27 ms / p95 52 ms @4× (czyli ~37 fps pod poczwórnym dławieniem).
   Wcześniejsze podejrzenie o „lag scrolla” na desktopie było artefaktem pomiaru w trybie headless.

---

*Utworzone: 2026-09-07.*
