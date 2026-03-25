# Priority Plan

Krotka notatka robocza z ustalona kolejnoscia funkcji do wdrazania.

## Pattern-aware Features

| Funkcja | Status | Notatka |
|---|---|---|
| Section Navigator | wdrozone | opcjonalna warstwa pomocnicza |
| Repeating Block Detector | beta | zalezne od poprawnego wiersza naglowka |
| Wide-to-Long View | beta | na razie widok analityczny bez edycji |
| KPI Extractor | pozniej | po domknieciu podstaw analitycznych |
| Cross-Sheet Dependency Explorer | pozniej | bardziej specjalistyczny modul inspektorski |

## Najwazniejsze Kolejne Rzeczy

Na podstawie ustalen:

1. `column profiler`
2. `multi-sort + presety`
3. `data quality scanner`
4. `formula workbench`
5. `compare / diff`
6. `KPI Extractor`
7. `Cross-Sheet Dependency Explorer`

## Dlaczego Taka Kolejnosc

- `column profiler`:
  - daje szybki wglad w prawie kazdy plik
  - buduje fundament pod kolejne moduly analityczne
- `multi-sort + presety`:
  - najmocniej poprawia codzienny workflow
  - rozwija obecny kierunek "Excel, ale lepiej"
- `data quality scanner`:
  - naturalnie korzysta z czesci sygnalow z profilera
  - daje szybkie przejscie od wykrycia problemu do filtrowania danych
- `formula workbench` i `compare / diff`:
  - bardzo wartosciowe, ale troche ciezsze i bardziej specjalistyczne
- `KPI Extractor` i `Cross-Sheet Dependency Explorer`:
  - nadal wartosciowe, ale mniej uniwersalne niz trzy pierwsze pozycje

## Notatka UX / IA

W przyszlosci warto polaczyc:
- `Profiler kolumn`
- `Nawigator sekcji`
- `Wykrywacz blokow`

w jeden bardziej spojny modul typu:
- `Inspektor arkusza`
- albo `Workbench Inspector`

Powod:
- te panele coraz czesciej odpowiadaja na podobne pytanie: "jak ten arkusz jest zbudowany i gdzie sa najwazniejsze rzeczy"
- obecnie sa praktyczne osobno, ale docelowo dobrze byloby uproscic sidebar i zredukowac liczbe podobnych paneli
