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
| Key Compare | pozniej | tylko lekka wersja compare danych po kluczu |
| Formula Compare | pozniej | tylko jako lekkie rozszerzenie Formula Workbench |

## Najwazniejsze Kolejne Rzeczy

Na podstawie ustalen:

1. `column profiler`
2. `multi-sort + presety`
3. `sheet / workbench inspector`
4. `formula workbench`
5. `KPI Extractor`
6. `Cross-Sheet Dependency Explorer`
7. `Key Compare / Formula Compare` w lekkiej wersji, tylko jesli pojawi sie realna potrzeba

## Dlaczego Taka Kolejnosc

- `column profiler`:
  - daje szybki wglad w prawie kazdy plik
  - buduje fundament pod kolejne moduly analityczne
- `multi-sort + presety`:
  - najmocniej poprawia codzienny workflow
  - rozwija obecny kierunek "Excel, ale lepiej"
- `sheet / workbench inspector`:
  - porzadkuje sidebar zanim dorzucimy kolejne duze moduly
  - scala podobne funkcje w bardziej spojny UX
- `formula workbench`:
  - bardzo wartosciowe i dobrze pasuje do obecnego kierunku projektu
- `KPI Extractor` i `Cross-Sheet Dependency Explorer`:
  - nadal wartosciowe, ale mniej uniwersalne niz trzy pierwsze pozycje
  - przy KPI warto myslec nie tylko o "gorze arkusza", ale o podsumowaniach wokol blokow danych / tabelek

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

Decyzja z 1 kwietnia 2026:
- `data quality scanner` wypada z najblizszego planu
- jego miejsce zajmuje priorytet na inteligentne polaczenie podobnych funkcji inspektorskich
- pierwszy krok tego laczenia jest juz wdrozony:
  - `Profiler kolumn`, `Nawigator sekcji` i `Wykrywacz blokow` sa teraz zebrane w panelu `Inspektor arkusza`
  - pod spodem nadal dzialaja jako osobne logiki, co zmniejsza ryzyko regresji
- po audycie modulow 2-5:
  - `Formula Workbench` nie jest duplikatem obecnych paneli i zostaje rozwijany dalej
  - `KPI Extractor` i `Cross-Sheet Dependency Explorer` zostaja w planie bez duzej przebudowy
- `Compare / Diff` wypada z planu na teraz:
  - z poziomu produktu nie daje na tym etapie az tak duzej wartosci jak inne moduly
  - mozna do niego wrocic pozniej, jesli pojawi sie realna potrzeba
- jesli temat compare wroci, to tylko jako:
  - `Key Compare`
  - `Formula Compare`
  - bez budowania duzego, osobnego modulu `Compare / Diff`
