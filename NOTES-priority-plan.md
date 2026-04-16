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

## Dodatkowa decyzja produktowa - 13 kwietnia 2026

W obszarze lekkiej analityki:

1. najpierw `Duration Analyzer+`
2. potem lekki `kreator agregacji`

Co ma wejsc do `Duration Analyzer+`:
- srednia
- mediana
- min / max
- liczba rekordow
- filtr `wszystkie / tylko zamkniete / tylko otwarte`
- sortowanie po wybranej metryce

Co ma wejsc pozniej do lekkiego `kreatora agregacji`:
- `Grupuj po`
- `Mierz`
- `Agregacja`
- `Zakres`

Notatka po pierwszym wdrozeniu:
- obecna wersja Etapu 2 jest dobra jako `v1`
- nie traktujemy jej jako zamknietej, tylko jako bazy pod dalsza rozbudowe
- kluczowa decyzja techniczna: rozwijac to tak, by kolejne funkcje dalo sie dokladac przez ogolny model danych i agregacji, a nie przez osobne specjal-case'y

Kierunek dalszego rozwoju:
- wieksza mozliwosc recznej kustomizacji
- wieksza kontrola nad grupa / miara / sposobem liczenia
- rozwijanie tego w strone narzedzia, ktore zastapi czesc workflow opartych dzis w Excelu o makra
- szczegolnie wazne dla pracy na tablecie i w PWA, gdzie makra nie sa realnym rozwiazaniem

Wazna uwaga UX dla tego etapu:
- nie zakladamy, ze wszystko musi wyjsc z panelu bocznego
- jesli pewne rzeczy z Etapu 2 dobrze mieszcza sie w sidebarze, to to jest w porzadku
- na obecnym etapie limit wynikow ma sens, bo zmniejsza scrollowanie
- ale jesli przy pelnym wykorzystaniu funkcji cos przestaje sie tam miescic albo robi sie nieludzkie w obsludze, trzeba to pokazac inaczej
- docelowe umiejscowienie ma byc przyjemne, logiczne i wygodne, a nie wymuszone jedna sztywna regula layoutu

Wazna zasada:
- robimy to przez reuse obecnych mechanik workbencha
- nie budujemy od razu duzego osobnego modulu pivot / BI

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
