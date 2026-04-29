# Module Overlap Audit

Audyt z 1 kwietnia 2026 po domknieciu pierwszej iteracji `Sheet / Workbench Inspector`.

## Wnioski Glowne

- `Sheet / Workbench Inspector` i `Formula Workbench`:
  - nie dubluja sie produktowo
  - dziela tylko fundamenty techniczne: analiza arkusza, statystyki, szybkie skoki po tabeli
  - inspektor odpowiada na pytanie `jak arkusz jest zbudowany`
  - formula workbench odpowiada na pytanie `jakie sa formuly i gdzie sa problemy`

- `KPI Extractor`:
  - czesciowo styka sie z `Analiza workbench`, ale nie jest duplikatem
  - obecna analiza pokazuje metadane i sygnaly struktury
  - KPI extractor ma wyciagac kluczowe liczby z workbookow typu kosztorys, dashboard, formularz

- `Cross-Sheet Dependency Explorer`:
  - nie dubluje obecnych funkcji
  - czesciowo opiera sie na danych workbookowych typu hidden sheets / names / pomocnicze arkusze
  - jego prawdziwa wartosc to relacje miedzy arkuszami, lookupami i potencjalnymi zrodlami danych

## Decyzja

Plan nie wymaga duzej przebudowy.

Najbardziej sensowna kolejnosc zostaje:
1. domknac lekko `Sheet / Workbench Inspector`
2. rozwijac `Formula Workbench`
3. potem `KPI Extractor`
4. na koncu `Cross-Sheet Dependency Explorer`

## Status Na Dzis

- `Sheet / Workbench Inspector`:
  - czesciowo wdrozony
  - ma juz wspolny panel i wspolne akcje

- `Formula Workbench`:
  - zaczete
  - pierwsza wersja ma liste formul, wyszukiwanie, filtry i skok do komorki

- `KPI Extractor`:
  - jeszcze nie ruszany

- `Cross-Sheet Dependency Explorer`:
  - jeszcze nie ruszany

## Dodatkowa Decyzja

`Compare / Diff` wypada na razie z planu wdrozen.

Powod:
- z poziomu produktu nie daje teraz az tak duzej wartosci jak pozostale moduly
- lepiej skupic energie na funkcjach bardziej codziennych i bardziej zgodnych z obecnym kierunkiem workbencha

Jesli temat compare wroci, to tylko jako lekka wersja:
- `Key Compare`:
  - compare danych po kluczu
  - raczej dla 2 arkuszy w jednym pliku albo 2 podobnych zestawow danych
- `Formula Compare`:
  - compare formul miedzy 2 arkuszami
  - najlepiej jako dodatki do `Formula Workbench`, a nie wielki osobny modul
