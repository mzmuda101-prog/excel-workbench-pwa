# Notes - Filter Press Workbook Ideas

Ten dokument zbiera pomysly i wnioski dla przyszlych funkcji `excel-workbench-pwa`
pod pliki typu:
- harmonogram pras filtracyjnych
- planowanie cykli pracy maszyn
- analiza konfliktow procesowych
- analiza urzadzen KKS i sekwencji procesu

Zalozenie:
- plik `.xlsx` moze byc wrazliwy i firmowy
- pomysly pochodza z PDF opisujacego proces, bez analizy samego arkusza Excel

## Kontekst procesowy z PDF

Na podstawie dokumentu `Opis działania pras filtracyjnych.pdf`:
- istnieja 3 prasy filtracyjne
- zazwyczaj pracuja 2 prasy, trzecia jest redundantna
- prasy maja pracowac 24/7
- kazdy przebieg ma powtarzalna strukture
- sa 4 glowne cykle:
  - praca wlasciwa
  - mycie plyt
  - CIP
  - przestoj
- praca wlasciwa sklada sie z 6 etapow w stalej kolejnosci:
  1. zamkniecie prasy
  2. podawanie i zageszczanie nadawy
  3. wyciskanie membranami
  4. wydmuch rdzenia
  5. spuszczanie wody z membran
  6. rozladunek

## Twarde reguly procesu, ktore aplikacja moglaby sprawdzac

- start CIP i start pracy wlasciwej nie moga zachodzic jednoczesnie
- mycie plyt nie powinno nakladac sie pomiedzy prasami
- etapy 1-6 musza isc zawsze we wlasciwej kolejnosci
- przenosniki tasmowe musza byc uruchomione 10 minut przed obciazeniem
- niektore urzadzenia dzialaja redundantnie, a nie rownolegle
- przebiegi dla prasy 1 i 2 maja ustalona, powtarzalna strukture

## Potencjalne funkcje dedykowane pod taki plik

### 1. Rule Checker

Automatyczne sprawdzanie logiki procesu:
- niedozwolone nakladki czasowe
- zla kolejnosc etapow
- brakujace etapy
- naruszenie ograniczen wspolnych mediow
- brak wymaganych wyprzedzen czasowych

### 2. Timeline Per Prasa

Widok osi czasu dla:
- FP03101
- FP03102
- FP03103

Na osi czasu:
- praca wlasciwa
- mycie plyt
- CIP
- przestoj
- opcjonalnie etapy 1-6 jako podpoziom

### 3. Conflict Scanner

Widok pokazujacy tylko problemy:
- kolizje CIP vs nadawa
- nakladki mycia
- konflikt uzycia urzadzen redundantnych
- brak uruchomienia tasm 10 minut przed obciazeniem
- niepelne lub niespojne cykle

### 4. Cycle Validator

Walidacja wzorcow przebiegow.

Prasa 1:
- 13x praca
- mycie
- 17x praca
- przestoj
- CIP

Prasa 2:
- przestoj
- 16x praca
- mycie
- 14x praca
- CIP

Aplikacja moglaby:
- sprawdzac zgodnosc z wzorcem
- pokazywac miejsce odchylenia
- oznaczac brakujace lub dodatkowe cykle

### 5. KKS Device Inspector

Inspektor urzadzen elektrycznych i napedow:
- wyszukaj po KKS, np. `P03121`
- pokaz kiedy urzadzenie pracuje
- dla ktorej prasy
- w jakim cyklu / etapie
- czy uzycie jest poprawne wzgledem zasad procesu

## Najbardziej wartosciowe funkcje pod taki plik

1. timeline pras
2. rule checker
3. cycle validator
4. KKS device inspector
5. conflict scanner

## Co Excel robi tu slabo, a workbench moze robic lepiej

- pilnowanie logiki procesu
- wykrywanie konfliktow czasowych
- pokazanie przebiegu kilku pras na jednej osi czasu
- sprawdzanie kolejnosci etapow
- wykrywanie brakow i odchylen od wzorca
- analiza pracy konkretnych urzadzen KKS

## Potencjalny tryb produktowy

Mozliwy osobny tryb:
- `Process-aware workbook mode`

Czyli:
- aplikacja rozpoznaje kolumny typu `prasa`, `cykl`, `etap`, `start`, `stop`, `urzadzenie`
- naklada reguly procesu
- pokazuje nie tylko dane, ale tez naruszenia logiki

## Przydatne przyszle pytania przy wdrozeniu

Jesli w przyszlosci bedziemy to budowac, warto ustalic:
- jakie sa realne kolumny w pliku `.xlsx`
- czy sa osobne wiersze dla etapow czy dla calych cykli
- czy czas jest jako `start/stop`, czy jako `start + duration`
- czy urzadzenia sa przypisane jawnie w tabeli
- czy sa osobne arkusze na prasy / cykle / urzadzenia
- czy harmonogram jest planem, wykonaniem czy jednym i drugim
