# Excel Workbench PWA Roadmap

Ten dokument trzyma praktyczny plan rozwoju projektu jako:
- lokalny i bezpieczny workbench do plikow `.xlsx` / `.xlsm`
- narzedzie "Excel, ale lepiej" tam, gdzie oryginalny Excel jest niewygodny
- inspektor i analityk plikow, a nie tylko podglad tabeli

## Kierunek produktu

Priorytety projektu:
- bezpieczenstwo samych wczytywanych arkuszy
- lokalne przetwarzanie danych w przegladarce
- funkcje, ktore w Excelu sa ukryte, trudne albo meczace w uzyciu
- funkcje stricte workbenchowe: diagnostyka, porownania, profilowanie, analiza struktury

Zasada produktowa:
- nie probujemy zrobic pelnego "Excela w przegladarce"
- robimy narzedzie do pracy z plikami Excel wygodniejsze niz Excel w wybranych obszarach

## MVP

Najblizsze rzeczy do zrobienia:
- multi-sort po wielu kolumnach
- presety filtrow i sortowania
- column profiler
- data quality scanner
- formula workbench
- compare / diff dwoch arkuszy lub plikow

## Pattern-aware Features

Opcjonalne warstwy pomocnicze ponad klasycznym widokiem tabeli:
- `Nawigator sekcji`:
  - status: wdrozone
  - wykrywa potencjalny wiersz naglowka, bloki layoutu i podsekcje
  - ma pomagac w nawigacji, a nie zastepowac podstawowy widok
- `Wykrywacz blokow`:
  - status: wdrozone, ale nadal beta
  - wykrywa powtarzalne grupy kolumn, np. `1 Cykl ... 8 Cykl`
  - ma pomagac w szerokich arkuszach z cyklami, etapami i seriami kolumn
  - obecnie mocno zalezy od poprawnie ustawionego wiersza naglowka
  - do dopracowania: lepsze auto-wykrywanie i czytelniejszy opis blokow dla roznych typow workbookow
- `Wide-to-Long View`:
  - status: wdrozone, ale nadal beta
  - opcjonalny widok roboczy dla arkuszy z powtarzalnymi blokami
  - obecnie to bezpieczny widok analityczny, bez edycji komorek
  - do dopracowania: lepszy opis rekordow long i szersze wsparcie dla roznych wzorcow blokow
- `KPI Extractor`:
  - status: planowane
  - opcjonalny panel podsumowan dla workbookow typu kosztorys / dashboard
- `Cross-Sheet Dependency Explorer`:
  - status: planowane
  - inspektor zaleznosci miedzy arkuszami pomocniczymi, lookupami i pivotami

## Pozniej

Kolejna warstwa wartosci:
- workbook inspector rozszerzony o named ranges, freeze panes, tabele, hidden / very hidden
- sheet map / heatmapa struktury arkusza
- filtrowanie po typie komorki, regexie, dlugosci tekstu i kolorze
- wykrywanie wielu blokow danych w jednym arkuszu
- batch tools do bezpiecznego czyszczenia danych z preview
- eksport raportow diagnostycznych do CSV / JSON / TXT

## Premium-fajne

Rzeczy bardzo mocne produktowo:
- wykrywanie odstajacych formul wzgledem wzorca w kolumnie
- porownania miesiac do miesiaca / snapshot to snapshot
- inteligentne wykrywanie prawie-duplikatow i niespojnosci nazw
- tryby robocze per use case, np. analiza finansowa / kontrola jakosci / import do systemu
- zapisywanie lokalnych sesji analitycznych i presetow per typ pliku

## Najlepsze 5 Na Start

### 1. Multi-sort + presety filtrow

Dlaczego:
- to jest naturalne rozwiniecie obecnego najmocniejszego obszaru projektu
- duzo bardziej praktyczne niz standardowy Excelowy flow

Zakres:
- sortowanie po wielu kolumnach z kolejnoscia priorytetow
- zapis presetow filtrow i sortowania
- szybkie przelaczanie miedzy presetami

Etapy:
1. doda lista sortowan: kolumna + kierunek
2. dodac UI do zmiany kolejnosci sortowan
3. zapis presetow do localStorage
4. lista presetow w sidebarze
5. eksport / import presetow JSON

Definition of done:
- uzytkownik potrafi w 2-3 klikach odtworzyc zlozony widok filtrowania i sortowania

### 2. Column Profiler

Dlaczego:
- daje natychmiastowy wglad w dane bez recznego scrollowania
- to jest typowa funkcja workbenchowa, nie tylko UI do tabeli

Zakres:
- typ kolumny
- procent pustych
- liczba unikalnych
- najczestsze wartosci
- min / max dla liczb i dat
- wykrywanie mixed type

Etapy:
1. policzyc podstawowe statystyki kolumn po zaladowaniu arkusza
2. dodac panel profilu kolumn
3. pokazac top wartosci i flagi ostrzegawcze
4. dodac sortowanie kolumn problematycznych
5. umozliwic klikniecie w profil i od razu odfiltrowanie danych

Definition of done:
- po zaladowaniu arkusza widac, ktore kolumny sa problematyczne i jakie maja cechy

### 3. Data Quality Scanner

Dlaczego:
- bardzo przydatne przy prawdziwej pracy z plikami
- w Excelu zrobienie tego dobrze jest zwykle wolne i nieprzyjemne

Zakres:
- duplikaty calych wierszy
- duplikaty po wskazanych kolumnach
- puste wartosci w waznych kolumnach
- liczby jako tekst
- daty jako tekst
- trailing spaces / podwojne spacje / niespojne warianty

Etapy:
1. modul flag jakosci danych
2. panel "Problemy" z licznikami
3. klik w problem => automatyczny filtr widoku
4. konfiguracja "waznych kolumn"
5. eksport raportu problemow

Definition of done:
- uzytkownik moze jednym kliknieciem przejsc od wykrycia problemu do listy problematycznych rekordow

### 4. Formula Workbench

Dlaczego:
- formuly to jeden z najbardziej upierdliwych obszarow Excela
- tu mozna zrobic realnie lepszy workflow niz w samym Excelu

Zakres:
- lista wszystkich formul
- wyszukiwarka po funkcjach
- filtry: brak wyniku, blad, odstajacy wzorzec, wybrany arkusz / kolumna
- podglad adresu komorki i samej formuly

Etapy:
1. indeks formul przy ladowaniu arkusza
2. panel listy formul
3. wyszukiwarka po tekstach typu `XLOOKUP`, `SUMIFS`, `INDIRECT`
4. flaga formul bez wyniku / cache
5. wykrywanie formul odstajacych od wzorca w kolumnie

Definition of done:
- uzytkownik moze szybko znalezc i przeanalizowac problematyczne formuly bez klikania po komorkach

### 5. Compare / Diff

Dlaczego:
- to jedna z najbardziej wartosciowych funkcji workbenchowych
- bardzo potrzebne przy pracy miesiac do miesiaca albo wersja do wersji

Zakres:
- porownanie 2 arkuszy z jednego pliku
- porownanie 2 plikow
- roznice w naglowkach
- roznice w danych po kluczu
- roznice w formulach

Etapy:
1. porownanie naglowkow i struktury
2. porownanie rekordow po wybranym kluczu
3. widok tylko zmian
4. raport dodane / usuniete / zmienione
5. porownanie formul i typow danych

Definition of done:
- uzytkownik potrafi znalezc roznice miedzy dwoma wersjami arkusza bez recznego przeklikiwania

## Kolejnosc Realizacji

Proponowana kolejnosc pracy:
1. column profiler
2. multi-sort + presety
3. data quality scanner
4. formula workbench
5. compare / diff

Powod:
- pierwsze trzy funkcje najszybciej podniosa codzienna uzytecznosc
- formula workbench i diff dadza najmocniejszy efekt "tego Excel sam dobrze nie daje"

## Najblizszy Etap

Etap 1:
- wdrozyc multi-sort
- zapis presetow filtrow / sortowan
- przygotowac strukture stanu pod dalsze moduly analityczne

Cel etapu 1:
- zrobic z obecnego filtrowania prawdziwy killer feature projektu

## Dodatkowe Notatki Domenowe

- Pomysly pod harmonogram pras filtracyjnych zapisano w [NOTES-filter-press-workbench.md](./NOTES-filter-press-workbench.md).
- Wnioski z analizy realnych workbookow zapisano w [NOTES-workbook-patterns-from-real-files.md](./NOTES-workbook-patterns-from-real-files.md).
- Biezaca tabela priorytetow rozwoju jest w [NOTES-priority-plan.md](./NOTES-priority-plan.md).
