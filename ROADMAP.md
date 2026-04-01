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
- polaczenie podobnych funkcji inspektorskich w 1-2 bardziej spojne moduly
- formula workbench

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
  - docelowo powinien rozumiec, ze podsumowania moga byc nad tabela, pod tabela albo przy mniejszych blokach danych
  - przy duplikatach powinien preferowac bardziej ludzka etykiete, a pozostale pokazywac jako aliasy
- `Cross-Sheet Dependency Explorer`:
  - status: planowane
  - inspektor zaleznosci miedzy arkuszami pomocniczymi, lookupami i pivotami
- `Key Compare`:
  - status: pozniej
  - lekki compare danych po kluczu dla 2 arkuszy w jednym pliku albo 2 podobnych zestawow danych
- `Formula Compare`:
  - status: pozniej
  - lekkie porownanie formul miedzy 2 arkuszami, najlepiej jako rozszerzenie `Formula Workbench`

## Pozniej

Kolejna warstwa wartosci:
- workbook inspector rozszerzony o named ranges, freeze panes, tabele, hidden / very hidden
- sheet map / heatmapa struktury arkusza
- filtrowanie po typie komorki, regexie, dlugosci tekstu i kolorze
- wykrywanie wielu blokow danych w jednym arkuszu
- batch tools do bezpiecznego czyszczenia danych z preview
- eksport raportow diagnostycznych do CSV / JSON / TXT
- lekki `Key Compare` dla 2 arkuszy w jednym pliku
- lekki `Formula Compare` jako rozszerzenie `Formula Workbench`

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

### 3. Sheet / Workbench Inspector

Dlaczego:
- sidebar zaczyna miec kilka paneli, ktore odpowiadaja na bardzo podobne pytania
- lepiej najpierw uproscic i scalic te funkcje niz dokladac kolejny osobny modul

Zakres:
- polaczenie `Profiler kolumn`, `Nawigator sekcji` i `Wykrywacz blokow`
- wspolny widok typu `Inspektor arkusza` albo `Workbench Inspector`
- bardziej spojna nawigacja po strukturze arkusza
- jedno miejsce na sygnaly: typy kolumn, bloki, sekcje, problematyczne obszary

Status na 1 kwietnia 2026:
- czesciowo wdrozone
- istnieje juz wspolny panel `Inspektor arkusza`
- obecna wersja nadal korzysta z trzech osobnych rendererow pod spodem, ale scala je w jeden bardziej spojny UX
- do dalszego dopracowania: odchudzenie duplikatow, lepsze priorytety sekcji i mocniejsze wspolne akcje

Etapy:
1. zaprojektowac wspolny model danych dla tych trzech paneli
2. ograniczyc duplikowanie informacji
3. zrobic 1 glowny panel + ewentualnie 1 pomocniczy
4. uporzadkowac akcje typu `Skocz`, `Ustaw naglowek`, `Skocz do kolumny`
5. zostawic droge do dalszych rozszerzen

Definition of done:
- uzytkownik szybciej rozumie arkusz, a sidebar jest krotszy i bardziej spojny

### 4. Formula Workbench

Dlaczego:
- formuly to jeden z najbardziej upierdliwych obszarow Excela
- tu mozna zrobic realnie lepszy workflow niz w samym Excelu

Status na 1 kwietnia 2026:
- zaczete
- pierwsza wersja ma osobny panel z lista formul, filtrem, wyszukiwaniem i skokiem do komorki
- do dalszego dopracowania: lepsze wykrywanie wzorcow formul, filtry po funkcjach i mocniejsza analiza bledow

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

## Kolejnosc Realizacji

Proponowana kolejnosc pracy:
1. column profiler
2. multi-sort + presety
3. sheet / workbench inspector
4. formula workbench
5. KPI Extractor
6. Cross-Sheet Dependency Explorer
7. Key Compare / Formula Compare (lekka wersja, tylko jesli bedzie realna potrzeba)

Powod:
- pierwsze trzy funkcje najszybciej podniosa codzienna uzytecznosc
- trzeci krok dodatkowo uporzadkuje sam interfejs
- formula workbench dalej daje najmocniejszy efekt "tego Excel sam dobrze nie daje"

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
- Audyt duplikacji modulow 2-5 jest w [NOTES-module-overlap-audit.md](./NOTES-module-overlap-audit.md).
