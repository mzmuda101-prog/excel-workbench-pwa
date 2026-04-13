# Excel Workbench PWA Roadmap

Ten dokument trzyma praktyczny plan rozwoju projektu jako:
- lokalny i bezpieczny workbench do plikow `.xlsx` / `.xlsm`
- narzedzie "Excel, ale lepiej" tam, gdzie oryginalny Excel jest niewygodny
- inspektor i warsztat do filtrowania, porzadkowania i rozumienia plikow, a nie tylko podglad tabeli

## Kierunek produktu

Priorytety projektu:
- bezpieczenstwo samych wczytywanych arkuszy
- lokalne przetwarzanie danych w przegladarce
- funkcje, ktore w Excelu sa ukryte, trudne albo meczace w uzyciu
- funkcje stricte workbenchowe: filtry, sortowanie, presety, analiza struktury i wygodniejsze widoki robocze

Zasada produktowa:
- nie probujemy zrobic pelnego "Excela w przegladarce"
- robimy narzedzie do pracy z plikami Excel wygodniejsze niz Excel w wybranych obszarach
- nie kopiujemy 1:1 desktopowych funkcji edycji, jesli w PWA dawałyby gorsza wiernosc zapisu albo mylne oczekiwania co do zachowania pliku po ponownym otwarciu w Excelu

## Desktop vs PWA

Wersja desktopowa Python ma przewage tam, gdzie liczy sie:
- wierniejsza edycja i zapis do pliku
- bezposrednia praca na workbooku
- bardziej zaawansowane podpowiedzi przy recznej edycji komorek

Dlatego do PWA przenosimy tylko te rzeczy, ktore maja tam realny sens.
Nie przenosimy na sile calego dzisiejszego pakietu sugestii i UX edycji, jesli w przegladarce dawalby slabszy efekt koncowy.

Najwiekszy sens ma rozwijanie w PWA tego, co:
- nie psuje pliku
- nie zalezy od bardzo wiernej edycji i zapisu
- daje realna przewage nad zwyklym Excelem przy przegladaniu, filtrowaniu i rozumieniu danych

## Co z dzisiejszych rzeczy ma jeszcze sens dla PWA

- lepsza informacja o skrotach i obsludze tam, gdzie rzeczywiscie pomaga
- wybrane lekkie inspiracje rankingiem podpowiedzi, ale tylko jesli nie komplikują zbyt mocno UX
- inspiracje porzadkowaniem sugestii, niekoniecznie pelny system z desktopu

## Czego nie traktowac jako priorytetu dla PWA

- rozbudowanego panelu podpowiedzi przy edycji komorek
- workflow sugerujacego desktopowy poziom wiernej edycji i zapisu
- funkcji opartych glownie o przewage bezposredniej edycji pliku, ktore w PWA tracilyby sens lub jakosc

## MVP

Najblizsze rzeczy do zrobienia:
- multi-sort po wielu kolumnach
- presety filtrow i sortowania
- mocniejsze filtry wielokolumnowe i datowe
- dalsze dopracowanie `Wide-to-Long View`
- mocniejszy `Sheet / Workbench Inspector`
- column profiler
- formula workbench

## Co ma w PWA najwiekszy sens

- filtry
- sortowanie
- presety
- `Wide-to-Long View`
- `Sheet / Workbench Inspector`
- `Formula Workbench`

Lekkie analizy sa mile jako dodatek, ale nie powinny dominowac kierunku produktu.

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
  - status: planowane, ale nizszy priorytet
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

## Decision Note - Duration Analyzer+ -> lekki kreator agregacji

Decyzja z 13 kwietnia 2026:
- najpierw rozwijamy obecny panel analizy czasu dla powtarzalnych blokow
- dopiero potem wyciagamy z niego bardziej ogolny, lekki `kreator agregacji`

Powod:
- to najnizszy koszt przy najwyzszej wartosci na realnych plikach typu obieg / cykle / etapy
- pozwala sprawdzic na prawdziwych workbookach, jakie agregacje i filtry sa naprawde potrzebne
- daje naturalna droge do reuse istniejacego kodu zamiast budowania osobnego modulu od zera

### Etap 1 - Duration Analyzer+

Najpierw rozwijamy obecny panel o:
- srednia
- mediana
- min / max
- liczba rekordow
- `wszystkie / tylko zamkniete / tylko otwarte`
- sortowanie po sredniej / medianie / liczbie rekordow

Zasada:
- to nadal ma byc lekki, bardzo ludzki panel pomocniczy
- bez zamiany sidebaru w pelny kreator pivotow

### Etap 2 - lekki kreator agregacji

Po ustabilizowaniu panelu czasu:
- `Grupuj po:` wybrana kolumna
- `Mierz:` zakres dat `od-do` albo wybrana kolumna liczbowa / czasowa
- `Agregacja:` srednia, mediana, min, max, liczebnosc
- `Zakres:` aktualny widok albo caly arkusz
- opcjonalnie: praca na wykrytych blokach albo na widoku `Wide-to-Long`

Zasada:
- nie budujemy pelnego `Pivot Builder`
- budujemy prosty, szybki i czytelny analizator do najczestszych pytan roboczych

Uwaga UX / IA:
- przy Etapie 2 nie zakladamy z gory, ze wszystko musi wyjsc z sidebaru
- jesli czesc funkcji wygodnie miesci sie w panelu bocznym, to moze tam zostac
- obecny panel pokazuje tylko wycinek wynikow, co ma sens na etapie lekkiego dodatku i ogranicza scrollowanie
- dopiero te elementy, ktore nie mieszcza sie dobrze przy wykorzystaniu 100% mozliwosci, powinny dostac inne miejsce w aplikacji
- decyzja o wyniesieniu czegos poza sidebar ma byc oparta o czytelnosc, wygode i logike pracy, nie o sztywna zasade
- docelowo powinno byc wygodne:
  - przegladanie dluzszej listy wynikow
  - sortowanie i filtrowanie wynikow agregacji
  - przechodzenie miedzy agregacja a tabela zrodlowa
  - praca bez wrazenia, ze wazne rzeczy sa schowane albo upchniete nieludzko
- mozliwe kierunki:
  - osobny tryb / widok roboczy
  - rozwijana sekcja w glownym obszarze
  - dedykowany ekran analityczny nad tabela albo obok niej

### Reuse istniejacego kodu

Przy tej sciezce maksymalnie wykorzystujemy:
- `Repeating Block Detector`
- `Wide-to-Long View`
- `Column Profiler`
- istniejacy panel analityczny czasu
- logike wykrywania `osoba / od / do / dlugosc`
- aktualne filtry, sortowanie i workflow pracy na przefiltrowanym widoku

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

Rzeczy z tej sekcji maja sens dopiero po dopracowaniu:
- filtrow
- sortowania
- presetow
- inspektora
- `Wide-to-Long`

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

### 2. Wide-to-Long + Sheet / Workbench Inspector

Dlaczego:
- to jest prawdziwa przewaga workbencha nad zwykla tabela
- pomaga rozumiec trudne arkusze bez ryzykownej edycji pliku

Zakres:
- dalsze dopracowanie `Wide-to-Long View`
- polaczenie `Nawigator sekcji`, `Wykrywacz blokow` i inspektora w bardziej spojny workflow
- bardziej spojna nawigacja po strukturze arkusza
- jedno miejsce na sygnaly: bloki, sekcje, problematyczne obszary i pomocnicze metryki

Etapy:
1. dopracowac wykrywanie blokow i sekcji
2. ograniczyc duplikowanie informacji
3. lepiej spiac `Wide-to-Long` z inspektorem
4. uporzadkowac akcje typu `Skocz`, `Ustaw naglowek`, `Skocz do kolumny`
5. zostawic droge do dalszych rozszerzen

Definition of done:
- uzytkownik szybciej rozumie trudny arkusz i moze pracowac na nim bardziej warsztatowo niz w zwyklym Excelu

### 3. Mocniejsze filtry + ergonomia codziennej pracy

Dlaczego:
- to jest glowny powod, dla ktorego taki workbench w PWA ma sens na co dzien
- daje wartosc od razu, bez wchodzenia w ciezsza analityke

Zakres:
- dopracowanie wielokolumnowych filtrow tekstowych
- dopracowanie filtrow dat
- szybkie odtwarzanie roboczych widokow
- lepsza ergonomia codziennej pracy na danych

Status na 1 kwietnia 2026:
- czesciowo wdrozone
- podstawy sa juz obecne
- do dalszego dopracowania: ergonomia filtrow, kolejnosc akcji, szybsze odtwarzanie widokow roboczych

Etapy:
1. poprawic ergonomie filtrow wielokolumnowych
2. dopracowac filtry dat
3. uporzadkowac szybkie odtwarzanie widokow
4. dopracowac presetowe workflow codziennej pracy

Definition of done:
- uzytkownik potrafi szybko zbudowac i odtwarzac robocze widoki danych lepiej niz w zwyklym Excelu

### 4. Column Profiler i lekkie analizy

Dlaczego:
- lekkie analizy sa przydatne, ale dopiero po dowiezieniu filtrowania, sortowania i widokow roboczych
- maja wspierac glowne workflow, a nie je zastepowac

Status na 1 kwietnia 2026:
- czesciowo zaczete
- profiler i lekkie sygnaly sa sensowne jako wsparcie
- nie sa obecnie najwazniejszym kierunkiem produktu PWA

Zakres:
- typ kolumny
- procent pustych
- liczba unikalnych
- najczestsze wartosci
- min / max dla liczb i dat
- lekkie flagi ostrzegawcze

Etapy:
1. policzyc podstawowe statystyki kolumn po zaladowaniu arkusza
2. pokazac top wartosci i flagi ostrzegawcze
3. spiac to lepiej z inspektorem
4. zostawic droge do dalszych lekkich analiz

Definition of done:
- uzytkownik dostaje szybki wglad pomocniczy, ale glowny nacisk produktu nadal idzie na filtrowanie i pracę z widokiem

### 5. Formula Workbench

Dlaczego:
- formuly to jeden z najbardziej upierdliwych obszarow Excela
- ale w PWA to nadal powinien byc dodatek do glownego workbenchowego workflow, nie glowny temat dnia codziennego

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
1. multi-sort + presety + filtry
2. `Wide-to-Long` + `Sheet / Workbench Inspector`
3. column profiler i lekkie analizy
4. formula workbench
5. KPI Extractor
6. Cross-Sheet Dependency Explorer
7. Key Compare / Formula Compare (lekka wersja, tylko jesli bedzie realna potrzeba)

Powod:
- pierwsze dwa obszary daja najwieksza codzienna wartosc w PWA
- lekkie analizy maja sens bardziej jako wsparcie niz jako glowny kierunek
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
# Dalsze przeniesienie dobrych rzeczy z desktopowego Python workbencha

- Dodac bardziej swiadomy UX/UI panelu podpowiedzi przy edycji komorek, na tyle na ile pozwala wersja PWA
- Dodac mozliwosc bardziej swiadomego pokazywania / ukrywania listy sugestii przy edycji
- Dodac opcje wlaczania / wylaczania wybranych zachowan sugestii, zamiast trzymac wszystko na stale
- Dodac lepsza informacje o skrotach i obsludze podpowiedzi dla uzytkownika
- Sprobowac przeniesc logike inteligentniejszych sugestii z desktopu do PWA tam, gdzie to realnie mozliwe bez backendu i bez pogorszenia UX
