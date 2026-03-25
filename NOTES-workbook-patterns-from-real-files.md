# Notes - Workbook Patterns From Real Files

Ten dokument zbiera wnioski z lokalnej analizy dwoch realnych workbookow:
- maly kosztorys / budzet z tabela i sekcja KPI
- duzy arkusz obiegu z powtarzalnymi cyklami, arkuszem pomocniczym i tabela przestawna

Cel:
- wyciagnac wzorce, ktore `excel-workbench-pwa` powinien umiec rozpoznawac
- zapisac pomysly na funkcje produktowe pod takie typy plikow
- uwzglednic nie tylko dane i formuly, ale tez style, mergy i uklad workbooka

## Wzorzec 1 - Kosztorys / budzet / prosty arkusz decyzyjny

### Co widac w strukturze

- jeden arkusz
- jedna tabela Excela `Dane`
- tabela w zakresie `B5:G23`
- sekcja podsumowania nad tabela
- scalone komorki dla KPI i etykiet
- formuly podsumowujace budzet
- formaty walutowe
- czesc rekordow jest niepelna lub "w trakcie uzupelniania"

### Co widac w stylach

- tytul arkusza jako osobny, szeroki blok
- wyraznie oddzielone KPI od tabeli danych
- naglowek tabeli ma inny styl niz dane
- sa osobne style dla kwot, ilosci i pol tekstowych
- arkusz jest projektowany jak mini-formularz / mini-dashboard, nie jak surowa baza

### Co to znaczy dla workbencha

To nie jest tylko tabela. To jest:
- tabela + sekcja decyzji
- tabela + status budzetowy
- tabela + pola nieuzupelnione / orientacyjne

### Pomysly na funkcje

#### 1. KPI Extractor

Rozpoznawanie arkuszy, gdzie nad tabela sa karty podsumowujace:
- budzet docelowy
- koszt calkowity
- roznica / zapas / przekroczenie

Workbench moglby pokazac je jako osobne karty, zamiast zmuszac do czytania komorek po layoutcie.

#### 2. Budget Mode

Specjalny tryb dla kosztorysow:
- suma pozycji
- pozycje bez kosztu
- pozycje bez ilosci
- pozycje z kwota wpisana recznie zamiast koszt x ilosc
- ile zostalo do budzetu
- ktore pozycje najbardziej "zjadaja" budzet

#### 3. Missing Cost Audit

Wykrywanie rekordow typu:
- jest nazwa, brak kosztu
- jest ilosc, brak kosztu
- jest uwaga, ale brak kwoty
- jest kwota, ale brak jasnej logiki wyliczenia

To jest bardzo przydatne dla kosztorysow roboczych.

#### 4. What-if Preview

Lekki tryb scenariuszy:
- co sie stanie, jesli zmienimy koszt lub ilosc
- podglad nowego budzetu bez psucia oryginalnego pliku
- porownanie "obecnie" vs "wariant A"

#### 5. Formula Intent View

Pokazanie formul w jezyku bardziej ludzkim:
- suma tabeli
- roznica budzetu
- przekroczono / nie przekroczono

To byloby bardzo fajne dla prostych arkuszy decyzyjnych.

## Wzorzec 2 - Powtarzalne cykle / obieg / harmonogram osob lub spraw

### Co widac w strukturze

- glowny arkusz z duza tabela Excela `opracowanie_terenow`
- arkusz pomocniczy `kolumny_pomocnicze`
- tabela przestawna
- maly arkusz lookup / testowy
- dodatkowy arkusz ukladany bardziej wizualnie niz tabelarycznie

### Najwazniejszy wzorzec danych

Glowny arkusz nie jest klasyczna plaska tabela.

Zawiera:
- numer rekordu
- 8 powtarzalnych blokow cyklu
- dla kazdego cyklu pola typu:
  - osoba
  - od
  - do
  - dlugosc

Czyli logicznie to jest:
- "jeden rekord"
- z osmioma powtorzeniami tej samej mini-struktury

To jest bardzo wazne, bo wiele takich plikow w Excelu jest budowanych "wszerz", a nie "w dol".

### Co widac w stylach

- scalone naglowki blokow `1 Cykl`, `2 Cykl`, ..., `8 Cykl`
- mocne centrowanie i siatkowy, raportowy uklad
- formaty dat i dlugosci trwania
- pomocniczy arkusz liczy dlugosci przez formuly
- sa zielone i czerwone fonty, czyli workbook komunikuje stany / statusy
- jest osobny arkusz wizualny z wieloma mergami i conditional formatting

To znaczy, ze workbook jest troche:
- baza danych
- troche kalkulator
- troche raport
- troche interfejs roboczy

### Co to znaczy dla workbencha

To jest typ workbooka, z ktorym zwykly grid radzi sobie srednio.

Najwiekszy problem:
- dane logicznie sa rekordami i cyklami
- w Excelu sa ulozone jako szeroki, wieloblokowy raport

### Pomysly na funkcje

#### 1. Repeating Block Detector

Rozpoznawanie arkuszy, gdzie wystepuja powtarzalne grupy kolumn:
- `Imie i Nazwisko`
- `od`
- `do`
- `Dlugosc`

z kolejnymi numerami:
- `Imie i Nazwisko2`
- `od2`
- `do2`
- `Dlugosc2`

Workbench moglby automatycznie powiedziec:
- "to jest jeden model danych z 8 powtarzalnymi cyklami"

#### 2. Wide-To-Long View

Jedno z najmocniejszych usprawnien:
- z szerokiego arkusza robic logiczny widok dlugi
- jeden rekord -> wiele wierszy cykli

Zamiast:
- 1 rekord z 8 blokami wszerz

uzytkownik moglby zobaczyc:
- rekord 14, cykl 1
- rekord 14, cykl 2
- rekord 14, cykl 3

To byloby ogromnie wygodniejsze do filtrowania, sortowania i analizy.

#### 3. Duration Analyzer

Poniewaz workbook liczy dlugosci:
- workbench moze pokazac rozklady czasow
- rekordy bez daty koncowej
- rekordy bez daty startu
- sredni czas na cykl
- rekordy odstajace
- porownanie dlugosci miedzy cyklami

#### 4. Cross-Sheet Dependency Explorer

Tu bardzo widac zaleznosci:
- glowny arkusz
- arkusz pomocniczy liczacy dlugosci
- tabela przestawna
- arkusz lookup

Workbench moglby pokazac:
- z czego wynikaja dane kolumny
- ktore arkusze sa pomocnicze
- ktore sa raportowe
- ktore sa tylko lookupiem

#### 5. Formula Network View

Przy takich plikach bardzo przydatny bylby widok:
- te kolumny sa wyliczane
- te kolumny sa wejsciowe
- te arkusze sa pochodne

To byloby lepsze niz reczne sledzenie formul po Excelu.

#### 6. Table + Pivot Awareness

Workbench powinien rozpoznawac zestaw:
- glowna tabela
- arkusz pomocniczy
- pivot / raport

i umiec pokazac to jako jeden "pakiet workbooka", a nie 3 oddzielne byty bez kontekstu.

#### 7. Privacy-Friendly Review Mode

Poniewaz plik ma charakter RODO:
- maskowanie nazwisk / danych osobowych w widoku
- analiza struktur, dat i dlugosci bez pokazywania danych wrazliwych
- szybkie przelaczenie `pelne dane / dane zmaskowane`

To bardzo pasuje do lokalnego workbencha.

#### 8. Section Navigator

Na podstawie mergow i stylow workbench moze wykrywac sekcje:
- grupy cykli
- podsumowania
- bloki wizualne
- arkusze typu dashboard

Czyli nie tylko "kolumny A-AG", ale tez:
- "blok cyklu 1"
- "blok cyklu 2"
- "widok raportowy"
- "arkusz pomocniczy"

#### 9. Style Map / Layout Map

W takich plikach style niosa znaczenie.

Workbench moglby umiec pokazac:
- gdzie sa pola wejscia
- gdzie sa pola wynikowe
- gdzie sa sekcje naglowkowe
- gdzie sa arkusze bardziej dashboardowe niz tabelaryczne

#### 10. Conditional Formatting Inspector

Jesli arkusz ma conditional formatting albo statusy kolorystyczne:
- pokaz reguly
- pokaz komorki objete regula
- pokaz znaczenie kolorow i wyroznien

To jest rzecz bardzo trudna do szybkiego ogarniecia w zwyklym Excelu.

## Wspolne wnioski z obu workbookow

Te dwa pliki pokazaly, ze workbench powinien umiec obslugiwac co najmniej 3 typy arkuszy:

1. `table-first`
- klasyczna tabela z prostym podsumowaniem

2. `form-first`
- arkusz wygladajacy jak formularz, mini-dashboard albo kalkulator

3. `wide-cycle-first`
- szeroki arkusz z powtarzalnymi blokami danych

## Najmocniejsze pomysly produktowe wynikajace z tej analizy

1. Repeating block detector
2. Wide-to-long view
3. KPI extractor
4. Cross-sheet dependency explorer
5. Privacy-friendly review mode
6. Section navigator oparty o style i merge
7. Conditional formatting inspector
8. Missing cost / missing data audit
9. Duration analyzer
10. Table + pivot awareness

## Co warto dopisac do roadmapy jako przyszly kierunek

Mozliwy nowy obszar roadmapy:
- `Workbook Pattern Recognition`

Czyli automatyczne rozpoznawanie, z jakim typem workbooka mamy do czynienia:
- kosztorys
- formularz
- harmonogram
- szeroki arkusz cykliczny
- workbook z arkuszem pomocniczym i raportowym

To mogloby potem dynamicznie wlaczac najlepszy tryb pracy dla konkretnego pliku.
