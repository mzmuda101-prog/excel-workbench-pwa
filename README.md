# Excel Workbench PWA

Lekka PWA do lokalnego przegladania, filtrowania i edycji arkuszy Excel.
Dziala w przegladarce (iPad i macOS), bez backendu. Pliki sa wybierane przez systemowy picker.

## Dla rekrutera (krotko)

Projekt pokazuje:
- PWA offline-first (Service Worker + manifest)
- Przetwarzanie .xlsx lokalnie w przegladarce (bez backendu)
- UX narzedziowy: filtry wielokolumnowe, zakresy dat, sortowanie, edycja
- Organizacja kodu pod dalszy rozwoj (czytelne warstwy UI/logic)

## Start lokalny

Uzyj dowolnego prostego serwera statycznego, np.:

```bash
python3 -m http.server 8001
```

Potem wejdz na:
```
http://127.0.0.1:8001/
```

## Deploy na Vercel

To jest statyczna strona. Na Vercel ustaw:
- Framework: `Other`
- Build: brak
- Output: root repo

## Dodanie na ekran glowny (iPad / iOS)

1. Otworz strone w Safari.
2. Kliknij ikone udostepniania.
3. Wybierz "Dodaj do ekranu poczatkowego".

## Bezpieczenstwo danych

Ta aplikacja nie wysyla plikow .xlsx na serwer — wszystko dzieje sie lokalnie w przegladarce.
Bezpieczenstwo zalezy od uruchamiania zaufanej wersji strony (bez dodatkowego kodu wysylajacego dane).

## Offline

Aplikacja ma prosty Service Worker. Po pierwszym uruchomieniu moze dzialac offline.
Uzywa **xlsx-js-style** (SheetJS + style); domyslnie z CDN. Dla pelnego offline skopiuj
`node_modules/xlsx-js-style/dist/xlsx.bundle.js` do `vendor/` i w `index.html` ustaw `<script src="vendor/xlsx.bundle.js">`.

## Funkcje

- Wczytanie .xlsx/.xlsm przez picker
- Wybor arkusza i wiersza naglowka
- Filtry tekstowe (2 niezalezne, wiele kolumn)
- Filtr dat (pomiedzy / przed / po / ostatnie N dni)
- Sortowanie
- Auto-szerokosc kolumn + reczne dopasowanie
- Edycja komorek (blokada formul)
- Zapis i Zapis jako...
- Eksport CSV
- Wykrywanie koloru komorek (fill) i subtelne podswietlenie w podgladzie

## Zapis a wersja Python (openpyxl)

W PWA zapis dziala tak: edytujesz komorki w pamieci (obiekt `workbook`), potem `XLSX.writeFile(workbook, plik)` zapisuje caly skoroszyt. Formuly i niezmienione komorki sa w obiekcie, wiec trafiaja do pliku. **Roznica:** W Pythonie openpyxl daje pelna wiernosc pliku (otwierasz → edytujesz obiekt → save() = ten sam plik + zmiany). W JS biblioteka przy odczycie i zapisie moze czegos nie odtworzyc 1:1 (np. skomplikowane formatowanie, nisze formuly). Da sie zblizyc do Pythona: uzywamy xlsx-js-style (zachowanie stylow), edytujemy tylko wartosci komorek (formul nie ruszamy) — round-trip jest wtedy lepszy. Pelna rownowaznosc z openpyxl w samej przegladarce nie jest mozliwa bez backendu (Node + openpyxl lub Excel).
