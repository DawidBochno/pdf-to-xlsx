# PDF to XLSX (program lokalny)

Program **lokalny** — działa w całości na Twoim komputerze (Python + tkinter),
nie łączy się z internetem i nigdzie nie wysyła Twoich plików PDF ani danych.

Program z GUI (tkinter) do wyciągania tabel z plików PDF do Excela (.xlsx).
Radzi sobie z PDF-ami bez linii siatki tabeli — dane ułożone "wierszami"
(wiersz, mały odstęp, kolejny wiersz) są rozpoznawane po współrzędnych
tekstu, nie po liniach.

## Szybki start

1. `install.bat` — instaluje Pythona (jeśli brak, prosi o ręczną instalację)
   oraz biblioteki `pdfplumber` i `openpyxl`, uruchamia self-test.
2. `uruchom.bat` — otwiera GUI.

Wymaga zainstalowanego Pythona 3.9+ z opcją "Add python.exe to PATH" oraz
"tcl/tk and IDLE" (potrzebne do GUI).

## Użycie

W GUI wskaż:
- **Plik PDF lub folder** — pojedynczy plik albo folder z wieloma PDF-ami
  (tryb wsadowy — każdy plik trafia do osobnego .xlsx).
- **Folder wyjściowy** — gdzie zapisać wynikowe pliki .xlsx.
- **Tryb**:
  - `auto` (domyślny) — najpierw próbuje wykryć tabele po liniach siatki
    PDF, a gdy ich nie ma, dzieli tekst po współrzędnych.
  - `tabele` — wymusza wykrywanie po liniach siatki.
  - `wiersze` — wymusza dzielenie po współrzędnych (dla PDF-ów bez siatki).

### Nagłówek / stopka i podział na kolumny

- **Obetnij górę / dół [pt]** — wycina pas strony (np. nagłówek firmy,
  stopka z numerem strony) przed ekstrakcją.
- **Tolerancja wiersza [pt]** — jak blisko pionowo muszą być fragmenty
  tekstu, żeby uznać je za ten sam wiersz (zwiększ, jeśli jeden wiersz
  danych rozjeżdża się na dwa).
- **Min. przerwa kolumny [pt]** — minimalna pionowa przerwa bez tekstu,
  żeby uznać ją za granicę kolumny (zwiększ, jeśli kolumny dzielą się
  za często; zmniejsz, jeśli sąsiednie kolumny się sklejają).
- **Usuń powtarzalne nagłówki/stopki** — automatycznie usuwa wiersze,
  które powtarzają się identycznie na większości stron (np. tytuł
  dokumentu i stopka drukowane na każdej stronie).
- **Wszystkie strony na jednym arkuszu** — inaczej każda strona trafia
  do osobnego arkusza w tym samym pliku .xlsx.
- **Konwertuj liczby** — zamienia tekst typu `1 234,56` na liczbę Excela.

## Ograniczenia

- Działa tylko na PDF-ach z warstwą tekstową — skany bez OCR nie są
  obsługiwane.
- Wiersze zawijane na dwie linie (np. długa nazwa produktu) nie są
  automatycznie scalane w jeden wiersz tabeli.

## Testy

```bash
python pdf_to_xlsx.py --selftest
```

Uruchamia test logiki grupowania w wiersze, wykrywania kolumn i
konwersji wartości liczbowych.
