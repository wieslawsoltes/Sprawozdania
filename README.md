# Sprawozdania i analizy (Racibórz / powiat raciborski)

## Źródła danych (katalog `data/`)
- Sprawozdania finansowe jednostek obsługiwanych (2024): https://zopo.bipraciborz.pl/bipkod/40495541  
  - Pobieranie: `download_reports.py` (wyjście w `data/sprawozdania_2024/<placówka>/`).
- Rocznik Demograficzny 2025 (GUS): https://stat.gov.pl/obszary-tematyczne/roczniki-statystyczne/roczniki-statystyczne/rocznik-demograficzny-2025,3,19.html  
  - Pobieranie/tablice: `data/Rocznik2025/` (PDF + pliki XLS/XLSX).
- Prognozy demograficzne GUS 2023–2060 / 2023–2040: https://demografia.stat.gov.pl/BazaDemografia/Prognoza_2023_2060.aspx  
  - Pobieranie/rozpakowanie: `data/GUS/` (pliki prognoz i scenariuszy, w tym powiat raciborski i miasto Racibórz).

## Skrypty
- `download_reports.py` – pobiera PDF-y sprawozdań finansowych 2024 dla wszystkich placówek i zapisuje w oddzielnych katalogach.
- `analyze_financials.py` – parsuje RZiS 2024, buduje arkusz `raport_finansowy_2024.xlsx` z:
  - arkuszem `Zbiorcze_porownanie` (przychody, koszty, wyniki, formuły koszt_na_ucznia),
  - arkuszami per placówka (tabele RZiS),
  - `Pivot_placowka` + `Wykresy` (koszty operacyjne, wynik netto, koszt/uczeń per placówka),
  - `uwagi_nieprawidlowosci.docx` (potencjalne uwagi).
- `fix_financials_excel.py` – poprawia formuły/formaty koszt_na_ucznia w `raport_finansowy_2024.xlsx`, przebudowuje pivot per placówka i wykresy.
- `extract_gus_children.py` – wyciąga z prognoz GUS liczebności dzieci (powiat raciborski 0–2/3–6/7–18; miasto Racibórz grupy dostępne 0–9, 10–19, 0–17) i zapisuje do `GUS/demografia_dzieci.xlsx`.
- `build_demand.py` – na bazie `demografia_dzieci.xlsx` tworzy:
  - `GUS/zapotrzebowanie_miejsc_2023_2060.xlsx` (zapotre­bowanie miejsc = 100% populacji),
  - `GUS/prezentacja_demografia_placowki.pptx` (slajdy z wykresami i założeniami).

## Pliki wynikowe (w `data/`)
- `data/raport_finansowy_2024.xlsx` – dane finansowe 2024 (z formułami), w tym koszt_na_ucznia; `Pivot_placowka` + wykresy per placówka.
- `data/uwagi_nieprawidlowisci.docx` – uwagi do sprawozdań (ujemny wynik, inne koszty operacyjne, brak liczby uczniów).
- `data/GUS/demografia_dzieci.xlsx` – agregaty dzieci (powiat/gmina) z prognoz GUS.
- `data/GUS/zapotrzebowanie_miejsc_2023_2060.xlsx` – zapotrzebowanie na miejsca (żłobek/przedszkole/szkoła) 2023–2060 (powiat) i 2023–2040 (miasto – brak rozbicia 0–2/3–6).
- `data/GUS/prezentacja_demografia_placowki.pptx` – slajdy z wykresami demograficznymi.
- `data/Rocznik2025/` – rocznik demograficzny 2025 (PDF + tablice).

## Uwagi analityczne / ograniczenia
- Miasto Racibórz: prognoza gmin 2023–2040 nie zawiera rozbicia 0–2 i 3–6; potrzebne lokalne dane, by precyzyjnie policzyć zapotrzebowanie żłobek/przedszkole.
- Koszt_na_ucznia: uzupełniony tylko dla placówek z podaną liczbą uczniów; reszta wymaga danych wejściowych.
- Zapotrzebowanie miejsc w `zapotrzebowanie_miejsc_2023_2060.xlsx` = 100% populacji danej grupy (brak współczynników partycypacji).
- Braki do pełnej analizy zamknięć/redukcji placówek:
  - brak liczby uczniów/dzieci dla większości placówek (znane: Przedszkole nr 10 – 150; Przedszkole nr 15 – 100),
  - brak pojemności/obłożenia (miejsca vs faktyczne dzieci) dla żłobków/przedszkoli/szkół,
  - demografia GUS jest agregowana (powiat/miasto), nie ma rozbicia na placówki.
- Gdy dostarczysz (1) liczby uczniów/dzieci i (2) pojemności miejsc per placówka, można wyliczyć:
  - koszt na ucznia dla wszystkich placówek (na bazie `raport_finansowy_2024.xlsx`),
  - wskaźniki wykorzystania miejsc,
  - projekcję popytu (GUS) vs podaż miejsc do 2060,
  - priorytety zamknięć/redukcji (wysoki koszt/uczeń + niskie obłożenie/spadek popytu),
  - rekomendacje utrzymania/inwestycji (niski koszt/uczeń + stabilny/wzrastający popyt).
- Jeśli nie ma pojemności, podaj chociaż aktualne liczby uczniów/dzieci per placówka – pozwoli to wskazać kandydatów do redukcji na podstawie koszt/uczeń i trendu demograficznego (spadek 0–6 / 7–18 w powiecie).

## Jak odtworzyć
```
# środowisko
python3 -m venv .venv
.venv/bin/pip install pandas openpyxl pdfplumber pypdf python-pptx xlrd

# finanse
.venv/bin/python download_reports.py      # zapisuje do data/sprawozdania_2024
.venv/bin/python analyze_financials.py
.venv/bin/python fix_financials_excel.py

# demografia/prognozy
.venv/bin/python extract_gus_children.py
.venv/bin/python build_demand.py
```
