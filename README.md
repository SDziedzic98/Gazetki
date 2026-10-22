# Gazetki → Excel

Prosta aplikacja w przeglądarce do odczytywania promocji z gazetek spożywczych i eksportowania ich do pliku Excel. Pliki są przetwarzane lokalnie w przeglądarce — nie są wysyłane na serwer aplikacji.

## Funkcje

- wczytywanie wielu gazetek PDF oraz zdjęć JPG, PNG i WEBP;
- odczyt tekstu z PDF, a dla skanów i zdjęć OCR w języku polskim i angielskim;
- wstępne wyodrębnianie cen, opisów, gramatur i warunków promocji;
- edycja, usuwanie i ręczne dodawanie ofert;
- automatyczne przeliczanie ceny promocyjnej na kilogram lub litr;
- eksport do `.xlsx` z arkuszem ofert i informacjami o eksporcie.

## Uruchomienie

Najprościej opublikować zawartość repozytorium na GitHub Pages (Settings → Pages → Deploy from a branch → `main` → `/ (root)`). Po chwili aplikacja będzie dostępna pod adresem Pages repozytorium.

Możesz też uruchomić ją lokalnie przez serwer statyczny, np.:

```bash
python -m http.server 8000
```

Następnie otwórz `http://localhost:8000`.

Aplikacja korzysta z bibliotek PDF.js, Tesseract.js i SheetJS ładowanych z CDN. Do pierwszego użycia OCR potrzebne jest połączenie z internetem. Tekstowe PDF-y mogą być analizowane bez OCR. W przypadku skanów OCR pobiera polskie dane językowe przy pierwszym uruchomieniu.

## Ograniczenia MVP

Gazetki mają różne układy, a OCR bywa niedokładny — szczególnie przy cenach, małych etykietach, przekreślonych cenach i ofertach wielosztukowych. Wydobywanie nazw i warunków jest heurystyczne. Zawsze sprawdź tabelę z gazetką przed użyciem danych. Ilość należy wpisać w kilogramach lub litrach (np. 0,5 dla 500 g lub 500 ml), aby cena jednostkowa była prawidłowa.

## Prywatność

Aplikacja nie ma backendu. PDF-y i zdjęcia pozostają w pamięci przeglądarki. Biblioteki i dane OCR są pobierane z wymienionych CDN-ów.

## Licencja

MIT — zobacz plik `LICENSE`.

## Automatyczne publikowanie

Repozytorium zawiera workflow GitHub Actions. Po włączeniu GitHub Pages w ustawieniach repozytorium i wskazaniu źródła **GitHub Actions**, każde wypchnięcie zmian do `main` opublikuje aplikację.

## Automatyczne pobieranie gazetki Biedronki

Workflow `Obserwuj gazetkę Biedronki` raz dziennie sprawdza [stronę Biedronki w iUlotka](https://iulotka.pl/biedronka) i pobiera bezpośredni PDF najnowszej gazetki spożywczej. Po wykryciu nowego pliku wykonuje OCR po polsku i angielsku oraz zapisuje skoroszyt Excel i źródłowy PDF jako artefakt uruchomienia GitHub Actions. Nie tworzy kolejnego Excela, jeśli zawartość PDF się nie zmieniła. Można go też uruchomić ręcznie z zakładki **Actions**; opcja „Wymuś ponowny OCR” przetwarza ponownie tę samą gazetkę.

Po zakończeniu znajdziesz pliki w **Actions → Obserwuj gazetkę Biedronki → ostatnie uruchomienie → Artifacts**. Artefakty są dostępne przez 30 dni. Workflow zapisuje w repozytorium sumę SHA-256 i adres źródła ostatniego PDF-u, aby rozpoznać, czy gazetka się zmieniła; nie publikuje samych PDF-ów ani Exceli w repozytorium.

Automatyczny Excel to arkusz roboczy. Każda strona jest odczytywana w rozdzielczości 300 dpi dwoma ustawieniami OCR; skoroszyt pokazuje ich zgodność, średnią pewność Tesseract dla całej linii, tekst z tej samej linii oraz pobliski kontekst. Ceny wykryte bez zgodnego drugiego odczytu lub z niejednoznacznym formatem trafiają do osobnej zakładki **Niejednoznaczne — sprawdź**. Są to wskaźniki jakości odczytu OCR, a nie prawdopodobieństwo prawidłowej oferty: nawet zgodnie odczytana cena może być przypisana do niewłaściwego produktu albo pomijać warunki promocji. Sprawdź wynik z dołączonym PDF-em przed dalszym użyciem. iUlotka podaje Biedronka.pl jako źródło gazetek, a same PDF-y udostępnia przez własne linki do plików.
