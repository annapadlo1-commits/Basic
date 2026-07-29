# Ustawienia 2.2.1 — Install Fix

## Naprawiony błąd

Rekord `SEZONOWY` w tabeli `SCENARIUSZE` zawierał 10 wartości przy 11 kolumnach. Brakowało pola `AKTYWNY`.

## Zabezpieczenie

Instalator przed każdym zapisem:

- odczytuje rzeczywistą liczbę kolumn arkusza;
- sprawdza szerokość każdego wiersza demonstracyjnego;
- uzupełnia brakujące końcowe wartości pustymi komórkami;
- zgłasza nazwę tabeli i numer wiersza, jeśli danych jest więcej niż kolumn.

Poprawka jest bezpieczna dla częściowo wykonanej instalacji. Wystarczy podmienić `Code.gs` i ponownie uruchomić `dbInstall`.
