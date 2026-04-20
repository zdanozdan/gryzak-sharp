# Historia zmian - Gryzak

## 1.6.7 (2026-04-20)

- Pobieranie zamówienia po numerze z API: przycisk lupy przy polu wyszukiwania oraz zaproszenie z lupą przy pustym wyniku filtra (zamiast samego komunikatu „Brak zamówień”).

## 1.6.6 (2026-03-11)

- Rozszerzenie akcji WyPLUwacz: teraz aktualizuje również pola SWW (`tw_SWW`) oraz podstawowy kod kreskowy (`tw_PodstKodKresk`), ustawiając je na wartość identyfikatora towaru.

## 1.6.5 (2026-03-11)

- Ograniczenie sprawdzania istnienia dokumentów tylko do typu ZK (Zamówienie od Klienta), aby uniknąć błędnego łączenia z fakturami zakupu lub innymi dokumentami o tym samym numerze oryginału.

## 1.6.4 (2026-03-11)

- Naprawa błędu SQL "Nieprawidłowa nazwa obiektu dbo.adr__Ewid" (dodanie InitialCatalog do połączenia)
- Gwarancja wyświetlania okna wyboru kontrahenta nawet w przypadku błędów SQL lub braku wyników
- Poprawienie logowania (naprawa ok. 100 błędnych formatowań z brakującym znakiem $)
- Usunięcie oznaczenia ".plu" z wersji aplikacji

## 1.6.3.plu (2026-02-05)

- Dodanie narzędzia WyPLUwacz (limit 100 produktów)
- Nowe modalne okno postępu z obsługą przerywania
- Dodanie parametrów obiektu .gt w ustawieniach (Produkt, Autentykacja, Tryb uruchomienia)
- Rozbudowane logowanie parametrów podczas testu połączenia
- Naprawa błędów kompilacji i ostrzeżeń instalatora

## 1.6.2 (2026-02-05)

- Aktualizacja wersji

## 1.0.1 (2025-11-04)

### Dodane
- Dodano menu "Akcje" z opcjami "Odśwież", "Dodaj ZK" i "Nowe ZK"
- Funkcja "Nowe ZK" - tworzy nowy dokument ZK bez sprawdzania czy dokument już istnieje w Subiekcie GT
- Opcje produktu (z JSON API) są teraz wyświetlane w opisie pozycji ZK w formacie `nazwa:wartość`
- Sprawdzanie czy koszty (KOSZTY/1 i KOSZTY/2) są większe od 0 przed dodaniem pozycji do ZK

### Zmienione
- Ulepszona obsługa opcji produktu - automatyczne dodawanie do opisu pozycji w ZK
- Lepsze logowanie błędów podczas dodawania opcji do pozycji

### Naprawione
- Usunięto ostrzeżenia kompilatora o nieużywanych zmiennych w obsłudze opcji produktu

## 1.0.0 (2025-01-XX)

### Pierwsza wersja
- Podstawowa funkcjonalność zarządzania zamówieniami
- Integracja z API sklepu internetowego
- Integracja z Subiekt GT poprzez Sfera API
- Automatyczne tworzenie dokumentów ZK w Subiekcie GT
- Sprawdzanie istnienia dokumentów ZK w bazie danych
- Interfejs użytkownika WPF
- Konfiguracja połączenia z API i Subiektem GT

