# Historia zmian - Gryzak

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

