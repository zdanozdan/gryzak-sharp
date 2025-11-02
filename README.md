# Gryzak - Menedżer Zamówień

Aplikacja desktopowa do zarządzania zamówieniami ze sklepu internetowego, napisana w C# i WPF.

## Funkcjonalności

- 📊 Wyświetlanie listy zamówień w formie tabeli
- 🔍 Filtrowanie zamówień po statusie
- 📈 Statystyki zamówień (łączna liczba i wartość)
- ⚙️ Konfiguracja API z testem połączenia
- 🔄 Automatyczne ładowanie danych
- 📋 Przykładowe dane testowe gdy API nie jest skonfigurowane

## Wymagania

- .NET 8.0 SDK lub nowszy
- Windows 10/11

## Kompilacja

```bash
dotnet build
```

## Budowanie instalatora

Aby utworzyć instalator aplikacji:

### Wymagania

- .NET 8.0 SDK
- Inno Setup Compiler (opcjonalnie, ale zalecane)
  - Pobierz z: https://innosetup.com/
  - Domyślna lokalizacja: `C:\Program Files (x86)\Inno Setup 6\`

### Szybki start

Uruchom w PowerShell:

```powershell
.\create-installer.ps1
```

Skrypt automatycznie:
1. Opublikuje aplikację w trybie Release (x64 i x86)
2. Utworzy instalator używając Inno Setup
3. Umieści plik instalatora w folderze `installer/`

### Opcje

**Tylko publikacja (bez instalatora):**
```powershell
.\create-installer.ps1 -PublishOnly
```
lub
```powershell
.\publish.ps1
```

**Jeśli Inno Setup jest w innej lokalizacji:**
```powershell
.\create-installer.ps1 -InnoSetupPath "C:\Ścieżka\Do\Inno Setup 6\ISCC.exe"
```

### Wynik

Po zakończeniu w folderze `installer/` znajdziesz plik:
- `Gryzak-Setup-1.0.0.exe` - gotowy instalator aplikacji

### Szczegółowa dokumentacja

Zobacz [INSTALLER.md](INSTALLER.md) dla bardziej szczegółowej instrukcji.

## Uruchomienie

```bash
dotnet run
```

Lub po kompilacji:
```bash
dotnet bin/Debug/net8.0-windows/Gryzak.exe
```

## Konfiguracja API

1. Otwórz aplikację
2. Kliknij przycisk "⚙️ Konfiguracja API"
3. Wprowadź:
   - URL API (wymagany)
   - Token API (opcjonalny)
   - Timeout (5-300 sekund)
   - Endpoint listy zamówień
   - Endpoint szczegółów zamówienia
4. Kliknij "🧪 Testuj połączenie" aby sprawdzić konfigurację
5. Zapisz konfigurację

## Przykładowe dane testowe

Gdy API nie jest skonfigurowane, aplikacja wyświetla przykładowe dane testowe z 3 zamówieniami.

## Struktura projektu

```
Gryzak/
├── Models/          # Modele danych
├── Services/        # Serwisy (API, Config)
├── ViewModels/      # ViewModele
├── Views/           # Okna i dialogi
└── App.xaml         # Główny plik aplikacji
```

