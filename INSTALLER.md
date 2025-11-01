# Instrukcja tworzenia instalatora dla Gryzak

## Wymagania

1. **.NET 8.0 SDK** - do kompilacji aplikacji
2. **Inno Setup** (opcjonalnie, ale zalecane) - do tworzenia profesjonalnego instalatora
   - Pobierz z: https://innosetup.com/
   - Domyślna lokalizacja: `C:\Program Files (x86)\Inno Setup 6\`

## Szybki start

### Opcja 1: Automatyczne tworzenie instalatora (zalecane)

Uruchom w PowerShell:

```powershell
.\create-installer.ps1
```

Skrypt automatycznie:
1. Opublikuje aplikację w trybie Release
2. Utworzy instalator używając Inno Setup
3. Umieści plik instalatora w folderze `installer/`

### Opcja 2: Tylko publikacja (bez instalatora)

```powershell
.\create-installer.ps1 -PublishOnly
```

lub użyj istniejącego skryptu:

```powershell
.\publish.ps1
```

### Opcja 3: Jeśli Inno Setup jest w innej lokalizacji

```powershell
.\create-installer.ps1 -InnoSetupPath "C:\Ścieżka\Do\Inno Setup 6\ISCC.exe"
```

## Struktura po kompilacji

```
project-root/
├── publish/
│   └── win-x64/
│       └── Gryzak.exe          # Opublikowana aplikacja
└── installer/
    └── Gryzak-Setup-1.0.0.exe  # Instalator
```

## Ręczne tworzenie instalatora

Jeśli nie masz Inno Setup, możesz:

1. **Użyć publikacji bezpośrednio:**
   - Folder `publish/win-x64/` zawiera gotową aplikację
   - Możesz ją spakować w ZIP i rozpakować na docelowym komputerze

2. **Użyć alternatywnych narzędzi:**
   - **NSIS** (Nullsoft Scriptable Install System)
   - **WiX Toolset** (Windows Installer XML)
   - **Advanced Installer**
   - **7-Zip** SFX (self-extracting archive)

## Konfiguracja instalatora

Plik `setup.iss` zawiera konfigurację instalatora. Możesz edytować:

- **Nazwę aplikacji:** `#define MyAppName`
- **Wersję:** `#define MyAppVersion`
- **Wydawcę:** `#define MyAppPublisher`
- **Lokalizację instalacji:** `DefaultDirName`
- **Ikony i skróty:** Sekcja `[Icons]`

## Instalacja na komputerze użytkownika

Użytkownik powinien:

1. Uruchomić `Gryzak-Setup-1.0.0.exe`
2. Postępować zgodnie z kreatorem instalacji
3. Wybrać folder docelowy (domyślnie: `C:\Program Files\Gryzak`)
4. Opcjonalnie utworzyć skróty na pulpicie i w menu Start

## Uwagi

- Instalator wymaga uprawnień administratora (UAC)
- Aplikacja jest self-contained (nie wymaga .NET Runtime)
- Rozmiar instalatora to około 50-100 MB (w zależności od bibliotek)
- Instalator automatycznie sprawdza, czy aplikacja jest uruchomiona i zamyka ją przed instalacją

