# Skrypt do tworzenia instalatora dla Gryzak
# Wymaga Inno Setup Compiler (ISCC.exe) w PATH lub podaj pelna sciezke

param(
    [string]$InnoSetupPath = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    [switch]$BuildFirst = $true,
    [switch]$PublishOnly = $false
)

Write-Host "=== Tworzenie instalatora Gryzak ===" -ForegroundColor Green

# Krok 1: Publikacja aplikacji
if ($BuildFirst -and -not $PublishOnly) {
    Write-Host ""
    Write-Host "1. Publikacja aplikacji w trybie Release..." -ForegroundColor Yellow
    
    if (Test-Path "publish\win-x64") {
        Remove-Item "publish\win-x64" -Recurse -Force -ErrorAction SilentlyContinue
    }
    
    dotnet publish -c Release -r win-x64 --self-contained true `
        -p:PublishSingleFile=true `
        -p:IncludeNativeLibrariesForSelfExtract=true `
        -p:EnableCompressionInSingleFile=true `
        -o publish/win-x64
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Blad publikacji aplikacji" -ForegroundColor Red
        exit 1
    }
    
    Write-Host "Aplikacja opublikowana do publish/win-x64" -ForegroundColor Green
    
    # Wyczysc niepotrzebne pliki
    Get-ChildItem -Path "publish\win-x64" -Filter "*.pdb" | Remove-Item -Force -ErrorAction SilentlyContinue
}

if ($PublishOnly) {
    Write-Host "Publikacja zakonczona (pomijam tworzenie instalatora)" -ForegroundColor Green
    exit 0
}

# Krok 2: Sprawdz czy Inno Setup jest dostepny
if (-not (Test-Path $InnoSetupPath)) {
    Write-Host ""
    Write-Host "Nie znaleziono Inno Setup Compiler w: $InnoSetupPath" -ForegroundColor Red
    Write-Host ""
    Write-Host "Mozesz:" -ForegroundColor Yellow
    Write-Host "  1. Zainstalowac Inno Setup z: https://innosetup.com/" -ForegroundColor White
    Write-Host "  2. Podac sciezke do ISCC.exe przez parametr -InnoSetupPath" -ForegroundColor White
    Write-Host "  3. Dodac Inno Setup do PATH systemowego" -ForegroundColor White
    Write-Host ""
    Write-Host "Przyklad:" -ForegroundColor Cyan
    Write-Host "  .\create-installer.ps1 -InnoSetupPath 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe'" -ForegroundColor White
    exit 1
}

# Krok 3: Sprawdz czy plik publish istnieje
if (-not (Test-Path "publish\win-x64\Gryzak.exe")) {
    Write-Host ""
    Write-Host "Nie znaleziono opublikowanej aplikacji w publish\win-x64\Gryzak.exe" -ForegroundColor Red
    Write-Host "Uruchom najpierw: .\create-installer.ps1 -BuildFirst" -ForegroundColor Yellow
    exit 1
}

# Krok 4: Sprawdz czy plik setup.iss istnieje
if (-not (Test-Path "setup.iss")) {
    Write-Host ""
    Write-Host "Nie znaleziono pliku setup.iss" -ForegroundColor Red
    exit 1
}

# Krok 5: Utworz folder installer jesli nie istnieje
if (-not (Test-Path "installer")) {
    New-Item -ItemType Directory -Path "installer" | Out-Null
}

# Krok 6: Kompiluj instalator
Write-Host ""
Write-Host "2. Kompilowanie instalatora..." -ForegroundColor Yellow
$issFile = Join-Path $PWD "setup.iss"
$argumentList = [string]::Format('"{0}"', $issFile)
$process = Start-Process -FilePath $InnoSetupPath -ArgumentList $argumentList -Wait -PassThru -NoNewWindow

if ($process.ExitCode -eq 0) {
    Write-Host "Instalator utworzony pomyslnie!" -ForegroundColor Green
    
    # Znajdz utworzony plik instalatora
    $installerFile = Get-ChildItem -Path "installer" -Filter "Gryzak-Setup-*.exe" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    
    if ($installerFile) {
        $sizeMB = [math]::Round($installerFile.Length / 1MB, 2)
        Write-Host ""
        Write-Host "=== Instalator gotowy ===" -ForegroundColor Green
        Write-Host "Plik: $($installerFile.FullName)" -ForegroundColor Cyan
        Write-Host "Rozmiar: $sizeMB MB" -ForegroundColor Cyan
        
        # Otworz folder z instalatorem
        Start-Process explorer.exe -ArgumentList "/select,$($installerFile.FullName)"
    }
} else {
    Write-Host "Blad kompilacji instalatora (kod wyjscia: $($process.ExitCode))" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "=== Gotowe ===" -ForegroundColor Green
