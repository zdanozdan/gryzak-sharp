# Skrypt do publikacji aplikacji Gryzak

Write-Host "=== Publikacja Gryzak do single-file executables ===" -ForegroundColor Green

# Tryb Release z single-file
Write-Host ""
Write-Host "1. Publikacja w trybie Release (Windows x64)..." -ForegroundColor Yellow
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o publish/win-x64

if ($LASTEXITCODE -eq 0) {
    Write-Host "Opublikowano do publish/win-x64" -ForegroundColor Green
} else {
    Write-Host "Blad publikacji" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "2. Publikacja w trybie Release (Windows x86)..." -ForegroundColor Yellow
dotnet publish -c Release -r win-x86 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o publish/win-x86

if ($LASTEXITCODE -eq 0) {
    Write-Host "Opublikowano do publish/win-x86" -ForegroundColor Green
} else {
    Write-Host "Blad publikacji" -ForegroundColor Red
    exit 1
}

# Wyswietl informacje o plikach
Write-Host ""
Write-Host "=== Publikacja zakonczona ===" -ForegroundColor Green
Write-Host ""
Write-Host "Pliki dostepne w:" -ForegroundColor Cyan
Write-Host "  - publish/win-x64/Gryzak.exe" -ForegroundColor White
Write-Host "  - publish/win-x86/Gryzak.exe" -ForegroundColor White

Write-Host ""
Write-Host "Aby utworzyc instalator:" -ForegroundColor Cyan
Write-Host "  Uruchom: .\create-installer.ps1" -ForegroundColor White
Write-Host "  (Wymaga Inno Setup - https://innosetup.com/)" -ForegroundColor Gray

# Pokaz rozmiary plikow
if (Test-Path "publish/win-x64/Gryzak.exe") {
    $x64Size = (Get-Item "publish/win-x64/Gryzak.exe").Length / 1MB
    Write-Host ""
    Write-Host "  Rozmiar win-x64: $([math]::Round($x64Size, 2)) MB" -ForegroundColor White
}

if (Test-Path "publish/win-x86/Gryzak.exe") {
    $x86Size = (Get-Item "publish/win-x86/Gryzak.exe").Length / 1MB
    Write-Host "  Rozmiar win-x86: $([math]::Round($x86Size, 2)) MB" -ForegroundColor White
}
