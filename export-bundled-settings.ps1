# Eksportuje bieżące ustawienia Gryzak (%AppData%) do gryzak-ustawienia.json
# i kopiuje plik do katalogów publish (bundel instalatora).

param(
    [Parameter(Mandatory = $true)]
    [string]$GryzakExe,
    [string]$OutputFile = ""
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $GryzakExe)) {
    Write-Host "Nie znaleziono Gryzak.exe: $GryzakExe" -ForegroundColor Red
    exit 1
}

if ([string]::IsNullOrWhiteSpace($OutputFile)) {
    $OutputFile = Join-Path $PSScriptRoot "gryzak-ustawienia.json"
}

Write-Host "  Eksport z AppData przez: $GryzakExe" -ForegroundColor Cyan
Write-Host "  Cel: $OutputFile" -ForegroundColor Cyan

$process = Start-Process -FilePath $GryzakExe `
    -ArgumentList @("--export-settings", $OutputFile) `
    -Wait -PassThru -NoNewWindow

if ($process.ExitCode -ne 0) {
    Write-Host "Eksport ustawien nie powiodl sie (kod $($process.ExitCode))" -ForegroundColor Red
    exit $process.ExitCode
}

if (-not (Test-Path $OutputFile)) {
    Write-Host "Brak pliku po eksporcie: $OutputFile" -ForegroundColor Red
    exit 1
}

foreach ($dir in @("publish\win-x64", "publish\win-x86")) {
    $targetDir = Join-Path $PSScriptRoot $dir
    if (Test-Path $targetDir) {
        Copy-Item -Path $OutputFile -Destination (Join-Path $targetDir "gryzak-ustawienia.json") -Force
        Write-Host "  Skopiowano do $dir" -ForegroundColor Green
    }
}

Write-Host "  Ustawienia wyeksportowane" -ForegroundColor Green
exit 0
