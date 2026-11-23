$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$python = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    Write-Host "Виртуальное окружение не найдено. Сначала выполните .\setup.ps1" -ForegroundColor Red
    exit 1
}

$env:PYTHONUTF8 = "1"
& $python -m app
