$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host "=== DyplomCheck: установка ==="

if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "Python не найден в PATH. Установите Python 3.11+ и повторите." -ForegroundColor Red
    exit 1
}

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    Write-Host "Создаю виртуальное окружение .venv ..."
    python -m venv .venv
}

$python = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
Write-Host "Устанавливаю зависимости ..."
& $python -m pip install --upgrade pip
& $python -m pip install -r requirements.txt

if (-not (Test-Path ".env")) {
    Copy-Item ".env.example" ".env"
    Write-Host "Создан .env. Вставьте в него BOT_TOKEN от @BotFather." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Готово." -ForegroundColor Green
Write-Host "1. Откройте .env и вставьте токен: BOT_TOKEN=123456:AA..."
Write-Host "2. Запустите: .\run.ps1"
