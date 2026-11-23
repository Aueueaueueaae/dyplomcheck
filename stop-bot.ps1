$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$pidFile = Join-Path $PSScriptRoot "bot.pid"
if (-not (Test-Path $pidFile)) {
    Write-Host "bot.pid не найден: похоже, бот запущен не через .\start-bot.ps1" -ForegroundColor Yellow
    exit 1
}

$botPid = 0
if (-not [int]::TryParse((Get-Content $pidFile -Raw).Trim(), [ref]$botPid)) {
    Write-Host "В bot.pid некорректный PID" -ForegroundColor Red
    Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
    exit 1
}

$process = Get-Process -Id $botPid -ErrorAction SilentlyContinue
if ($process) {
    Stop-Process -Id $botPid -Force
    Start-Sleep -Milliseconds 500
    Write-Host "Бот остановлен, PID $botPid" -ForegroundColor Green
} else {
    Write-Host "Процесс с PID $botPid уже не запущен" -ForegroundColor Yellow
}

Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
