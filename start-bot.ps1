$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$python = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    Write-Host "Виртуальное окружение не найдено. Сначала выполните .\setup.ps1" -ForegroundColor Red
    exit 1
}

$pidFile = Join-Path $PSScriptRoot "bot.pid"
if (Test-Path $pidFile) {
    $oldPid = 0
    if ([int]::TryParse((Get-Content $pidFile -Raw).Trim(), [ref]$oldPid)) {
        if (Get-Process -Id $oldPid -ErrorAction SilentlyContinue) {
            Write-Host "Бот уже запущен, PID $oldPid. Сначала выполните .\stop-bot.ps1" -ForegroundColor Yellow
            exit 0
        }
    }
    Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
}

$env:PYTHONUTF8 = "1"
$p = Start-Process -FilePath $python `
    -ArgumentList "-m", "app" `
    -WorkingDirectory $PSScriptRoot `
    -RedirectStandardOutput (Join-Path $PSScriptRoot "bot-run.log") `
    -RedirectStandardError (Join-Path $PSScriptRoot "bot-run.err") `
    -WindowStyle Hidden `
    -PassThru

$p.Id | Set-Content $pidFile
Start-Sleep -Seconds 2

if ($p.HasExited) {
    Write-Host "Бот не запустился. Смотрите bot-run.err:" -ForegroundColor Red
    Get-Content (Join-Path $PSScriptRoot "bot-run.err") -Encoding UTF8 -Tail 30
    Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
    exit 1
}

Write-Host "Бот запущен в фоне. PID $($p.Id)" -ForegroundColor Green
Write-Host "Логи: bot-run.err"
