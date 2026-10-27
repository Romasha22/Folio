# Folio — установщик. Собирает программу и ставит её только для текущего пользователя.
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$here  = Split-Path -Parent $MyInvocation.MyCommand.Path
$cache = Join-Path $env:LOCALAPPDATA 'Folio-build'
$app   = Join-Path $env:LOCALAPPDATA 'Programs\Folio'
$exe   = Join-Path $app 'Folio.exe'

function Step($n, $text) { Write-Host ""; Write-Host "[$n/3] $text" -ForegroundColor Cyan }

Write-Host ""
Write-Host "  Folio — установка" -ForegroundColor White
Write-Host "  Это займёт 3–5 минут. Окно закроется само в конце." -ForegroundColor DarkGray

# 1. Python
Step 1 "Ищу Python…"
function Get-Python {
    $cands = @()
    if (Get-Command py -ErrorAction SilentlyContinue) { $cands += ,@('py', '-3') }
    if (Get-Command python -ErrorAction SilentlyContinue) { $cands += ,@('python') }
    Get-ChildItem "$env:LOCALAPPDATA\Programs\Python\Python3*\python.exe" -ErrorAction SilentlyContinue |
        Sort-Object FullName -Descending | ForEach-Object { $cands += ,@($_.FullName) }
    foreach ($c in $cands) {
        try {
            $args2 = @(); if ($c.Count -gt 1) { $args2 = $c[1..($c.Count-1)] }
            $out = & $c[0] @args2 -c "import sys; print(sys.version_info[:2] >= (3,9)); print(sys.executable)" 2>$null
            if ($LASTEXITCODE -eq 0 -and $out -and $out[0] -eq 'True') { return $out[1] }
        } catch {}
    }
    return $null
}
$py = Get-Python
if (-not $py) {
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Write-Host "  Python не найден — ставлю его через winget (нужен один раз)…"
        winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements --silent | Out-Host
        $env:Path = [Environment]::GetEnvironmentVariable('Path','User') + ';' + [Environment]::GetEnvironmentVariable('Path','Machine')
        $py = Get-Python
    }
}
if (-not $py) {
    Write-Host ""
    Write-Host "  Не получилось найти или поставить Python." -ForegroundColor Yellow
    Write-Host "  Открою сайт python.org: скачай Python 3.12, при установке поставь галочку 'Add python.exe to PATH',"
    Write-Host "  потом запусти Install.bat ещё раз."
    Start-Process 'https://www.python.org/downloads/windows/'
    Read-Host "  Нажми Enter, чтобы закрыть"
    exit 1
}
Write-Host "  Python: $py" -ForegroundColor DarkGray

# 2. Библиотеки
Step 2 "Скачиваю библиотеки (PySide6, python-docx)…"
if (Test-Path $cache) { Remove-Item $cache -Recurse -Force }
New-Item -ItemType Directory -Path $cache | Out-Null
& $py -m venv "$cache\venv"
$vpy = "$cache\venv\Scripts\python.exe"
& $vpy -m pip install --upgrade pip --quiet --disable-pip-version-check
& $vpy -m pip install pyside6-essentials python-docx pyinstaller --quiet --disable-pip-version-check
if ($LASTEXITCODE -ne 0) { throw "pip не смог поставить библиотеки (проверь интернет)." }

# 3. Сборка
Step 3 "Собираю Folio.exe…"
Copy-Item "$here\..\folio.py", "$here\..\icon.ico" $cache
Push-Location $cache
& $vpy -m PyInstaller --noconfirm --clean --windowed --name Folio --icon icon.ico `
    --add-data "icon.ico;." --log-level WARN folio.py
Pop-Location
if (-not (Test-Path "$cache\dist\Folio\Folio.exe")) { throw "Сборка не удалась." }

# 4–5. Установка готовой сборки
Copy-Item "$here\Uninstall.bat", "$here\uninstall.ps1" "$cache\dist\Folio"
& "$here\setup.ps1" -Source "$cache\dist\Folio"
Remove-Item $cache -Recurse -Force -ErrorAction SilentlyContinue
