# Folio — установка готовой программы (без Python).
param([string]$Source)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Source) { $Source = Join-Path $here 'Folio' }
$app = Join-Path $env:LOCALAPPDATA 'Programs\Folio'
$exe = Join-Path $app 'Folio.exe'
function Step($n, $text) { Write-Host ""; Write-Host "[$n/2] $text" -ForegroundColor Cyan }
if (-not (Test-Path (Join-Path $Source 'Folio.exe'))) { throw "Не нашёл Folio.exe рядом с установщиком. Сначала распакуй zip целиком." }
Write-Host ""
Write-Host "  Folio — установка" -ForegroundColor White
# 4. Установка
Step 1 "Устанавливаю в $app…"
Get-Process Folio -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 400
if (Test-Path $app) { Remove-Item $app -Recurse -Force }
New-Item -ItemType Directory -Path (Split-Path $app) -Force | Out-Null
Copy-Item $Source $app -Recurse

$ws = New-Object -ComObject WScript.Shell
$links = @(
    (Join-Path ([Environment]::GetFolderPath('Programs')) 'Folio.lnk'),
    (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Folio.lnk')
)
foreach ($l in $links) {
    $s = $ws.CreateShortcut($l); $s.TargetPath = $exe; $s.WorkingDirectory = $app
    $s.IconLocation = "$exe,0"; $s.Description = 'Folio — текстовый редактор'; $s.Save()
}

# 5. «Открыть с помощью»
Step 2 "Добавляю Folio в «Открыть с помощью»…"
$cls = 'HKCU:\Software\Classes'
$open = "`"$exe`" `"%1`""
New-Item "$cls\Folio.Document\DefaultIcon" -Force | Out-Null
New-Item "$cls\Folio.Document\shell\open\command" -Force | Out-Null
Set-ItemProperty "$cls\Folio.Document" '(default)' 'Документ'
Set-ItemProperty "$cls\Folio.Document" 'FriendlyTypeName' 'Документ'
Set-ItemProperty "$cls\Folio.Document\DefaultIcon" '(default)' "$exe,0"
Set-ItemProperty "$cls\Folio.Document\shell\open\command" '(default)' $open

New-Item "$cls\Applications\Folio.exe\shell\open\command" -Force | Out-Null
New-Item "$cls\Applications\Folio.exe\SupportedTypes" -Force | Out-Null
Set-ItemProperty "$cls\Applications\Folio.exe" 'FriendlyAppName' 'Folio'
Set-ItemProperty "$cls\Applications\Folio.exe\shell\open\command" '(default)' $open

foreach ($ext in '.docx', '.txt', '.md', '.markdown', '.log', '.htm', '.html') {
    New-Item "$cls\$ext\OpenWithProgids" -Force | Out-Null
    New-ItemProperty "$cls\$ext\OpenWithProgids" -Name 'Folio.Document' -Value '' -PropertyType String -Force | Out-Null
    New-ItemProperty "$cls\Applications\Folio.exe\SupportedTypes" -Name $ext -Value '' -PropertyType String -Force | Out-Null
}
# .docx без Word: сделаем Folio программой по умолчанию, если другой нет
$docx = (Get-ItemProperty "$cls\.docx" -ErrorAction SilentlyContinue).'(default)'
if (-not $docx) { Set-ItemProperty "$cls\.docx" '(default)' 'Folio.Document' }

$un = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Folio'
New-Item $un -Force | Out-Null
Set-ItemProperty $un 'DisplayName' 'Folio'
Set-ItemProperty $un 'DisplayIcon' "$exe,0"
Set-ItemProperty $un 'Publisher' 'Folio'
Set-ItemProperty $un 'InstallLocation' $app
Set-ItemProperty $un 'UninstallString' "`"$app\Uninstall.bat`""
New-ItemProperty $un -Name 'NoModify' -Value 1 -PropertyType DWord -Force | Out-Null

Add-Type -Namespace W -Name S -MemberDefinition '[DllImport("shell32.dll")] public static extern void SHChangeNotify(int e, int f, System.IntPtr a, System.IntPtr b);'
[W.S]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)

Write-Host ""
Write-Host "  Готово! Folio на рабочем столе и в меню Пуск." -ForegroundColor Green
Start-Process $exe
Start-Sleep -Seconds 2
