$ErrorActionPreference = 'SilentlyContinue'
$app = Join-Path $env:LOCALAPPDATA 'Programs\Folio'
Get-Process Folio | Stop-Process -Force
Start-Sleep -Milliseconds 400
$cls = 'HKCU:\Software\Classes'
foreach ($ext in '.docx', '.txt', '.md', '.markdown', '.log', '.htm', '.html') {
    Remove-ItemProperty "$cls\$ext\OpenWithProgids" -Name 'Folio.Document'
}
if ((Get-ItemProperty "$cls\.docx").'(default)' -eq 'Folio.Document') { Set-ItemProperty "$cls\.docx" '(default)' '' }
Remove-Item "$cls\Folio.Document" -Recurse -Force
Remove-Item "$cls\Applications\Folio.exe" -Recurse -Force
Remove-Item 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Folio' -Recurse -Force
Remove-Item 'HKCU:\Software\Folio' -Recurse -Force
Remove-Item (Join-Path ([Environment]::GetFolderPath('Programs')) 'Folio.lnk') -Force
Remove-Item (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Folio.lnk') -Force
# папку программы удаляем после выхода из скрипта (из неё он и запущен)
Start-Process cmd -ArgumentList "/c timeout /t 2 >nul & rmdir /s /q `"$app`"" -WindowStyle Hidden
Write-Host "Folio удалён." -ForegroundColor Green
Start-Sleep -Seconds 2
