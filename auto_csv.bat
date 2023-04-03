@echo off
powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -File "auto_csv.ps1" -EncodedCommand $(Get-Content "auto_csv.ps1" -Encoding UTF8 | Out-String | ConvertTo-Base64)
pause