@echo off
cd /d "%~dp0"
if exist "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\native\powershell\pwsh.exe" (
  "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\native\powershell\pwsh.exe" -NoProfile -File "%~dp0Start-IAXG.ps1"
) else (
  where pwsh.exe >nul 2>nul
  if errorlevel 1 (
    powershell.exe -NoProfile -File "%~dp0Start-IAXG.ps1"
  ) else (
    pwsh.exe -NoProfile -File "%~dp0Start-IAXG.ps1"
  )
)
if errorlevel 1 pause
