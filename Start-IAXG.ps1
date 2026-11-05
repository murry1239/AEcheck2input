$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$candidates = @(
    (Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'),
    (Join-Path $env:ProgramFiles 'LibreOffice\program\python.exe'),
    (Join-Path ${env:ProgramFiles(x86)} 'LibreOffice\program\python.exe')
)
$pythonPath = $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $pythonPath) {
    $pythonCommand = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($pythonCommand -and $pythonCommand.Source -notmatch 'WindowsApps') { $pythonPath = $pythonCommand.Source }
}
if (-not $pythonPath) {
    Add-Type -AssemblyName System.Windows.Forms
    [void][System.Windows.Forms.MessageBox]::Show('IAXG 실행에 필요한 Python이 없습니다. Python 또는 LibreOffice 설치 후 다시 실행해 주세요.', 'IAXG')
    exit 1
}
$appPath = Join-Path $root 'iaxg\app.py'
Start-Process -FilePath $pythonPath -ArgumentList @('"' + $appPath + '"') -WorkingDirectory $root -WindowStyle Hidden
