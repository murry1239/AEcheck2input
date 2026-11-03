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
Add-Type -AssemblyName System.Windows.Forms
if (-not $pythonPath) {
    [void][System.Windows.Forms.MessageBox]::Show('AAI 실행에 필요한 Python이 없습니다. Python 설치 후 다시 실행해 주세요.', 'AAI')
    exit 1
}
# ACI 작업 폴더의 outputs에서 시작해 ACI 결과 엑셀을 고른다.
$start = $root
$settings = Join-Path $root 'aci_settings.json'
if (Test-Path -LiteralPath $settings) {
    $workspace = (Get-Content -LiteralPath $settings -Raw -Encoding UTF8 | ConvertFrom-Json).workspace
    if ($workspace -and (Test-Path -LiteralPath (Join-Path $workspace 'outputs'))) { $start = Join-Path $workspace 'outputs' }
    elseif ($workspace -and (Test-Path -LiteralPath $workspace)) { $start = $workspace }
}
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = 'AAI: ACI 결과 엑셀(IDBI_YYYYMMDD_ACI.xls)을 선택하세요'
$dialog.Filter = 'ACI 결과 (*_ACI.xls;*_ACI.xlsx)|*_ACI.xls;*_ACI.xlsx|Excel (*.xls;*.xlsx)|*.xls;*.xlsx'
$dialog.InitialDirectory = $start
if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { exit 0 }
Write-Host 'AAI 0.1: Excel로 파일을 읽고 E2B(R3) XML을 만드는 중입니다. 1~2분 걸릴 수 있습니다...'
& $pythonPath (Join-Path $root 'aai\main.py') $dialog.FileName
if ($LASTEXITCODE) {
    [void][System.Windows.Forms.MessageBox]::Show("AAI 처리 중 오류가 났습니다. logs 폴더의 aai_날짜.log를 확인해 주세요.", 'AAI')
    exit 1
}
