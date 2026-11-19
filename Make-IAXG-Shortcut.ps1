param([switch]$Quiet)  # -Quiet: 완료 안내 창을 띄우지 않음
$ErrorActionPreference = 'Stop'
# 바탕화면에 IAXG 바로가기를 만든다. 프로그램 폴더를 옮긴 뒤에는 다시 실행하면 된다.
$root = $PSScriptRoot
$desktop = [Environment]::GetFolderPath('Desktop')  # OneDrive로 옮겨진 바탕화면도 찾는다
$link = Join-Path $desktop 'IAXG.lnk'
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($link)
$shortcut.TargetPath = Join-Path $root 'IAXG_실행.cmd'
$shortcut.WorkingDirectory = $root
$shortcut.WindowStyle = 7  # 실행 창은 최소화로 잠깐만 뜬다
$shortcut.IconLocation = (Join-Path $root 'iaxg\iaxg.ico') + ',0'
$shortcut.Description = 'IAXG (IDBI AE XML Generator) 실행'
$shortcut.Save()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
if ($Quiet) { Write-Output $link; return }
Add-Type -AssemblyName System.Windows.Forms
[void][System.Windows.Forms.MessageBox]::Show("바탕화면에 IAXG 바로가기를 만들었습니다.`n$link`n`n작업 표시줄이나 시작 메뉴에 고정하려면 바로가기를 마우스 오른쪽 버튼으로 누르세요.", 'IAXG')
