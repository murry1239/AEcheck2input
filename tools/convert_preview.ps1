param([string]$InputPath,[string]$OutputPath)
$ErrorActionPreference='Stop'
$app=$null; $book=$null
try {
    $app=New-Object -ComObject Excel.Application
    $app.Visible=$false; $app.DisplayAlerts=$false; $app.AutomationSecurity=3
    $book=$app.Workbooks.Open((Resolve-Path -LiteralPath $InputPath).Path,0,$true)
    $book.SaveAs([IO.Path]::GetFullPath($OutputPath),51)
} finally {
    if ($book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($app) { $app.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
