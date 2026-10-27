param([string]$Path)
$ErrorActionPreference = 'Stop'
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open((Resolve-Path -LiteralPath $Path).Path, 0, $true)
    $result = @()
    foreach ($sheet in $book.Worksheets) {
        $used = $sheet.UsedRange
        $values = $used.Value2
        $rows = @()
        for ($r = 1; $r -le $used.Rows.Count; $r++) {
            $line = @()
            for ($c = 1; $c -le [Math]::Min(13, $used.Columns.Count); $c++) {
                $line += $values.GetValue($r, $c)
            }
            $rows += ,$line
        }
        $result += [pscustomobject]@{ name=$sheet.Name; count=$used.Rows.Count; columns=$used.Columns.Count; rows=$rows }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($used)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)
    }
    $result | ConvertTo-Json -Depth 8
} finally {
    if ($book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
