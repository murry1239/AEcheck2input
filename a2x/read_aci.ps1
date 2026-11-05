param(
    [string]$InputPath,
    [string]$JsonOut
)
# ACI 결과 엑셀의 A~M열을 읽기 전용으로 읽어 JSON으로 저장한다.
$ErrorActionPreference = 'Stop'
$excel = $null; $book = $null
$headers = @('구분','일자','성별','나이','품목명','문의종류','불만사항','AE')

function Write-Json($Object, $Path) {
    $text = ConvertTo-Json -InputObject $Object -Depth 30
    [IO.File]::WriteAllText($Path, $text, (New-Object Text.UTF8Encoding($false)))
}

try {
    $inputFile = (Resolve-Path -LiteralPath $InputPath).Path
    $hash = (Get-FileHash -LiteralPath $inputFile -Algorithm SHA256).Hash.ToLowerInvariant()
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($inputFile,0,$true)
    $sheets = @()
    foreach ($sheet in $book.Worksheets) {
        $used = $sheet.UsedRange
        $rowCount = $used.Row + $used.Rows.Count - 1
        if ($rowCount -lt 2) { continue }
        $range = $sheet.Range($sheet.Cells.Item(1,1),$sheet.Cells.Item($rowCount,13))
        $data = $range.Value2
        $first = @(1..8 | ForEach-Object { ([string]$data.GetValue(1,$_)).Replace(' ','').Trim() })
        if (($first -join '|') -eq ($headers -join '|')) {
            $records = @()
            for ($r=2;$r -le $rowCount;$r++) {
                $values = @(1..13 | ForEach-Object { $v=$data.GetValue($r,$_); if ($null -eq $v) { '' } else { [string]$v } })
                if (-not ($values | Where-Object { $_ -ne '' })) { continue }
                $records += @{row=$r;values=$values}
            }
            $sheets += @{name=$sheet.Name;records=$records}
        }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($range)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($used)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)
    }
    if (-not $sheets.Count) { throw '첫 행 A~H 제목이 ACI 결과 양식(구분…불만사항, AE)과 일치하지 않습니다. ACI 결과 파일(_ACI.xls)을 선택하세요.' }
    Write-Json @{path=$inputFile;name=[IO.Path]::GetFileName($inputFile);source_hash=$hash;sheets=$sheets} $JsonOut
} catch {
    Write-Json @{error=$_.Exception.Message} $JsonOut
    throw
} finally {
    if ($book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
