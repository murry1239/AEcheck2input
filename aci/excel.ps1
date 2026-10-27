param(
    [ValidateSet('read','export')][string]$Mode,
    [string]$InputPath,
    [string]$JsonOut,
    [string]$PayloadPath,
    [string]$OutputPath,
    [string]$PreviewPath
)
$ErrorActionPreference = 'Stop'
$excel = $null; $book = $null
$headers = @('구분','일자','성별','나이','품목명','문의종류','불만사항')
$extra = @('AE','Cube','SAE','relatedness','QC','comment')

function Write-Json($Object, $Path) {
    $text = ConvertTo-Json -InputObject $Object -Depth 30
    [IO.File]::WriteAllText($Path, $text, (New-Object Text.UTF8Encoding($false)))
}
function Get-Hash($Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
function Get-Values($Sheet) {
    $used = $Sheet.UsedRange
    $rowCount = $used.Row + $used.Rows.Count - 1
    $colCount = $used.Column + $used.Columns.Count - 1
    $range = $Sheet.Range($Sheet.Cells.Item(1,1),$Sheet.Cells.Item($rowCount,$colCount))
    $data = $range.Value2; $formulas = $range.Formula
    $rows = @(); $formulaRows = @()
    for ($r=1;$r -le $rowCount;$r++) {
        $line = @(); $f = @()
        for ($c=1;$c -le $colCount;$c++) {
            if ($rowCount -eq 1 -and $colCount -eq 1) { $value=$data; $formula=$formulas }
            else { $value=$data.GetValue($r,$c); $formula=$formulas.GetValue($r,$c) }
            if ($null -eq $value) { $value='' }
            if ($null -eq $formula) { $formula='' }
            $line += $value; $f += $formula
        }
        $rows += ,$line; $formulaRows += ,$f
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($range)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($used)
    return @{name=$Sheet.Name; rows=$rows; formulas=$formulaRows; count=$rowCount; columns=$colCount}
}
function Get-Snapshot($Book) {
    $list = @()
    foreach ($sheet in $Book.Worksheets) {
        $list += (Get-Values $sheet)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)
    }
    return ,$list
}

try {
    $inputFile = (Resolve-Path -LiteralPath $InputPath).Path
    $sourceHash = Get-Hash $inputFile
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($inputFile,0,$true)
    $snapshot = Get-Snapshot $book
    if ($Mode -eq 'read') {
        $sheets = @()
        foreach ($s in $snapshot) {
            if ($s.columns -lt 7) { continue }
            $first = @($s.rows[0][0..6] | ForEach-Object { ([string]$_).Replace(' ','').Trim() })
            if (($first -join '|') -ne ($headers -join '|')) { continue }
            $records = @()
            for ($row=1;$row -lt $s.count;$row++) {
                $values = @($s.rows[$row][0..6])
                if (-not ($values | Where-Object { $_ -ne '' })) { continue }
                $records += @{row=$row+1;values=$values}
            }
            $sheets += @{name=$s.name;records=$records}
        }
        if (-not $sheets.Count) { throw '첫 행 A~G 제목이 IDBI 양식과 일치하지 않습니다.' }
        Write-Json @{path=$inputFile;name=[IO.Path]::GetFileName($inputFile);source_hash=$sourceHash;sheets=$sheets} $JsonOut
    } else {
        $payload = Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($sourceHash -ne $payload.source_hash) { throw '검토 이후 원본이 변경되었습니다. 다시 불러오세요.' }
        $target = [IO.Path]::GetFullPath($OutputPath)
        if ($target -eq $inputFile -or (Test-Path -LiteralPath $target)) { throw '원본 또는 기존 결과는 덮어쓸 수 없습니다.' }
        $groups = $payload.records | Group-Object sheet
        foreach ($group in $groups) {
            $sheet = $book.Worksheets.Item($group.Name)
            $sheet.Activate()
            for ($c=0;$c -lt 6;$c++) {
                $cell = $sheet.Cells.Item(1,$c+8)
                if ($cell.Value2 -and $cell.Value2 -ne $extra[$c]) { throw 'H~M에 다른 제목이 있습니다. 기존 열을 덮어쓰지 않습니다.' }
                $reference = $sheet.Cells.Item(1,7)
                $cell.Style = $reference.Style
                $cell.Font.Name = $reference.Font.Name
                $cell.Font.Size = $reference.Font.Size
                $cell.Font.Color = $reference.Font.Color
                $cell.Interior.Color = $reference.Interior.Color
                foreach ($edge in @(7,8,9,10)) {
                    $cell.Borders.Item($edge).LineStyle = $reference.Borders.Item($edge).LineStyle
                    $cell.Borders.Item($edge).Weight = $reference.Borders.Item($edge).Weight
                    $cell.Borders.Item($edge).Color = $reference.Borders.Item($edge).Color
                }
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($reference)
                $cell.Value2 = $extra[$c]
                $cell.Font.Bold = $true
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cell)
            }
            $sheet.Columns.Item(7).ColumnWidth = 78
            $sheet.Columns.Item(8).ColumnWidth = 28
            $sheet.Columns.Item(11).ColumnWidth = 18
            foreach ($c in @(9,10,12,13)) { if ($sheet.Columns.Item($c).ColumnWidth -lt 12) { $sheet.Columns.Item($c).ColumnWidth = 12 } }
            foreach ($r in $group.Group) {
                $g = $sheet.Cells.Item([int]$r.row,7)
                if (([string]$g.Value2) -cne ([string]$r.g_text)) { $g.Value2 = [string]$r.g_text } # 개인정보 익명화
                $g.WrapText = $true
                $g.VerticalAlignment = -4160 # top
                if ($r.ae -eq 'O') { $g.Interior.Color = 65535 } # RGB yellow
                $h = $sheet.Cells.Item([int]$r.row,8)
                $h.NumberFormat = '@'
                if ($r.ae -eq 'O') { $h.Value2='O — '+([string]$r.summary).Trim() } else { $h.Value2='X' }
                $h.WrapText=$true; $h.VerticalAlignment=-4160
                $k = $sheet.Cells.Item([int]$r.row,11)
                $k.Value2=[string]$r.relatedness
                $k.WrapText=$true; $k.VerticalAlignment=-4160
                $sheet.Rows.Item([int]$r.row).AutoFit() | Out-Null
                foreach ($cell in @($g,$h,$k)) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cell) }
            }
            foreach ($width in @(100,120,150)) {
                $capped=@($group.Group | Where-Object { $sheet.Rows.Item([int]$_.row).RowHeight -ge 409.5 })
                if (-not $capped.Count) { break }
                $sheet.Columns.Item(7).ColumnWidth=$width
                foreach ($r in $group.Group) { $sheet.Rows.Item([int]$r.row).AutoFit() | Out-Null }
            }
            $book.Windows.Item(1).FreezePanes=$false
            $book.Windows.Item(1).SplitColumn=0; $book.Windows.Item(1).SplitRow=1
            $book.Windows.Item(1).FreezePanes=$true
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)
        }
        $excel.CalculateFull()
        if ([IO.Path]::GetExtension($target) -eq '.xls') { $format=56 } else { $format=51 }
        $book.SaveAs($target,$format)
        if ($PreviewPath) { $book.SaveCopyAs([IO.Path]::GetFullPath($PreviewPath)) }
        $book.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book); $book=$null
        $book=$excel.Workbooks.Open($target,0,$true)
        $after=Get-Snapshot $book
        if ($snapshot.Count -ne $after.Count) { throw '저장 후 시트 수 검증 실패' }
        for ($si=0;$si -lt $snapshot.Count;$si++) {
            $b=$snapshot[$si]; $a=$after[$si]
            if ($a.name -ne $b.name) { throw '저장 후 시트 이름 검증 실패' }
            $changed=@($payload.records | Where-Object { $_.sheet -eq $b.name } | ForEach-Object { [int]$_.row-1 })
            for ($row=0;$row -lt $b.count;$row++) { for ($col=0;$col -lt $b.columns;$col++) {
                $edited=$changed.Count -gt 0 -and (($row -eq 0 -and $col -ge 7 -and $col -le 12) -or ($changed -contains $row -and $col -in @(6,7,10)))
                if ($edited) { continue }
                if ($b.rows[$row][$col] -cne $a.rows[$row][$col] -or $b.formulas[$row][$col] -cne $a.formulas[$row][$col]) { throw "원본 값/수식 보존 검증 실패: $($b.name) 행 $($row+1) 열 $($col+1)" }
            } }
        }
        $details=@()
        foreach ($group in $groups) {
            $sheet=$book.Worksheets.Item($group.Name); $sheet.Activate()
            if (-not $book.Windows.Item(1).FreezePanes -or $book.Windows.Item(1).SplitRow -ne 1 -or $book.Windows.Item(1).SplitColumn -ne 0) { throw '첫 행 고정 저장 검증 실패' }
            for ($c=0;$c -lt 6;$c++) { if ($sheet.Cells.Item(1,$c+8).Value2 -cne $extra[$c]) { throw '추가 제목 검증 실패' } }
            foreach ($r in $group.Group) {
                $g=$sheet.Cells.Item([int]$r.row,7); $h=$sheet.Cells.Item([int]$r.row,8); $k=$sheet.Cells.Item([int]$r.row,11)
                $expected='X'; if ($r.ae -eq 'O') { $expected='O — '+([string]$r.summary).Trim() }
                if ($h.Value2 -cne $expected -or ([string]$k.Value2) -cne ([string]$r.relatedness) -or -not $g.WrapText) { throw "AE/인과관계/줄 바꿈 검증 실패: $($r.row)행" }
                if (([string]$g.Value2) -cne ([string]$r.g_text)) { throw "G열 익명화 검증 실패: $($r.row)행" }
                if ($r.ae -eq 'O' -and $g.Interior.Color -ne 65535) { throw "노란색 검증 실패: $($r.row)행" }
                $details += @{sheet=$group.Name;row=[int]$r.row;h=$h.Value2;k=[string]$k.Value2}
                foreach ($cell in @($g,$h,$k)) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cell) }
            }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)
        }
        if ((Get-Hash $inputFile) -ne $payload.source_hash) { throw '원본 해시 검증 실패' }
        Write-Json @{output=$target;verified=$true;rows=$details.Count;details=$details} $JsonOut
    }
} catch {
    Write-Json @{error=$_.Exception.Message} $JsonOut
    throw
} finally {
    if ($book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
