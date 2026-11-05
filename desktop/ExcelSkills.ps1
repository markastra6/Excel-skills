[CmdletBinding()]
param(
    [ValidateSet('Menu','NewTracker','SetTargets','ImportCsv','ExportPdf','Recalculate','Inspect')]
    [string]$Action = 'Menu',
    [string]$Source,
    [string]$Destination,
    [string]$Project = 'New Project',
    [ValidateRange(1,100000)][int]$CompSB = 30,
    [ValidateRange(1,100000)][int]$CompWS = 20,
    [ValidateRange(1,100000)][int]$QCSB = 40,
    [ValidateRange(1,100000)][int]$QCWS = 25,
    [ValidateRange(1,100000)][int]$JA = 1,
    [ValidateRange(1,24)][int]$Hours = 8
)
# Requires Windows desktop Excel. Compatible with Excel 2010 COM APIs.
# Never overwrites originals or destination files. Macro auto-execution disabled.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ask([string]$label, [string]$default = '') {
    if ($default) {
        $answer = Read-Host "$label [$default]"
        if ([string]::IsNullOrWhiteSpace($answer)) { return $default }
        return $answer
    }
    return (Read-Host $label)
}

if ($Action -eq 'Menu') {
    Write-Host ''
    Write-Host 'Excel Skills - Windows Excel 2010+ automation'
    Write-Host '  1. New reusable Production/QC tracker'
    Write-Host '  2. Change shared targets in a tracker copy'
    Write-Host '  3. Import CSV as editable Excel workbook'
    Write-Host '  4. Export Excel workbook to PDF'
    Write-Host '  5. Recalculate workbook copy'
    Write-Host '  6. Inspect workbook'
    $choice = Ask 'Choose 1-6'
    $actionMap = @{'1'='NewTracker';'2'='SetTargets';'3'='ImportCsv';'4'='ExportPdf';'5'='Recalculate';'6'='Inspect'}
    if (-not $actionMap.ContainsKey($choice)) { throw 'Invalid menu choice' }
    $Action = $actionMap[$choice]
    if ($Action -ne 'NewTracker') { $Source = Ask 'Source file (path including extension)' }
    if ($Action -notin @('Inspect')) { $Destination = Ask 'New output file (will not overwrite)' }
    if ($Action -eq 'NewTracker') { $Project = Ask 'Project name' $Project }
    if ($Action -in @('NewTracker','SetTargets')) {
        $CompSB = [int](Ask 'Comp Student Book pages/day' "$CompSB")
        $CompWS = [int](Ask 'Comp Worksheets pages/day' "$CompWS")
        $QCSB = [int](Ask 'QC Student Book pages/day' "$QCSB")
        $QCWS = [int](Ask 'QC Worksheets pages/day' "$QCWS")
        $JA = [int](Ask 'JA units/day' "$JA")
        $Hours = [int](Ask 'Working hours/day' "$Hours")
    }
}

function ExistingPath([string]$candidate, [string]$expected) {
    if (-not $candidate) { throw "Missing input $expected" }
    $full = [IO.Path]::GetFullPath($candidate.Trim('"'))
    if (-not [IO.File]::Exists($full)) { throw "Source not found: $full" }
    if ([IO.Path]::GetExtension($full).ToLowerInvariant() -ne $expected) {
        throw "Source must have extension $expected"
    }
    return $full
}
function NewOutput([string]$candidate, [string]$expected, [string]$input = '') {
    if (-not $candidate) { throw 'Missing destination path' }
    $full = [IO.Path]::GetFullPath($candidate.Trim('"'))
    if ([IO.Path]::GetExtension($full).ToLowerInvariant() -ne $expected) {
        throw "Destination must have extension $expected"
    }
    if ($input -and $full -ieq $input) { throw 'Source and destination must be different' }
    if ([IO.File]::Exists($full)) { throw "Destination already exists: $full" }
    $folder = [IO.Path]::GetDirectoryName($full)
    if (-not [IO.Directory]::Exists($folder)) { [IO.Directory]::CreateDirectory($folder) | Out-Null }
    return $full
}
function StagePath([string]$final) {
    return [IO.Path]::Combine([IO.Path]::GetDirectoryName($final),
        ('.excel-skills-' + [Guid]::NewGuid().ToString('N') + [IO.Path]::GetExtension($final)))
}
function OleColor([int]$r,[int]$g,[int]$b) { return ($r + 256*$g + 65536*$b) }
function SetInputStyle($range) {
    $range.Interior.Color = (OleColor 236 250 246)
    $range.Font.Color = (OleColor 15 92 81)
}
function SetupExcel {
    try { $app = New-Object -ComObject Excel.Application }
    catch { throw 'Microsoft Excel desktop is required. Install/activate Excel on Windows first.' }
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $app.AskToUpdateLinks = $false
    $app.EnableEvents = $false
    try { $app.AutomationSecurity = 3 } catch { }
    return $app
}
function AddValidation($range,[string]$list) {
    try {
        $range.Validation.Delete()
        # xlValidateList = 3, xlValidAlertStop = 1, xlBetween = 1.
        $range.Validation.Add(3,1,1,$list)
        $range.Validation.IgnoreBlank = $true
        $range.Validation.InCellDropdown = $true
    } catch { Write-Warning 'One dropdown could not be created in this Excel installation' }
}
function CreateTracker($book) {
    $sheet = $book.Worksheets.Item(1)
    $sheet.Name = 'Production QC'
    while ($book.Worksheets.Count -gt 1) { $book.Worksheets.Item($book.Worksheets.Count).Delete() }
    $sheet.Activate()
    $sheet.Range('A1:AC40').Font.Name = 'Calibri'
    $sheet.Range('A1:AC40').Font.Size = 9
    $sheet.Range('A1:AC40').RowHeight = 19
    $sheet.Range('A1:AC1').Merge()
    $sheet.Range('A1').Value2 = 'PRODUCTION + QUALITY CONTROL'
    $sheet.Range('A1:AC1').Interior.Color = (OleColor 17 32 51)
    $sheet.Range('A1:AC1').Font.Color = (OleColor 255 255 255)
    $sheet.Range('A1:AC1').Font.Bold = $true
    $sheet.Range('A1:AC1').Font.Size = 15
    $sheet.Range('A1:AC1').RowHeight = 34
    $sheet.Range('A2').Value2 = 'Project'
    $sheet.Range('B2:F2').Merge()
    $sheet.Range('B2').Value2 = $Project
    $sheet.Range('G2').Value2 = 'Edit mint cells only. One set of shared targets below.'
    $sheet.Range('G2:N2').Merge()
    $sheet.Range('A3:N3').Merge()
    $sheet.Range('A3').Value2 = 'SHARED PRODUCTIVITY TARGETS'
    $sheet.Range('A3:N3').Interior.Color = (OleColor 32 76 92)
    $sheet.Range('A3:N3').Font.Color = (OleColor 255 255 255)
    $labels = @('Comp SB/day','Comp WS/day','QC SB/day','QC WS/day','JA units/day','Hours/day')
    $cols = @(1,3,5,7,9,11)
    $vals = @($CompSB,$CompWS,$QCSB,$QCWS,$JA,$Hours)
    for ($i=0; $i -lt $labels.Count; $i++) {
        $sheet.Cells.Item(4,$cols[$i]).Value2 = $labels[$i]
        $cell = $sheet.Cells.Item(4,($cols[$i]+1))
        $cell.Value2 = [double]$vals[$i]
        SetInputStyle $cell
        $cell.Font.Bold = $true
    }
    $sheet.Range('A6:N6').Merge()
    $sheet.Range('A6').Value2 = 'PROJECT OVERVIEW'
    $sheet.Range('A6:N6').Interior.Color = (OleColor 234 240 246)
    $sheet.Range('A6:N6').Font.Bold = $true
    $sheet.Range('A7').Value2 = 'Proof scope'
    $sheet.Range('B7').Formula = '=SUM(C11:C32)'
    $sheet.Range('D7').Value2 = 'Comp left'
    $sheet.Range('E7').Formula = '=SUM(F11:F32)'
    $sheet.Range('G7').Value2 = 'QC left'
    $sheet.Range('H7').Formula = '=SUM(I11:I32)'
    $sheet.Range('J7').Value2 = 'Review'
    $sheet.Range('K7').Formula = '=COUNTIF(M11:M32,"Review")'
    $sheet.Range('A9:N9').Merge()
    $sheet.Range('A9').Value2 = 'CHAPTER / UNIT PROGRESS'
    $sheet.Range('A9:N9').Interior.Color = (OleColor 32 76 92)
    $sheet.Range('A9:N9').Font.Color = (OleColor 255 255 255)
    $headers = @('Unit','Type','Proof pages','Composer','Comp pages','Comp left',
                 'QC employee','QC pages','QC left','Corrections','PE pages',
                 'Quality %','Status','Comp target hours')
    for ($i=0; $i -lt $headers.Count; $i++) { $sheet.Cells.Item(10,$i+1).Value2 = $headers[$i] }
    $sheet.Range('A10:N10').Interior.Color = (OleColor 221 232 238)
    $sheet.Range('A10:N10').Font.Bold = $true
    $sheet.Range('A10:N10').WrapText = $true
    $sheet.Range('A10:N10').RowHeight = 30
    foreach ($column in @('A','B','C','D','E','G','H','J','K')) {
        SetInputStyle $sheet.Range(('{0}11:{0}32' -f $column))
    }
    for ($r = 11; $r -le 32; $r++) {
        $sheet.Range("F$r").Formula = '=IF(C'+$r+'="","",MAX(0,C'+$r+'-E'+$r+'))'
        $sheet.Range("I$r").Formula = '=IF(C'+$r+'="","",MAX(0,C'+$r+'-H'+$r+'))'
        $sheet.Range("L$r").Formula = '=IF(OR(C'+$r+'="",K'+$r+'=""),"",IF(OR(K'+$r+'<0,K'+$r+'>C'+$r+'),"CHECK",1-K'+$r+'/C'+$r+'))'
        $sheet.Range("M$r").Formula = '=IF(C'+$r+'="","",IF(OR(E'+$r+'>C'+$r+',H'+$r+'>C'+$r+'),"Review",IF(AND(F'+$r+'=0,I'+$r+'=0),"Done",IF(OR(E'+$r+'>0,H'+$r+'>0),"In Progress","Not Started"))))'
        $sheet.Range("N$r").Formula = '=IF(OR(C'+$r+'="",B'+$r+'=""),"",IF(B'+$r+'="SB",C'+$r+'/$B$4*$L$4,IF(B'+$r+'="WS",C'+$r+'/$D$4*$L$4,"")))'
    }
    $sheet.Range('L11:L32').NumberFormat = '0%'
    $sheet.Range('N11:N32').NumberFormat = '0.0'
    AddValidation $sheet.Range('B11:B32') 'SB,WS,SA'
    $sheet.Range('P9:AB9').Merge()
    $sheet.Range('P9').Value2 = 'DAY-WISE TEAM OUTPUT'
    $sheet.Range('P9:AB9').Interior.Color = (OleColor 32 76 92)
    $sheet.Range('P9:AB9').Font.Color = (OleColor 255 255 255)
    $sheet.Range('P10').Value2 = 'Employee'
    $sheet.Range('Q10').Value2 = 'Role'
    $sheet.Range('R10').Value2 = 'Type'
    for ($d = 1; $d -le 8; $d++) { $sheet.Cells.Item(10,18+$d).Value2 = "Day $d" }
    $sheet.Range('AA10').Value2 = 'Total'
    $sheet.Range('AB10').Value2 = 'Target %'
    $sheet.Range('P10:AB10').Interior.Color = (OleColor 221 232 238)
    $sheet.Range('P10:AB10').Font.Bold = $true
    SetInputStyle $sheet.Range('P11:Z24')
    AddValidation $sheet.Range('Q11:Q24') 'Comp,QC,JA'
    AddValidation $sheet.Range('R11:R24') 'SB,WS,SA,Mixed'
    for ($r=11; $r -le 24; $r++) {
        $sheet.Range("AA$r").Formula = '=IF(COUNT(S'+$r+':Z'+$r+')=0,"",SUM(S'+$r+':Z'+$r+'))'
        $sheet.Range("AC$r").Formula = '=IF(Q'+$r+'="JA",$J$4,IF(AND(Q'+$r+'="Comp",R'+$r+'="SB"),$B$4,IF(AND(Q'+$r+'="Comp",R'+$r+'="WS"),$D$4,IF(AND(Q'+$r+'="QC",R'+$r+'="SB"),$F$4,IF(AND(Q'+$r+'="QC",R'+$r+'="WS"),$H$4,"")))))'
        $sheet.Range("AB$r").Formula = '=IF(OR(AC'+$r+'="",COUNT(S'+$r+':Z'+$r+')=0),"",AA'+$r+'/(COUNT(S'+$r+':Z'+$r+')*AC'+$r+'))'
    }
    $sheet.Range('AB11:AB24').NumberFormat='0%'
    $sheet.Range('P26:AB26').Merge()
    $sheet.Range('P26').Value2='SECTION COVERAGE - PLAN, COMP, QC'
    $sheet.Range('P26:AB26').Interior.Color = (OleColor 234 240 246)
    $sheet.Range('P26:AB26').Font.Bold=$true
    $sheet.Range('AE10:AH10').Value2 = @('Section','Planned','Comp','QC')
    for($i=11;$i -le 13;$i++){
        $t = @('SB','WS','SA')[$i-11]
        $sheet.Range("AE$i").Value2=$t
        $sheet.Range("AF$i").Formula='=SUMIF($B$11:$B$32,AE'+$i+',$C$11:$C$32)'
        $sheet.Range("AG$i").Formula='=SUMIF($B$11:$B$32,AE'+$i+',$E$11:$E$32)'
        $sheet.Range("AH$i").Formula='=SUMIF($B$11:$B$32,AE'+$i+',$H$11:$H$32)'
    }
    $chartObj = $sheet.ChartObjects().Add(
        [double]$sheet.Range('P27').Left, [double]$sheet.Range('P27').Top, 530.0, 205.0)
    $chartObj.Chart.ChartType = 51
    $chartObj.Chart.SetSourceData($sheet.Range('AE10:AH13'))
    $chartObj.Chart.PlotVisibleOnly = $false
    $chartObj.Chart.HasLegend = $true
    $sheet.Range('AC:AH').EntireColumn.Hidden = $true
    $sizes = @{'A'=17;'B'=9;'C'=11;'D'=14;'E'=11;'F'=10;'G'=14;'H'=11;
               'I'=10;'J'=11;'K'=10;'L'=11;'M'=14;'N'=14;'O'=3;'P'=15;
               'Q'=9;'R'=9;'S'=8;'T'=8;'U'=8;'V'=8;'W'=8;'X'=8;
               'Y'=8;'Z'=8;'AA'=10;'AB'=11}
    foreach ($column in $sizes.Keys) { $sheet.Columns.Item($column).ColumnWidth=$sizes[$column] }
    $sheet.PageSetup.Orientation=2
    $sheet.PageSetup.PrintArea='$A$1:$AB$39'
    $sheet.PageSetup.Zoom=$false
    $sheet.PageSetup.FitToPagesWide=1
    $sheet.PageSetup.FitToPagesTall=1
    $sheet.Range('A1').Select() | Out-Null
    $sheet.Application.ActiveWindow.DisplayGridlines=$false
    $sheet.Application.ActiveWindow.Zoom=85
    $sheet.Application.ActiveWindow.FreezePanes=$false
    $sheet.Range('A11').Select() | Out-Null
    $sheet.Application.ActiveWindow.FreezePanes=$true
    $sheet.Range('A1').Select() | Out-Null
}
function CopyToStage([string]$input,[string]$final) {
    $stage=StagePath $final
    Copy-Item -LiteralPath $input -Destination $stage -ErrorAction Stop
    return $stage
}
function CsvIntoWorkbook($book,[string]$csvPath) {
    Add-Type -AssemblyName Microsoft.VisualBasic
    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser(
       $csvPath, [System.Text.Encoding]::UTF8)
    try {
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters(',')
        $parser.HasFieldsEnclosedInQuotes = $true
        $rows = New-Object System.Collections.ArrayList
        $maxCols=0
        while (-not $parser.EndOfData) {
            if ($rows.Count -ge 100001) { throw 'CSV exceeds supported import limit of 100,000 data rows' }
            $fields=$parser.ReadFields()
            if ($fields.Length -gt 200) { throw 'CSV has over 200 columns' }
            if ($fields.Length -gt $maxCols) { $maxCols=$fields.Length }
            [void]$rows.Add($fields)
        }
    } finally { $parser.Close() }
    if ($rows.Count -eq 0) { throw 'CSV is empty' }
    $ws=$book.Worksheets.Item(1)
    $ws.Name='Imported Data'
    $matrix=New-Object 'object[,]' $rows.Count,$maxCols
    for ($i=0; $i -lt $rows.Count; $i++){
        for ($j=0; $j -lt $maxCols; $j++){
            if ($j -lt $rows[$i].Length) { $matrix[$i,$j]=[string]$rows[$i][$j] }
            else { $matrix[$i,$j]='' }
        }
    }
    $dest=$ws.Range('A1').Resize($rows.Count,$maxCols)
    $dest.NumberFormat='@'
    $dest.Value2=$matrix
    $ws.Range('A1').Resize(1,$maxCols).Font.Bold=$true
    $ws.Range('A1').Resize(1,$maxCols).Interior.Color=(OleColor 32 76 92)
    $ws.Range('A1').Resize(1,$maxCols).Font.Color=(OleColor 255 255 255)
    $ws.Range('A1').Resize(1,$maxCols).AutoFilter() | Out-Null
    $ws.Columns.AutoFit() | Out-Null
    while ($book.Worksheets.Count -gt 1) { $book.Worksheets.Item($book.Worksheets.Count).Delete() }
}
$excel=$null; $workbook=$null; $stage=$null
try {
    if ($Action -eq 'NewTracker') {
        $target=NewOutput $Destination '.xlsx'
        $stage=StagePath $target
        $excel=SetupExcel
        $workbook=$excel.Workbooks.Add()
        CreateTracker $workbook
        $workbook.SaveAs($stage,51)
        $workbook.Close($false);$workbook=$null
        Move-Item -LiteralPath $stage -Destination $target -ErrorAction Stop
        $stage=$null
        Write-Host "Created: $target"
    } elseif ($Action -eq 'ImportCsv') {
        $input=ExistingPath $Source '.csv'
        $target=NewOutput $Destination '.xlsx' $input
        $stage=StagePath $target
        $excel=SetupExcel
        $workbook=$excel.Workbooks.Add()
        CsvIntoWorkbook $workbook $input
        $workbook.SaveAs($stage,51)
        $workbook.Close($false);$workbook=$null
        Move-Item -LiteralPath $stage -Destination $target -ErrorAction Stop
        $stage=$null
        Write-Host "Created: $target"
    } elseif ($Action -eq 'ExportPdf') {
        $input=ExistingPath $Source '.xlsx'
        $target=NewOutput $Destination '.pdf' $input
        $stage=StagePath $target
        $excel=SetupExcel
        $workbook=$excel.Workbooks.Open($input,0,$true)
        $workbook.ExportAsFixedFormat(0,$stage)
        $workbook.Close($false);$workbook=$null
        Move-Item -LiteralPath $stage -Destination $target -ErrorAction Stop
        $stage=$null
        Write-Host "Exported: $target"
    } elseif ($Action -eq 'Inspect') {
        $input=ExistingPath $Source '.xlsx'
        $excel=SetupExcel
        $workbook=$excel.Workbooks.Open($input,0,$true)
        Write-Host ("Workbook: {0}" -f $input)
        Write-Host ("Worksheets: {0}" -f $workbook.Worksheets.Count)
        foreach($sheet in $workbook.Worksheets){
            Write-Host ("  {0}: used={1}, charts={2}" -f $sheet.Name,$sheet.UsedRange.Address(),$sheet.ChartObjects().Count)
        }
        $workbook.Close($false);$workbook=$null
    } else {
        $input=ExistingPath $Source '.xlsx'
        $target=NewOutput $Destination '.xlsx' $input
        $stage=CopyToStage $input $target
        $excel=SetupExcel
        $workbook=$excel.Workbooks.Open($stage,0,$false)
        if($Action -eq 'SetTargets') {
            $sheet=$workbook.Worksheets.Item('Production QC')
            if([string]$sheet.Range('A4').Value2 -ne 'Comp SB/day' -or
               [string]$sheet.Range('K4').Value2 -ne 'Hours/day'){
                throw 'Not an Excel Skills tracker: shared target header cells are missing'
            }
            $sheet.Range('B4').Value2=[double]$CompSB
            $sheet.Range('D4').Value2=[double]$CompWS
            $sheet.Range('F4').Value2=[double]$QCSB
            $sheet.Range('H4').Value2=[double]$QCWS
            $sheet.Range('J4').Value2=[double]$JA
            $sheet.Range('L4').Value2=[double]$Hours
        }
        $excel.CalculateFullRebuild()
        $workbook.Save()
        $workbook.Close($false);$workbook=$null
        Move-Item -LiteralPath $stage -Destination $target -ErrorAction Stop
        $stage=$null
        Write-Host "Saved: $target"
    }
} catch {
    Write-Error ("Excel Skills failed: " + $_.Exception.Message) -ErrorAction Continue
    exit 1
} finally {
    if ($workbook) { try {$workbook.Close($false)} catch {} }
    if ($excel) { try {$excel.Quit()} catch {} }
    if ($workbook) {try {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)} catch {}}
    if ($excel) {try {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)} catch {}}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if ($stage -and [IO.File]::Exists($stage)) {Remove-Item -LiteralPath $stage -Force -ErrorAction SilentlyContinue}
}
