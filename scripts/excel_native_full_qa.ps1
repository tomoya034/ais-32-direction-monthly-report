<# Read-only native QA. SpecialCells errors never mean zero formula errors.
PNG requires actual decoding. This does not prove visual/reference equality. #>
[CmdletBinding(DefaultParameterSetName = 'Audit')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Audit')][string]$Root,
    [Parameter(Mandatory = $true, ParameterSetName = 'Audit')][string]$OutputDirectory,
    [Parameter(ParameterSetName = 'Audit')][ValidateRange(-1, 10000)][int]$ExpectedWorkbookCount = -1,
    [Parameter(ParameterSetName = 'Audit')][ValidateRange(-1, 10000)][int]$ExpectedChartCountPerWorkbook = -1,
    [Parameter(ParameterSetName = 'Audit')][string]$ModernContractReport,
    [Parameter(ParameterSetName = 'Audit')][string]$HistoricalContractReport,
    [Parameter(ParameterSetName = 'Audit')][switch]$PlanOnly,
    [Parameter(Mandatory = $true, ParameterSetName = 'Helpers')][switch]$SelfTest
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$nativeSourceHashes = [ordered]@{
    full_qa = (Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash
    helpers = (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'excel_native_audit_helpers.ps1') -Algorithm SHA256).Hash
    historical_helpers = (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'excel_native_historical_audit_helpers.ps1') -Algorithm SHA256).Hash
}
. (Join-Path $PSScriptRoot 'excel_native_audit_helpers.ps1')
. (Join-Path $PSScriptRoot 'excel_native_historical_audit_helpers.ps1')
if ($SelfTest) {
    Assert-NativeAuditExpectedCount 6 6 'workbooks'
    Assert-NativeAuditExpectedCount 2 2 'charts'
    Assert-NativeAuditExpectedCount 0 -1 'generic inventory'
    foreach ($case in @(@(5,6), @(7,6), @(0,2), @(1,2), @(3,2))) {
        $blocked = $false
        try { Assert-NativeAuditExpectedCount $case[0] $case[1] 'synthetic missing/duplicate output' } catch { $blocked = $true }
        if (-not $blocked) { throw 'Missing/duplicate workbook/chart count falsely passed' }
    }
    if ((Assert-NativeAuditCount 0 10) -ne 0 -or (Assert-NativeAuditCount 2.0 2) -ne 2) { throw 'Count success test' }
    foreach ($value in @($true, -1, 1.5, [double]::NaN, 11, '0', [Runtime.InteropServices.ErrorWrapper]::new(-2146826281))) {
        $blocked = $false
        try { Assert-NativeAuditCount $value 10 | Out-Null } catch { $blocked = $true }
        if (-not $blocked) { throw "Invalid count accepted: $value" }
    }
    $protected = Join-Path ([IO.Path]::GetTempPath()) 'ais-nativeqa-protected'
    $blocked = $false
    try { Assert-NativeAuditOutput (Join-Path $protected 'new') @($protected) | Out-Null } catch { $blocked = $true }
    if (-not $blocked -or (Test-Path -LiteralPath (Join-Path $protected 'new'))) { throw 'Output guard failed before mkdir' }
    $empty = [IO.Path]::GetTempFileName()
    try { if ((Get-NativeAuditPng $empty).status -ne 'UNVERIFIED') { throw 'Empty PNG false success' } }
    finally { Remove-Item -LiteralPath $empty }
    $fakeUsed = [pscustomobject]@{ HasFormula = [DBNull]::Value }
    $fakeUsed | Add-Member ScriptMethod SpecialCells { param($kind) throw [Runtime.InteropServices.COMException]::new('synthetic SpecialCells failure', -2146827284) }
    $fakeSheet = [pscustomobject]@{ UsedRange = $fakeUsed }
    $missing = Get-NativeAuditFormulaErrors $fakeSheet
    if ($missing.status -ne 'UNVERIFIED' -or $missing.has_formula_variant_type -ne 'System.DBNull') { throw 'Mixed-formula COM failure became no-errors' }
    $fakeArea = [pscustomobject]@{ HasFormula = $true; CountLarge = 3 }
    $fakeArea | Add-Member ScriptMethod Address { param($a,$b,$c,$d) return '$B$1:$B$3' }
    $fakeAreas = [pscustomobject]@{ Count = 1; Area = $fakeArea }
    $fakeAreas | Add-Member ScriptMethod Item { param($i) return $this.Area }
    $fakeUsed = [pscustomobject]@{ HasFormula = [DBNull]::Value; Scope = [pscustomobject]@{ Areas = $fakeAreas } }
    $fakeUsed | Add-Member ScriptMethod SpecialCells { param($kind) return $this.Scope }
    $fakeSheet = [pscustomobject]@{ UsedRange = $fakeUsed }
    $fakeSheet | Add-Member ScriptMethod Evaluate { param($formula) return 1.0 }
    $positive = Get-NativeAuditFormulaErrors $fakeSheet
    if ($positive.status -ne 'FORMULA_ERRORS' -or $positive.error_count -ne 1 -or $positive.formula_cell_count -ne 3) { throw 'Formula-only independent error count failed' }
    $fakeProtectedUsed = [pscustomobject]@{ HasFormula=[DBNull]::Value; CountLarge=12 }
    $fakeProtectedUsed | Add-Member ScriptMethod Address { param($a,$b,$c,$d) return '$A$1:$C$4' }
    $fakeProtectedUsed | Add-Member ScriptMethod SpecialCells { param($kind) throw 'Must not select protected cells' }
    $fakeProtected = [pscustomobject]@{ UsedRange=$fakeProtectedUsed; ProtectContents=$true; FormulaCount=3; FormulaErrors=1 }
    $fakeProtected | Add-Member ScriptMethod Evaluate { param($formula) if($formula.Contains('ISERROR')) {return $this.FormulaErrors}; return $this.FormulaCount }
    $protectedResult = Get-NativeAuditFormulaErrors $fakeProtected
    if ($protectedResult.status -ne 'FORMULA_ERRORS' -or $protectedResult.error_count -ne 1 -or $protectedResult.formula_cell_count -ne 3) { throw 'Protected mixed range formula masking failed' }
    foreach ($invalidCount in @($true, -1, 0, 13, [double]::NaN, '0')) {
        $fakeProtected.FormulaCount = $invalidCount
        if ((Get-NativeAuditFormulaErrors $fakeProtected).status -ne 'UNVERIFIED') { throw 'Protected invalid/zero formula count falsely accepted' }
    }
    $fakeProtected.FormulaCount=3; $fakeProtected.FormulaErrors=4
    if ((Get-NativeAuditFormulaErrors $fakeProtected).status -ne 'UNVERIFIED') { throw 'Protected invalid error count falsely accepted' }
    Write-Host 'SELFTEST PASS: invalid counts rejected, path guard before mkdir, empty PNG rejected.'
    exit 0
}
$resolvedRoot = Resolve-NativeAuditPath $Root
if (-not (Test-Path -LiteralPath $resolvedRoot -PathType Container)) { throw "Missing Root: $resolvedRoot" }
$resolvedOutput = Assert-NativeAuditOutput $OutputDirectory @($resolvedRoot)
$files = @(Get-ChildItem -LiteralPath $resolvedRoot -Recurse -File -Filter '*.xlsx' | Sort-Object FullName)
if ($files.Count -eq 0) { throw 'No XLSX workbooks in input tree' }
if (@($files | Where-Object Name -like '~$*').Count -gt 0) { throw 'Open Excel temporary workbook present' }
Assert-NativeAuditExpectedCount $files.Count $ExpectedWorkbookCount 'input workbooks before mkdir/COM'
if ($PlanOnly) { [pscustomobject]@{ workbooks = $files.Count; root = $resolvedRoot; output = $resolvedOutput } | ConvertTo-Json; exit 0 }
New-Item -ItemType Directory -Path $resolvedOutput | Out-Null
$imageDirectory = Join-Path $resolvedOutput 'charts'
New-Item -ItemType Directory -Path $imageDirectory | Out-Null
$reportPath = Join-Path $resolvedOutput 'excel_native_qa.json'
$excel = $null; $excelBooks = $null
$workbooks = [Collections.Generic.List[object]]::new()
$started = Get-Date
$report = [ordered]@{
    schema = 'AIS_NATIVE_FULL_QA_3'; generated_at = $started.ToString('o'); root = $resolvedRoot; owner_powershell_pid = $PID
    expected_workbook_count = $files.Count; workbook_count = 0; pass_count = 0; fail_count = 0; workbooks = $workbooks; state = 'IN_PROGRESS'
    asserted_workbook_count = $ExpectedWorkbookCount; asserted_chart_count_per_workbook = $ExpectedChartCountPerWorkbook
    formula_contract = 'FullRebuild then independent all-formula ISERROR count retained; intentional Modern chart NA requires current hash-bound independent matrix, exact helper scope/formula/source/nativeCOUNT0/ISNA and every actual Series.Values missing=null, numeric0=0. All other errors FAIL; diagnostic failure UNVERIFIED.'
    modern_contract_report = $ModernContractReport
    historical_contract_report = $HistoricalContractReport
    verifier_source_sha256 = $nativeSourceHashes
    png_contract = 'Export=true plus signature+actualdecode+positive dimensions; failed first attempt preserved; optional private Activate/Refresh retry.'
    visual_reference_equality = 'UNVERIFIED: native formula/chart access and export is not a visual/reference equality verdict'
}
Write-NativeAuditReport $reportPath $report
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false; $excel.AutomationSecurity = 3
    $excelBooks = $excel.Workbooks
    $report.excel_version = [string]$excel.Version; $report.excel_build = [string]$excel.Build
    $report.excel_process = Get-NativeAuditExcelIdentity $excel
    Write-NativeAuditReport $reportPath $report
    $position = 0
    foreach ($file in $files) {
        $position++; Write-Host ("[{0}/{1}] {2}" -f $position, $files.Count, $file.FullName)
        $before = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        $book = $null; $sheets = $null; $chartSheets = $null
        $entry = [ordered]@{ path = $file.FullName; sha256_before = $before; opened_read_only = $false; sheets = @(); formula_errors = @(); formula_checks = @(); charts = @(); modern_chart_missing = $null; historical_chart_missing = $null; historical_chart_stages = @(); status = 'FAIL'; error = $null }
        try {
            $modernContract = Get-NativeAuditModernContract $ModernContractReport (Resolve-NativeAuditPath $file.FullName) $before
            $historicalContract = Get-NativeAuditHistoricalContract $HistoricalContractReport (Resolve-NativeAuditPath $file.FullName) $before
            if($null -ne $historicalContract -and $null -ne $modernContract){throw 'Workbook cannot have both Modern and Historical native contracts'}
            $book = $excelBooks.Open($file.FullName, 0, $true)
            $entry.opened_read_only = [bool]$book.ReadOnly
            if (-not $entry.opened_read_only) { throw 'Excel did not open read-only' }
            $excel.Calculation = -4135; $excel.CalculateFullRebuild()
            if ([int]$excel.CalculationState -ne 0) { throw 'UNVERIFIED calculation not complete' }
            if($null -ne $historicalContract){
                $entry.historical_chart_missing=Get-NativeAuditHistoricalMissing $book $historicalContract
                if($entry.historical_chart_missing.status -ne 'PASS'){throw "Historical native all-source/point/geometry gate failed: $($entry.historical_chart_missing.error)"}
                $entry.historical_chart_stages += [pscustomobject]@{stage='AFTER_FULL_REBUILD_BEFORE_EXPORT';gate=$entry.historical_chart_missing}
            }
            $sheets = $book.Worksheets; $chartSheets = $book.Charts
            if ($chartSheets.Count -gt 0) { throw 'UNVERIFIED standalone chart-sheet traversal not implemented' }
            for ($sheetIndex = 1; $sheetIndex -le $sheets.Count; $sheetIndex++) {
                $sheet = $null; $objects = $null
                try {
                    $sheet = $sheets.Item($sheetIndex); $entry.sheets += [string]$sheet.Name
                    $missingGate = $null
                    if($null -ne $historicalContract -and $sheet.Name -ceq '總表'){$missingGate=$entry.historical_chart_missing}
                    if ($null -ne $modernContract -and $sheet.Name -ceq '總表') {
                        if ($null -ne $entry.modern_chart_missing) { throw 'Duplicate Modern summary native gate' }
                        $missingGate = Get-NativeAuditModernMissing $sheet $modernContract
                        $entry.modern_chart_missing = $missingGate
                        if ($missingGate.status -ne 'PASS') { throw "Modern native missing-data gate failed: $($missingGate.error)" }
                    }
                    $check = Get-NativeAuditFormulaDisposition (Get-NativeAuditFormulaErrors $sheet) $missingGate
                    $entry.formula_checks += [pscustomobject]@{ sheet = [string]$sheet.Name; result = $check }
                    if ($check.status -eq 'UNVERIFIED') { throw "UNVERIFIED formula scan $($sheet.Name): $($check.error)" }
                    if ($check.unexpected_error_count -gt 0) {
                        $entry.formula_errors += [pscustomobject]@{ sheet=[string]$sheet.Name; unexpected_error_count=$check.unexpected_error_count; raw_error_count=$check.error_count; areas=$check.areas }
                    }
                    $objects = $sheet.ChartObjects()
                    for ($chartIndex = 1; $chartIndex -le $objects.Count; $chartIndex++) {
                        $chartObject = $null; $chart = $null; $plot = $null; $series = $null; $title = $null
                        try {
                            $chartObject = $objects.Item($chartIndex); $chart = $chartObject.Chart
                            $identity = "$($file.FullName)`n$sheetIndex`n$chartIndex"
                            $token = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($identity))).ToLowerInvariant()
                            $png = Join-Path $imageDirectory ($token + '.png')
                            $export = Export-NativeAuditChart $chart $chartObject $sheet $png
                            $plot = $chart.PlotArea; $series = $chart.FullSeriesCollection(); $titleText = ''
                            if ($chart.HasTitle) { $title = $chart.ChartTitle; $titleText = [string]$title.Text }
                            $entry.charts += [pscustomobject]@{
                                sheet = [string]$sheet.Name; index = $chartIndex; chart_type = [int]$chart.ChartType; chart_style = [int]$chart.ChartStyle; title = $titleText
                                left_points = [double]$chartObject.Left; top_points = [double]$chartObject.Top; width_points = [double]$chartObject.Width; height_points = [double]$chartObject.Height
                                plot_left_points = [double]$plot.InsideLeft; plot_top_points = [double]$plot.InsideTop; plot_width_points = [double]$plot.InsideWidth; plot_height_points = [double]$plot.InsideHeight
                                series_count = [int]$series.Count; png_exported = $export.png_exported; png = if ($export.png_exported) { $export.selected.path } else { $png }
                                png_validation = $export.selected; png_export_attempts = $export.attempts
                            }
                        } finally { foreach ($child in @($title, $series, $plot, $chart, $chartObject)) { Release-NativeAuditCom $child } }
                    }
                } finally { Release-NativeAuditCom $objects; Release-NativeAuditCom $sheet }
            }
            Assert-NativeAuditExpectedCount $entry.charts.Count $ExpectedChartCountPerWorkbook $file.FullName
            if ($null -ne $modernContract -and $null -eq $entry.modern_chart_missing) { throw 'Independent Modern contract never received exact native summary/series coverage' }
            if($null -ne $historicalContract -and $null -eq $entry.historical_chart_missing){throw 'Independent Historical contract never received exact native source/series/geometry coverage'}
            if($null -ne $historicalContract){
                $postExport=Get-NativeAuditHistoricalMissing $book $historicalContract
                $entry.historical_chart_stages += [pscustomobject]@{stage='AFTER_NATIVE_EXPORT';gate=$postExport}
                if($postExport.status -ne 'PASS'){throw "Historical post-export native point gate failed: $($postExport.error)"}
                $summary=$null;$historicalObjects=$null
                try{
                    $summary=$sheets.Item('總表');$historicalObjects=$summary.ChartObjects()
                    for($position=1;$position -le $historicalObjects.Count;$position++){
                        $historicalObject=$null;$historicalChart=$null
                        try{$historicalObject=$historicalObjects.Item($position);$historicalChart=$historicalObject.Chart;$summary.Activate()|Out-Null;$historicalObject.Activate()|Out-Null;$historicalChart.Refresh()}
                        finally{Release-NativeAuditCom $historicalChart;Release-NativeAuditCom $historicalObject}
                    }
                }finally{Release-NativeAuditCom $historicalObjects;Release-NativeAuditCom $summary}
                $postRefresh=Get-NativeAuditHistoricalMissing $book $historicalContract
                $entry.historical_chart_stages += [pscustomobject]@{stage='AFTER_EXPLICIT_REFRESH';gate=$postRefresh}
                if($postRefresh.status -ne 'PASS'){throw "Historical post-refresh native point gate failed: $($postRefresh.error)"}
            }
            $entry.status = if ($entry.formula_errors.Count -eq 0 -and @($entry.charts | Where-Object { -not $_.png_exported }).Count -eq 0) { 'PASS' } else { 'FAIL' }
        } catch { $entry.error = $_.Exception.Message; $entry.status = 'FAIL'; $entry.hresult = $_.Exception.HResult }
        finally {
            Release-NativeAuditCom $chartSheets; Release-NativeAuditCom $sheets
            if ($null -ne $book) {
                try { $book.Close($false) } catch { $entry.status = 'FAIL'; $entry.error = 'Close without saving failed: ' + $_.Exception.Message }
                Release-NativeAuditCom $book
            }
        }
        $entry.sha256_after = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        $entry.file_unchanged = $before -eq $entry.sha256_after
        if (-not $entry.file_unchanged) { $entry.status = 'FAIL'; $entry.error = 'Workbook hash changed.' }
        $workbooks.Add([pscustomobject]$entry); $report.workbook_count = $workbooks.Count
        Write-NativeAuditReport $reportPath $report
    }
    if(-not [string]::IsNullOrWhiteSpace($HistoricalContractReport)){
        $historicalEntries=@($workbooks|Where-Object {$null -ne $_.historical_chart_missing})
        if($historicalEntries.Count -ne 5 -or @($historicalEntries.historical_chart_missing.role|Sort-Object -Unique).Count -ne 5){throw 'Historical native mandatory five-role contract census failed'}
    }
    $report.state = 'COMPLETE'
} catch { $report.state = 'INTERRUPTED_OR_FAILED'; $report.fatal_error = $_.Exception.Message }
finally {
    Release-NativeAuditCom $excelBooks
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { $report.cleanup_error = $_.Exception.Message; $report.state = 'INTERRUPTED_OR_FAILED' }
        Release-NativeAuditCom $excel
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    $report.workbook_count = $workbooks.Count; $report.pass_count = @($workbooks | Where-Object status -eq 'PASS').Count
    $report.fail_count = @($workbooks | Where-Object status -ne 'PASS').Count
    $report.elapsed_seconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 3)
    Write-NativeAuditReport $reportPath $report
}
[pscustomobject]$report | Select-Object workbook_count, pass_count, fail_count, state, excel_version, excel_build, elapsed_seconds | Format-List
if ($report.state -ne 'COMPLETE' -or $report.workbook_count -ne $files.Count -or $report.fail_count -ne 0) { exit 1 }
exit 0
