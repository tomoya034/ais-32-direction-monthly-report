<# Read-only observation, NOT a Historical correctness verdict. No workbook is saved.
   Capture actual series and every point before/after native export/refresh because
   Series.Values alone does not establish whether Excel rendered a blank as zero. #>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$Root,
    [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [Parameter(Mandatory=$true)][string]$PauseFlag,
    [string[]]$AdditionalRoot=@(),
    [int]$ExpectedWorkbookCount=5
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'excel_native_audit_helpers.ps1')
$sourceHash=(Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash
$helperHash=(Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'excel_native_audit_helpers.ps1') -Algorithm SHA256).Hash
$inputRoot=Resolve-NativeAuditPath $Root
$inputRoots=@($inputRoot)+@($AdditionalRoot|ForEach-Object {Resolve-NativeAuditPath $_})
$output=Assert-NativeAuditOutput $OutputDirectory $inputRoots
$files=@($inputRoots|ForEach-Object {Get-ChildItem -LiteralPath $_ -File -Filter '*.xlsx'}|Sort-Object FullName)
Assert-NativeAuditExpectedCount $files.Count $ExpectedWorkbookCount 'Historical observation inputs'
if(@($files|Where-Object Name -like '~$*').Count -gt 0){throw 'Temporary/open Excel file is not an input'}
if(Test-Path -LiteralPath $PauseFlag){throw 'Quota pause flag present; no native work can start'}
New-Item -ItemType Directory -Path $output|Out-Null
$owner=Get-CimInstance Win32_Process -Filter ('ProcessId='+$PID)
$reportPath=Join-Path $output 'historical_point_observation.json'
$report=[ordered]@{schema='AIS_NATIVE_HISTORICAL_POINT_OBSERVATION_1';state='IN_PROGRESS';
    boundary='Observation only; no PASS result, no replacement for independently reconstructed period matrix or accepted-reference gate';
    roots=$inputRoots;owner=[pscustomobject]@{pid=$PID;creation=$owner.CreationDate.ToString('o');command=$owner.CommandLine};
    probe_sha256=$sourceHash;helper_sha256=$helperHash;workbooks=@();started_at=(Get-Date).ToString('o')}
$excel=$null;$books=$null
function Assert-ProbeActive {
    if(Test-Path -LiteralPath $PauseFlag){throw 'QUOTA_PAUSED'}
    if((Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash -cne $sourceHash){throw 'Probe source changed during observation'}
}
function Get-PointObservation([object]$Sheet,[object]$ChartObject,[int]$ChartIndex,[string]$Stage){
    $chart=$null;$collection=$null;$plot=$null
    $result=[ordered]@{stage=$Stage;chart_index=$ChartIndex;series=@();geometry_errors=@()}
    try{
        $chart=$ChartObject.Chart;$collection=$chart.FullSeriesCollection();$plot=$chart.PlotArea
        $result.chart_type=[int]$chart.ChartType;$result.display_blanks_as=[int]$chart.DisplayBlanksAs
        $result.plot_visible_only=[bool]$chart.PlotVisibleOnly
        $result.plot=[pscustomobject]@{left=[double]$plot.InsideLeft;top=[double]$plot.InsideTop;width=[double]$plot.InsideWidth;height=[double]$plot.InsideHeight}
        for($seriesIndex=1;$seriesIndex -le $collection.Count;$seriesIndex++){
            Assert-ProbeActive
            $series=$null;$points=$null
            try{
                $series=$collection.Item($seriesIndex);$points=$series.Points()
                $values=@($series.Values);$categories=@($series.XValues)
                $entry=[ordered]@{index=$seriesIndex;name=[string]$series.Name;formula=[string]$series.Formula;chart_type=[int]$series.ChartType;
                    values=$values;categories=$categories;point_count=[int]$points.Count;points=@()}
                for($pointIndex=1;$pointIndex -le $points.Count;$pointIndex++){
                    $point=$null;$geometry=[ordered]@{index=$pointIndex;error=$null}
                    try{
                        $point=$points.Item($pointIndex)
                        foreach($field in @('Left','Top','Width','Height')){
                            try{$geometry[$field.ToLowerInvariant()]=[double]$point.$field}
                            catch{$geometry[$field.ToLowerInvariant()]=$null;$geometry.error=($geometry.error+' '+$field+': '+$_.Exception.Message).Trim()}
                        }
                    }finally{Release-NativeAuditCom $point}
                    $entry.points += [pscustomobject]$geometry
                    if($null -ne $geometry.error){$result.geometry_errors += "series$seriesIndex/point$pointIndex"}
                }
                $result.series += [pscustomobject]$entry
            }finally{Release-NativeAuditCom $points;Release-NativeAuditCom $series}
        }
    }finally{Release-NativeAuditCom $plot;Release-NativeAuditCom $collection;Release-NativeAuditCom $chart}
    return [pscustomobject]$result
}
Write-NativeAuditReport $reportPath $report
try{
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    $books=$excel.Workbooks;$report.excel_process=Get-NativeAuditExcelIdentity $excel
    $report.excel_version=[string]$excel.Version;$report.excel_build=[string]$excel.Build
    Write-NativeAuditReport $reportPath $report
    foreach($file in $files){
        Assert-ProbeActive
        $before=(Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        $entry=[ordered]@{path=$file.FullName;sha256_before=$before;stages=@();status='IN_PROGRESS';error=$null}
        $book=$null;$sheet=$null;$objects=$null;$sheets=$null
        try{
            $book=$books.Open($file.FullName,0,$true)
            if(-not $book.ReadOnly){throw 'Native workbook was not opened read-only'}
            $excel.Calculation=-4135;$excel.CalculateFullRebuild()
            if($excel.CalculationState -ne 0){throw 'Native full rebuild remains incomplete'}
            $sheets=$book.Worksheets;$sheet=$sheets.Item('總表');$objects=$sheet.ChartObjects()
            Assert-NativeAuditExpectedCount $objects.Count 2 'Historical native observation charts'
            for($chartIndex=1;$chartIndex -le $objects.Count;$chartIndex++){
                $object=$null;$chart=$null
                try{
                    $object=$objects.Item($chartIndex);$chart=$object.Chart
                    $entry.stages += Get-PointObservation $sheet $object $chartIndex 'AFTER_FULL_REBUILD_BEFORE_EXPORT'
                    $name=[IO.Path]::GetFileNameWithoutExtension($file.Name)+'_chart'+$chartIndex+'.png'
                    $export=Export-NativeAuditChart $chart $object $sheet (Join-Path $output $name)
                    $entry.stages += Get-PointObservation $sheet $object $chartIndex 'AFTER_NATIVE_EXPORT'
                    $sheet.Activate()|Out-Null;$object.Activate()|Out-Null;$chart.Refresh()
                    $entry.stages += Get-PointObservation $sheet $object $chartIndex 'AFTER_EXPLICIT_REFRESH'
                    $entry['chart'+$chartIndex+'_export']=$export
                }finally{Release-NativeAuditCom $chart;Release-NativeAuditCom $object}
            }
            $entry.status='OBSERVATION_COMPLETE'
        }catch{$entry.status='OBSERVATION_FAILED';$entry.error=$_.Exception.Message}
        finally{
            Release-NativeAuditCom $objects;Release-NativeAuditCom $sheet;Release-NativeAuditCom $sheets
            if($null -ne $book){try{$book.Close($false)}catch{$entry.status='OBSERVATION_FAILED';$entry.error='Read-only close failed: '+$_.Exception.Message};Release-NativeAuditCom $book}
        }
        $entry.sha256_after=(Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        $entry.file_unchanged=$before -ceq $entry.sha256_after
        if(-not $entry.file_unchanged){$entry.status='OBSERVATION_FAILED';$entry.error='Workbook changed during read-only observation'}
        $report.workbooks += [pscustomobject]$entry;Write-NativeAuditReport $reportPath $report
        if($entry.status -ne 'OBSERVATION_COMPLETE'){throw $entry.error}
    }
    $report.state='COMPLETE_OBSERVATION_ONLY'
}catch{$report.state=if($_.Exception.Message -eq 'QUOTA_PAUSED'){'PAUSED'}else{'INTERRUPTED_OR_FAILED'};$report.error=$_.Exception.Message}
finally{
    Release-NativeAuditCom $books
    if($null -ne $excel){try{$excel.Quit()}catch{$report.cleanup_error=$_.Exception.Message};Release-NativeAuditCom $excel}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    $report.finished_at=(Get-Date).ToString('o');Write-NativeAuditReport $reportPath $report
}
if($report.state -ne 'COMPLETE_OBSERVATION_ONLY'){exit 1}
exit 0
