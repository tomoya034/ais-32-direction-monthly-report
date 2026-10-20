<# Sequential, checkpointed native QA queue. No workbook is saved.
   Only explicitly supplied completed port-months enter this queue. #>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [Parameter(Mandatory=$true)][string]$QueueName,
    [Parameter(Mandatory=$true)][string]$JobsJson,
    [Parameter(Mandatory=$true)][string]$PauseFlag
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'excel_native_audit_helpers.ps1')
$started=Get-Date
$jobs=@($JobsJson | ConvertFrom-Json)
if ($jobs.Count -eq 0) { throw 'No queue jobs supplied' }
$evidence=Resolve-NativeAuditPath $EvidenceRoot
if ($QueueName -notmatch '^[A-Za-z0-9_-]+$') { throw 'Unsafe queue name' }
$reportPath=Join-Path $evidence ($QueueName+'.queue.json')
if(Test-Path -LiteralPath $reportPath){throw 'Queue checkpoint already exists'}
$records=[Collections.Generic.List[object]]::new()
$owner=Get-CimInstance Win32_Process -Filter ('ProcessId='+$PID)
$report=[ordered]@{schema='AIS_NATIVE_QUEUE_1';started_at=$started.ToString('o');owner_pid=$PID;owner_creation=$owner.CreationDate.ToString('o');owner_command=$owner.CommandLine;state='IN_PROGRESS';planned_jobs=$jobs.Count;completed_jobs=0;records=$records;pause_flag=$PauseFlag}
function Test-PauseRequested {
    return (Test-Path -LiteralPath $PauseFlag) -and (Get-Item -LiteralPath $PauseFlag).LastWriteTime -ge $started
}
function Stop-ExactProcess([int]$ProcessId,[string]$Creation,[string]$Command) {
    $current=Get-CimInstance Win32_Process -Filter ('ProcessId='+$ProcessId)
    if($null -eq $current){return}
    if($current.CreationDate.ToString('o') -ne $Creation -or $current.CommandLine -ne $Command){throw "Process identity changed: $ProcessId"}
    Stop-Process -Id $ProcessId -Force
}
function Stop-PrivateExcel([string]$NativeReport) {
    if(-not(Test-Path -LiteralPath $NativeReport)){return}
    # PowerShell 7.5+ auto-parses ISO strings as DateTime, which loses the
    # fractional identity when later coerced to a culture-formatted string.
    # Preserve exact recorded PID start-time text for safe shutdown matching.
    $capture=Get-Content -LiteralPath $NativeReport -Raw|ConvertFrom-Json -DateKind String
    if(-not($capture.PSObject.Properties.Name -contains 'excel_process')){return}
    $private=$capture.excel_process
    $actual=Get-Process -Id $private.pid -ErrorAction SilentlyContinue
    if($null -eq $actual){return}
    if($actual.ProcessName -ne 'EXCEL' -or $actual.StartTime.ToString('o') -ne $private.start_time){throw 'Private Excel identity changed; will not terminate another process'}
    Stop-Process -Id $private.pid -Force
}
Write-NativeAuditReport $reportPath $report
try {
    foreach($job in $jobs){
        if(Test-PauseRequested){$report.state='PAUSED';break}
        if($job.kind -notin @('fullqa','style')){throw 'Unknown native queue job kind'}
        $root=Resolve-NativeAuditPath $job.root
        $output=Assert-NativeAuditOutput $job.output @($root)
        if(-not $output.StartsWith($evidence+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Output is not inside evidence root'}
        $scriptName=if($job.kind -eq 'fullqa'){'excel_native_full_qa.ps1'}else{'excel_native_chart_style_audit.ps1'}
        $nativeReport=Join-Path $output $(if($job.kind -eq 'fullqa'){'excel_native_qa.json'}else{'excel_native_chart_style_audit.json'})
        $log=$output+'.stdout.log';$err=$output+'.stderr.log'
        foreach($path in @($log,$err)){if(Test-Path -LiteralPath $path){throw "Queue evidence log collision: $path"}}
        $launchArguments=@('-NoProfile','-File',('"'+(Join-Path $PSScriptRoot $scriptName)+'"'),'-Root',('"'+$root+'"'),'-OutputDirectory',('"'+$output+'"'),'-ExpectedWorkbookCount','6','-ExpectedChartCountPerWorkbook','2')
        if($job.kind -eq 'fullqa' -and $job.PSObject.Properties.Name -contains 'modern_contract_report'){
            $contractPath=Resolve-NativeAuditPath $job.modern_contract_report
            if(-not(Test-Path -LiteralPath $contractPath -PathType Leaf)){throw 'Missing independent Modern contract report for full QA job'}
            $launchArguments += @('-ModernContractReport',('"'+$contractPath+'"'))
        }
        $pwsh=(Get-Process -Id $PID).Path
        $child=Start-Process -FilePath $pwsh -ArgumentList $launchArguments -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError $err -PassThru
        $identity=Get-CimInstance Win32_Process -Filter ('ProcessId='+$child.Id)
        $record=[ordered]@{kind=$job.kind;root=$root;output=$output;native_report=$nativeReport;pid=$identity.ProcessId;creation=$identity.CreationDate.ToString('o');command=$identity.CommandLine;state='IN_PROGRESS';stdout=$log;stderr=$err}
        $records.Add($record);Write-NativeAuditReport $reportPath $report
        while(-not $child.WaitForExit(1000)){
            if(Test-PauseRequested){
                Stop-ExactProcess $record.pid $record.creation $record.command
                Stop-PrivateExcel $nativeReport
                $record.state='PAUSED';$report.state='PAUSED';break
            }
        }
        if($report.state -eq 'PAUSED'){break}
        $record.exit_code=$child.ExitCode
        $record.state=if($child.ExitCode -eq 0){'COMPLETE'}else{'FAILED'}
        $report.completed_jobs++
        Write-NativeAuditReport $reportPath $report
        # Failures remain evidence, and independent later months are still
        # inspected rather than silently skipped or discarded.
    }
    if($report.state -ne 'PAUSED'){$report.state='COMPLETE'}
} catch {$report.state='INTERRUPTED_OR_FAILED';$report.fatal_error=$_.Exception.Message}
finally {Write-NativeAuditReport $reportPath $report}
if($report.state -ne 'COMPLETE' -or @($records|Where-Object state -ne 'COMPLETE').Count -gt 0){exit 1}
exit 0
