<# Deterministic native-contract negative tests. No Excel/book/raw file is edited. #>
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'excel_native_audit_helpers.ps1')

function New-Fixture {
    $labels=@('北','北微東','東北偏北','東北微北','東北','東北微東','東北偏東','東微北','東','東微南','東南偏東','東南微東','東南','東南微南','東南偏南','南微東','南','南微西','西南偏南','西南微南','西南','西南微西','西南偏西','西微南','西','西微北','西北偏西','西北微西','西北','西北微北','西北偏北','北微西')
    $values=@(for($index=0;$index -lt 32;$index++){if($index -in @(7,31)){$null}elseif($index -eq 0){0.0}else{[double]($index+1)}})
    $snapshots=@{}
    $headers=[Array]::CreateInstance([object],@(1,32),@(1,1));$grid=[Array]::CreateInstance([object],@(1,32),@(1,1))
    for($index=0;$index -lt 32;$index++){$headers[1,($index+1)]=$labels[$index];$grid[1,($index+1)]=if($null -eq $values[$index]){''}else{$values[$index]}}
    foreach($address in @('B4:AG4','AO4:BT4')){$snapshots[$address]=[pscustomobject]@{Value2=$headers}}
    $snapshots['AO:BT']=[pscustomobject]@{EntireColumn=[pscustomobject]@{Hidden=$true}}
    $snapshots['B5:AG5']=[pscustomobject]@{Value2=$grid}
    foreach($pair in @(@('AI4','納入方位'),@('AJ4','MAX'),@('AK4','MIN'),@('AL4','平均'))){$snapshots[$pair[0]]=[pscustomobject]@{Formula=$pair[1];HasFormula=$false;Value2=$pair[1]}}
    $snapshots['A5']=[pscustomobject]@{Formula=0;HasFormula=$false;Value2=[double]([DateTime]'2026-01-01').ToOADate()}
    $helpers=@();$radarHelpers=@();$intentional=@()
    for($direction=0;$direction -lt 32;$direction++){
        $row=$direction+5;$labelCell="AI$row"
        $snapshots[$labelCell]=[pscustomobject]@{Formula=$labels[$direction];HasFormula=$false;Value2=$labels[$direction]}
        $expected=$values[$direction];$column=Get-NativeAuditExcelColumn ($direction+2)
        $dailySource="${column}5";$snapshots[$dailySource]=[pscustomobject]@{Formula='=IF(COUNT(source)=0,"",source)';HasFormula=$true;Value2=$grid[1,($direction+1)]}
        for($stat=0;$stat -lt 3;$stat++){
            $function=@('MAX','MIN','AVERAGE')[$stat];$source="$column$($stat+7)";$sourceRange="${column}5:${column}5"
            $sourceFormula='=IF(COUNT('+$sourceRange+')<1,"",'+$function+'('+$sourceRange+'))'
            $cell="$(Get-NativeAuditExcelColumn (36+$stat))$row";$formula='=IF(COUNT('+$source+')=0,NA(),'+$source+')'
            $snapshots[$source]=[pscustomobject]@{Formula=$sourceFormula;HasFormula=$true;Value2=if($null -eq $expected){''}else{$expected}}
            $snapshots[$cell]=[pscustomobject]@{Formula=$formula;HasFormula=$true;Value2=if($null -eq $expected){-2146826246}else{$expected}}
            $helpers += [pscustomobject]@{verified=$true;sheet='總表';cell=$cell;direction_index=$direction;direction=$labels[$direction];statistic=$function;source=$source;source_range=$sourceRange;source_formula=$sourceFormula;formula=$formula;expected_value=$expected;expected_missing=$null -eq $expected}
            if($null -eq $expected){$intentional+=$cell}
        }
        $cell="$(Get-NativeAuditExcelColumn ($direction+41))5";$formula='=IF(COUNT('+$dailySource+')=0,NA(),'+$dailySource+')'
        $snapshots[$cell]=[pscustomobject]@{Formula=$formula;HasFormula=$true;Value2=if($null -eq $expected){-2146826246}else{$expected}}
        $radarHelpers += [pscustomobject]@{verified=$true;sheet='總表';cell=$cell;source=$dailySource;source_formula=$snapshots[$dailySource].Formula;formula=$formula;direction_index=$direction;direction=$labels[$direction];day='2026-01-01';expected_missing=$null -eq $expected;expected_value=$expected}
        if($null -eq $expected){$intentional+=$cell}
    }
    $seriesList=@()
    for($index=1;$index -le 4;$index++){
        if($index -eq 1){$name='1/1';$formula='=SERIES("1/1",總表!$B$4:$AG$4,總表!$AO$5:$BT$5,1)';$type=82}
        else{$name=@('最遠','最近','平均')[$index-2];$column=@('AJ','AK','AL')[$index-2];$formula='=SERIES("'+$name+'",總表!$AI$5:$AI$36,總表!$'+$column+'$5:$'+$column+'$36,'+($index-1)+')';$type=@(51,51,65)[$index-2]}
        $series=[pscustomobject]@{Name=$name;Formula=$formula;ChartType=$type;Values=$values.Clone();XValues=$labels.Clone()}
        $series|Add-Member ScriptMethod Points {return [pscustomobject]@{Count=32}}
        $seriesList+=$series
    }
    $chartList=@()
    foreach($items in @(@($seriesList[0]),@($seriesList[1..3]))){
        $collection=[pscustomobject]@{Count=$items.Count;Items=$items}
        $collection|Add-Member ScriptMethod Item {param($index)return $this.Items[$index-1]}
        $chart=[pscustomobject]@{Collection=$collection;PlotVisibleOnly=$false}
        $chart|Add-Member ScriptMethod FullSeriesCollection {return $this.Collection}
        $chartList += [pscustomobject]@{Chart=$chart}
    }
    $objects=[pscustomobject]@{Count=2;Items=$chartList};$objects|Add-Member ScriptMethod Item {param($index)return $this.Items[$index-1]}
    $sheet=[pscustomobject]@{Name='總表';Snapshots=$snapshots;Objects=$objects}
    $sheet|Add-Member ScriptMethod Range {param($address)if(-not $this.Snapshots.ContainsKey($address)){throw "Missing fixture address $address"};return $this.Snapshots[$address]}
    $sheet|Add-Member ScriptMethod ChartObjects {return $this.Objects}
    $sheet|Add-Member ScriptMethod Evaluate {param($formula)
        if($formula -match '^COUNT\((\w+)\)$'){$value=$this.Snapshots[$Matches[1]].Value2;return [double]$(if($value -is [ValueType] -and $value -isnot [bool] -and [double]$value -ge 0){1}else{0})}
        if($formula -match '^ISNA\((\w+)\)$'){return $this.Snapshots[$Matches[1]].Value2 -ceq -2146826246}
        throw "Unsupported fixture Evaluate $formula"
    }
    $contract=[pscustomobject]@{direction_labels=$labels;included_indices=@(0..31);logical_days=@('2026-01-01');expected_matrix=@(,$values);helpers=$helpers;radar_helpers=$radarHelpers;verified_intentional_missing_count=6;radar_verified_intentional_missing_count=2;native_contract_report='fixture';native_contract_report_sha256='fixture';workbook_sha256='fixture';raw_selector_prerequisite='NOT_A_REAL_RAW_GATE'}
    return [pscustomobject]@{Sheet=$sheet;Contract=$contract;Series=$seriesList;Intentional=$intentional}
}

$tests=0
$fixture=New-Fixture;$gate=Get-NativeAuditModernMissing $fixture.Sheet $fixture.Contract
if($gate.status -ne 'PASS' -or $gate.intentional_error_addresses.Count -ne 8 -or $gate.primary_daily_cells -ne 32 -or $gate.series_points -ne 128 -or $gate.genuine_zero_points -ne 4){throw "Positive native missing/zero fixture failed: $($gate.error)"};$tests++
foreach($case in @('shifted_ref','different_formula','nonformula','numeric_NA','missing_literalNA','source_error','source_missing_zero','header','direction_order','matrix_width','extra_helper','missing_helper','radar_ref','radar_nonformula','radar_numericNA','missing_series_zero','zero_series_null','series_range','series_order','category_order','missing_chart','extra_chart','hidden_not_enabled','helper_visible')){
    $fixture=New-Fixture
    switch($case){
        'shifted_ref'{$fixture.Contract.helpers[0].source='C7'}
        'different_formula'{$fixture.Sheet.Snapshots['AJ5'].Formula='=NA()'}
        'nonformula'{$fixture.Sheet.Snapshots['AJ12'].HasFormula=$false}
        'numeric_NA'{$fixture.Sheet.Snapshots['AJ5'].Value2=-2146826246}
        'missing_literalNA'{$fixture.Sheet.Snapshots['AJ12'].Value2='#N/A'}
        'source_error'{$fixture.Sheet.Snapshots['B7'].Value2=-2146826246}
        'source_missing_zero'{$fixture.Sheet.Snapshots['I7'].Value2=0.0}
        'header'{$fixture.Sheet.Snapshots['AJ4'].Value2='MIN'}
        'direction_order'{$fixture.Contract.direction_labels[31]='北'}
        'matrix_width'{$fixture.Contract.expected_matrix=@(,@(0..30))}
        'extra_helper'{$fixture.Contract.helpers+= $fixture.Contract.helpers[0]}
        'missing_helper'{$fixture.Contract.helpers=$fixture.Contract.helpers[0..94]}
        'radar_ref'{$fixture.Contract.radar_helpers[0].source='C5'}
        'radar_nonformula'{$fixture.Sheet.Snapshots['AO5'].HasFormula=$false}
        'radar_numericNA'{$fixture.Sheet.Snapshots['AO5'].Value2=-2146826246}
        'missing_series_zero'{$fixture.Series[0].Values[7]=0.0}
        'zero_series_null'{$fixture.Series[0].Values[0]=$null}
        'series_range'{$fixture.Series[0].Formula='=SERIES("1/1",總表!$B$4:$AG$4,總表!$B$5:$AG$5,1)'}
        'series_order'{$fixture.Series[1].Name='最近'}
        'category_order'{$fixture.Series[3].XValues[31]='北'}
        'missing_chart'{$fixture.Sheet.Objects.Count=1}
        'extra_chart'{$fixture.Sheet.Objects.Count=3}
        'hidden_not_enabled'{$fixture.Sheet.Objects.Items[0].Chart.PlotVisibleOnly=$true}
        'helper_visible'{$fixture.Sheet.Snapshots['AO:BT'].EntireColumn.Hidden=$false}
    }
    $gate=Get-NativeAuditModernMissing $fixture.Sheet $fixture.Contract
    if($gate.status -ne 'FAIL'){throw "Native negative falsely passed: $case"};$tests++
}
$fixture=New-Fixture;$gate=Get-NativeAuditModernMissing $fixture.Sheet $fixture.Contract
$raw=[pscustomobject]@{status='FORMULA_ERRORS';formula_cell_count=200;error_count=8;areas=@();error=$null}
$classified=Get-NativeAuditFormulaDisposition $raw $gate
if($classified.status -ne 'VERIFIED_INTENTIONAL_CHART_MISSING' -or $classified.error_count -ne 8 -or $classified.unexpected_error_count -ne 0){throw 'Exact intentional error census not preserved'};$tests++
$raw.error_count=9;$classified=Get-NativeAuditFormulaDisposition $raw $gate
if($classified.status -ne 'FORMULA_ERRORS' -or $classified.unexpected_error_count -ne 1){throw 'Unexpected formula error hidden by NA exception'};$tests++
foreach($case in @('excess_allowed','duplicate','outside','unverified')){
    $fixture=New-Fixture;$gate=Get-NativeAuditModernMissing $fixture.Sheet $fixture.Contract;$raw.error_count=8
    switch($case){'excess_allowed'{$raw.error_count=7};'duplicate'{$gate.intentional_error_addresses+= $gate.intentional_error_addresses[0]};'outside'{$gate.intentional_error_addresses[0]='B7'};'unverified'{$gate.independent_contract_status='UNVERIFIED'}}
    $blocked=$false;try{Get-NativeAuditFormulaDisposition $raw $gate|Out-Null}catch{$blocked=$true}
    if(-not $blocked){throw "Native formula negative falsely passed $case"};$tests++
}
Write-Host "SELFTEST PASS: $tests strict native matrix/helper/series/count tests; genuine0 retained; unexpectedNA cannot pass."
