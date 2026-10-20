Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'excel_native_audit_helpers.ps1')
. (Join-Path $PSScriptRoot 'excel_native_historical_audit_helpers.ps1')
$tests=0
Assert-NativeAuditNumericValue 0 0 'genuinezero';$tests++
Assert-NativeAuditNumericValue 1e-14 1e-14 'tinypositive';$tests++
foreach($function in @('MAXIFS','STDEV.S')){
    $expected='=IF(COUNT(B2:B32)<1,"",_xlfn.'+$function+'(B2:B32))';$actual=$expected.Replace('_xlfn.','')
    if((Get-NativeHistoricalFormulaCanonical $expected) -cne (Get-NativeHistoricalFormulaCanonical $actual)){throw 'Known native prefix representation rejected'};$tests++
    foreach($wrong in @($actual.Replace('B2:B32','C2:C32'),$actual.Replace($function,'SUM'))){if((Get-NativeHistoricalFormulaCanonical $expected) -ceq (Get-NativeHistoricalFormulaCanonical $wrong)){throw 'Function/reference drift hidden by native prefix normalizer'};$tests++}
}
foreach($literal in @('="_xlfn.MAXIFS(B2)"','=_xlfn.UNKNOWN(B2)')){if((Get-NativeHistoricalFormulaCanonical $literal) -cne $literal){throw 'Literal/unknown function compatibility prefix was stripped'};$tests++}
function Assert-Blocked([scriptblock]$Action,[string]$Name){$blocked=$false;try{& $Action}catch{$blocked=$true};if(-not $blocked){throw "Historical negative control falsely PASS: $Name"};$script:tests++}
Assert-Blocked {Assert-NativeAuditNumericValue 0 1e-14 'tinypositivebecamezero'} 'tinypositivecannotbecomezero'
Assert-Blocked {Assert-NativeAuditNumericValue 1e-14 0 'zerobecametinypositive'} 'zerocannotbecometinypositive'
$plot=[pscustomobject]@{left=20.;top=30.;width=100.;height=100.}
$zero=[pscustomobject]@{left=70.;top=80.;width=0.;height=0.}
Assert-NativeHistoricalPointGeometry $zero $plot 0 1 1 1 32 0 100;$tests++
$north=[pscustomobject]@{left=70.;top=55.;width=0.;height=0.}
Assert-NativeHistoricalPointGeometry $north $plot 50 1 1 1 32 0 100;$tests++
$east=[pscustomobject]@{left=95.;top=80.;width=0.;height=0.}
Assert-NativeHistoricalPointGeometry $east $plot 50 1 1 9 32 0 100;$tests++
foreach($position in @(@(70000.,47000.),@(-70000.,47000.),@(70.,-84000.),@(70.,84000.))){
    $missing=[pscustomobject]@{left=$position[0];top=$position[1];width=0.;height=0.}
    Assert-NativeHistoricalPointGeometry $missing $plot $null 1 1 12 32 0 100;$tests++
}
$bar=[pscustomobject]@{left=30.;top=80.;width=5.;height=50.}
Assert-NativeHistoricalPointGeometry $bar $plot 50 2 1 1 32 0 100;$tests++
$nullBar=[pscustomobject]@{left=60.;top=-84390.;width=4.;height=84520.}
Assert-NativeHistoricalPointGeometry $nullBar $plot $null 2 1 29 32 0 100;$tests++
Assert-Blocked {Assert-NativeHistoricalPointGeometry $nullBar $plot 0 2 1 29 32 0 100} 'phantomNULLbarNotGenuineZero'
$line=[pscustomobject]@{left=30.;top=80.;width=0.;height=0.}
Assert-NativeHistoricalPointGeometry $line $plot 50 2 3 1 32 0 100;$tests++
Assert-Blocked {Assert-NativeHistoricalPointGeometry $zero $plot $null 1 1 1 32 0 100} 'fabricatedmissingatcenter'
Assert-Blocked {Assert-NativeAuditSeriesValue 0 $null 'Historical fabricatedmissing'} 'missingnativezero'
Assert-Blocked {Assert-NativeAuditSeriesValue '' $null 'Historical emptystringnotmissing'} 'missingnativestring'
Assert-Blocked {Assert-NativeHistoricalPointGeometry $north $plot 50 1 1 9 32 0 100} 'shifteddirection'
Assert-Blocked {Assert-NativeHistoricalPointGeometry $zero $plot 50 1 1 1 32 0 100} 'numericwrongradius'
Assert-Blocked {Assert-NativeHistoricalPointGeometry $bar $plot 25 2 1 1 32 0 100} 'barwrongvalue'
Assert-Blocked {Assert-NativeHistoricalPointGeometry $line $plot 25 2 3 1 32 0 100} 'linewrongvalue'
Assert-Blocked {Assert-NativeHistoricalPointGeometry $line $plot 50 2 3 1 32 0 0} 'invalidaxis'
Assert-Blocked {Assert-NativeHistoricalPointGeometry ([pscustomobject]@{left=[double]::NaN;top=80.;width=0.;height=0.}) $plot 50 2 3 1 32 0 100} 'nonfinitegeometry'
Assert-Blocked {Assert-NativeHistoricalPointGeometry ([pscustomobject]@{left=30.;top=80.;width=$null;height=0.}) $plot 50 2 3 1 32 0 100} 'nullgeometry'
$raw=[pscustomobject]@{status='FORMULA_ERRORS';error_count=1;formula_cell_count=32;error=$null;areas=@()}
$gate=[pscustomobject]@{status='PASS';independent_contract_status='PASS';missing_gate_kind='Historical/radar';role='full_a';intentional_error_addresses=@('AO2')}
$verdict=Get-NativeAuditFormulaDisposition $raw $gate
if($verdict.unexpected_error_count -ne 0 -or $verdict.intentional_chart_missing_count -ne 1){throw 'Historical exactNA permission positive failed'};$tests++
foreach($address in @('B2','AJ2','AO1','AO33','BU2')){$gate.intentional_error_addresses=@($address);Assert-Blocked {Get-NativeAuditFormulaDisposition $raw $gate|Out-Null} ('NAwrongscope'+$address)}
$gate.intentional_error_addresses=@('AO2');$gate.role='summary_a';Assert-Blocked {Get-NativeAuditFormulaDisposition $raw $gate|Out-Null} 'NAnotformulaSourceRole'
$gate.role='full_a';$gate.status='UNVERIFIED';Assert-Blocked {Get-NativeAuditFormulaDisposition $raw $gate|Out-Null} 'NAunverifiedNativeGate'
$gate.status='PASS';$gate.intentional_error_addresses=@('AO2','AO2');Assert-Blocked {Get-NativeAuditFormulaDisposition $raw $gate|Out-Null} 'NAduplicatedaddress'
Write-Host "SELFTEST PASS: $tests Historical geometry/value/exactNA positive and negative controls"
