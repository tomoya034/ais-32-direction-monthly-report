# Shared read-only native QA helpers. No workbook is saved or reformatted.
Set-StrictMode -Version Latest

function Release-NativeAuditCom([object]$Object) {
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($Object) | Out-Null } catch { Write-Warning $_.Exception.Message }
    }
}

function Resolve-NativeAuditPath([string]$Path) {
    $full = [IO.Path]::GetFullPath($Path)
    $current = [IO.Path]::GetPathRoot($full)
    foreach ($part in $full.Substring($current.Length).Split([char[]]@('\', '/'), [StringSplitOptions]::RemoveEmptyEntries)) {
        $current = Join-Path $current $part
        if (Test-Path -LiteralPath $current) {
            $item = Get-Item -LiteralPath $current -Force
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                $target = $item.ResolveLinkTarget($true)
                if ($null -eq $target) { throw "Unresolved reparse point: $current" }
                $current = $target.FullName
            }
        }
    }
    return [IO.Path]::GetFullPath($current).TrimEnd([char[]]@('\', '/'))
}

function Assert-NativeAuditOutput([string]$Output, [string[]]$Inputs) {
    $destination = Resolve-NativeAuditPath $Output
    foreach ($inputPath in $Inputs) {
        $protected = Resolve-NativeAuditPath $inputPath
        if ($destination.Equals($protected, [StringComparison]::OrdinalIgnoreCase) -or
            $destination.StartsWith($protected + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
            throw "Output must be outside input tree before mkdir: $destination"
        }
    }
    if (Test-Path -LiteralPath $destination) {
        if (-not (Test-Path -LiteralPath $destination -PathType Container) -or @(Get-ChildItem -LiteralPath $destination -Force).Count -gt 0) {
            throw 'New or empty output directory required; prior evidence cannot be overwritten.'
        }
    }
    return $destination
}

function Get-NativeAuditExcelIdentity([object]$Application) {
    if ($null -eq ('AisNativeQA.WindowOwner' -as [type])) {
        Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; namespace AisNativeQA { public static class WindowOwner { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); } }'
    }
    [uint32]$nativeExcelId = 0
    $window = [IntPtr]([long]$Application.Hwnd)
    [AisNativeQA.WindowOwner]::GetWindowThreadProcessId($window, [ref]$nativeExcelId) | Out-Null
    if ($nativeExcelId -eq 0) { throw 'Private Excel PID cannot be identified' }
    $process = Get-Process -Id $nativeExcelId
    return [pscustomobject]@{ pid = $nativeExcelId; hwnd = $window.ToInt64(); start_time = $process.StartTime.ToString('o') }
}

function Get-NativeAuditPng([string]$Path) {
    $result = [ordered]@{ status = 'UNVERIFIED'; path = $Path; bytes = 0; decoded = $false; error = $null }
    $picture = $null
    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw 'PNG file missing after Chart.Export' }
        $result.bytes = (Get-Item -LiteralPath $Path).Length
        if ($result.bytes -lt 33) { throw 'PNG is empty/truncated, although Chart.Export may return true' }
        $stream = [IO.File]::OpenRead($Path)
        try {
            $magic = [byte[]]::new(8)
            if ($stream.Read($magic, 0, 8) -ne 8 -or [Convert]::ToHexString($magic) -ne '89504E470D0A1A0A') { throw 'Invalid PNG signature' }
        } finally { $stream.Dispose() }
        Add-Type -AssemblyName System.Drawing.Common
        $picture = [Drawing.Image]::FromFile($Path)
        if ($picture.Width -le 0 -or $picture.Height -le 0) { throw 'Decoded PNG has no image dimensions' }
        $result.width_pixels = $picture.Width; $result.height_pixels = $picture.Height
        $result.decoded = $true; $result.status = 'VALID_PNG'
        $result.sha256 = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    } catch { $result.error = $_.Exception.Message }
    finally { if ($null -ne $picture) { $picture.Dispose() } }
    return [pscustomobject]$result
}

function Export-NativeAuditChart([object]$Chart, [object]$ChartObject, [object]$Sheet, [string]$Path) {
    $attempts = @()
    $selected = $null
    for ($attempt = 1; $attempt -le 2; $attempt++) {
        $target = if ($attempt -eq 1) { $Path } else { Join-Path (Split-Path -Parent $Path) (([IO.Path]::GetFileNameWithoutExtension($Path)) + '_retry2.png') }
        if (Test-Path -LiteralPath $target) { throw "PNG evidence collision: $target" }
        $exported = $false; $exportError = $null; $activation = 'NOT_REQUESTED'
        try {
            if ($attempt -gt 1) {
                # A zero-byte success is not success. Preserve it; retry in a
                # separate evidence file with this private workbook activated.
                $Sheet.Activate() | Out-Null
                $ChartObject.Activate() | Out-Null
                $Chart.Refresh()
                $activation = 'SHEET_CHART_ACTIVATE_AND_REFRESH'
            }
            $exported = [bool]$Chart.Export($target, 'PNG', $false)
        } catch { $exportError = $_.Exception.Message }
        $validation = Get-NativeAuditPng $target
        $attempts += [pscustomobject]@{ attempt = $attempt; activation = $activation; export_returned = $exported; export_error = $exportError; validation = $validation }
        if ($exported -and $validation.status -eq 'VALID_PNG') { $selected = $validation; break }
    }
    return [pscustomobject]@{ png_exported = $null -ne $selected; selected = $selected; attempts = $attempts }
}

function Assert-NativeAuditCount([object]$Value, [double]$Maximum) {
    if ($Value -is [bool] -or $Value -is [Array] -or $Value -is [Runtime.InteropServices.ErrorWrapper] -or
        $Value -isnot [ValueType] -or $Value -is [DateTime]) { throw 'Excel ISERROR count did not return a numeric scalar' }
    $count = [double]$Value
    if (-not [double]::IsFinite($count) -or $count -lt 0 -or $count -gt $Maximum -or $count -ne [Math]::Floor($count)) {
        throw 'Excel ISERROR count was nonfinite, negative, nonintegral or out of range'
    }
    return [long]$count
}

function Assert-NativeAuditExpectedCount([long]$Actual, [int]$Expected, [string]$Context) {
    # -1 retains the generic inventory workflow. AIS acceptance callers set
    # explicit 6 workbooks / 2 charts, so an empty chart collection cannot
    # satisfy a vacuous "all exported charts succeeded" predicate.
    if ($Actual -lt 0 -or $Expected -lt -1) { throw "Invalid expected-count contract: $Context" }
    if ($Expected -ge 0 -and $Actual -ne $Expected) {
        throw "Expected-count contract failed ($Context): expected $Expected, actual $Actual"
    }
}

function Get-NativeAuditFormulaErrors([object]$Sheet) {
    # HasFormula: false means no formulas; mixed returns null. Never infer
    # no errors from SpecialCells HRESULT 0x800A03EC. With formulas present,
    # independently count ISERROR over every formula-only area after recalc.
    # Docs: https://learn.microsoft.com/en-us/office/vba/api/excel.range.hasformula
    # https://learn.microsoft.com/en-us/office/vba/api/excel.range.specialcells
    # https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.evaluate
    $used = $null; $formulas = $null; $areas = $null
    $result = [ordered]@{ status = 'UNVERIFIED'; formula_cell_count = 0; error_count = 0; areas = @(); error = $null }
    try {
        $used = $Sheet.UsedRange
        $hasFormula = $used.HasFormula
        $result.has_formula_variant_type = if ($null -eq $hasFormula) { 'null' } else { $hasFormula.GetType().FullName }
        if ($hasFormula -is [bool] -and -not $hasFormula) {
            $result.status = 'VERIFIED_NO_FORMULAS'; return [pscustomobject]$result
        }
        # COM VT_NULL marshals as DBNull.Value on PowerShell/.NET, not only $null.
        if ($null -ne $hasFormula -and $hasFormula -isnot [DBNull] -and $hasFormula -isnot [bool]) { throw 'Unexpected UsedRange.HasFormula type' }
        # A protected ledger may reject SpecialCells even in a ReadOnly book.
        # Evaluate inspects the full bounded range without selecting/unprotecting
        # or editing it. ISFORMULA excludes constant error cells from the count.
        # Docs: https://support.microsoft.com/en-us/office/isformula-function-e4d1355f-7121-4ef2-801e-3839bfd6b1e5
        $isProtected = $false
        try { $isProtected = $Sheet.ProtectContents -is [bool] -and $Sheet.ProtectContents } catch { }
        if ($isProtected) {
            $address = [string]$used.Address($true, $true, 1, $false)
            $cells = [long]$used.CountLarge
            $formulaCount = Assert-NativeAuditCount ($Sheet.Evaluate("SUMPRODUCT(--ISFORMULA($address))")) $cells
            if ($formulaCount -le 0 -or ($hasFormula -is [bool] -and $hasFormula -and $formulaCount -ne $cells)) {
                throw 'Protected-range formula count contradicts native HasFormula'
            }
            $count = Assert-NativeAuditCount ($Sheet.Evaluate("SUMPRODUCT(--ISFORMULA($address),--ISERROR($address))")) $formulaCount
            $result.formula_cell_count = $formulaCount
            $result.error_count = $count
            $result.areas = @([pscustomobject]@{ address=$address; formula_cells=$formulaCount; errors=$count; method='READONLY_ISFORMULA_MASK' })
            $result.method = 'READONLY_ISFORMULA_MASK'
            $result.status = if ($count -eq 0) { 'VERIFIED_ZERO_ERRORS' } else { 'FORMULA_ERRORS' }
            return [pscustomobject]$result
        }
        $formulas = $used.SpecialCells(-4123)
        $areas = $formulas.Areas
        if ($areas.Count -le 0) { throw 'Formula scope empty despite HasFormula true/mixed' }
        for ($index = 1; $index -le $areas.Count; $index++) {
            $area = $null
            try {
                $area = $areas.Item($index)
                if ($area.HasFormula -isnot [bool] -or -not $area.HasFormula) { throw 'SpecialCells area was not formula-only' }
                $address = [string]$area.Address($true, $true, 1, $false)
                $cells = [long]$area.CountLarge
                $count = Assert-NativeAuditCount ($Sheet.Evaluate("SUMPRODUCT(--ISERROR($address))")) $cells
                $result.formula_cell_count += $cells; $result.error_count += $count
                $result.areas += [pscustomobject]@{ address = $address; formula_cells = $cells; errors = $count }
            } finally { Release-NativeAuditCom $area }
        }
        $result.status = if ($result.error_count -eq 0) { 'VERIFIED_ZERO_ERRORS' } else { 'FORMULA_ERRORS' }
    } catch { $result.error = $_.Exception.Message; $result.hresult = $_.Exception.HResult; $result.status = 'UNVERIFIED' }
    finally { Release-NativeAuditCom $areas; Release-NativeAuditCom $formulas; Release-NativeAuditCom $used }
    return [pscustomobject]$result
}

function Write-NativeAuditReport([string]$Path, [object]$Report) {
    $temporary = Join-Path (Split-Path -Parent $Path) ([Guid]::NewGuid().ToString('N') + '.json.tmp')
    try {
        $Report | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $temporary -Encoding utf8
        [IO.File]::Move($temporary, $Path, $true)
    } finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary } }
}

function Assert-NativeAuditNumericValue([object]$Actual, [object]$Expected, [string]$Context) {
    # A COM Excel error may marshal as Int32. Never accept error sentinels as
    # numeric data, and never coerce null/string/bool into a plausible zero.
    foreach ($value in @($Actual, $Expected)) {
        if ($null -eq $value -or $value -is [bool] -or $value -is [string] -or
            $value -is [Runtime.InteropServices.ErrorWrapper] -or $value -isnot [ValueType] -or
            $value -is [DateTime] -or -not [double]::IsFinite([double]$value) -or [double]$value -lt 0) {
            throw "Invalid native numeric value ($Context)"
        }
    }
    # Excel recomputes binary double statistics while writer caches are 16G.
    # Only a rounding-sized relative bound is permitted; missingness and zero
    # are exact and cannot disappear under this numerical tolerance.
    if ((([double]$Expected -eq 0) -ne ([double]$Actual -eq 0)) -or
        [Math]::Abs([double]$Actual - [double]$Expected) -gt 1e-12 * [Math]::Max(1, [Math]::Abs([double]$Expected))) {
        throw "Native numeric mismatch ($Context): expected $Expected actual $Actual"
    }
}

function Assert-NativeAuditSeriesValue([object]$Actual, [object]$Expected, [string]$Context) {
    if ($null -eq $Expected) {
        # The actual Excel Series.Values API must expose a genuine missing
        # point. Formula "" becoming0 is explicitly a producer FAIL.
        if ($null -ne $Actual) { throw "Missing chart point was not native null ($Context)" }
        return
    }
    Assert-NativeAuditNumericValue $Actual $Expected $Context
}

function Get-NativeAuditFormulaDisposition([object]$RawCheck, [object]$MissingGate) {
    $result = [ordered]@{}
    foreach ($property in $RawCheck.PSObject.Properties) { $result[$property.Name] = $property.Value }
    $result.intentional_chart_missing_count = 0
    $result.intentional_chart_missing_addresses = @()
    $result.unexpected_error_count = $RawCheck.error_count
    if ($RawCheck.status -eq 'UNVERIFIED') { return [pscustomobject]$result }
    if ($null -ne $MissingGate) {
        if ($MissingGate.status -ne 'PASS' -or $MissingGate.independent_contract_status -ne 'PASS') {
            throw 'Intentional chart missing markers require a complete independent/native gate'
        }
        $addresses = @($MissingGate.intentional_error_addresses)
        if (@($addresses | Sort-Object -Unique).Count -ne $addresses.Count) { throw 'Duplicate intentional error address' }
        foreach ($address in $addresses) {
            $historical = $MissingGate.PSObject.Properties.Name -contains 'missing_gate_kind' -and $MissingGate.missing_gate_kind -ceq 'Historical/radar'
            $validHistorical = $historical -and $MissingGate.role -in @('full_a','full_b','integrated') -and
                               $address -match '^(A[O-Z]|B[A-T])([2-9]|[12][0-9]|3[0-2])$'
            $validModern = -not $historical -and ($address -match '^(AJ|AK|AL)([5-9]|[12][0-9]|3[0-6])$' -or
                                                 $address -match '^(A[O-Z]|B[A-T])([5-9]|[12][0-9]|3[0-5])$')
            if (-not $validHistorical -and -not $validModern) {
                throw "Invalid Modern statistic helper allowlist address: $address"
            }
        }
        $rawCount = Assert-NativeAuditCount $RawCheck.error_count $RawCheck.formula_cell_count
        if ($addresses.Count -gt $rawCount) { throw 'Intentional marker count exceeds independently counted formula errors' }
        $result.intentional_chart_missing_count = $addresses.Count
        $result.intentional_chart_missing_addresses = $addresses
        $result.unexpected_error_count = $rawCount - $addresses.Count
        $result.status = if ($result.unexpected_error_count -gt 0) { 'FORMULA_ERRORS' }
                         elseif ($addresses.Count -gt 0) { 'VERIFIED_INTENTIONAL_CHART_MISSING' }
                         else { $RawCheck.status }
    }
    return [pscustomobject]$result
}

function Get-NativeAuditModernContract([string]$ReportPath, [string]$WorkbookPath, [string]$WorkbookHash) {
    if ([string]::IsNullOrWhiteSpace($ReportPath)) { return $null }
    $path = Resolve-NativeAuditPath $ReportPath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw 'Independent Modern contract report missing' }
    $reportHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    $report = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -DateKind String
    $candidates = @()
    if ($report.PSObject.Properties.Name -contains 'modern_chart_missing_contract') {
        $candidates = @($report.modern_chart_missing_contract)
    } elseif ($report.PSObject.Properties.Name -contains 'jobs') {
        foreach ($job in $report.jobs) {
            if ($job.PSObject.Properties.Name -contains 'modern_chart_missing_contract') {
                if ($job.pass -isnot [bool] -or -not $job.pass) { throw 'Independent workbook contract job did not PASS' }
                $candidates += $job.modern_chart_missing_contract
            }
        }
    } else { throw 'Unsupported independent Modern contract report shape' }
    $matched = @($candidates | Where-Object {
        $null -ne $_ -and (Resolve-NativeAuditPath $_.workbook).Equals($WorkbookPath,[StringComparison]::OrdinalIgnoreCase)
    })
    if ($matched.Count -eq 0) { return $null }
    if ($matched.Count -ne 1) { throw 'Duplicate independent Modern workbook contract' }
    $contract = $matched[0]
    if ($contract.schema -ne 1 -or $contract.role -ne 'Modern/default' -or
        $contract.state -ne 'PASS_SERIALIZATION_PREREQUISITES' -or
        $contract.workbook_sha256 -notmatch '^[0-9a-fA-F]{64}$' -or $contract.workbook_sha256 -ine $WorkbookHash) {
        throw 'Independent Modern contract schema/state/current-workbook hash failed'
    }
    if (@($contract.cache_sources).Count -eq 0) { throw 'Missing independent cache-source binding' }
    foreach ($source in $contract.cache_sources) {
        $sourcePath = Resolve-NativeAuditPath $source.path
        if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf) -or
            $source.sha256 -notmatch '^[0-9a-fA-F]{64}$' -or
            (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash -ine $source.sha256) {
            throw "Independent cache source changed/missing: $sourcePath"
        }
    }
    if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -cne $reportHash) { throw 'Independent contract report changed while reading' }
    $contract | Add-Member -NotePropertyName native_contract_report -NotePropertyValue $path
    $contract | Add-Member -NotePropertyName native_contract_report_sha256 -NotePropertyValue $reportHash
    return $contract
}

function Get-NativeAuditCellSnapshot([object]$Sheet, [string]$Address) {
    $cell = $null
    try {
        $cell = $Sheet.Range($Address)
        return [pscustomobject]@{ address=$Address; formula=[string]$cell.Formula; has_formula=$cell.HasFormula; value=$cell.Value2 }
    } finally { Release-NativeAuditCom $cell }
}

function Get-NativeAuditExcelColumn([int]$Index) {
    if ($Index -lt 1 -or $Index -gt 16384) { throw 'Invalid Excel column index' }
    $text = ''
    while ($Index -gt 0) {
        $Index--; $text = [char](65 + $Index % 26) + $text; $Index = [int][Math]::Floor($Index / 26)
    }
    return $text
}

function Assert-NativeAuditPrimaryValue([object]$Actual, [object]$Expected, [string]$Context) {
    if ($null -eq $Expected) {
        if ($null -ne $Actual -and ($Actual -isnot [string] -or $Actual -ne '')) { throw "Primary missing source was not blank ($Context)" }
    } else { Assert-NativeAuditNumericValue $Actual $Expected $Context }
}

function Get-NativeAuditModernMissing([object]$Sheet, [object]$Contract) {
    # The contract originates from independently reconstructed selected values,
    # not from Excel's own COUNT or a trusted filename. Serialized-cache/raw3/
    # selector3 gates remain separate mandatory final acceptance prerequisites.
    $result = [ordered]@{status='UNVERIFIED';independent_contract_status='UNVERIFIED';error=$null;
        intentional_error_addresses=@();helper_observations=@();series_observations=@();
        primary_daily_cells=0;statistic_helper_cells=0;radar_helper_cells=0;series_points=0;genuine_zero_points=0}
    try {
        if ($Sheet.Name -cne '總表') { throw 'Modern contract can apply only to 總表' }
        $labels = @('北','北微東','東北偏北','東北微北','東北','東北微東','東北偏東','東微北','東','東微南','東南偏東','東南微東','東南','東南微南','東南偏南','南微東','南','南微西','西南偏南','西南微南','西南','西南微西','西南偏西','西微南','西','西微北','西北偏西','西北微西','西北','西北微北','西北偏北','北微西')
        if (@($Contract.direction_labels).Count -ne 32 -or ($Contract.direction_labels -join "`n") -cne ($labels -join "`n")) { throw 'Independent32-direction order/labels failed' }
        $indices = @($Contract.included_indices)
        if ($indices.Count -eq 0 -or ($indices -join ',') -cne (@($indices | Sort-Object -Unique) -join ',')) { throw 'Included indices must be nonempty sorted unique' }
        foreach ($index in $indices) { if ($index -isnot [long] -and $index -isnot [int]) { throw 'Direction index was not an integer' }; if ($index -lt 0 -or $index -gt 31) { throw 'Direction index outside32' } }
        $days = @($Contract.logical_days); $matrix = @($Contract.expected_matrix)
        if ($days.Count -eq 0 -or $days.Count -gt 31 -or $matrix.Count -ne $days.Count -or
            ($days -join ',') -cne (@($days | Sort-Object -Unique) -join ',')) { throw 'Missing/duplicate/unordered independent logical-day matrix' }
        foreach ($day in $days) { if ($day -notmatch '^\d{4}-\d{2}-\d{2}$') { throw 'Invalid independent logical-day key' } }
        $headers = $null; $daily = $null
        try {
            $headers = $Sheet.Range('B4:AG4'); $headerValues = $headers.Value2
            for ($direction=0;$direction -lt 32;$direction++) { if ($headerValues[1,($direction+1)] -cne $labels[$direction]) { throw 'Native primary direction header drift' } }
            $last = $days.Count + 4; $daily = $Sheet.Range("B5:AG$last"); $values = $daily.Value2
            for ($day=0;$day -lt $days.Count;$day++) {
                if (@($matrix[$day]).Count -ne 32) { throw 'Independent daily matrix row not32' }
                for ($direction=0;$direction -lt 32;$direction++) {
                    $expected = $matrix[$day][$direction]
                    if ($direction -notin $indices -and $null -ne $expected) { throw 'Excluded direction had independent numeric result' }
                    Assert-NativeAuditPrimaryValue $values[($day+1),($direction+1)] $expected "$($days[$day])/$($labels[$direction])"
                    $result.primary_daily_cells++
                }
            }
        } finally { Release-NativeAuditCom $daily; Release-NativeAuditCom $headers }
        foreach ($pair in @(@('AI4','納入方位'),@('AJ4','MAX'),@('AK4','MIN'),@('AL4','平均'))) {
            $header = Get-NativeAuditCellSnapshot $Sheet $pair[0]
            if ($header.has_formula -isnot [bool] -or $header.has_formula -or $header.value -cne $pair[1]) { throw 'Native chart helper header drift' }
        }
        $helpers = @($Contract.helpers)
        if ($helpers.Count -ne 3*$indices.Count) { throw 'Missing/extra independent statistic helpers' }
        $position=0
        foreach ($direction in $indices) {
            $offset = [int]($position/3); $row=$offset+5
            $labelCell=Get-NativeAuditCellSnapshot $Sheet "AI$row"
            if($labelCell.has_formula -isnot [bool] -or $labelCell.has_formula -or $labelCell.value -cne $labels[$direction]){throw 'Native helper direction mapping drift'}
            for($stat=0;$stat -lt 3;$stat++) {
                $helper=$helpers[$position];$position++
                $function=@('MAX','MIN','AVERAGE')[$stat];$column=Get-NativeAuditExcelColumn ($direction+2)
                $sourceRow=$days.Count+6+$stat;$source="$column$sourceRow";$sourceRange="${column}5:$column$($days.Count+4)"
                $cell="$(Get-NativeAuditExcelColumn (36+$stat))$row"
                $sourceFormula='=IF(COUNT('+ $sourceRange +')<1,"",'+$function+'('+$sourceRange+'))'
                $formula='=IF(COUNT('+$source+')=0,NA(),'+$source+')'
                $numeric=@(for($dayIndex=0;$dayIndex -lt $days.Count;$dayIndex++){if($null -ne $matrix[$dayIndex][$direction]){$matrix[$dayIndex][$direction]}})
                $independentStat=$null
                if($numeric.Count -gt 0){
                    $summary=$numeric|Measure-Object -Minimum -Maximum -Average
                    $independentStat=@($summary.Maximum,$summary.Minimum,$summary.Average)[$stat]
                }
                Assert-NativeAuditPrimaryValue $helper.expected_value $independentStat 'Independent matrix statistic'
                if($helper.verified -isnot [bool] -or -not $helper.verified -or $helper.sheet -cne '總表' -or
                   $helper.cell -cne $cell -or $helper.direction_index -ne $direction -or $helper.direction -cne $labels[$direction] -or
                   $helper.statistic -cne $function -or $helper.source -cne $source -or $helper.source_range -cne $sourceRange -or
                   $helper.source_formula -cne $sourceFormula -or $helper.formula -cne $formula -or
                   $helper.expected_missing -isnot [bool] -or $helper.expected_missing -ne ($null -eq $helper.expected_value)){throw 'Independent helper scope/reference/formula contract drift'}
                $sourceCell=Get-NativeAuditCellSnapshot $Sheet $source;$helperCell=Get-NativeAuditCellSnapshot $Sheet $cell
                if($sourceCell.has_formula -isnot [bool] -or -not $sourceCell.has_formula -or $sourceCell.formula -cne $sourceFormula -or
                   $helperCell.has_formula -isnot [bool] -or -not $helperCell.has_formula -or $helperCell.formula -cne $formula){throw 'Native helper/source formula drift'}
                Assert-NativeAuditPrimaryValue $sourceCell.value $helper.expected_value $source
                $count=Assert-NativeAuditCount ($Sheet.Evaluate("COUNT($source)")) 1
                if($helper.expected_missing) {
                    if($count -ne 0 -or $Sheet.Evaluate("ISNA($cell)") -isnot [bool] -or -not $Sheet.Evaluate("ISNA($cell)")){throw 'Native intended missing helper lacks absent source/exactNA evidence'}
                    $result.intentional_error_addresses += $cell
                } else { if($count -ne 1){throw 'Native numeric source COUNT not1'};Assert-NativeAuditNumericValue $helperCell.value $helper.expected_value $cell }
                $result.helper_observations += [pscustomobject]@{cell=$cell;source=$source;formula=$helperCell.formula;expected_missing=$helper.expected_missing;source_count=$count;source_value=$sourceCell.value;helper_value=$helperCell.value}
                $result.statistic_helper_cells++
            }
        }
        if($result.intentional_error_addresses.Count -ne $Contract.verified_intentional_missing_count){throw 'Independent/native intentional marker census differs'}
        $radarHelpers=@($Contract.radar_helpers)
        if($radarHelpers.Count -ne 32*$days.Count){throw 'Missing/extra independent radar helpers'}
        $radarHeader=$null
        try{
            $radarHeader=$Sheet.Range('AO4:BT4');$headerValues=$radarHeader.Value2
            for($direction=0;$direction -lt 32;$direction++){if($headerValues[1,($direction+1)] -cne $labels[$direction]){throw 'Native radar helper direction header drift'}}
        }finally{Release-NativeAuditCom $radarHeader}
        $hiddenRange=$null;$hiddenColumns=$null
        try{
            $hiddenRange=$Sheet.Range('AO:BT');$hiddenColumns=$hiddenRange.EntireColumn
            if($hiddenColumns.Hidden -isnot [bool] -or -not $hiddenColumns.Hidden){throw 'Native radar chart-only helper columns were not hidden'}
        }finally{Release-NativeAuditCom $hiddenColumns;Release-NativeAuditCom $hiddenRange}
        $radarMissingCount=0
        for($dayIndex=0;$dayIndex -lt $days.Count;$dayIndex++){
            $row=$dayIndex+5
            $dateCell=Get-NativeAuditCellSnapshot $Sheet "A$row"
            if($dateCell.value -isnot [double] -or [DateTime]::FromOADate($dateCell.value).ToString('yyyy-MM-dd') -cne $days[$dayIndex]){throw 'Native logical-day row/date drift'}
            for($direction=0;$direction -lt 32;$direction++){
                $helper=$radarHelpers[32*$dayIndex+$direction];$expected=$matrix[$dayIndex][$direction]
                $source="$(Get-NativeAuditExcelColumn ($direction+2))$row";$cell="$(Get-NativeAuditExcelColumn ($direction+41))$row"
                $formula='=IF(COUNT('+$source+')=0,NA(),'+$source+')'
                if($helper.verified -isnot [bool] -or -not $helper.verified -or $helper.sheet -cne '總表' -or
                   $helper.cell -cne $cell -or $helper.source -cne $source -or $helper.formula -cne $formula -or
                   $helper.direction_index -ne $direction -or $helper.direction -cne $labels[$direction] -or
                   $helper.day -cne $days[$dayIndex] -or $helper.expected_missing -isnot [bool] -or
                   $helper.expected_missing -ne ($null -eq $expected)){throw 'Independent radar helper scope/reference/formula contract drift'}
                Assert-NativeAuditPrimaryValue $helper.expected_value $expected 'Independent radar matrix value'
                $sourceCell=Get-NativeAuditCellSnapshot $Sheet $source;$helperCell=Get-NativeAuditCellSnapshot $Sheet $cell
                if($sourceCell.formula -cne [string]$helper.source_formula -or $sourceCell.has_formula -isnot [bool] -or
                   $sourceCell.has_formula -ne ($direction -in $indices)){throw 'Native radar primary source formula drift'}
                if($helperCell.has_formula -isnot [bool] -or -not $helperCell.has_formula -or $helperCell.formula -cne $formula){throw 'Native radar helper formula drift'}
                Assert-NativeAuditPrimaryValue $sourceCell.value $expected $source
                $count=Assert-NativeAuditCount ($Sheet.Evaluate("COUNT($source)")) 1
                if($null -eq $expected){
                    $isNa=$Sheet.Evaluate("ISNA($cell)")
                    if($count -ne 0 -or $isNa -isnot [bool] -or -not $isNa){throw 'Native radar intendedNA lacks independently missing source'}
                    $result.intentional_error_addresses += $cell;$radarMissingCount++
                }else{if($count -ne 1){throw 'Native radar numeric source COUNT not1'};Assert-NativeAuditNumericValue $helperCell.value $expected $cell}
                $result.helper_observations += [pscustomobject]@{role='radar';cell=$cell;source=$source;formula=$helperCell.formula;expected_missing=$null -eq $expected;source_count=$count;source_value=$sourceCell.value;helper_value=$helperCell.value}
                $result.radar_helper_cells++
            }
        }
        if($radarMissingCount -ne $Contract.radar_verified_intentional_missing_count){throw 'Independent/native radar marker census differs'}
        $objects=$null
        try{
            $objects=$Sheet.ChartObjects()
            if($objects.Count -ne 2){throw 'Modern must have exactly2 charts'}
            for($chartIndex=1;$chartIndex -le 2;$chartIndex++){
                $chartObject=$null;$chart=$null;$collection=$null
                try{
                    $chartObject=$objects.Item($chartIndex);$chart=$chartObject.Chart;$collection=$chart.FullSeriesCollection()
                    $radarDays=@(for($dayIndex=0;$dayIndex -lt $days.Count;$dayIndex++){if(@($matrix[$dayIndex]|Where-Object {$null -ne $_}).Count -gt 0){$dayIndex}})
                    $emptyMonth=$radarDays.Count -eq 0
                    if($emptyMonth){$radarDays=@(0)}
                    $expectedCount=if($chartIndex -eq 1){$radarDays.Count}else{3}
                    if($collection.Count -ne $expectedCount){throw 'Native series census differs from independent matrix'}
                    if($chartIndex -eq 1 -and ($chart.PlotVisibleOnly -isnot [bool] -or $chart.PlotVisibleOnly)){throw 'Native radar hidden helper data not enabled'}
                    for($seriesIndex=1;$seriesIndex -le $expectedCount;$seriesIndex++){
                        $series=$null;$points=$null
                        try{
                            $series=$collection.Item($seriesIndex);$points=$series.Points();$nativeValues=@($series.Values);$categories=@($series.XValues)
                            if($chartIndex -eq 1){
                                $dayIndex=$radarDays[$seriesIndex-1];$row=$dayIndex+5;$date=[DateTime]::ParseExact($days[$dayIndex],'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)
                                $name=if($emptyMonth){'本時段無有效資料'}else{"$($date.Month)/$($date.Day)"}
                                $formula='=SERIES("'+$name+'",總表!$B$4:$AG$4,總表!$AO$'+$row+':$BT$'+$row+','+$seriesIndex+')'
                                $expectedValues=@($matrix[$dayIndex]);$expectedLabels=$labels;$pointCount=32
                            }else{
                                $name=@('最遠','最近','平均')[$seriesIndex-1];$column=@('AJ','AK','AL')[$seriesIndex-1];$last=$indices.Count+4
                                $formula='=SERIES("'+$name+'",總表!$AI$5:$AI$'+$last+',總表!$'+$column+'$5:$'+$column+'$'+$last+','+$seriesIndex+')'
                                $expectedValues=@(for($offset=0;$offset -lt $indices.Count;$offset++){$helpers[3*$offset+$seriesIndex-1].expected_value})
                                $expectedLabels=@($indices|ForEach-Object {$labels[$_]});$pointCount=$indices.Count
                                if([int]$series.ChartType -ne @(51,51,65)[$seriesIndex-1]){throw 'Native statistic series type/order drift'}
                            }
                            if([string]$series.Name -cne $name -or [string]$series.Formula -cne $formula -or
                               $nativeValues.Count -ne $pointCount -or $categories.Count -ne $pointCount -or $points.Count -ne $pointCount){throw 'Native chart series source/category/count drift'}
                            $missing=0;$zeros=0
                            for($point=0;$point -lt $pointCount;$point++){
                                if($categories[$point] -cne $expectedLabels[$point]){throw 'Native chart category mapping/order drift'}
                                Assert-NativeAuditSeriesValue $nativeValues[$point] $expectedValues[$point] "chart$chartIndex/series$seriesIndex/point$point"
                                if($null -eq $expectedValues[$point]){$missing++}elseif($expectedValues[$point] -eq 0){$zeros++}
                                $result.series_points++
                            }
                            $result.genuine_zero_points+=$zeros
                            $result.series_observations += [pscustomobject]@{chart_index=$chartIndex;series_index=$seriesIndex;name=$name;formula=[string]$series.Formula;point_count=$pointCount;missing_native_null_points=$missing;genuine_zero_points=$zeros;values=$nativeValues;categories=$categories}
                        }finally{Release-NativeAuditCom $points;Release-NativeAuditCom $series}
                    }
                }finally{Release-NativeAuditCom $collection;Release-NativeAuditCom $chart;Release-NativeAuditCom $chartObject}
            }
        }finally{Release-NativeAuditCom $objects}
        $result.independent_contract_status='PASS'
        $result.contract_report=$Contract.native_contract_report;$result.contract_report_sha256=$Contract.native_contract_report_sha256
        $result.workbook_sha256=$Contract.workbook_sha256
        $result.raw_selector_prerequisite=$Contract.raw_selector_prerequisite
        $result.status='PASS'
    } catch {$result.error=$_.Exception.Message;$result.status='FAIL'}
    return [pscustomobject]$result
}
