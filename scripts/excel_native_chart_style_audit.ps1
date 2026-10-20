<#
Read-only native Excel chart inventory. No workbook formatting or expected
results are changed. Property read errors remain UNVERIFIED. Successful PNG
export is not a visual/reference equality verdict.

Official object-model references used:
https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.format
https://learn.microsoft.com/en-us/office/vba/api/excel.series.format
https://learn.microsoft.com/en-us/office/vba/api/excel.series.formula
https://learn.microsoft.com/en-us/office/vba/api/excel.chart.export
https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.gapwidth
https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.overlap
https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea(object)
https://learn.microsoft.com/en-us/office/vba/api/word.colorformat
https://learn.microsoft.com/en-us/office/vba/api/word.lineformat.transparency

ChartFormat exposes Office LineFormat/FillFormat/ColorFormat. Transparency is
0=opaque, 1=clear. RGB is VBA's RGB long (red in the least-significant byte).
Computed HSV saturation is derived from resolved RGB; it is not a separate
Excel palette/tint setting. All size/position values are points.
#>
[CmdletBinding(DefaultParameterSetName = 'Audit')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Audit')][string]$Root,
    [Parameter(ParameterSetName = 'Audit')][ValidateSet('generated', 'reference')][string]$RootRole = 'generated',
    [Parameter(ParameterSetName = 'Audit')][string[]]$ReferenceRoot = @(),
    [Parameter(Mandatory = $true, ParameterSetName = 'Audit')][string]$OutputDirectory,
    [Parameter(ParameterSetName = 'Audit')][ValidateRange(-1, 10000)][int]$ExpectedWorkbookCount = -1,
    [Parameter(ParameterSetName = 'Audit')][ValidateRange(-1, 10000)][int]$ExpectedChartCountPerWorkbook = -1,
    [Parameter(ParameterSetName = 'Audit')][switch]$PlanOnly,
    [Parameter(ParameterSetName = 'Audit')][switch]$GeometryOnly,
    [Parameter(Mandatory = $true, ParameterSetName = 'Helpers')][switch]$SelfTest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'excel_native_audit_helpers.ps1')
$script:styleReadErrorCount = 0

function Get-ExcelProcessIdentity([object]$Application) {
    if ($null -eq ('AisNativeChartAudit.WindowOwner' -as [type])) {
        Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; namespace AisNativeChartAudit { public static class WindowOwner { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); } }'
    }
    [uint32]$styleExcelProcessId = 0
    $window = [IntPtr]([long]$Application.Hwnd)
    [AisNativeChartAudit.WindowOwner]::GetWindowThreadProcessId($window, [ref]$styleExcelProcessId) | Out-Null
    if ($styleExcelProcessId -eq 0) { throw 'Could not identify this automation instance by HWND' }
    $process = Get-Process -Id $styleExcelProcessId
    return [pscustomobject]@{ pid = $styleExcelProcessId; hwnd = $window.ToInt64(); start_time = $process.StartTime.ToString('o'); process_name = $process.ProcessName }
}

function Release-Com([object]$Object) {
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($Object) | Out-Null }
        catch { Write-Warning ('COM release failed: ' + $_.Exception.Message) }
    }
}

function Resolve-CanonicalPath([string]$Path) {
    $full = [IO.Path]::GetFullPath($Path)
    $base = [IO.Path]::GetPathRoot($full)
    $parts = $full.Substring($base.Length).Split([char[]]@('\', '/'), [StringSplitOptions]::RemoveEmptyEntries)
    $current = $base
    foreach ($part in $parts) {
        $current = Join-Path $current $part
        if (Test-Path -LiteralPath $current) {
            $item = Get-Item -LiteralPath $current -Force
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                $target = $item.ResolveLinkTarget($true)
                if ($null -eq $target) { throw "Cannot resolve reparse point: $current" }
                $current = $target.FullName
            }
        }
    }
    $canonical = [IO.Path]::GetFullPath($current)
    if ($canonical -eq [IO.Path]::GetPathRoot($canonical)) { return $canonical }
    return $canonical.TrimEnd([char[]]@('\', '/'))
}

function Assert-OutputOutsideInputs([string]$Output, [string[]]$Inputs) {
    $destination = Resolve-CanonicalPath $Output
    foreach ($inputPath in $Inputs) {
        $protected = Resolve-CanonicalPath $inputPath
        if ($destination.Equals($protected, [StringComparison]::OrdinalIgnoreCase) -or
            $destination.StartsWith($protected.TrimEnd([char[]]@('\', '/')) + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
            throw "Output must be outside every input tree: $destination / $protected"
        }
    }
    return $destination
}

function Get-ChartToken([string]$Role, [string]$Path, [int]$Sheet, [int]$Chart) {
    $identity = "$Role`n$Path`n$Sheet`n$Chart"
    $bytes = [Text.Encoding]::UTF8.GetBytes($identity)
    return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes)).Substring(0, 24).ToLowerInvariant()
}

function Convert-Rgb([object]$Value) {
    if ($null -eq $Value -or [long]$Value -lt 0 -or [long]$Value -gt 0xFFFFFF) {
        return [pscustomobject]@{ status = 'UNVERIFIED'; reason = 'Automatic/mixed/system legacy color is not a resolved RGB long'; raw = $Value }
    }
    $number = [long]$Value
    $r = [int]($number -band 255)
    $g = [int](($number -shr 8) -band 255)
    $b = [int](($number -shr 16) -band 255)
    $maximum = [Math]::Max($r, [Math]::Max($g, $b))
    $minimum = [Math]::Min($r, [Math]::Min($g, $b))
    return [pscustomobject]@{
        status = 'VERIFIED_READ'; rgb_hex = ('#{0:X2}{1:X2}{2:X2}' -f $r, $g, $b)
        red = $r; green = $g; blue = $b
        hsv_saturation = if ($maximum -eq 0) { 0.0 } else { [double]($maximum - $minimum) / $maximum }
        hsv_value = [double]$maximum / 255.0
    }
}

function Convert-PlainValue([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ([Runtime.InteropServices.Marshal]::IsComObject($Value)) { throw 'Property returned a COM object instead of a scalar/array' }
    if ($Value -is [Array]) {
        $items = [Collections.Generic.List[object]]::new()
        foreach ($item in $Value) { $items.Add((Convert-PlainValue $item)) }
        return ,$items.ToArray()
    }
    if ($Value -is [DateTime]) { return $Value.ToString('o') }
    if ($Value -is [double] -and (-not [double]::IsFinite($Value))) {
        return [pscustomobject]@{ status = 'UNVERIFIED'; nonfinite_value = [string]$Value }
    }
    if ($Value -is [string] -or $Value -is [ValueType]) { return $Value }
    throw ('Unsupported scalar type: ' + $Value.GetType().FullName)
}

function Read-PropertyPath([object]$Object, [string]$Path) {
    $cursor = $Object
    $acquired = [Collections.Generic.List[object]]::new()
    try {
        foreach ($member in $Path.Split('.')) {
            if ($null -eq $cursor) { throw "Null parent reading $Path" }
            $cursor = $cursor.$member
            if ($null -ne $cursor -and [Runtime.InteropServices.Marshal]::IsComObject($cursor)) { $acquired.Add($cursor) }
        }
        return [pscustomobject]@{ value = (Convert-PlainValue $cursor) }
    }
    finally {
        for ($index = $acquired.Count - 1; $index -ge 0; $index--) { Release-Com $acquired[$index] }
    }
}

function Get-Metadata([object]$Object, [string[]]$Properties, [string]$Context) {
    $metadata = [ordered]@{ read_status = 'VERIFIED_READ'; unverified = @() }
    foreach ($property in $Properties) {
        try { $metadata[$property] = (Read-PropertyPath $Object $property).value }
        catch {
            $script:styleReadErrorCount++
            $metadata.read_status = 'UNVERIFIED'
            $metadata[$property] = $null
            $metadata.unverified += [pscustomobject]@{
                context = $Context; property = $property; status = 'UNVERIFIED'
                error = $_.Exception.Message; hresult = $_.Exception.HResult
            }
        }
    }
    return $metadata
}

function Get-ColorMetadata([object]$Object, [string]$Prefix, [string]$Context) {
    $names = @('RGB', 'Type', 'ObjectThemeColor', 'Brightness', 'TintAndShade') | ForEach-Object { "$Prefix.$_" }
    $metadata = Get-Metadata $Object $names $Context
    $metadata.resolved_rgb = Convert-Rgb $metadata["$Prefix.RGB"]
    $metadata.effective_color_status = if ($null -eq $metadata["$Prefix.Type"] -or $metadata["$Prefix.Type"] -le 0) { 'UNVERIFIED_MIXED_OR_INHERITED' } else { 'VERIFIED_READ' }
    if ($metadata.effective_color_status -ne 'VERIFIED_READ') {
        $metadata.resolved_rgb.status = 'UNVERIFIED_MIXED_OR_INHERITED'
        $metadata.resolved_rgb | Add-Member NoteProperty reason 'Numeric RGB bits are observed, but mixed/inherited Color.Type is not an effective resolved appearance.' -Force
    }
    if ($metadata.resolved_rgb.status -eq 'UNVERIFIED') { $metadata.read_status = 'UNVERIFIED'; $script:styleReadErrorCount++ }
    return $metadata
}

function Get-LineMetadata([object]$Object, [string]$Context) {
    $metadata = Get-Metadata $Object @('Format.Line.Visible', 'Format.Line.Weight', 'Format.Line.DashStyle',
        'Format.Line.Style', 'Format.Line.Transparency') $Context
    $metadata.foreground = Get-ColorMetadata $Object 'Format.Line.ForeColor' $Context
    $metadata.background = Get-ColorMetadata $Object 'Format.Line.BackColor' $Context
    $metadata.effective_appearance_status = if ($null -eq $metadata['Format.Line.Weight'] -or $metadata['Format.Line.Weight'] -lt 0 -or
        $null -eq $metadata['Format.Line.Transparency'] -or $metadata['Format.Line.Transparency'] -lt 0 -or $metadata['Format.Line.Transparency'] -gt 1 -or
        $metadata['Format.Line.Visible'] -eq -2) { 'UNVERIFIED_MIXED_OR_INHERITED' } else { 'VERIFIED_READ' }
    return $metadata
}

function Get-FillMetadata([object]$Object, [string]$Context) {
    $metadata = Get-Metadata $Object @('Format.Fill.Visible', 'Format.Fill.Type', 'Format.Fill.Transparency') $Context
    $metadata.foreground = Get-ColorMetadata $Object 'Format.Fill.ForeColor' $Context
    $metadata.background = Get-ColorMetadata $Object 'Format.Fill.BackColor' $Context
    $metadata.effective_appearance_status = if ($null -eq $metadata['Format.Fill.Transparency'] -or $metadata['Format.Fill.Transparency'] -lt 0 -or
        $metadata['Format.Fill.Transparency'] -gt 1 -or $metadata['Format.Fill.Type'] -le 0) { 'UNVERIFIED_MIXED_OR_INHERITED' } else { 'VERIFIED_READ' }
    if ($metadata['Format.Fill.Type'] -eq 3) {
        $metadata.gradient = Get-Metadata $Object @('Format.Fill.GradientColorType', 'Format.Fill.GradientDegree',
            'Format.Fill.GradientStyle', 'Format.Fill.GradientVariant', 'Format.Fill.GradientAngle') $Context
    }
    else { $metadata.gradient = [pscustomobject]@{ status = 'NOT_APPLICABLE'; reason = 'Fill.Type is not gradient (3)' } }
    return $metadata
}

function Get-FontMetadata([object]$Object, [string]$Context) {
    $metadata = Get-Metadata $Object @('Font.Name', 'Font.Size', 'Font.Bold', 'Font.Italic', 'Font.Underline',
        'Font.Color', 'Font.ColorIndex', 'Font.ThemeColor', 'Font.TintAndShade') $Context
    $metadata.legacy_rgb = Convert-Rgb $metadata['Font.Color']
    if ($metadata.legacy_rgb.status -eq 'UNVERIFIED') { $metadata.read_status = 'UNVERIFIED'; $script:styleReadErrorCount++ }
    return $metadata
}

function Get-AreaMetadata([object]$Object, [string]$Context, [switch]$Plot) {
    $properties = @('Left', 'Top', 'Width', 'Height')
    if ($Plot) { $properties += @('InsideLeft', 'InsideTop', 'InsideWidth', 'InsideHeight', 'Position') }
    $metadata = Get-Metadata $Object $properties $Context
    $metadata.aspect_ratio = if ($null -ne $metadata.Width -and $metadata.Height -gt 0) { [double]$metadata.Width / $metadata.Height } else { $null }
    if ($Plot) {
        $metadata.font = [pscustomobject]@{
            status = 'NOT_APPLICABLE_BY_EXCEL_API'
            reason = 'PlotArea has no Font property. Visible text fonts belong to chart/axis titles, tick labels, legend or data labels.'
            source = 'https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea(object)'
        }
    }
    else { $metadata.font = Get-FontMetadata $Object $Context }
    $metadata.fill = Get-FillMetadata $Object $Context
    $metadata.line = Get-LineMetadata $Object $Context
    $metadata.legacy_background_border = Get-Metadata $Object @('Interior.Color', 'Interior.ColorIndex', 'Interior.Pattern',
        'Border.Color', 'Border.ColorIndex', 'Border.LineStyle', 'Border.Weight') $Context
    return $metadata
}

function Get-AxisMetadata([object]$Axis, [string]$Context) {
    $metadata = Get-Metadata $Axis @('Type', 'AxisGroup', 'Crosses', 'CrossesAt', 'ReversePlotOrder',
        'MajorTickMark', 'MinorTickMark', 'TickLabelPosition', 'HasTitle', 'HasMajorGridlines', 'HasMinorGridlines') $Context
    if ($metadata.Type -eq 2) {
        $metadata.scale = Get-Metadata $Axis @('MinimumScale', 'MaximumScale', 'MinimumScaleIsAuto', 'MaximumScaleIsAuto',
            'MajorUnit', 'MinorUnit', 'MajorUnitIsAuto', 'MinorUnitIsAuto', 'ScaleType', 'LogBase', 'DisplayUnit') $Context
    }
    else {
        $metadata.scale = Get-Metadata $Axis @('CategoryType', 'TickLabelSpacing', 'TickMarkSpacing', 'AxisBetweenCategories',
            'BaseUnit', 'BaseUnitIsAuto') $Context
        $metadata.numeric_scale = [pscustomobject]@{ status = 'NOT_APPLICABLE'; reason = 'Axis.Type is not a value axis (2)' }
    }
    $metadata.tick_labels = Get-Metadata $Axis @('TickLabels.NumberFormat', 'TickLabels.NumberFormatLinked',
        'TickLabels.Orientation', 'TickLabels.Font.Name', 'TickLabels.Font.Size', 'TickLabels.Font.Bold',
        'TickLabels.Font.Italic', 'TickLabels.Font.Color', 'TickLabels.Font.ColorIndex', 'TickLabels.Font.ThemeColor') $Context
    $metadata.line = Get-LineMetadata $Axis $Context
    foreach ($kind in @('Major', 'Minor')) {
        $presence = $metadata["Has${kind}Gridlines"]
        if ($presence) {
            $grid = $null
            try { $grid = $Axis."${kind}Gridlines"; $metadata["${kind}_gridlines"] = Get-LineMetadata $grid $Context }
            finally { Release-Com $grid }
        }
        else { $metadata["${kind}_gridlines"] = [pscustomobject]@{ status = 'NOT_APPLICABLE'; present = $presence } }
    }
    if ($metadata.HasTitle) {
        $title = $null
        try {
            $title = $Axis.AxisTitle
            $metadata.title = Get-Metadata $title @('Text', 'Left', 'Top', 'Width', 'Height', 'Orientation') $Context
            $metadata.title.font = Get-FontMetadata $title $Context
            $metadata.title.fill = Get-FillMetadata $title $Context
            $metadata.title.line = Get-LineMetadata $title $Context
        }
        finally { Release-Com $title }
    }
    return $metadata
}

function Get-SeriesMetadata([object]$Series, [int]$Order, [string]$Context) {
    $metadata = Get-Metadata $Series @('Name', 'Formula', 'FormulaLocal', 'ChartType', 'AxisGroup', 'PlotOrder',
        'IsFiltered', 'Values', 'XValues', 'Smooth', 'MarkerStyle', 'MarkerSize') $Context
    $metadata.collection_order = $Order
    $metadata.line = Get-LineMetadata $Series $Context
    $metadata.fill = Get-FillMetadata $Series $Context
    $metadata.points = @()
    $points = $null
    try {
        $points = $Series.Points()
        $metadata.point_count = [int]$points.Count
        for ($pointIndex = 1; $pointIndex -le $metadata.point_count; $pointIndex++) {
            $point = $null
            try {
                $point = $points.Item($pointIndex)
                $pointContext = "$Context/point:$pointIndex"
                $pointData = Get-Metadata $point @('MarkerStyle', 'MarkerSize', 'HasDataLabel') $pointContext
                $pointData.index = $pointIndex
                $pointData.line = Get-LineMetadata $point $pointContext
                $pointData.fill = Get-FillMetadata $point $pointContext
                if ($pointData.MarkerStyle -ne -4142) {
                    $pointData.marker_colors = Get-Metadata $point @('MarkerForegroundColor', 'MarkerForegroundColorIndex',
                        'MarkerBackgroundColor', 'MarkerBackgroundColorIndex') $pointContext
                }
                else { $pointData.marker_colors = [pscustomobject]@{ status = 'NOT_APPLICABLE'; reason = 'MarkerStyle is none (-4142)' } }
                $metadata.points += $pointData
            }
            finally { Release-Com $point }
        }
    }
    finally { Release-Com $points }
    return $metadata
}

function Get-ChartGroupMetadata([object]$Chart, [string]$Context) {
    # Gap/overlap belong to the chart group, not the individual series. Read
    # only documented 2-D bar/column properties; other groups are explicit N/A.
    $groups = $null
    $result = @()
    try {
        $groups = $Chart.ChartGroups()
        for ($groupIndex = 1; $groupIndex -le $groups.Count; $groupIndex++) {
            $group = $null; $series = $null; $first = $null
            try {
                $group = $groups.Item($groupIndex)
                $series = $group.SeriesCollection()
                if ($series.Count -lt 1) { throw 'Chart group has no series; cannot identify applicability' }
                $first = $series.Item(1)
                $kind = Get-Metadata $first @('ChartType', 'Name') "$Context/group:$groupIndex/first-series"
                $entry = [ordered]@{ collection_order = $groupIndex; series_count = [int]$series.Count; first_series = $kind }
                if ($kind.ChartType -in @(51, 52, 53, 57, 58, 59)) {
                    $entry.column_spacing = Get-Metadata $group @('GapWidth', 'Overlap') "$Context/group:$groupIndex"
                } else {
                    $entry.column_spacing = [pscustomobject]@{ status = 'NOT_APPLICABLE'; reason = 'Group is not a documented 2-D bar/column chart'; chart_type = $kind.ChartType }
                }
                $result += [pscustomobject]$entry
            }
            finally { Release-Com $first; Release-Com $series; Release-Com $group }
        }
    }
    finally { Release-Com $groups }
    return ,$result
}

function Write-AuditReport([string]$Path, [object]$Report) {
    # The output directory starts empty. A unique temporary cannot alias a
    # pre-existing raw/reference/cache file; replace only our own report.
    $temporary = Join-Path (Split-Path -Parent $Path) ([Guid]::NewGuid().ToString('N') + '.json.tmp')
    try {
        if (Test-Path -LiteralPath $temporary) { throw "Unexpected temporary collision: $temporary" }
        $Report | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $temporary -Encoding utf8
        [IO.File]::Move($temporary, $Path, $true)
    }
    finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary } }
}

if ($SelfTest) {
    Assert-NativeAuditExpectedCount 6 6 'style inventory workbooks'
    Assert-NativeAuditExpectedCount 2 2 'style inventory charts'
    foreach ($actual in @(0, 1, 3)) {
        $blocked = $false
        try { Assert-NativeAuditExpectedCount $actual 2 'style inventory missing/duplicate charts' } catch { $blocked = $true }
        if (-not $blocked) { throw 'Style chart cardinality false PASS' }
    }
    $red = Convert-Rgb 255
    if ($red.rgb_hex -ne '#FF0000' -or $red.hsv_saturation -ne 1) { throw 'RGB byte-order/saturation test failed' }
    if ((Convert-Rgb 0).hsv_saturation -ne 0 -or (Convert-Rgb -4105).status -ne 'UNVERIFIED') { throw 'RGB zero/automatic test failed' }
    if ((Get-ChartToken 'a' 'path' 1 1) -eq (Get-ChartToken 'a' 'path' 1 2)) { throw 'Chart-token index collision' }
    $fixture = [pscustomobject]@{ Width = 10; Nested = [pscustomobject]@{ Number = 0 } }
    $valid = Get-Metadata $fixture @('Width', 'Nested.Number', 'DoesNotExist') 'synthetic'
    if ($valid.Width -ne 10 -or $valid['Nested.Number'] -ne 0 -or $valid.unverified.Count -ne 1) { throw 'Missing-property/zero-value test failed' }
    $inputPath = Join-Path ([IO.Path]::GetTempPath()) 'ais-style-protected'
    $blocked = $false
    try { Assert-OutputOutsideInputs (Join-Path $inputPath 'reports') @($inputPath) | Out-Null } catch { $blocked = $true }
    if (-not $blocked) { throw 'Input/output containment test failed' }
    Write-Host 'SELFTEST PASS: RGB/order, saturation, automatic color, unique chart identity, explicit UNVERIFIED, path boundary.'
    exit 0
}

$resolvedRoot = Resolve-CanonicalPath $Root
if (-not (Test-Path -LiteralPath $resolvedRoot -PathType Container)) { throw "Missing Root: $resolvedRoot" }
$resolvedReferences = @($ReferenceRoot | ForEach-Object {
    $path = Resolve-CanonicalPath $_
    if (-not (Test-Path -LiteralPath $path -PathType Container)) { throw "Missing ReferenceRoot: $path" }
    $path
})
$resolvedOutput = Assert-OutputOutsideInputs $OutputDirectory @($resolvedRoot; $resolvedReferences)
$files = [Collections.Generic.List[object]]::new()
$seenPaths = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($inputSpec in @([pscustomobject]@{ role = $RootRole; root = $resolvedRoot }) + @($resolvedReferences | ForEach-Object { [pscustomobject]@{ role = 'reference'; root = $_ } })) {
    foreach ($file in @(Get-ChildItem -LiteralPath $inputSpec.root -Recurse -File -Filter '*.xlsx' | Sort-Object FullName)) {
        if ($file.Name.StartsWith('~$')) { throw "Open Excel temporary present; finish/close that workbook first: $($file.FullName)" }
        $canonical = Resolve-CanonicalPath $file.FullName
        if (-not $seenPaths.Add($canonical)) { throw "Input workbook selected twice: $canonical" }
        $files.Add([pscustomobject]@{ role = $inputSpec.role; root = $inputSpec.root; path = $canonical })
    }
}
if ($files.Count -eq 0) { throw 'No workbooks found' }
Assert-NativeAuditExpectedCount $files.Count $ExpectedWorkbookCount 'style input workbooks before mkdir/COM'
if ($PlanOnly) {
    [pscustomobject]@{ generated_workbooks = @($files | Where-Object role -eq 'generated').Count;
        reference_workbooks = @($files | Where-Object role -eq 'reference').Count; output = $resolvedOutput } | ConvertTo-Json
    exit 0
}
if (Test-Path -LiteralPath $resolvedOutput) {
    if (-not (Test-Path -LiteralPath $resolvedOutput -PathType Container) -or @(Get-ChildItem -LiteralPath $resolvedOutput -Force).Count -gt 0) {
        throw 'OutputDirectory must be new or empty to prevent overwriting unrelated artifacts.'
    }
}
New-Item -ItemType Directory -Path $resolvedOutput -Force | Out-Null
$imageDirectory = Join-Path $resolvedOutput 'charts'
New-Item -ItemType Directory -Path $imageDirectory | Out-Null
$reportPath = Join-Path $resolvedOutput 'excel_native_chart_style_audit.json'
$chartNames = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$workbooks = [Collections.Generic.List[object]]::new()
$excel = $null
$excelBooks = $null
$started = Get-Date
$report = [ordered]@{
    schema = 'AIS_NATIVE_CHART_STYLE_INVENTORY_1'; root = $resolvedRoot; root_role = $RootRole; reference_roots = $resolvedReferences
    capture_scope = if ($GeometryOnly) { 'GEOMETRY_AND_SERIES_FORMULAS_NO_POINT_STYLES' } else { 'DETAILED_STYLE_WITH_ALL_POINT_STYLES' }
    created_at = $started.ToString('o'); property_read_errors = 0; workbooks = $workbooks
    units = 'points'; transparency = '0 opaque; 1 clear'
    visual_reference_equality = 'UNVERIFIED: inventory/export only; requires explicit role-aware metadata and visual comparison'
    source_count_expected = $files.Count; completed_workbooks = 0
    asserted_workbook_count = $ExpectedWorkbookCount; asserted_chart_count_per_workbook = $ExpectedChartCountPerWorkbook
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $report.excel_version = [string]$excel.Version
    $report.excel_build = [string]$excel.Build
    $report.owner_powershell_pid = $PID
    $report.excel_process = Get-ExcelProcessIdentity $excel
    Write-AuditReport $reportPath $report
    $excelBooks = $excel.Workbooks
    foreach ($file in $files) {
        Write-Host ("[{0}/{1}] {2}: {3}" -f ($workbooks.Count + 1), $files.Count, $file.role, $file.path)
        $before = (Get-FileHash -LiteralPath $file.path -Algorithm SHA256).Hash
        $errorStart = $script:styleReadErrorCount
        $entry = [ordered]@{ role = $file.role; path = $file.path; sha256_before = $before;
            opened_read_only = $false; charts = @(); fatal_error = $null; native_access_status = 'FAIL'; style_capture_status = 'UNVERIFIED' }
        $book = $null
        $sheets = $null
        $chartSheets = $null
        try {
            $book = $excelBooks.Open($file.path, 0, $true)
            $entry.opened_read_only = [bool]$book.ReadOnly
            if (-not $entry.opened_read_only) { throw 'Excel did not open read-only' }
            $normalStyle = $null; $bookStyles = $null
            try {
                $bookStyles = $book.Styles
                $normalStyle = $bookStyles.Item('Normal')
                $entry.normal_style = Get-Metadata $normalStyle @('Name', 'NameLocal', 'BuiltIn', 'Font.Name', 'Font.Size', 'Font.Bold', 'Font.Italic') $file.path
            } finally { Release-Com $normalStyle; Release-Com $bookStyles }
            $entry.sheet_geometry = @()
            $excel.Calculation = -4135
            $excel.CalculateFullRebuild()
            $sheets = $book.Worksheets
            $chartSheets = $book.Charts
            $entry.chart_sheet_count = [int]$chartSheets.Count
            if ($entry.chart_sheet_count -gt 0) {
                throw 'UNVERIFIED: standalone chart sheets require additional traversal; embedded-only inventory cannot pass this workbook.'
            }
            for ($sheetIndex = 1; $sheetIndex -le $sheets.Count; $sheetIndex++) {
                $sheet = $null
                $objects = $null
                try {
                    $sheet = $sheets.Item($sheetIndex)
                    $objects = $sheet.ChartObjects()
                    if ($objects.Count -gt 0) {
                        # All AIS reference anchors end by column 24/row 57. Record
                        # the complete enclosing grid and explicit bounded scope.
                        $grid = Get-Metadata $sheet @('Name', 'StandardWidth', 'StandardHeight') $file.path
                        $grid.columns = @(); $grid.rows = @(); $grid.column_scope = '1..40'; $grid.row_scope = '1..64'
                        for ($columnNumber = 1; $columnNumber -le 40; $columnNumber++) {
                            $columnRange = $null
                            try { $columnRange = $sheet.Columns.Item($columnNumber); $grid.columns += [pscustomobject]@{ index = $columnNumber; metadata = (Get-Metadata $columnRange @('Left', 'Width', 'ColumnWidth', 'Hidden') "$($file.path)/column:$columnNumber") } }
                            finally { Release-Com $columnRange }
                        }
                        for ($rowNumber = 1; $rowNumber -le 64; $rowNumber++) {
                            $rowRange = $null
                            try { $rowRange = $sheet.Rows.Item($rowNumber); $grid.rows += [pscustomobject]@{ index = $rowNumber; metadata = (Get-Metadata $rowRange @('Top', 'Height', 'RowHeight', 'Hidden') "$($file.path)/row:$rowNumber") } }
                            finally { Release-Com $rowRange }
                        }
                        $entry.sheet_geometry += $grid
                    }
                    for ($chartIndex = 1; $chartIndex -le $objects.Count; $chartIndex++) {
                        $chartObject = $null; $chart = $null; $area = $null; $plot = $null; $axes = $null; $seriesCollection = $null
                        $context = "$($file.role):$($file.path)/sheet:$sheetIndex/chart:$chartIndex"
                        $chartData = [ordered]@{ sheet = [string]$sheet.Name; sheet_index = $sheetIndex; chart_index = $chartIndex;
                            metadata_errors = 0; fatal_error = $null; png_exported = $false }
                        $chartErrorStart = $script:styleReadErrorCount
                        try {
                            $chartObject = $objects.Item($chartIndex)
                            $chart = $chartObject.Chart
                            $chartData.chart_object = Get-Metadata $chartObject @('Name', 'Left', 'Top', 'Width', 'Height', 'Placement') $context
                            $chartData.chart = Get-Metadata $chart @('ChartType', 'ChartStyle', 'HasTitle', 'HasLegend', 'PlotVisibleOnly', 'DisplayBlanksAs') $context
                            $chartData.chart_groups = Get-ChartGroupMetadata $chart $context
                            $area = $chart.ChartArea
                            $plot = $chart.PlotArea
                            $chartData.chart_area = Get-AreaMetadata $area $context
                            $chartData.plot_area = Get-AreaMetadata $plot $context -Plot
                            if ($chartData.chart.HasTitle) {
                                $title = $null
                                try {
                                    $title = $chart.ChartTitle
                                    $chartData.title = Get-Metadata $title @('Text', 'Left', 'Top', 'Width', 'Height', 'Orientation', 'IncludeInLayout') $context
                                    $chartData.title.font = Get-FontMetadata $title $context
                                    $chartData.title.fill = Get-FillMetadata $title $context
                                    $chartData.title.line = Get-LineMetadata $title $context
                                } finally { Release-Com $title }
                            } else { $chartData.title = [pscustomobject]@{ status = 'NOT_APPLICABLE'; present = $chartData.chart.HasTitle } }
                            if ($chartData.chart.HasLegend) {
                                $legend = $null
                                try {
                                    $legend = $chart.Legend
                                    $chartData.legend = Get-Metadata $legend @('Position', 'Left', 'Top', 'Width', 'Height', 'IncludeInLayout') $context
                                    $chartData.legend.font = Get-FontMetadata $legend $context
                                    $chartData.legend.fill = Get-FillMetadata $legend $context
                                    $chartData.legend.line = Get-LineMetadata $legend $context
                                } finally { Release-Com $legend }
                            } else { $chartData.legend = [pscustomobject]@{ status = 'NOT_APPLICABLE'; present = $chartData.chart.HasLegend } }
                            $chartData.axes = @()
                            $axes = $chart.Axes()
                            $chartData.axis_count = [int]$axes.Count
                            $axisIndex = 0
                            foreach ($axis in $axes) {
                                $axisIndex++
                                try { $chartData.axes += (Get-AxisMetadata $axis "$context/axis:$axisIndex") }
                                finally { Release-Com $axis }
                            }
                            try { $seriesCollection = $chart.FullSeriesCollection(); $chartData.series_collection_type = 'FullSeriesCollection' }
                            catch {
                                $chartData.full_series_unverified = [pscustomobject]@{ status = 'UNVERIFIED'; error = $_.Exception.Message }
                                $script:styleReadErrorCount++
                                $seriesCollection = $chart.SeriesCollection(); $chartData.series_collection_type = 'SeriesCollection fallback'
                            }
                            $chartData.series_count = [int]$seriesCollection.Count
                            $chartData.series = @()
                            for ($seriesIndex = 1; $seriesIndex -le $seriesCollection.Count; $seriesIndex++) {
                                $series = $null
                                try {
                                    $series = $seriesCollection.Item($seriesIndex)
                                    if ($GeometryOnly) {
                                        $seriesData = Get-Metadata $series @('Name', 'Formula', 'FormulaLocal', 'ChartType', 'IsFiltered') "$context/series:$seriesIndex"
                                        $seriesData.order = $seriesIndex; $seriesData.point_styles = 'NOT_CAPTURED_GEOMETRY_ONLY'
                                        $chartData.series += $seriesData
                                    } else { $chartData.series += (Get-SeriesMetadata $series $seriesIndex "$context/series:$seriesIndex") }
                                }
                                finally { Release-Com $series }
                            }
                            $token = Get-ChartToken $file.role $file.path $sheetIndex $chartIndex
                            $pngName = "$($file.role)_$token.png"
                            if (-not $chartNames.Add($pngName)) { throw "PNG identity collision: $pngName" }
                            $png = Join-Path $imageDirectory $pngName
                            if (Test-Path -LiteralPath $png) { throw "PNG destination already exists: $png" }
                            $verifiedExport = Export-NativeAuditChart $chart $chartObject $sheet $png
                            $chartData.png_export_attempts = $verifiedExport.attempts
                            $chartData.png_exported = $verifiedExport.png_exported
                            $chartData.png = if ($verifiedExport.png_exported) { $verifiedExport.selected.path } else { $png }
                            $chartData.png_bytes = if ($verifiedExport.png_exported) { $verifiedExport.selected.bytes } else { 0 }
                            $chartData.png_sha256 = if ($verifiedExport.png_exported) { $verifiedExport.selected.sha256 } else { $null }
                            $chartData.png_validation = if ($verifiedExport.png_exported) { $verifiedExport.selected } else { [pscustomobject]@{ status = 'UNVERIFIED_EXPORT_FAILED_AFTER_RETRY' } }
                        }
                        catch { $chartData.fatal_error = $_.Exception.Message }
                        finally {
                            foreach ($child in @($seriesCollection, $axes, $plot, $area, $chart, $chartObject)) { Release-Com $child }
                            $chartData.metadata_errors = $script:styleReadErrorCount - $chartErrorStart
                            $chartData.style_capture_status = if ($GeometryOnly) { 'LIMITED_GEOMETRY_ONLY' } elseif ($chartData.metadata_errors -eq 0 -and $null -eq $chartData.fatal_error) { 'COMPLETE' } else { 'UNVERIFIED' }
                            $entry.charts += $chartData
                        }
                    }
                }
                finally { Release-Com $objects; Release-Com $sheet }
            }
        }
        catch { $entry.fatal_error = $_.Exception.Message }
        finally {
            Release-Com $sheets
            Release-Com $chartSheets
            if ($null -ne $book) {
                try { $book.Close($false) } catch { $entry.fatal_error = 'Close without saving failed: ' + $_.Exception.Message }
                Release-Com $book
            }
        }
        $entry.sha256_after = (Get-FileHash -LiteralPath $file.path -Algorithm SHA256).Hash
        $entry.file_unchanged = $entry.sha256_before -eq $entry.sha256_after
        $entry.property_read_errors = $script:styleReadErrorCount - $errorStart
        try { Assert-NativeAuditExpectedCount $entry.charts.Count $ExpectedChartCountPerWorkbook $file.path }
        catch { $entry.fatal_error = $_.Exception.Message }
        $entry.native_access_status = if ($entry.opened_read_only -and $entry.file_unchanged -and $null -eq $entry.fatal_error -and
            @($entry.charts | Where-Object { -not $_.png_exported -or $null -ne $_.fatal_error }).Count -eq 0) { 'PASS' } else { 'FAIL' }
        $entry.style_capture_status = if ($GeometryOnly) { 'LIMITED_GEOMETRY_ONLY' } elseif ($entry.property_read_errors -eq 0 -and $entry.native_access_status -eq 'PASS') { 'COMPLETE' } else { 'UNVERIFIED' }
        $workbooks.Add($entry)
        $report.completed_workbooks = $workbooks.Count
        $report.property_read_errors = $script:styleReadErrorCount
        Write-AuditReport $reportPath $report
    }
}
finally {
    Release-Com $excelBooks
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { Write-Warning ('Excel.Quit failed: ' + $_.Exception.Message) }
        Release-Com $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $report.completed_workbooks = $workbooks.Count
    $report.elapsed_seconds = [Math]::Round(((Get-Date) - $started).TotalSeconds, 3)
    $report.native_access_pass_count = @($workbooks | Where-Object native_access_status -eq 'PASS').Count
    $report.native_access_fail_count = @($workbooks | Where-Object native_access_status -ne 'PASS').Count
    $report.style_capture_complete_count = @($workbooks | Where-Object style_capture_status -eq 'COMPLETE').Count
    $report.geometry_only_count = @($workbooks | Where-Object style_capture_status -eq 'LIMITED_GEOMETRY_ONLY').Count
    $report.style_capture_unverified_count = @($workbooks | Where-Object style_capture_status -eq 'UNVERIFIED').Count
    $report.property_read_errors = $script:styleReadErrorCount
    Write-AuditReport $reportPath $report
}
[pscustomobject]$report | Select-Object completed_workbooks, native_access_pass_count, native_access_fail_count, style_capture_complete_count,
    style_capture_unverified_count, property_read_errors, elapsed_seconds | Format-List
if ($report.completed_workbooks -ne $files.Count -or $report.native_access_fail_count -gt 0) { exit 1 }
if ($report.style_capture_unverified_count -gt 0) { exit 2 }
exit 0
