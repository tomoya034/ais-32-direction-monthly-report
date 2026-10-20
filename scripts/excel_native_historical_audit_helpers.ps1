# Independent Historical native contracts. Dot-source after excel_native_audit_helpers.ps1.
# No original/reference/workbook is modified or saved. NA permission is chart-only.
Set-StrictMode -Version Latest

function Get-NativeHistoricalFormulaCanonical([string]$Formula){
    # Excel16.0 resolves these two supported OOXML compatibility prefixes.
    # Normalize only function tokens outside string literals; preserve all references.
    $parts=[regex]::Split($Formula,'("(?:[^"]|"")*")')
    for($index=0;$index -lt $parts.Count;$index+=2){$parts[$index]=[regex]::Replace($parts[$index],'(?<![A-Za-z0-9_.])_xlfn\.(MAXIFS|STDEV\.S)\(','$1(')}
    return $parts -join ''
}

function Assert-NativeHistoricalPointGeometry([object]$Geometry,[object]$Plot,[object]$Expected,[int]$ChartIndex,[int]$SeriesIndex,[int]$PointIndex,[int]$PointCount,[double]$AxisMinimum,[double]$AxisMaximum){
    foreach($field in @('left','top','width','height')){
        $value=$Geometry.$field
        if($null -eq $value -or $value -is [bool] -or $value -is [string] -or -not [double]::IsFinite([double]$value)){throw 'Historical point geometry unavailable/nonfinite'}
    }
    if($Geometry.width -lt 0 -or $Geometry.height -lt 0 -or $Plot.width -le 0 -or $Plot.height -le 0 -or
       -not [double]::IsFinite($AxisMinimum) -or -not [double]::IsFinite($AxisMaximum) -or $AxisMaximum -le $AxisMinimum){throw 'Historical invalid point/plot/axis geometry'}
    # Missing filled-radar points can have very large POSITIVE or NEGATIVE coordinates.
    # Checking only a negative sentinel misses six of the eleven KLNG excluded directions.
    $intersects=$Geometry.left+$Geometry.width -ge $Plot.left-0.02 -and $Geometry.left -le $Plot.left+$Plot.width+0.02 -and
                $Geometry.top+$Geometry.height -ge $Plot.top-0.02 -and $Geometry.top -le $Plot.top+$Plot.height+0.02
    if($null -eq $Expected){
        # A native NULL column point reports a far-offscreen Top but a phantom
        # Height stretching back to baseline. That placeholder is not a painted
        # rectangle. Require the actual point anchor off-plot; native NULL was
        # independently proved above and may never be replaced by numeric0.
        $anchorInside=$Geometry.left -ge $Plot.left-0.02 -and $Geometry.left -le $Plot.left+$Plot.width+0.02 -and
                      $Geometry.top -ge $Plot.top-0.02 -and $Geometry.top -le $Plot.top+$Plot.height+0.02
        if($anchorInside){throw 'Historical missing point anchor intersects actual plot area'};return
    }
    if(-not $intersects){throw 'Historical numeric point lies outside actual plot area'}
    if($Expected -lt $AxisMinimum -or $Expected -gt $AxisMaximum){throw 'Historical numeric value outside native axis bounds'}
    $ratio=([double]$Expected-$AxisMinimum)/($AxisMaximum-$AxisMinimum)
    if($ChartIndex -eq 1){
        $angle=2*[Math]::PI*($PointIndex-1)/$PointCount
        $radius=[Math]::Min($Plot.width,$Plot.height)/2*$ratio
        $expectedX=$Plot.left+$Plot.width/2+$radius*[Math]::Sin($angle)
        $expectedY=$Plot.top+$Plot.height/2-$radius*[Math]::Cos($angle)
        $actualX=$Geometry.left+$Geometry.width/2;$actualY=$Geometry.top+$Geometry.height/2
        if([Math]::Abs($actualX-$expectedX) -gt 0.02 -or [Math]::Abs($actualY-$expectedY) -gt 0.02){throw 'Historical radar geometry does not represent independent direction/value'}
    }else{
        $expectedY=$Plot.top+$Plot.height*(1-$ratio)
        $actualY=if($SeriesIndex -eq 3){$Geometry.top+$Geometry.height/2}else{$Geometry.top}
        if([Math]::Abs($actualY-$expectedY) -gt 0.02){throw 'Historical statistic point geometry does not represent independent value'}
        if($SeriesIndex -lt 3 -and [Math]::Abs($Geometry.height-($Plot.top+$Plot.height-$expectedY)) -gt 0.02){throw 'Historical bar height does not represent independent value'}
    }
}

function Get-NativeAuditHistoricalContract([string]$ReportPath,[string]$WorkbookPath,[string]$WorkbookHash){
    if([string]::IsNullOrWhiteSpace($ReportPath)){return $null}
    $path=Resolve-NativeAuditPath $ReportPath
    $reportHash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    $report=Get-Content -LiteralPath $path -Raw|ConvertFrom-Json -DateKind String
    $candidates=@()
    if($report.PSObject.Properties.Name -contains 'historical_chart_contracts'){$candidates=@($report.historical_chart_contracts)}
    elseif($report.PSObject.Properties.Name -contains 'jobs'){
        foreach($job in $report.jobs){
            if($job.PSObject.Properties.Name -contains 'historical_chart_contracts'){
                if($job.pass -isnot [bool] -or -not $job.pass){throw 'Historical independent whole-job prerequisite FAIL'}
                if(@($job.historical_chart_contracts).Count -ne 5 -or @($job.historical_chart_contracts.role|Sort-Object -Unique).Count -ne 5){throw 'Historical independent role census is not exactly five'}
                $candidates += $job.historical_chart_contracts
            }
        }
    }else{throw 'Missing Historical independent chart contracts'}
    $matched=@($candidates|Where-Object {(Resolve-NativeAuditPath $_.workbook).Equals($WorkbookPath,[StringComparison]::OrdinalIgnoreCase)})
    if($matched.Count -eq 0){return $null}
    if($matched.Count -ne 1){throw 'Duplicate Historical independent workbook contract'}
    $contract=$matched[0]
    if($contract.schema -ne 1 -or $contract.state -cne 'PASS_SERIALIZATION_PREREQUISITES' -or
       $contract.role -notin @('full_a','full_b','summary_a','summary_b','integrated') -or
       $contract.workbook_sha256 -ine $WorkbookHash){throw 'Historical independent schema/role/state/workbook hash failed'}
    if(@($contract.cache_sources).Count -eq 0){throw 'Historical independent cache binding absent'}
    foreach($source in $contract.cache_sources){
        if($source.sha256 -notmatch '^[0-9a-fA-F]{64}$' -or (Get-FileHash -LiteralPath (Resolve-NativeAuditPath $source.path) -Algorithm SHA256).Hash -ine $source.sha256){throw 'Historical independent cache changed'}
    }
    if((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -cne $reportHash){throw 'Historical independent report changed while reading'}
    $contract|Add-Member -NotePropertyName native_contract_report -NotePropertyValue $path
    $contract|Add-Member -NotePropertyName native_contract_report_sha256 -NotePropertyValue $reportHash
    return $contract
}

function Get-NativeAuditHistoricalMissing([object]$Book,[object]$Contract){
    $result=[ordered]@{status='UNVERIFIED';independent_contract_status='UNVERIFIED';error=$null;missing_gate_kind='Historical/radar';
        role=$Contract.role;scope=$Contract.scope;intentional_error_addresses=@();source_observations=@();helper_observations=@();
        series_observations=@();source_cells=0;radar_helper_cells=0;series_points=0;missing_points=0;genuine_zero_points=0}
    $sheets=$null;$sheet=$null
    try{
        $labels=@('N','NbE','NNE','NEbN','NE','NEbE','ENE','EbN','E','EbS','ESE','SEbE','SE','SEbS','SSE','SbE','S','SbW','SSW','SWbS','SW','SWbW','WSW','WbS','W','WbN','WNW','NWbW','NW','NWbN','NNW','NbW')
        $days=@($Contract.logical_days);$matrix=@($Contract.expected_matrix);$included=@($Contract.included_indices)
        $wantedIncluded=if($Contract.port -ceq 'KLNG'){@(0..10)+@(22..31)}else{@(0..31)}
        if($days.Count -lt 1 -or $days.Count -gt 31 -or $matrix.Count -ne $days.Count -or
           ($Contract.direction_labels -join '|') -cne ($labels -join '|') -or ($included -join ',') -cne ($wantedIncluded -join ',')){throw 'Historical independent matrix/direction policy invalid'}
        $wantedOrder=if($Contract.port -ceq 'KLNG'){@(22..31)+@(0..10)}else{@(0..31)}
        if(($Contract.statistic_order -join ',') -cne ($wantedOrder -join ',') -or @($Contract.statistic_matrix).Count -ne $included.Count){throw 'Historical statistic category order invalid'}
        for($dayIndex=0;$dayIndex -lt $days.Count;$dayIndex++){
            if(@($matrix[$dayIndex]).Count -ne 32){throw 'Historical matrix does not have exactly32 directions'}
            $date=[DateTime]::ParseExact($days[$dayIndex],'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)
            if($dayIndex -gt 0 -and $days[$dayIndex] -cle $days[$dayIndex-1]){throw 'Historical logical days unsorted/duplicated'}
            for($direction=0;$direction -lt 32;$direction++){
                $value=$matrix[$dayIndex][$direction]
                if($direction -notin $included -and $null -ne $value){throw 'Historical excluded direction became data'}
                if($null -ne $value){Assert-NativeAuditNumericValue $value $value 'Historical independent matrix'}
            }
        }
        $sheets=$Book.Worksheets;$sheet=$sheets.Item('總表')
        for($direction=0;$direction -lt 32;$direction++){
            $header=Get-NativeAuditCellSnapshot $sheet ((Get-NativeAuditExcelColumn ($direction+2))+'1')
            if($header.has_formula -isnot [bool] -or $header.has_formula -or $header.value -cne $labels[$direction]){throw 'Historical native32 direction/header order drift'}
        }
        for($dayIndex=0;$dayIndex -lt $days.Count;$dayIndex++){
            $snapshot=Get-NativeAuditCellSnapshot $sheet ('A'+($dayIndex+2))
            if($snapshot.value -isnot [double] -or [DateTime]::FromOADate($snapshot.value).ToString('yyyy-MM-dd') -cne $days[$dayIndex]){throw 'Historical native date/row mapping drift'}
        }
        $cells=@($Contract.source_cells)
        if($cells.Count -eq 0 -or $Contract.verified_source_cell_count -ne $cells.Count -or @($cells|ForEach-Object {$_.sheet+'/'+$_.cell}|Sort-Object -Unique).Count -ne $cells.Count){throw 'Historical source-cell proof missing/duplicated'}
        foreach($source in $cells){
            if($source.verified -isnot [bool] -or -not $source.verified){throw 'Historical source-cell prerequisite not verified'}
            $sourceSheet=$null
            try{
                $sourceSheet=$sheets.Item($source.sheet);$snapshot=Get-NativeAuditCellSnapshot $sourceSheet $source.cell
                $hasFormula=$null -ne $source.formula
                if($snapshot.has_formula -isnot [bool] -or $snapshot.has_formula -ne $hasFormula -or
                   ($hasFormula -and (Get-NativeHistoricalFormulaCanonical $snapshot.formula) -cne (Get-NativeHistoricalFormulaCanonical $source.formula))){throw ('Historical native primary/source/statistic formula drift '+$source.sheet+'/'+$source.cell+' expected='+$source.formula+' actual='+$snapshot.formula)}
                Assert-NativeAuditPrimaryValue $snapshot.value $source.expected_value ($source.sheet+'/'+$source.cell)
                $result.source_observations += [pscustomobject]@{sheet=$source.sheet;cell=$source.cell;formula=$snapshot.formula;has_formula=$snapshot.has_formula;value=$snapshot.value}
                $result.source_cells++
            }finally{Release-NativeAuditCom $sourceSheet}
        }
        # Independently recompute each statistic from the period matrix before accepting chart helpers.
        for($offset=0;$offset -lt $wantedOrder.Count;$offset++){
            $direction=$wantedOrder[$offset];$numbers=@(for($row=0;$row -lt $days.Count;$row++){if($null -ne $matrix[$row][$direction]){[double]$matrix[$row][$direction]}})
            $wanted=@($null,$null,$null)
            if($numbers.Count -gt 0){$measure=$numbers|Measure-Object -Minimum -Maximum -Average;$wanted=@($measure.Maximum,$measure.Minimum,$measure.Average)}
            if(@($Contract.statistic_matrix[$offset]).Count -ne 3){throw 'Historical independent statistic vector dimension drift'}
            $category=Get-NativeAuditCellSnapshot $sheet ('AI'+($offset+2))
            if($category.value -cne $labels[$direction] -or $category.has_formula){throw 'Historical native rotated statistic category drift'}
            for($stat=0;$stat -lt 3;$stat++){
                Assert-NativeAuditPrimaryValue $Contract.statistic_matrix[$offset][$stat] $wanted[$stat] 'Historical independently recomputed statistic'
                $helper=Get-NativeAuditCellSnapshot $sheet ((@('AJ','AK','AL')[$stat])+($offset+2))
                if($helper.has_formula -isnot [bool] -or $helper.has_formula){throw 'Historical statistic helper unexpectedly became formula'}
                Assert-NativeAuditPrimaryValue $helper.value $wanted[$stat] 'Historical literal statistic helper'
            }
        }
        $formulaBacked=$Contract.role -in @('full_a','full_b','integrated')
        $helpers=@($Contract.radar_helpers)
        if($helpers.Count -ne $(if($formulaBacked){32*$days.Count}else{0})){throw 'Historical chart-only radar helper census invalid'}
        if($formulaBacked){
            $columns=$null;$columnRange=$null
            try{$columnRange=$sheet.Range('AO:BT');$columns=$columnRange.EntireColumn;if($columns.Hidden -isnot [bool] -or -not $columns.Hidden){throw 'Historical chart-only helpers are not hidden'}}finally{Release-NativeAuditCom $columns;Release-NativeAuditCom $columnRange}
            for($direction=0;$direction -lt 32;$direction++){
                $header=Get-NativeAuditCellSnapshot $sheet ((Get-NativeAuditExcelColumn ($direction+41))+'1')
                if($header.has_formula -or $header.value -cne $labels[$direction]){throw 'Historical chart-only helper header drift'}
            }
        }
        for($offset=0;$offset -lt $helpers.Count;$offset++){
            $dayIndex=[int][Math]::Floor($offset/32);$direction=$offset%32;$row=$dayIndex+2
            $address=(Get-NativeAuditExcelColumn ($direction+41))+$row;$sourceAddress=(Get-NativeAuditExcelColumn ($direction+2))+$row
            $expected=$matrix[$dayIndex][$direction];$missing=$null -eq $expected;$formula='=IF(COUNT('+$sourceAddress+')=0,NA(),'+$sourceAddress+')'
            $helper=$helpers[$offset]
            if($helper.cell -cne $address -or $helper.source -cne $sourceAddress -or $helper.formula -cne $formula -or
               $helper.expected_missing -isnot [bool] -or $helper.expected_missing -ne $missing -or $helper.verified -isnot [bool] -or -not $helper.verified){throw 'Historical independent helper scope/formula/prerequisite drift'}
            Assert-NativeAuditPrimaryValue $helper.expected_value $expected 'Historical independent helper expected value'
            $sourceSnapshot=Get-NativeAuditCellSnapshot $sheet $sourceAddress;$snapshot=Get-NativeAuditCellSnapshot $sheet $address
            $count=Assert-NativeAuditCount ($sheet.Evaluate('COUNT('+$sourceAddress+')')) 1
            if($snapshot.has_formula -isnot [bool] -or -not $snapshot.has_formula -or $snapshot.formula -cne $formula){throw 'Historical native chart helper exact formula drift'}
            Assert-NativeAuditPrimaryValue $sourceSnapshot.value $expected 'Historical native chart helper source'
            $isNA=$sheet.Evaluate('ISNA('+$address+')')
            if($isNA -isnot [bool]){throw 'Historical native ISNA getter unavailable'}
            if($missing){
                if($count -ne 0 -or -not $isNA -or $snapshot.value -ne -2146826246){throw 'Historical chart-onlyNA missing source proof failed'}
                $result.intentional_error_addresses += $address
            }else{
                if($count -ne 1 -or $isNA){throw 'Historical numeric source becameNA or COUNT0'}
                Assert-NativeAuditNumericValue $snapshot.value $expected 'Historical native chart helper numeric value'
            }
            $result.helper_observations += [pscustomobject]@{cell=$address;source=$sourceAddress;formula=$snapshot.formula;expected_missing=$missing;source_count=$count;source_value=$sourceSnapshot.value;helper_value=$snapshot.value}
            $result.radar_helper_cells++
        }
        if($result.intentional_error_addresses.Count -ne $Contract.radar_verified_intentional_missing_count){throw 'Historical intentionalNA census differs from independent matrix'}
        $objects=$null
        try{
            $objects=$sheet.ChartObjects();Assert-NativeAuditExpectedCount $objects.Count 2 'Historical native actual chart collection'
            for($chartIndex=1;$chartIndex -le 2;$chartIndex++){
                $object=$null;$chart=$null;$collection=$null;$plot=$null;$axis=$null
                try{
                    $object=$objects.Item($chartIndex);$chart=$object.Chart;$collection=$chart.FullSeriesCollection();$plot=$chart.PlotArea;$axis=$chart.Axes(2,1)
                    $plotGeometry=[pscustomobject]@{left=[double]$plot.InsideLeft;top=[double]$plot.InsideTop;width=[double]$plot.InsideWidth;height=[double]$plot.InsideHeight}
                    $axisMin=[double]$axis.MinimumScale;$axisMax=[double]$axis.MaximumScale
                    if($chart.DisplayBlanksAs -ne 1){throw 'Historical actual chart does not preserve gaps'}
                    if($chartIndex -eq 1 -and $formulaBacked -and $chart.PlotVisibleOnly){throw 'Historical native radar does not expose hidden chart helpers'}
                    $plan=if($chartIndex -eq 1){@($Contract.radar_series)}else{@($Contract.statistic_series)}
                    if($collection.Count -ne $plan.Count -or ($chartIndex -eq 2 -and $plan.Count -ne 3)){throw 'Historical native series count/order differs from independent plan'}
                    for($seriesIndex=1;$seriesIndex -le $plan.Count;$seriesIndex++){
                        $series=$null;$points=$null
                        try{
                            $wanted=$plan[$seriesIndex-1];$series=$collection.Item($seriesIndex);$points=$series.Points();$values=@($series.Values);$categories=@($series.XValues);$expected=@($wanted.expected_values)
                            $expectedCategories=if($chartIndex -eq 1){$labels}else{@($wantedOrder|ForEach-Object {$labels[$_]})}
                            $cat=[regex]::Replace($wanted.categories,'([A-Z]+)([0-9]+)','$$$1$$$2');$val=[regex]::Replace($wanted.values,'([A-Z]+)([0-9]+)','$$$1$$$2')
                            $formula='=SERIES("'+$wanted.name+'",'+$cat+','+$val+','+$seriesIndex+')'
                            if([string]$series.Name -cne $wanted.name -or [string]$series.Formula -cne $formula -or $values.Count -ne $expected.Count -or $points.Count -ne $expected.Count -or
                               ($categories -join '|') -cne ($expectedCategories -join '|')){throw 'Historical actual native series source/category/point mapping drift'}
                            # All five accepted copied June references use xlLine=4,
                            # not Modern's xlLineMarkers=65, for the unmarked mean.
                            if($chartIndex -eq 2 -and $series.ChartType -ne @([int]51,[int]51,[int]4)[$seriesIndex-1]){throw 'Historical statistic native series type/order drift'}
                            if($chartIndex -eq 1 -and $series.ChartType -notin @(81,82)){throw 'Historical native radar series type drift'}
                            $geometry=@();$missingPoints=0;$zeroPoints=0
                            for($pointIndex=1;$pointIndex -le $points.Count;$pointIndex++){
                                Assert-NativeAuditSeriesValue $values[$pointIndex-1] $expected[$pointIndex-1] 'Historical actual native series value'
                                $point=$null
                                try{$point=$points.Item($pointIndex);$snapshot=[pscustomobject]@{index=$pointIndex;left=[double]$point.Left;top=[double]$point.Top;width=[double]$point.Width;height=[double]$point.Height}}
                                finally{Release-NativeAuditCom $point}
                                Assert-NativeHistoricalPointGeometry $snapshot $plotGeometry $expected[$pointIndex-1] $chartIndex $seriesIndex $pointIndex $points.Count $axisMin $axisMax
                                $geometry += $snapshot;$result.series_points++
                                if($null -eq $expected[$pointIndex-1]){$missingPoints++;$result.missing_points++}
                                elseif($expected[$pointIndex-1] -eq 0){$zeroPoints++;$result.genuine_zero_points++}
                            }
                            $result.series_observations += [pscustomobject]@{chart_index=$chartIndex;series_index=$seriesIndex;name=$wanted.name;formula=[string]$series.Formula;
                                point_count=$points.Count;values=$values;categories=$categories;missing_native_null_points=$missingPoints;genuine_zero_points=$zeroPoints;
                                axis_minimum=$axisMin;axis_maximum=$axisMax;plot=$plotGeometry;points=$geometry}
                        }finally{Release-NativeAuditCom $points;Release-NativeAuditCom $series}
                    }
                }finally{foreach($item in @($axis,$plot,$collection,$chart,$object)){Release-NativeAuditCom $item}}
            }
        }finally{Release-NativeAuditCom $objects}
        $result.independent_contract_status='PASS';$result.status='PASS'
        $result.contract_report=$Contract.native_contract_report;$result.contract_report_sha256=$Contract.native_contract_report_sha256;$result.workbook_sha256=$Contract.workbook_sha256
    }catch{$result.error=$_.Exception.Message}
    finally{Release-NativeAuditCom $sheet;Release-NativeAuditCom $sheets}
    return [pscustomobject]$result
}
