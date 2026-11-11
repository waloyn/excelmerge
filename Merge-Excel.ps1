<#
.SYNOPSIS
    合并指定目录下所有表格到一个新工作簿。自动适配 Microsoft Excel、WPS 表格,
    都没有安装时回落到 ImportExcel 模块(无需安装 Office)。

.DESCRIPTION
    引擎选择顺序(-Engine Auto):
      1. Microsoft Excel   (COM: Excel.Application)
      2. WPS 表格          (COM: KET.Application / ET.Application)
      3. ImportExcel 模块  (基于 EPPlus,不依赖任何办公软件;仅支持 .xlsx/.xlsm)

    功能:
      - 指定表头行数:表头只保留第一个表格的,其余表格自动跳过前 N 行
      - 指定要合并的工作表名(支持多个、通配符);未指定时合并每个文件的第一个工作表

.PARAMETER Path
    待合并表格所在目录。

.PARAMETER OutputPath
    合并结果保存路径(.xlsx)。默认保存在源目录下,文件名带时间戳。

.PARAMETER HeaderRows
    每个表格的表头行数,默认 1。从工作表"已使用区域"的第一行开始计算。

.PARAMETER SheetName
    要合并的工作表名,支持多个和通配符,例如 "销售","Sheet*"。
    不指定时,只合并每个文件的第一个工作表。

.PARAMETER AllSheets
    合并每个文件中的所有工作表(优先级高于 -SheetName)。

.PARAMETER Recurse
    同时处理子目录中的表格。

.PARAMETER AddSourceColumns
    在结果最前面增加两列:"来源文件"、"来源工作表"。

.PARAMETER ValuesOnly
    仅 Excel/WPS 引擎有效:只合并单元格的值,不保留格式(更快)。

.PARAMETER Engine
    Auto(默认) | Excel | WPS | ImportExcel。强制指定引擎时,不可用会直接报错而不回落。

.PARAMETER AutoInstall
    回落到 ImportExcel 且未安装该模块时,自动安装到当前用户(否则交互式询问)。

.PARAMETER Extensions
    要处理的文件扩展名,默认 .xlsx .xlsm .xls。
    Excel/WPS 引擎可额外加入 ".csv"。ImportExcel 引擎只处理 .xlsx/.xlsm。

.EXAMPLE
    .\Merge-Excel.ps1 -Path "D:\data"

.EXAMPLE
    .\Merge-Excel.ps1 -Path "D:\data" -HeaderRows 2 -SheetName "销售明细" -AddSourceColumns

.EXAMPLE
    .\Merge-Excel.ps1 "D:\data" -SheetName "Q*","汇总" -Recurse -Engine ImportExcel -AutoInstall
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory, Position = 0)]
    [string]$Path,

    [string]$OutputPath,

    [ValidateRange(0, 1000)]
    [int]$HeaderRows = 1,

    [string[]]$SheetName,

    [switch]$AllSheets,

    [switch]$Recurse,

    [switch]$AddSourceColumns,

    [switch]$ValuesOnly,

    [ValidateSet('Auto', 'Excel', 'WPS', 'ImportExcel')]
    [string]$Engine = 'Auto',

    [switch]$AutoInstall,

    [string[]]$Extensions = @('.xlsx', '.xlsm', '.xls')
)

$ErrorActionPreference = 'Stop'
$MaxRows = 1048576
$srcOffset = if ($AddSourceColumns) { 2 } else { 0 }

# 合并过程的共享状态
$script:nextRow = 1
$script:headerWritten = $false
$script:baseCols = 0
$script:mergedSheets = 0

# ======================================================================
#  公共函数
# ======================================================================

# 判断工作表是否被选中
function Test-SheetSelected {
    param([string]$Name, [int]$Index)
    if ($AllSheets) { return $true }
    if ($SheetName) {
        foreach ($p in $SheetName) {
            if ($Name -eq $p) { return $true }
            try { if ($Name -like $p) { return $true } } catch { }
        }
        return $false
    }
    return ($Index -eq 1)
}

# 计算某个工作表的写入计划(跳过几行表头、写入几行);无内容可写时返回 $null
function Get-MergePlan {
    param([int]$Rows, [int]$Cols, [string]$Label)

    $skip = if ($script:headerWritten) { $HeaderRows } else { 0 }
    if ($skip -ge $Rows) {
        Write-Verbose "[$Label] 只有表头,没有数据行,已跳过"
        return $null
    }
    $nr = $Rows - $skip
    if ($script:nextRow + $nr - 1 -gt $MaxRows) {
        throw "合并后的行数超过 Excel 上限($MaxRows 行),已在 [$Label] 处停止。"
    }
    if (-not $script:headerWritten) {
        $script:baseCols = $Cols
    } elseif ($Cols -ne $script:baseCols) {
        Write-Warning "[$Label] 列数为 $Cols,与第一个表格的 $($script:baseCols) 列不一致,请检查数据是否对齐。"
    }
    return [pscustomobject]@{ Skip = $skip; NR = $nr }
}

# 本次写入中"数据行"(不含首个表格的表头行)的起止行号
function Get-DataRowRange {
    param($Plan)
    $start = $script:nextRow
    if (-not $script:headerWritten) { $start += $HeaderRows }
    $end = $script:nextRow + $Plan.NR - 1
    return [pscustomobject]@{ Start = $start; End = $end }
}

function Complete-Block {
    param($Plan, [string]$File, [string]$Sheet)
    $script:nextRow += $Plan.NR
    $script:headerWritten = $true
    $script:mergedSheets++
    Write-Host ("  + {0} [{1}]  {2} 行" -f $File, $Sheet, $Plan.NR)
}

# ======================================================================
#  引擎一:Excel / WPS(COM)
# ======================================================================

function New-SpreadsheetApp {
    param([string]$Engine)

    $candidates = switch ($Engine) {
        'Excel' { @('Excel.Application') }
        'WPS'   { @('KET.Application', 'ET.Application') }
        default { @('Excel.Application', 'KET.Application', 'ET.Application') }
    }

    foreach ($prog in $candidates) {
        $type = [Type]::GetTypeFromProgID($prog)
        if (-not $type) { continue }
        try {
            $app = [Activator]::CreateInstance($type)
        } catch {
            Write-Verbose "$prog 已注册但无法启动: $($_.Exception.Message)"
            continue
        }
        $label = $prog
        try { $label = "$($app.Name) $($app.Version)".Trim() } catch { }
        return [pscustomobject]@{ App = $app; Label = $label }
    }
    return $null
}

function Test-ComRangeEmpty {
    param($App, $Range)
    try {
        return ($App.WorksheetFunction.CountA($Range) -eq 0)
    } catch {
        try { return ($Range.Cells.Count -eq 1 -and [string]::IsNullOrEmpty([string]$Range.Cells.Item(1, 1).Value2)) }
        catch { return $false }
    }
}

function Merge-WithCom {
    param($Handle, $Files)

    $app = $Handle.App
    foreach ($prop in 'Visible', 'DisplayAlerts', 'ScreenUpdating') {
        try { $app.$prop = $false } catch { }
    }

    $outWb = $null
    $fileIndex = 0
    try {
        $outWb = $app.Workbooks.Add()
        $outWs = $outWb.Worksheets.Item(1)
        try { $outWs.Name = '合并结果' } catch { }

        foreach ($f in $Files) {
            $fileIndex++
            Write-Progress -Activity '正在合并表格' -Status $f.Name -PercentComplete ($fileIndex * 100 / $Files.Count)

            $wb = $null
            try {
                try { $wb = $app.Workbooks.Open($f.FullName, 0, $true) }   # 不更新链接、只读
                catch { $wb = $app.Workbooks.Open($f.FullName) }

                $targets = New-Object System.Collections.Generic.List[object]
                $allNames = @()
                $count = $wb.Worksheets.Count
                for ($i = 1; $i -le $count; $i++) {
                    $ws = $wb.Worksheets.Item($i)
                    $allNames += $ws.Name
                    if (Test-SheetSelected -Name $ws.Name -Index $i) { $targets.Add($ws) }
                }

                if ($targets.Count -eq 0) {
                    Write-Warning "[$($f.Name)] 未找到指定的工作表,已跳过。该文件的工作表: $($allNames -join ', ')"
                    continue
                }

                foreach ($ws in $targets) {
                    $used = $ws.UsedRange
                    if (Test-ComRangeEmpty -App $app -Range $used) {
                        Write-Verbose "[$($f.Name)/$($ws.Name)] 空工作表,跳过"
                        continue
                    }

                    $firstRow = $used.Row
                    $firstCol = $used.Column
                    $rows = $used.Rows.Count
                    $cols = $used.Columns.Count

                    $plan = Get-MergePlan -Rows $rows -Cols $cols -Label "$($f.Name)/$($ws.Name)"
                    if (-not $plan) { continue }

                    $src = $ws.Range(
                        $ws.Cells.Item($firstRow + $plan.Skip, $firstCol),
                        $ws.Cells.Item($firstRow + $rows - 1, $firstCol + $cols - 1))
                    $dst = $outWs.Range(
                        $outWs.Cells.Item($script:nextRow, 1 + $srcOffset),
                        $outWs.Cells.Item($script:nextRow + $plan.NR - 1, $srcOffset + $cols))

                    if ($ValuesOnly) {
                        $dst.Value2 = $src.Value2
                    } else {
                        [void]$src.Copy($dst)
                        $dst.Value2 = $dst.Value2   # 公式转为值,避免引用已关闭的源文件
                    }

                    if ($AddSourceColumns) {
                        if (-not $script:headerWritten -and $HeaderRows -ge 1) {
                            $outWs.Cells.Item($script:nextRow + $HeaderRows - 1, 1).Value2 = '来源文件'
                            $outWs.Cells.Item($script:nextRow + $HeaderRows - 1, 2).Value2 = '来源工作表'
                        }
                        $dr = Get-DataRowRange -Plan $plan
                        if ($dr.Start -le $dr.End) {
                            $outWs.Range($outWs.Cells.Item($dr.Start, 1), $outWs.Cells.Item($dr.End, 1)).Value2 = $f.Name
                            $outWs.Range($outWs.Cells.Item($dr.Start, 2), $outWs.Cells.Item($dr.End, 2)).Value2 = $ws.Name
                        }
                    }

                    Complete-Block -Plan $plan -File $f.Name -Sheet $ws.Name
                }
            }
            catch {
                if ($_.Exception.Message -like '*超过 Excel 上限*') { throw }
                Write-Warning "处理 [$($f.Name)] 时出错: $($_.Exception.Message)"
            }
            finally {
                try { $app.CutCopyMode = $false } catch { }
                if ($wb) {
                    try { $wb.Close($false) } catch { }
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
                }
            }
        }

        Write-Progress -Activity '正在合并表格' -Completed

        if (-not $script:headerWritten) {
            Write-Warning '没有合并到任何数据,未生成结果文件。'
            return
        }

        try { [void]$outWs.Columns.AutoFit() } catch { }
        try { $outWb.SaveAs($OutputPath, 51) }   # 51 = xlOpenXMLWorkbook (.xlsx)
        catch { $outWb.SaveAs($OutputPath) }
    }
    finally {
        if ($outWb) {
            try { $outWb.Close($false) } catch { }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outWb)
        }
        try { $app.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# ======================================================================
#  引擎二:ImportExcel 模块(EPPlus,无需安装 Office)
# ======================================================================

function Initialize-ImportExcel {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        $install = [bool]$AutoInstall
        if (-not $install -and [Environment]::UserInteractive -and -not [Console]::IsInputRedirected) {
            $ans = Read-Host '未检测到 Excel / WPS,也未安装 ImportExcel 模块。是否现在安装 ImportExcel 到当前用户? (Y/N)'
            $install = ($ans -match '^[Yy]')
        }
        if (-not $install) {
            throw "未找到可用引擎。请安装 Excel/WPS,或先执行: Install-Module ImportExcel -Scope CurrentUser   (也可加 -AutoInstall 参数自动安装)"
        }
        try {
            if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
                [void](Install-PackageProvider -Name NuGet -Scope CurrentUser -Force)
            }
            Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber
        } catch {
            throw "安装 ImportExcel 失败: $($_.Exception.Message)"
        }
    }
    Import-Module ImportExcel -ErrorAction Stop
    if (-not ('OfficeOpenXml.ExcelPackage' -as [type])) {
        throw 'ImportExcel 已加载,但找不到 EPPlus 组件(OfficeOpenXml.ExcelPackage)。'
    }
}

function Merge-WithImportExcel {
    param($Files)

    Initialize-ImportExcel

    $outPkg = New-Object OfficeOpenXml.ExcelPackage
    $fileIndex = 0
    try {
        $outWs = $outPkg.Workbook.Worksheets.Add('合并结果')

        foreach ($f in $Files) {
            $fileIndex++
            Write-Progress -Activity '正在合并表格' -Status $f.Name -PercentComplete ($fileIndex * 100 / $Files.Count)

            $stream = $null
            $pkg = $null
            try {
                # 允许读取正被 Excel/WPS 打开的文件
                $stream = [IO.File]::Open($f.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
                $pkg = New-Object OfficeOpenXml.ExcelPackage -ArgumentList $stream

                $targets = New-Object System.Collections.Generic.List[object]
                $allNames = @()
                $i = 0
                foreach ($ws in $pkg.Workbook.Worksheets) {
                    $i++
                    $allNames += $ws.Name
                    if (Test-SheetSelected -Name $ws.Name -Index $i) { $targets.Add($ws) }
                }

                if ($targets.Count -eq 0) {
                    Write-Warning "[$($f.Name)] 未找到指定的工作表,已跳过。该文件的工作表: $($allNames -join ', ')"
                    continue
                }

                $fmtCache = @{}   # 样式ID -> 数字格式,按工作簿缓存

                foreach ($ws in $targets) {
                    $dim = $ws.Dimension
                    if ($null -eq $dim) {
                        Write-Verbose "[$($f.Name)/$($ws.Name)] 空工作表,跳过"
                        continue
                    }

                    $sr = $dim.Start.Row
                    $sc = $dim.Start.Column
                    $rows = $dim.End.Row - $sr + 1
                    $cols = $dim.End.Column - $sc + 1

                    $plan = Get-MergePlan -Rows $rows -Cols $cols -Label "$($f.Name)/$($ws.Name)"
                    if (-not $plan) { continue }

                    $base = $script:nextRow
                    for ($r = 0; $r -lt $plan.NR; $r++) {
                        $srow = $sr + $plan.Skip + $r
                        $drow = $base + $r
                        for ($c = 0; $c -lt $cols; $c++) {
                            $scol = $sc + $c
                            $cell = $ws.Cells[$srow, $scol]
                            $v = $cell.Value
                            if ($null -eq $v) { continue }

                            $dcol = 1 + $srcOffset + $c
                            $outWs.SetValue($drow, $dcol, $v)

                            # 保留数字/日期格式
                            $sid = $cell.StyleID
                            if (-not $fmtCache.ContainsKey($sid)) {
                                $fmtCache[$sid] = [string]$cell.Style.Numberformat.Format
                            }
                            $fmt = $fmtCache[$sid]
                            if ($v -is [datetime] -and ([string]::IsNullOrEmpty($fmt) -or $fmt -eq 'General')) {
                                $fmt = 'yyyy-mm-dd hh:mm:ss'
                            }
                            if ($fmt -and $fmt -ne 'General') {
                                $outWs.Cells[$drow, $dcol].Style.Numberformat.Format = $fmt
                            }
                        }
                    }

                    if ($AddSourceColumns) {
                        if (-not $script:headerWritten -and $HeaderRows -ge 1) {
                            $outWs.SetValue(($base + $HeaderRows - 1), 1, '来源文件')
                            $outWs.SetValue(($base + $HeaderRows - 1), 2, '来源工作表')
                        }
                        $dr = Get-DataRowRange -Plan $plan
                        for ($row = $dr.Start; $row -le $dr.End; $row++) {
                            $outWs.SetValue($row, 1, $f.Name)
                            $outWs.SetValue($row, 2, $ws.Name)
                        }
                    }

                    Complete-Block -Plan $plan -File $f.Name -Sheet $ws.Name
                }
            }
            catch {
                if ($_.Exception.Message -like '*超过 Excel 上限*') { throw }
                Write-Warning "处理 [$($f.Name)] 时出错: $($_.Exception.Message)"
            }
            finally {
                if ($pkg) { $pkg.Dispose() }
                if ($stream) { $stream.Dispose() }
            }
        }

        Write-Progress -Activity '正在合并表格' -Completed

        if (-not $script:headerWritten) {
            Write-Warning '没有合并到任何数据,未生成结果文件。'
            return
        }

        if ($outWs.Dimension -and $outWs.Dimension.End.Row -le 100000) {
            try { $outWs.Cells[$outWs.Dimension.Address].AutoFitColumns() } catch { }
        }
        $outPkg.SaveAs((New-Object IO.FileInfo -ArgumentList $OutputPath))
    }
    finally {
        $outPkg.Dispose()
    }
}

# ======================================================================
#  主流程
# ======================================================================

$dir = (Resolve-Path -LiteralPath $Path).Path
if (-not (Test-Path -LiteralPath $dir -PathType Container)) {
    throw "指定路径不是目录: $dir"
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $dir ("合并结果_{0:yyyyMMdd_HHmmss}.xlsx" -f (Get-Date))
}
$OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
if ([IO.Path]::GetExtension($OutputPath).ToLower() -ne '.xlsx') { $OutputPath += '.xlsx' }

$exts = $Extensions | ForEach-Object { $_.ToLower() }
$files = @(Get-ChildItem -LiteralPath $dir -File -Recurse:$Recurse |
    Where-Object {
        $exts -contains $_.Extension.ToLower() -and
        $_.Name -notlike '~$*' -and
        $_.Name -notlike '.~*' -and
        $_.FullName -ne $OutputPath
    } | Sort-Object FullName)

if ($files.Count -eq 0) {
    Write-Warning "目录中没有找到可合并的表格文件: $dir"
    return
}

# 选择引擎
$handle = $null
if ($Engine -ne 'ImportExcel') {
    $handle = New-SpreadsheetApp -Engine $Engine
    if (-not $handle -and $Engine -ne 'Auto') {
        throw "指定的引擎 [$Engine] 不可用:未检测到对应软件,或 COM 组件未注册。"
    }
}

$sw = [Diagnostics.Stopwatch]::StartNew()

if ($handle) {
    Write-Host "使用引擎: $($handle.Label)(COM)" -ForegroundColor Cyan
    Write-Host "找到 $($files.Count) 个文件,表头行数 = $HeaderRows" -ForegroundColor Cyan
    Merge-WithCom -Handle $handle -Files $files
}
else {
    Write-Host '未检测到 Excel / WPS,使用 ImportExcel 模块处理' -ForegroundColor Yellow
    $supported = @($files | Where-Object { $_.Extension -match '^\.(xlsx|xlsm)$' })
    $skipped = $files.Count - $supported.Count
    if ($skipped -gt 0) {
        Write-Warning "ImportExcel 模式不支持 .xls / .csv 等格式,已忽略 $skipped 个文件。"
    }
    if ($supported.Count -eq 0) {
        Write-Warning '没有可处理的 .xlsx / .xlsm 文件。'
        return
    }
    Write-Host "找到 $($supported.Count) 个文件,表头行数 = $HeaderRows" -ForegroundColor Cyan
    Merge-WithImportExcel -Files $supported
}

$sw.Stop()
if ($script:headerWritten) {
    Write-Host ''
    Write-Host ("合并完成:共合并 {0} 个工作表,结果 {1} 行(含表头),用时 {2:N1} 秒。" -f $script:mergedSheets, ($script:nextRow - 1), $sw.Elapsed.TotalSeconds) -ForegroundColor Green
    Write-Host "输出文件: $OutputPath" -ForegroundColor Green
}
