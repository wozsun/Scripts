#requires -Version 7.0
<#
用途：
  使用本机 Office 原生应用，将指定文件或文件夹下的 .doc、.xls、.ppt 转换为 .docx、.xlsx、.pptx。

参数：
  -h     显示帮助信息。
  -s     包含隐藏文件和隐藏文件夹。
  Path   一个或多个文件或文件夹绝对路径；未提供时会引导交互输入。

关键规则：
  默认保留原始文件。
  新文件已存在时跳过，不覆盖。
  旧格式原始文件不会被删除、移动或覆盖。
#>

# ========== 参数区 ==========

[CmdletBinding()]
param(
    [Alias('h')]
    [switch]$Help,

    [Alias('s')]
    [switch]$IncludeHidden,

    [Parameter(Mandatory = $false, Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$PathList
)

# ========== 可调整配置 ==========

# 转换完成后统一移动文件前的单文件等待时间，给 OneDrive 等同步目录留出短暂缓冲。
$FileMoveDelayMilliseconds = 1000

# 输入根目录下的脚本专属临时目录名，避免误用用户原本已有的 tmp 文件夹。
$TempRootDirectoryName = '.convert-officefiles-tmp'

# 旧格式扩展名到 Office 应用、新格式扩展名和保存格式编号的映射。
# FileFormat 编号来自 Office COM 接口：Word 16=docx，Excel 51=xlsx，PowerPoint 24=pptx。
$OfficeFormatMap = @{
    '.doc' = [pscustomobject]@{
        AppName         = 'Word'
        TargetExtension = '.docx'
        FileFormat      = 16
    }
    '.xls' = [pscustomobject]@{
        AppName         = 'Excel'
        TargetExtension = '.xlsx'
        FileFormat      = 51
    }
    '.ppt' = [pscustomobject]@{
        AppName         = 'PowerPoint'
        TargetExtension = '.pptx'
        FileFormat      = 24
    }
}

# ========== 运行环境设置 ==========

Set-StrictMode -Version Latest

# 遇到未处理异常时立即进入 catch/退出流程，避免后续步骤继续处理不完整状态。
$ErrorActionPreference = 'Stop'

# 加载 Windows 脚本公共工具函数。
Import-Module -Name (Join-Path $PSScriptRoot 'common.psm1') -Force

# ========== 参数派生选项 ==========

# 是否扫描隐藏文件和隐藏文件夹，由 -s 参数决定。
$ShouldIncludeHiddenItems = [bool]$IncludeHidden

# 输出脚本用途、参数和关键安全规则。
function Show-HelpText {
    Write-Host @'
用途：
  使用本机 Office 原生应用，将指定文件或文件夹下的 .doc、.xls、.ppt 转换为 .docx、.xlsx、.pptx。

用法：
  pwsh -File .\Convert-OfficeFiles.ps1 [-s] [Path1] [Path2 ...]

参数：
  Path
    一个或多个文件或文件夹绝对路径。文件路径会直接转换；文件夹路径会递归扫描。
    未提供时会引导交互输入，交互时可在同一行输入多个路径；路径含空格请使用英文引号。

  -s
    包含隐藏文件和隐藏文件夹。默认只扫描未隐藏项。

规则：
  默认保留原始 .doc、.xls、.ppt 文件。
  如果目标 .docx、.xlsx、.pptx 已存在，则跳过，不覆盖。
  脚本扫描后直接转换，不需要预览确认。
  扫描前会检查输入目录或直接文件所在目录下的专属临时文件夹；如已有内容，会等待用户清理。
  脚本会先转换到该专属临时文件夹，退出 Office 后再统一移动到目标位置。
  每个文件单独转换，单个文件失败时记录错误并继续处理后续文件。
'@
}

# 校验输入路径：仅接受 Windows 绝对路径，并确保最终指向单个文件或文件夹。
function Resolve-InputPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathText,

        [Parameter(Mandatory = $false)]
        [string]$ParameterName = 'Path'
    )

    $normalizedPathText = ConvertTo-UnquotedPathText -PathText $PathText
    $isAbsolutePath = $normalizedPathText -match '^[a-zA-Z]:[\\/]' -or $normalizedPathText -match '^[\\/]{2}'
    if (-not $isAbsolutePath) {
        throw "$ParameterName 必须是 Windows 文件或文件夹绝对路径。"
    }

    $resolvedPaths = @(Resolve-Path -LiteralPath $normalizedPathText -ErrorAction Stop)
    if ($resolvedPaths.Count -ne 1) {
        throw "$ParameterName 必须只能解析到一个文件或文件夹。"
    }

    return (Get-Item -LiteralPath $resolvedPaths[0].ProviderPath -ErrorAction Stop)
}

# 逐个校验输入路径，并返回文件系统对象。
function Resolve-InputPathList {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$InputPathList
    )

    return @(
        for ($index = 0; $index -lt $InputPathList.Count; $index++) {
            Resolve-InputPath -PathText $InputPathList[$index] -ParameterName "Path$($index + 1)"
        }
    )
}

# 递归扫描旧格式 Office 文件；扫描错误只记录，不中断整个脚本。
function Get-LegacyOfficeFileList {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath,

        [Parameter(Mandatory = $false)]
        [System.IO.FileInfo[]]$SourceFileList
    )

    if ($null -ne $SourceFileList) {
        Write-StageMessage "开始处理直接输入文件，文件数: $($SourceFileList.Count)"
        $files = @($SourceFileList)
    }
    else {
        Write-StageMessage "开始扫描目录: $RootPath"
        $scanErrors = $null
        if ($ShouldIncludeHiddenItems) {
            $files = @(Get-ChildItem -LiteralPath $RootPath -File -Recurse -Force -ErrorAction SilentlyContinue -ErrorVariable scanErrors)
        }
        else {
            $files = @(Get-ChildItem -LiteralPath $RootPath -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable scanErrors)
        }

        $tempRootDirectory = Get-TempRootDirectory -RootPath $RootPath
        $separatorCharacters = [char[]]@([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
        $tempRootPrefix = $tempRootDirectory.TrimEnd($separatorCharacters) + [System.IO.Path]::DirectorySeparatorChar
        $files = @(
            $files | Where-Object {
                -not $_.FullName.StartsWith($tempRootPrefix, [System.StringComparison]::OrdinalIgnoreCase)
            }
        )

        $scanWarningList = New-DeferredScanWarningList
        foreach ($scanError in @($scanErrors)) {
            Add-DeferredScanWarning `
                -WarningList $scanWarningList `
                -Message '扫描跳过' `
                -Path $scanError.TargetObject `
                -Reason $scanError.Exception.Message
        }
        Write-DeferredScanWarningList -WarningList $scanWarningList -Title 'Office 扫描跳过汇总'
    }

    $legacyFiles = @(
        $files | Where-Object {
            $OfficeFormatMap.ContainsKey($_.Extension.ToLowerInvariant())
        }
    )

    $hiddenModeText = if ($ShouldIncludeHiddenItems) { '包含隐藏项' } else { '不包含隐藏项' }
    if ($null -ne $SourceFileList) {
        Write-StageMessage "直接输入文件检查完成，文件数: $($files.Count)，旧格式 Office 文件: $($legacyFiles.Count)"
    }
    else {
        Write-StageMessage "扫描完成，文件数: $($files.Count)，旧格式 Office 文件: $($legacyFiles.Count)，$hiddenModeText"
    }

    return $legacyFiles
}

# 根据扫描结果生成转换计划；目标文件已存在时标记为跳过，避免覆盖。
function New-ConversionPlanList {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath,

        [Parameter(Mandatory = $false)]
        [System.IO.FileInfo[]]$SourceFileList,

        [Parameter(Mandatory = $false)]
        [System.Collections.Generic.HashSet[string]]$ProcessedSourcePathSet
    )

    $legacyFiles = @(Get-LegacyOfficeFileList -RootPath $RootPath -SourceFileList $SourceFileList)
    $plans = [System.Collections.Generic.List[object]]::new()
    $processedCount = 0
    $lastPercent = -1

    foreach ($file in $legacyFiles) {
        $processedCount++
        Write-ProgressBar -Activity '转换计划生成' -Status '正在检查目标文件' -ProcessedCount $processedCount -TotalCount $legacyFiles.Count -LastPercent ([ref]$lastPercent)

        if ($null -ne $ProcessedSourcePathSet -and -not $ProcessedSourcePathSet.Add($file.FullName)) {
            continue
        }

        $format = $OfficeFormatMap[$file.Extension.ToLowerInvariant()]
        $targetPath = [System.IO.Path]::ChangeExtension($file.FullName, $format.TargetExtension)
        $status = if (Test-Path -LiteralPath $targetPath) { 'SkipExists' } else { 'Convert' }

        $plans.Add([pscustomobject]@{
                AppName        = $format.AppName
                SourceFile     = $file
                SourcePathText = Get-RelativePathText -RootPath $RootPath -FilePath $file.FullName
                TargetPath     = $targetPath
                TargetPathText = Get-RelativePathText -RootPath $RootPath -FilePath $targetPath
                FileFormat     = $format.FileFormat
                Status         = $status
            })
    }

    Complete-DynamicStatusLine
    return $plans.ToArray()
}

# 将输入文件/目录整理为转换作用域；直接文件按父目录分组，共用同一个临时目录。
function New-ConversionScopeList {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileSystemInfo[]]$InputItemList
    )

    $scopeList = [System.Collections.Generic.List[object]]::new()
    $directoryRootSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $directFileScopeByRoot = @{}

    foreach ($inputItem in $InputItemList) {
        if ($inputItem -is [System.IO.DirectoryInfo]) {
            if ($directoryRootSet.Add($inputItem.FullName)) {
                $scopeList.Add([pscustomobject]@{
                        RootPath       = $inputItem.FullName
                        SourceFileList = $null
                        Label          = "目录: $($inputItem.FullName)"
                    })
            }
            continue
        }

        if (-not $OfficeFormatMap.ContainsKey($inputItem.Extension.ToLowerInvariant())) {
            Write-Host "跳过不支持的文件: $($inputItem.FullName)" -ForegroundColor Yellow
            continue
        }

        $rootPath = Split-Path -Parent $inputItem.FullName
        if (-not $directFileScopeByRoot.ContainsKey($rootPath)) {
            $directFileScopeByRoot[$rootPath] = [pscustomobject]@{
                RootPath       = $rootPath
                SourceFileList = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
                Label          = "直接文件所在目录: $rootPath"
            }
            $scopeList.Add($directFileScopeByRoot[$rootPath])
        }

        $directFileScopeByRoot[$rootPath].SourceFileList.Add([System.IO.FileInfo]$inputItem)
    }

    return $scopeList.ToArray()
}

# 输出转换计划摘要。新建文件操作不需要预览确认，但仍给出数量概览。
function Write-ConversionPlanSummary {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$ConversionPlans
    )

    $convertPlans = @($ConversionPlans | Where-Object { $_.Status -eq 'Convert' })
    $skipPlans = @($ConversionPlans | Where-Object { $_.Status -eq 'SkipExists' })

    Write-Host "`n转换计划:" -ForegroundColor Cyan
    Write-Host "待转换文件: $($convertPlans.Count)，目标已存在跳过: $($skipPlans.Count)" -ForegroundColor Cyan

    return $convertPlans.Count
}

# 获取输入根目录下的脚本专属临时目录。
function Get-TempRootDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    return [System.IO.Path]::Combine($RootPath, $TempRootDirectoryName)
}

# 确保脚本专属临时目录存在，并避免误用同名文件。
function New-TempRootDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    $tempRootDirectory = Get-TempRootDirectory -RootPath $RootPath

    if (Test-Path -LiteralPath $tempRootDirectory -PathType Leaf) {
        throw "临时目录路径被同名文件占用: $tempRootDirectory"
    }

    New-Item -ItemType Directory -Path $tempRootDirectory -Force -ErrorAction Stop | Out-Null
    $createdDirectory = Get-Item -LiteralPath $tempRootDirectory -ErrorAction Stop
    if (-not $createdDirectory.PSIsContainer) {
        throw "临时路径不是文件夹: $tempRootDirectory"
    }

    return $createdDirectory.FullName
}

# 获取目录中的项目数量，扫描隐藏项以确保临时目录真正为空。
function Get-DirectoryItemCount {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath,

        [Parameter(Mandatory = $false)]
        [switch]$TreatErrorAsNonEmpty
    )

    try {
        return @(Get-ChildItem -LiteralPath $DirectoryPath -Force -ErrorAction Stop).Count
    }
    catch {
        if ($TreatErrorAsNonEmpty) {
            return 1
        }

        throw
    }
}

# 扫描前准备临时目录；如果已有内容，暂停等待用户清理后再继续。
function Initialize-TempRootDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    $tempRootDirectory = Get-TempRootDirectory -RootPath $RootPath
    if (Test-Path -LiteralPath $tempRootDirectory -PathType Leaf) {
        throw "临时目录路径被同名文件占用: $tempRootDirectory"
    }

    if (-not (Test-Path -LiteralPath $tempRootDirectory)) {
        return (New-TempRootDirectory -RootPath $RootPath)
    }

    $tempRootItem = Get-Item -LiteralPath $tempRootDirectory -ErrorAction Stop
    if (-not $tempRootItem.PSIsContainer) {
        throw "临时路径不是文件夹: $tempRootDirectory"
    }

    while (Get-DirectoryItemCount -DirectoryPath $tempRootDirectory) {
        Write-Host ""
        Write-Host "临时目录已存在且不为空: $tempRootDirectory" -ForegroundColor Yellow
        Write-Host "请先处理或清空此文件夹中的内容；处理完成后按 Enter 继续。" -ForegroundColor Yellow
        Write-Host "如需退出脚本，请按 Ctrl+C。" -ForegroundColor DarkGray
        $continueInput = Read-ColoredLine -Prompt "等待处理完成: "
        if ($null -eq $continueInput) {
            throw "输入流已结束，临时目录仍非空: $tempRootDirectory"
        }

        if (-not (Test-Path -LiteralPath $tempRootDirectory)) {
            return (New-TempRootDirectory -RootPath $RootPath)
        }

        if (Test-Path -LiteralPath $tempRootDirectory -PathType Leaf) {
            throw "临时目录路径被同名文件占用: $tempRootDirectory"
        }
    }

    return $tempRootDirectory
}

# 尝试删除空的脚本专属临时目录；非空时保留，避免误删用户仍需处理的内容。
function Remove-TempRootDirectoryIfEmpty {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    $tempRootDirectory = Get-TempRootDirectory -RootPath $RootPath
    if (-not (Test-Path -LiteralPath $tempRootDirectory)) {
        return $true
    }

    if (Test-Path -LiteralPath $tempRootDirectory -PathType Leaf) {
        Write-Host "临时目录清理跳过: 路径被同名文件占用: $tempRootDirectory" -ForegroundColor Yellow
        return $false
    }

    if (Get-DirectoryItemCount -DirectoryPath $tempRootDirectory) {
        Write-Host "临时目录未清理: 目录仍有内容: $tempRootDirectory" -ForegroundColor Yellow
        return $false
    }

    Remove-Item -LiteralPath $tempRootDirectory -Force -ErrorAction SilentlyContinue
    return -not (Test-Path -LiteralPath $tempRootDirectory)
}

# 先保存到脚本专属临时目录，确认生成成功后再移动为目标文件，减少半成品残留风险。
function New-TempOutputPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempRootDirectory,

        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    if (-not (Test-Path -LiteralPath $TempRootDirectory -PathType Container)) {
        throw "临时目录不存在或不可用: $TempRootDirectory"
    }

    $fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
    $extension = [System.IO.Path]::GetExtension($TargetPath)
    return [System.IO.Path]::Combine($TempRootDirectory, "$fileNameWithoutExtension.tmp-$([guid]::NewGuid().ToString('N'))$extension")
}

# 尝试清理转换临时文件；OneDrive 目录可能短暂占用文件，因此失败时做少量重试。
function Remove-TempOutputFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempOutputPath
    )

    if (-not (Test-Path -LiteralPath $TempOutputPath)) {
        return $true
    }

    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            Remove-Item -LiteralPath $TempOutputPath -Force -ErrorAction Stop
            return $true
        }
        catch {
            Start-Sleep -Milliseconds (200 * $attempt)
        }
    }

    return -not (Test-Path -LiteralPath $TempOutputPath)
}

# 根据临时输出文件路径清理其所在的空临时目录；如果里面还有文件则保留，避免误删用户内容。
function Remove-TempOutputDirectoryIfEmpty {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempOutputPath
    )

    $tempDirectory = [System.IO.Path]::GetDirectoryName($TempOutputPath)
    if ([string]::IsNullOrWhiteSpace($tempDirectory) -or -not (Test-Path -LiteralPath $tempDirectory)) {
        return
    }

    if ([System.IO.Path]::GetFileName($tempDirectory) -ne $TempRootDirectoryName) {
        return
    }

    if ((Get-DirectoryItemCount -DirectoryPath $tempDirectory -TreatErrorAsNonEmpty) -eq 0) {
        Remove-Item -LiteralPath $tempDirectory -Force -ErrorAction SilentlyContinue
    }
}

# 获取临时文件移动结果，确保目标文件存在，并尽量清理移动后仍残留的临时文件。
function Get-TempOutputMoveResult {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempOutputPath,

        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    if (-not (Test-Path -LiteralPath $TargetPath)) {
        return [pscustomobject]@{
            IsValid     = $false
            HasWarning  = $false
            Message     = '目标文件未生成'
            CleanupPath = $null
        }
    }

    if (Test-Path -LiteralPath $TempOutputPath) {
        if (-not (Remove-TempOutputFile -TempOutputPath $TempOutputPath)) {
            return [pscustomobject]@{
                IsValid     = $true
                HasWarning  = $true
                Message     = '目标文件已生成，但临时文件清理失败'
                CleanupPath = $TempOutputPath
            }
        }
    }

    return [pscustomobject]@{
        IsValid     = $true
        HasWarning  = $false
        Message     = '移动结果正常'
        CleanupPath = $null
    }
}

# 释放 Office COM 对象，降低后台残留 Office 进程的概率。
function Remove-ComObject {
    param(
        [Parameter(Mandatory = $false)]
        [object]$ComObject
    )

    if ($null -ne $ComObject) {
        while ([System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) -gt 0) {}
    }
}

# 按需启动并缓存 Office 应用实例，避免每个文件都重复拉起 Office。
function Get-OfficeApplication {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppName,

        [Parameter(Mandatory = $true)]
        [hashtable]$ApplicationCache
    )

    if ($ApplicationCache.ContainsKey($AppName)) {
        return $ApplicationCache[$AppName]
    }

    $application = switch ($AppName) {
        'Word' {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $word.DisplayAlerts = 0
            $word
        }
        'Excel' {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $excel
        }
        'PowerPoint' {
            $powerPoint = New-Object -ComObject PowerPoint.Application
            $powerPoint.DisplayAlerts = 1
            $powerPoint
        }
    }

    $ApplicationCache[$AppName] = $application
    return $application
}

# 使用对应 Office 应用转换单个文件到专属临时目录；单文件失败时返回失败结果，由调用方继续处理后续文件。
function Convert-OfficeFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempRootDirectory,

        [Parameter(Mandatory = $true)]
        [object]$Plan,

        [Parameter(Mandatory = $true)]
        [hashtable]$ApplicationCache
    )

    $tempOutputPath = $null
    $document = $null

    try {
        if (Test-Path -LiteralPath $Plan.TargetPath) {
            return [pscustomobject]@{ Status = 'Skipped'; Message = '目标文件已存在' }
        }

        $tempOutputPath = New-TempOutputPath -TempRootDirectory $TempRootDirectory -TargetPath $Plan.TargetPath
        $application = Get-OfficeApplication -AppName $Plan.AppName -ApplicationCache $ApplicationCache
        switch ($Plan.AppName) {
            'Word' {
                $document = $application.Documents.Open($Plan.SourceFile.FullName, $false, $true, $false)
                $document.SaveAs2($tempOutputPath, $Plan.FileFormat)
                $document.Close($false)
            }
            'Excel' {
                $document = $application.Workbooks.Open($Plan.SourceFile.FullName, 0, $true)
                $document.SaveAs($tempOutputPath, $Plan.FileFormat)
                $document.Close($false)
            }
            'PowerPoint' {
                $document = $application.Presentations.Open($Plan.SourceFile.FullName, $true, $false, $false)
                $document.SaveAs($tempOutputPath, $Plan.FileFormat)
                $document.Close()
            }
        }

        if (-not (Test-Path -LiteralPath $tempOutputPath)) {
            throw 'Office 未生成临时输出文件。'
        }

        return [pscustomobject]@{ Status = 'ConvertedToTemp'; Message = '临时文件转换成功'; TempOutputPath = $tempOutputPath }
    }
    catch {
        if ($null -ne $document) {
            try {
                if ($Plan.AppName -eq 'PowerPoint') {
                    $document.Close()
                }
                else {
                    $document.Close($false)
                }
            }
            catch {
                Write-Debug "关闭 Office 文档失败: $($_.Exception.Message)"
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($tempOutputPath)) {
            [void](Remove-TempOutputFile -TempOutputPath $tempOutputPath)
            Remove-TempOutputDirectoryIfEmpty -TempOutputPath $tempOutputPath
        }

        return [pscustomobject]@{ Status = 'Failed'; Message = $_.Exception.Message }
    }
    finally {
        Remove-ComObject -ComObject $document
        $document = $null
    }
}

# 统一退出已启动的 Office 应用，并触发垃圾回收清理 COM 引用。
function Close-OfficeApplicationCache {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$ApplicationCache
    )

    foreach ($appName in @($ApplicationCache.Keys)) {
        $application = $ApplicationCache[$appName]
        try {
            $application.Quit()
        }
        catch {
            Write-Debug "退出 Office 应用失败: $($_.Exception.Message)"
        }
        finally {
            Remove-ComObject -ComObject $application
        }
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# 在 Office 应用退出后统一移动临时文件到目标位置，并用数量进度条展示移动进度。
function Move-ConvertedOfficeFileList {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempRootDirectory,

        [Parameter(Mandatory = $true)]
        [object[]]$TempConversionResults
    )

    $movedCount = 0
    $skippedCount = 0
    $failedCount = 0
    $warningMessages = [System.Collections.Generic.List[string]]::new()
    $failureMessages = [System.Collections.Generic.List[string]]::new()
    $processedCount = 0
    $lastPercent = -1

    if ($TempConversionResults.Count -eq 0) {
        return [pscustomobject]@{
            MovedCount   = 0
            SkippedCount = 0
            FailedCount  = 0
        }
    }

    Write-StageMessage "开始移动转换文件，待移动: $($TempConversionResults.Count)，单文件延时: ${FileMoveDelayMilliseconds}ms"

    foreach ($conversionResult in $TempConversionResults) {
        $processedCount++
        $plan = $conversionResult.Plan
        $tempOutputPath = $conversionResult.TempOutputPath

        Write-ProgressBar -Activity '移动转换文件' -Status '正在移动文件' -ProcessedCount $processedCount -TotalCount $TempConversionResults.Count -LastPercent ([ref]$lastPercent)
        Start-Sleep -Milliseconds $FileMoveDelayMilliseconds

        if (Test-Path -LiteralPath $plan.TargetPath) {
            if (Remove-TempOutputFile -TempOutputPath $tempOutputPath) {
                Remove-TempOutputDirectoryIfEmpty -TempOutputPath $tempOutputPath
                $warningMessages.Add("跳过移动: $($plan.TargetPathText)`n  原因: 目标文件已存在，临时文件已清理")
            }
            else {
                $warningMessages.Add("跳过移动: $($plan.TargetPathText)`n  原因: 目标文件已存在，但临时文件清理失败`n  临时文件: $tempOutputPath")
            }

            $skippedCount++
            continue
        }

        try {
            Move-Item -LiteralPath $tempOutputPath -Destination $plan.TargetPath -ErrorAction Stop
            $moveResult = Get-TempOutputMoveResult -TempOutputPath $tempOutputPath -TargetPath $plan.TargetPath
            if (-not $moveResult.IsValid) {
                throw $moveResult.Message
            }

            if ($moveResult.HasWarning) {
                $warningMessages.Add("移动完成但临时文件清理失败: $($plan.TargetPathText)`n  临时文件: $($moveResult.CleanupPath)")
            }

            Remove-TempOutputDirectoryIfEmpty -TempOutputPath $tempOutputPath
            $movedCount++
        }
        catch {
            $failedCount++
            $failureMessages.Add("移动失败: $($plan.TargetPathText)`n  原因: $($_.Exception.Message)`n  临时文件: $tempOutputPath")
        }
    }

    Complete-DynamicStatusLine

    foreach ($warningMessage in $warningMessages) {
        Write-Host $warningMessage -ForegroundColor Yellow
    }

    foreach ($failureMessage in $failureMessages) {
        Write-Host $failureMessage -ForegroundColor Red
    }

    if (-not (Test-Path -LiteralPath $TempRootDirectory)) {
        Write-Host "临时文件夹已清理。" -ForegroundColor Green
    }

    return [pscustomobject]@{
        MovedCount   = $movedCount
        SkippedCount = $skippedCount
        FailedCount  = $failedCount
    }
}

# 执行转换计划：先输出跳过项，再逐个转换待处理文件并汇总结果。
function Invoke-ConversionPlanList {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempRootDirectory,

        [Parameter(Mandatory = $true)]
        [object[]]$ConversionPlans
    )

    $convertPlans = @($ConversionPlans | Where-Object { $_.Status -eq 'Convert' })
    $skipPlans = @($ConversionPlans | Where-Object { $_.Status -eq 'SkipExists' })
    $applicationCache = @{}
    $tempConversionResults = [System.Collections.Generic.List[object]]::new()
    $movedCount = 0
    $skippedCount = 0
    $failedCount = 0

    foreach ($plan in $skipPlans) {
        Write-Host "跳过: $($plan.TargetPathText)" -ForegroundColor Yellow
        Write-Host "  原因: 目标文件已存在" -ForegroundColor DarkGray
        Write-Host "  来源: $($plan.SourcePathText)" -ForegroundColor DarkGray
        $skippedCount++
    }

    if ($convertPlans.Count -gt 0) {
        Write-StageMessage "开始转换 Office 文件，待转换: $($convertPlans.Count)"
    }

    try {
        foreach ($plan in $convertPlans) {
            $workingMessage = "正在转换: $($plan.SourcePathText)"
            [void](Write-RefreshStatusLine -Message $workingMessage -Color White -NoNewLine)
            $result = Convert-OfficeFile -TempRootDirectory $TempRootDirectory -Plan $plan -ApplicationCache $applicationCache
            switch ($result.Status) {
                'ConvertedToTemp' {
                    [void](Write-RefreshStatusLine -Message "转换完成: $($plan.SourcePathText)" -Color Green)
                    $tempConversionResults.Add([pscustomobject]@{
                            Plan           = $plan
                            TempOutputPath = $result.TempOutputPath
                        })
                }
                'Skipped' {
                    [void](Write-RefreshStatusLine -Message "跳过: $($plan.TargetPathText)" -Color Yellow)
                    Write-Host "  原因: $($result.Message)" -ForegroundColor DarkGray
                    $skippedCount++
                }
                'Failed' {
                    [void](Write-RefreshStatusLine -Message "转换失败: $($plan.SourcePathText)" -Color Red)
                    Write-Host "  原因: $($result.Message)" -ForegroundColor DarkGray
                    $failedCount++
                }
            }
        }
    }
    finally {
        if ($applicationCache.Count -gt 0) {
            [void](Write-RefreshStatusLine -Message '正在退出 Office 应用...' -Color White -NoNewLine)
        }

        Close-OfficeApplicationCache -ApplicationCache $applicationCache

        if ($applicationCache.Count -gt 0) {
            [void](Write-RefreshStatusLine -Message 'Office 应用已退出' -Color Green)
        }
    }

    if ($tempConversionResults.Count -gt 0) {
        $moveSummary = Move-ConvertedOfficeFileList -TempRootDirectory $TempRootDirectory -TempConversionResults $tempConversionResults.ToArray()
        $movedCount += $moveSummary.MovedCount
        $skippedCount += $moveSummary.SkippedCount
        $failedCount += $moveSummary.FailedCount
    }

    Write-Host ""
    Write-Host "转换完成。成功: $movedCount，跳过: $skippedCount，失败: $failedCount" -ForegroundColor Green
    if ($failedCount -gt 0) {
        Write-Host "存在转换失败文件，请查看上方失败原因。" -ForegroundColor Red
    }
}

if ($Help) {
    # 用户显式请求帮助时只输出帮助信息，不进入扫描流程。
    Show-HelpText
    exit 0
}

if ($null -eq $PathList -or $PathList.Count -eq 0) {
    # 未传入路径时，引导用户输入文件或文件夹，直接回车则安全退出。
    Write-Host "请输入文件或目录绝对路径。可在同一行输入多个路径。" -ForegroundColor Cyan
    Write-Host "多个路径可用空格或英文分号分隔；路径含空格请使用英文引号。" -ForegroundColor DarkGray
    Write-Host "直接回车退出；输入 0 退出脚本。" -ForegroundColor DarkGray
    $pathInputRaw = Read-ColoredLine -Prompt 'Path: '
    if ($null -eq $pathInputRaw) {
        Write-Host "已退出，未执行扫描。" -ForegroundColor Yellow
        exit 0
    }

    $pathInput = $pathInputRaw.Trim()
    if ([string]::IsNullOrWhiteSpace($pathInput)) {
        Write-Host "已退出，未执行扫描。" -ForegroundColor Yellow
        exit 0
    }

    if ($pathInput -eq '0') {
        Write-Host "已退出，未执行扫描。" -ForegroundColor Yellow
        exit 0
    }

    $PathList = @(Split-InteractivePathInput -PathInput $pathInput)
    if ($PathList.Count -gt 1) {
        Write-Host "识别到 $($PathList.Count) 个路径。" -ForegroundColor DarkGray
    }
}

try {
    # 在正式扫描前集中校验路径，避免后续函数反复处理无效输入。
    $effectivePathList = @($PathList | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($effectivePathList.Count -eq 0) {
        Write-Host "未提供有效路径，已退出。" -ForegroundColor Yellow
        exit 0
    }

    $resolvedInputItemList = @(Resolve-InputPathList -InputPathList $effectivePathList)
    $conversionScopeList = @(New-ConversionScopeList -InputItemList $resolvedInputItemList)
}
catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

if ($conversionScopeList.Count -eq 0) {
    Write-Host "未发现可处理的 .doc、.xls、.ppt 文件或目录。" -ForegroundColor Green
    exit 0
}

$processedSourcePathSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$scopeIndex = 0
foreach ($conversionScope in $conversionScopeList) {
    $scopeIndex++
    if ($conversionScopeList.Count -gt 1) {
        Write-Host ""
        Write-Host "转换范围 $scopeIndex / $($conversionScopeList.Count): $($conversionScope.Label)" -ForegroundColor Cyan
    }

    try {
        $tempRootDirectory = Initialize-TempRootDirectory -RootPath $conversionScope.RootPath
    }
    catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        exit 1
    }

    try {
        $sourceFileList = if ($null -eq $conversionScope.SourceFileList) { $null } else { $conversionScope.SourceFileList.ToArray() }
        $conversionPlans = @(New-ConversionPlanList -RootPath $conversionScope.RootPath -SourceFileList $sourceFileList -ProcessedSourcePathSet $processedSourcePathSet)
        if ($conversionPlans.Count -eq 0) {
            # 没有可处理文件时直接成功结束，但仍会在 finally 中清理空临时目录。
            Write-Host "未发现 .doc、.xls、.ppt 旧格式 Office 文件。" -ForegroundColor Green
        }
        else {
            [void](Write-ConversionPlanSummary -ConversionPlans $conversionPlans)
            Invoke-ConversionPlanList -TempRootDirectory $tempRootDirectory -ConversionPlans $conversionPlans
        }
    }
    finally {
        [void](Remove-TempRootDirectoryIfEmpty -RootPath $conversionScope.RootPath)
    }
}
