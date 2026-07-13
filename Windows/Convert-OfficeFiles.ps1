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

# 转换后移动临时文件的最大尝试次数；仅在文件被 Office、OneDrive 等短暂占用时重试。
$FileMoveRetryCount = 5

# 移动失败后的基础退避时间；第 N 次失败后等待 N 倍该值。
$FileMoveRetryBaseDelayMilliseconds = 200

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
    命令行传参时，路径包含空格、括号等 PowerShell 特殊字符，请使用英文引号包裹路径。
    未提供时会引导交互输入，交互时可在同一行输入多个路径；交互输入中路径含空格或英文分号时，也请使用英文引号包裹路径。

  -s
    包含隐藏文件和隐藏文件夹。默认只扫描未隐藏项。

规则：
  默认保留原始 .doc、.xls、.ppt 文件。
  如果目标 .docx、.xlsx、.pptx 已存在，则跳过，不覆盖。
  脚本扫描后直接转换，不需要预览确认。
  存在待转换文件时会创建输入目录或直接文件所在目录下的专属临时文件夹；如该临时目录已存在，会清空复用。
  脚本会先转换到该专属临时文件夹，退出 Office 后再统一移动到目标位置。
  临时文件正常时立即移动；遇到 Office、OneDrive 等短暂占用时按配置退避重试。
  每个文件单独转换，单个文件失败时记录错误并继续处理后续文件。
  单个输入范围初始化或扫描失败时记录错误，并继续处理后续独立范围。
'@
}

# 校验 Office 转换输入路径：仅接受 Windows 绝对路径，并确保最终指向单个文件或文件夹。
function Resolve-OfficeInputPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathText,

        [Parameter(Mandatory = $false)]
        [string]$ParameterName = 'Path'
    )

    $resolvedResult = common\Resolve-InputPath -PathText $PathText -PathType Any
    if ($resolvedResult.Success) {
        return $resolvedResult.Item
    }

    $normalizedPathText = ConvertTo-UnquotedPathText -PathText $PathText
    if ($resolvedResult.Error -eq '请输入 Windows 绝对路径。') {
        throw "$ParameterName 必须是 Windows 文件或文件夹绝对路径。"
    }

    if ($resolvedResult.Error -eq '路径必须只能解析到一个文件或文件夹。') {
        throw "$ParameterName 必须只能解析到一个文件或文件夹。"
    }

    throw "$ParameterName 无法读取: $normalizedPathText。原因: $($resolvedResult.Error)。命令行传参时，路径包含空格、括号等 PowerShell 特殊字符，请使用英文引号包裹路径。"
}

# 逐个校验 Office 转换输入路径，自动去重后返回文件系统对象。
function Resolve-OfficeInputPathList {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$InputPathList
    )

    $resolvedResult = common\Resolve-InputPathList -PathList $InputPathList -PathType Any
    if ($resolvedResult.Success) {
        return @($resolvedResult.Items)
    }

    $errorMessageList = [System.Collections.Generic.List[string]]::new()
    for ($index = 0; $index -lt $InputPathList.Count; $index++) {
        $singleResult = common\Resolve-InputPath -PathText $InputPathList[$index] -PathType Any
        if (-not $singleResult.Success) {
            try {
                [void](Resolve-OfficeInputPath -PathText $InputPathList[$index] -ParameterName "Path$($index + 1)")
            }
            catch {
                $errorMessageList.Add($_.Exception.Message)
            }
        }
    }

    if ($errorMessageList.Count -gt 0) {
        throw "输入路径校验失败:`n$($errorMessageList -join "`n")"
    }

    throw $resolvedResult.Error
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

        $filteredFileList = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
        foreach ($file in $files) {
            $relativePath = [System.IO.Path]::GetRelativePath($RootPath, $file.FullName)
            $separatorIndex = $relativePath.IndexOfAny([char[]]@('\', '/'))
            if ($separatorIndex -ge 0) {
                $firstPathSegment = $relativePath.Substring(0, $separatorIndex)
            }
            else {
                $firstPathSegment = $relativePath
            }

            if (-not (Test-SiblingTempDirectoryName -DirectoryName $firstPathSegment -TempDirectoryName $TempRootDirectoryName)) {
                $filteredFileList.Add($file)
            }
        }

        $files = $filteredFileList.ToArray()

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

    $legacyFileList = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
    foreach ($file in $files) {
        if ($OfficeFormatMap.ContainsKey($file.Extension.ToLowerInvariant())) {
            $legacyFileList.Add($file)
        }
    }
    $legacyFiles = $legacyFileList.ToArray()

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
        $status = if (Test-FileSystemPath -Path $targetPath) { 'SkipExists' } else { 'Convert' }

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
        if (-not $ShouldIncludeHiddenItems -and (Test-HiddenFileSystemItem -Item $inputItem)) {
            Write-Host "跳过隐藏输入项: $($inputItem.FullName)；如需处理隐藏项，请传入 -s。" -ForegroundColor Yellow
            continue
        }

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

# 扫描前准备本轮临时目录；已有残留临时目录时会清空复用。
function Initialize-TempRootDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    return Initialize-SiblingTempDirectory -ParentPath $RootPath -TempDirectoryName $TempRootDirectoryName
}

# 尝试删除空的脚本专属临时目录；非空时保留，避免误删用户仍需处理的内容。
function Remove-TempRootDirectoryIfEmpty {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempRootDirectory
    )

    $tempRootParentPath = Split-Path -Parent $TempRootDirectory
    try {
        $removed = Remove-EmptySiblingTempDirectory `
            -TempDirectoryPath $TempRootDirectory `
            -ParentPath $tempRootParentPath `
            -TempDirectoryName $TempRootDirectoryName
        if (-not $removed) {
            Write-Host "临时目录未清理: 目录仍有内容: $TempRootDirectory" -ForegroundColor Yellow
            return $false
        }

        return $true
    }
    catch {
        Write-Host "临时目录清理跳过: $($_.Exception.Message)" -ForegroundColor Yellow
        return $false
    }
}

# 先保存到脚本专属临时目录，确认生成成功后再移动为目标文件，减少半成品残留风险。
function New-TempOutputPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempRootDirectory,

        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    if (-not (Test-FileSystemDirectory -Path $TempRootDirectory)) {
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

    if (-not (Test-FileSystemPath -Path $TempOutputPath)) {
        return $true
    }

    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            Remove-FileSystemItem -Path $TempOutputPath
            return $true
        }
        catch {
            Start-Sleep -Milliseconds (200 * $attempt)
        }
    }

    return -not (Test-FileSystemPath -Path $TempOutputPath)
}

# 将转换后的临时文件移动到最终位置；正常情况立即完成，短暂占用时才按配置退避重试。
function Move-TempOutputFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempOutputPath,

        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    $effectiveRetryCount = [Math]::Max(1, [int]$FileMoveRetryCount)
    $lastErrorMessage = $null
    for ($attempt = 1; $attempt -le $effectiveRetryCount; $attempt++) {
        if (-not (Test-FileSystemFile -Path $TempOutputPath)) {
            throw "临时输出文件已不存在: $TempOutputPath"
        }

        if (Test-FileSystemPath -Path $TargetPath) {
            throw "目标文件已存在: $TargetPath"
        }

        try {
            [System.IO.File]::Move($TempOutputPath, $TargetPath)
            return
        }
        catch {
            $lastErrorMessage = $_.Exception.Message
            if ($attempt -lt $effectiveRetryCount) {
                $retryDelayMilliseconds = [Math]::Max(0, [int]$FileMoveRetryBaseDelayMilliseconds) * $attempt
                if ($retryDelayMilliseconds -gt 0) {
                    Start-Sleep -Milliseconds $retryDelayMilliseconds
                }
            }
        }
    }

    throw "移动临时输出文件失败，已尝试 $effectiveRetryCount 次。$lastErrorMessage"
}

# 根据临时输出文件路径清理其所在的空临时目录；如果里面还有文件则保留，避免误删用户内容。
function Remove-TempOutputDirectoryIfEmpty {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempOutputPath,

        [Parameter(Mandatory = $true)]
        [string]$TempRootDirectory
    )

    $tempDirectory = [System.IO.Path]::GetDirectoryName($TempOutputPath)
    if ([string]::IsNullOrWhiteSpace($tempDirectory) -or -not (Test-FileSystemPath -Path $tempDirectory)) {
        return
    }

    $normalizedTempDirectory = ConvertTo-NormalizedPath -Path $tempDirectory
    $normalizedTempRootDirectory = ConvertTo-NormalizedPath -Path $TempRootDirectory
    if (-not $normalizedTempDirectory.Equals($normalizedTempRootDirectory, [System.StringComparison]::OrdinalIgnoreCase)) {
        return
    }

    $tempRootParentPath = Split-Path -Parent $TempRootDirectory
    [void](Remove-EmptySiblingTempDirectory `
            -TempDirectoryPath $TempRootDirectory `
            -ParentPath $tempRootParentPath `
            -TempDirectoryName $TempRootDirectoryName)
}

# 获取临时文件移动结果，确保目标文件存在，并尽量清理移动后仍残留的临时文件。
function Get-TempOutputMoveResult {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempOutputPath,

        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    if (-not (Test-FileSystemFile -Path $TargetPath)) {
        return [pscustomobject]@{
            IsValid     = $false
            HasWarning  = $false
            Message     = '目标文件未生成'
            CleanupPath = $null
        }
    }

    if (Test-FileSystemPath -Path $TempOutputPath) {
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
            # PowerPoint 使用 PpAlertLevel 枚举；1 表示 ppAlertsNone，和 Word/Excel 一样静默批量转换。
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
    $documentCollection = $null
    $documentClosed = $false

    try {
        if (Test-FileSystemPath -Path $Plan.TargetPath) {
            return [pscustomobject]@{ Status = 'Skipped'; Message = '目标文件已存在' }
        }

        $tempOutputPath = New-TempOutputPath -TempRootDirectory $TempRootDirectory -TargetPath $Plan.TargetPath
        $application = Get-OfficeApplication -AppName $Plan.AppName -ApplicationCache $ApplicationCache
        switch ($Plan.AppName) {
            'Word' {
                $documentCollection = $application.Documents
                $document = $documentCollection.Open($Plan.SourceFile.FullName, $false, $true, $false)
                $document.SaveAs2($tempOutputPath, $Plan.FileFormat)
                $document.Close($false)
                $documentClosed = $true
            }
            'Excel' {
                $documentCollection = $application.Workbooks
                $document = $documentCollection.Open($Plan.SourceFile.FullName, 0, $true)
                $document.SaveAs($tempOutputPath, $Plan.FileFormat)
                $document.Close($false)
                $documentClosed = $true
            }
            'PowerPoint' {
                $documentCollection = $application.Presentations
                $document = $documentCollection.Open($Plan.SourceFile.FullName, $true, $false, $false)
                $document.SaveAs($tempOutputPath, $Plan.FileFormat)
                $document.Close()
                $documentClosed = $true
            }
        }

        if (-not (Test-FileSystemFile -Path $tempOutputPath)) {
            throw 'Office 未生成临时输出文件。'
        }

        return [pscustomobject]@{ Status = 'ConvertedToTemp'; Message = '临时文件转换成功'; TempOutputPath = $tempOutputPath }
    }
    catch {
        if ($null -ne $document -and -not $documentClosed) {
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
            Remove-TempOutputDirectoryIfEmpty -TempOutputPath $tempOutputPath -TempRootDirectory $TempRootDirectory
        }

        return [pscustomobject]@{ Status = 'Failed'; Message = $_.Exception.Message }
    }
    finally {
        Remove-ComObject -ComObject $document
        $document = $null
        Remove-ComObject -ComObject $documentCollection
        $documentCollection = $null
    }
}

# 统一退出已启动的 Office 应用，释放 COM 引用并返回退出失败信息。
function Close-OfficeApplicationCache {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$ApplicationCache
    )

    $failureMessageList = [System.Collections.Generic.List[string]]::new()
    foreach ($appName in @($ApplicationCache.Keys)) {
        $application = $ApplicationCache[$appName]
        try {
            $application.Quit()
        }
        catch {
            $failureMessageList.Add("$appName 退出失败: $($_.Exception.Message)")
        }
        finally {
            try {
                Remove-ComObject -ComObject $application
            }
            catch {
                $failureMessageList.Add("$appName COM 引用释放失败: $($_.Exception.Message)")
            }
        }
    }

    $ApplicationCache.Clear()
    for ($collectionAttempt = 0; $collectionAttempt -lt 2; $collectionAttempt++) {
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    return [pscustomobject]@{
        Success         = ($failureMessageList.Count -eq 0)
        FailureMessages = $failureMessageList.ToArray()
    }
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

    Write-StageMessage "开始移动转换文件，待移动: $($TempConversionResults.Count)，占用失败时最多重试 $FileMoveRetryCount 次"

    foreach ($conversionResult in $TempConversionResults) {
        $processedCount++
        $plan = $conversionResult.Plan
        $tempOutputPath = $conversionResult.TempOutputPath

        Write-ProgressBar -Activity '移动转换文件' -Status '正在移动文件' -ProcessedCount $processedCount -TotalCount $TempConversionResults.Count -LastPercent ([ref]$lastPercent)

        if (Test-FileSystemPath -Path $plan.TargetPath) {
            if (Remove-TempOutputFile -TempOutputPath $tempOutputPath) {
                Remove-TempOutputDirectoryIfEmpty -TempOutputPath $tempOutputPath -TempRootDirectory $TempRootDirectory
                $warningMessages.Add("跳过移动: $($plan.TargetPathText)`n  原因: 目标文件已存在，临时文件已清理")
            }
            else {
                $warningMessages.Add("跳过移动: $($plan.TargetPathText)`n  原因: 目标文件已存在，但临时文件清理失败`n  临时文件: $tempOutputPath")
            }

            $skippedCount++
            continue
        }

        try {
            Move-TempOutputFile -TempOutputPath $tempOutputPath -TargetPath $plan.TargetPath
            $moveResult = Get-TempOutputMoveResult -TempOutputPath $tempOutputPath -TargetPath $plan.TargetPath
            if (-not $moveResult.IsValid) {
                throw $moveResult.Message
            }

            if ($moveResult.HasWarning) {
                $warningMessages.Add("移动完成但临时文件清理失败: $($plan.TargetPathText)`n  临时文件: $($moveResult.CleanupPath)")
            }

            Remove-TempOutputDirectoryIfEmpty -TempOutputPath $tempOutputPath -TempRootDirectory $TempRootDirectory
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

    if (-not (Test-FileSystemPath -Path $TempRootDirectory)) {
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
        [Parameter(Mandatory = $false)]
        [AllowNull()]
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

    if ($convertPlans.Count -gt 0 -and [string]::IsNullOrWhiteSpace($TempRootDirectory)) {
        throw '存在待转换文件，但未提供可用的临时目录。'
    }

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
        $hadOfficeApplications = ($applicationCache.Count -gt 0)
        if ($hadOfficeApplications) {
            [void](Write-RefreshStatusLine -Message '正在退出 Office 应用...' -Color White -NoNewLine)
        }

        $closeResult = Close-OfficeApplicationCache -ApplicationCache $applicationCache

        if ($hadOfficeApplications) {
            if ($closeResult.Success) {
                [void](Write-RefreshStatusLine -Message 'Office 应用退出请求已完成' -Color Green)
            }
            else {
                [void](Write-RefreshStatusLine -Message '部分 Office 应用退出或 COM 释放失败，移动阶段将按需重试' -Color Yellow)
                foreach ($failureMessage in @($closeResult.FailureMessages)) {
                    Write-Host "  $failureMessage" -ForegroundColor DarkGray
                }
            }
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
    Write-Host "多个路径可用空格或英文分号分隔；路径含空格或英文分号时，请使用英文引号包裹路径。" -ForegroundColor DarkGray
    Write-Host "直接回车退出；输入 0 退出脚本。" -ForegroundColor DarkGray
    $PathInputRaw = Read-ColoredLine -Prompt 'Path: '
    if ($null -eq $PathInputRaw) {
        Write-Host "已退出，未执行扫描。" -ForegroundColor Yellow
        exit 0
    }

    $PathInput = $PathInputRaw.Trim()
    if ([string]::IsNullOrWhiteSpace($PathInput)) {
        Write-Host "已退出，未执行扫描。" -ForegroundColor Yellow
        exit 0
    }

    if ($PathInput -eq '0') {
        Write-Host "已退出，未执行扫描。" -ForegroundColor Yellow
        exit 0
    }

    $PathList = @(Split-InteractivePathInput -PathInput $PathInput)
    if ($PathList.Count -gt 1) {
        Write-Host "识别到 $($PathList.Count) 个路径。" -ForegroundColor DarkGray
    }
}

try {
    # 在正式扫描前集中校验路径，避免后续函数反复处理无效输入。
    $EffectivePathList = @($PathList | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($EffectivePathList.Count -eq 0) {
        Write-Host "未提供有效路径，已退出。" -ForegroundColor Yellow
        exit 0
    }

    $ResolvedInputItemList = @(Resolve-OfficeInputPathList -InputPathList $EffectivePathList)
    $ConversionScopeList = @(New-ConversionScopeList -InputItemList $ResolvedInputItemList)
}
catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

if ($ConversionScopeList.Count -eq 0) {
    Write-Host "未发现可处理的 .doc、.xls、.ppt 文件或目录。" -ForegroundColor Green
    exit 0
}

$ProcessedSourcePathSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$FinalExitCode = 0
$ScopeIndex = 0
foreach ($ConversionScope in $ConversionScopeList) {
    $ScopeIndex++
    if ($ConversionScopeList.Count -gt 1) {
        Write-Host ""
        Write-Host "转换范围 $ScopeIndex / $($ConversionScopeList.Count): $($ConversionScope.Label)" -ForegroundColor Cyan
    }

    $TempRootDirectory = $null
    try {
        $SourceFileList = if ($null -eq $ConversionScope.SourceFileList) { $null } else { $ConversionScope.SourceFileList.ToArray() }
        $ConversionPlans = @(New-ConversionPlanList -RootPath $ConversionScope.RootPath -SourceFileList $SourceFileList -ProcessedSourcePathSet $ProcessedSourcePathSet)
        if ($ConversionPlans.Count -eq 0) {
            Write-Host "未发现 .doc、.xls、.ppt 旧格式 Office 文件。" -ForegroundColor Green
        }
        else {
            $ConvertPlanCount = Write-ConversionPlanSummary -ConversionPlans $ConversionPlans
            if ($ConvertPlanCount -gt 0) {
                $TempRootDirectory = Initialize-TempRootDirectory -RootPath $ConversionScope.RootPath
            }

            Invoke-ConversionPlanList -TempRootDirectory $TempRootDirectory -ConversionPlans $ConversionPlans
        }
    }
    catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        $FinalExitCode = 1
    }
    finally {
        if (-not [string]::IsNullOrWhiteSpace($TempRootDirectory)) {
            [void](Remove-TempRootDirectoryIfEmpty -TempRootDirectory $TempRootDirectory)
        }
    }

}

if ($FinalExitCode -ne 0) {
    exit $FinalExitCode
}
