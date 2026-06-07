#requires -Version 7.0
<#
用途：
  删除一个或多个指定文件夹下的空子文件夹。

参数：
  Path   一个或多个文件夹绝对路径；未提供时会引导交互输入。
  -s     把隐藏文件夹也作为删除候选。
  -yes   跳过详细预览和菜单，输出汇总后等待 10 秒，再执行默认删除。
  -h     显示帮助信息。

规则：
  空文件夹指目录内没有任何子项；隐藏项和系统项也会参与空目录判断。
  脚本只处理输入目录下的子文件夹，不删除输入目录本身。
#>

# ========== 参数区 ==========

[CmdletBinding()]
param(
    [Alias('h')]
    [switch]$Help,

    [Alias('s')]
    [switch]$IncludeHidden,

    [Alias('yes')]
    [switch]$AssumeYes,

    [Parameter(Mandatory = $false, Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$PathList
)

# ========== 运行环境设置 ==========

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# 加载 Windows 脚本公共工具函数。
Import-Module -Name (Join-Path $PSScriptRoot 'common.psm1') -Force

# 是否把隐藏文件夹也作为可删除候选；空目录判断始终会检查隐藏和系统项。
$IncludeHiddenFolderCandidates = [bool]$IncludeHidden

# ========== 输出与工具函数 ==========

function Show-HelpText {
    Write-Host @'
用途：
  删除一个或多个指定文件夹下的空子文件夹。

用法：
  pwsh -File .\Remove-EmptyFolders.ps1 [-s] [-yes] [Path1] [Path2 ...]

参数：
  Path   一个或多个文件夹绝对路径；未提供时会引导交互输入；路径含空格请使用英文引号。
  -s     把隐藏文件夹也作为删除候选；空目录判断始终会检查隐藏项和系统项。
  -yes   跳过详细预览和菜单，输出汇总后等待 10 秒，再执行默认删除。
  -h     显示帮助信息。

规则：
  空文件夹指目录内没有任何子项；隐藏项和系统项也会参与空目录判断。
  默认不把隐藏文件夹作为删除候选；传入 -s 后包含隐藏文件夹。
  脚本只处理输入目录下的子文件夹，不删除输入目录本身。
'@
}

# 校验并解析单个输入目录。
function Resolve-InputDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RawPath
    )

    $cleanPath = ConvertTo-UnquotedPathText -PathText $RawPath
    if ([string]::IsNullOrWhiteSpace($cleanPath)) {
        return [pscustomobject]@{
            Success = $false
            Path    = $null
            Error   = '路径为空。'
        }
    }

    try {
        if (-not [System.IO.Path]::IsPathFullyQualified($cleanPath)) {
            return [pscustomobject]@{
                Success = $false
                Path    = $null
                Error   = '请输入存在的 Windows 文件夹绝对路径。'
            }
        }
    }
    catch {
        return [pscustomobject]@{
            Success = $false
            Path    = $null
            Error   = '路径格式无效。'
        }
    }

    if (-not (Test-Path -LiteralPath $cleanPath -PathType Container)) {
        return [pscustomobject]@{
            Success = $false
            Path    = $null
            Error   = '文件夹不存在。'
        }
    }

    try {
        $resolvedPath = (Resolve-Path -LiteralPath $cleanPath -ErrorAction Stop).ProviderPath
        return [pscustomobject]@{
            Success = $true
            Path    = (Get-DirectoryKey -DirectoryPath $resolvedPath)
            Error   = $null
        }
    }
    catch {
        return [pscustomobject]@{
            Success = $false
            Path    = $null
            Error   = $_.Exception.Message
        }
    }
}

# 校验一批输入目录；同一批输入只要有错误，本批路径都不保留。
function Resolve-InputDirectoryList {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$RawPathList,

        [string[]]$ExistingPathList = @()
    )

    $resolvedPathList = [System.Collections.Generic.List[string]]::new()

    foreach ($rawPath in $RawPathList) {
        $resolvedResult = Resolve-InputDirectory -RawPath $rawPath
        if (-not $resolvedResult.Success) {
            return [pscustomobject]@{
                Success = $false
                Paths   = @()
                Error   = "路径无效: $rawPath。$($resolvedResult.Error) 本次输入不保留。"
            }
        }

        foreach ($existingPath in $ExistingPathList) {
            if (Test-DirectoryOverlap -LeftPath $resolvedResult.Path -RightPath $existingPath) {
                return [pscustomobject]@{
                    Success = $false
                    Paths   = @()
                    Error   = '路径无效，请不要重复输入或输入互相包含的目录。本次输入不保留。'
                }
            }
        }

        foreach ($newPath in $resolvedPathList) {
            if (Test-DirectoryOverlap -LeftPath $resolvedResult.Path -RightPath $newPath) {
                return [pscustomobject]@{
                    Success = $false
                    Paths   = @()
                    Error   = '路径无效，请不要重复输入或输入互相包含的目录。本次输入不保留。'
                }
            }
        }

        $resolvedPathList.Add($resolvedResult.Path)
    }

    return [pscustomobject]@{
        Success = $true
        Paths   = @($resolvedPathList.ToArray())
        Error   = $null
    }
}

# 交互读取一个或多个目录路径。
function Read-InteractiveDirectoryList {
    Write-Host "请输入目录绝对路径。可在同一行输入多个路径。" -ForegroundColor Cyan
    Write-Host "多个路径可用空格或英文分号分隔；路径含空格请使用英文引号。" -ForegroundColor DarkGray
    Write-Host "直接回车开始执行或退出；输入 0 退出脚本。" -ForegroundColor DarkGray

    $pathList = [System.Collections.Generic.List[string]]::new()

    while ($true) {
        $promptIndex = $pathList.Count + 1
        $inputLine = Read-ColoredLine -Prompt "Path${promptIndex}: "

        if ([string]::IsNullOrWhiteSpace($inputLine)) {
            if ($pathList.Count -eq 0) {
                Write-Host "已退出，未执行扫描。" -ForegroundColor Yellow
                exit 0
            }

            return $pathList.ToArray()
        }

        if ($inputLine.Trim() -eq '0') {
            Write-Host "已退出，未执行扫描。" -ForegroundColor Yellow
            exit 0
        }

        $rawPathList = @(Split-InteractivePathInput -PathInput $inputLine)
        if ($rawPathList.Count -eq 0) {
            Write-Host "输入无效，请重新输入目录绝对路径。" -ForegroundColor Red
            continue
        }

        $resolvedResult = Resolve-InputDirectoryList -RawPathList $rawPathList -ExistingPathList $pathList.ToArray()
        if (-not $resolvedResult.Success) {
            Write-Host $resolvedResult.Error -ForegroundColor Red
            continue
        }

        foreach ($resolvedPath in @($resolvedResult.Paths)) {
            $pathList.Add($resolvedPath)
        }

        $resolvedPathCount = @($resolvedResult.Paths).Count
        if ($resolvedPathCount -gt 1) {
            Write-Host "识别到 $resolvedPathCount 个路径。" -ForegroundColor DarkGray
        }
    }
}

# 判断目录在计划删除其子空目录后是否会变空。
function Test-DirectoryWillBeEmpty {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath,

        [System.Collections.Generic.HashSet[string]]$PlannedDirectoryKeySet,

        [System.Collections.Generic.List[object]]$ErrorList
    )

    try {
        $childItems = @(Get-ChildItem -LiteralPath $DirectoryPath -Force -ErrorAction Stop)
    }
    catch {
        $ErrorList.Add([pscustomobject]@{
                Path    = $DirectoryPath
                Message = $_.Exception.Message
                Stage   = '判断空文件夹'
            })
        return $false
    }

    foreach ($childItem in $childItems) {
        if ($childItem.PSIsContainer) {
            $childKey = Get-DirectoryKey -DirectoryPath $childItem.FullName
            if ($PlannedDirectoryKeySet.Contains($childKey)) {
                continue
            }
        }

        return $false
    }

    return $true
}

# 扫描所有输入目录并计算默认删除计划。
function New-EmptyFolderDeletionPlan {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$RootPathList
    )

    $allDirectoryList = [System.Collections.Generic.List[object]]::new()
    $errorList = [System.Collections.Generic.List[object]]::new()

    foreach ($rootPath in $RootPathList) {
        Write-StageMessage "开始收集子文件夹: $rootPath"

        $getChildItemParameters = @{
            LiteralPath = $rootPath
            Directory   = $true
            Recurse     = $true
            ErrorAction = 'SilentlyContinue'
        }
        if ($IncludeHiddenFolderCandidates) {
            $getChildItemParameters.Force = $true
        }

        $rootScanErrors = @()
        $childDirectoryList = @(Get-ChildItem @getChildItemParameters -ErrorVariable rootScanErrors)

        foreach ($scanError in $rootScanErrors) {
            $errorList.Add([pscustomobject]@{
                    Path    = $rootPath
                    Message = $scanError.Exception.Message
                    Stage   = '扫描子文件夹'
                })
        }

        foreach ($childDirectory in $childDirectoryList) {
            $allDirectoryList.Add([pscustomobject]@{
                    FullName   = $childDirectory.FullName
                    Root       = $rootPath
                    Attributes = $childDirectory.Attributes
                })
        }

        $hiddenText = if ($IncludeHiddenFolderCandidates) { '包含隐藏项' } else { '不包含隐藏项' }
        Write-StageMessage "子文件夹收集完成，数量: $($childDirectoryList.Count)，$hiddenText"
    }

    $plannedDirectoryKeySet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $plannedDirectoryList = [System.Collections.Generic.List[object]]::new()

    if ($allDirectoryList.Count -eq 0) {
        return [pscustomobject]@{
            Items  = @()
            Errors = $errorList.ToArray()
        }
    }

    $sortedDirectoryList = @($allDirectoryList | Sort-Object `
            @{ Expression = { Get-DirectoryPathDepth -DirectoryPath $_.FullName }; Descending = $true }, `
            @{ Expression = { $_.FullName }; Descending = $true })

    $activity = '空文件夹判断'
    $currentIndex = 0
    $lastFolderCheckPercent = -1
    foreach ($directory in $sortedDirectoryList) {
        $currentIndex++
        Write-ProgressBar -Activity $activity -Status '正在判断空文件夹' -ProcessedCount $currentIndex -TotalCount $sortedDirectoryList.Count -LastPercent ([ref]$lastFolderCheckPercent)

        if (($directory.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            continue
        }

        if (Test-DirectoryWillBeEmpty -DirectoryPath $directory.FullName -PlannedDirectoryKeySet $plannedDirectoryKeySet -ErrorList $errorList) {
            $directoryKey = Get-DirectoryKey -DirectoryPath $directory.FullName
            if ($plannedDirectoryKeySet.Add($directoryKey)) {
                $plannedDirectoryList.Add($directory)
            }
        }
    }
    Write-ProgressBar -Activity $activity -Status '空文件夹判断完成' -ProcessedCount $sortedDirectoryList.Count -TotalCount $sortedDirectoryList.Count -LastPercent ([ref]$lastFolderCheckPercent) -Force
    Complete-DynamicStatusLine

    return [pscustomobject]@{
        Items  = $plannedDirectoryList.ToArray()
        Errors = $errorList.ToArray()
    }
}

# 输出扫描期间记录的错误。
function Write-ErrorRecordList {
    param(
        [object[]]$ErrorRecordList
    )

    if ($ErrorRecordList.Count -eq 0) {
        return
    }

    $warningList = New-DeferredScanWarningList
    foreach ($errorRecord in $ErrorRecordList) {
        Add-DeferredScanWarning `
            -WarningList $warningList `
            -Message "[$($errorRecord.Stage)] 跳过目录" `
            -Path $errorRecord.Path `
            -Reason $errorRecord.Message
    }

    Write-Host
    Write-DeferredScanWarningList -WarningList $warningList -Title '空文件夹扫描跳过汇总'
}

# 输出删除预览。
function Write-DeletionPreview {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$DeletionPlanList
    )

    if ($DeletionPlanList.Count -eq 0) {
        Write-Host "未发现可删除的空文件夹。" -ForegroundColor Green
        return
    }

    Write-Host
    Write-Host "删除预览:" -ForegroundColor Yellow
    Write-PreviewSeparator -NoLeadingBlank

    $index = 0
    foreach ($planItem in $DeletionPlanList) {
        $index++
        Write-Host -NoNewline ("{0,4}. " -f $index) -ForegroundColor Cyan
        Write-Host $planItem.FullName -ForegroundColor White
    }

    Write-PreviewSeparator -NoLeadingBlank
    Write-Host
    Write-Host -NoNewline "空文件夹列举完成。默认计划删除空文件夹数: " -ForegroundColor White
    Write-Host $DeletionPlanList.Count -ForegroundColor Magenta
}

# 读取操作菜单。
function Read-OperationChoice {
    $choice = Read-MenuChoice -Title '请选择操作:' -EndOfInputChoice '0' -MenuOptionList @(
        [pscustomobject]@{ Value = '1'; Label = '默认删除' }
        [pscustomobject]@{ Value = '2'; Label = '手动删除' }
        [pscustomobject]@{ Value = '0'; Label = '退出脚本' }
    )

    switch ($choice) {
        '1' { return 'Default' }
        '2' { return 'Manual' }
        '0' { return 'Exit' }
    }
}

# 手动选择要删除的预览项。
function Read-ManualDeletionSelection {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$DeletionPlanList
    )

    while ($true) {
        Write-Host
        Write-Host "请输入要删除的编号，多个编号用英文逗号分隔；输入 0 取消，输入 00 退出脚本。" -ForegroundColor Cyan
        $inputText = Read-ColoredLine -Prompt '编号: '
        if ($null -eq $inputText) {
            Write-Host "输入流已结束，程序退出。" -ForegroundColor Yellow
            exit 0
        }

        $trimmedInput = $inputText.Trim()

        if ($trimmedInput -eq '00') {
            Write-Host "已退出，未继续删除。" -ForegroundColor Yellow
            exit 0
        }

        if ($trimmedInput -eq '0') {
            Write-Host "已取消手动删除。" -ForegroundColor Yellow
            return @()
        }

        $selectedIndexSet = [System.Collections.Generic.HashSet[int]]::new()
        $isValid = $true

        foreach ($indexText in ($trimmedInput -split ',')) {
            $parsedIndex = 0
            if (-not [int]::TryParse($indexText.Trim(), [ref]$parsedIndex) -or
                $parsedIndex -lt 1 -or
                $parsedIndex -gt $DeletionPlanList.Count) {
                $isValid = $false
                break
            }

            [void]$selectedIndexSet.Add($parsedIndex)
        }

        if (-not $isValid -or $selectedIndexSet.Count -eq 0) {
            Write-Host "输入无效，请输入预览列表中的编号。" -ForegroundColor Red
            continue
        }

        $selectedItemList = [System.Collections.Generic.List[object]]::new()
        foreach ($selectedIndex in ($selectedIndexSet | Sort-Object)) {
            $selectedItemList.Add($DeletionPlanList[$selectedIndex - 1])
        }

        return $selectedItemList.ToArray()
    }
}

# 删除选定的空文件夹。
function Invoke-EmptyFolderDeletion {
    param(
        [object[]]$TargetDirectoryList
    )

    if ($null -eq $TargetDirectoryList -or $TargetDirectoryList.Count -eq 0) {
        Write-Host "没有需要删除的空文件夹。" -ForegroundColor Green
        return
    }

    $orderedTargetList = @($TargetDirectoryList | Sort-Object `
            @{ Expression = { Get-DirectoryPathDepth -DirectoryPath $_.FullName }; Descending = $true }, `
            @{ Expression = { $_.FullName }; Descending = $true })

    $deletedList = [System.Collections.Generic.List[string]]::new()
    $skippedList = [System.Collections.Generic.List[object]]::new()
    $failedList = [System.Collections.Generic.List[object]]::new()

    $activity = '空文件夹删除'
    $currentIndex = 0
    $lastDeletionPercent = -1

    foreach ($targetDirectory in $orderedTargetList) {
        $currentIndex++
        Write-ProgressBar -Activity $activity -Status '正在删除空文件夹' -ProcessedCount $currentIndex -TotalCount $orderedTargetList.Count -LastPercent ([ref]$lastDeletionPercent)

        if (-not (Test-Path -LiteralPath $targetDirectory.FullName -PathType Container)) {
            $skippedList.Add([pscustomobject]@{
                    Path    = $targetDirectory.FullName
                    Message = '目录已不存在。'
                })
            continue
        }

        try {
            $firstChild = Get-ChildItem -LiteralPath $targetDirectory.FullName -Force -ErrorAction Stop | Select-Object -First 1
            if ($null -ne $firstChild) {
                $skippedList.Add([pscustomobject]@{
                        Path    = $targetDirectory.FullName
                        Message = '目录当前不是空文件夹。'
                    })
                continue
            }

            Remove-Item -LiteralPath $targetDirectory.FullName -Force -ErrorAction Stop
            $deletedList.Add($targetDirectory.FullName)
        }
        catch {
            $failedList.Add([pscustomobject]@{
                    Path    = $targetDirectory.FullName
                    Message = $_.Exception.Message
                })
        }
    }

    Write-ProgressBar -Activity $activity -Status '空文件夹删除完成' -ProcessedCount $orderedTargetList.Count -TotalCount $orderedTargetList.Count -LastPercent ([ref]$lastDeletionPercent) -Force
    Complete-DynamicStatusLine

    if ($deletedList.Count -gt 0) {
        Write-Host
        Write-Host "已删除空文件夹:" -ForegroundColor Magenta
        foreach ($deletedPath in $deletedList) {
            Write-Host "  $deletedPath" -ForegroundColor Magenta
        }
    }

    if ($skippedList.Count -gt 0) {
        Write-Host
        Write-Host "跳过项目:" -ForegroundColor Yellow
        foreach ($skippedItem in $skippedList) {
            Write-Host "  $($skippedItem.Path)" -ForegroundColor Yellow
            Write-Host "    原因: $($skippedItem.Message)" -ForegroundColor DarkGray
        }
    }

    if ($failedList.Count -gt 0) {
        Write-Host
        Write-Host "删除失败:" -ForegroundColor Red
        foreach ($failedItem in $failedList) {
            Write-Host "  $($failedItem.Path)" -ForegroundColor Red
            Write-Host "    原因: $($failedItem.Message)" -ForegroundColor DarkGray
        }
    }

    Write-Host
    Write-Host -NoNewline "删除完成。已删除 " -ForegroundColor White
    Write-Host -NoNewline $deletedList.Count -ForegroundColor Magenta
    Write-Host -NoNewline "，跳过 " -ForegroundColor White
    Write-Host -NoNewline $skippedList.Count -ForegroundColor Yellow
    Write-Host -NoNewline "，失败 " -ForegroundColor White
    Write-Host $failedList.Count -ForegroundColor $(if ($failedList.Count -gt 0) { 'Red' } else { 'Green' })
}

# ========== 主逻辑 ==========

if ($Help) {
    Show-HelpText
    exit 0
}

if ($null -eq $PathList -or $PathList.Count -eq 0) {
    $RootPathList = @(Read-InteractiveDirectoryList)
}
else {
    $resolvedPathResult = Resolve-InputDirectoryList -RawPathList $PathList
    if (-not $resolvedPathResult.Success) {
        Write-Host $resolvedPathResult.Error -ForegroundColor Red
        exit 1
    }

    $RootPathList = @($resolvedPathResult.Paths)
}

Write-StageMessage "开始扫描输入目录，数量: $($RootPathList.Count)"
$deletionPlanResult = New-EmptyFolderDeletionPlan -RootPathList $RootPathList
$DeletionPlanList = @($deletionPlanResult.Items)

Write-ErrorRecordList -ErrorRecordList @($deletionPlanResult.Errors)

if ($DeletionPlanList.Count -eq 0) {
    Write-Host
    Write-Host "扫描完成，未发现可删除的空文件夹。" -ForegroundColor Green
    exit 0
}

if ($AssumeYes) {
    Write-Host
    Write-Host -NoNewline "扫描完成。默认计划删除空文件夹数: " -ForegroundColor White
    Write-Host $DeletionPlanList.Count -ForegroundColor Magenta

    if (Wait-AssumeYesDeletionGracePeriod `
            -WarningMessage '已启用 -yes，将跳过详细预览并执行默认删除。' `
            -AdditionalWarningMessage "计划删除空文件夹数: $($DeletionPlanList.Count)" `
            -CancelledMessage '已取消默认删除。') {
        Invoke-EmptyFolderDeletion -TargetDirectoryList $DeletionPlanList
    }
    exit 0
}

Write-DeletionPreview -DeletionPlanList $DeletionPlanList
$operationChoice = Read-OperationChoice

switch ($operationChoice) {
    'Default' {
        Invoke-EmptyFolderDeletion -TargetDirectoryList $DeletionPlanList
    }
    'Manual' {
        $selectedDeletionList = @(Read-ManualDeletionSelection -DeletionPlanList $DeletionPlanList)
        Invoke-EmptyFolderDeletion -TargetDirectoryList $selectedDeletionList
    }
    'Exit' {
        Write-Host "已退出，未删除任何文件夹。" -ForegroundColor Yellow
    }
}
