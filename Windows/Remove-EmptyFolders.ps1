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
  Path   一个或多个文件夹绝对路径；未提供时会引导交互输入；路径含空格时，请使用英文引号包裹路径。
  -s     把隐藏文件夹也作为删除候选；空目录判断始终会检查隐藏项和系统项。
  -yes   跳过详细预览和菜单，输出汇总后等待 10 秒，再执行默认删除。
  -h     显示帮助信息。

规则：
  空文件夹指目录内没有任何子项；隐藏项和系统项也会参与空目录判断。
  默认不把隐藏文件夹作为删除候选；传入 -s 后包含隐藏文件夹。
  脚本只处理输入目录下的子文件夹，不删除输入目录本身。
'@
}

# 校验一批输入目录；同一批输入只要有错误，本批路径都不保留。
function Resolve-InputDirectoryList {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$RawPathList,

        [string[]]$ExistingPathList = @()
    )

    $resolvedResult = common\Resolve-IndependentInputDirectoryList `
        -PathList $RawPathList `
        -ExistingDirectoryPathList $ExistingPathList
    if (-not $resolvedResult.Success) {
        $errorMessage = if ($resolvedResult.Error -eq '目录不能互为父子目录。') {
            '路径无效，请不要输入父子目录。'
        }
        else {
            $resolvedResult.Error
        }

        return [pscustomobject]@{
            Success = $false
            Paths   = @()
            Error   = "$errorMessage 本次输入不保留。"
        }
    }

    if (-not $IncludeHiddenFolderCandidates) {
        foreach ($resolvedItem in @($resolvedResult.Items)) {
            if (Test-HiddenFileSystemItem -Item $resolvedItem) {
                return [pscustomobject]@{
                    Success = $false
                    Paths   = @()
                    Error   = "路径无效: $($resolvedItem.FullName) 是隐藏文件夹；如需处理隐藏项，请传入 -s。本次输入不保留。"
                }
            }
        }
    }

    return [pscustomobject]@{
        Success = $true
        Paths   = @($resolvedResult.Paths)
        Error   = $null
    }
}

# 交互读取一个或多个目录路径。
function Read-InteractiveDirectoryList {
    Write-Host "请输入目录绝对路径。可在同一行输入多个路径。" -ForegroundColor Cyan
    Write-Host "多个路径可用空格或英文分号分隔；路径含空格时，请使用英文引号包裹路径。" -ForegroundColor DarkGray
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
        if ($resolvedPathCount -eq 0) {
            Write-Host "输入路径已存在，本次未新增。" -ForegroundColor DarkGray
        }
        elseif ($rawPathList.Count -gt 1) {
            Write-Host "识别到 $resolvedPathCount 个新路径。" -ForegroundColor DarkGray
        }
    }
}

# 判断目录在计划删除其子空目录后是否会变空。
function Test-DirectoryWillBeEmpty {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[string]]$PlannedDirectoryKeySet,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$ErrorList
    )

    try {
        $childItems = @(Get-DirectoryChildItemList -DirectoryPath $DirectoryPath -ResolvedDirectoryPath $DirectoryPath)
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
        if ($childItem -is [System.IO.DirectoryInfo]) {
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

# 将 Windows 路径转换为扩展长度路径，用于兜底处理末尾空格等普通 Win32 路径难以删除的目录。
function ConvertTo-ExtendedLengthPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ($Path.StartsWith('\\?\', [System.StringComparison]::Ordinal)) {
        return $Path
    }

    if ($Path.StartsWith('\\', [System.StringComparison]::Ordinal)) {
        $trimmedPath = $Path.TrimStart([char[]]@('\', '/'))
        return "\\?\UNC\$trimmedPath"
    }

    return "\\?\$Path"
}

# 普通 Test-Path 失败时使用扩展路径重试，避免异常目录名被误判为不存在。
function Resolve-DeletionDirectoryPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath
    )

    try {
        if ([System.IO.Directory]::Exists($DirectoryPath)) {
            return [pscustomobject]@{
                Exists       = $true
                Path         = $DirectoryPath
                UsedExtended = $false
            }
        }
    }
    catch {
        # 继续走扩展路径兜底。
        Write-Debug "普通路径目录检测失败，准备使用扩展路径兜底: $DirectoryPath。原因: $($_.Exception.Message)"
    }

    $extendedPath = ConvertTo-ExtendedLengthPath -Path $DirectoryPath
    if (-not $extendedPath.Equals($DirectoryPath, [System.StringComparison]::Ordinal)) {
        try {
            if ([System.IO.Directory]::Exists($extendedPath)) {
                return [pscustomobject]@{
                    Exists       = $true
                    Path         = $extendedPath
                    UsedExtended = $true
                }
            }
        }
        catch {
            # 统一返回不存在，由调用方按跳过处理。
            Write-Debug "扩展路径目录检测失败: $extendedPath。原因: $($_.Exception.Message)"
        }
    }

    return [pscustomobject]@{
        Exists       = $false
        Path         = $DirectoryPath
        UsedExtended = $false
    }
}

# 读取目录子项；普通路径失败时使用扩展路径兜底。
function Get-DirectoryChildItemList {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath,

        [Parameter(Mandatory = $true)]
        [string]$ResolvedDirectoryPath
    )

    try {
        return @(([System.IO.DirectoryInfo]::new($ResolvedDirectoryPath)).EnumerateFileSystemInfos())
    }
    catch {
        $extendedPath = ConvertTo-ExtendedLengthPath -Path $DirectoryPath
        if ($extendedPath.Equals($ResolvedDirectoryPath, [System.StringComparison]::Ordinal)) {
            throw
        }

        return @(([System.IO.DirectoryInfo]::new($extendedPath)).EnumerateFileSystemInfos())
    }
}

# 删除前读取第一个子项；普通路径失败时再用扩展路径兜底，失败则不删除。
function Get-FirstDirectoryChildItem {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath,

        [Parameter(Mandatory = $true)]
        [string]$ResolvedDirectoryPath
    )

    foreach ($childItem in (Get-DirectoryChildItemList -DirectoryPath $DirectoryPath -ResolvedDirectoryPath $ResolvedDirectoryPath)) {
        return $childItem
    }

    return $null
}

# 删除空目录；普通路径删除失败时，用扩展路径重新确认为空后再删除。
function Remove-EmptyDirectoryWithFallback {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath,

        [Parameter(Mandatory = $true)]
        [string]$ResolvedDirectoryPath
    )

    try {
        [System.IO.Directory]::Delete($ResolvedDirectoryPath)
        return
    }
    catch {
        $extendedPath = ConvertTo-ExtendedLengthPath -Path $DirectoryPath
        if ($extendedPath.Equals($ResolvedDirectoryPath, [System.StringComparison]::Ordinal)) {
            throw
        }

        $firstChild = $null
        foreach ($childPath in [System.IO.Directory]::EnumerateFileSystemEntries($extendedPath)) {
            $firstChild = $childPath
            break
        }

        if ($null -ne $firstChild) {
            throw '目录当前不是空文件夹。'
        }

        [System.IO.Directory]::Delete($extendedPath)
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

        $directoryPathState = Resolve-DeletionDirectoryPath -DirectoryPath $targetDirectory.FullName
        if (-not $directoryPathState.Exists) {
            $skippedList.Add([pscustomobject]@{
                    Path    = $targetDirectory.FullName
                    Message = '目录已不存在。'
                })
            continue
        }

        try {
            $firstChild = Get-FirstDirectoryChildItem `
                -DirectoryPath $targetDirectory.FullName `
                -ResolvedDirectoryPath $directoryPathState.Path
            if ($null -ne $firstChild) {
                $skippedList.Add([pscustomobject]@{
                        Path    = $targetDirectory.FullName
                        Message = '目录当前不是空文件夹。'
                    })
                continue
            }

            Remove-EmptyDirectoryWithFallback `
                -DirectoryPath $targetDirectory.FullName `
                -ResolvedDirectoryPath $directoryPathState.Path
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
    $ResolvedPathResult = Resolve-InputDirectoryList -RawPathList $PathList
    if (-not $ResolvedPathResult.Success) {
        Write-Host $ResolvedPathResult.Error -ForegroundColor Red
        exit 1
    }

    $RootPathList = @($ResolvedPathResult.Paths)
}

Write-StageMessage "开始扫描输入目录，数量: $($RootPathList.Count)"
$DeletionPlanResult = New-EmptyFolderDeletionPlan -RootPathList $RootPathList
$DeletionPlanList = @($DeletionPlanResult.Items)

Write-ErrorRecordList -ErrorRecordList @($DeletionPlanResult.Errors)

if ($DeletionPlanList.Count -eq 0) {
    Write-Host
    Write-Host "扫描完成，未发现可删除的空文件夹。" -ForegroundColor Green
    exit 0
}

if ($AssumeYes) {
    Write-Host
    Write-Host -NoNewline "扫描完成。默认计划删除空文件夹数: " -ForegroundColor White
    Write-Host $DeletionPlanList.Count -ForegroundColor Magenta

    if (Wait-DangerousOperationGracePeriod `
            -WarningMessage '已启用 -yes，将跳过详细预览并执行默认删除。' `
            -AdditionalWarningMessage "计划删除空文件夹数: $($DeletionPlanList.Count)" `
            -CancelledMessage '已取消默认删除。' `
            -CompletedMessage '倒计时结束，开始执行默认删除。' `
            -CountdownMessageFormat '倒计时 {0} 秒后开始删除，按 Enter 取消...') {
        Invoke-EmptyFolderDeletion -TargetDirectoryList $DeletionPlanList
    }
    exit 0
}

Write-DeletionPreview -DeletionPlanList $DeletionPlanList
$OperationChoice = Read-OperationChoice

switch ($OperationChoice) {
    'Default' {
        Invoke-EmptyFolderDeletion -TargetDirectoryList $DeletionPlanList
    }
    'Manual' {
        $SelectedDeletionList = @(Read-ManualDeletionSelection -DeletionPlanList $DeletionPlanList)
        if ($SelectedDeletionList.Count -gt 0) {
            Invoke-EmptyFolderDeletion -TargetDirectoryList $SelectedDeletionList
        }
    }
    'Exit' {
        Write-Host "已退出，未删除任何文件夹。" -ForegroundColor Yellow
    }
}
