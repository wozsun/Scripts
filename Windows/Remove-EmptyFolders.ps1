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
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Path,

    [switch]$s,

    [switch]$yes,

    [switch]$h
)

# ========== 可调整配置 ==========

# 单行进度条宽度。
$ProgressBarWidth = 32

# 使用 -yes 时，默认删除前等待的秒数。
$YesCountdownSeconds = 10

# ========== 运行环境设置 ==========

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# 是否把隐藏文件夹也作为可删除候选；空目录判断始终会检查隐藏和系统项。
$IncludeHiddenFolderCandidates = [bool]$s

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

# 输出阶段信息。
function Write-StageMessage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host "[进度] $Message" -ForegroundColor Cyan
}

# 输出菜单项，保持菜单编号醒目。
function Write-MenuItem {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Number,

        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    Write-Host -NoNewline "  $Number " -ForegroundColor Cyan
    Write-Host $Text -ForegroundColor White
}

# 输出彩色输入提示并读取一行文本。
function Read-ColoredLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt
    )

    Write-Host -NoNewline $Prompt -ForegroundColor Cyan
    return [Console]::ReadLine()
}

# 输出单行动态进度条。
function Write-ProgressLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Activity,

        [Parameter(Mandatory = $true)]
        [int]$Current,

        [Parameter(Mandatory = $true)]
        [int]$Total,

        [Parameter(Mandatory = $true)]
        [string]$Status,

        [switch]$Force
    )

    if ($Total -le 0) {
        return
    }

    $percent = [int][Math]::Floor(($Current / $Total) * 100)
    $stateKey = "ProgressPercent:$Activity"
    $previousPercent = if (Test-Path -LiteralPath "variable:script:$stateKey") {
        Get-Variable -Name $stateKey -Scope Script -ValueOnly
    }
    else {
        $null
    }

    if (-not $Force -and $null -ne $previousPercent -and $previousPercent -eq $percent) {
        return
    }

    Set-Variable -Name $stateKey -Scope Script -Value $percent

    $filledWidth = if ($percent -ge 100) {
        $ProgressBarWidth
    }
    else {
        [int][Math]::Floor($ProgressBarWidth * $percent / 100)
    }
    $emptyWidth = $ProgressBarWidth - $filledWidth
    $bar = ('#' * $filledWidth) + ('-' * $emptyWidth)
    $line = "[进度] $Activity [$bar] $percent% $Status ($Current / $Total)"

    Write-Host -NoNewline "`r$line" -ForegroundColor Cyan
}

# 结束单行进度条，避免后续输出和进度行混在一起。
function Complete-ProgressLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Activity
    )

    $stateKey = "ProgressPercent:$Activity"
    if (Test-Path -LiteralPath "variable:script:$stateKey") {
        Remove-Variable -Name $stateKey -Scope Script -Force
    }

    Write-Host
}

# 去掉路径两端的英文引号。
function ConvertTo-UnquotedPathText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $trimmedText = $Text.Trim()
    if ($trimmedText.Length -ge 2) {
        $firstCharacterCode = [int][char]$trimmedText[0]
        $lastCharacterCode = [int][char]$trimmedText[$trimmedText.Length - 1]
        if (($firstCharacterCode -eq 34 -and $lastCharacterCode -eq 34) -or
            ($firstCharacterCode -eq 39 -and $lastCharacterCode -eq 39)) {
            return $trimmedText.Substring(1, $trimmedText.Length - 2).Trim()
        }
    }

    return $trimmedText
}

# 将一行路径输入拆分为多个路径；英文引号只在路径片段开头生效，避免误伤路径中的撇号。
function Split-InputPathLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputText
    )

    $pathList = [System.Collections.Generic.List[string]]::new()
    $builder = [System.Text.StringBuilder]::new()
    $closingQuote = $null

    foreach ($character in $InputText.ToCharArray()) {
        if ($null -ne $closingQuote) {
            if ($character -eq $closingQuote) {
                $closingQuote = $null
                continue
            }

            [void]$builder.Append($character)
            continue
        }

        $characterCode = [int][char]$character

        $canStartQuotedPath = $builder.Length -eq 0
        if ($canStartQuotedPath -and $characterCode -eq 34) {
            $closingQuote = [char]34
            continue
        }
        elseif ($canStartQuotedPath -and $characterCode -eq 39) {
            $closingQuote = [char]39
            continue
        }
        if ([char]::IsWhiteSpace($character) -or $character -eq ';') {
            $currentPath = ConvertTo-UnquotedPathText -Text $builder.ToString()
            if (-not [string]::IsNullOrWhiteSpace($currentPath)) {
                $pathList.Add($currentPath)
            }

            [void]$builder.Clear()
            continue
        }

        [void]$builder.Append($character)
    }

    $lastPath = ConvertTo-UnquotedPathText -Text $builder.ToString()
    if (-not [string]::IsNullOrWhiteSpace($lastPath)) {
        $pathList.Add($lastPath)
    }

    return $pathList.ToArray()
}

# 生成用于比较的目录键，统一去除尾部分隔符。
function Get-DirectoryKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath
    )

    return [System.IO.Path]::TrimEndingDirectorySeparator([System.IO.Path]::GetFullPath($DirectoryPath))
}

# 给目录路径追加结尾分隔符，避免 C:\A 和 C:\AB 这种前缀误判；根目录已带分隔符时不再重复追加。
function Add-TrailingDirectorySeparator {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath
    )

    if ($DirectoryPath.EndsWith('\', [System.StringComparison]::Ordinal) -or
        $DirectoryPath.EndsWith('/', [System.StringComparison]::Ordinal)) {
        return $DirectoryPath
    }

    return "$DirectoryPath$([System.IO.Path]::DirectorySeparatorChar)"
}

# 计算路径层级，用于从深到浅处理目录。
function Get-PathDepth {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath
    )

    return ((Get-DirectoryKey -DirectoryPath $DirectoryPath) -split '[\\/]').Count
}

# 判断两个目录是否相同或互相包含。
function Test-DirectoryOverlap {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LeftPath,

        [Parameter(Mandatory = $true)]
        [string]$RightPath
    )

    $leftKey = Get-DirectoryKey -DirectoryPath $LeftPath
    $rightKey = Get-DirectoryKey -DirectoryPath $RightPath

    if ($leftKey.Equals($rightKey, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }

    $leftPrefix = Add-TrailingDirectorySeparator -DirectoryPath $leftKey
    $rightPrefix = Add-TrailingDirectorySeparator -DirectoryPath $rightKey

    return ($leftPrefix.StartsWith($rightPrefix, [System.StringComparison]::OrdinalIgnoreCase) -or
        $rightPrefix.StartsWith($leftPrefix, [System.StringComparison]::OrdinalIgnoreCase))
}

# 校验并解析单个输入目录。
function Resolve-InputDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RawPath
    )

    $cleanPath = ConvertTo-UnquotedPathText -Text $RawPath
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

        $rawPathList = @(Split-InputPathLine -InputText $inputLine)
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
            @{ Expression = { Get-PathDepth -DirectoryPath $_.FullName }; Descending = $true }, `
            @{ Expression = { $_.FullName }; Descending = $true })

    $activity = '空文件夹判断'
    $currentIndex = 0
    foreach ($directory in $sortedDirectoryList) {
        $currentIndex++
        Write-ProgressLine -Activity $activity -Current $currentIndex -Total $sortedDirectoryList.Count -Status '正在判断空文件夹'

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
    Write-ProgressLine -Activity $activity -Current $sortedDirectoryList.Count -Total $sortedDirectoryList.Count -Status '空文件夹判断完成' -Force
    Complete-ProgressLine -Activity $activity

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

    Write-Host
    Write-Host "扫描或判断过程中有 $($ErrorRecordList.Count) 项失败，已跳过相关目录:" -ForegroundColor Yellow
    foreach ($errorRecord in $ErrorRecordList) {
        Write-Host "  [$($errorRecord.Stage)] $($errorRecord.Path)" -ForegroundColor Red
        Write-Host "    原因: $($errorRecord.Message)" -ForegroundColor DarkGray
    }
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
    Write-Host "================================================================" -ForegroundColor DarkGray

    $index = 0
    foreach ($planItem in $DeletionPlanList) {
        $index++
        Write-Host -NoNewline ("{0,4}. " -f $index) -ForegroundColor Cyan
        Write-Host $planItem.FullName -ForegroundColor White
    }

    Write-Host "================================================================" -ForegroundColor DarkGray
    Write-Host
    Write-Host -NoNewline "空文件夹列举完成。默认计划删除空文件夹数: " -ForegroundColor White
    Write-Host $DeletionPlanList.Count -ForegroundColor Magenta
}

# 读取操作菜单。
function Read-OperationChoice {
    while ($true) {
        Write-Host
        Write-Host "请选择操作:" -ForegroundColor Cyan
        Write-MenuItem -Number '1' -Text '默认删除'
        Write-MenuItem -Number '2' -Text '手动删除'
        Write-MenuItem -Number '0' -Text '退出脚本'

        $choice = Read-ColoredLine -Prompt '请输入选项: '
        switch ($choice.Trim()) {
            '1' { return 'Default' }
            '2' { return 'Manual' }
            '0' { return 'Exit' }
            default {
                Write-Host "输入无效，请输入 1、2 或 0。" -ForegroundColor Red
            }
        }
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

# 使用 -yes 时给出倒计时，留出中止机会。
function Wait-DefaultDeletionCountdown {
    param(
        [Parameter(Mandatory = $true)]
        [int]$DeletionCount
    )

    Write-Host
    Write-Host "已启用 -yes，将跳过详细预览并执行默认删除。" -ForegroundColor Yellow
    Write-Host "计划删除空文件夹数: $DeletionCount" -ForegroundColor Yellow

    for ($remainingSeconds = $YesCountdownSeconds; $remainingSeconds -gt 0; $remainingSeconds--) {
        Write-Host -NoNewline "`r$remainingSeconds 秒后开始删除；按 Enter 取消。" -ForegroundColor Yellow

        $deadline = (Get-Date).AddSeconds(1)
        while ((Get-Date) -lt $deadline) {
            try {
                if (-not [Console]::IsInputRedirected -and [Console]::KeyAvailable) {
                    $key = [Console]::ReadKey($true)
                    if ($key.Key -eq [ConsoleKey]::Enter) {
                        Write-Host
                        Write-Host "已取消默认删除。" -ForegroundColor Yellow
                        return $false
                    }
                }
            }
            catch {
                Start-Sleep -Seconds $remainingSeconds
                Write-Host
                return $true
            }

            Start-Sleep -Milliseconds 50
        }
    }

    Write-Host
    return $true
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
            @{ Expression = { Get-PathDepth -DirectoryPath $_.FullName }; Descending = $true }, `
            @{ Expression = { $_.FullName }; Descending = $true })

    $deletedList = [System.Collections.Generic.List[string]]::new()
    $skippedList = [System.Collections.Generic.List[object]]::new()
    $failedList = [System.Collections.Generic.List[object]]::new()

    $activity = '空文件夹删除'
    $currentIndex = 0

    foreach ($targetDirectory in $orderedTargetList) {
        $currentIndex++
        Write-ProgressLine -Activity $activity -Current $currentIndex -Total $orderedTargetList.Count -Status '正在删除空文件夹'

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

    Write-ProgressLine -Activity $activity -Current $orderedTargetList.Count -Total $orderedTargetList.Count -Status '空文件夹删除完成' -Force
    Complete-ProgressLine -Activity $activity

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

if ($h) {
    Show-HelpText
    exit 0
}

if ($null -eq $Path -or $Path.Count -eq 0) {
    $RootPathList = @(Read-InteractiveDirectoryList)
}
else {
    $resolvedPathResult = Resolve-InputDirectoryList -RawPathList $Path
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

if ($yes) {
    Write-Host
    Write-Host -NoNewline "扫描完成。默认计划删除空文件夹数: " -ForegroundColor White
    Write-Host $DeletionPlanList.Count -ForegroundColor Magenta

    if (Wait-DefaultDeletionCountdown -DeletionCount $DeletionPlanList.Count) {
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
