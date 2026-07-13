#requires -Version 5.1
<#
用途：
  查找文件夹中的重复文件，先计算默认删除计划，再由用户确认默认删除、手动删除或退出。

参数：
  -h     显示帮助信息。
  Path   一个或多个文件夹绝对路径；未提供时会引导交互输入。
  -a     多目录合并模式，把多个目录抽象为一个大目录进行重复文件扫描。
  -c     参考目录模式，第一个目录为参考目录，其余目录为目标目录。
  -s     包含隐藏文件和隐藏文件夹。
  -yes   跳过详细预览和菜单，输出汇总后等待 10 秒，再执行默认删除计划。

提示：
  如遇脚本无法执行的问题，先在管理员PowerShell中运行以下命令允许本地执行脚本：
  Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy RemoteSigned
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

    [Alias('a')]
    [switch]$AggregateMode,

    [Alias('c')]
    [switch]$ReferenceMode,

    [Parameter(Mandatory = $false, Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$PathList
)

# ========== 可调整配置 ==========

# 部分哈希预筛选时读取文件头尾每段的字节数；值越大越稳，扫描成本也越高。
$PartialHashSegmentByteCount = 64KB

# 哈希计算并发数；设为 1 可退回串行，机械硬盘或网络盘可适当调低。
$HashParallelThrottleLimit = 8

# 删除预览中分隔不同重复文件组的横线长度。
$PreviewSeparatorCellCount = 64

# 删除预览中分隔不同重复文件组的字符。
$PreviewSeparatorCharacter = '='

# 文本进度条宽度。
$ProgressBarCellCount = 32

# 文本进度条已完成部分的字符。
$ProgressBarFilledCharacter = '#'

# 文本进度条未完成部分的字符。
$ProgressBarEmptyCharacter = '-'

# 进度条同百分比状态下的强制刷新间隔；设为 0 或负数则只在百分比变化时刷新。
$ProgressBarTimedRefreshMilliseconds = 1000

# 使用 -yes 时的默认删除倒计时秒数，给用户留出取消窗口。
$AssumeYesGraceSeconds = 10

# -yes 倒计时期间检查 Enter 输入的间隔。
$AssumeYesInputPollIntervalMilliseconds = 100

# ========== 运行环境设置 ==========

Set-StrictMode -Version Latest

# 遇到未处理异常时立即进入 catch/退出流程，避免继续执行危险操作。
$ErrorActionPreference = 'Stop'

# ========== 参数派生选项 ==========

# 是否扫描隐藏文件和隐藏文件夹，由 -s 参数决定。
$ShouldIncludeHiddenItems = [bool]$IncludeHidden

# 是否启用无交互默认删除，由 -yes 参数决定。
$ShouldAssumeYesDeletion = [bool]$AssumeYes

# ========== 运行状态 ==========

# 记录 -yes 倒计时是否被 Enter 取消；多目录分别执行时用于停止后续目录。
$AssumeYesDeletionCancelled = $false

# 输出相关运行状态；由动态状态行和进度条函数维护，请勿手动修改。
$script:OutputRuntimeState = [pscustomobject]@{
    # 上一次动态状态行的显示宽度；用于本次刷新时清掉旧尾巴。
    DynamicStatusLastCellWidth            = 0
    # 记录上一条动态状态是否真的以内联方式输出；输出重定向或降级时不额外补换行。
    DynamicStatusLastWriteWasInline       = $false
    # 记录各进度条最近一次刷新时间；用于同百分比但耗时较长时按间隔刷新计数。
    ProgressBarLastRefreshMillisecondsByKey  = @{}
}

# ========== 输出与路径工具 ==========

function Show-HelpText {
    Write-Host @'
用途：
  查找重复文件，先预览，再选择删除或退出。

用法：
  powershell -File .\Remove-DuplicateFiles.ps1 [-a] [-s] [-yes] [Path1] [Path2 ...]
  powershell -File .\Remove-DuplicateFiles.ps1 -c [-s] [-yes] [ReferencePath] [TargetPath1] [TargetPath2 ...]

参数：
  Path   文件夹绝对路径；命令行传参时，路径包含空格、括号等 PowerShell 特殊字符，请使用英文引号包裹路径。
  -a     多目录合并模式，把多个目录视作一个大目录。
  -c     参考目录模式；第一个目录为参考目录，其余为目标目录。
  -s     包含隐藏文件和隐藏文件夹。
  -yes   跳过详细预览和菜单，输出汇总后等待 10 秒再默认删除；可按 Enter 取消。
  -h     显示帮助信息。

说明：
  无路径且未指定 -a/-c 时，会先显示模式菜单。
  预览和删除摘要会显示计划删除数量以及预计可释放空间。
  -yes 模式没有可删除项时会直接跳过，不进入倒计时。
  交互模式每轮流程完成后会返回模式菜单；命令行带路径运行时执行一次后退出。
  所有交互位置中，00 表示退出脚本，0 只表示返回或跳过。
  路径输入阶段可输入 0 返回模式菜单，输入 00 退出脚本。
  单目录和多目录合并模式可默认删除、手动删除、跳过本次操作或退出。
  多个单目录逐个操作时，0 跳过当前目录，00 退出脚本。
  参考目录模式只删除目标目录文件，可默认删除、跳过本次操作或退出。
  交互输入多个路径时可分行，也可用空格或英文分号分隔；路径含空格或英文分号时，请使用英文引号包裹路径。
  输入根目录不能是符号链接或目录联接点；删除前会重新核对保留文件与待删文件的完整 SHA-256。
  重复文件候选哈希计算默认使用 8 个并发 worker；每个 worker 批量处理文件，可修改脚本顶部 $HashParallelThrottleLimit 调整，设为 1 可退回串行。
'@
}

# 输出已启用的关键参数，避免进入交互菜单后忘记当前运行选项。
function Write-EnabledOptionNotice {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$IncludeHiddenItems,

        [Parameter(Mandatory = $true)]
        [bool]$AssumeYesDeletion
    )

    if (-not $IncludeHiddenItems -and -not $AssumeYesDeletion) {
        return
    }

    Write-Host ""
    if ($IncludeHiddenItems) {
        Write-Host "已启用 -s：扫描将包含隐藏文件和隐藏文件夹。" -ForegroundColor Cyan
    }

    if ($AssumeYesDeletion) {
        Write-Host "已启用 -yes：扫描完成后将输出汇总，并跳过详细预览和删除选择菜单。" -ForegroundColor Red
    }
}

# 输出当前执行阶段，避免大目录扫描时长时间无反馈。
function Write-StageMessage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    # 阶段消息是普通整行输出；先结束可能仍停留在同一行的动态进度，避免拼接到进度尾部。
    Complete-DynamicStatusLine
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

    Complete-DynamicStatusLine
    Write-Host -NoNewline "  $Number " -ForegroundColor Cyan
    Write-Host $Text -ForegroundColor White
}

# 输出彩色输入提示并读取一行文本。
function Read-ColoredLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt
    )

    Complete-DynamicStatusLine
    Write-Host -NoNewline $Prompt -ForegroundColor Cyan
    return [Console]::ReadLine()
}

# 估算字符在控制台中占用的单元格宽度；中文和全角字符通常占两格。
function Get-ConsoleCharacterCellWidth {
    param(
        [Parameter(Mandatory = $true)]
        [char]$Character
    )

    $codePoint = [int]$Character
    if (
        ($codePoint -ge 0x1100 -and $codePoint -le 0x115F) -or
        ($codePoint -ge 0x2E80 -and $codePoint -le 0xA4CF) -or
        ($codePoint -ge 0xAC00 -and $codePoint -le 0xD7A3) -or
        ($codePoint -ge 0xF900 -and $codePoint -le 0xFAFF) -or
        ($codePoint -ge 0xFE10 -and $codePoint -le 0xFE6F) -or
        ($codePoint -ge 0xFF00 -and $codePoint -le 0xFF60) -or
        ($codePoint -ge 0xFFE0 -and $codePoint -le 0xFFE6)
    ) {
        return 2
    }

    return 1
}

# 将文本限制在指定控制台宽度内，避免动态进度行因过长而换行。
function Get-ConsoleTextWithinCellWidth {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [int]$MaxCellWidth
    )

    $textBuilder = [System.Text.StringBuilder]::new()
    $cellWidth = 0
    foreach ($character in $Text.ToCharArray()) {
        $characterWidth = Get-ConsoleCharacterCellWidth -Character $character
        if (($cellWidth + $characterWidth) -gt $MaxCellWidth) {
            break
        }

        [void]$textBuilder.Append($character)
        $cellWidth += $characterWidth
    }

    return [pscustomobject]@{
        Text      = $textBuilder.ToString()
        CellWidth = $cellWidth
    }
}

# 刷新单行动态状态；写完整行后回到行首，减少行尾光标闪烁。
function Write-DynamicStatusLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [System.ConsoleColor]$Color
    )

    # 动态状态必须始终是单行；即使调用方传入异常文本，也不能把进度刷新打成多行。
    $statusMessage = $Message -replace '[\r\n]+', ' '

    try {
        if ([Console]::IsOutputRedirected) {
            Complete-DynamicStatusLine
            Write-Host $statusMessage -ForegroundColor $Color
            $script:OutputRuntimeState.DynamicStatusLastCellWidth = 0
            $script:OutputRuntimeState.DynamicStatusLastWriteWasInline = $false
            return
        }

        $maxLineWidth = [Math]::Max(1, [Console]::WindowWidth - 1)
        $lineText = Get-ConsoleTextWithinCellWidth -Text $statusMessage -MaxCellWidth $maxLineWidth

        # 写“当前文本 + 必要补空格”清除上一条长文本尾巴，结尾回到行首等待下次覆盖。
        $clearWidth = [Math]::Min($maxLineWidth, [Math]::Max($script:OutputRuntimeState.DynamicStatusLastCellWidth, $lineText.CellWidth))
        $paddingWidth = [Math]::Max(0, $clearWidth - $lineText.CellWidth)
        $paddingText = ' ' * $paddingWidth
        Write-Host -NoNewline "`r$($lineText.Text)$paddingText`r" -ForegroundColor $Color

        $script:OutputRuntimeState.DynamicStatusLastCellWidth = $lineText.CellWidth
        $script:OutputRuntimeState.DynamicStatusLastWriteWasInline = $true
        return
    }
    catch {
        # 某些宿主不支持读取控制台宽度；静默降级为普通整行输出，避免异常打断进度行。
        Complete-DynamicStatusLine
        Write-Host $statusMessage -ForegroundColor $Color
        $script:OutputRuntimeState.DynamicStatusLastCellWidth = 0
        $script:OutputRuntimeState.DynamicStatusLastWriteWasInline = $false
    }
}

# 结束当前动态状态行；只有确实使用了内联刷新时才补换行。
function Complete-DynamicStatusLine {
    if ($script:OutputRuntimeState.DynamicStatusLastWriteWasInline) {
        Write-Host ""
    }

    $script:OutputRuntimeState.DynamicStatusLastCellWidth = 0
    $script:OutputRuntimeState.DynamicStatusLastWriteWasInline = $false
}

# 更新百分比进度条；百分比变化或同百分比超过配置间隔时刷新，避免长任务看起来卡住。
function Write-ProgressBar {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Activity,

        [Parameter(Mandatory = $true)]
        [string]$Status,

        [Parameter(Mandatory = $true)]
        [int]$ProcessedCount,

        [Parameter(Mandatory = $true)]
        [int]$TotalCount,

        [Parameter(Mandatory = $true)]
        [ref]$LastPercent,

        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    if ($TotalCount -le 0) {
        return
    }

    $percent = [Math]::Min(100, [int][Math]::Floor(($ProcessedCount / $TotalCount) * 100))
    $refreshKey = "$Activity`0$Status"
    $currentMilliseconds = [long]([DateTime]::UtcNow.Ticks / [TimeSpan]::TicksPerMillisecond)
    $refreshIntervalMilliseconds = [Math]::Max(0, [int]$ProgressBarTimedRefreshMilliseconds)
    $shouldRefreshByPercent = ($percent -ne $LastPercent.Value)
    $shouldRefreshByTime = $false

    if ($refreshIntervalMilliseconds -gt 0 -and $script:OutputRuntimeState.ProgressBarLastRefreshMillisecondsByKey.ContainsKey($refreshKey)) {
        $lastRefreshMilliseconds = [long]$script:OutputRuntimeState.ProgressBarLastRefreshMillisecondsByKey[$refreshKey]
        $shouldRefreshByTime = (($currentMilliseconds - $lastRefreshMilliseconds) -ge $refreshIntervalMilliseconds)
    }

    if (-not $Force -and -not $shouldRefreshByPercent -and -not $shouldRefreshByTime) {
        return
    }

    $filledWidth = [Math]::Floor(($percent / 100) * $ProgressBarCellCount)
    $emptyWidth = $ProgressBarCellCount - $filledWidth
    $bar = ($ProgressBarFilledCharacter * $filledWidth) + ($ProgressBarEmptyCharacter * $emptyWidth)
    $progressText = "[进度] $Activity [$bar] $percent% $Status ($ProcessedCount / $TotalCount)"

    Write-DynamicStatusLine -Message $progressText -Color Cyan
    $LastPercent.Value = $percent
    $script:OutputRuntimeState.ProgressBarLastRefreshMillisecondsByKey[$refreshKey] = $currentMilliseconds
}

# 新建延迟输出的扫描警告列表，避免进度条刷新时被错误信息打断。
function New-DeferredScanWarningList {
    $warningList = [System.Collections.Generic.List[object]]::new()
    return , $warningList
}

# 记录扫描或哈希阶段的可跳过错误；未传入列表时退化为立即输出。
function Add-DeferredScanWarning {
    param(
        [Parameter(Mandatory = $false)]
        [System.Collections.Generic.List[object]]$WarningList,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Reason
    )

    if ($null -eq $WarningList) {
        Complete-DynamicStatusLine
        Write-Host "$($Message): $Path" -ForegroundColor Yellow
        Write-Host "  原因: $Reason" -ForegroundColor DarkGray
        return
    }

    $WarningList.Add([pscustomobject]@{
            Message = $Message
            Path    = $Path
            Reason  = $Reason
        })
}

# 在当前进度段结束后统一输出扫描警告，保持动态进度条单行刷新。
function Write-DeferredScanWarningList {
    param(
        [Parameter(Mandatory = $false)]
        [System.Collections.Generic.List[object]]$WarningList,

        [Parameter(Mandatory = $false)]
        [string]$Title = '扫描跳过汇总'
    )

    if ($null -eq $WarningList -or $WarningList.Count -eq 0) {
        return
    }

    Complete-DynamicStatusLine
    Write-Host "$($Title): $($WarningList.Count)" -ForegroundColor Yellow
    foreach ($warning in $WarningList) {
        Write-Host "$($warning.Message): $($warning.Path)" -ForegroundColor Yellow
        Write-Host "  原因: $($warning.Reason)" -ForegroundColor DarkGray
    }
}

# 输出用于区分不同预览或结果块的分隔线。
function Write-PreviewSeparator {
    Complete-DynamicStatusLine
    Write-Host ""
    Write-Host ($PreviewSeparatorCharacter * $PreviewSeparatorCellCount) -ForegroundColor DarkGray
}

# 输出阶段汇总；前置空行让它和列表内容、删除明细保持清楚间隔。
function Write-StatusSummary {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [System.ConsoleColor]$Color
    )

    Complete-DynamicStatusLine
    Write-Host ""
    Write-Host $Message -ForegroundColor $Color
}

# 检查倒计时期间是否按下 Enter；不支持读取键盘状态时静默退化为只支持 Ctrl+C。
function Test-EnterKeyPressed {
    try {
        if ([Console]::IsInputRedirected) {
            return $false
        }

        while ([Console]::KeyAvailable) {
            $keyInfo = [Console]::ReadKey($true)
            if ($keyInfo.Key -eq [ConsoleKey]::Enter) {
                return $true
            }
        }
    }
    catch {
        return $false
    }

    return $false
}

# -yes 会跳过人工确认，因此执行删除前提供醒目的中止窗口。
function Wait-AssumeYesDeletionGracePeriod {
    param(
        [Parameter(Mandatory = $false)]
        [int]$Seconds = $AssumeYesGraceSeconds
    )

    Complete-DynamicStatusLine
    Write-Host ""
    Write-Host "危险操作: 已启用 -yes，将跳过详细预览和菜单并执行默认删除。" -ForegroundColor Red
    Write-Host "如需取消，请在倒计时结束前按 Enter；也可按 Ctrl+C 强制中止。" -ForegroundColor Yellow

    $pollInterval = [Math]::Max(1, $AssumeYesInputPollIntervalMilliseconds)
    $pollCountPerSecond = [Math]::Max(1, [int][Math]::Ceiling(1000 / $pollInterval))
    for ($remainingSeconds = $Seconds; $remainingSeconds -gt 0; $remainingSeconds--) {
        Write-DynamicStatusLine -Message "倒计时 $remainingSeconds 秒后开始删除，按 Enter 取消..." -Color Yellow

        for ($pollIndex = 0; $pollIndex -lt $pollCountPerSecond; $pollIndex++) {
            if (Test-EnterKeyPressed) {
                $script:AssumeYesDeletionCancelled = $true
                Write-DynamicStatusLine -Message '已取消 -yes 默认删除，未删除任何文件。' -Color Yellow
                Complete-DynamicStatusLine
                return $false
            }

            Start-Sleep -Milliseconds $pollInterval
        }
    }

    Write-DynamicStatusLine -Message '倒计时结束，开始执行默认删除。' -Color Magenta
    Complete-DynamicStatusLine
    return $true
}

# 兼容交互输入时复制带首尾英文引号的路径；这里只移除成对包裹符号。
function ConvertTo-UnquotedPathText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathText
    )

    $normalizedPathText = $PathText.Trim()
    if ($normalizedPathText.Length -lt 2) {
        return $normalizedPathText
    }

    $firstCharacterCode = [int][char]$normalizedPathText[0]
    $lastCharacterCode = [int][char]$normalizedPathText[$normalizedPathText.Length - 1]
    if (($firstCharacterCode -eq 34 -and $lastCharacterCode -eq 34) -or
        ($firstCharacterCode -eq 39 -and $lastCharacterCode -eq 39)) {
        return $normalizedPathText.Substring(1, $normalizedPathText.Length - 2).Trim()
    }

    return $normalizedPathText
}

# 判断输入文本是否为 Windows 绝对路径；Windows PowerShell 5.1 缺少 IsPathFullyQualified 时使用兼容判断。
function Test-WindowsAbsolutePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathText
    )

    try {
        $method = [System.IO.Path].GetMethod('IsPathFullyQualified', [type[]]@([string]))
        if ($null -ne $method) {
            return [bool]$method.Invoke($null, @($PathText))
        }
    }
    catch {
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($PathText)) {
        return $false
    }

    return ($PathText -match '^[a-zA-Z]:[\\/]' -or $PathText -match '^[\\/]{2}[^\\/]+[\\/]+[^\\/]')
}

# 拆分交互输入的路径行：英文引号用于包裹含空格或分隔符的路径；路径可用英文分号分隔，也可在下一个片段看起来是绝对路径时按空格分隔。
function Split-InteractivePathInput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathInput
    )

    $trimmedPathInput = $PathInput.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmedPathInput)) {
        return @()
    }

    $pathPartList = New-Object System.Collections.Generic.List[string]
    $currentInputPart = New-Object System.Text.StringBuilder
    $quoteCloseByOpen = @{}
    $quoteCloseByOpen.Add([string][char]34, [string][char]34)
    $quoteCloseByOpen.Add([string][char]39, [string][char]39)
    $activeClosingQuote = $null

    for ($index = 0; $index -lt $trimmedPathInput.Length; $index++) {
        $currentChar = $trimmedPathInput[$index].ToString()

        if ($null -ne $activeClosingQuote) {
            if ($currentChar -eq $activeClosingQuote) {
                $activeClosingQuote = $null
            }
            else {
                [void]$currentInputPart.Append($currentChar)
            }
            continue
        }

        $currentInputText = $currentInputPart.ToString()
        $canStartQuotedPath = [string]::IsNullOrWhiteSpace($currentInputText)
        if (-not $canStartQuotedPath) {
            $canStartQuotedPath = [char]::IsWhiteSpace($currentInputText[$currentInputText.Length - 1])
        }

        if ($canStartQuotedPath -and $quoteCloseByOpen.ContainsKey($currentChar)) {
            $activeClosingQuote = $quoteCloseByOpen[$currentChar]
            continue
        }

        if ($currentChar -eq ';') {
            $pathPart = $currentInputPart.ToString().Trim()
            if (-not [string]::IsNullOrWhiteSpace($pathPart)) {
                $pathPartList.Add($pathPart)
            }
            [void]$currentInputPart.Clear()
            continue
        }

        [void]$currentInputPart.Append($currentChar)
    }

    $lastPathPart = $currentInputPart.ToString().Trim()
    if (-not [string]::IsNullOrWhiteSpace($lastPathPart)) {
        $pathPartList.Add($lastPathPart)
    }

    $resultPathList = New-Object System.Collections.Generic.List[string]
    $quoteStartPattern = @(
        [regex]::Escape([string][char]34)
        [regex]::Escape([string][char]39)
    ) -join '|'
    $absolutePathSeparatorPattern = '\s+(?=(?:' + $quoteStartPattern + ')?(?:[a-zA-Z]:[\\/]|[\\/]{2}))'
    foreach ($pathPart in $pathPartList) {
        foreach ($pathSegment in [regex]::Split($pathPart, $absolutePathSeparatorPattern)) {
            $normalizedPath = ConvertTo-UnquotedPathText -PathText $pathSegment
            if (-not [string]::IsNullOrWhiteSpace($normalizedPath)) {
                $resultPathList.Add($normalizedPath)
            }
        }
    }

    return $resultPathList.ToArray()
}

# 校验输入路径是否为存在的 Windows 绝对目录，并返回规范化后的完整路径。
function Resolve-InputDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$ParameterName
    )

    $normalizedPathText = ConvertTo-UnquotedPathText -PathText $Path

    if (-not (Test-WindowsAbsolutePath -PathText $normalizedPathText)) {
        throw "$ParameterName 必须是 Windows 文件夹绝对路径: $normalizedPathText"
    }

    try {
        $fullPath = [System.IO.Path]::GetFullPath($normalizedPathText)
    }
    catch {
        throw "$ParameterName 无法读取: $normalizedPathText。原因: $($_.Exception.Message)"
    }

    if ([System.IO.File]::Exists($fullPath)) {
        throw "$ParameterName 必须是文件夹: $normalizedPathText"
    }

    if (-not [System.IO.Directory]::Exists($fullPath)) {
        throw "$ParameterName 不存在或无法访问: $normalizedPathText。请确认路径存在；命令行传参时，路径包含空格、括号等 PowerShell 特殊字符，请使用英文引号包裹路径；交互输入多个路径可分行，或在同一行用空格/英文分号分隔。"
    }

    $directoryInfo = [System.IO.DirectoryInfo]::new($fullPath)
    if (($directoryInfo.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "$ParameterName 不能是符号链接或联接点: $normalizedPathText"
    }

    return $directoryInfo.FullName
}

# 逐个校验路径列表，自动去重后返回规范化后的完整目录路径。
function Resolve-InputDirectoryList {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$PathList,

        [Parameter(Mandatory = $false)]
        [string]$ParameterNamePrefix = 'Path',

        [Parameter(Mandatory = $false)]
        [int]$StartIndex = 1,

        [Parameter(Mandatory = $false)]
        [string[]]$ExistingPathList = @()
    )

    $resolvedDirectoryList = New-Object System.Collections.Generic.List[string]
    $resolvedDirectoryKeySet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($existingPath in $ExistingPathList) {
        [void]$resolvedDirectoryKeySet.Add((ConvertTo-NormalizedPath -Path $existingPath))
    }

    for ($index = 0; $index -lt $PathList.Count; $index++) {
        $resolvedDirectory = Resolve-InputDirectory -Path $PathList[$index] -ParameterName "$($ParameterNamePrefix)[$($StartIndex + $index)]"
        $directoryKey = ConvertTo-NormalizedPath -Path $resolvedDirectory
        if ($resolvedDirectoryKeySet.Add($directoryKey)) {
            $resolvedDirectoryList.Add($resolvedDirectory)
        }
    }

    try {
        Assert-NoParentChildDirectorySet `
            -DirectoryPathList $resolvedDirectoryList.ToArray() `
            -ExistingDirectoryPathList $ExistingPathList
    }
    catch {
        throw '输入目录不能互为父子目录。'
    }

    return @(
        $resolvedDirectoryList.ToArray()
    )
}

# 拆分并校验交互输入的一行路径，返回本行识别数量和规范化后的目录列表。
function Resolve-InteractivePathInputLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathInput,

        [Parameter(Mandatory = $true)]
        [int]$StartIndex,

        [Parameter(Mandatory = $false)]
        [string[]]$ExistingPathList = @()
    )

    $pathInputList = @(Split-InteractivePathInput -PathInput $PathInput)
    if ($pathInputList.Count -eq 0) {
        throw '未识别到可用路径。'
    }

    $resolvedPathInputList = @(Resolve-InputDirectoryList `
            -PathList $pathInputList `
            -ParameterNamePrefix 'Path' `
            -StartIndex $StartIndex `
            -ExistingPathList $ExistingPathList)

    return [pscustomobject]@{
        InputCount = $pathInputList.Count
        PathList   = $resolvedPathInputList
    }
}

# 交互读取一个或多个目录路径；0 返回模式菜单，00 直接退出脚本。
function Read-InteractivePathList {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModePrompt,

        [Parameter(Mandatory = $true)]
        [int]$MinimumCount
    )

    $inputPathList = New-Object System.Collections.Generic.List[string]
    Write-Host "进入$ModePrompt。" -ForegroundColor Cyan
    Write-Host "请输入目录绝对路径。可在同一行输入多个路径。" -ForegroundColor Yellow
    Write-Host "多个路径可用空格或英文分号分隔；路径含空格或英文分号时，请使用英文引号包裹路径。" -ForegroundColor DarkGray
    Write-Host "直接回车开始执行；输入 0 返回上级菜单；输入 00 退出脚本。" -ForegroundColor DarkGray

    while ($true) {
        $pathInputRaw = Read-ColoredLine -Prompt "Path$($inputPathList.Count + 1): "
        if ($null -eq $pathInputRaw) {
            return [pscustomobject]@{
                Action   = 'Exit'
                PathList = @()
            }
        }

        $pathInput = $pathInputRaw.Trim()
        if ($pathInput -eq '00') {
            return [pscustomobject]@{
                Action   = 'Exit'
                PathList = @()
            }
        }

        if ($pathInput -eq '0') {
            return [pscustomobject]@{
                Action   = 'Back'
                PathList = @()
            }
        }

        if ([string]::IsNullOrWhiteSpace($pathInput)) {
            if ($inputPathList.Count -eq 0) {
                Write-Host "尚未输入目录；请输入目录，或输入 0 返回上级菜单，输入 00 退出脚本。" -ForegroundColor Yellow
                continue
            }

            if ($inputPathList.Count -lt $MinimumCount) {
                Write-Host "当前模式至少需要输入 $MinimumCount 个目录。" -ForegroundColor Red
                continue
            }

            return [pscustomobject]@{
                Action   = 'Submit'
                PathList = $inputPathList.ToArray()
            }
        }

        try {
            $lineInputResult = Resolve-InteractivePathInputLine `
                -PathInput $pathInput `
                -StartIndex ($inputPathList.Count + 1) `
                -ExistingPathList $inputPathList.ToArray()
        }
        catch {
            Write-Host $_.Exception.Message -ForegroundColor Yellow
            continue
        }

        $resolvedPathInputList = @($lineInputResult.PathList)
        if ($resolvedPathInputList.Count -eq 0) {
            Write-Host "输入路径已存在，本次未新增。" -ForegroundColor DarkGray
            continue
        }

        foreach ($resolvedInputPath in $resolvedPathInputList) {
            $inputPathList.Add($resolvedInputPath)
        }

        if ($lineInputResult.InputCount -gt 1) {
            Write-Host "识别到 $($resolvedPathInputList.Count) 个新路径。" -ForegroundColor DarkGray
        }
    }
}

# 将路径标准化为便于比较的形式；文件路径同样适用，去除尾部分隔符对文件路径是空操作。
function ConvertTo-NormalizedPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return [System.IO.Path]::GetFullPath($Path).TrimEnd([char[]]@('\', '/'))
}

# 给目录路径追加结尾分隔符，避免 C:\A 和 C:\AB 这种前缀误判。
function Add-TrailingDirectorySeparator {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ($Path.EndsWith('\', [System.StringComparison]::Ordinal) -or
        $Path.EndsWith('/', [System.StringComparison]::Ordinal)) {
        return $Path
    }

    return "$Path\"
}

# 使用 Uri 计算相对路径，兼容 Windows PowerShell 5.1 中缺失的 Path.GetRelativePath。
function Get-RelativePathCompat {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath,

        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    $rootFullPath = Add-TrailingDirectorySeparator -Path ([System.IO.Path]::GetFullPath($RootPath))
    $targetFullPath = [System.IO.Path]::GetFullPath($TargetPath)
    $rootUri = New-Object System.Uri($rootFullPath)
    $targetUri = New-Object System.Uri($targetFullPath)
    $relativeUri = $rootUri.MakeRelativeUri($targetUri)
    return [System.Uri]::UnescapeDataString($relativeUri.ToString()).Replace('/', [System.IO.Path]::DirectorySeparatorChar)
}

# 判断两个目录是否存在父子关系；相同目录不算父子关系，重复路径由路径解析去重处理。
function Test-ParentChildDirectoryPair {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LeftPath,

        [Parameter(Mandatory = $true)]
        [string]$RightPath
    )

    $normalizedLeftPath = ConvertTo-NormalizedPath -Path $LeftPath
    $normalizedRightPath = ConvertTo-NormalizedPath -Path $RightPath

    if ($normalizedLeftPath.Equals($normalizedRightPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $false
    }

    $leftPathPrefix = Add-TrailingDirectorySeparator -Path $normalizedLeftPath
    $rightPathPrefix = Add-TrailingDirectorySeparator -Path $normalizedRightPath

    return ($normalizedRightPath.StartsWith($leftPathPrefix, [System.StringComparison]::OrdinalIgnoreCase) -or
        $normalizedLeftPath.StartsWith($rightPathPrefix, [System.StringComparison]::OrdinalIgnoreCase))
}

# 拒绝目录集合中的父子目录关系；相同目录应由路径解析阶段去重，不在这里视为错误。
function Assert-NoParentChildDirectorySet {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$DirectoryPathList,

        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [string[]]$ExistingDirectoryPathList = @()
    )

    for ($existingIndex = 0; $existingIndex -lt $ExistingDirectoryPathList.Count; $existingIndex++) {
        for ($directoryIndex = 0; $directoryIndex -lt $DirectoryPathList.Count; $directoryIndex++) {
            if (Test-ParentChildDirectoryPair -LeftPath $ExistingDirectoryPathList[$existingIndex] -RightPath $DirectoryPathList[$directoryIndex]) {
                throw '目录不能互为父子目录。'
            }
        }
    }

    for ($leftIndex = 0; $leftIndex -lt $DirectoryPathList.Count; $leftIndex++) {
        for ($rightIndex = $leftIndex + 1; $rightIndex -lt $DirectoryPathList.Count; $rightIndex++) {
            if (Test-ParentChildDirectoryPair -LeftPath $DirectoryPathList[$leftIndex] -RightPath $DirectoryPathList[$rightIndex]) {
                throw '目录不能互为父子目录。'
            }
        }
    }
}

# ========== 哈希与重复文件识别 ==========

# 判断部分哈希是否已经覆盖完整文件；覆盖时可直接把部分哈希当作最终哈希。
function Test-PartialHashCoversFullFile {
    param(
        [Parameter(Mandatory = $true)]
        [long]$Length
    )

    return $Length -le ([int64]$PartialHashSegmentByteCount * 2)
}

# 计算文件首尾片段的 SHA-256，用作快速筛选候选重复文件。
function Get-PartialContentHash {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    $fileStream = [System.IO.File]::Open($File.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)

    try {
        # 文件不超过两段采样总长时直接计算完整文件，避免头尾片段重叠读取。
        if (Test-PartialHashCoversFullFile -Length $File.Length) {
            $hashBytes = $sha256.ComputeHash($fileStream)
            return [BitConverter]::ToString($hashBytes).Replace('-', '').ToLowerInvariant()
        }

        # 大文件只做快速预筛选；真正删除前仍会使用完整 SHA-256 确认。
        $hashBuffer = [byte[]]::new($PartialHashSegmentByteCount)

        $firstRead = $fileStream.Read($hashBuffer, 0, $hashBuffer.Length)
        if ($firstRead -gt 0) {
            [void]$sha256.TransformBlock($hashBuffer, 0, $firstRead, $null, 0)
        }

        if ($File.Length -gt $PartialHashSegmentByteCount) {
            $tailOffset = [Math]::Max(0, $File.Length - $PartialHashSegmentByteCount)
            [void]$fileStream.Seek($tailOffset, [System.IO.SeekOrigin]::Begin)
            $lastRead = $fileStream.Read($hashBuffer, 0, $hashBuffer.Length)
            if ($lastRead -gt 0) {
                [void]$sha256.TransformBlock($hashBuffer, 0, $lastRead, $null, 0)
            }
        }

        [void]$sha256.TransformFinalBlock([byte[]]::new(0), 0, 0)
        return [BitConverter]::ToString($sha256.Hash).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $fileStream.Dispose()
        $sha256.Dispose()
    }
}

# 计算完整文件内容的 SHA-256，用于最终确认文件内容完全一致。
function Get-FullContentHash {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    $fileStream = [System.IO.File]::Open($File.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)

    try {
        $hashBytes = $sha256.ComputeHash($fileStream)
        return [BitConverter]::ToString($hashBytes).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $fileStream.Dispose()
        $sha256.Dispose()
    }
}

# 在独立 runspace 中批量计算文件哈希；脚本块不依赖外部函数，便于受控并行执行。
$script:ContentHashWorkerScript = {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$FilePathList,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Partial', 'Full')]
        [string]$HashKind,

        [Parameter(Mandatory = $true)]
        [int]$PartialHashSegmentByteCount
    )

    foreach ($filePath in $FilePathList) {
        $sha256 = $null
        $fileStream = $null
        try {
            $sha256 = [System.Security.Cryptography.SHA256]::Create()
            $fileStream = [System.IO.File]::Open($filePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)

            if ($HashKind -eq 'Partial') {
                if ($fileStream.Length -le ([int64]$PartialHashSegmentByteCount * 2)) {
                    $hashBytes = $sha256.ComputeHash($fileStream)
                }
                else {
                    $hashBuffer = [byte[]]::new($PartialHashSegmentByteCount)
                    $firstRead = $fileStream.Read($hashBuffer, 0, $hashBuffer.Length)
                    if ($firstRead -gt 0) {
                        [void]$sha256.TransformBlock($hashBuffer, 0, $firstRead, $null, 0)
                    }

                    $tailOffset = [Math]::Max(0, $fileStream.Length - $PartialHashSegmentByteCount)
                    [void]$fileStream.Seek($tailOffset, [System.IO.SeekOrigin]::Begin)
                    $lastRead = $fileStream.Read($hashBuffer, 0, $hashBuffer.Length)
                    if ($lastRead -gt 0) {
                        [void]$sha256.TransformBlock($hashBuffer, 0, $lastRead, $null, 0)
                    }

                    [void]$sha256.TransformFinalBlock([byte[]]::new(0), 0, 0)
                    $hashBytes = $sha256.Hash
                }
            }
            else {
                $hashBytes = $sha256.ComputeHash($fileStream)
            }

            [pscustomobject]@{
                Success = $true
                Path    = $filePath
                Hash    = [BitConverter]::ToString($hashBytes).Replace('-', '').ToLowerInvariant()
                Message = $null
            }
        }
        catch {
            [pscustomobject]@{
                Success = $false
                Path    = $filePath
                Hash    = $null
                Message = $_.Exception.Message
            }
        }
        finally {
            if ($null -ne $fileStream) {
                $fileStream.Dispose()
            }

            if ($null -ne $sha256) {
                $sha256.Dispose()
            }
        }
    }
}

# 启动一个受控并行哈希批次任务。
function Start-ContentHashRunspaceBatchJob {
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.Runspaces.RunspacePool]$RunspacePool,

        [Parameter(Mandatory = $true)]
        [string[]]$FilePathList,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Partial', 'Full')]
        [string]$HashKind
    )

    $powerShell = [System.Management.Automation.PowerShell]::Create()
    $powerShell.RunspacePool = $RunspacePool
    [void]$powerShell.AddScript($script:ContentHashWorkerScript.ToString())
    [void]$powerShell.AddArgument([string[]]$FilePathList)
    [void]$powerShell.AddArgument($HashKind)
    [void]$powerShell.AddArgument([int]$PartialHashSegmentByteCount)

    return [pscustomobject]@{
        PowerShell = $powerShell
        Handle     = $powerShell.BeginInvoke()
        PathList   = $FilePathList
    }
}

# 将哈希结果加入结果列表或警告列表，并刷新主线程进度。
function Add-ContentHashResult {
    param(
        [Parameter(Mandatory = $true)]
        [object]$HashResult,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$HashRecordList,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$WarningList,

        [Parameter(Mandatory = $true)]
        [string]$WarningMessage,

        [Parameter(Mandatory = $true)]
        [string]$Activity,

        [Parameter(Mandatory = $true)]
        [string]$Status,

        [Parameter(Mandatory = $true)]
        [int]$TotalCount,

        [Parameter(Mandatory = $true)]
        [ref]$ProcessedCount,

        [Parameter(Mandatory = $true)]
        [ref]$LastPercent
    )

    $ProcessedCount.Value++
    Write-ProgressBar -Activity $Activity -Status $Status -ProcessedCount $ProcessedCount.Value -TotalCount $TotalCount -LastPercent $LastPercent

    if ($HashResult.Success) {
        $HashRecordList.Add([pscustomobject]@{
            File = [System.IO.FileInfo]::new($HashResult.Path)
            Hash = $HashResult.Hash
        })
        return
    }

    Add-DeferredScanWarning -WarningList $WarningList -Message $WarningMessage -Path $HashResult.Path -Reason $HashResult.Message
}

# 按配置的并发数计算文件哈希；并发只用于读取和哈希，进度和警告仍由主线程统一处理。
function Get-ContentHashRecordList {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.IO.FileInfo[]]$FileList,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Partial', 'Full')]
        [string]$HashKind,

        [Parameter(Mandatory = $true)]
        [string]$Activity,

        [Parameter(Mandatory = $true)]
        [string]$Status,

        [Parameter(Mandatory = $true)]
        [int]$TotalCount,

        [Parameter(Mandatory = $true)]
        [ref]$ProcessedCount,

        [Parameter(Mandatory = $true)]
        [ref]$LastPercent,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$WarningList,

        [Parameter(Mandatory = $true)]
        [string]$WarningMessage
    )

    $hashRecordList = New-Object System.Collections.Generic.List[object]
    if ($FileList.Count -eq 0) {
        return $hashRecordList.ToArray()
    }

    $effectiveThrottleLimit = [Math]::Max(1, [int]$HashParallelThrottleLimit)
    if ($effectiveThrottleLimit -le 1 -or $FileList.Count -eq 1) {
        foreach ($file in $FileList) {
            try {
                if ($HashKind -eq 'Partial') {
                    $hashValue = Get-PartialContentHash -File $file
                }
                else {
                    $hashValue = Get-FullContentHash -File $file
                }

                Add-ContentHashResult `
                    -HashResult ([pscustomobject]@{ Success = $true; Path = $file.FullName; Hash = $hashValue; Message = $null }) `
                    -HashRecordList $hashRecordList `
                    -WarningList $WarningList `
                    -WarningMessage $WarningMessage `
                    -Activity $Activity `
                    -Status $Status `
                    -TotalCount $TotalCount `
                    -ProcessedCount $ProcessedCount `
                    -LastPercent $LastPercent
            }
            catch {
                Add-ContentHashResult `
                    -HashResult ([pscustomobject]@{ Success = $false; Path = $file.FullName; Hash = $null; Message = $_.Exception.Message }) `
                    -HashRecordList $hashRecordList `
                    -WarningList $WarningList `
                    -WarningMessage $WarningMessage `
                    -Activity $Activity `
                    -Status $Status `
                    -TotalCount $TotalCount `
                    -ProcessedCount $ProcessedCount `
                    -LastPercent $LastPercent
            }
        }

        return $hashRecordList.ToArray()
    }

    $batchCount = [Math]::Min($effectiveThrottleLimit, $FileList.Count)
    $batchPathBuilderList = New-Object System.Collections.Generic.List[object]
    $batchByteCountList = New-Object 'long[]' $batchCount
    for ($batchIndex = 0; $batchIndex -lt $batchCount; $batchIndex++) {
        $batchPathBuilderList.Add((New-Object System.Collections.Generic.List[string]))
    }

    if ($HashKind -eq 'Full') {
        # 完整哈希成本主要由字节数决定；按文件大小贪心分配，减少大文件集中到同一 worker 的长尾。
        foreach ($file in @($FileList | Sort-Object -Property Length -Descending)) {
            $targetBatchIndex = 0
            for ($batchIndex = 1; $batchIndex -lt $batchCount; $batchIndex++) {
                if ($batchByteCountList[$batchIndex] -lt $batchByteCountList[$targetBatchIndex]) {
                    $targetBatchIndex = $batchIndex
                }
            }

            $batchPathBuilderList[$targetBatchIndex].Add($file.FullName)
            $batchByteCountList[$targetBatchIndex] += [long]$file.Length
        }
    }
    else {
        # 部分哈希读取量接近固定，按文件数量轮询分配即可。
        for ($fileIndex = 0; $fileIndex -lt $FileList.Count; $fileIndex++) {
            $batchPathBuilderList[$fileIndex % $batchCount].Add($FileList[$fileIndex].FullName)
        }
    }

    $filePathBatchList = New-Object 'System.Collections.Generic.List[string[]]'
    foreach ($batchPathBuilder in $batchPathBuilderList) {
        if ($batchPathBuilder.Count -gt 0) {
            $filePathBatchList.Add($batchPathBuilder.ToArray())
        }
    }

    $runspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $effectiveThrottleLimit)
    $pendingJobList = New-Object System.Collections.Generic.List[object]
    $nextBatchIndex = 0

    try {
        $runspacePool.Open()

        while ($nextBatchIndex -lt $filePathBatchList.Count -or $pendingJobList.Count -gt 0) {
            while ($nextBatchIndex -lt $filePathBatchList.Count -and $pendingJobList.Count -lt $effectiveThrottleLimit) {
                $pendingJobList.Add((Start-ContentHashRunspaceBatchJob -RunspacePool $runspacePool -FilePathList $filePathBatchList[$nextBatchIndex] -HashKind $HashKind))
                $nextBatchIndex++
            }

            $completedAnyJob = $false
            for ($jobIndex = $pendingJobList.Count - 1; $jobIndex -ge 0; $jobIndex--) {
                $pendingJob = $pendingJobList[$jobIndex]
                if (-not $pendingJob.Handle.IsCompleted) {
                    continue
                }

                $completedAnyJob = $true
                try {
                    $jobOutputList = @($pendingJob.PowerShell.EndInvoke($pendingJob.Handle))
                    if ($jobOutputList.Count -gt 0) {
                        $hashResultList = $jobOutputList
                    }
                    else {
                        $hashResultList = @(
                            foreach ($filePath in $pendingJob.PathList) {
                                [pscustomobject]@{
                                    Success = $false
                                    Path    = $filePath
                                    Hash    = $null
                                    Message = '哈希任务未返回结果。'
                                }
                            }
                        )
                    }
                }
                catch {
                    $hashResultList = @(
                        foreach ($filePath in $pendingJob.PathList) {
                            [pscustomobject]@{
                                Success = $false
                                Path    = $filePath
                                Hash    = $null
                                Message = $_.Exception.Message
                            }
                        }
                    )
                }
                finally {
                    $pendingJob.PowerShell.Dispose()
                    $pendingJobList.RemoveAt($jobIndex)
                }

                foreach ($hashResult in $hashResultList) {
                    Add-ContentHashResult `
                        -HashResult $hashResult `
                        -HashRecordList $hashRecordList `
                        -WarningList $WarningList `
                        -WarningMessage $WarningMessage `
                        -Activity $Activity `
                        -Status $Status `
                        -TotalCount $TotalCount `
                        -ProcessedCount $ProcessedCount `
                        -LastPercent $LastPercent
                }
            }

            if (-not $completedAnyJob -and $pendingJobList.Count -gt 0) {
                Start-Sleep -Milliseconds 50
            }
        }
    }
    finally {
        foreach ($pendingJob in $pendingJobList) {
            try {
                $pendingJob.PowerShell.Stop()
            }
            catch {
                Write-Debug "停止哈希任务失败: $($_.Exception.Message)"
            }
            finally {
                $pendingJob.PowerShell.Dispose()
            }
        }

        $runspacePool.Close()
        $runspacePool.Dispose()
    }

    return $hashRecordList.ToArray()
}

# 递归获取目录下所有普通文件；默认不包含隐藏项，传入 -s 时包含隐藏文件和隐藏文件夹。
function Get-ScannedFileList {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath,

        [Parameter(Mandatory = $false)]
        [string]$ProgressLabel = '目录',

        [Parameter(Mandatory = $false)]
        [switch]$SuppressScanStageMessages,

        [Parameter(Mandatory = $false)]
        [System.Collections.Generic.List[object]]$WarningList
    )

    if (-not $SuppressScanStageMessages) {
        Write-StageMessage "开始扫描$($ProgressLabel): $RootPath"
    }

    $scanErrorList = $null
    if ($ShouldIncludeHiddenItems) {
        $scannedFiles = @(Get-ChildItem -LiteralPath $RootPath -File -Recurse -Force -ErrorAction SilentlyContinue -ErrorVariable scanErrorList)
    }
    else {
        $scannedFiles = @(Get-ChildItem -LiteralPath $RootPath -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable scanErrorList)
    }

    $scanErrors = @($scanErrorList)
    if ($scanErrors.Count -gt 0) {
        $scanWarningList = if ($null -ne $WarningList) { $WarningList } else { New-DeferredScanWarningList }
        foreach ($scanError in $scanErrors) {
            Add-DeferredScanWarning -WarningList $scanWarningList -Message '扫描跳过' -Path $scanError.TargetObject -Reason $scanError.Exception.Message
        }

        if ($null -eq $WarningList) {
            if ($SuppressScanStageMessages) {
                Complete-DynamicStatusLine
            }
            Write-DeferredScanWarningList -WarningList $scanWarningList -Title "$($ProgressLabel)扫描跳过汇总"
        }
    }

    $hiddenScopeText = if ($ShouldIncludeHiddenItems) { '包含隐藏项' } else { '不包含隐藏项' }
    if (-not $SuppressScanStageMessages) {
        Write-StageMessage "$($ProgressLabel)扫描完成，文件数: $($scannedFiles.Count)，$hiddenScopeText"
    }

    return $scannedFiles
}

# ========== 路径显示与默认保留规则 ==========

# 将完整文件路径转换为便于日志阅读的相对路径，可附加目录名前缀。
function Get-RelativePathText {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File,

        [Parameter(Mandatory = $true)]
        [string]$RootPath,

        [Parameter(Mandatory = $false)]
        [string]$PathPrefix
    )

    $relativePathText = Get-RelativePathCompat -RootPath $RootPath -TargetPath $File.FullName
    if ($relativePathText.StartsWith('..', [System.StringComparison]::Ordinal) -or [System.IO.Path]::IsPathRooted($relativePathText)) {
        $relativePathText = $File.Name
    }

    if ([string]::IsNullOrWhiteSpace($PathPrefix)) {
        return $relativePathText
    }

    return "$PathPrefix\$relativePathText"
}

# 获取目录最后一级名称，用于多目录日志前缀。
function Get-DirectoryLabel {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    $name = Split-Path -Leaf $RootPath.TrimEnd('\')
    if ([string]::IsNullOrWhiteSpace($name)) {
        return $RootPath.TrimEnd('\')
    }

    return $name
}

# 使用 .NET 删除文件；删除前移除只读属性，避免大量删除时引入 Remove-Item 的额外开销。
function Remove-FileSystemFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    if (-not [System.IO.File]::Exists($FilePath)) {
        throw '文件已不存在。'
    }

    [System.IO.File]::SetAttributes($FilePath, [System.IO.FileAttributes]::Normal)
    [System.IO.File]::Delete($FilePath)
    if ([System.IO.File]::Exists($FilePath)) {
        throw '删除后文件仍存在。'
    }
}

# 获取文件所在目录的规范化路径，用于默认保留优先级判断。
function Get-FileParentDirectoryPath {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    return ConvertTo-NormalizedPath -Path (Split-Path -Parent $File.FullName)
}

# 统计目录路径层级；层级越少，默认保留优先级越高。
function Get-DirectoryPathDepth {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $depth = 0
    foreach ($pathSegment in $Path.Split([char[]]@('\', '/'))) {
        if (-not [string]::IsNullOrWhiteSpace($pathSegment)) {
            $depth++
        }
    }

    return $depth
}

# 按默认保留规则排序文件：目录越靠上越优先，同目录内文件名越短越优先。
function Get-FileListByKeepPriority {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$FileList
    )

    $keepPriorityRecordList = @(
        foreach ($file in $FileList) {
            $parentDirectoryPath = Get-FileParentDirectoryPath -File $file
            [pscustomobject]@{
                File             = $file
                ParentDepth      = Get-DirectoryPathDepth -Path $parentDirectoryPath
                ParentPathLength = $parentDirectoryPath.Length
                NameLength       = $file.Name.Length
                Name             = $file.Name
                FullName         = $file.FullName
            }
        }
    )

    return $keepPriorityRecordList |
        Sort-Object @{ Expression = { $_.ParentDepth }; Ascending = $true },
                    @{ Expression = { $_.ParentPathLength }; Ascending = $true },
                    @{ Expression = { $_.NameLength }; Ascending = $true },
                    @{ Expression = { $_.Name }; Ascending = $true },
                    @{ Expression = { $_.FullName }; Ascending = $true } |
        ForEach-Object { $_.File }
}

# 从一组重复文件中选出默认应保留的文件。
function Select-DefaultKeepFile {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$FileList
    )

    $orderedFileList = @(Get-FileListByKeepPriority -FileList $FileList)
    if ($orderedFileList.Count -eq 0) {
        return $null
    }

    return $orderedFileList[0]
}

# 按单目录根路径、目录前缀或预先建立的映射获取文件显示路径。
function Get-FileDisplayPath {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File,

        [Parameter(Mandatory = $false)]
        [string]$RootPath,

        [Parameter(Mandatory = $false)]
        [string]$PathPrefix,

        [Parameter(Mandatory = $false)]
        [hashtable]$DisplayPathByFullName
    )

    if ($null -ne $DisplayPathByFullName -and $DisplayPathByFullName.ContainsKey($File.FullName)) {
        return $DisplayPathByFullName[$File.FullName]
    }

    if (-not [string]::IsNullOrWhiteSpace($RootPath)) {
        return Get-RelativePathText -File $File -RootPath $RootPath -PathPrefix $PathPrefix
    }

    return $File.FullName
}

# 将待删除文件封装为包含显示路径、预期哈希和保留文件列表的删除项。
function New-DeletionItemList {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$FileList,

        [Parameter(Mandatory = $false)]
        [string]$RootPath,

        [Parameter(Mandatory = $false)]
        [string]$PathPrefix,

        [Parameter(Mandatory = $false)]
        [hashtable]$DisplayPathByFullName,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedHash,

        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$KeepFileList
    )

    return @(
        $FileList | ForEach-Object {
            $displayPath = Get-FileDisplayPath -File $_ -RootPath $RootPath -PathPrefix $PathPrefix -DisplayPathByFullName $DisplayPathByFullName
            $byteCount = 0
            try {
                $byteCount = [long]$_.Length
            }
            catch {
                Write-Debug "读取文件大小失败: $($_.Exception.Message)"
            }

            [pscustomobject]@{
                File         = $_
                DisplayPath  = $displayPath
                ByteCount    = $byteCount
                ExpectedHash = $ExpectedHash
                KeepFileList = $KeepFileList
            }
        }
    )
}

# 将文件加入指定键的小桶；用哈希表直接分组，避免大批量 Group-Object 的管道和对象开销。
function Add-FileToGroupBucket {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$FileListByKey,

        [Parameter(Mandatory = $true)]
        [object]$Key,

        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    if (-not $FileListByKey.ContainsKey($Key)) {
        $FileListByKey[$Key] = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    }

    $FileListByKey[$Key].Add($File)
}

# 按文件大小、部分哈希、完整哈希分层筛选出内容完全一致的重复文件组。
function Find-DuplicateFileGroup {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$FileList,

        [Parameter(Mandatory = $false)]
        [string]$ProgressLabel = '文件'
    )

    # 分层匹配可以避免在大目录中对每个文件都计算完整哈希。
    Write-StageMessage "$($ProgressLabel)按文件大小分组中..."
    $filesByLength = @{}
    $processedLengthFileCount = 0
    $lastLengthPercent = -1
    $lengthProgressActivity = "$($ProgressLabel)文件大小分组"

    foreach ($file in $FileList) {
        $processedLengthFileCount++
        Write-ProgressBar -Activity $lengthProgressActivity -Status '正在按文件大小归类' -ProcessedCount $processedLengthFileCount -TotalCount $FileList.Count -LastPercent ([ref]$lastLengthPercent)

        if (-not $filesByLength.ContainsKey($file.Length)) {
            $filesByLength[$file.Length] = New-Object System.Collections.Generic.List[System.IO.FileInfo]
        }
        $filesByLength[$file.Length].Add($file)
    }

    Complete-DynamicStatusLine

    # 只有大小相同的文件才可能重复；不同大小的文件无需继续计算哈希。
    $sameLengthGroupList = New-Object System.Collections.Generic.List[object]
    $partialHashCandidateCount = 0
    foreach ($size in $filesByLength.Keys) {
        $sameLengthFileList = $filesByLength[$size]
        if ($sameLengthFileList.Count -gt 1) {
            $sameLengthGroupList.Add($sameLengthFileList)
            $partialHashCandidateCount += $sameLengthFileList.Count
        }
    }

    Write-StageMessage "$($ProgressLabel)大小相同的候选文件数: $partialHashCandidateCount，候选大小组数: $($sameLengthGroupList.Count)"

    $partialHashCandidateFileList = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    foreach ($sameLengthFileList in $sameLengthGroupList) {
        foreach ($file in $sameLengthFileList) {
            $partialHashCandidateFileList.Add($file)
        }
    }

    $processedPartialHashCount = 0
    $lastPartialHashPercent = -1
    $hashProgressName = "$($ProgressLabel)哈希计算"
    $hashWarningList = New-DeferredScanWarningList
    $fullHashCandidateGroupList = New-Object System.Collections.Generic.List[object]
    $fileListByLengthAndPartialHash = @{}

    foreach ($partialHashRecord in @(Get-ContentHashRecordList `
                -FileList $partialHashCandidateFileList.ToArray() `
                -HashKind 'Partial' `
                -Activity $hashProgressName `
                -Status '正在并行筛选候选文件' `
                -TotalCount $partialHashCandidateCount `
                -ProcessedCount ([ref]$processedPartialHashCount) `
                -LastPercent ([ref]$lastPartialHashPercent) `
                -WarningList $hashWarningList `
                -WarningMessage '跳过文件，无法计算部分哈希')) {
        $lengthKey = [long]$partialHashRecord.File.Length
        if (-not $fileListByLengthAndPartialHash.ContainsKey($lengthKey)) {
            $fileListByLengthAndPartialHash[$lengthKey] = @{}
        }

        Add-FileToGroupBucket `
            -FileListByKey $fileListByLengthAndPartialHash[$lengthKey] `
            -Key $partialHashRecord.Hash `
            -File $partialHashRecord.File
    }

    foreach ($fileListByPartialHash in $fileListByLengthAndPartialHash.Values) {
        foreach ($partialHash in $fileListByPartialHash.Keys) {
            $partialHashFileList = $fileListByPartialHash[$partialHash]
            if ($partialHashFileList.Count -le 1) {
                continue
            }

            # 文件不大于两段采样总长时，部分哈希已按完整文件计算，可直接确认为重复文件组。
            if (Test-PartialHashCoversFullFile -Length $partialHashFileList[0].Length) {
                [pscustomobject]@{
                    Hash  = $partialHash
                    Files = @($partialHashFileList.ToArray())
                }
                continue
            }

            $fullHashCandidateGroupList.Add($partialHashFileList.ToArray())
        }
    }

    $fullHashCandidateCount = 0
    $fullHashCandidateFileList = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    foreach ($fullHashCandidateGroup in $fullHashCandidateGroupList) {
        $fullHashCandidateCount += $fullHashCandidateGroup.Count
        foreach ($file in $fullHashCandidateGroup) {
            $fullHashCandidateFileList.Add($file)
        }
    }

    $processedFullHashCount = [Math]::Max(0, $partialHashCandidateCount - $fullHashCandidateCount)
    $lastFullHashPercent = -1
    $fileListByLengthAndContentHash = @{}
    foreach ($contentHashRecord in @(Get-ContentHashRecordList `
                -FileList $fullHashCandidateFileList.ToArray() `
                -HashKind 'Full' `
                -Activity $hashProgressName `
                -Status '正在并行确认完整哈希' `
                -TotalCount $partialHashCandidateCount `
                -ProcessedCount ([ref]$processedFullHashCount) `
                -LastPercent ([ref]$lastFullHashPercent) `
                -WarningList $hashWarningList `
                -WarningMessage '跳过文件，无法计算完整哈希')) {
        $lengthKey = [long]$contentHashRecord.File.Length
        if (-not $fileListByLengthAndContentHash.ContainsKey($lengthKey)) {
            $fileListByLengthAndContentHash[$lengthKey] = @{}
        }

        Add-FileToGroupBucket `
            -FileListByKey $fileListByLengthAndContentHash[$lengthKey] `
            -Key $contentHashRecord.Hash `
            -File $contentHashRecord.File
    }

    foreach ($fileListByContentHash in $fileListByLengthAndContentHash.Values) {
        foreach ($contentHash in $fileListByContentHash.Keys) {
            $contentHashFileList = $fileListByContentHash[$contentHash]
            if ($contentHashFileList.Count -le 1) {
                continue
            }

            [pscustomobject]@{
                Hash  = $contentHash
                Files = @($contentHashFileList.ToArray())
            }
        }
    }

    if ($partialHashCandidateCount -gt 0) {
        Write-ProgressBar `
            -Activity $hashProgressName `
            -Status '哈希计算完成' `
            -ProcessedCount $partialHashCandidateCount `
            -TotalCount $partialHashCandidateCount `
            -LastPercent ([ref]$lastFullHashPercent) `
            -Force
    }

    Complete-DynamicStatusLine
    Write-DeferredScanWarningList -WarningList $hashWarningList -Title "$($ProgressLabel)哈希跳过汇总"
}

# 输出默认预览中的单个重复文件块。
function Write-DuplicatePreviewBlock {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Hash,

        [Parameter(Mandatory = $true)]
        [string]$KeepPathText,

        [Parameter(Mandatory = $true)]
        [string[]]$DeletePathTexts
    )

    Write-PreviewSeparator
    Write-Host "SHA-256: $Hash" -ForegroundColor DarkGray
    Write-Host -NoNewline "保留文件:" -ForegroundColor Green
    Write-Host " $KeepPathText"
    Write-Host "删除文件:" -ForegroundColor Yellow
    foreach ($path in $DeletePathTexts) {
        Write-Host "  - $path"
    }
}

# 输出默认删除计划摘要，并返回预计删除数量。
function Write-DeletionPlanSummary {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$DeletionPlanList,

        [Parameter(Mandatory = $true)]
        [string]$SummaryMessageTemplate
    )

    $deletionPlanStatistics = Get-DeletionPlanSummary -DeletionPlanList $DeletionPlanList
    $plannedDeletionCount = $deletionPlanStatistics.DeletionItemCount

    Write-Host ""
    Write-Host -NoNewline "重复组数: " -ForegroundColor White
    Write-Host -NoNewline $deletionPlanStatistics.GroupCount -ForegroundColor Cyan
    Write-Host -NoNewline "，默认保留文件数: " -ForegroundColor White
    Write-Host -NoNewline $deletionPlanStatistics.KeepFileCount -ForegroundColor Green
    Write-Host -NoNewline "，默认计划删除文件数: " -ForegroundColor White
    Write-Host -NoNewline $plannedDeletionCount -ForegroundColor Red
    Write-Host -NoNewline "，可释放 " -ForegroundColor White
    Write-Host $deletionPlanStatistics.ReclaimableSizeText -ForegroundColor Magenta

    $summaryMessage = $SummaryMessageTemplate -f $plannedDeletionCount
    if ($plannedDeletionCount -gt 0) {
        $summaryMessage = "$summaryMessage，可释放 $($deletionPlanStatistics.ReclaimableSizeText)"
    }

    Write-StatusSummary -Message $summaryMessage -Color Yellow
    return $plannedDeletionCount
}

# 输出默认删除计划的详细预览与摘要，并返回预计删除数量。
function Write-DeletionPlanPreview {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$DeletionPlanList,

        [Parameter(Mandatory = $true)]
        [string]$SummaryMessageTemplate
    )

    Write-Host "`n删除预览:" -ForegroundColor Yellow
    foreach ($deletionPlan in $DeletionPlanList) {
        $plannedDeletionItems = @($deletionPlan.DeletionItems)
        $plannedDeletePathTexts = @($plannedDeletionItems | ForEach-Object { $_.DisplayPath })
        Write-DuplicatePreviewBlock -Hash $deletionPlan.Hash -KeepPathText $deletionPlan.KeepPathText -DeletePathTexts $plannedDeletePathTexts
    }

    return (Write-DeletionPlanSummary -DeletionPlanList $DeletionPlanList -SummaryMessageTemplate $SummaryMessageTemplate)
}

# 输出手动模式下本组已删除文件的原编号和相对路径。
function Write-ManualDeletionResult {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$DeletedSelectionList,

        [Parameter(Mandatory = $false)]
        [string]$RootPath,

        [Parameter(Mandatory = $false)]
        [hashtable]$DisplayPathByFullName
    )

    Write-Host "已删除文件:" -ForegroundColor Magenta
    foreach ($manualSelection in $DeletedSelectionList) {
        $displayPath = Get-FileDisplayPath -File $manualSelection.File -RootPath $RootPath -DisplayPathByFullName $DisplayPathByFullName
        Write-Host "  - [$($manualSelection.Number)] $displayPath" -ForegroundColor Magenta
    }
}

# 在手动模式下列出重复文件，并读取用户选择的删除动作。
function Read-ManualDeletionSelection {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$DuplicateFileGroup,

        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$DefaultKeepFile,

        [Parameter(Mandatory = $false)]
        [string]$RootPath,

        [Parameter(Mandatory = $false)]
        [hashtable]$DisplayPathByFullName,

        [Parameter(Mandatory = $true)]
        [string]$Hash
    )

    $orderedDuplicateFiles = @(Get-FileListByKeepPriority -FileList $DuplicateFileGroup)
    $defaultDeletionSelections = @(
        for ($index = 0; $index -lt $orderedDuplicateFiles.Count; $index++) {
            if ($orderedDuplicateFiles[$index].FullName -ne $DefaultKeepFile.FullName) {
                [pscustomobject]@{
                    Number = $index + 1
                    File   = $orderedDuplicateFiles[$index]
                }
            }
        }
    )

    Write-PreviewSeparator
    Write-Host "SHA-256: $Hash" -ForegroundColor DarkGray
    Write-Host "重复文件:" -ForegroundColor Yellow

    for ($index = 0; $index -lt $orderedDuplicateFiles.Count; $index++) {
        $fileNumber = $index + 1
        $displayPath = Get-FileDisplayPath -File $orderedDuplicateFiles[$index] -RootPath $RootPath -DisplayPathByFullName $DisplayPathByFullName
        if ($orderedDuplicateFiles[$index].FullName -eq $DefaultKeepFile.FullName) {
            Write-Host "  [$fileNumber] $displayPath  (默认保留)" -ForegroundColor Green
        }
        else {
            Write-Host -NoNewline "  [$fileNumber] " -ForegroundColor Yellow
            Write-Host $displayPath
        }
    }

    while ($true) {
        $manualInputRaw = Read-ColoredLine -Prompt '请输入要删除的编号，多个编号用逗号分隔；直接回车使用默认规则；输入 0 跳过；输入 00 退出脚本: '
        if ($null -eq $manualInputRaw) {
            return [pscustomobject]@{
                Action     = 'Exit'
                Selections = @()
            }
        }

        $trimmedInputText = $manualInputRaw.Trim()

        if ([string]::IsNullOrWhiteSpace($trimmedInputText)) {
            return [pscustomobject]@{
                Action     = 'Delete'
                Selections = $defaultDeletionSelections
            }
        }

        if ($trimmedInputText -eq '0') {
            return [pscustomobject]@{
                Action     = 'Skip'
                Selections = @()
            }
        }

        if ($trimmedInputText -eq '00') {
            return [pscustomobject]@{
                Action     = 'Exit'
                Selections = @()
            }
        }

        $selectedFileNumbers = New-Object System.Collections.Generic.List[int]
        $isValidSelection = $true

        foreach ($inputPart in ($trimmedInputText -split '[,，\s]+')) {
            if ([string]::IsNullOrWhiteSpace($inputPart)) {
                continue
            }

            $fileNumber = 0
            if (-not [int]::TryParse($inputPart, [ref]$fileNumber)) {
                $isValidSelection = $false
                break
            }

            if ($fileNumber -lt 1 -or $fileNumber -gt $orderedDuplicateFiles.Count) {
                $isValidSelection = $false
                break
            }

            if (-not $selectedFileNumbers.Contains($fileNumber)) {
                $selectedFileNumbers.Add($fileNumber)
            }
        }

        if (-not $isValidSelection -or $selectedFileNumbers.Count -eq 0) {
            Write-Host "输入无效，请输入列表中的编号，例如: 2 或 2,3；输入 0 跳过；输入 00 退出脚本" -ForegroundColor Red
            continue
        }

        if ($selectedFileNumbers.Count -ge $orderedDuplicateFiles.Count) {
            Write-Host "不能删除该组中的所有文件，请至少保留一个文件。" -ForegroundColor Red
            continue
        }

        return [pscustomobject]@{
            Action     = 'Delete'
            Selections = @(
                $selectedFileNumbers | ForEach-Object {
                    [pscustomobject]@{
                        Number = $_
                        File   = $orderedDuplicateFiles[$_ - 1]
                    }
                }
            )
        }
    }
}

# 获取文件当前完整哈希并缓存结果，供删除前安全复核复用。
function Get-CachedFullContentHash {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [hashtable]$HashByPath
    )

    if ($HashByPath.ContainsKey($FilePath)) {
        return $HashByPath[$FilePath]
    }

    if (-not [System.IO.File]::Exists($FilePath)) {
        throw "文件已不存在: $FilePath"
    }

    $currentHash = Get-FullContentHash -File ([System.IO.FileInfo]::new($FilePath))
    $HashByPath[$FilePath] = $currentHash
    return $currentHash
}

# 执行一批删除项；删除前重新确认至少一份保留文件和待删文件仍匹配原完整哈希。
function Remove-DeletionItemList {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$DeletionItemList,

        [Parameter(Mandatory = $false)]
        [switch]$Quiet
    )

    $deletedFileCount = 0
    $failedFileCount = 0
    $deletedItemList = New-Object System.Collections.Generic.List[object]
    $currentHashByPath = @{}
    foreach ($deletionItem in $DeletionItemList) {
        try {
            $hasValidKeepFile = $false
            foreach ($keepFile in @($deletionItem.KeepFileList)) {
                try {
                    $keepHash = Get-CachedFullContentHash -FilePath $keepFile.FullName -HashByPath $currentHashByPath
                    if ($keepHash.Equals($deletionItem.ExpectedHash, [System.StringComparison]::OrdinalIgnoreCase)) {
                        $hasValidKeepFile = $true
                        break
                    }
                }
                catch {
                    Write-Debug "保留文件复核失败: $($keepFile.FullName)。原因: $($_.Exception.Message)"
                }
            }

            if (-not $hasValidKeepFile) {
                throw '未找到仍与扫描结果一致的保留文件，已拒绝删除。'
            }

            $deletionFilePath = $deletionItem.File.FullName
            $currentDeletionHash = Get-CachedFullContentHash -FilePath $deletionFilePath -HashByPath $currentHashByPath
            if (-not $currentDeletionHash.Equals($deletionItem.ExpectedHash, [System.StringComparison]::OrdinalIgnoreCase)) {
                throw '待删文件内容已变化，已拒绝删除。'
            }

            Remove-FileSystemFile -FilePath $deletionFilePath
            if (-not $Quiet) {
                Write-Host "已删除: $($deletionItem.DisplayPath)" -ForegroundColor Magenta
            }
            $deletedFileCount++
            $deletedItemList.Add($deletionItem)
        }
        catch {
            Write-Host "删除失败: $($deletionItem.DisplayPath)" -ForegroundColor Red
            Write-Host "  原因: $($_.Exception.Message)" -ForegroundColor DarkGray
            $failedFileCount++
        }
    }

    return [pscustomobject]@{
        DeletedCount = $deletedFileCount
        FailedCount  = $failedFileCount
        DeletedItems = $deletedItemList.ToArray()
    }
}

# 执行默认删除计划并输出统一的删除汇总。
function Invoke-DefaultDeletionPlan {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$DeletionPlanList,

        [Parameter(Mandatory = $true)]
        [string]$SummaryMessageTemplate
    )

    $deletionResult = Remove-DeletionItemList -DeletionItemList @($DeletionPlanList | ForEach-Object { $_.DeletionItems })
    $deletedByteCount = Get-DeletionItemTotalByteCount -DeletionItemList @($deletionResult.DeletedItems)
    $summaryMessage = $SummaryMessageTemplate -f $deletionResult.DeletedCount
    if ($deletionResult.DeletedCount -gt 0) {
        $summaryMessage = "$summaryMessage，已释放 $(Format-ByteSize -ByteCount $deletedByteCount)"
    }

    Write-StatusSummary -Message $summaryMessage -Color Magenta
    if ($deletionResult.FailedCount -gt 0) {
        Write-Host "删除失败文件: $($deletionResult.FailedCount)" -ForegroundColor Red
    }

    return $deletionResult
}

# 生成删除操作菜单选项；0 只表示跳过或返回，00 统一表示退出脚本。
function New-DeletionActionMenuOptionList {
    param(
        [Parameter(Mandatory = $false)]
        [switch]$IncludeManualDeletion,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeSkipCurrentDirectory
    )

    $menuOptionList = New-Object System.Collections.Generic.List[object]
    $menuOptionList.Add([pscustomobject]@{ Value = '1'; Label = '默认删除' })
    if ($IncludeManualDeletion) {
        $menuOptionList.Add([pscustomobject]@{ Value = '2'; Label = '手动删除' })
    }
    if ($IncludeSkipCurrentDirectory) {
        $menuOptionList.Add([pscustomobject]@{ Value = '0'; Label = '跳过当前目录' })
    }
    else {
        $menuOptionList.Add([pscustomobject]@{ Value = '0'; Label = '跳过本次操作' })
    }
    $menuOptionList.Add([pscustomobject]@{ Value = '00'; Label = '退出脚本' })

    return $menuOptionList.ToArray()
}

# 输出通用菜单并读取用户选择。
function Read-MenuChoice {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [object[]]$MenuOptionList
    )

    Complete-DynamicStatusLine
    Write-Host ""
    Write-Host $Title -ForegroundColor Cyan
    foreach ($menuOption in $MenuOptionList) {
        Write-MenuItem -Number $menuOption.Value -Text $menuOption.Label
    }

    $validMenuChoices = @($MenuOptionList | ForEach-Object { $_.Value })
    while ($true) {
        $inputRaw = Read-ColoredLine -Prompt '请输入选项: '

        if ($null -eq $inputRaw) {
            Write-Host "输入流已结束，程序退出。" -ForegroundColor Yellow
            return "00"
        }

        $menuChoice = $inputRaw.Trim()

        if ($validMenuChoices -contains $menuChoice) {
            return $menuChoice
        }

        Write-Host "输入无效，请输入: $($validMenuChoices -join ', ')" -ForegroundColor Red
    }
}

# 输出操作菜单并读取用户选择。
function Read-DeletionAction {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$MenuOptionList
    )

    return Read-MenuChoice -Title '请选择操作:' -MenuOptionList $MenuOptionList
}

# 无路径且未指定模式时，先让用户选择扫描模式，避免默认落入单目录模式造成误解。
function Read-InteractiveScanMode {
    $modeChoice = Read-MenuChoice -Title '请选择扫描模式:' -MenuOptionList @(
        [pscustomobject]@{ Value = '1'; Label = '单目录模式（每个目录分别扫描）' }
        [pscustomobject]@{ Value = '2'; Label = '多目录合并模式（多个目录视作一个大目录）' }
        [pscustomobject]@{ Value = '3'; Label = '参考目录模式（首个目录作为参考目录）' }
        [pscustomobject]@{ Value = '00'; Label = '退出脚本' }
    )

    switch ($modeChoice) {
        '1' { return 'Single' }
        '2' { return 'Aggregate' }
        '3' { return 'Reference' }
        '00' { return 'Exit' }
    }
}

# 根据当前模式返回交互输入所需的最少目录数量。
function Get-ScanModeMinimumPathCount {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$UseReferenceMode
    )

    if ($UseReferenceMode) {
        return 2
    }

    return 1
}

# 根据当前模式返回路径输入阶段的模式提示。
function Get-ScanModePrompt {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$UseReferenceMode,

        [Parameter(Mandatory = $true)]
        [bool]$UseAggregateMode
    )

    if ($UseReferenceMode) {
        return '参考目录模式：首个目录作为参考目录，其余目录作为目标目录'
    }

    if ($UseAggregateMode) {
        return '多目录合并模式：多个目录将视作一个大目录扫描'
    }

    return '单目录模式：每个目录将分别扫描'
}

# 为单目录模式生成默认删除计划。
function New-SingleDirectoryDeletionPlan {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    $scannedFiles = @(Get-ScannedFileList -RootPath $RootPath -ProgressLabel '单目录')
    if ($scannedFiles.Count -lt 2) {
        return @()
    }

    foreach ($duplicateGroupRecord in Find-DuplicateFileGroup -FileList $scannedFiles -ProgressLabel '单目录') {
        $duplicateFiles = @($duplicateGroupRecord.Files)
        $defaultKeepFile = Select-DefaultKeepFile -FileList $duplicateFiles
        $filesToDelete = New-Object System.Collections.Generic.List[System.IO.FileInfo]
        foreach ($duplicateFile in $duplicateFiles) {
            if (-not $duplicateFile.FullName.Equals($defaultKeepFile.FullName, [System.StringComparison]::OrdinalIgnoreCase)) {
                $filesToDelete.Add($duplicateFile)
            }
        }

        [pscustomobject]@{
            Hash           = $duplicateGroupRecord.Hash
            KeepFile      = $defaultKeepFile
            KeepPathText  = Get-RelativePathText -File $defaultKeepFile -RootPath $RootPath
            DeletionItems = New-DeletionItemList `
                -FileList ($filesToDelete.ToArray()) `
                -RootPath $RootPath `
                -ExpectedHash $duplicateGroupRecord.Hash `
                -KeepFileList @($defaultKeepFile)
            DuplicateFiles = $duplicateFiles
        }
    }
}

# 为多目录合并模式生成默认删除计划；所有输入目录会被视作一个虚拟大目录。
function New-MergedDirectoryDeletionPlan {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$RootPathList
    )

    $fileRecordList = New-Object System.Collections.Generic.List[object]
    $mergedScanFileCount = 0
    $lastMergedScanPercent = -1
    $hiddenScopeText = if ($ShouldIncludeHiddenItems) { '包含隐藏项' } else { '不包含隐藏项' }
    $mergedScanWarningList = New-DeferredScanWarningList

    Write-StageMessage "开始扫描合并目录，目录数: $($RootPathList.Count)"
    for ($rootIndex = 0; $rootIndex -lt $RootPathList.Count; $rootIndex++) {
        $rootPath = $RootPathList[$rootIndex]
        $rootNumber = $rootIndex + 1
        $rootLabel = Get-DirectoryLabel -RootPath $rootPath
        $pathPrefix = "{0}-{1}" -f $rootNumber, $rootLabel

        Write-ProgressBar -Activity '合并目录扫描' -Status "正在扫描目录 $rootNumber：$rootLabel" -ProcessedCount $rootIndex -TotalCount $RootPathList.Count -LastPercent ([ref]$lastMergedScanPercent)
        $scannedFiles = @(Get-ScannedFileList -RootPath $rootPath -ProgressLabel "合并目录$rootNumber" -SuppressScanStageMessages -WarningList $mergedScanWarningList)
        $mergedScanFileCount += $scannedFiles.Count

        foreach ($file in $scannedFiles) {
            $fileRecordList.Add([pscustomobject]@{
                File        = $file
                DisplayPath = Get-RelativePathText -File $file -RootPath $rootPath -PathPrefix $pathPrefix
            })
        }
    }
    Write-ProgressBar -Activity '合并目录扫描' -Status '扫描完成' -ProcessedCount $RootPathList.Count -TotalCount $RootPathList.Count -LastPercent ([ref]$lastMergedScanPercent) -Force
    Complete-DynamicStatusLine
    Write-DeferredScanWarningList -WarningList $mergedScanWarningList -Title '合并目录扫描跳过汇总'
    Write-StageMessage "合并目录扫描完成，目录数: $($RootPathList.Count)，文件数: $mergedScanFileCount，$hiddenScopeText"

    $scannedFileList = @($fileRecordList | ForEach-Object { $_.File })
    if ($scannedFileList.Count -lt 2) {
        return @()
    }

    $displayPathByFullName = @{}
    foreach ($fileRecord in $fileRecordList) {
        $displayPathByFullName[$fileRecord.File.FullName] = $fileRecord.DisplayPath
    }

    foreach ($duplicateGroupRecord in Find-DuplicateFileGroup -FileList $scannedFileList -ProgressLabel '多目录合并') {
        $duplicateFiles = @($duplicateGroupRecord.Files)
        $defaultKeepFile = Select-DefaultKeepFile -FileList $duplicateFiles
        $filesToDelete = New-Object System.Collections.Generic.List[System.IO.FileInfo]
        foreach ($duplicateFile in $duplicateFiles) {
            if (-not $duplicateFile.FullName.Equals($defaultKeepFile.FullName, [System.StringComparison]::OrdinalIgnoreCase)) {
                $filesToDelete.Add($duplicateFile)
            }
        }

        [pscustomobject]@{
            Hash           = $duplicateGroupRecord.Hash
            KeepFile       = $defaultKeepFile
            KeepPathText   = $displayPathByFullName[$defaultKeepFile.FullName]
            DeletionItems  = New-DeletionItemList `
                -FileList ($filesToDelete.ToArray()) `
                -DisplayPathByFullName $displayPathByFullName `
                -ExpectedHash $duplicateGroupRecord.Hash `
                -KeepFileList @($defaultKeepFile)
            DuplicateFiles = $duplicateFiles
        }
    }
}

# 从删除计划中提取显示路径映射，供手动删除模式复用。
function New-DisplayPathMapFromDeletionPlan {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$DeletionPlanList
    )

    $displayPathByFullName = @{}
    foreach ($deletionPlan in $DeletionPlanList) {
        $displayPathByFullName[$deletionPlan.KeepFile.FullName] = $deletionPlan.KeepPathText
        foreach ($deletionItem in @($deletionPlan.DeletionItems)) {
            $displayPathByFullName[$deletionItem.File.FullName] = $deletionItem.DisplayPath
        }
    }

    return $displayPathByFullName
}

# 执行单目录手动删除流程：每组选择后立即删除并输出结果。
function Invoke-ManualDeletion {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$DeletionPlanList,

        [Parameter(Mandatory = $false)]
        [string]$RootPath,

        [Parameter(Mandatory = $false)]
        [hashtable]$DisplayPathByFullName
    )

    $deletedFileCount = 0
    $failedFileCount = 0
    $deletedItemList = New-Object System.Collections.Generic.List[object]
    foreach ($deletionPlan in $DeletionPlanList) {
        $manualSelection = Read-ManualDeletionSelection -DuplicateFileGroup $deletionPlan.DuplicateFiles -DefaultKeepFile $deletionPlan.KeepFile -RootPath $RootPath -DisplayPathByFullName $DisplayPathByFullName -Hash $deletionPlan.Hash
        if ($manualSelection.Action -eq 'Exit') {
            Write-Host "已退出脚本。" -ForegroundColor Yellow
            if ($deletedFileCount -gt 0) {
                Write-Host "本次已手动删除重复文件: $deletedFileCount" -ForegroundColor Magenta
            }
            if ($failedFileCount -gt 0) {
                Write-Host "本次删除失败文件: $failedFileCount" -ForegroundColor Red
            }
            return 'Exit'
        }

        if ($manualSelection.Action -eq 'Skip') {
            continue
        }

        $selectedDeletionEntries = @($manualSelection.Selections)
        $selectedFilesToDelete = @($selectedDeletionEntries | ForEach-Object { $_.File })
        if ($selectedFilesToDelete.Count -eq 0) {
            continue
        }

        $selectedFilePathSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($selectedFile in $selectedFilesToDelete) {
            [void]$selectedFilePathSet.Add($selectedFile.FullName)
        }

        $keepFileList = @(
            $deletionPlan.DuplicateFiles | Where-Object {
                -not $selectedFilePathSet.Contains($_.FullName)
            }
        )
        $manualDeletionItems = New-DeletionItemList `
            -FileList $selectedFilesToDelete `
            -RootPath $RootPath `
            -DisplayPathByFullName $DisplayPathByFullName `
            -ExpectedHash $deletionPlan.Hash `
            -KeepFileList $keepFileList

        $deletionResult = Remove-DeletionItemList -DeletionItemList $manualDeletionItems -Quiet
        $deletedFilePathSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($deletedItem in @($deletionResult.DeletedItems)) {
            [void]$deletedFilePathSet.Add($deletedItem.File.FullName)
        }

        $deletedSelectionList = @(
            foreach ($selectedDeletionEntry in $selectedDeletionEntries) {
                if ($deletedFilePathSet.Contains($selectedDeletionEntry.File.FullName)) {
                    $selectedDeletionEntry
                }
            }
        )
        if ($deletedSelectionList.Count -gt 0) {
            Write-ManualDeletionResult -DeletedSelectionList $deletedSelectionList -RootPath $RootPath -DisplayPathByFullName $DisplayPathByFullName
        }
        $deletedFileCount += $deletionResult.DeletedCount
        $failedFileCount += $deletionResult.FailedCount
        foreach ($deletedItem in @($deletionResult.DeletedItems)) {
            $deletedItemList.Add($deletedItem)
        }
    }

    if ($deletedFileCount -eq 0) {
        Write-Host "未选择删除任何文件。" -ForegroundColor Yellow
    }
    else {
        $deletedByteCount = Get-DeletionItemTotalByteCount -DeletionItemList $deletedItemList.ToArray()
        Write-StatusSummary -Message "手动删除完成。已删除重复文件: $deletedFileCount，已释放 $(Format-ByteSize -ByteCount $deletedByteCount)" -Color Magenta
    }

    if ($failedFileCount -gt 0) {
        Write-Host "删除失败文件: $failedFileCount" -ForegroundColor Red
    }

    return 'Continue'
}

# 统计默认删除计划中的待删除文件数，用于多目录扫描完成后的汇总输出。
function Get-DeletionPlanItemCount {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$DeletionPlanList
    )

    $deletionItemCount = 0
    foreach ($deletionPlan in $DeletionPlanList) {
        $deletionItemCount += @($deletionPlan.DeletionItems).Count
    }

    return $deletionItemCount
}

# 汇总删除项文件大小；文件已不可读时跳过，避免摘要统计影响主流程。
function Get-DeletionItemTotalByteCount {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$DeletionItemList
    )

    [long]$totalByteCount = 0
    foreach ($deletionItem in $DeletionItemList) {
        try {
            if ($null -ne $deletionItem.PSObject.Properties['ByteCount']) {
                $totalByteCount += [long]$deletionItem.ByteCount
                continue
            }

            if ($null -ne $deletionItem.File) {
                $totalByteCount += [long]$deletionItem.File.Length
            }
        }
        catch {
            Write-Debug "统计文件大小失败: $($_.Exception.Message)"
        }
    }

    return $totalByteCount
}

# 将字节数格式化为便于判断释放价值的大小文本。
function Format-ByteSize {
    param(
        [Parameter(Mandatory = $true)]
        [long]$ByteCount
    )

    if ($ByteCount -lt 1KB) {
        return "$ByteCount B"
    }

    $unitList = @(
        [pscustomobject]@{ Name = 'TB'; Size = 1TB }
        [pscustomobject]@{ Name = 'GB'; Size = 1GB }
        [pscustomobject]@{ Name = 'MB'; Size = 1MB }
        [pscustomobject]@{ Name = 'KB'; Size = 1KB }
    )

    foreach ($unit in $unitList) {
        if ($ByteCount -ge $unit.Size) {
            return ('{0:N2} {1}' -f ($ByteCount / $unit.Size), $unit.Name)
        }
    }
}

# 统一统计删除计划数量和预计可释放空间，供预览、-yes 和最终删除流程复用。
function Get-DeletionPlanSummary {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$DeletionPlanList
    )

    $plannedDeletionCount = 0
    $plannedKeepCount = 0
    $deletionItemList = [System.Collections.Generic.List[object]]::new()

    foreach ($deletionPlan in $DeletionPlanList) {
        $plannedDeletionItems = @($deletionPlan.DeletionItems)
        $plannedDeletionCount += $plannedDeletionItems.Count
        $plannedKeepCount++
        foreach ($plannedDeletionItem in $plannedDeletionItems) {
            $deletionItemList.Add($plannedDeletionItem)
        }
    }

    $reclaimableByteCount = Get-DeletionItemTotalByteCount -DeletionItemList $deletionItemList.ToArray()
    return [pscustomobject]@{
        GroupCount           = $DeletionPlanList.Count
        KeepFileCount        = $plannedKeepCount
        DeletionItemCount    = $plannedDeletionCount
        ReclaimableByteCount = $reclaimableByteCount
        ReclaimableSizeText  = Format-ByteSize -ByteCount $reclaimableByteCount
    }
}

# 对已经生成的删除计划执行预览、默认删除、手动删除、跳过或退出。
function Invoke-DeletionPlanAction {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$DeletionPlanList,

        [Parameter(Mandatory = $true)]
        [string]$EmptyMessage,

        [Parameter(Mandatory = $true)]
        [string]$PreviewSummaryMessageTemplate,

        [Parameter(Mandatory = $true)]
        [string]$DefaultDeletionSummaryMessageTemplate,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeManualDeletion,

        [Parameter(Mandatory = $false)]
        [switch]$AllowSkipCurrentDirectory,

        [Parameter(Mandatory = $false)]
        [string]$ManualRootPath,

        [Parameter(Mandatory = $false)]
        [hashtable]$ManualDisplayPathByFullName
    )

    $deletionPlanStatistics = Get-DeletionPlanSummary -DeletionPlanList $DeletionPlanList
    if ($DeletionPlanList.Count -eq 0 -or $deletionPlanStatistics.DeletionItemCount -eq 0) {
        Write-Host $EmptyMessage -ForegroundColor Green
        return 'Continue'
    }

    if ($ShouldAssumeYesDeletion) {
        [void](Write-DeletionPlanSummary -DeletionPlanList $DeletionPlanList -SummaryMessageTemplate $PreviewSummaryMessageTemplate)

        if (-not (Wait-AssumeYesDeletionGracePeriod)) {
            return 'Exit'
        }

        [void](Invoke-DefaultDeletionPlan -DeletionPlanList $DeletionPlanList -SummaryMessageTemplate $DefaultDeletionSummaryMessageTemplate)
        return 'Continue'
    }

    [void](Write-DeletionPlanPreview -DeletionPlanList $DeletionPlanList -SummaryMessageTemplate $PreviewSummaryMessageTemplate)
    $menuOptionList = New-DeletionActionMenuOptionList -IncludeManualDeletion:$IncludeManualDeletion -IncludeSkipCurrentDirectory:$AllowSkipCurrentDirectory
    $menuChoice = Read-DeletionAction -MenuOptionList $menuOptionList

    if ($menuChoice -eq '0') {
        if ($AllowSkipCurrentDirectory) {
            Write-Host "已跳过当前目录，未删除任何文件。" -ForegroundColor Yellow
            return 'Skip'
        }

        Write-Host "已跳过本次操作，未删除任何文件。" -ForegroundColor Yellow
        return 'Continue'
    }

    if ($menuChoice -eq '00') {
        Write-Host "已退出脚本，未删除任何文件。" -ForegroundColor Yellow
        return 'Exit'
    }

    if ($menuChoice -eq '1') {
        [void](Invoke-DefaultDeletionPlan -DeletionPlanList $DeletionPlanList -SummaryMessageTemplate $DefaultDeletionSummaryMessageTemplate)
        return 'Continue'
    }

    if ($IncludeManualDeletion) {
        return (Invoke-ManualDeletion -DeletionPlanList $DeletionPlanList -RootPath $ManualRootPath -DisplayPathByFullName $ManualDisplayPathByFullName)
    }

    return 'Continue'
}

# 对单目录删除计划执行操作；多个单目录逐个操作时可允许跳过当前目录。
function Invoke-SingleDirectoryDeletionAction {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$DeletionPlanList,

        [Parameter(Mandatory = $false)]
        [switch]$AllowSkipCurrentDirectory
    )

    $actionParameters = @{
        DeletionPlanList                      = $DeletionPlanList
        EmptyMessage                          = '未发现重复文件。'
        PreviewSummaryMessageTemplate         = '重复文件列举完成。默认计划删除重复文件: {0}'
        DefaultDeletionSummaryMessageTemplate = '删除完成。已删除重复文件: {0}'
        IncludeManualDeletion                 = $true
        AllowSkipCurrentDirectory             = $AllowSkipCurrentDirectory
        ManualRootPath                        = $RootPath
    }

    return (Invoke-DeletionPlanAction @actionParameters)
}

# 单目录去重入口：扫描一个目录后预览默认删除计划，再按用户选择执行删除。
function Invoke-SingleDirectoryMode {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    $deletionPlanList = @(New-SingleDirectoryDeletionPlan -RootPath $RootPath)
    return (Invoke-SingleDirectoryDeletionAction -RootPath $RootPath -DeletionPlanList $deletionPlanList)
}

# 多个单目录先全部扫描，再按目录逐个预览和确认，避免扫描过程中被菜单反复打断。
function Invoke-SeparateSingleDirectoryMode {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$RootPathList
    )

    if ($RootPathList.Count -eq 1) {
        return (Invoke-SingleDirectoryMode -RootPath $RootPathList[0])
    }

    $singleDirectoryPlanRecordList = New-Object System.Collections.Generic.List[object]
    for ($rootIndex = 0; $rootIndex -lt $RootPathList.Count; $rootIndex++) {
        Write-PreviewSeparator
        Write-Host "单目录模式扫描 $($rootIndex + 1) / $($RootPathList.Count): $($RootPathList[$rootIndex])" -ForegroundColor Cyan

        $deletionPlanList = @(New-SingleDirectoryDeletionPlan -RootPath $RootPathList[$rootIndex])
        $plannedDeletionCount = Get-DeletionPlanItemCount -DeletionPlanList $deletionPlanList
        if ($deletionPlanList.Count -eq 0) {
            Write-Host "扫描结果: 未发现重复文件。" -ForegroundColor Green
        }
        else {
            Write-Host "扫描结果: 重复组数 $($deletionPlanList.Count)，默认计划删除文件数 $plannedDeletionCount。" -ForegroundColor DarkGray
        }

        $singleDirectoryPlanRecordList.Add([pscustomobject]@{
            RootPath           = $RootPathList[$rootIndex]
            DeletionPlanList   = $deletionPlanList
        })
    }

    $actionPlanRecordList = @($singleDirectoryPlanRecordList | Where-Object { $_.DeletionPlanList.Count -gt 0 })
    $actionDirectoryCount = $actionPlanRecordList.Count
    Write-StatusSummary -Message "单目录模式扫描完成。待操作目录: $actionDirectoryCount / $($RootPathList.Count)" -Color Cyan
    if ($actionDirectoryCount -eq 0) {
        return 'Continue'
    }

    for ($recordIndex = 0; $recordIndex -lt $actionPlanRecordList.Count; $recordIndex++) {
        $planRecord = $actionPlanRecordList[$recordIndex]
        Write-PreviewSeparator
        Write-Host "单目录模式操作 $($recordIndex + 1) / $($actionPlanRecordList.Count): $($planRecord.RootPath)" -ForegroundColor Cyan

        $operationResult = Invoke-SingleDirectoryDeletionAction -RootPath $planRecord.RootPath -DeletionPlanList @($planRecord.DeletionPlanList) -AllowSkipCurrentDirectory
        if ($operationResult -eq 'Exit' -or $AssumeYesDeletionCancelled) {
            return 'Exit'
        }
    }

    return 'Continue'
}

# 多目录合并模式入口：先把所有目录视作一个虚拟大目录，再按用户选择默认或手动删除。
function Invoke-MergedDirectoryMode {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$RootPathList
    )

    $deletionPlanList = @(New-MergedDirectoryDeletionPlan -RootPathList $RootPathList)
    $displayPathByFullName = New-DisplayPathMapFromDeletionPlan -DeletionPlanList $deletionPlanList
    $actionParameters = @{
        DeletionPlanList                      = $deletionPlanList
        EmptyMessage                          = '未发现重复文件。'
        PreviewSummaryMessageTemplate         = '重复文件列举完成。默认计划从多目录合并结果中删除重复文件: {0}'
        DefaultDeletionSummaryMessageTemplate = '删除完成。已从多目录合并结果中删除重复文件: {0}'
        IncludeManualDeletion                 = $true
        ManualDisplayPathByFullName           = $displayPathByFullName
    }
    return (Invoke-DeletionPlanAction @actionParameters)
}

# 为参考目录模式建立轻量参考索引；这里只按文件大小分组，不读取文件内容。
function New-ReferenceDirectoryIndex {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ReferenceRootPath
    )

    $referenceFileList = @(Get-ScannedFileList -RootPath $ReferenceRootPath -ProgressLabel '参考目录')
    $referenceFilesByLength = @{}
    $indexedReferenceFileCount = 0

    if ($referenceFileList.Count -gt 0) {
        Write-StageMessage "参考目录模式构建文件大小索引..."
        $processedReferenceFileCount = 0
        $lastReferenceLengthPercent = -1
        foreach ($file in $referenceFileList) {
            $processedReferenceFileCount++
            Write-ProgressBar -Activity '参考目录大小索引' -Status '正在按文件大小归类' -ProcessedCount $processedReferenceFileCount -TotalCount $referenceFileList.Count -LastPercent ([ref]$lastReferenceLengthPercent)

            if (-not $referenceFilesByLength.ContainsKey($file.Length)) {
                $referenceFilesByLength[$file.Length] = New-Object System.Collections.Generic.List[System.IO.FileInfo]
            }
            $referenceFilesByLength[$file.Length].Add($file)
            $indexedReferenceFileCount++
        }
        Complete-DynamicStatusLine
        Write-StageMessage "参考目录大小索引完成，已索引文件数: $indexedReferenceFileCount"
    }

    return [pscustomobject]@{
        RootPath            = $ReferenceRootPath
        PathPrefix          = Get-DirectoryLabel -RootPath $ReferenceRootPath
        FileList            = $referenceFileList
        FilesByLength       = $referenceFilesByLength
        PartialHashIndexByLength = @{}
        FullHashIndexCache       = @{}
        PartialHashByFullName    = @{}
        FullHashByFullName       = @{}
        IndexedFileCount         = $indexedReferenceFileCount
    }
}

# 按目标实际命中的文件大小，批量懒加载参考目录部分哈希索引。
function Initialize-ReferencePartialHashIndexForLengthList {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ReferenceIndex,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [long[]]$LengthList,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$WarningList
    )

    $pendingLengthList = New-Object System.Collections.Generic.List[long]
    $referenceFileToHashList = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    foreach ($length in $LengthList) {
        if ($ReferenceIndex.PartialHashIndexByLength.ContainsKey($length)) {
            continue
        }

        if (-not $ReferenceIndex.FilesByLength.ContainsKey($length)) {
            $ReferenceIndex.PartialHashIndexByLength[$length] = @{}
            continue
        }

        $pendingLengthList.Add($length)
        foreach ($referenceFile in @($ReferenceIndex.FilesByLength[$length])) {
            if (-not $ReferenceIndex.PartialHashByFullName.ContainsKey($referenceFile.FullName)) {
                $referenceFileToHashList.Add($referenceFile)
            }
        }
    }

    if ($referenceFileToHashList.Count -gt 0) {
        $processedReferencePartialHashCount = 0
        $lastReferencePartialHashPercent = -1
        foreach ($partialHashRecord in @(Get-ContentHashRecordList `
                    -FileList $referenceFileToHashList.ToArray() `
                    -HashKind 'Partial' `
                    -Activity '参考目录懒加载哈希' `
                    -Status '正在并行计算参考部分哈希' `
                    -TotalCount $referenceFileToHashList.Count `
                    -ProcessedCount ([ref]$processedReferencePartialHashCount) `
                    -LastPercent ([ref]$lastReferencePartialHashPercent) `
                    -WarningList $WarningList `
                    -WarningMessage '跳过参考文件，无法计算部分哈希')) {
            $ReferenceIndex.PartialHashByFullName[$partialHashRecord.File.FullName] = $partialHashRecord.Hash
        }
    }

    foreach ($length in $pendingLengthList) {
        $partialHashIndex = @{}
        foreach ($referenceFile in @($ReferenceIndex.FilesByLength[$length])) {
            if (-not $ReferenceIndex.PartialHashByFullName.ContainsKey($referenceFile.FullName)) {
                continue
            }

            Add-FileToGroupBucket `
                -FileListByKey $partialHashIndex `
                -Key $ReferenceIndex.PartialHashByFullName[$referenceFile.FullName] `
                -File $referenceFile
        }

        $ReferenceIndex.PartialHashIndexByLength[$length] = $partialHashIndex
    }
}

# 返回已缓存的参考目录部分哈希索引；调用方应先执行懒加载初始化。
function Get-ReferencePartialHashIndexFromCache {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ReferenceIndex,

        [Parameter(Mandatory = $true)]
        [long]$Length
    )

    if (-not $ReferenceIndex.PartialHashIndexByLength.ContainsKey($Length)) {
        return @{}
    }

    return $ReferenceIndex.PartialHashIndexByLength[$Length]
}

# 按目标实际命中的“文件大小 + 部分哈希”，批量懒加载参考目录完整哈希索引。
function Initialize-ReferenceFullHashIndexForMatchList {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ReferenceIndex,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$MatchList,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$WarningList
    )

    $pendingMatchList = New-Object System.Collections.Generic.List[object]
    $pendingMatchKeySet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $referenceFileToHashList = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    foreach ($match in $MatchList) {
        if (Test-PartialHashCoversFullFile -Length $match.Length) {
            continue
        }

        if (-not $ReferenceIndex.FullHashIndexCache.ContainsKey($match.Length)) {
            $ReferenceIndex.FullHashIndexCache[$match.Length] = @{}
        }

        $fullHashIndexByPartialHash = $ReferenceIndex.FullHashIndexCache[$match.Length]
        if ($fullHashIndexByPartialHash.ContainsKey($match.PartialHash)) {
            continue
        }

        $matchKey = "$($match.Length)`0$($match.PartialHash)"
        if (-not $pendingMatchKeySet.Add($matchKey)) {
            continue
        }

        $partialHashIndex = Get-ReferencePartialHashIndexFromCache -ReferenceIndex $ReferenceIndex -Length $match.Length
        if (-not $partialHashIndex.ContainsKey($match.PartialHash)) {
            $fullHashIndexByPartialHash[$match.PartialHash] = @{}
            continue
        }

        $pendingMatchList.Add($match)
        foreach ($referenceFile in @($partialHashIndex[$match.PartialHash])) {
            if (-not $ReferenceIndex.FullHashByFullName.ContainsKey($referenceFile.FullName)) {
                $referenceFileToHashList.Add($referenceFile)
            }
        }
    }

    if ($referenceFileToHashList.Count -gt 0) {
        $processedReferenceFullHashCount = 0
        $lastReferenceFullHashPercent = -1
        foreach ($fullHashRecord in @(Get-ContentHashRecordList `
                    -FileList $referenceFileToHashList.ToArray() `
                    -HashKind 'Full' `
                    -Activity '参考目录懒加载哈希' `
                    -Status '正在并行计算参考完整哈希' `
                    -TotalCount $referenceFileToHashList.Count `
                    -ProcessedCount ([ref]$processedReferenceFullHashCount) `
                    -LastPercent ([ref]$lastReferenceFullHashPercent) `
                    -WarningList $WarningList `
                    -WarningMessage '跳过参考文件，无法计算完整哈希')) {
            $ReferenceIndex.FullHashByFullName[$fullHashRecord.File.FullName] = $fullHashRecord.Hash
        }
    }

    foreach ($match in $pendingMatchList) {
        $partialHashIndex = Get-ReferencePartialHashIndexFromCache -ReferenceIndex $ReferenceIndex -Length $match.Length
        $fullHashIndex = @{}
        foreach ($referenceFile in @($partialHashIndex[$match.PartialHash])) {
            if (-not $ReferenceIndex.FullHashByFullName.ContainsKey($referenceFile.FullName)) {
                continue
            }

            Add-FileToGroupBucket `
                -FileListByKey $fullHashIndex `
                -Key $ReferenceIndex.FullHashByFullName[$referenceFile.FullName] `
                -File $referenceFile
        }

        $ReferenceIndex.FullHashIndexCache[$match.Length][$match.PartialHash] = $fullHashIndex
    }
}

# 返回已缓存的参考目录完整哈希索引；调用方应先执行懒加载初始化。
function Get-ReferenceFullHashIndexFromCache {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ReferenceIndex,

        [Parameter(Mandatory = $true)]
        [long]$Length,

        [Parameter(Mandatory = $true)]
        [string]$PartialHash
    )

    if (-not $ReferenceIndex.FullHashIndexCache.ContainsKey($Length)) {
        return @{}
    }

    $fullHashIndexByPartialHash = $ReferenceIndex.FullHashIndexCache[$Length]
    if (-not $fullHashIndexByPartialHash.ContainsKey($PartialHash)) {
        return @{}
    }

    return $fullHashIndexByPartialHash[$PartialHash]
}

# 为参考目录模式生成删除计划：参考目录只参与比较，目标目录才会进入删除列表。
function New-ReferenceDirectoryDeletionPlan {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ReferenceIndex,

        [Parameter(Mandatory = $true)]
        [string]$TargetRootPath
    )

    $referenceFileList = @($ReferenceIndex.FileList)
    $targetFileList = @(Get-ScannedFileList -RootPath $TargetRootPath -ProgressLabel '目标目录')
    if ($referenceFileList.Count -eq 0 -or $ReferenceIndex.IndexedFileCount -eq 0 -or $targetFileList.Count -eq 0) {
        return @()
    }

    $referenceRootPath = $ReferenceIndex.RootPath
    $referencePathPrefix = $ReferenceIndex.PathPrefix
    $referenceFilesByLength = $ReferenceIndex.FilesByLength
    $targetPathPrefix = Get-DirectoryLabel -RootPath $TargetRootPath

    Write-StageMessage "参考目录模式使用批量懒加载索引筛选目标目录..."
    $matchedTargetFilesByHash = @{}
    $matchWarningList = New-DeferredScanWarningList
    $targetPartialCandidateFileList = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    $targetCandidateLengthSet = @{}

    $processedTargetFileCount = 0
    $lastTargetMatchPercent = -1
    foreach ($file in $targetFileList) {
        $processedTargetFileCount++
        Write-ProgressBar -Activity '目标目录重复文件筛选' -Status '正在筛选目标大小候选' -ProcessedCount $processedTargetFileCount -TotalCount $targetFileList.Count -LastPercent ([ref]$lastTargetMatchPercent)

        # 目标文件只有在参考目录存在相同大小文件时，才需要进入后续哈希比较。
        if (-not $referenceFilesByLength.ContainsKey($file.Length)) {
            continue
        }

        $targetPartialCandidateFileList.Add($file)
        $targetCandidateLengthSet[$file.Length] = $true
    }

    Initialize-ReferencePartialHashIndexForLengthList `
        -ReferenceIndex $ReferenceIndex `
        -LengthList ([long[]]@($targetCandidateLengthSet.Keys)) `
        -WarningList $matchWarningList

    $processedTargetPartialHashCount = 0
    $lastTargetPartialHashPercent = -1
    $targetFullCandidateList = New-Object System.Collections.Generic.List[object]
    foreach ($partialHashRecord in @(Get-ContentHashRecordList `
                -FileList $targetPartialCandidateFileList.ToArray() `
                -HashKind 'Partial' `
                -Activity '目标目录重复文件筛选' `
                -Status '正在并行计算目标部分哈希' `
                -TotalCount $targetPartialCandidateFileList.Count `
                -ProcessedCount ([ref]$processedTargetPartialHashCount) `
                -LastPercent ([ref]$lastTargetPartialHashPercent) `
                -WarningList $matchWarningList `
                -WarningMessage '跳过文件，无法计算部分哈希')) {
        $file = $partialHashRecord.File
        $partialHash = $partialHashRecord.Hash
        $partialHashIndex = Get-ReferencePartialHashIndexFromCache -ReferenceIndex $ReferenceIndex -Length $file.Length
        if (-not $partialHashIndex.ContainsKey($partialHash)) {
            continue
        }

        if (Test-PartialHashCoversFullFile -Length $file.Length) {
            $fullHash = $partialHash
            if (-not $matchedTargetFilesByHash.ContainsKey($fullHash)) {
                $targetFileListForHash = New-Object System.Collections.Generic.List[System.IO.FileInfo]
                $matchedTargetFilesByHash[$fullHash] = [pscustomobject]@{
                    Hash           = $fullHash
                    ReferenceFiles = @($partialHashIndex[$partialHash])
                    TargetFiles    = $targetFileListForHash
                }
            }

            $matchedTargetFilesByHash[$fullHash].TargetFiles.Add($file)
            continue
        }

        $targetFullCandidateList.Add([pscustomobject]@{
                File        = $file
                Length      = $file.Length
                PartialHash = $partialHash
            })
    }

    Initialize-ReferenceFullHashIndexForMatchList `
        -ReferenceIndex $ReferenceIndex `
        -MatchList $targetFullCandidateList.ToArray() `
        -WarningList $matchWarningList

    $targetFullCandidateContextByFullName = @{}
    $targetFullCandidateFileList = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    foreach ($targetFullCandidate in $targetFullCandidateList) {
        $fullHashIndex = Get-ReferenceFullHashIndexFromCache `
            -ReferenceIndex $ReferenceIndex `
            -Length $targetFullCandidate.Length `
            -PartialHash $targetFullCandidate.PartialHash
        if ($fullHashIndex.Count -eq 0) {
            continue
        }

        $targetFullCandidateContextByFullName[$targetFullCandidate.File.FullName] = $targetFullCandidate
        $targetFullCandidateFileList.Add($targetFullCandidate.File)
    }

    $processedTargetFullHashCount = 0
    $lastTargetFullHashPercent = -1
    foreach ($fullHashRecord in @(Get-ContentHashRecordList `
                -FileList $targetFullCandidateFileList.ToArray() `
                -HashKind 'Full' `
                -Activity '目标目录重复文件筛选' `
                -Status '正在并行计算目标完整哈希' `
                -TotalCount $targetFullCandidateFileList.Count `
                -ProcessedCount ([ref]$processedTargetFullHashCount) `
                -LastPercent ([ref]$lastTargetFullHashPercent) `
                -WarningList $matchWarningList `
                -WarningMessage '跳过文件，无法计算完整哈希')) {
        $targetFullCandidate = $targetFullCandidateContextByFullName[$fullHashRecord.File.FullName]
        $fullHashIndex = Get-ReferenceFullHashIndexFromCache `
            -ReferenceIndex $ReferenceIndex `
            -Length $targetFullCandidate.Length `
            -PartialHash $targetFullCandidate.PartialHash
        if (-not $fullHashIndex.ContainsKey($fullHashRecord.Hash)) {
            continue
        }

        if (-not $matchedTargetFilesByHash.ContainsKey($fullHashRecord.Hash)) {
            $targetFileListForHash = New-Object System.Collections.Generic.List[System.IO.FileInfo]
            $matchedTargetFilesByHash[$fullHashRecord.Hash] = [pscustomobject]@{
                Hash           = $fullHashRecord.Hash
                ReferenceFiles = @($fullHashIndex[$fullHashRecord.Hash])
                TargetFiles    = $targetFileListForHash
            }
        }

        $matchedTargetFilesByHash[$fullHashRecord.Hash].TargetFiles.Add($fullHashRecord.File)
    }

    Complete-DynamicStatusLine
    Write-DeferredScanWarningList -WarningList $matchWarningList -Title '参考匹配跳过汇总'

    $matchedTargetFileCount = 0
    foreach ($matchRecord in $matchedTargetFilesByHash.Values) {
        $matchedTargetFileCount += $matchRecord.TargetFiles.Count
    }
    Write-StageMessage "目标目录匹配参考目录的重复文件数: $matchedTargetFileCount，完整哈希组数: $($matchedTargetFilesByHash.Count)"

    foreach ($matchRecord in $matchedTargetFilesByHash.Values) {
        $matchingReferenceFiles = @($matchRecord.ReferenceFiles)
        $referenceKeepFile = Select-DefaultKeepFile -FileList $matchingReferenceFiles
        $matchingTargetFiles = @($matchRecord.TargetFiles)

        [pscustomobject]@{
            Hash          = $matchRecord.Hash
            KeepPathText = Get-RelativePathText -File $referenceKeepFile -RootPath $referenceRootPath -PathPrefix $referencePathPrefix
            DeletionItems = New-DeletionItemList `
                -FileList $matchingTargetFiles `
                -RootPath $TargetRootPath `
                -PathPrefix $targetPathPrefix `
                -ExpectedHash $matchRecord.Hash `
                -KeepFileList @($referenceKeepFile)
        }
    }
}

# 参考目录去重入口：以第一个目录为参考，只删除后续目标目录中的重复文件。
function Invoke-ReferenceDirectoryMode {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ReferenceRootPath,

        [Parameter(Mandatory = $true)]
        [string[]]$TargetRootPathList
    )

    $referenceIndex = New-ReferenceDirectoryIndex -ReferenceRootPath $ReferenceRootPath
    if (@($referenceIndex.FileList).Count -eq 0) {
        Write-Host "未发现目标目录中存在与参考目录重复的文件。" -ForegroundColor Green
        return 'Continue'
    }

    $deletionPlanList = @(
        for ($targetIndex = 0; $targetIndex -lt $TargetRootPathList.Count; $targetIndex++) {
            if ($TargetRootPathList.Count -gt 1) {
                Write-PreviewSeparator
                Write-Host "参考目录模式目标 $($targetIndex + 1) / $($TargetRootPathList.Count): $($TargetRootPathList[$targetIndex])" -ForegroundColor Cyan
            }

            New-ReferenceDirectoryDeletionPlan -ReferenceIndex $referenceIndex -TargetRootPath $TargetRootPathList[$targetIndex]
        }
    )

    $actionParameters = @{
        DeletionPlanList                      = $deletionPlanList
        EmptyMessage                          = '未发现目标目录中存在与参考目录重复的文件。'
        PreviewSummaryMessageTemplate         = '重复文件列举完成。默认计划从目标目录删除重复文件: {0}'
        DefaultDeletionSummaryMessageTemplate = '删除完成。已从目标目录删除重复文件: {0}'
    }
    return (Invoke-DeletionPlanAction @actionParameters)
}

# 校验路径并执行一轮扫描流程；返回 Continue 或 Exit，供交互菜单判断后续动作。
function Invoke-DuplicateScanRun {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$InputPathList,

        [Parameter(Mandatory = $true)]
        [bool]$UseAggregateMode,

        [Parameter(Mandatory = $true)]
        [bool]$UseReferenceMode
    )

    $minimumPathCount = Get-ScanModeMinimumPathCount -UseReferenceMode $UseReferenceMode
    if ($InputPathList.Count -lt $minimumPathCount) {
        throw "当前模式至少需要输入 $minimumPathCount 个目录。"
    }

    $resolvedPathList = @(Resolve-InputDirectoryList -PathList $InputPathList)
    if ($resolvedPathList.Count -lt $minimumPathCount) {
        throw "当前模式至少需要输入 $minimumPathCount 个不重复目录。"
    }

    if ($UseReferenceMode) {
        $targetRootPathList = @()
        if ($resolvedPathList.Count -gt 1) {
            $targetRootPathList = @($resolvedPathList[1..($resolvedPathList.Count - 1)])
        }

        return (Invoke-ReferenceDirectoryMode -ReferenceRootPath $resolvedPathList[0] -TargetRootPathList $targetRootPathList)
    }
    elseif ($UseAggregateMode) {
        return (Invoke-MergedDirectoryMode -RootPathList $resolvedPathList)
    }
    else {
        return (Invoke-SeparateSingleDirectoryMode -RootPathList $resolvedPathList)
    }
}

if ($Help) {
    Show-HelpText
    exit 0
}

if ($AggregateMode -and $ReferenceMode) {
    Write-Host "-a 和 -c 不能同时使用。" -ForegroundColor Red
    exit 1
}

$useAggregateMode = [bool]$AggregateMode
$useReferenceMode = [bool]$ReferenceMode
$inputPathList = @($PathList | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
$hasCommandLinePathList = $inputPathList.Count -gt 0

Write-EnabledOptionNotice -IncludeHiddenItems $ShouldIncludeHiddenItems -AssumeYesDeletion $ShouldAssumeYesDeletion

if (-not $hasCommandLinePathList) {
    $shouldReadScanMode = -not $useAggregateMode -and -not $useReferenceMode

    while ($true) {
        $hasInteractivePathList = $false

        while (-not $hasInteractivePathList) {
            if ($shouldReadScanMode) {
                $selectedScanMode = Read-InteractiveScanMode
                switch ($selectedScanMode) {
                    'Single' {
                        $useAggregateMode = $false
                        $useReferenceMode = $false
                    }
                    'Aggregate' {
                        $useAggregateMode = $true
                        $useReferenceMode = $false
                    }
                    'Reference' {
                        $useAggregateMode = $false
                        $useReferenceMode = $true
                    }
                    'Exit' {
                        Write-Host "已退出，未执行扫描。" -ForegroundColor Yellow
                        exit 0
                    }
                }
            }

            $minimumPathCount = Get-ScanModeMinimumPathCount -UseReferenceMode $useReferenceMode
            $modePrompt = Get-ScanModePrompt -UseReferenceMode $useReferenceMode -UseAggregateMode $useAggregateMode
            $pathInputResult = Read-InteractivePathList -ModePrompt $modePrompt -MinimumCount $minimumPathCount

            switch ($pathInputResult.Action) {
                'Submit' {
                    $inputPathList = @($pathInputResult.PathList)
                    $hasInteractivePathList = $true
                }
                'Back' {
                    Write-Host "已返回上级菜单。" -ForegroundColor Yellow
                    $useAggregateMode = $false
                    $useReferenceMode = $false
                    $shouldReadScanMode = $true
                }
                'Exit' {
                    Write-Host "已退出，未执行扫描。" -ForegroundColor Yellow
                    exit 0
                }
                default {
                    Write-Host "路径输入返回了未知状态: $($pathInputResult.Action)" -ForegroundColor Red
                    exit 1
                }
            }
        }

        try {
            $scanRunResult = Invoke-DuplicateScanRun -InputPathList $inputPathList -UseAggregateMode $useAggregateMode -UseReferenceMode $useReferenceMode
        }
        catch {
            Write-Host $_.Exception.Message -ForegroundColor Red
            $scanRunResult = 'Continue'
        }

        if ($scanRunResult -eq 'Exit' -or $AssumeYesDeletionCancelled) {
            exit 0
        }

        Write-Host ""
        Write-Host "本轮流程完成，返回扫描模式菜单。" -ForegroundColor Cyan
        $inputPathList = @()
        $useAggregateMode = $false
        $useReferenceMode = $false
        $shouldReadScanMode = $true
    }
}

try {
    $scanRunResult = Invoke-DuplicateScanRun -InputPathList $inputPathList -UseAggregateMode $useAggregateMode -UseReferenceMode $useReferenceMode
}
catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

if ($scanRunResult -eq 'Exit' -or $AssumeYesDeletionCancelled) {
    exit 0
}
