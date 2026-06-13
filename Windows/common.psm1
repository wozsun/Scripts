# Windows 脚本通用工具函数。
# 仅放跨脚本复用的输出、进度、延迟警告和路径处理逻辑；业务流程仍保留在各脚本内。
# 设计原则：
# - common 只提供“怎么输出、怎么解析路径、怎么比较目录”这类基础能力。
# - 删除、转换、哈希等业务规则留在具体脚本里，避免公共模块变成隐式流程控制器。

Set-StrictMode -Version Latest

# ========== 可调整配置 ==========
# 以下变量可根据需要修改，控制输出样式和交互行为。

# 文本进度条的槽位数量；32 在常见终端宽度下足够清楚，也不容易挤占状态文本。
$script:ProgressBarCellCount = 32

# 进度条已完成部分使用的字符；使用 ASCII，避免不同 Windows 终端字体显示宽度不一致。
$script:ProgressBarFilledCharacter = '#'

# 进度条未完成部分使用的字符；和已完成字符保持同宽，动态刷新时不容易抖动。
$script:ProgressBarEmptyCharacter = '-'

# 进度条同百分比状态下的强制刷新间隔；设为 0 或负数则只在百分比变化时刷新。
$script:ProgressBarTimedRefreshMilliseconds = 1000

# 预览分隔线长度；用于重复文件组、扫描阶段等块状输出的视觉分隔。
$script:PreviewSeparatorCellCount = 64

# 预览分隔线字符；使用 ASCII，便于复制日志到纯文本环境。
$script:PreviewSeparatorCharacter = '='

# -yes 默认删除前的等待秒数；给用户留出按 Enter 取消的窗口。
$script:AssumeYesGraceSeconds = 10

# -yes 倒计时期间检查键盘输入的间隔；越小响应越快，但会更频繁轮询控制台。
$script:AssumeYesInputPollIntervalMilliseconds = 100

# ========== 运行状态 ==========
# 以下变量由函数内部维护，用于跨调用追踪动态输出状态；请勿手动修改。

# 输出相关运行状态；由动态状态行和进度条函数维护，请勿手动修改。
$script:OutputRuntimeState = [pscustomobject]@{
    # 上一次动态状态行的显示宽度；用于本次刷新时清掉旧尾巴。
    DynamicStatusLastCellWidth            = 0
    # 记录上一条动态状态是否真的以内联方式输出；输出重定向或降级时不额外补换行。
    DynamicStatusLastWriteWasInline       = $false
    # 记录各进度条最近一次刷新时间；用于同百分比但耗时较长时按间隔刷新计数。
    ProgressBarLastRefreshMillisecondsByKey  = @{}
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
    # [Console]::ReadLine() 在输入流结束时返回 $null；调用方需要明确决定是退出、跳过还是报错。
    return [Console]::ReadLine()
}

# 输出通用菜单并读取用户选择。
function Read-MenuChoice {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [object[]]$MenuOptionList,

        [Parameter(Mandatory = $false)]
        [string]$EndOfInputChoice = '00'
    )

    # 菜单项约定为带 Value / Label 的对象；这样调用方可以用业务语义封装返回值。
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
            # 不在 common 中直接 exit，避免公共函数替调用方决定进程生命周期。
            return $EndOfInputChoice
        }

        $menuChoice = $inputRaw.Trim()

        if ($validMenuChoices -contains $menuChoice) {
            return $menuChoice
        }

        Write-Host "输入无效，请输入: $($validMenuChoices -join ', ')" -ForegroundColor Red
    }
}

# 估算字符在控制台中占用的单元格宽度；中文和全角字符通常占两格。
function Get-ConsoleCharacterCellWidth {
    param(
        [Parameter(Mandatory = $true)]
        [char]$Character
    )

    $codePoint = [int]$Character
    # 这里只覆盖常见 CJK / 全角区间，目标是让进度行“足够不换行”，不是实现完整 Unicode 宽度算法。
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

# 将文本限制在指定控制台宽度内，避免动态状态行因过长而换行。
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
        # 先判断再追加，保证返回文本本身不会越过 MaxCellWidth。
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

# 估算字符串显示宽度。
function Get-ConsoleTextWidth {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    return (Get-ConsoleTextWithinCellWidth -Text $Text -MaxCellWidth ([int]::MaxValue)).CellWidth
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

        # WindowWidth - 1 可减少刚好写满控制台宽度时自动换行的概率。
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
    $refreshIntervalMilliseconds = [Math]::Max(0, [int]$script:ProgressBarTimedRefreshMilliseconds)
    $shouldRefreshByPercent = ($percent -ne $LastPercent.Value)
    $shouldRefreshByTime = $false

    if ($refreshIntervalMilliseconds -gt 0 -and $script:OutputRuntimeState.ProgressBarLastRefreshMillisecondsByKey.ContainsKey($refreshKey)) {
        $lastRefreshMilliseconds = [long]$script:OutputRuntimeState.ProgressBarLastRefreshMillisecondsByKey[$refreshKey]
        $shouldRefreshByTime = (($currentMilliseconds - $lastRefreshMilliseconds) -ge $refreshIntervalMilliseconds)
    }

    # 大目录扫描时大量重复绘制会明显拖慢输出；百分比变化或同百分比超过配置间隔时再刷新。
    if (-not $Force -and -not $shouldRefreshByPercent -and -not $shouldRefreshByTime) {
        return
    }

    $filledWidth = [Math]::Floor(($percent / 100) * $script:ProgressBarCellCount)
    $emptyWidth = $script:ProgressBarCellCount - $filledWidth
    $bar = ($script:ProgressBarFilledCharacter * $filledWidth) + ($script:ProgressBarEmptyCharacter * $emptyWidth)
    $progressText = "[进度] $Activity [$bar] $percent% $Status ($ProcessedCount / $TotalCount)"

    Write-DynamicStatusLine -Message $progressText -Color Cyan
    $LastPercent.Value = $percent
    $script:OutputRuntimeState.ProgressBarLastRefreshMillisecondsByKey[$refreshKey] = $currentMilliseconds
}

# 使用同一行刷新状态，适合“正在处理 -> 处理完成”这种短状态。
function Write-RefreshStatusLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [System.ConsoleColor]$Color,

        [Parameter(Mandatory = $false)]
        [switch]$NoNewLine
    )

    # 返回宽度是给调用方在同一行追加文本或做状态覆盖时使用的。
    $messageWidth = Get-ConsoleTextWidth -Text ($Message -replace '[\r\n]+', ' ')
    Write-DynamicStatusLine -Message $Message -Color $Color

    if ($NoNewLine) {
        return $messageWidth
    }

    Complete-DynamicStatusLine
    return 0
}

# 输出用于区分不同预览或结果块的分隔线。
function Write-PreviewSeparator {
    param(
        [Parameter(Mandatory = $false)]
        [switch]$NoLeadingBlank
    )

    # 部分调用点已经手动输出了标题和空行，可以用 -NoLeadingBlank 避免多出一行。
    Complete-DynamicStatusLine
    if (-not $NoLeadingBlank) {
        Write-Host ""
    }

    Write-Host ($script:PreviewSeparatorCharacter * $script:PreviewSeparatorCellCount) -ForegroundColor DarkGray
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
        # 输入重定向时 KeyAvailable/ReadKey 不适用；直接返回 false，避免异常打断倒计时。
        if ([Console]::IsInputRedirected) {
            return $false
        }

        # 清空缓冲区里已经按下的键；只把 Enter 视为取消信号。
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

# 危险操作执行前提供醒目的中止窗口，调用方负责传入符合具体操作的提示文本。
function Wait-DangerousOperationGracePeriod {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WarningMessage,

        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$AdditionalWarningMessage = '',

        [Parameter(Mandatory = $false)]
        [string]$CancelHintMessage = '如需取消，请在倒计时结束前按 Enter；也可按 Ctrl+C 强制中止。',

        [Parameter(Mandatory = $true)]
        [string]$CancelledMessage,

        [Parameter(Mandatory = $true)]
        [string]$CompletedMessage,

        [Parameter(Mandatory = $true)]
        [string]$CountdownMessageFormat
    )

    # 这里只负责“等待并返回是否继续”。具体取消后是退出本轮还是退出进程，由调用方决定。
    Complete-DynamicStatusLine
    Write-Host ""
    Write-Host $WarningMessage -ForegroundColor Red
    if (-not [string]::IsNullOrWhiteSpace($AdditionalWarningMessage)) {
        Write-Host $AdditionalWarningMessage -ForegroundColor Yellow
    }
    Write-Host $CancelHintMessage -ForegroundColor Yellow

    $pollInterval = [Math]::Max(1, $script:AssumeYesInputPollIntervalMilliseconds)
    $pollCountPerSecond = [Math]::Max(1, [int][Math]::Ceiling(1000 / $pollInterval))
    for ($remainingSeconds = $script:AssumeYesGraceSeconds; $remainingSeconds -gt 0; $remainingSeconds--) {
        Write-DynamicStatusLine -Message ($CountdownMessageFormat -f $remainingSeconds) -Color Yellow

        # 每秒拆成多个短 sleep，保证按 Enter 后能较快响应。
        for ($pollIndex = 0; $pollIndex -lt $pollCountPerSecond; $pollIndex++) {
            if (Test-EnterKeyPressed) {
                Write-DynamicStatusLine -Message $CancelledMessage -Color Yellow
                Complete-DynamicStatusLine
                return $false
            }

            Start-Sleep -Milliseconds $pollInterval
        }
    }

    Write-DynamicStatusLine -Message $CompletedMessage -Color Magenta
    Complete-DynamicStatusLine
    return $true
}

# 新建延迟输出的扫描警告列表，避免进度条刷新时被错误信息打断。
function New-DeferredScanWarningList {
    # 前置逗号确保 PowerShell 不会在只有一个元素时把集合自动展开成单个对象。
    $warningList = [System.Collections.Generic.List[object]]::new()
    return , $warningList
}

# 记录扫描、哈希、移动等可跳过错误；未传入列表时退化为立即输出。
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
        # 允许调用方不建立延迟列表；此时退化为即时输出，方便小工具复用。
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

# 在当前进度段结束后统一输出警告，保持动态进度条单行刷新。
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

# 判断文件路径是否存在；使用 .NET 避免 Test-Path 的管道和 Provider 开销。
function Test-FileSystemFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return [System.IO.File]::Exists($Path)
}

# 判断目录路径是否存在；使用 .NET 避免 Test-Path 的管道和 Provider 开销。
function Test-FileSystemDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return [System.IO.Directory]::Exists($Path)
}

# 判断文件或目录路径是否存在。
function Test-FileSystemPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return ((Test-FileSystemFile -Path $Path) -or (Test-FileSystemDirectory -Path $Path))
}

# 获取文件系统对象；路径不存在时返回 null。
function Get-FileSystemItem {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    if ([System.IO.Directory]::Exists($fullPath)) {
        return [System.IO.DirectoryInfo]::new($fullPath)
    }

    if ([System.IO.File]::Exists($fullPath)) {
        return [System.IO.FileInfo]::new($fullPath)
    }

    return $null
}

# 创建目录并返回 DirectoryInfo。
function New-FileSystemDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return [System.IO.Directory]::CreateDirectory($Path)
}

# 删除文件或目录；删除前尽量移除常见保护属性，适合清理脚本临时目录和临时文件。
function Remove-FileSystemItem {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [switch]$Recurse
    )

    if ([System.IO.File]::Exists($Path)) {
        [System.IO.File]::SetAttributes($Path, [System.IO.FileAttributes]::Normal)
        [System.IO.File]::Delete($Path)
        if ([System.IO.File]::Exists($Path)) {
            throw "删除后文件仍存在: $Path"
        }

        return
    }

    if (-not [System.IO.Directory]::Exists($Path)) {
        return
    }

    $directoryInfo = [System.IO.DirectoryInfo]::new($Path)
    $attributesToClear = [int](
        [System.IO.FileAttributes]::ReadOnly -bor
        [System.IO.FileAttributes]::Hidden -bor
        [System.IO.FileAttributes]::System
    )

    if ($Recurse) {
        foreach ($childFile in $directoryInfo.EnumerateFiles('*', [System.IO.SearchOption]::AllDirectories)) {
            try {
                $childFile.Attributes = [System.IO.FileAttributes]::Normal
            }
            catch {
                # 删除时会再次抛出真实错误；这里不提前中断。
                Write-Debug "清理文件属性失败: $($childFile.FullName)。原因: $($_.Exception.Message)"
            }
        }

        foreach ($childDirectory in $directoryInfo.EnumerateDirectories('*', [System.IO.SearchOption]::AllDirectories)) {
            try {
                $childDirectory.Attributes = [System.IO.FileAttributes](([int]$childDirectory.Attributes) -band (-bnot $attributesToClear))
            }
            catch {
                # 删除时会再次抛出真实错误；这里不提前中断。
                Write-Debug "清理目录属性失败: $($childDirectory.FullName)。原因: $($_.Exception.Message)"
            }
        }
    }

    try {
        $directoryInfo.Attributes = [System.IO.FileAttributes](([int]$directoryInfo.Attributes) -band (-bnot $attributesToClear))
    }
    catch {
        # 删除时会再次抛出真实错误；这里不提前中断。
        Write-Debug "清理目录属性失败: $Path。原因: $($_.Exception.Message)"
    }

    [System.IO.Directory]::Delete($Path, [bool]$Recurse)
    if ([System.IO.Directory]::Exists($Path)) {
        throw "删除后目录仍存在: $Path"
    }
}

# 获取目录中的项目数量；可选择在扫描失败时按非空处理，适合临时目录安全判断。
function Get-DirectoryItemCount {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath,

        [Parameter(Mandatory = $false)]
        [switch]$TreatErrorAsNonEmpty
    )

    try {
        $itemCount = 0
        foreach ($itemPath in [System.IO.Directory]::EnumerateFileSystemEntries($DirectoryPath)) {
            $itemCount++
        }

        return $itemCount
    }
    catch {
        if ($TreatErrorAsNonEmpty) {
            return 1
        }

        throw
    }
}

# 校验脚本专属临时目录名，避免调用方传入路径导致清理越界。
function Assert-SiblingTempDirectoryName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempDirectoryName
    )

    if ([string]::IsNullOrWhiteSpace($TempDirectoryName)) {
        throw '临时目录名不能为空。'
    }

    if ([System.IO.Path]::IsPathRooted($TempDirectoryName) -or
        $TempDirectoryName.Contains([System.IO.Path]::DirectorySeparatorChar) -or
        $TempDirectoryName.Contains([System.IO.Path]::AltDirectorySeparatorChar)) {
        throw "临时目录名不能包含路径或路径分隔符: $TempDirectoryName"
    }

    if ($TempDirectoryName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0) {
        throw "临时目录名包含非法字符: $TempDirectoryName"
    }
}

# 获取同目录下的脚本专属临时目录路径。
function Get-SiblingTempDirectoryPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ParentPath,

        [Parameter(Mandatory = $true)]
        [string]$TempDirectoryName
    )

    Assert-SiblingTempDirectoryName -TempDirectoryName $TempDirectoryName
    return [System.IO.Path]::Combine($ParentPath, $TempDirectoryName)
}

# 判断目录名是否为脚本专属临时目录名。
function Test-SiblingTempDirectoryName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryName,

        [Parameter(Mandatory = $true)]
        [string]$TempDirectoryName
    )

    Assert-SiblingTempDirectoryName -TempDirectoryName $TempDirectoryName
    if ([string]::IsNullOrWhiteSpace($DirectoryName)) {
        return $false
    }

    return $DirectoryName.Equals($TempDirectoryName, [System.StringComparison]::OrdinalIgnoreCase)
}

# 判断路径是否为指定父目录下的脚本专属临时目录。
function Test-SiblingTempDirectoryPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$ParentPath,

        [Parameter(Mandatory = $true)]
        [string]$TempDirectoryName
    )

    Assert-SiblingTempDirectoryName -TempDirectoryName $TempDirectoryName

    try {
        $actualPath = ConvertTo-NormalizedPath -Path $Path
        $actualParentPath = [System.IO.Path]::GetDirectoryName($actualPath)
        if ([string]::IsNullOrWhiteSpace($actualParentPath)) {
            return $false
        }

        $actualParentKey = ConvertTo-NormalizedPath -Path $actualParentPath
        $expectedParentKey = ConvertTo-NormalizedPath -Path $ParentPath
        if (-not $actualParentKey.Equals($expectedParentKey, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $false
        }

        $actualDirectoryName = [System.IO.Path]::GetFileName($actualPath)
        return Test-SiblingTempDirectoryName -DirectoryName $actualDirectoryName -TempDirectoryName $TempDirectoryName
    }
    catch {
        return $false
    }
}

# 创建同目录脚本专属临时目录；已存在的专属临时目录会被清空并复用。
function Initialize-SiblingTempDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ParentPath,

        [Parameter(Mandatory = $true)]
        [string]$TempDirectoryName
    )

    if (-not (Test-FileSystemDirectory -Path $ParentPath)) {
        throw "临时目录父目录不存在或不可用: $ParentPath"
    }

    $tempDirectoryPath = Get-SiblingTempDirectoryPath -ParentPath $ParentPath -TempDirectoryName $TempDirectoryName
    if (Test-FileSystemFile -Path $tempDirectoryPath) {
        throw "临时目录路径被同名文件占用: $tempDirectoryPath"
    }

    if (Test-FileSystemDirectory -Path $tempDirectoryPath) {
        Clear-SiblingTempDirectoryContents `
            -TempDirectoryPath $tempDirectoryPath `
            -ParentPath $ParentPath `
            -TempDirectoryName $TempDirectoryName
        return (Get-FileSystemItem -Path $tempDirectoryPath).FullName
    }

    return (New-FileSystemDirectory -Path $tempDirectoryPath).FullName
}

# 清空脚本专属临时目录内容；调用前会校验目录名，避免清空普通目录。
function Clear-SiblingTempDirectoryContents {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempDirectoryPath,

        [Parameter(Mandatory = $true)]
        [string]$ParentPath,

        [Parameter(Mandatory = $true)]
        [string]$TempDirectoryName
    )

    if (Test-FileSystemFile -Path $TempDirectoryPath) {
        throw "临时路径不是文件夹: $TempDirectoryPath"
    }

    if (-not (Test-FileSystemDirectory -Path $TempDirectoryPath)) {
        return
    }

    if (-not (Test-SiblingTempDirectoryPath -Path $TempDirectoryPath -ParentPath $ParentPath -TempDirectoryName $TempDirectoryName)) {
        throw "拒绝清理非脚本临时目录: $TempDirectoryPath"
    }

    foreach ($childItem in ([System.IO.DirectoryInfo]::new($TempDirectoryPath)).EnumerateFileSystemInfos()) {
        Remove-FileSystemItem -Path $childItem.FullName -Recurse
    }
}

# 删除脚本专属临时目录；调用前会校验目录名，避免删除普通目录。
function Remove-SiblingTempDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempDirectoryPath,

        [Parameter(Mandatory = $true)]
        [string]$ParentPath,

        [Parameter(Mandatory = $true)]
        [string]$TempDirectoryName
    )

    if (-not (Test-FileSystemPath -Path $TempDirectoryPath)) {
        return
    }

    if (Test-FileSystemFile -Path $TempDirectoryPath) {
        throw "临时路径不是文件夹: $TempDirectoryPath"
    }

    if (-not (Test-SiblingTempDirectoryPath -Path $TempDirectoryPath -ParentPath $ParentPath -TempDirectoryName $TempDirectoryName)) {
        throw "拒绝清理非脚本临时目录: $TempDirectoryPath"
    }

    Remove-FileSystemItem -Path $TempDirectoryPath -Recurse
}

# 仅当脚本专属临时目录为空时删除；非空返回 false，交由调用方决定提示方式。
function Remove-EmptySiblingTempDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempDirectoryPath,

        [Parameter(Mandatory = $true)]
        [string]$ParentPath,

        [Parameter(Mandatory = $true)]
        [string]$TempDirectoryName
    )

    if (-not (Test-FileSystemPath -Path $TempDirectoryPath)) {
        return $true
    }

    if (Test-FileSystemFile -Path $TempDirectoryPath) {
        throw "临时目录路径被同名文件占用: $TempDirectoryPath"
    }

    if (-not (Test-SiblingTempDirectoryPath -Path $TempDirectoryPath -ParentPath $ParentPath -TempDirectoryName $TempDirectoryName)) {
        throw "拒绝清理非脚本临时目录: $TempDirectoryPath"
    }

    if (Get-DirectoryItemCount -DirectoryPath $TempDirectoryPath -TreatErrorAsNonEmpty) {
        return $false
    }

    Remove-FileSystemItem -Path $TempDirectoryPath
    return -not (Test-FileSystemPath -Path $TempDirectoryPath)
}

# 输出相对路径用于展示，避免终端日志被完整绝对路径撑得过长。
function Get-RelativePathText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath,

        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $false)]
        [string]$PathPrefix
    )

    $relativePathText = [System.IO.Path]::GetRelativePath($RootPath, $FilePath)
    if ($relativePathText.StartsWith('..', [System.StringComparison]::Ordinal) -or
        [System.IO.Path]::IsPathRooted($relativePathText)) {
        # FilePath 不在 RootPath 下时只显示文件名，避免把跨目录绝对路径混进预览列表。
        $relativePathText = [System.IO.Path]::GetFileName($FilePath)
    }

    if ([string]::IsNullOrWhiteSpace($PathPrefix)) {
        return $relativePathText
    }

    return "$PathPrefix\$relativePathText"
}

# 处理交互输入时复制带首尾英文引号的路径；这里只移除成对包裹符号。
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
    # 用字符码判断英文单/双引号，避免源文件编码或中文引号造成误判。
    if (($firstCharacterCode -eq 34 -and $lastCharacterCode -eq 34) -or
        ($firstCharacterCode -eq 39 -and $lastCharacterCode -eq 39)) {
        return $normalizedPathText.Substring(1, $normalizedPathText.Length - 2).Trim()
    }

    return $normalizedPathText
}

# 判断输入文本是否为 Windows 绝对路径。
function Test-WindowsAbsolutePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathText
    )

    try {
        return [System.IO.Path]::IsPathFullyQualified($PathText)
    }
    catch {
        return $false
    }
}

# 判断文件系统对象是否为隐藏项。
function Test-HiddenFileSystemItem {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileSystemInfo]$Item
    )

    return (($Item.Attributes -band [System.IO.FileAttributes]::Hidden) -ne 0)
}

# 解析单个用户输入路径，并按需要限制文件或目录类型。
function Resolve-InputPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathText,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Any', 'File', 'Directory')]
        [string]$PathType = 'Any'
    )

    $cleanPath = ConvertTo-UnquotedPathText -PathText $PathText
    if ([string]::IsNullOrWhiteSpace($cleanPath)) {
        return [pscustomobject]@{
            Success = $false
            Item    = $null
            Error   = '路径为空。'
        }
    }

    if (-not (Test-WindowsAbsolutePath -PathText $cleanPath)) {
        return [pscustomobject]@{
            Success = $false
            Item    = $null
            Error   = '请输入 Windows 绝对路径。'
        }
    }

    $resolvedItem = $null
    try {
        $resolvedItem = Get-FileSystemItem -Path $cleanPath
    }
    catch {
        return [pscustomobject]@{
            Success = $false
            Item    = $null
            Error   = $_.Exception.Message
        }
    }

    if ($null -eq $resolvedItem) {
        return [pscustomobject]@{
            Success = $false
            Item    = $null
            Error   = '路径不存在。'
        }
    }

    try {
        if ($PathType -eq 'File' -and $resolvedItem -is [System.IO.DirectoryInfo]) {
            return [pscustomobject]@{
                Success = $false
                Item    = $null
                Error   = '请输入文件路径。'
            }
        }

        if ($PathType -eq 'Directory' -and $resolvedItem -isnot [System.IO.DirectoryInfo]) {
            return [pscustomobject]@{
                Success = $false
                Item    = $null
                Error   = '请输入文件夹路径。'
            }
        }

        return [pscustomobject]@{
            Success = $true
            Item    = $resolvedItem
            Error   = $null
        }
    }
    catch {
        return [pscustomobject]@{
            Success = $false
            Item    = $null
            Error   = $_.Exception.Message
        }
    }
}

# 批量解析用户输入路径，并按规范化后的完整路径静默去重。
function Resolve-InputPathList {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$PathList,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Any', 'File', 'Directory')]
        [string]$PathType = 'Any',

        [Parameter(Mandatory = $false)]
        [System.Collections.Generic.HashSet[string]]$ExistingKeySet
    )

    if ($null -eq $ExistingKeySet) {
        $ExistingKeySet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }

    $resolvedItemList = [System.Collections.Generic.List[System.IO.FileSystemInfo]]::new()
    $pendingKeySet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($pathText in $PathList) {
        $resolvedResult = Resolve-InputPath -PathText $pathText -PathType $PathType
        if (-not $resolvedResult.Success) {
            return [pscustomobject]@{
                Success = $false
                Items   = @()
                Error   = "路径无效: $pathText。$($resolvedResult.Error)"
            }
        }

        $itemKey = ConvertTo-NormalizedPath -Path $resolvedResult.Item.FullName
        if (-not $ExistingKeySet.Contains($itemKey) -and $pendingKeySet.Add($itemKey)) {
            $resolvedItemList.Add($resolvedResult.Item)
        }
    }

    foreach ($itemKey in $pendingKeySet) {
        [void]$ExistingKeySet.Add($itemKey)
    }

    return [pscustomobject]@{
        Success = $true
        Items   = $resolvedItemList.ToArray()
        Error   = $null
    }
}

# 批量解析目录输入、按真实路径去重，并拒绝父子目录。
function Resolve-IndependentInputDirectoryList {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$PathList,

        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [string[]]$ExistingDirectoryPathList = @()
    )

    $existingKeySet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($existingDirectoryPath in $ExistingDirectoryPathList) {
        [void]$existingKeySet.Add((ConvertTo-NormalizedPath -Path $existingDirectoryPath))
    }

    $resolvedResult = Resolve-InputPathList `
        -PathList $PathList `
        -PathType Directory `
        -ExistingKeySet $existingKeySet
    if (-not $resolvedResult.Success) {
        return [pscustomobject]@{
            Success = $false
            Items   = @()
            Paths   = @()
            Error   = $resolvedResult.Error
        }
    }

    $resolvedDirectoryPathList = [System.Collections.Generic.List[string]]::new()
    foreach ($resolvedItem in @($resolvedResult.Items)) {
        $resolvedDirectoryPathList.Add((Get-DirectoryKey -DirectoryPath $resolvedItem.FullName))
    }

    try {
        Assert-NoParentChildDirectorySet `
            -DirectoryPathList $resolvedDirectoryPathList.ToArray() `
            -ExistingDirectoryPathList $ExistingDirectoryPathList
    }
    catch {
        return [pscustomobject]@{
            Success = $false
            Items   = @()
            Paths   = @()
            Error   = $_.Exception.Message
        }
    }

    return [pscustomobject]@{
        Success = $true
        Items   = $resolvedResult.Items
        Paths   = $resolvedDirectoryPathList.ToArray()
        Error   = $null
    }
}

# 拆分交互输入的路径行：英文引号用于包裹含空格或分隔符的路径；路径可用英文分号分隔，也可在下一个片段是绝对路径时按空格分隔。
function Split-InteractivePathInput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathInput
    )

    $trimmedPathInput = $PathInput.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmedPathInput)) {
        return @()
    }

    $pathPartList = [System.Collections.Generic.List[string]]::new()
    $currentPartBuilder = [System.Text.StringBuilder]::new()
    $quoteCloseByOpen = @{}
    $quoteCloseByOpen.Add([string][char]34, [string][char]34)
    $quoteCloseByOpen.Add([string][char]39, [string][char]39)
    $activeClosingQuote = $null

    # 第一段先按英文分号拆；引号内的分号视为路径内容。
    for ($index = 0; $index -lt $trimmedPathInput.Length; $index++) {
        $currentChar = [string]$trimmedPathInput[$index]
        if ($null -ne $activeClosingQuote) {
            if ($currentChar -eq $activeClosingQuote) {
                $activeClosingQuote = $null
            }
            else {
                [void]$currentPartBuilder.Append($currentChar)
            }
            continue
        }

        $currentInputText = $currentPartBuilder.ToString()
        $canStartQuotedPath = [string]::IsNullOrWhiteSpace($currentInputText)
        if (-not $canStartQuotedPath) {
            $canStartQuotedPath = [char]::IsWhiteSpace($currentInputText[$currentInputText.Length - 1])
        }

        if ($canStartQuotedPath -and $quoteCloseByOpen.ContainsKey($currentChar)) {
            $activeClosingQuote = $quoteCloseByOpen[$currentChar]
            continue
        }

        if ($currentChar -eq ';') {
            $pathPart = $currentPartBuilder.ToString().Trim()
            if (-not [string]::IsNullOrWhiteSpace($pathPart)) {
                $pathPartList.Add($pathPart)
            }

            [void]$currentPartBuilder.Clear()
            continue
        }

        [void]$currentPartBuilder.Append($currentChar)
    }

    $lastPathPart = $currentPartBuilder.ToString().Trim()
    if (-not [string]::IsNullOrWhiteSpace($lastPathPart)) {
        $pathPartList.Add($lastPathPart)
    }

    $expandedPathList = [System.Collections.Generic.List[string]]::new()
    $quoteStartPattern = @(
        [regex]::Escape([string][char]34)
        [regex]::Escape([string][char]39)
    ) -join '|'
    # 第二段按“空白后面紧跟 Windows 绝对路径”拆，支持同一行粘贴多个未加引号的绝对路径。
    $absolutePathSeparatorPattern = '\s+(?=(?:' + $quoteStartPattern + ')?(?:[a-zA-Z]:[\\/]|[\\/]{2}))'
    foreach ($pathPart in $pathPartList) {
        foreach ($expandedPath in ($pathPart -split $absolutePathSeparatorPattern)) {
            $normalizedPath = ConvertTo-UnquotedPathText -PathText $expandedPath
            if (-not [string]::IsNullOrWhiteSpace($normalizedPath)) {
                $expandedPathList.Add($normalizedPath)
            }
        }
    }

    return $expandedPathList.ToArray()
}

# 将文件或目录路径标准化为便于比较的形式；文件路径调用时去除尾部分隔符是空操作。
function ConvertTo-NormalizedPath {
    param(
        [Parameter(Mandatory = $true)]
        [Alias('DirectoryPath')]
        [string]$Path
    )

    # GetFullPath 先消解 . 和 ..；TrimEndingDirectorySeparator 再统一去掉尾部分隔符。
    return [System.IO.Path]::TrimEndingDirectorySeparator([System.IO.Path]::GetFullPath($Path))
}

# 生成用于比较的目录键，统一去除尾部分隔符。
function Get-DirectoryKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath
    )

    return ConvertTo-NormalizedPath -Path $DirectoryPath
}

# 给目录路径追加结尾分隔符，避免 C:\A 和 C:\AB 这种前缀误判；根目录已带分隔符时不再重复追加。
function Add-TrailingDirectorySeparator {
    param(
        [Parameter(Mandatory = $true)]
        [Alias('Path')]
        [string]$DirectoryPath
    )

    if ($DirectoryPath.EndsWith('\', [System.StringComparison]::Ordinal) -or
        $DirectoryPath.EndsWith('/', [System.StringComparison]::Ordinal)) {
        return $DirectoryPath
    }

    return "$DirectoryPath$([System.IO.Path]::DirectorySeparatorChar)"
}

# 计算目录路径层级，用于从深到浅处理目录。
function Get-DirectoryPathDepth {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath
    )

    # 过滤空段后，C:\A、\\server\share\A 等不同前缀形式的层级比较更稳定。
    $depth = 0
    foreach ($pathSegment in (Get-DirectoryKey -DirectoryPath $DirectoryPath).Split([char[]]@('\', '/'))) {
        if (-not [string]::IsNullOrWhiteSpace($pathSegment)) {
            $depth++
        }
    }

    return $depth
}

# 判断两个目录是否存在父子关系；相同目录不算父子关系，重复路径由路径解析去重处理。
function Test-ParentChildDirectoryPair {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LeftPath,

        [Parameter(Mandatory = $true)]
        [string]$RightPath
    )

    $leftKey = Get-DirectoryKey -DirectoryPath $LeftPath
    $rightKey = Get-DirectoryKey -DirectoryPath $RightPath

    if ($leftKey.Equals($rightKey, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $false
    }

    $leftPrefix = Add-TrailingDirectorySeparator -DirectoryPath $leftKey
    $rightPrefix = Add-TrailingDirectorySeparator -DirectoryPath $rightKey

    # 非父子目录直接视为独立目录；追加尾部分隔符可避免 C:\A 误判为 C:\AB 的父目录。
    return ($rightKey.StartsWith($leftPrefix, [System.StringComparison]::OrdinalIgnoreCase) -or
        $leftKey.StartsWith($rightPrefix, [System.StringComparison]::OrdinalIgnoreCase))
}

# 从目录集合中找出第一组父子目录；可传入已有目录集合用于交互式增量校验。
function Find-ParentChildDirectoryPair {
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
                return [pscustomobject]@{
                    Found       = $true
                    LeftPath    = $ExistingDirectoryPathList[$existingIndex]
                    RightPath   = $DirectoryPathList[$directoryIndex]
                    LeftIndex   = $existingIndex
                    RightIndex  = $directoryIndex
                    LeftSource  = 'Existing'
                    RightSource = 'Input'
                }
            }
        }
    }

    for ($leftIndex = 0; $leftIndex -lt $DirectoryPathList.Count; $leftIndex++) {
        for ($rightIndex = $leftIndex + 1; $rightIndex -lt $DirectoryPathList.Count; $rightIndex++) {
            if (Test-ParentChildDirectoryPair -LeftPath $DirectoryPathList[$leftIndex] -RightPath $DirectoryPathList[$rightIndex]) {
                return [pscustomobject]@{
                    Found       = $true
                    LeftPath    = $DirectoryPathList[$leftIndex]
                    RightPath   = $DirectoryPathList[$rightIndex]
                    LeftIndex   = $leftIndex
                    RightIndex  = $rightIndex
                    LeftSource  = 'Input'
                    RightSource = 'Input'
                }
            }
        }
    }

    return [pscustomobject]@{
        Found       = $false
        LeftPath    = $null
        RightPath   = $null
        LeftIndex   = -1
        RightIndex  = -1
        LeftSource  = $null
        RightSource = $null
    }
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

    $parentChildPair = Find-ParentChildDirectoryPair `
        -DirectoryPathList $DirectoryPathList `
        -ExistingDirectoryPathList $ExistingDirectoryPathList
    if ($parentChildPair.Found) {
        throw '目录不能互为父子目录。'
    }
}

Export-ModuleMember -Function @(
    'Write-StageMessage'
    'Read-ColoredLine'
    'Read-MenuChoice'
    'Complete-DynamicStatusLine'
    'Write-ProgressBar'
    'Write-RefreshStatusLine'
    'Write-PreviewSeparator'
    'Write-StatusSummary'
    'Wait-DangerousOperationGracePeriod'
    'New-DeferredScanWarningList'
    'Add-DeferredScanWarning'
    'Write-DeferredScanWarningList'
    'Test-FileSystemFile'
    'Test-FileSystemDirectory'
    'Test-FileSystemPath'
    'Remove-FileSystemItem'
    'Test-SiblingTempDirectoryName'
    'Initialize-SiblingTempDirectory'
    'Clear-SiblingTempDirectoryContents'
    'Remove-SiblingTempDirectory'
    'Remove-EmptySiblingTempDirectory'
    'Get-RelativePathText'
    'ConvertTo-UnquotedPathText'
    'Test-HiddenFileSystemItem'
    'Resolve-InputPath'
    'Resolve-InputPathList'
    'Resolve-IndependentInputDirectoryList'
    'Split-InteractivePathInput'
    'ConvertTo-NormalizedPath'
    'Get-DirectoryKey'
    'Add-TrailingDirectorySeparator'
    'Get-DirectoryPathDepth'
)
