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

# 预览分隔线长度；用于重复文件组、扫描阶段等块状输出的视觉分隔。
$script:PreviewSeparatorCellCount = 64

# 预览分隔线字符；使用 ASCII，便于复制日志到纯文本环境。
$script:PreviewSeparatorCharacter = '='

# -yes 默认删除前的等待秒数；给用户留出按 Enter 取消的窗口。
$script:AssumeYesGraceSeconds = 10

# -yes 倒计时期间检查键盘输入的间隔；越小响应越快，但会更频繁轮询控制台。
$script:AssumeYesInputPollIntervalMilliseconds = 100

# ========== 内部状态变量 ==========
# 以下变量由函数内部维护，用于跨调用追踪动态输出状态；请勿手动修改。

# 上一次动态状态行的显示宽度；用于本次刷新时清掉旧尾巴。
$script:DynamicStatusLastCellWidth = 0

# 记录上一条动态状态是否真的以内联方式输出；输出重定向或降级时不额外补换行。
$script:DynamicStatusLastWriteWasInline = $false

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

# 刷新单行动态状态；不使用 SetCursorPosition，避免光标定位失败后停在补空格末尾。
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
            $script:DynamicStatusLastCellWidth = 0
            $script:DynamicStatusLastWriteWasInline = $false
            return
        }

        # WindowWidth - 1 可减少刚好写满控制台宽度时自动换行的概率。
        $maxLineWidth = [Math]::Max(1, [Console]::WindowWidth - 1)
        $lineText = Get-ConsoleTextWithinCellWidth -Text $statusMessage -MaxCellWidth $maxLineWidth

        # 先写“当前文本 + 必要补空格”清除上一条长文本尾巴，再回到行首写当前文本。
        # 这样光标自然停在当前文本末尾，不依赖宿主是否支持 SetCursorPosition。
        $clearWidth = [Math]::Min($maxLineWidth, [Math]::Max($script:DynamicStatusLastCellWidth, $lineText.CellWidth))
        $paddingWidth = [Math]::Max(0, $clearWidth - $lineText.CellWidth)
        $paddingText = ' ' * $paddingWidth
        Write-Host -NoNewline "`r$($lineText.Text)$paddingText`r$($lineText.Text)" -ForegroundColor $Color

        $script:DynamicStatusLastCellWidth = $lineText.CellWidth
        $script:DynamicStatusLastWriteWasInline = $true
        return
    }
    catch {
        # 某些宿主不支持读取控制台宽度；静默降级为普通整行输出，避免异常打断进度行。
        Complete-DynamicStatusLine
        Write-Host $statusMessage -ForegroundColor $Color
        $script:DynamicStatusLastCellWidth = 0
        $script:DynamicStatusLastWriteWasInline = $false
    }
}

# 结束当前动态状态行；只有确实使用了内联刷新时才补换行。
function Complete-DynamicStatusLine {
    if ($script:DynamicStatusLastWriteWasInline) {
        Write-Host ""
    }

    $script:DynamicStatusLastCellWidth = 0
    $script:DynamicStatusLastWriteWasInline = $false
}

# 更新百分比进度条；调用方通过 ref 记录上次百分比，避免重复刷新。
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
    # 大目录扫描时大量重复绘制会明显拖慢输出；百分比变化时再刷新即可。
    if (-not $Force -and $percent -eq $LastPercent.Value) {
        return
    }

    $filledWidth = [Math]::Floor(($percent / 100) * $script:ProgressBarCellCount)
    $emptyWidth = $script:ProgressBarCellCount - $filledWidth
    $bar = ($script:ProgressBarFilledCharacter * $filledWidth) + ($script:ProgressBarEmptyCharacter * $emptyWidth)
    $progressText = "[进度] $Activity [$bar] $percent% $Status ($ProcessedCount / $TotalCount)"

    Write-DynamicStatusLine -Message $progressText -Color Cyan
    $LastPercent.Value = $percent
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

# -yes 会跳过人工确认，因此执行删除前提供醒目的中止窗口。
function Wait-AssumeYesDeletionGracePeriod {
    param(
        [Parameter(Mandatory = $false)]
        [string]$WarningMessage = '危险操作: 已启用 -yes，将跳过详细预览和菜单并执行默认删除。',

        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$AdditionalWarningMessage = '',

        [Parameter(Mandatory = $false)]
        [string]$CancelHintMessage = '如需取消，请在倒计时结束前按 Enter；也可按 Ctrl+C 强制中止。',

        [Parameter(Mandatory = $false)]
        [string]$CancelledMessage = '已取消 -yes 默认删除，未删除任何项目。',

        [Parameter(Mandatory = $false)]
        [string]$CompletedMessage = '倒计时结束，开始执行默认删除。'
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
        Write-DynamicStatusLine -Message "倒计时 $remainingSeconds 秒后开始删除，按 Enter 取消..." -Color Yellow

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
    # 用字符码判断英文单/双引号，避免源文件编码或中文引号造成误判。
    if (($firstCharacterCode -eq 34 -and $lastCharacterCode -eq 34) -or
        ($firstCharacterCode -eq 39 -and $lastCharacterCode -eq 39)) {
        return $normalizedPathText.Substring(1, $normalizedPathText.Length - 2).Trim()
    }

    return $normalizedPathText
}

# 拆分交互输入的路径行：英文引号用于包裹含空格路径；路径可用英文分号分隔，也可在下一个片段是绝对路径时按空格分隔。
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

# 将目录路径标准化为便于比较的形式，用于目录重叠检查。
function ConvertTo-NormalizedDirectoryPath {
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

    return ConvertTo-NormalizedDirectoryPath -Path $DirectoryPath
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
    return @((Get-DirectoryKey -DirectoryPath $DirectoryPath) -split '[\\/]' | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_)
        }).Count
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

    # 先追加尾部分隔符再做 StartsWith，避免 C:\A 被误判为 C:\AB 的父目录。
    return ($leftPrefix.StartsWith($rightPrefix, [System.StringComparison]::OrdinalIgnoreCase) -or
        $rightPrefix.StartsWith($leftPrefix, [System.StringComparison]::OrdinalIgnoreCase))
}

Export-ModuleMember -Function @(
    'Write-StageMessage'
    'Write-MenuItem'
    'Read-ColoredLine'
    'Read-MenuChoice'
    'Get-ConsoleCharacterCellWidth'
    'Get-ConsoleTextWithinCellWidth'
    'Get-ConsoleTextWidth'
    'Write-DynamicStatusLine'
    'Complete-DynamicStatusLine'
    'Write-ProgressBar'
    'Write-RefreshStatusLine'
    'Write-PreviewSeparator'
    'Write-StatusSummary'
    'Test-EnterKeyPressed'
    'Wait-AssumeYesDeletionGracePeriod'
    'New-DeferredScanWarningList'
    'Add-DeferredScanWarning'
    'Write-DeferredScanWarningList'
    'Get-RelativePathText'
    'ConvertTo-UnquotedPathText'
    'Split-InteractivePathInput'
    'ConvertTo-NormalizedDirectoryPath'
    'Get-DirectoryKey'
    'Add-TrailingDirectorySeparator'
    'Get-DirectoryPathDepth'
    'Test-DirectoryOverlap'
)
