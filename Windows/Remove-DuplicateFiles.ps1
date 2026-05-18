#requires -Version 7.0
<#
用途：
  查找文件夹中的重复文件，先计算默认删除计划，再由用户确认默认删除、手动删除或退出。

参数：
  -h     显示帮助信息。
  Path   一个或多个文件夹绝对路径。
  -a     多目录合并模式，把多个目录抽象为一个大目录进行重复文件扫描。
  -c     参考目录模式，第一个目录为参考目录，其余目录为目标目录。
  -s     包含隐藏文件和隐藏文件夹。
  -yes   跳过预览和菜单，等待 10 秒后执行默认删除计划。
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
$PartialHashSegmentByteCount = 256KB

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

# 倒计时状态结束后附加的清理空格数量，用于覆盖上一轮较长输出的尾巴。
$ConsoleLineClearPadding = 20

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

# ========== 输出与路径工具 ==========

function Show-HelpText {
    Write-Host @'
用途：
  查找重复文件，先预览，再选择删除或退出。

用法：
  pwsh -File .\Remove-DuplicateFiles.ps1 [-s] [-yes] [Path1] [Path2 ...]
  pwsh -File .\Remove-DuplicateFiles.ps1 -a [-s] [-yes] <Path1> [Path2 ...]
  pwsh -File .\Remove-DuplicateFiles.ps1 -c [-s] [-yes] <ReferencePath> <TargetPath1> [TargetPath2 ...]
  pwsh -File .\Remove-DuplicateFiles.ps1 -h

参数：
  Path   文件夹绝对路径；不带 -a/-c 时，多个目录先全部扫描，再逐个操作。
  -a     多目录合并模式，把多个目录视作一个大目录。
  -c     参考目录模式；第一个目录为参考目录，其余为目标目录。
  -s     包含隐藏文件和隐藏文件夹。
  -yes   跳过预览和菜单，等待 10 秒后默认删除；可按 Enter 取消。
  -h     显示帮助信息。

说明：
  无路径且未指定 -a/-c 时，会先显示模式菜单。
  单目录和多目录合并模式可默认删除、手动删除或退出。
  多个单目录逐个操作时，0 跳过当前目录，00 退出脚本。
  参考目录模式只删除目标目录文件，可默认删除或退出。
  交互输入多个路径时可分行，也可用空格或英文分号分隔；路径含空格请加引号。
'@
}

# 输出当前执行阶段，避免大目录扫描时长时间无反馈。
function Write-StageMessage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host "[进度] $Message" -ForegroundColor Cyan
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

# 刷新单行动态状态：清理旧尾巴后把光标放回文本末尾，避免补空格导致光标漂移。
function Write-DynamicStatusLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [System.ConsoleColor]$Color
    )

    try {
        if (-not [Console]::IsOutputRedirected) {
            $maxLineWidth = [Math]::Max(1, [Console]::WindowWidth - 1)
            $lineText = Get-ConsoleTextWithinCellWidth -Text $Message -MaxCellWidth $maxLineWidth

            $cursorTop = [Console]::CursorTop
            Write-Host -NoNewline "`r$($lineText.Text)" -ForegroundColor $Color
            $remainingWidth = [Math]::Max(0, $maxLineWidth - $lineText.CellWidth)
            if ($remainingWidth -gt 0) {
                Write-Host -NoNewline (' ' * $remainingWidth)
                [Console]::SetCursorPosition($lineText.CellWidth, $cursorTop)
            }
            return
        }
    }
    catch {
        # 部分宿主不支持读取控制台宽度，回退为普通回车刷新。
        Write-Debug "动态状态行刷新已回退: $($_.Exception.Message)"
    }

    Write-Host -NoNewline "`r$Message" -ForegroundColor $Color
}

# 更新百分比进度条；使用普通文本单行刷新，避免 Write-Progress 改变控制台背景色。
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
        [ref]$LastPercent
    )

    if ($TotalCount -le 0) {
        return
    }

    $percent = [Math]::Min(100, [int][Math]::Floor(($ProcessedCount / $TotalCount) * 100))
    if ($percent -eq $LastPercent.Value) {
        return
    }

    $filledWidth = [Math]::Floor(($percent / 100) * $ProgressBarCellCount)
    $emptyWidth = $ProgressBarCellCount - $filledWidth
    $bar = ($ProgressBarFilledCharacter * $filledWidth) + ($ProgressBarEmptyCharacter * $emptyWidth)
    $progressText = "[进度] $Activity [$bar] $percent% $Status ($ProcessedCount / $TotalCount)"

    Write-DynamicStatusLine -Message $progressText -Color Cyan
    $LastPercent.Value = $percent
}

# 结束当前进度条并换行，避免后续日志和动态进度混在一起。
function Complete-ProgressBar {
    Write-Host ""
}

# 输出用于区分不同预览或结果块的分隔线。
function Write-PreviewSeparator {
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

    Write-Host ""
    Write-Host $Message -ForegroundColor $Color
}

# 检查倒计时期间是否按下 Enter；不支持读取键盘状态时静默退化为只支持 Ctrl+C。
function Test-EnterKeyPressed {
    try {
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

    Write-Host ""
    Write-Host "危险操作: 已启用 -yes，将跳过预览和菜单并执行默认删除。" -ForegroundColor Red
    Write-Host "如需取消，请在倒计时结束前按 Enter；也可按 Ctrl+C 强制中止。" -ForegroundColor Yellow

    for ($remainingSeconds = $Seconds; $remainingSeconds -gt 0; $remainingSeconds--) {
        Write-Host -NoNewline "`r倒计时 $remainingSeconds 秒后开始删除，按 Enter 取消..." -ForegroundColor Yellow

        $pollCountPerSecond = [Math]::Max(1, [int][Math]::Ceiling(1000 / $AssumeYesInputPollIntervalMilliseconds))
        for ($pollIndex = 0; $pollIndex -lt $pollCountPerSecond; $pollIndex++) {
            if (Test-EnterKeyPressed) {
                $script:AssumeYesDeletionCancelled = $true
                Write-Host "`r已取消 -yes 默认删除，未删除任何文件。$(' ' * $ConsoleLineClearPadding)" -ForegroundColor Yellow
                return $false
            }

            Start-Sleep -Milliseconds $AssumeYesInputPollIntervalMilliseconds
        }
    }

    Write-Host "`r倒计时结束，开始执行默认删除。$(' ' * $ConsoleLineClearPadding)" -ForegroundColor Magenta
    return $true
}

# 兼容交互输入时复制带首尾引号的路径；这里只移除成对包裹符号。
function ConvertTo-UnquotedPathText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathText
    )

    $normalizedPathText = $PathText.Trim()
    if ($normalizedPathText.Length -lt 2) {
        return $normalizedPathText
    }

    $quotePairs = @(
        @{ Open = '"'; Close = '"' }
        @{ Open = "'"; Close = "'" }
        @{ Open = '“'; Close = '”' }
        @{ Open = '‘'; Close = '’' }
    )

    foreach ($quotePair in $quotePairs) {
        if ($normalizedPathText.StartsWith($quotePair.Open, [System.StringComparison]::Ordinal) -and
            $normalizedPathText.EndsWith($quotePair.Close, [System.StringComparison]::Ordinal)) {
            return $normalizedPathText.Substring(1, $normalizedPathText.Length - 2).Trim()
        }
    }

    return $normalizedPathText
}

# 拆分交互输入的路径行：支持分号分隔，也支持在下一个片段看起来是绝对路径时按空格分隔。
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
    $currentInputPart = [System.Text.StringBuilder]::new()
    $quoteCloseByOpen = @{}
    $quoteCloseByOpen.Add([string][char]34, [string][char]34)
    $quoteCloseByOpen.Add([string][char]39, [string][char]39)
    $quoteCloseByOpen.Add([string][char]0x201C, [string][char]0x201D)
    $quoteCloseByOpen.Add([string][char]0x2018, [string][char]0x2019)
    $activeClosingQuote = $null

    for ($index = 0; $index -lt $trimmedPathInput.Length; $index++) {
        $currentChar = $trimmedPathInput[$index].ToString()

        if ($null -ne $activeClosingQuote) {
            [void]$currentInputPart.Append($currentChar)
            if ($currentChar -eq $activeClosingQuote) {
                $activeClosingQuote = $null
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
            [void]$currentInputPart.Append($currentChar)
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

    $resultPathList = [System.Collections.Generic.List[string]]::new()
    $quoteStartPattern = @(
        [regex]::Escape([string][char]34)
        [regex]::Escape([string][char]39)
        [regex]::Escape([string][char]0x201C)
        [regex]::Escape([string][char]0x2018)
    ) -join '|'
    $absolutePathSeparatorPattern = '\s+(?=(?:' + $quoteStartPattern + ')?(?:[a-zA-Z]:[\\/]|[\\/]{2}))'
    foreach ($pathPart in $pathPartList) {
        foreach ($pathSegment in [regex]::Split($pathPart, $absolutePathSeparatorPattern)) {
            $trimmedPathSegment = $pathSegment.Trim()
            if (-not [string]::IsNullOrWhiteSpace($trimmedPathSegment)) {
                $resultPathList.Add($trimmedPathSegment)
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

    $isAbsolutePath = $normalizedPathText -match '^[a-zA-Z]:[\\/]' -or $normalizedPathText -match '^[\\/]{2}'
    if (-not $isAbsolutePath) {
        throw "$ParameterName 必须是 Windows 文件夹绝对路径: $normalizedPathText"
    }

    try {
        $resolvedPaths = @(Resolve-Path -LiteralPath $normalizedPathText -ErrorAction Stop)
    }
    catch {
        throw "$ParameterName 不存在或无法访问: $normalizedPathText。请确认路径存在；多个路径可分行输入，或在同一行用空格/英文分号分隔；路径含空格请加引号。原始错误: $($_.Exception.Message)"
    }

    if ($resolvedPaths.Count -ne 1) {
        throw "$ParameterName 必须只能解析到一个目录。"
    }

    try {
        $item = Get-Item -LiteralPath $resolvedPaths[0].ProviderPath -ErrorAction Stop
    }
    catch {
        throw "$ParameterName 无法读取: $normalizedPathText。原因: $($_.Exception.Message)"
    }

    if (-not $item.PSIsContainer) {
        throw "$ParameterName 必须是文件夹: $normalizedPathText"
    }

    return $item.FullName
}

# 逐个校验路径列表，并返回规范化后的完整目录路径。
function Resolve-InputDirectoryList {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$PathList,

        [Parameter(Mandatory = $false)]
        [string]$ParameterNamePrefix = 'Path'
    )

    return @(
        for ($index = 0; $index -lt $PathList.Count; $index++) {
            Resolve-InputDirectory -Path $PathList[$index] -ParameterName "$($ParameterNamePrefix)[$($index + 1)]"
        }
    )
}

# 无参数执行时按行读取一个或多个目录路径；已输入至少一个路径后，空回车开始执行。
function Read-InteractivePathList {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModeName,

        [Parameter(Mandatory = $true)]
        [int]$MinimumCount
    )

    $inputPathList = [System.Collections.Generic.List[string]]::new()
    Write-Host "进入$ModeName。" -ForegroundColor Cyan
    Write-Host "请输入目录绝对路径。可在同一行输入多个路径。" -ForegroundColor Yellow
    Write-Host "多个路径可用空格或英文分号分隔；路径含空格请加引号；直接回车开始执行或退出。" -ForegroundColor DarkGray

    while ($true) {
        $pathInput = (Read-Host "Path$($inputPathList.Count + 1)").Trim()
        if ([string]::IsNullOrWhiteSpace($pathInput)) {
            if ($inputPathList.Count -eq 0) {
                Write-Host "已退出，未执行扫描。" -ForegroundColor Yellow
                exit 0
            }

            if ($inputPathList.Count -lt $MinimumCount) {
                Write-Host "$ModeName 至少需要输入 $MinimumCount 个目录。" -ForegroundColor Red
                continue
            }

            return $inputPathList.ToArray()
        }

        $pathInputList = @(Split-InteractivePathInput -PathInput $pathInput)
        foreach ($inputPath in $pathInputList) {
            $inputPathList.Add($inputPath)
        }

        if ($pathInputList.Count -gt 1) {
            Write-Host "已识别 $($pathInputList.Count) 个路径。" -ForegroundColor DarkGray
        }
    }
}

# 将目录路径标准化为便于比较的形式，用于目录重叠检查。
function ConvertTo-NormalizedDirectoryPath {
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

# 合并或参考模式必须使用互不包含的目录，避免同一文件被重复扫描后匹配到自身。
function Test-IndependentDirectoryPair {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LeftRootPath,

        [Parameter(Mandatory = $true)]
        [string]$RightRootPath,

        [Parameter(Mandatory = $false)]
        [string]$LeftName = 'Path1',

        [Parameter(Mandatory = $false)]
        [string]$RightName = 'Path2'
    )

    $normalizedLeftPath = ConvertTo-NormalizedDirectoryPath -Path $LeftRootPath
    $normalizedRightPath = ConvertTo-NormalizedDirectoryPath -Path $RightRootPath

    if ($normalizedLeftPath.Equals($normalizedRightPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "$LeftName 和 $RightName 不能是同一个目录。"
    }

    $leftPathPrefix = Add-TrailingDirectorySeparator -Path $normalizedLeftPath
    $rightPathPrefix = Add-TrailingDirectorySeparator -Path $normalizedRightPath

    if ($normalizedRightPath.StartsWith($leftPathPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "$RightName 不能位于 $LeftName 子目录中。"
    }

    if ($normalizedLeftPath.StartsWith($rightPathPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "$LeftName 不能位于 $RightName 子目录中。"
    }
}

# 单目录模式允许父子目录分别处理，但不允许同一目录重复输入。
function Test-UniqueDirectorySet {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$RootPathList
    )

    for ($leftIndex = 0; $leftIndex -lt $RootPathList.Count; $leftIndex++) {
        $normalizedLeftPath = ConvertTo-NormalizedDirectoryPath -Path $RootPathList[$leftIndex]
        for ($rightIndex = $leftIndex + 1; $rightIndex -lt $RootPathList.Count; $rightIndex++) {
            $normalizedRightPath = ConvertTo-NormalizedDirectoryPath -Path $RootPathList[$rightIndex]
            if ($normalizedLeftPath.Equals($normalizedRightPath, [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "Path$($leftIndex + 1) 和 Path$($rightIndex + 1) 不能是同一个目录。"
            }
        }
    }
}

# 校验一组目录两两独立；用于多目录合并和多目标参考模式。
function Test-IndependentDirectorySet {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$RootPathList
    )

    for ($leftIndex = 0; $leftIndex -lt $RootPathList.Count; $leftIndex++) {
        for ($rightIndex = $leftIndex + 1; $rightIndex -lt $RootPathList.Count; $rightIndex++) {
            Test-IndependentDirectoryPair `
                -LeftRootPath $RootPathList[$leftIndex] `
                -RightRootPath $RootPathList[$rightIndex] `
                -LeftName "Path$($leftIndex + 1)" `
                -RightName "Path$($rightIndex + 1)"
        }
    }
}

# ========== 哈希与重复文件识别 ==========

# 计算文件首尾片段的 SHA-256，用作快速筛选候选重复文件。
function Get-PartialContentHash {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    $fileStream = [System.IO.File]::Open($File.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)

    try {
        # 这里只做快速预筛选；真正删除前仍会使用完整 SHA-256 确认。
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

# 递归获取目录下所有普通文件；默认不包含隐藏项，传入 -s 时包含隐藏文件和隐藏文件夹。
function Get-ScannedFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath,

        [Parameter(Mandatory = $false)]
        [string]$ProgressLabel = '目录'
    )

    Write-StageMessage "开始扫描$($ProgressLabel): $RootPath"
    $scanErrorList = $null
    if ($ShouldIncludeHiddenItems) {
        $scannedFiles = @(Get-ChildItem -LiteralPath $RootPath -File -Recurse -Force -ErrorAction SilentlyContinue -ErrorVariable scanErrorList)
    }
    else {
        $scannedFiles = @(Get-ChildItem -LiteralPath $RootPath -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable scanErrorList)
    }

    foreach ($scanError in @($scanErrorList)) {
        Write-Host "扫描跳过: $($scanError.TargetObject)" -ForegroundColor Yellow
        Write-Host "  原因: $($scanError.Exception.Message)" -ForegroundColor DarkGray
    }

    $hiddenScopeText = if ($ShouldIncludeHiddenItems) { '包含隐藏项' } else { '不包含隐藏项' }
    Write-StageMessage "$($ProgressLabel)扫描完成，文件数: $($scannedFiles.Count)，$hiddenScopeText"
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

    $relativePathText = [System.IO.Path]::GetRelativePath($RootPath, $File.FullName)
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
    return [string]::IsNullOrWhiteSpace($name) ? $RootPath.TrimEnd('\') : $name
}

# 获取文件所在目录的规范化路径，用于默认保留优先级判断。
function Get-FileParentDirectoryPath {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    return ConvertTo-NormalizedDirectoryPath -Path (Split-Path -Parent $File.FullName)
}

# 统计目录路径层级；层级越少，默认保留优先级越高。
function Get-DirectoryPathDepth {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return @($Path -split '[\\/]' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count
}

# 按默认保留规则排序文件：先比较所在目录，再用文件名作为同目录兜底。
function Get-FilesByKeepPriority {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$FileList
    )

    return $FileList |
        Sort-Object @{ Expression = { Get-DirectoryPathDepth -Path (Get-FileParentDirectoryPath -File $_) }; Ascending = $true },
                    @{ Expression = { (Get-FileParentDirectoryPath -File $_).Length }; Ascending = $true },
                    @{ Expression = { $_.Name.Length }; Ascending = $true },
                    @{ Expression = { $_.Name }; Ascending = $true },
                    @{ Expression = { $_.FullName }; Ascending = $true }
}

# 从一组重复文件中选出默认应保留的文件。
function Select-DefaultKeepFile {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$FileList
    )

    return Get-FilesByKeepPriority -FileList $FileList | Select-Object -First 1
}

# 按单目录根路径或预先建立的映射获取文件显示路径。
function Get-FileDisplayPath {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File,

        [Parameter(Mandatory = $false)]
        [string]$RootPath,

        [Parameter(Mandatory = $false)]
        [hashtable]$DisplayPathByFullName
    )

    if ($null -ne $DisplayPathByFullName -and $DisplayPathByFullName.ContainsKey($File.FullName)) {
        return $DisplayPathByFullName[$File.FullName]
    }

    if (-not [string]::IsNullOrWhiteSpace($RootPath)) {
        return Get-RelativePathText -File $File -RootPath $RootPath
    }

    return $File.FullName
}

# 将待删除文件封装为包含文件对象和显示路径的删除项。
function New-DeletionItems {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$FileList,

        [Parameter(Mandatory = $false)]
        [string]$RootPath,

        [Parameter(Mandatory = $false)]
        [string]$PathPrefix,

        [Parameter(Mandatory = $false)]
        [hashtable]$DisplayPathByFullName
    )

    return @(
        $FileList | ForEach-Object {
            $displayPath = if ($null -ne $DisplayPathByFullName -and $DisplayPathByFullName.ContainsKey($_.FullName)) {
                $DisplayPathByFullName[$_.FullName]
            }
            else {
                if ([string]::IsNullOrWhiteSpace($RootPath)) {
                    $_.FullName
                }
                else {
                    Get-RelativePathText -File $_ -RootPath $RootPath -PathPrefix $PathPrefix
                }
            }

            [pscustomobject]@{
                File        = $_
                DisplayPath = $displayPath
            }
        }
    )
}

# 按文件大小、部分哈希、完整哈希分层筛选出内容完全一致的重复文件组。
function Find-DuplicateFileGroups {
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

        $filesByLength[$file.Length] ??= [System.Collections.Generic.List[System.IO.FileInfo]]::new()
        $filesByLength[$file.Length].Add($file)
    }

    Complete-ProgressBar

    # 只有大小相同的文件才可能重复；不同大小的文件无需继续计算哈希。
    $sameLengthGroups = @(
        foreach ($size in $filesByLength.Keys) {
            if ($filesByLength[$size].Count -gt 1) {
                [pscustomobject]@{
                    Name  = $size
                    Count = $filesByLength[$size].Count
                    Group = @($filesByLength[$size])
                }
            }
        }
    )
    $partialHashCandidateCount = @($sameLengthGroups | ForEach-Object { $_.Group }).Count
    Write-StageMessage "$($ProgressLabel)大小相同的候选文件数: $partialHashCandidateCount，候选大小组数: $($sameLengthGroups.Count)"

    $processedPartialHashCount = 0
    $lastPartialHashPercent = -1
    $hashProgressName = "$($ProgressLabel)哈希计算"

    foreach ($sizeGroup in $sameLengthGroups) {
        $partialHashRecords = @(
            foreach ($file in $sizeGroup.Group) {
                $processedPartialHashCount++
                Write-ProgressBar -Activity $hashProgressName -Status '正在筛选候选文件' -ProcessedCount $processedPartialHashCount -TotalCount $partialHashCandidateCount -LastPercent ([ref]$lastPartialHashPercent)

                try {
                    [pscustomobject]@{
                        File        = $file
                        PartialHash = Get-PartialContentHash -File $file
                    }
                }
                catch {
                    Write-Host "跳过文件，无法计算部分哈希: $($file.FullName)" -ForegroundColor Yellow
                    Write-Host "  原因: $($_.Exception.Message)" -ForegroundColor DarkGray
                }
            }
        )

        $partialHashGroups = @($partialHashRecords |
            Group-Object -Property PartialHash |
            Where-Object { $_.Count -gt 1 })

        foreach ($partialHashGroup in $partialHashGroups) {
            # 部分哈希只用于减少候选范围，最终仍按完整 SHA-256 分组确认。
            $contentHashRecords = @(
                foreach ($partialHashRecord in $partialHashGroup.Group) {
                    try {
                        [pscustomobject]@{
                            File        = $partialHashRecord.File
                            ContentHash = Get-FullContentHash -File $partialHashRecord.File
                        }
                    }
                    catch {
                        Write-Host "跳过文件，无法计算完整哈希: $($partialHashRecord.File.FullName)" -ForegroundColor Yellow
                        Write-Host "  原因: $($_.Exception.Message)" -ForegroundColor DarkGray
                    }
                }
            )

            $contentHashGroups = @($contentHashRecords |
                Group-Object -Property ContentHash |
                Where-Object { $_.Count -gt 1 })

            foreach ($contentHashGroup in $contentHashGroups) {
                [pscustomobject]@{
                    Hash  = $contentHashGroup.Name
                    Files = @($contentHashGroup.Group | ForEach-Object { $_.File })
                }
            }
        }
    }

    Complete-ProgressBar
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

# 输出所有默认删除计划的预览，并返回预计删除数量。
function Write-DeletionPlanPreview {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$DeletionPlanList,

        [Parameter(Mandatory = $true)]
        [string]$SummaryFormat
    )

    Write-Host "`n删除预览:" -ForegroundColor Yellow
    $plannedDeletionCount = 0
    $plannedKeepCount = 0
    foreach ($deletionPlan in $DeletionPlanList) {
        $plannedDeletionItems = @($deletionPlan.DeletionItems)
        $plannedDeletePathTexts = @($plannedDeletionItems | ForEach-Object { $_.DisplayPath })
        Write-DuplicatePreviewBlock -Hash $deletionPlan.Hash -KeepPathText $deletionPlan.KeepPathText -DeletePathTexts $plannedDeletePathTexts
        $plannedDeletionCount += $plannedDeletionItems.Count
        $plannedKeepCount++
    }

    Write-Host "重复组数: $($DeletionPlanList.Count)，默认保留文件数: $plannedKeepCount，默认计划删除文件数: $plannedDeletionCount" -ForegroundColor DarkGray
    Write-StatusSummary -Message ($SummaryFormat -f $plannedDeletionCount) -Color Yellow
    return $plannedDeletionCount
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

    $orderedDuplicateFiles = @(Get-FilesByKeepPriority -FileList $DuplicateFileGroup)
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
        $manualInputText = Read-Host "请输入要删除的编号，多个编号用逗号分隔；直接回车使用默认规则；输入 0 跳过；输入 00 退出脚本"
        $trimmedInputText = $manualInputText.Trim()

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

        $selectedFileNumbers = [System.Collections.Generic.List[int]]::new()
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

# 执行一批删除项，并返回实际删除与失败数量。
function Remove-DeletionItems {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$DeletionItemList,

        [Parameter(Mandatory = $false)]
        [switch]$Quiet
    )

    $deletedFileCount = 0
    $failedFileCount = 0
    $deletedItemList = [System.Collections.Generic.List[object]]::new()
    foreach ($deletionItem in $DeletionItemList) {
        try {
            Remove-Item -LiteralPath $deletionItem.File.FullName -Force -ErrorAction Stop
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
        [string]$SummaryFormat
    )

    $deletionResult = Remove-DeletionItems -DeletionItemList @($DeletionPlanList | ForEach-Object { $_.DeletionItems })
    Write-StatusSummary -Message ($SummaryFormat -f $deletionResult.DeletedCount) -Color Magenta
    if ($deletionResult.FailedCount -gt 0) {
        Write-Host "删除失败文件: $($deletionResult.FailedCount)" -ForegroundColor Red
    }

    return $deletionResult
}

# 生成删除操作菜单选项；多单目录操作时可把 0 作为跳过当前目录。
function New-DeletionActionMenuOptions {
    param(
        [Parameter(Mandatory = $false)]
        [switch]$IncludeManualDeletion,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeSkipCurrentDirectory,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeExitScript
    )

    $menuOptionList = [System.Collections.Generic.List[object]]::new()
    $menuOptionList.Add([pscustomobject]@{ Value = '1'; Label = '默认删除' })
    if ($IncludeManualDeletion) {
        $menuOptionList.Add([pscustomobject]@{ Value = '2'; Label = '手动删除' })
    }
    if ($IncludeSkipCurrentDirectory) {
        $menuOptionList.Add([pscustomobject]@{ Value = '0'; Label = '跳过当前目录' })
    }
    else {
        $menuOptionList.Add([pscustomobject]@{ Value = '0'; Label = '退出' })
    }
    if ($IncludeExitScript) {
        $menuOptionList.Add([pscustomobject]@{ Value = '00'; Label = '退出脚本' })
    }

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

    Write-Host ""
    Write-Host $Title -ForegroundColor Cyan
    foreach ($menuOption in $MenuOptionList) {
        Write-Host -NoNewline "  $($menuOption.Value) " -ForegroundColor Cyan
        Write-Host $menuOption.Label
    }

    $validMenuChoices = @($MenuOptionList | ForEach-Object { $_.Value })
    while ($true) {
        $menuChoice = (Read-Host "请输入选项").Trim()
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
        [pscustomobject]@{ Value = '1'; Label = '单目录模式（多个目录先扫描后操作）' }
        [pscustomobject]@{ Value = '2'; Label = '多目录合并模式（多个目录视作一个大目录）' }
        [pscustomobject]@{ Value = '3'; Label = '参考目录模式（第一个目录为参考目录）' }
        [pscustomobject]@{ Value = '0'; Label = '退出脚本' }
    )

    switch ($modeChoice) {
        '1' { return 'Single' }
        '2' { return 'Aggregate' }
        '3' { return 'Reference' }
        '0' { return 'Exit' }
    }
}

# 为单目录模式生成默认删除计划。
function New-SingleDirectoryDeletionPlan {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    $scannedFiles = @(Get-ScannedFiles -RootPath $RootPath -ProgressLabel '单目录')
    if ($scannedFiles.Count -lt 2) {
        return @()
    }

    foreach ($duplicateGroupRecord in Find-DuplicateFileGroups -FileList $scannedFiles -ProgressLabel '单目录') {
        $duplicateFiles = @($duplicateGroupRecord.Files)
        $defaultKeepFile = Select-DefaultKeepFile -FileList $duplicateFiles
        $filesToDelete = @($duplicateFiles | Where-Object { $_.FullName -ne $defaultKeepFile.FullName })

        [pscustomobject]@{
            Hash           = $duplicateGroupRecord.Hash
            KeepFile      = $defaultKeepFile
            KeepPathText  = Get-RelativePathText -File $defaultKeepFile -RootPath $RootPath
            DeletionItems = New-DeletionItems -FileList $filesToDelete -RootPath $RootPath
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

    $fileRecordList = [System.Collections.Generic.List[object]]::new()
    for ($rootIndex = 0; $rootIndex -lt $RootPathList.Count; $rootIndex++) {
        $rootPath = $RootPathList[$rootIndex]
        $pathPrefix = "{0}-{1}" -f ($rootIndex + 1), (Get-DirectoryLabel -RootPath $rootPath)
        $scannedFiles = @(Get-ScannedFiles -RootPath $rootPath -ProgressLabel "合并目录$($rootIndex + 1)")

        foreach ($file in $scannedFiles) {
            $fileRecordList.Add([pscustomobject]@{
                File        = $file
                DisplayPath = Get-RelativePathText -File $file -RootPath $rootPath -PathPrefix $pathPrefix
            })
        }
    }

    $scannedFileList = @($fileRecordList | ForEach-Object { $_.File })
    if ($scannedFileList.Count -lt 2) {
        return @()
    }

    $displayPathByFullName = @{}
    foreach ($fileRecord in $fileRecordList) {
        $displayPathByFullName[$fileRecord.File.FullName] = $fileRecord.DisplayPath
    }

    foreach ($duplicateGroupRecord in Find-DuplicateFileGroups -FileList $scannedFileList -ProgressLabel '多目录合并') {
        $duplicateFiles = @($duplicateGroupRecord.Files)
        $defaultKeepFile = Select-DefaultKeepFile -FileList $duplicateFiles
        $filesToDelete = @($duplicateFiles | Where-Object { $_.FullName -ne $defaultKeepFile.FullName })

        [pscustomobject]@{
            Hash           = $duplicateGroupRecord.Hash
            KeepFile       = $defaultKeepFile
            KeepPathText   = $displayPathByFullName[$defaultKeepFile.FullName]
            DeletionItems  = New-DeletionItems -FileList $filesToDelete -DisplayPathByFullName $displayPathByFullName
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

        $manualDeletionItems = New-DeletionItems -FileList $selectedFilesToDelete -RootPath $RootPath -DisplayPathByFullName $DisplayPathByFullName

        $deletionResult = Remove-DeletionItems -DeletionItemList $manualDeletionItems -Quiet
        $deletedSelectionList = @(
            foreach ($selectedDeletionEntry in $selectedDeletionEntries) {
                if (@($deletionResult.DeletedItems | Where-Object { $_.File.FullName -eq $selectedDeletionEntry.File.FullName }).Count -gt 0) {
                    $selectedDeletionEntry
                }
            }
        )
        if ($deletedSelectionList.Count -gt 0) {
            Write-ManualDeletionResult -DeletedSelectionList $deletedSelectionList -RootPath $RootPath -DisplayPathByFullName $DisplayPathByFullName
        }
        $deletedFileCount += $deletionResult.DeletedCount
        $failedFileCount += $deletionResult.FailedCount
    }

    if ($deletedFileCount -eq 0) {
        Write-Host "未选择删除任何文件。" -ForegroundColor Yellow
    }
    else {
        Write-StatusSummary -Message "手动删除完成。已删除重复文件: $deletedFileCount" -Color Magenta
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

# 对已经生成的删除计划执行预览、默认删除、手动删除、跳过或退出。
function Invoke-DeletionPlanAction {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$DeletionPlanList,

        [Parameter(Mandatory = $true)]
        [string]$EmptyMessage,

        [Parameter(Mandatory = $true)]
        [string]$PreviewSummaryFormat,

        [Parameter(Mandatory = $true)]
        [string]$DefaultDeletionSummaryFormat,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeManualDeletion,

        [Parameter(Mandatory = $false)]
        [switch]$AllowSkipCurrentDirectory,

        [Parameter(Mandatory = $false)]
        [string]$ManualRootPath,

        [Parameter(Mandatory = $false)]
        [hashtable]$ManualDisplayPathByFullName
    )

    if ($DeletionPlanList.Count -eq 0) {
        Write-Host $EmptyMessage -ForegroundColor Green
        return 'Continue'
    }

    if ($ShouldAssumeYesDeletion) {
        if (-not (Wait-AssumeYesDeletionGracePeriod)) {
            return 'Exit'
        }

        [void](Invoke-DefaultDeletionPlan -DeletionPlanList $DeletionPlanList -SummaryFormat $DefaultDeletionSummaryFormat)
        return 'Continue'
    }

    [void](Write-DeletionPlanPreview -DeletionPlanList $DeletionPlanList -SummaryFormat $PreviewSummaryFormat)
    $menuOptionList = New-DeletionActionMenuOptions -IncludeManualDeletion:$IncludeManualDeletion -IncludeSkipCurrentDirectory:$AllowSkipCurrentDirectory -IncludeExitScript:$AllowSkipCurrentDirectory
    $menuChoice = Read-DeletionAction -MenuOptionList $menuOptionList

    if ($menuChoice -eq '0') {
        if ($AllowSkipCurrentDirectory) {
            Write-Host "已跳过当前目录，未删除任何文件。" -ForegroundColor Yellow
            return 'Skip'
        }

        Write-Host "已退出脚本，未删除任何文件。" -ForegroundColor Yellow
        return 'Exit'
    }

    if ($menuChoice -eq '00') {
        Write-Host "已退出脚本，未删除任何文件。" -ForegroundColor Yellow
        return 'Exit'
    }

    if ($menuChoice -eq '1') {
        [void](Invoke-DefaultDeletionPlan -DeletionPlanList $DeletionPlanList -SummaryFormat $DefaultDeletionSummaryFormat)
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
        DeletionPlanList             = $DeletionPlanList
        EmptyMessage                 = '未发现重复文件。'
        PreviewSummaryFormat         = '重复文件列举完成。默认计划删除重复文件: {0}'
        DefaultDeletionSummaryFormat = '删除完成。已删除重复文件: {0}'
        IncludeManualDeletion        = $true
        AllowSkipCurrentDirectory    = $AllowSkipCurrentDirectory
        ManualRootPath               = $RootPath
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
    [void](Invoke-SingleDirectoryDeletionAction -RootPath $RootPath -DeletionPlanList $deletionPlanList)
}

# 多个单目录先全部扫描，再按目录逐个预览和确认，避免扫描过程中被菜单反复打断。
function Invoke-IndependentSingleDirectoryMode {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$RootPathList
    )

    if ($RootPathList.Count -eq 1) {
        Invoke-SingleDirectoryMode -RootPath $RootPathList[0]
        return
    }

    $singleDirectoryPlanRecordList = [System.Collections.Generic.List[object]]::new()
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
            RootPath         = $RootPathList[$rootIndex]
            DeletionPlanList = $deletionPlanList
        })
    }

    $actionPlanRecordList = @($singleDirectoryPlanRecordList | Where-Object { $_.DeletionPlanList.Count -gt 0 })
    $actionDirectoryCount = $actionPlanRecordList.Count
    Write-StatusSummary -Message "单目录模式扫描完成。待操作目录: $actionDirectoryCount / $($RootPathList.Count)" -Color Cyan
    if ($actionDirectoryCount -eq 0) {
        return
    }

    for ($recordIndex = 0; $recordIndex -lt $actionPlanRecordList.Count; $recordIndex++) {
        $planRecord = $actionPlanRecordList[$recordIndex]
        Write-PreviewSeparator
        Write-Host "单目录模式操作 $($recordIndex + 1) / $($actionPlanRecordList.Count): $($planRecord.RootPath)" -ForegroundColor Cyan

        $operationResult = Invoke-SingleDirectoryDeletionAction -RootPath $planRecord.RootPath -DeletionPlanList @($planRecord.DeletionPlanList) -AllowSkipCurrentDirectory
        if ($operationResult -eq 'Exit' -or $AssumeYesDeletionCancelled) {
            return
        }
    }
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
        DeletionPlanList            = $deletionPlanList
        EmptyMessage                = '未发现重复文件。'
        PreviewSummaryFormat        = '重复文件列举完成。默认计划从多目录合并结果中删除重复文件: {0}'
        DefaultDeletionSummaryFormat = '删除完成。已从多目录合并结果中删除重复文件: {0}'
        IncludeManualDeletion       = $true
        ManualDisplayPathByFullName = $displayPathByFullName
    }
    [void](Invoke-DeletionPlanAction @actionParameters)
}

# 为参考目录模式建立轻量参考索引；这里只按文件大小分组，不读取文件内容。
function New-ReferenceDirectoryIndex {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ReferenceRootPath
    )

    $referenceFileList = @(Get-ScannedFiles -RootPath $ReferenceRootPath -ProgressLabel '参考目录')
    $referenceFilesByLength = @{}
    $indexedReferenceFileCount = 0

    if ($referenceFileList.Count -gt 0) {
        Write-StageMessage "参考目录模式构建文件大小索引..."
        $processedReferenceFileCount = 0
        $lastReferenceLengthPercent = -1
        foreach ($file in $referenceFileList) {
            $processedReferenceFileCount++
            Write-ProgressBar -Activity '参考目录大小索引' -Status '正在按文件大小归类' -ProcessedCount $processedReferenceFileCount -TotalCount $referenceFileList.Count -LastPercent ([ref]$lastReferenceLengthPercent)

            $referenceFilesByLength[$file.Length] ??= [System.Collections.Generic.List[System.IO.FileInfo]]::new()
            $referenceFilesByLength[$file.Length].Add($file)
            $indexedReferenceFileCount++
        }
        Complete-ProgressBar
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

# 按参考文件完整路径缓存部分哈希，避免同一次运行中同一文件重复计算。
function Get-CachedReferencePartialHash {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ReferenceIndex,

        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    if ($ReferenceIndex.PartialHashByFullName.ContainsKey($File.FullName)) {
        return $ReferenceIndex.PartialHashByFullName[$File.FullName]
    }

    $partialHash = Get-PartialContentHash -File $File
    $ReferenceIndex.PartialHashByFullName[$File.FullName] = $partialHash
    return $partialHash
}

# 按参考文件完整路径缓存完整哈希，避免同一次运行中同一文件重复计算。
function Get-CachedReferenceFullHash {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ReferenceIndex,

        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    if ($ReferenceIndex.FullHashByFullName.ContainsKey($File.FullName)) {
        return $ReferenceIndex.FullHashByFullName[$File.FullName]
    }

    $fullHash = Get-FullContentHash -File $File
    $ReferenceIndex.FullHashByFullName[$File.FullName] = $fullHash
    return $fullHash
}

# 按需为参考目录中的某个文件大小组建立部分哈希索引，并缓存结果供后续目标目录复用。
function Get-ReferencePartialHashIndex {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ReferenceIndex,

        [Parameter(Mandatory = $true)]
        [long]$Length
    )

    if ($ReferenceIndex.PartialHashIndexByLength.ContainsKey($Length)) {
        return $ReferenceIndex.PartialHashIndexByLength[$Length]
    }

    $partialHashIndex = @{}
    if (-not $ReferenceIndex.FilesByLength.ContainsKey($Length)) {
        $ReferenceIndex.PartialHashIndexByLength[$Length] = $partialHashIndex
        return $partialHashIndex
    }

    $referenceFiles = @($ReferenceIndex.FilesByLength[$Length])
    foreach ($referenceFile in $referenceFiles) {
        try {
            $partialHash = Get-CachedReferencePartialHash -ReferenceIndex $ReferenceIndex -File $referenceFile
            $partialHashIndex[$partialHash] ??= [System.Collections.Generic.List[System.IO.FileInfo]]::new()
            $partialHashIndex[$partialHash].Add($referenceFile)
        }
        catch {
            Write-Host "跳过参考文件，无法计算部分哈希: $($referenceFile.FullName)" -ForegroundColor Yellow
            Write-Host "  原因: $($_.Exception.Message)" -ForegroundColor DarkGray
        }
    }

    $ReferenceIndex.PartialHashIndexByLength[$Length] = $partialHashIndex
    return $partialHashIndex
}

# 按需为参考目录中的某个“文件大小 + 部分哈希”组建立完整哈希索引，并缓存结果供后续目标目录复用。
function Get-ReferenceFullHashIndex {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ReferenceIndex,

        [Parameter(Mandatory = $true)]
        [long]$Length,

        [Parameter(Mandatory = $true)]
        [string]$PartialHash
    )

    if (-not $ReferenceIndex.FullHashIndexCache.ContainsKey($Length)) {
        $ReferenceIndex.FullHashIndexCache[$Length] = @{}
    }

    $fullHashCacheByPartialHash = $ReferenceIndex.FullHashIndexCache[$Length]
    if ($fullHashCacheByPartialHash.ContainsKey($PartialHash)) {
        return $fullHashCacheByPartialHash[$PartialHash]
    }

    $partialHashIndex = Get-ReferencePartialHashIndex -ReferenceIndex $ReferenceIndex -Length $Length
    $fullHashIndex = @{}
    if (-not $partialHashIndex.ContainsKey($PartialHash)) {
        $fullHashCacheByPartialHash[$PartialHash] = $fullHashIndex
        return $fullHashIndex
    }

    $referenceFiles = @($partialHashIndex[$PartialHash])
    foreach ($referenceFile in $referenceFiles) {
        try {
            $fullHash = Get-CachedReferenceFullHash -ReferenceIndex $ReferenceIndex -File $referenceFile
            $fullHashIndex[$fullHash] ??= [System.Collections.Generic.List[System.IO.FileInfo]]::new()
            $fullHashIndex[$fullHash].Add($referenceFile)
        }
        catch {
            Write-Host "跳过参考文件，无法计算完整哈希: $($referenceFile.FullName)" -ForegroundColor Yellow
            Write-Host "  原因: $($_.Exception.Message)" -ForegroundColor DarkGray
        }
    }

    $fullHashCacheByPartialHash[$PartialHash] = $fullHashIndex
    return $fullHashIndex
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
    $targetFileList = @(Get-ScannedFiles -RootPath $TargetRootPath -ProgressLabel '目标目录')
    if ($referenceFileList.Count -eq 0 -or $ReferenceIndex.IndexedFileCount -eq 0 -or $targetFileList.Count -eq 0) {
        return @()
    }

    $referenceRootPath = $ReferenceIndex.RootPath
    $referencePathPrefix = $ReferenceIndex.PathPrefix
    $referenceFilesByLength = $ReferenceIndex.FilesByLength
    $targetPathPrefix = Get-DirectoryLabel -RootPath $TargetRootPath

    Write-StageMessage "参考目录模式使用懒加载索引筛选目标目录..."
    $matchedTargetFilesByHash = @{}
    $processedTargetFileCount = 0
    $lastTargetMatchPercent = -1
    foreach ($file in $targetFileList) {
        $processedTargetFileCount++
        Write-ProgressBar -Activity '目标目录重复文件筛选' -Status '正在匹配参考目录索引' -ProcessedCount $processedTargetFileCount -TotalCount $targetFileList.Count -LastPercent ([ref]$lastTargetMatchPercent)

        # 目标文件只有在参考目录存在相同大小文件时，才需要进入后续哈希比较。
        if (-not $referenceFilesByLength.ContainsKey($file.Length)) {
            continue
        }

        try {
            $partialHash = Get-PartialContentHash -File $file
        }
        catch {
            Write-Host "跳过文件，无法计算部分哈希: $($file.FullName)" -ForegroundColor Yellow
            Write-Host "  原因: $($_.Exception.Message)" -ForegroundColor DarkGray
            continue
        }

        $partialHashIndex = Get-ReferencePartialHashIndex -ReferenceIndex $ReferenceIndex -Length $file.Length
        if (-not $partialHashIndex.ContainsKey($partialHash)) {
            continue
        }

        $fullHashIndex = Get-ReferenceFullHashIndex -ReferenceIndex $ReferenceIndex -Length $file.Length -PartialHash $partialHash
        if ($fullHashIndex.Count -eq 0) {
            continue
        }

        try {
            $fullHash = Get-FullContentHash -File $file
        }
        catch {
            Write-Host "跳过文件，无法计算完整哈希: $($file.FullName)" -ForegroundColor Yellow
            Write-Host "  原因: $($_.Exception.Message)" -ForegroundColor DarkGray
            continue
        }

        if (-not $fullHashIndex.ContainsKey($fullHash)) {
            continue
        }

        if (-not $matchedTargetFilesByHash.ContainsKey($fullHash)) {
            $matchedTargetFilesByHash[$fullHash] = [pscustomobject]@{
                Hash           = $fullHash
                ReferenceFiles = @($fullHashIndex[$fullHash])
                TargetFiles    = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
            }
        }

        $matchedTargetFilesByHash[$fullHash].TargetFiles.Add($file)
    }

    Complete-ProgressBar

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
            DeletionItems = New-DeletionItems -FileList $matchingTargetFiles -RootPath $TargetRootPath -PathPrefix $targetPathPrefix
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
        return
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
        DeletionPlanList             = $deletionPlanList
        EmptyMessage                 = '未发现目标目录中存在与参考目录重复的文件。'
        PreviewSummaryFormat         = '重复文件列举完成。默认计划从目标目录删除重复文件: {0}'
        DefaultDeletionSummaryFormat = '删除完成。已从目标目录删除重复文件: {0}'
    }
    [void](Invoke-DeletionPlanAction @actionParameters)
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

if ($inputPathList.Count -eq 0 -and -not $useAggregateMode -and -not $useReferenceMode) {
    $selectedScanMode = Read-InteractiveScanMode
    switch ($selectedScanMode) {
        'Aggregate' {
            $useAggregateMode = $true
        }
        'Reference' {
            $useReferenceMode = $true
        }
        'Exit' {
            Write-Host "已退出，未执行扫描。" -ForegroundColor Yellow
            exit 0
        }
    }
}

$minimumPathCount = if ($useReferenceMode) { 2 } else { 1 }
$modeName = if ($useReferenceMode) {
    '参考目录模式'
}
elseif ($useAggregateMode) {
    '多目录合并模式'
}
else {
    '单目录模式'
}

if ($inputPathList.Count -eq 0) {
    $inputPathList = @(Read-InteractivePathList -ModeName $modeName -MinimumCount $minimumPathCount)
}

if ($inputPathList.Count -lt $minimumPathCount) {
    Write-Host "$modeName 至少需要输入 $minimumPathCount 个目录。" -ForegroundColor Red
    exit 1
}

try {
    $resolvedPathList = @(Resolve-InputDirectoryList -PathList $inputPathList)
}
catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

if ($useReferenceMode -or $useAggregateMode) {
    try {
        Test-IndependentDirectorySet -RootPathList $resolvedPathList
    }
    catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        exit 1
    }
}
else {
    try {
        Test-UniqueDirectorySet -RootPathList $resolvedPathList
    }
    catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        exit 1
    }
}

if ($useReferenceMode) {
    Invoke-ReferenceDirectoryMode -ReferenceRootPath $resolvedPathList[0] -TargetRootPathList @($resolvedPathList | Select-Object -Skip 1)
}
elseif ($useAggregateMode) {
    Invoke-MergedDirectoryMode -RootPathList $resolvedPathList
}
else {
    Invoke-IndependentSingleDirectoryMode -RootPathList $resolvedPathList
}
