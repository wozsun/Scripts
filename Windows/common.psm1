# Windows 脚本通用工具函数。
# 仅放跨脚本复用的输出、进度、延迟警告和路径处理逻辑；业务流程仍保留在各脚本内。

Set-StrictMode -Version Latest

$script:ProgressBarCellCount = 32
$script:ProgressBarFilledCharacter = '#'
$script:ProgressBarEmptyCharacter = '-'

# 配置通用控制台输出样式。
function Set-ConsoleOutputConfig {
    param(
        [Parameter(Mandatory = $false)]
        [int]$ProgressBarCellCount = $script:ProgressBarCellCount,

        [Parameter(Mandatory = $false)]
        [string]$ProgressBarFilledCharacter = $script:ProgressBarFilledCharacter,

        [Parameter(Mandatory = $false)]
        [string]$ProgressBarEmptyCharacter = $script:ProgressBarEmptyCharacter
    )

    if ($ProgressBarCellCount -gt 0) {
        $script:ProgressBarCellCount = $ProgressBarCellCount
    }

    if (-not [string]::IsNullOrEmpty($ProgressBarFilledCharacter)) {
        $script:ProgressBarFilledCharacter = $ProgressBarFilledCharacter
    }

    if (-not [string]::IsNullOrEmpty($ProgressBarEmptyCharacter)) {
        $script:ProgressBarEmptyCharacter = $ProgressBarEmptyCharacter
    }
}

# 输出当前执行阶段，避免大目录扫描时长时间无反馈。
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
        Write-Debug "动态状态行刷新已回退: $($_.Exception.Message)"
    }

    Write-Host -NoNewline "`r$Message" -ForegroundColor $Color
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

# 结束当前动态进度条并换行，避免后续日志和进度行混在一起。
function Complete-ProgressBar {
    Write-Host ""
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

    $messageWidth = Get-ConsoleTextWidth -Text $Message
    Write-DynamicStatusLine -Message $Message -Color $Color

    if ($NoNewLine) {
        return $messageWidth
    }

    Write-Host ""
    return 0
}

# 新建延迟输出的扫描警告列表，避免进度条刷新时被错误信息打断。
function New-DeferredScanWarningList {
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

    Write-Host "$($Title): $($WarningList.Count)" -ForegroundColor Yellow
    foreach ($warning in $WarningList) {
        Write-Host "$($warning.Message): $($warning.Path)" -ForegroundColor Yellow
        Write-Host "  原因: $($warning.Reason)" -ForegroundColor DarkGray
    }
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

        $canStartQuotedPath = $currentPartBuilder.Length -eq 0
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

Export-ModuleMember -Function @(
    'Set-ConsoleOutputConfig'
    'Write-StageMessage'
    'Write-MenuItem'
    'Read-ColoredLine'
    'Get-ConsoleCharacterCellWidth'
    'Get-ConsoleTextWithinCellWidth'
    'Get-ConsoleTextWidth'
    'Write-DynamicStatusLine'
    'Write-ProgressBar'
    'Complete-ProgressBar'
    'Write-RefreshStatusLine'
    'New-DeferredScanWarningList'
    'Add-DeferredScanWarning'
    'Write-DeferredScanWarningList'
    'ConvertTo-UnquotedPathText'
    'Split-InteractivePathInput'
    'ConvertTo-NormalizedDirectoryPath'
    'Get-DirectoryKey'
    'Add-TrailingDirectorySeparator'
    'Get-DirectoryPathDepth'
    'Test-DirectoryOverlap'
)
