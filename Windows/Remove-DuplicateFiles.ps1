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

Set-StrictMode -Version Latest

# 遇到未处理异常时立即进入 catch/退出流程，避免继续执行危险操作。
$ErrorActionPreference = 'Stop'

# 部分哈希预筛选时读取文件头尾的字节数；值越大越稳，扫描成本也越高。
$SampleHashByteCount = 1MB

# 删除预览中分隔不同重复文件组的横线文本。
$PreviewSeparatorText = '=' * 64

# 文本进度条宽度；统一使用 # 和 -，便于在不同终端中保持可读。
$ProgressBarCellCount = 28

# 使用 -yes 时的默认删除倒计时秒数，给用户留出取消窗口。
$AssumeYesGraceSeconds = 10

# 记录 -yes 倒计时是否被 Enter 取消；多目录分别执行时用于停止后续目录。
$AssumeYesDeletionCancelled = $false

# ========== 输出与路径工具 ==========

function Show-HelpText {
    Write-Host @'
用途：
  查找文件夹中的重复文件，先展示默认删除预览，再根据模式由用户确认默认删除、手动删除或退出。

用法：
  pwsh -File .\Remove-DuplicateFiles.ps1 [-s] [-yes] [Path1] [Path2 ...]
  pwsh -File .\Remove-DuplicateFiles.ps1 -a [-s] [-yes] <Path1> [Path2 ...]
  pwsh -File .\Remove-DuplicateFiles.ps1 -c [-s] [-yes] <ReferencePath> <TargetPath1> [TargetPath2 ...]
  pwsh -File .\Remove-DuplicateFiles.ps1 -h

参数：
  Path
    一个或多个文件夹绝对路径。不带 -a/-c 时，每个目录会分别执行单目录扫描。
  -a
    多目录合并模式：把所有输入目录抽象为一个大目录，可查找跨目录重复文件。
    输入目录不能相同，也不能互为父子目录。
  -c
    参考目录模式：第一个目录为参考目录，其余目录为目标目录。
    只删除目标目录中与参考目录内容完全相同的文件。
  -s
    包含隐藏文件和隐藏文件夹。默认只扫描未隐藏项。
  -yes
    跳过预览和菜单，等待 10 秒后执行默认删除计划；倒计时期间可按 Enter 取消，或按 Ctrl+C 中止。

模式：
  单目录模式：
    不带 -a/-c 时启用。可输入一个或多个目录，每个目录会单独扫描和处理。
    默认保留文件名最短的文件；如长度相同，再按文件名和完整路径排序。

  多目录合并模式：
    传入 -a 时启用。所有输入目录会合并为一个虚拟目录扫描，可删除跨目录重复文件。
    该模式提供默认删除、手动删除或退出选项。

  参考目录模式：
    传入 -c 时启用。第一个目录为参考目录，后续目录为目标目录。
    参考目录不参与删除，也不提供手动删除选项。

交互：
  脚本会先展示删除预览；单目录和多目录合并模式提供默认删除、手动删除或退出选项，参考目录模式提供默认删除或退出选项。
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
    $bar = ('#' * $filledWidth) + ('-' * $emptyWidth)
    $progressText = "[进度] $Activity [$bar] $percent% $Status ($ProcessedCount / $TotalCount)"

    Write-Host -NoNewline "`r$progressText$(' ' * 20)" -ForegroundColor Cyan
    $LastPercent.Value = $percent
}

# 结束当前进度条并换行，避免后续日志和动态进度混在一起。
function Complete-ProgressBar {
    Write-Host ""
}

# 输出用于区分不同预览或结果块的分隔线。
function Write-PreviewSeparator {
    Write-Host ""
    Write-Host $PreviewSeparatorText -ForegroundColor DarkGray
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

        for ($pollIndex = 0; $pollIndex -lt 10; $pollIndex++) {
            if (Test-EnterKeyPressed) {
                $script:AssumeYesDeletionCancelled = $true
                Write-Host "`r已取消 -yes 默认删除，未删除任何文件。$(' ' * 20)" -ForegroundColor Yellow
                return $false
            }

            Start-Sleep -Milliseconds 100
        }
    }

    Write-Host "`r倒计时结束，开始执行默认删除。$(' ' * 20)" -ForegroundColor Magenta
    return $true
}

# 兼容交互输入时复制带引号的路径；Windows 路径本身不允许包含引号。
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
        throw "$ParameterName 必须是 Windows 文件夹绝对路径。"
    }

    $resolvedPaths = @(Resolve-Path -LiteralPath $normalizedPathText -ErrorAction Stop)
    if ($resolvedPaths.Count -ne 1) {
        throw "$ParameterName 必须只能解析到一个目录。"
    }

    $item = Get-Item -LiteralPath $resolvedPaths[0].ProviderPath -ErrorAction Stop
    if (-not $item.PSIsContainer) {
        throw "$ParameterName 必须是文件夹。"
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
    Write-Host "未提供目录参数。请输入${ModeName}目录绝对路径，每行一个；直接回车开始执行或退出。" -ForegroundColor Yellow

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

        $inputPathList.Add($pathInput)
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
function Get-SampleContentHash {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    $fileStream = [System.IO.File]::Open($File.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)

    try {
        # 这里只做快速预筛选；真正删除前仍会使用完整 SHA-256 确认。
        $hashBuffer = [byte[]]::new($SampleHashByteCount)

        $firstRead = $fileStream.Read($hashBuffer, 0, $hashBuffer.Length)
        if ($firstRead -gt 0) {
            [void]$sha256.TransformBlock($hashBuffer, 0, $firstRead, $null, 0)
        }

        if ($File.Length -gt $SampleHashByteCount) {
            $tailOffset = [Math]::Max(0, $File.Length - $SampleHashByteCount)
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
    if ($IncludeHidden) {
        $scannedFiles = @(Get-ChildItem -LiteralPath $RootPath -File -Recurse -Force -ErrorAction SilentlyContinue -ErrorVariable scanErrorList)
    }
    else {
        $scannedFiles = @(Get-ChildItem -LiteralPath $RootPath -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable scanErrorList)
    }

    foreach ($scanError in @($scanErrorList)) {
        Write-Host "扫描跳过: $($scanError.TargetObject)" -ForegroundColor Yellow
        Write-Host "  原因: $($scanError.Exception.Message)" -ForegroundColor DarkGray
    }

    $hiddenScopeText = if ($IncludeHidden) { '包含隐藏项' } else { '不包含隐藏项' }
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

# 按默认保留规则排序文件：文件名更短者优先，其次按文件名和完整路径排序。
function Get-FilesByKeepPriority {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$FileList
    )

    return $FileList |
        Sort-Object @{ Expression = { $_.Name.Length }; Ascending = $true },
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
    $sampleCandidateCount = @($sameLengthGroups | ForEach-Object { $_.Group }).Count
    Write-StageMessage "$($ProgressLabel)大小相同的候选文件数: $sampleCandidateCount，候选大小组数: $($sameLengthGroups.Count)"

    $processedSampleHashCount = 0
    $lastSampleHashPercent = -1
    $hashProgressName = "$($ProgressLabel)哈希计算"

    foreach ($sizeGroup in $sameLengthGroups) {
        $sampleHashRecords = @(
            foreach ($file in $sizeGroup.Group) {
                $processedSampleHashCount++
                Write-ProgressBar -Activity $hashProgressName -Status '正在筛选候选文件' -ProcessedCount $processedSampleHashCount -TotalCount $sampleCandidateCount -LastPercent ([ref]$lastSampleHashPercent)

                try {
                    [pscustomobject]@{
                        File        = $file
                        SampleHash = Get-SampleContentHash -File $file
                    }
                }
                catch {
                    Write-Host "跳过文件，无法计算部分哈希: $($file.FullName)" -ForegroundColor Yellow
                    Write-Host "  原因: $($_.Exception.Message)" -ForegroundColor DarkGray
                }
            }
        )

        $sampleHashGroups = @($sampleHashRecords |
            Group-Object -Property SampleHash |
            Where-Object { $_.Count -gt 1 })

        foreach ($sampleHashGroup in $sampleHashGroups) {
            # 部分哈希只用于减少候选范围，最终仍按完整 SHA-256 分组确认。
            $contentHashRecords = @(
                foreach ($sampleHashRecord in $sampleHashGroup.Group) {
                    try {
                        [pscustomobject]@{
                            File     = $sampleHashRecord.File
                            ContentHash = Get-FullContentHash -File $sampleHashRecord.File
                        }
                    }
                    catch {
                        Write-Host "跳过文件，无法计算完整哈希: $($sampleHashRecord.File.FullName)" -ForegroundColor Yellow
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
        $manualInputText = Read-Host "请输入要删除的编号，多个编号用逗号分隔；直接回车使用默认规则；输入 0 跳过；输入 00 退出"
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
            Write-Host "输入无效，请输入列表中的编号，例如: 2 或 2,3；输入 0 跳过；输入 00 退出" -ForegroundColor Red
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

# 输出操作菜单并读取用户选择。
function Read-DeletionAction {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$MenuOptionList
    )

    Write-Host ""
    Write-Host "请选择操作:" -ForegroundColor Cyan
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
            return
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
}

# 单目录去重入口：先预览默认删除计划，再按用户选择执行默认或手动删除。
function Invoke-SingleDirectoryMode {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    $deletionPlanList = @(New-SingleDirectoryDeletionPlan -RootPath $RootPath)
    if ($deletionPlanList.Count -eq 0) {
        Write-Host "未发现重复文件。" -ForegroundColor Green
        return
    }

    if ($AssumeYes) {
        if (-not (Wait-AssumeYesDeletionGracePeriod)) {
            return
        }

        $deletionResult = Remove-DeletionItems -DeletionItemList @($deletionPlanList | ForEach-Object { $_.DeletionItems })
        Write-StatusSummary -Message "删除完成。已删除重复文件: $($deletionResult.DeletedCount)" -Color Magenta
        if ($deletionResult.FailedCount -gt 0) {
            Write-Host "删除失败文件: $($deletionResult.FailedCount)" -ForegroundColor Red
        }
        return
    }

    [void](Write-DeletionPlanPreview -DeletionPlanList $deletionPlanList -SummaryFormat '重复文件列举完成。默认计划删除重复文件: {0}')
    $menuChoice = Read-DeletionAction -MenuOptionList @(
        [pscustomobject]@{ Value = '1'; Label = '默认删除' }
        [pscustomobject]@{ Value = '2'; Label = '手动删除' }
        [pscustomobject]@{ Value = '0'; Label = '退出' }
    )

    if ($menuChoice -eq '0') {
        Write-Host "已退出，未删除任何文件。" -ForegroundColor Yellow
        return
    }

    if ($menuChoice -eq '1') {
        $deletionResult = Remove-DeletionItems -DeletionItemList @($deletionPlanList | ForEach-Object { $_.DeletionItems })
        Write-StatusSummary -Message "删除完成。已删除重复文件: $($deletionResult.DeletedCount)" -Color Magenta
        if ($deletionResult.FailedCount -gt 0) {
            Write-Host "删除失败文件: $($deletionResult.FailedCount)" -ForegroundColor Red
        }
        return
    }

    Invoke-ManualDeletion -DeletionPlanList $deletionPlanList -RootPath $RootPath
}

# 多个目录分别执行单目录模式；每个目录独立预览、独立确认。
function Invoke-IndependentSingleDirectoryMode {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$RootPathList
    )

    for ($rootIndex = 0; $rootIndex -lt $RootPathList.Count; $rootIndex++) {
        if ($RootPathList.Count -gt 1) {
            Write-PreviewSeparator
            Write-Host "单目录模式 $($rootIndex + 1) / $($RootPathList.Count): $($RootPathList[$rootIndex])" -ForegroundColor Cyan
        }

        Invoke-SingleDirectoryMode -RootPath $RootPathList[$rootIndex]
        if ($AssumeYesDeletionCancelled) {
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
    if ($deletionPlanList.Count -eq 0) {
        Write-Host "未发现重复文件。" -ForegroundColor Green
        return
    }

    if ($AssumeYes) {
        if (-not (Wait-AssumeYesDeletionGracePeriod)) {
            return
        }

        $deletionResult = Remove-DeletionItems -DeletionItemList @($deletionPlanList | ForEach-Object { $_.DeletionItems })
        Write-StatusSummary -Message "删除完成。已从多目录合并结果中删除重复文件: $($deletionResult.DeletedCount)" -Color Magenta
        if ($deletionResult.FailedCount -gt 0) {
            Write-Host "删除失败文件: $($deletionResult.FailedCount)" -ForegroundColor Red
        }
        return
    }

    [void](Write-DeletionPlanPreview -DeletionPlanList $deletionPlanList -SummaryFormat '重复文件列举完成。默认计划从多目录合并结果中删除重复文件: {0}')
    $menuChoice = Read-DeletionAction -MenuOptionList @(
        [pscustomobject]@{ Value = '1'; Label = '默认删除' }
        [pscustomobject]@{ Value = '2'; Label = '手动删除' }
        [pscustomobject]@{ Value = '0'; Label = '退出' }
    )

    if ($menuChoice -eq '0') {
        Write-Host "已退出，未删除任何文件。" -ForegroundColor Yellow
        return
    }

    if ($menuChoice -eq '1') {
        $deletionResult = Remove-DeletionItems -DeletionItemList @($deletionPlanList | ForEach-Object { $_.DeletionItems })
        Write-StatusSummary -Message "删除完成。已从多目录合并结果中删除重复文件: $($deletionResult.DeletedCount)" -Color Magenta
        if ($deletionResult.FailedCount -gt 0) {
            Write-Host "删除失败文件: $($deletionResult.FailedCount)" -ForegroundColor Red
        }
        return
    }

    $displayPathByFullName = New-DisplayPathMapFromDeletionPlan -DeletionPlanList $deletionPlanList
    Invoke-ManualDeletion -DeletionPlanList $deletionPlanList -DisplayPathByFullName $displayPathByFullName
}

# 为参考目录模式生成删除计划：参考目录只参与比较，目标目录才会进入删除列表。
function New-ReferenceDirectoryDeletionPlan {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ReferenceRootPath,

        [Parameter(Mandatory = $true)]
        [string]$TargetRootPath
    )

    $referenceFileList = @(Get-ScannedFiles -RootPath $ReferenceRootPath -ProgressLabel '参考目录')
    $targetFileList = @(Get-ScannedFiles -RootPath $TargetRootPath -ProgressLabel '目标目录')
    if ($referenceFileList.Count -eq 0 -or $targetFileList.Count -eq 0) {
        return @()
    }

    $referencePathPrefix = Get-DirectoryLabel -RootPath $ReferenceRootPath
    $targetPathPrefix = Get-DirectoryLabel -RootPath $TargetRootPath

    # 只用参考目录建立查找表；实际删除计划只从目标目录文件生成。
    $referenceFilesByLength = @{}
    $processedReferenceFileCount = 0
    $lastReferenceLengthPercent = -1
    foreach ($file in $referenceFileList) {
        $processedReferenceFileCount++
        Write-ProgressBar -Activity '参考目录文件大小分组' -Status '正在建立参考目录索引' -ProcessedCount $processedReferenceFileCount -TotalCount $referenceFileList.Count -LastPercent ([ref]$lastReferenceLengthPercent)

        $referenceFilesByLength[$file.Length] ??= [System.Collections.Generic.List[System.IO.FileInfo]]::new()
        $referenceFilesByLength[$file.Length].Add($file)
    }
    Complete-ProgressBar

    Write-StageMessage "参考目录模式按文件大小筛选目标目录候选文件..."
    $targetCandidatesByLength = @{}
    $processedTargetFileCount = 0
    $lastTargetLengthPercent = -1
    foreach ($file in $targetFileList) {
        $processedTargetFileCount++
        Write-ProgressBar -Activity '目标目录文件大小筛选' -Status '正在查找大小匹配文件' -ProcessedCount $processedTargetFileCount -TotalCount $targetFileList.Count -LastPercent ([ref]$lastTargetLengthPercent)

        # 目标文件只有在参考目录存在相同大小文件时，才需要进入后续哈希比较。
        if ($referenceFilesByLength.ContainsKey($file.Length)) {
            $targetCandidatesByLength[$file.Length] ??= [System.Collections.Generic.List[System.IO.FileInfo]]::new()
            $targetCandidatesByLength[$file.Length].Add($file)
        }
    }
    Complete-ProgressBar

    $targetCandidateLengthGroups = @(
        foreach ($size in $targetCandidatesByLength.Keys) {
            [pscustomobject]@{
                Name  = $size
                Count = $targetCandidatesByLength[$size].Count
                Group = @($targetCandidatesByLength[$size])
            }
        }
    )
    $targetCandidateFileCount = @($targetCandidateLengthGroups | ForEach-Object { $_.Group }).Count
    Write-StageMessage "目标目录大小匹配候选文件数: $targetCandidateFileCount，候选大小组数: $($targetCandidateLengthGroups.Count)"
    Write-StageMessage "参考目录哈希阶段会同时计算参考目录候选文件和目标目录候选文件。"

    $comparisonFileCount = 0
    foreach ($targetLengthGroup in $targetCandidateLengthGroups) {
        $comparisonFileCount += @($referenceFilesByLength[[int64]$targetLengthGroup.Name]).Count + $targetLengthGroup.Group.Count
    }

    $processedSampleHashCount = 0
    $lastSampleHashPercent = -1
    $hashProgressName = '参考目录哈希计算'

    foreach ($targetLengthGroup in $targetCandidateLengthGroups) {
        $sameLengthReferenceFiles = @($referenceFilesByLength[[int64]$targetLengthGroup.Name])
        $comparisonFileRecords = @(
            $sameLengthReferenceFiles | ForEach-Object {
                [pscustomobject]@{
                    File = $_
                    Side = 'Reference'
                }
            }

            $targetLengthGroup.Group | ForEach-Object {
                [pscustomobject]@{
                    File = $_
                    Side = 'Target'
                }
            }
        )

        $sampleHashRecords = @(
            foreach ($comparisonFileRecord in $comparisonFileRecords) {
                $processedSampleHashCount++
                Write-ProgressBar -Activity $hashProgressName -Status '正在比较参考目录和目标目录' -ProcessedCount $processedSampleHashCount -TotalCount $comparisonFileCount -LastPercent ([ref]$lastSampleHashPercent)

                try {
                    [pscustomobject]@{
                        File        = $comparisonFileRecord.File
                        Side        = $comparisonFileRecord.Side
                        SampleHash = Get-SampleContentHash -File $comparisonFileRecord.File
                    }
                }
                catch {
                    Write-Host "跳过文件，无法计算部分哈希: $($comparisonFileRecord.File.FullName)" -ForegroundColor Yellow
                    Write-Host "  原因: $($_.Exception.Message)" -ForegroundColor DarkGray
                }
            }
        )

        $sampleHashGroups = @($sampleHashRecords |
            Group-Object -Property SampleHash |
            Where-Object {
                @($_.Group | Where-Object { $_.Side -eq 'Reference' }).Count -gt 0 -and
                @($_.Group | Where-Object { $_.Side -eq 'Target' }).Count -gt 0
            })

        foreach ($sampleHashGroup in $sampleHashGroups) {
            # 只有同一完整哈希里同时存在参考文件和目标文件时，才生成目标目录删除计划。
            $contentHashRecords = @(
                foreach ($sampleHashRecord in $sampleHashGroup.Group) {
                    try {
                        [pscustomobject]@{
                            File     = $sampleHashRecord.File
                            Side     = $sampleHashRecord.Side
                            ContentHash = Get-FullContentHash -File $sampleHashRecord.File
                        }
                    }
                    catch {
                        Write-Host "跳过文件，无法计算完整哈希: $($sampleHashRecord.File.FullName)" -ForegroundColor Yellow
                        Write-Host "  原因: $($_.Exception.Message)" -ForegroundColor DarkGray
                    }
                }
            )

            $contentHashGroups = @($contentHashRecords |
                Group-Object -Property ContentHash |
                Where-Object {
                    @($_.Group | Where-Object { $_.Side -eq 'Reference' }).Count -gt 0 -and
                    @($_.Group | Where-Object { $_.Side -eq 'Target' }).Count -gt 0
                })

            foreach ($contentHashGroup in $contentHashGroups) {
                $matchingReferenceFiles = @($contentHashGroup.Group | Where-Object { $_.Side -eq 'Reference' } | ForEach-Object { $_.File })
                $referenceKeepFile = Select-DefaultKeepFile -FileList $matchingReferenceFiles
                $matchingTargetFiles = @($contentHashGroup.Group | Where-Object { $_.Side -eq 'Target' } | ForEach-Object { $_.File })

                [pscustomobject]@{
                    Hash          = $contentHashGroup.Name
                    KeepPathText = Get-RelativePathText -File $referenceKeepFile -RootPath $ReferenceRootPath -PathPrefix $referencePathPrefix
                    DeletionItems = New-DeletionItems -FileList $matchingTargetFiles -RootPath $TargetRootPath -PathPrefix $targetPathPrefix
                }
            }
        }
    }

    Complete-ProgressBar
}

# 参考目录去重入口：以第一个目录为参考，只删除后续目标目录中的重复文件。
function Invoke-ReferenceDirectoryMode {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ReferenceRootPath,

        [Parameter(Mandatory = $true)]
        [string[]]$TargetRootPathList
    )

    $deletionPlanList = @(
        for ($targetIndex = 0; $targetIndex -lt $TargetRootPathList.Count; $targetIndex++) {
            if ($TargetRootPathList.Count -gt 1) {
                Write-PreviewSeparator
                Write-Host "参考目录模式目标 $($targetIndex + 1) / $($TargetRootPathList.Count): $($TargetRootPathList[$targetIndex])" -ForegroundColor Cyan
            }

            New-ReferenceDirectoryDeletionPlan -ReferenceRootPath $ReferenceRootPath -TargetRootPath $TargetRootPathList[$targetIndex]
        }
    )

    if ($deletionPlanList.Count -eq 0) {
        Write-Host "未发现目标目录中存在与参考目录重复的文件。" -ForegroundColor Green
        return
    }

    if ($AssumeYes) {
        if (-not (Wait-AssumeYesDeletionGracePeriod)) {
            return
        }

        $deletionResult = Remove-DeletionItems -DeletionItemList @($deletionPlanList | ForEach-Object { $_.DeletionItems })
        Write-StatusSummary -Message "删除完成。已从目标目录删除重复文件: $($deletionResult.DeletedCount)" -Color Magenta
        if ($deletionResult.FailedCount -gt 0) {
            Write-Host "删除失败文件: $($deletionResult.FailedCount)" -ForegroundColor Red
        }
        return
    }

    [void](Write-DeletionPlanPreview -DeletionPlanList $deletionPlanList -SummaryFormat '重复文件列举完成。默认计划从目标目录删除重复文件: {0}')
    $menuChoice = Read-DeletionAction -MenuOptionList @(
        [pscustomobject]@{ Value = '1'; Label = '默认删除' }
        [pscustomobject]@{ Value = '0'; Label = '退出' }
    )

    if ($menuChoice -eq '0') {
        Write-Host "已退出，未删除任何文件。" -ForegroundColor Yellow
        return
    }

    $deletionResult = Remove-DeletionItems -DeletionItemList @($deletionPlanList | ForEach-Object { $_.DeletionItems })
    Write-StatusSummary -Message "删除完成。已从目标目录删除重复文件: $($deletionResult.DeletedCount)" -Color Magenta
    if ($deletionResult.FailedCount -gt 0) {
        Write-Host "删除失败文件: $($deletionResult.FailedCount)" -ForegroundColor Red
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

$minimumPathCount = if ($ReferenceMode) { 2 } else { 1 }
$modeName = if ($ReferenceMode) {
    '参考目录模式'
}
elseif ($AggregateMode) {
    '多目录合并模式'
}
else {
    '单目录模式'
}

$inputPathList = @($PathList | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
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

if ($ReferenceMode -or $AggregateMode) {
    try {
        Test-IndependentDirectorySet -RootPathList $resolvedPathList
    }
    catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        exit 1
    }
}

if ($ReferenceMode) {
    Invoke-ReferenceDirectoryMode -ReferenceRootPath $resolvedPathList[0] -TargetRootPathList @($resolvedPathList | Select-Object -Skip 1)
}
elseif ($AggregateMode) {
    Invoke-MergedDirectoryMode -RootPathList $resolvedPathList
}
else {
    Invoke-IndependentSingleDirectoryMode -RootPathList $resolvedPathList
}
