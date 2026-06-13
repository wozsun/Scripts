#requires -Version 7.0
<#
用途：
  批量解压指定 .zip / .rar / .7z 压缩包，或递归扫描指定文件夹中的 .zip / .rar / .7z 压缩包并批量解压。

参数：
  Path   一个或多个 .zip / .rar / .7z 压缩包或文件夹绝对路径；未提供时会引导交互输入。
  -s     扫描目录时包含隐藏文件和隐藏文件夹。
  -yes   跳过解压预览菜单，并在解压成功后默认删除原压缩包。
  -h     显示帮助信息。

规则：
  .zip 使用 .NET 原生能力解压，并对旧中文 ZIP 文件名编码做 GBK 回退；.rar / .7z 优先使用 Windows 内置 tar.exe，失败时回退到其他可用工具。
  每个压缩包解压到自身所在目录下的同名文件夹。
  如果目标文件夹名已存在，则自动使用带序号的新文件夹名，不覆盖已有内容。
  解压时先写入同目录下的 .expand-archivefiles-tmp 临时文件夹；若该目录已存在则清空复用。成功后再移动为最终文件夹，失败时清理本轮临时文件。
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

# ========== 可调整配置 ==========

# 支持的压缩包扩展名。.zip 使用 .NET 原生能力，.rar / .7z 优先使用 Windows 内置 tar.exe，失败时回退到其他工具。
$SupportedArchiveExtensionList = @('.zip', '.rar', '.7z')

# 解压时在压缩包所在目录下创建的脚本专属临时目录名。
$TempDirectoryName = '.expand-archivefiles-tmp'

# 目标文件夹重名时使用的命名格式，例如 Archive (2)。
$NumberedDirectoryNameFormat = '{0} ({1})'

# 旧 ZIP 若没有 UTF-8 标记，中文 Windows 压缩软件常用 CP936/GBK 存储文件名。
$ZipLegacyEntryNameCodePageList = @(936)

# ========== 运行环境设置 ==========

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# 加载 Windows 脚本公共工具函数。
Import-Module -Name (Join-Path $PSScriptRoot 'common.psm1') -Force

# ========== 参数派生选项 ==========

$ShouldIncludeHiddenItems = [bool]$IncludeHidden
$SupportedArchiveExtensionSet = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]$SupportedArchiveExtensionList,
    [System.StringComparer]::OrdinalIgnoreCase
)

# 按扩展名缓存已发现的外部解压工具列表，避免批量处理时重复搜索 PATH 和常见安装目录。
$script:ArchiveExtractorListCacheByExtension = @{}

# 记录是否已注册 .NET 代码页编码提供器，避免批量解压时重复注册。
$script:ZipLegacyEncodingProviderRegistered = $false

# ========== 输出与工具函数 ==========

# 输出脚本帮助文本。
function Show-HelpText {
    Write-Host @'
用途：
  批量解压指定 .zip / .rar / .7z 压缩包，或递归扫描指定文件夹中的 .zip / .rar / .7z 压缩包并批量解压。

用法：
  pwsh -File .\Expand-ArchiveFiles.ps1 [-s] [-yes] [Path1] [Path2 ...]

参数：
  Path   一个或多个 .zip / .rar / .7z 压缩包或文件夹绝对路径；文件夹会递归扫描；未输入路径时会引导交互输入。
  -s     扫描目录时包含隐藏文件和隐藏文件夹。默认只扫描未隐藏项。
  -yes   跳过解压预览和删除确认菜单，解压后默认删除成功解压的原压缩包。
  -h     显示帮助信息。

规则：
  支持 .zip、.rar 和 .7z 压缩包。
  .zip 使用 .NET 原生能力解压，并对旧中文 ZIP 文件名编码做 GBK 回退；.rar / .7z 优先使用 Windows 内置 tar.exe，失败或缺失时回退到 7-Zip 或 WinRAR；.rar 还会尝试 UnRAR。
  输入路径必须是 Windows 文件或文件夹绝对路径；命令行传参时，路径包含空格、括号等 PowerShell 特殊字符，请使用英文引号包裹路径。
  多个输入路径可包含相同或互相包含的目录，脚本会按压缩包真实路径自动去重。
  每个压缩包会解压到自身所在目录下的同名文件夹。
  如果目标文件夹名已存在，则自动追加序号，例如 Archive (2)，不会覆盖已有文件夹或文件。
  解压时先写入同目录下的 .expand-archivefiles-tmp 临时文件夹；若该目录已存在则清空复用。
  解压完成后会询问是否删除成功解压的原压缩包；启用 -yes 时默认删除。
  单个压缩包失败时会记录原因并继续处理后续压缩包。
'@
}

# 判断文件是否为当前支持的压缩包格式。
function Test-SupportedArchiveFile {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    return $SupportedArchiveExtensionSet.Contains($File.Extension)
}

# 判断路径是否位于本脚本创建的临时目录内。
function Test-PathInsideScriptTempDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    foreach ($pathSegment in ($Path -split '[\\/]')) {
        if (Test-SiblingTempDirectoryName -DirectoryName $pathSegment -TempDirectoryName $TempDirectoryName) {
            return $true
        }
    }

    return $false
}

# 从 PATH 中查找第一个可执行应用程序路径。
function Get-CommandApplicationPath {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$CommandNameList
    )

    foreach ($commandName in $CommandNameList) {
        $command = Get-Command -Name $commandName -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $command -and -not [string]::IsNullOrWhiteSpace($command.Source)) {
            return $command.Source
        }
    }

    return $null
}

# 将常见安装目录下的工具路径加入候选列表；是否存在由后续函数统一判断。
function Add-ToolPathCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[string]]$CandidatePathList,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$BasePath,

        [Parameter(Mandatory = $true)]
        [string]$RelativePath
    )

    if ([string]::IsNullOrWhiteSpace($BasePath)) {
        return
    }

    $CandidatePathList.Add([System.IO.Path]::Combine($BasePath, $RelativePath))
}

# 从候选路径中返回第一个真实存在的文件路径。
function Get-ExistingFilePathCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$CandidatePathList
    )

    foreach ($candidatePath in $CandidatePathList) {
        if (-not [string]::IsNullOrWhiteSpace($candidatePath) -and
            (Test-FileSystemFile -Path $candidatePath)) {
            return $candidatePath
        }
    }

    return $null
}

# 获取 Windows 内置 tar.exe 路径；找不到系统路径时再查 PATH。
function Get-WindowsTarPath {
    if (-not [string]::IsNullOrWhiteSpace($env:SystemRoot)) {
        $systemTarPath = [System.IO.Path]::Combine($env:SystemRoot, 'System32', 'tar.exe')
        if (Test-FileSystemFile -Path $systemTarPath) {
            return $systemTarPath
        }
    }

    return Get-CommandApplicationPath -CommandNameList @('tar')
}

# 获取指定扩展名可用的外部解压工具列表，按优先级排序。
function Get-ArchiveExtractorList {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Extension
    )

    $extensionKey = $Extension.ToLowerInvariant()
    if ($script:ArchiveExtractorListCacheByExtension.ContainsKey($extensionKey)) {
        return $script:ArchiveExtractorListCacheByExtension[$extensionKey]
    }

    $archiveExtractorList = [System.Collections.Generic.List[object]]::new()

    $windowsTarPath = Get-WindowsTarPath
    if (-not [string]::IsNullOrWhiteSpace($windowsTarPath)) {
        $archiveExtractorList.Add([pscustomobject]@{
            Kind        = 'WindowsTar'
            Path        = $windowsTarPath
            DisplayName = 'Windows 内置 tar.exe'
        })
    }

    $sevenZipPath = Get-CommandApplicationPath -CommandNameList @('7z', '7za', '7zz')
    if ([string]::IsNullOrWhiteSpace($sevenZipPath)) {
        $sevenZipPathCandidateList = [System.Collections.Generic.List[string]]::new()
        Add-ToolPathCandidate -CandidatePathList $sevenZipPathCandidateList -BasePath $env:ProgramFiles -RelativePath '7-Zip\7z.exe'
        Add-ToolPathCandidate -CandidatePathList $sevenZipPathCandidateList -BasePath ([Environment]::GetEnvironmentVariable('ProgramFiles(x86)')) -RelativePath '7-Zip\7z.exe'
        Add-ToolPathCandidate -CandidatePathList $sevenZipPathCandidateList -BasePath $env:LocalAppData -RelativePath 'Programs\7-Zip\7z.exe'
        $sevenZipPath = Get-ExistingFilePathCandidate -CandidatePathList $sevenZipPathCandidateList.ToArray()
    }

    if (-not [string]::IsNullOrWhiteSpace($sevenZipPath)) {
        $archiveExtractorList.Add([pscustomobject]@{
            Kind        = 'SevenZip'
            Path        = $sevenZipPath
            DisplayName = '7-Zip'
        })
    }

    if ($extensionKey -eq '.rar') {
        $unRarPath = Get-CommandApplicationPath -CommandNameList @('UnRAR', 'unrar', 'rar')
        if ([string]::IsNullOrWhiteSpace($unRarPath)) {
            $unRarPathCandidateList = [System.Collections.Generic.List[string]]::new()
            Add-ToolPathCandidate -CandidatePathList $unRarPathCandidateList -BasePath $env:ProgramFiles -RelativePath 'WinRAR\UnRAR.exe'
            Add-ToolPathCandidate -CandidatePathList $unRarPathCandidateList -BasePath ([Environment]::GetEnvironmentVariable('ProgramFiles(x86)')) -RelativePath 'WinRAR\UnRAR.exe'
            Add-ToolPathCandidate -CandidatePathList $unRarPathCandidateList -BasePath $env:ProgramFiles -RelativePath 'WinRAR\Rar.exe'
            Add-ToolPathCandidate -CandidatePathList $unRarPathCandidateList -BasePath ([Environment]::GetEnvironmentVariable('ProgramFiles(x86)')) -RelativePath 'WinRAR\Rar.exe'
            $unRarPath = Get-ExistingFilePathCandidate -CandidatePathList $unRarPathCandidateList.ToArray()
        }

        if (-not [string]::IsNullOrWhiteSpace($unRarPath)) {
            $archiveExtractorList.Add([pscustomobject]@{
                Kind        = 'UnRar'
                Path        = $unRarPath
                DisplayName = 'UnRAR / RAR'
            })
        }
    }

    $winRarPath = Get-CommandApplicationPath -CommandNameList @('WinRAR')
    if ([string]::IsNullOrWhiteSpace($winRarPath)) {
        $winRarPathCandidateList = [System.Collections.Generic.List[string]]::new()
        Add-ToolPathCandidate -CandidatePathList $winRarPathCandidateList -BasePath $env:ProgramFiles -RelativePath 'WinRAR\WinRAR.exe'
        Add-ToolPathCandidate -CandidatePathList $winRarPathCandidateList -BasePath ([Environment]::GetEnvironmentVariable('ProgramFiles(x86)')) -RelativePath 'WinRAR\WinRAR.exe'
        $winRarPath = Get-ExistingFilePathCandidate -CandidatePathList $winRarPathCandidateList.ToArray()
    }

    if (-not [string]::IsNullOrWhiteSpace($winRarPath)) {
        $archiveExtractorList.Add([pscustomobject]@{
            Kind        = 'WinRAR'
            Path        = $winRarPath
            DisplayName = 'WinRAR'
        })
    }

    $script:ArchiveExtractorListCacheByExtension[$extensionKey] = $archiveExtractorList.ToArray()
    return $script:ArchiveExtractorListCacheByExtension[$extensionKey]
}

# 交互读取一个或多个输入路径。
function Read-InteractivePathList {
    Write-Host "请输入 .zip / .rar / .7z 压缩包或目录绝对路径。可在同一行输入多个路径。" -ForegroundColor Cyan
    Write-Host "多个路径可用空格或英文分号分隔；路径含空格或英文分号时，请使用英文引号包裹路径。" -ForegroundColor DarkGray
    Write-Host "直接回车开始执行或退出；输入 0 退出脚本。" -ForegroundColor DarkGray

    $inputItemList = [System.Collections.Generic.List[System.IO.FileSystemInfo]]::new()
    $inputItemKeySet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    while ($true) {
        $promptIndex = $inputItemList.Count + 1
        $inputLine = Read-ColoredLine -Prompt "Path${promptIndex}: "

        if ([string]::IsNullOrWhiteSpace($inputLine)) {
            if ($inputItemList.Count -eq 0) {
                Write-Host "已退出，未执行扫描。" -ForegroundColor Yellow
                exit 0
            }

            return $inputItemList.ToArray()
        }

        if ($inputLine.Trim() -eq '0') {
            Write-Host "已退出，未执行扫描。" -ForegroundColor Yellow
            exit 0
        }

        $rawPathList = @(Split-InteractivePathInput -PathInput $inputLine)
        if ($rawPathList.Count -eq 0) {
            Write-Host "输入无效，请重新输入 .zip / .rar / .7z 压缩包或目录绝对路径。" -ForegroundColor Red
            continue
        }

        $resolvedResult = Resolve-InputPathList -PathList $rawPathList -ExistingKeySet $inputItemKeySet
        if (-not $resolvedResult.Success) {
            Write-Host "$($resolvedResult.Error) 本次输入不保留。" -ForegroundColor Red
            continue
        }

        foreach ($resolvedItem in @($resolvedResult.Items)) {
            $inputItemList.Add($resolvedItem)
        }

        $resolvedPathCount = @($resolvedResult.Items).Count
        if ($resolvedPathCount -eq 0) {
            Write-Host "输入路径已存在，本次未新增。" -ForegroundColor DarkGray
        }
        elseif ($rawPathList.Count -gt 1) {
            Write-Host "识别到 $resolvedPathCount 个新路径。" -ForegroundColor DarkGray
        }
    }
}

# 将压缩包加入候选列表，并统计重复项。
function Add-ArchiveCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$ArchiveFile,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[string]]$ArchivePathSet,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[System.IO.FileInfo]]$ArchiveFileList,

        [Parameter(Mandatory = $true)]
        [ref]$DuplicateCount
    )

    $archiveKey = ConvertTo-NormalizedPath -Path $ArchiveFile.FullName
    if ($ArchivePathSet.Add($archiveKey)) {
        $ArchiveFileList.Add($ArchiveFile)
        return
    }

    $DuplicateCount.Value++
}

# 校验直接输入的文件是否为可处理压缩包。
function Add-InputFileArchiveCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$InputFile,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[string]]$ArchivePathSet,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[System.IO.FileInfo]]$ArchiveFileList,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$WarningList,

        [Parameter(Mandatory = $true)]
        [ref]$DuplicateCount
    )

    if (-not $ShouldIncludeHiddenItems -and (Test-HiddenFileSystemItem -Item $InputFile)) {
        Add-DeferredScanWarning `
            -WarningList $WarningList `
            -Message '跳过输入文件' `
            -Path $InputFile.FullName `
            -Reason '文件为隐藏项；如需处理隐藏项，请传入 -s。'
        return
    }

    if (-not (Test-SupportedArchiveFile -File $InputFile)) {
        Add-DeferredScanWarning `
            -WarningList $WarningList `
            -Message '跳过输入文件' `
            -Path $InputFile.FullName `
            -Reason "当前仅支持: $($SupportedArchiveExtensionList -join ', ')。"
        return
    }

    Add-ArchiveCandidate `
        -ArchiveFile $InputFile `
        -ArchivePathSet $ArchivePathSet `
        -ArchiveFileList $ArchiveFileList `
        -DuplicateCount $DuplicateCount
}

# 递归扫描目录中的可处理压缩包。
function Add-DirectoryArchiveCandidates {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$InputDirectory,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[string]]$ArchivePathSet,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[System.IO.FileInfo]]$ArchiveFileList,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$WarningList,

        [Parameter(Mandatory = $true)]
        [ref]$DuplicateCount
    )

    if (-not $ShouldIncludeHiddenItems -and (Test-HiddenFileSystemItem -Item $InputDirectory)) {
        Add-DeferredScanWarning `
            -WarningList $WarningList `
            -Message '跳过输入目录' `
            -Path $InputDirectory.FullName `
            -Reason '目录为隐藏项；如需处理隐藏项，请传入 -s。'
        return
    }

    Write-StageMessage "开始扫描目录: $($InputDirectory.FullName)"
    $scanErrors = @()
    $getChildItemParameters = @{
        LiteralPath = $InputDirectory.FullName
        File        = $true
        Recurse     = $true
        ErrorAction = 'SilentlyContinue'
    }
    if ($ShouldIncludeHiddenItems) {
        $getChildItemParameters.Force = $true
    }

    $allFileList = @(Get-ChildItem @getChildItemParameters -ErrorVariable scanErrors)
    $archiveFileList = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
    foreach ($file in $allFileList) {
        if ((Test-PathInsideScriptTempDirectory -Path $file.FullName) -or
            -not (Test-SupportedArchiveFile -File $file)) {
            continue
        }

        $archiveFileList.Add($file)
    }
    $archiveFileListInDirectory = $archiveFileList.ToArray()

    foreach ($scanError in @($scanErrors)) {
        Add-DeferredScanWarning `
            -WarningList $WarningList `
            -Message '扫描跳过' `
            -Path $scanError.TargetObject `
            -Reason $scanError.Exception.Message
    }

    foreach ($archiveFile in $archiveFileListInDirectory) {
        Add-ArchiveCandidate `
            -ArchiveFile $archiveFile `
            -ArchivePathSet $ArchivePathSet `
            -ArchiveFileList $ArchiveFileList `
            -DuplicateCount $DuplicateCount
    }

    $hiddenModeText = if ($ShouldIncludeHiddenItems) { '包含隐藏项' } else { '不包含隐藏项' }
    Write-StageMessage "目录扫描完成，文件数: $($allFileList.Count)，压缩包: $($archiveFileListInDirectory.Count)，$hiddenModeText"
}

# 根据压缩包文件名生成默认解压目录名。
function Get-ArchiveTargetBaseName {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$ArchiveFile
    )

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ArchiveFile.Name)
    if ([string]::IsNullOrWhiteSpace($baseName)) {
        return '解压文件'
    }

    if (Test-SiblingTempDirectoryName -DirectoryName $baseName -TempDirectoryName $TempDirectoryName) {
        return "$baseName 解压"
    }

    return $baseName
}

# 生成不会覆盖现有目录、也不会和当前计划冲突的目标目录路径。
function Get-UniqueDirectoryPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ParentPath,

        [Parameter(Mandatory = $true)]
        [string]$BaseName,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[string]]$ReservedPathSet
    )

    $directoryIndex = 1
    while ($true) {
        $directoryName = if ($directoryIndex -eq 1) {
            $BaseName
        }
        else {
            $NumberedDirectoryNameFormat -f $BaseName, $directoryIndex
        }

        $candidatePath = [System.IO.Path]::Combine($ParentPath, $directoryName)
        $candidateKey = Get-DirectoryKey -DirectoryPath $candidatePath
        if (Test-SiblingTempDirectoryName -DirectoryName $directoryName -TempDirectoryName $TempDirectoryName) {
            $directoryIndex++
            continue
        }

        if (-not (Test-FileSystemPath -Path $candidatePath) -and $ReservedPathSet.Add($candidateKey)) {
            return $candidatePath
        }

        $directoryIndex++
    }
}

# 扫描输入项并生成最终解压计划。
function New-ArchiveExtractionPlan {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileSystemInfo[]]$InputItemList
    )

    # 扫描警告延迟输出，避免打断进度条；压缩包路径集合用于跨输入去重。
    $warningList = New-DeferredScanWarningList
    $archivePathSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $archiveFileList = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
    $duplicateCount = 0

    foreach ($inputItem in $InputItemList) {
        if ($inputItem -is [System.IO.DirectoryInfo]) {
            Add-DirectoryArchiveCandidates `
                -InputDirectory $inputItem `
                -ArchivePathSet $archivePathSet `
                -ArchiveFileList $archiveFileList `
                -WarningList $warningList `
                -DuplicateCount ([ref]$duplicateCount)
            continue
        }

        Add-InputFileArchiveCandidate `
            -InputFile ([System.IO.FileInfo]$inputItem) `
            -ArchivePathSet $archivePathSet `
            -ArchiveFileList $archiveFileList `
            -WarningList $warningList `
            -DuplicateCount ([ref]$duplicateCount)
    }

    # 目标目录也要提前保留，避免两个同名压缩包计划到同一目录。
    $reservedTargetPathSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $planList = [System.Collections.Generic.List[object]]::new()
    $processedCount = 0
    $lastPercent = -1

    foreach ($archiveFile in $archiveFileList) {
        $processedCount++
        Write-ProgressBar `
            -Activity '解压计划生成' `
            -Status '正在分配目标目录' `
            -ProcessedCount $processedCount `
            -TotalCount $archiveFileList.Count `
            -LastPercent ([ref]$lastPercent)

        $targetBaseName = Get-ArchiveTargetBaseName -ArchiveFile $archiveFile
        $parentPath = $archiveFile.DirectoryName
        $targetPath = Get-UniqueDirectoryPath `
            -ParentPath $parentPath `
            -BaseName $targetBaseName `
            -ReservedPathSet $reservedTargetPathSet

        $planList.Add([pscustomobject]@{
                ArchivePath    = $archiveFile.FullName
                ArchiveName    = $archiveFile.Name
                Extension      = $archiveFile.Extension.ToLowerInvariant()
                ParentPath     = $parentPath
                TargetBaseName = $targetBaseName
                TargetPath     = $targetPath
                Length         = $archiveFile.Length
            })
    }

    Write-ProgressBar `
        -Activity '解压计划生成' `
        -Status '解压计划生成完成' `
        -ProcessedCount $archiveFileList.Count `
        -TotalCount $archiveFileList.Count `
        -LastPercent ([ref]$lastPercent) `
        -Force
    Complete-DynamicStatusLine

    return [pscustomobject]@{
        Items          = $planList.ToArray()
        Warnings       = $warningList
        DuplicateCount = $duplicateCount
    }
}

# 输出解压计划预览。
function Write-ExtractionPreview {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$ExtractionPlanList,

        [Parameter(Mandatory = $true)]
        [int]$DuplicateCount
    )

    Write-Host
    Write-Host "解压预览:" -ForegroundColor Yellow
    Write-PreviewSeparator -NoLeadingBlank

    $index = 0
    foreach ($planItem in $ExtractionPlanList) {
        $index++
        Write-Host -NoNewline ("{0,4}. " -f $index) -ForegroundColor Cyan
        Write-Host $planItem.ArchivePath -ForegroundColor White
        Write-Host "      -> $($planItem.TargetPath)" -ForegroundColor DarkGray
    }

    Write-PreviewSeparator -NoLeadingBlank
    Write-Host
    Write-Host -NoNewline "压缩包列举完成。默认计划解压: " -ForegroundColor White
    Write-Host -NoNewline $ExtractionPlanList.Count -ForegroundColor Magenta
    Write-Host -NoNewline "，已去重: " -ForegroundColor White
    Write-Host $DuplicateCount -ForegroundColor Yellow
}

# 输出非 zip 压缩包将使用的外部解压工具信息。
function Write-ArchiveExtractorStatus {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$ExtractionPlanList
    )

    $externalArchivePlanList = [System.Collections.Generic.List[object]]::new()
    foreach ($planItem in $ExtractionPlanList) {
        if ($planItem.Extension -ne '.zip') {
            $externalArchivePlanList.Add($planItem)
        }
    }

    $nativeArchiveGroupList = @($externalArchivePlanList.ToArray() | Group-Object -Property Extension | Sort-Object -Property Name)

    foreach ($archiveGroup in $nativeArchiveGroupList) {
        $extension = $archiveGroup.Name
        $archiveExtractorList = @(Get-ArchiveExtractorList -Extension $extension)
        if ($archiveExtractorList.Count -eq 0) {
            Write-Host "检测到 $extension 压缩包 $($archiveGroup.Count) 个，但未找到可用解压工具；解压会失败。" -ForegroundColor Yellow
            continue
        }

        $archiveExtractorSummary = @(
            $archiveExtractorList | ForEach-Object { "$($_.DisplayName): $($_.Path)" }
        ) -join '；备用 '
        Write-Host "检测到 $extension 压缩包 $($archiveGroup.Count) 个，解压工具: $archiveExtractorSummary" -ForegroundColor DarkGray
    }
}

# 读取解压阶段操作选择。
function Read-OperationChoice {
    $choice = Read-MenuChoice -Title '请选择操作:' -EndOfInputChoice '0' -MenuOptionList @(
        [pscustomobject]@{ Value = '1'; Label = '默认解压' }
        [pscustomobject]@{ Value = '2'; Label = '手动解压' }
        [pscustomobject]@{ Value = '0'; Label = '退出脚本' }
    )

    switch ($choice) {
        '1' { return 'Default' }
        '2' { return 'Manual' }
        '0' { return 'Exit' }
    }
}

# 读取手动解压编号选择。
function Read-ManualExtractionSelection {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ExtractionPlanList
    )

    while ($true) {
        Write-Host
        Write-Host "请输入要解压的编号，多个编号用英文逗号分隔；输入 0 取消，输入 00 退出脚本。" -ForegroundColor Cyan
        $inputText = Read-ColoredLine -Prompt '编号: '
        if ($null -eq $inputText) {
            Write-Host "输入流已结束，程序退出。" -ForegroundColor Yellow
            exit 0
        }

        $trimmedInput = $inputText.Trim()
        if ($trimmedInput -eq '00') {
            Write-Host "已退出，未继续解压。" -ForegroundColor Yellow
            exit 0
        }

        if ($trimmedInput -eq '0') {
            Write-Host "已取消手动解压。" -ForegroundColor Yellow
            return @()
        }

        $selectedIndexSet = [System.Collections.Generic.HashSet[int]]::new()
        $isValid = $true

        foreach ($indexText in ($trimmedInput -split ',')) {
            $parsedIndex = 0
            if (-not [int]::TryParse($indexText.Trim(), [ref]$parsedIndex) -or
                $parsedIndex -lt 1 -or
                $parsedIndex -gt $ExtractionPlanList.Count) {
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
            $selectedItemList.Add($ExtractionPlanList[$selectedIndex - 1])
        }

        return $selectedItemList.ToArray()
    }
}

# 准备本轮临时目录；已有残留临时目录时会清空复用。
function Initialize-ExtractionTempDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ParentPath
    )

    return Initialize-SiblingTempDirectory `
        -ParentPath $ParentPath `
        -TempDirectoryName $TempDirectoryName
}

# 删除本轮解压使用的临时目录。
function Remove-ExtractionTempDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempDirectoryPath,

        [Parameter(Mandatory = $true)]
        [string]$ParentPath
    )

    if (-not (Test-FileSystemPath -Path $TempDirectoryPath)) {
        return $true
    }

    try {
        Remove-SiblingTempDirectory `
            -TempDirectoryPath $TempDirectoryPath `
            -ParentPath $ParentPath `
            -TempDirectoryName $TempDirectoryName
        return $true
    }
    catch {
        Write-Host "临时目录清理失败: $TempDirectoryPath" -ForegroundColor Yellow
        Write-Host "  原因: $($_.Exception.Message)" -ForegroundColor DarkGray
        return $false
    }
}

# 清空脚本临时解压目录内容，用于外部工具失败后的回退重试。
function Clear-ExtractionTempDirectoryContents {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempDirectoryPath,

        [Parameter(Mandatory = $true)]
        [string]$ParentPath
    )

    if (-not (Test-FileSystemDirectory -Path $TempDirectoryPath)) {
        return
    }

    Clear-SiblingTempDirectoryContents `
        -TempDirectoryPath $TempDirectoryPath `
        -ParentPath $ParentPath `
        -TempDirectoryName $TempDirectoryName
}

# 注册旧 ZIP 文件名编码支持；PowerShell 7/.NET 默认不一定启用 Windows 代码页。
function Initialize-ZipLegacyEncodingProvider {
    if ($script:ZipLegacyEncodingProviderRegistered) {
        return
    }

    try {
        Add-Type -AssemblyName System.Text.Encoding.CodePages -ErrorAction Stop
        [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance)
        $script:ZipLegacyEncodingProviderRegistered = $true
    }
    catch {
        throw "无法加载旧 ZIP 文件名编码支持。$($_.Exception.Message)"
    }
}

# 根据代码页获取 ZIP 文件名编码对象。
function Get-ZipEntryNameEncoding {
    param(
        [Parameter(Mandatory = $true)]
        [int]$CodePage
    )

    Initialize-ZipLegacyEncodingProvider

    try {
        return [System.Text.Encoding]::GetEncoding($CodePage)
    }
    catch {
        throw "无法加载 ZIP 文件名代码页 $CodePage。$($_.Exception.Message)"
    }
}

# 检查默认读取到的 ZIP 条目名是否已经出现替换字符；这通常表示旧中文 ZIP 被按错编码读取。
function Test-ZipArchiveEntryNameHasReplacementCharacter {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ArchivePath
    )

    $zipArchive = $null
    try {
        $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($ArchivePath)
        foreach ($zipEntry in $zipArchive.Entries) {
            if ($zipEntry.FullName.Contains([char]0xFFFD)) {
                return $true
            }
        }

        return $false
    }
    finally {
        if ($null -ne $zipArchive) {
            $zipArchive.Dispose()
        }
    }
}

# 生成 ZIP 解压编码尝试列表；明显乱码时优先用旧中文编码，否则先用默认编码。
function Get-ZipExtractionEncodingPlanList {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ArchivePath
    )

    $defaultEncodingPlan = [pscustomobject]@{
        DisplayName = '默认文件名编码'
        Encoding    = $null
    }

    $legacyEncodingPlanList = [System.Collections.Generic.List[object]]::new()
    foreach ($codePage in $ZipLegacyEntryNameCodePageList) {
        $encoding = Get-ZipEntryNameEncoding -CodePage $codePage
        $encodingDisplayName = switch ($codePage) {
            936 { 'GBK / CP936' }
            default { "$($encoding.WebName) / CP$codePage" }
        }

        $legacyEncodingPlanList.Add([pscustomobject]@{
            DisplayName = $encodingDisplayName
            Encoding    = $encoding
        })
    }

    $hasReplacementCharacter = $false
    try {
        $hasReplacementCharacter = Test-ZipArchiveEntryNameHasReplacementCharacter -ArchivePath $ArchivePath
    }
    catch {
        # 如果预读失败，仍交给正式解压流程给出完整错误。
        $hasReplacementCharacter = $false
    }

    if ($hasReplacementCharacter) {
        return @($legacyEncodingPlanList.ToArray()) + @($defaultEncodingPlan)
    }

    return @($defaultEncodingPlan) + @($legacyEncodingPlanList.ToArray())
}

# 按指定文件名编码执行一次 ZIP 解压。
function Invoke-ZipExtractionAttempt {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ArchivePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [System.Text.Encoding]$EntryNameEncoding
    )

    if ($null -eq $EntryNameEncoding) {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($ArchivePath, $DestinationPath, $false)
        return
    }

    [System.IO.Compression.ZipFile]::ExtractToDirectory($ArchivePath, $DestinationPath, $EntryNameEncoding, $false)
}

# 检查已解压到临时目录中的文件名是否含替换字符；出现时通常说明 ZIP 文件名编码选择错误。
function Test-ZipExtractedItemNameHasReplacementCharacter {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    foreach ($itemPath in [System.IO.Directory]::EnumerateFileSystemEntries($DestinationPath, '*', [System.IO.SearchOption]::AllDirectories)) {
        $itemName = [System.IO.Path]::GetFileName($itemPath)
        if ($itemName.Contains([char]0xFFFD)) {
            return $true
        }
    }

    return $false
}

# 使用 .NET 原生能力解压 zip，并对旧中文 ZIP 文件名编码做回退处理。
function Expand-ZipArchiveToDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ArchivePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,

        [Parameter(Mandatory = $true)]
        [string]$ParentPath
    )

    $encodingPlanList = @(Get-ZipExtractionEncodingPlanList -ArchivePath $ArchivePath)
    $attemptFailureList = [System.Collections.Generic.List[string]]::new()
    $fallbackGarbledEncodingPlan = $null
    $attemptIndex = 0

    foreach ($encodingPlan in $encodingPlanList) {
        $attemptIndex++
        if ($attemptIndex -gt 1) {
            Clear-ExtractionTempDirectoryContents -TempDirectoryPath $DestinationPath -ParentPath $ParentPath
        }

        try {
            Invoke-ZipExtractionAttempt `
                -ArchivePath $ArchivePath `
                -DestinationPath $DestinationPath `
                -EntryNameEncoding $encodingPlan.Encoding
            if (Test-ZipExtractedItemNameHasReplacementCharacter -DestinationPath $DestinationPath) {
                if ($null -eq $fallbackGarbledEncodingPlan) {
                    $fallbackGarbledEncodingPlan = $encodingPlan
                }

                $attemptFailureList.Add("$($encodingPlan.DisplayName): 可解压，但文件名可能仍有乱码。")
                continue
            }

            return
        }
        catch {
            $attemptFailureList.Add("$($encodingPlan.DisplayName): $($_.Exception.Message)")
        }
    }

    if ($null -ne $fallbackGarbledEncodingPlan) {
        Clear-ExtractionTempDirectoryContents -TempDirectoryPath $DestinationPath -ParentPath $ParentPath
        try {
            Invoke-ZipExtractionAttempt `
                -ArchivePath $ArchivePath `
                -DestinationPath $DestinationPath `
                -EntryNameEncoding $fallbackGarbledEncodingPlan.Encoding
            Write-Host "ZIP 文件名编码无法完全修复，已保留可解压结果: $ArchivePath" -ForegroundColor Yellow
            return
        }
        catch {
            $attemptFailureList.Add("$($fallbackGarbledEncodingPlan.DisplayName) 回退重试: $($_.Exception.Message)")
        }
    }

    throw "ZIP 解压失败，已尝试 $($encodingPlanList.Count) 种文件名编码。$($attemptFailureList -join '；')"
}

# 调用外部解压命令，并把退出码和关键输出整理为异常。
function Invoke-ArchiveExtractorCommand {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Extractor,

        [Parameter(Mandatory = $true)]
        [string[]]$ArgumentList
    )

    try {
        $extractorOutput = @(& $Extractor.Path @ArgumentList 2>&1)
        $extractorExitCode = $LASTEXITCODE
    }
    catch {
        throw "$($Extractor.DisplayName) 调用失败。$($_.Exception.Message)"
    }

    if ($null -eq $extractorExitCode) {
        $extractorExitCode = 0
    }

    if ($extractorExitCode -eq 0) {
        return
    }

    $outputText = (($extractorOutput | Select-Object -Last 6 | ForEach-Object {
                $_.ToString().Trim()
            }) -join ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($outputText)) {
        $outputText = '工具未输出详细错误。'
    }

    throw "$($Extractor.DisplayName) 解压失败，退出码: $extractorExitCode。$outputText"
}

# 使用可用外部工具解压非 zip 压缩包；前一个工具失败时清空临时目录后尝试下一个。
function Expand-ExternalArchiveToDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ArchivePath,

        [Parameter(Mandatory = $true)]
        [string]$Extension,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,

        [Parameter(Mandatory = $true)]
        [string]$ParentPath
    )

    $archiveExtractorList = @(Get-ArchiveExtractorList -Extension $Extension)
    if ($archiveExtractorList.Count -eq 0) {
        throw "未找到可用的 $Extension 解压工具。请确认 Windows 内置 tar.exe 可用，或安装 7-Zip、WinRAR 并将对应命令加入 PATH。RAR 还可安装 UnRAR。"
    }

    $attemptFailureList = [System.Collections.Generic.List[string]]::new()
    $attemptIndex = 0
    foreach ($archiveExtractor in $archiveExtractorList) {
        $attemptIndex++
        if ($attemptIndex -gt 1) {
            Clear-ExtractionTempDirectoryContents -TempDirectoryPath $DestinationPath -ParentPath $ParentPath
        }

        try {
            $destinationPathWithSeparator = Add-TrailingDirectorySeparator -DirectoryPath $DestinationPath
            switch ($archiveExtractor.Kind) {
                'WindowsTar' {
                    Invoke-ArchiveExtractorCommand `
                        -Extractor $archiveExtractor `
                        -ArgumentList @('-xf', $ArchivePath, '-C', $DestinationPath)
                }
                'SevenZip' {
                    Invoke-ArchiveExtractorCommand `
                        -Extractor $archiveExtractor `
                        -ArgumentList @('x', '-y', '-p-', '-bso0', '-bsp0', '-bse2', "-o$DestinationPath", $ArchivePath)
                }
                'UnRar' {
                    Invoke-ArchiveExtractorCommand `
                        -Extractor $archiveExtractor `
                        -ArgumentList @('x', '-y', '-p-', '-idq', $ArchivePath, $destinationPathWithSeparator)
                }
                'WinRAR' {
                    Invoke-ArchiveExtractorCommand `
                        -Extractor $archiveExtractor `
                        -ArgumentList @('x', '-ibck', '-y', '-p-', '-idq', $ArchivePath, $destinationPathWithSeparator)
                }
                default {
                    throw "不支持的解压器类型: $($archiveExtractor.Kind)"
                }
            }

            return
        }
        catch {
            $attemptFailureList.Add("$($archiveExtractor.DisplayName): $($_.Exception.Message)")
        }
    }

    throw "$Extension 解压失败，已尝试 $($archiveExtractorList.Count) 个工具。$($attemptFailureList -join '；')"
}

# 按压缩包扩展名选择解压实现。
function Expand-ArchiveFileToDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [object]$PlanItem,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    switch ($PlanItem.Extension) {
        '.zip' {
            Expand-ZipArchiveToDirectory -ArchivePath $PlanItem.ArchivePath -DestinationPath $DestinationPath -ParentPath $PlanItem.ParentPath
        }
        '.rar' {
            Expand-ExternalArchiveToDirectory -ArchivePath $PlanItem.ArchivePath -Extension $PlanItem.Extension -DestinationPath $DestinationPath -ParentPath $PlanItem.ParentPath
        }
        '.7z' {
            Expand-ExternalArchiveToDirectory -ArchivePath $PlanItem.ArchivePath -Extension $PlanItem.Extension -DestinationPath $DestinationPath -ParentPath $PlanItem.ParentPath
        }
        default {
            throw "不支持的压缩包格式: $($PlanItem.Extension)"
        }
    }
}

# 执行解压计划，并按成功、跳过、失败分类汇总。
function Invoke-ArchiveExtraction {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ExtractionPlanList
    )

    if ($ExtractionPlanList.Count -eq 0) {
        Write-Host "没有需要解压的压缩包。" -ForegroundColor Green
        return
    }

    # 三类结果分别记录，确保单个压缩包失败不会中断批量处理。
    $successList = [System.Collections.Generic.List[object]]::new()
    $skippedList = [System.Collections.Generic.List[object]]::new()
    $failedList = [System.Collections.Generic.List[object]]::new()
    $runtimeReservedTargetPathSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $processedCount = 0
    $lastPercent = -1

    foreach ($planItem in $ExtractionPlanList) {
        $processedCount++
        Write-ProgressBar `
            -Activity '批量解压' `
            -Status '正在处理压缩包' `
            -ProcessedCount $processedCount `
            -TotalCount $ExtractionPlanList.Count `
            -LastPercent ([ref]$lastPercent)

        if (-not (Test-FileSystemFile -Path $planItem.ArchivePath)) {
            $skippedList.Add([pscustomobject]@{
                    ArchivePath = $planItem.ArchivePath
                    Message     = '压缩包已不存在。'
                })
            continue
        }

        $targetPath = $planItem.TargetPath
        $targetKey = Get-DirectoryKey -DirectoryPath $targetPath
        if ((Test-FileSystemPath -Path $targetPath) -or -not $runtimeReservedTargetPathSet.Add($targetKey)) {
            $targetPath = Get-UniqueDirectoryPath `
                -ParentPath $planItem.ParentPath `
                -BaseName $planItem.TargetBaseName `
                -ReservedPathSet $runtimeReservedTargetPathSet
        }

        $tempDirectoryPath = $null
        try {
            $tempDirectoryPath = Initialize-ExtractionTempDirectory -ParentPath $planItem.ParentPath
            Expand-ArchiveFileToDirectory -PlanItem $planItem -DestinationPath $tempDirectoryPath

            if (Test-FileSystemPath -Path $targetPath) {
                $targetPath = Get-UniqueDirectoryPath `
                    -ParentPath $planItem.ParentPath `
                    -BaseName $planItem.TargetBaseName `
                    -ReservedPathSet $runtimeReservedTargetPathSet
            }

            [System.IO.Directory]::Move($tempDirectoryPath, $targetPath)
            $tempDirectoryPath = $null

            if (-not (Test-FileSystemDirectory -Path $targetPath)) {
                throw "移动完成后未找到目标文件夹: $targetPath"
            }

            $successList.Add([pscustomobject]@{
                    ArchivePath = $planItem.ArchivePath
                    TargetPath  = $targetPath
                })
        }
        catch {
            $failedList.Add([pscustomobject]@{
                    ArchivePath = $planItem.ArchivePath
                    TargetPath  = $targetPath
                    Message     = $_.Exception.Message
                })
        }
        finally {
            if (-not [string]::IsNullOrWhiteSpace($tempDirectoryPath)) {
                [void](Remove-ExtractionTempDirectory -TempDirectoryPath $tempDirectoryPath -ParentPath $planItem.ParentPath)
            }
        }
    }

    Write-ProgressBar `
        -Activity '批量解压' `
        -Status '批量解压完成' `
        -ProcessedCount $ExtractionPlanList.Count `
        -TotalCount $ExtractionPlanList.Count `
        -LastPercent ([ref]$lastPercent) `
        -Force
    Complete-DynamicStatusLine

    if ($successList.Count -gt 0) {
        Write-Host
        Write-Host "已解压:" -ForegroundColor Green
        foreach ($successItem in $successList) {
            Write-Host "  $($successItem.ArchivePath)" -ForegroundColor White
            Write-Host "    -> $($successItem.TargetPath)" -ForegroundColor Green
        }
    }

    if ($skippedList.Count -gt 0) {
        Write-Host
        Write-Host "跳过项目:" -ForegroundColor Yellow
        foreach ($skippedItem in $skippedList) {
            Write-Host "  $($skippedItem.ArchivePath)" -ForegroundColor Yellow
            Write-Host "    原因: $($skippedItem.Message)" -ForegroundColor DarkGray
        }
    }

    if ($failedList.Count -gt 0) {
        Write-Host
        Write-Host "解压失败:" -ForegroundColor Red
        foreach ($failedItem in $failedList) {
            Write-Host "  $($failedItem.ArchivePath)" -ForegroundColor Red
            Write-Host "    目标: $($failedItem.TargetPath)" -ForegroundColor DarkGray
            Write-Host "    原因: $($failedItem.Message)" -ForegroundColor DarkGray
        }
    }

    Write-Host
    Write-Host -NoNewline "解压完成。成功 " -ForegroundColor White
    Write-Host -NoNewline $successList.Count -ForegroundColor Green
    Write-Host -NoNewline "，跳过 " -ForegroundColor White
    Write-Host -NoNewline $skippedList.Count -ForegroundColor Yellow
    Write-Host -NoNewline "，失败 " -ForegroundColor White
    Write-Host $failedList.Count -ForegroundColor $(if ($failedList.Count -gt 0) { 'Red' } else { 'Green' })

    return [pscustomobject]@{
        SuccessList = $successList.ToArray()
        SkippedList = $skippedList.ToArray()
        FailedList  = $failedList.ToArray()
    }
}

# 根据成功解压结果生成原压缩包删除候选列表。
function Get-ArchiveDeletionCandidateList {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ExtractionSuccessList
    )

    $candidateList = [System.Collections.Generic.List[object]]::new()
    foreach ($successItem in $ExtractionSuccessList) {
        $candidateList.Add([pscustomobject]@{
                ArchivePath = $successItem.ArchivePath
                TargetPath  = $successItem.TargetPath
            })
    }

    return $candidateList.ToArray()
}

# 输出原压缩包删除预览。
function Write-ArchiveDeletionPreview {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ArchiveDeletionCandidateList
    )

    if ($ArchiveDeletionCandidateList.Count -eq 0) {
        Write-Host "没有成功解压的原压缩包可删除。" -ForegroundColor Green
        return
    }

    Write-Host
    Write-Host "原压缩包删除预览:" -ForegroundColor Yellow
    Write-PreviewSeparator -NoLeadingBlank

    $index = 0
    foreach ($candidate in $ArchiveDeletionCandidateList) {
        $index++
        Write-Host -NoNewline ("{0,4}. " -f $index) -ForegroundColor Cyan
        Write-Host $candidate.ArchivePath -ForegroundColor White
        Write-Host "      解压目录: $($candidate.TargetPath)" -ForegroundColor DarkGray
    }

    Write-PreviewSeparator -NoLeadingBlank
    Write-Host
    Write-Host -NoNewline "成功解压的原压缩包数: " -ForegroundColor White
    Write-Host $ArchiveDeletionCandidateList.Count -ForegroundColor Magenta
    Write-Host "删除操作不会进入回收站；删除前会再次确认解压目录仍存在。" -ForegroundColor Yellow
}

# 读取是否删除原压缩包的选择。
function Read-ArchiveDeletionChoice {
    $choice = Read-MenuChoice -Title '是否删除成功解压的原压缩包?' -EndOfInputChoice '0' -MenuOptionList @(
        [pscustomobject]@{ Value = '1'; Label = '删除原压缩包' }
        [pscustomobject]@{ Value = '0'; Label = '保留原压缩包' }
    )

    switch ($choice) {
        '1' { return 'Delete' }
        '0' { return 'Keep' }
    }
}

# 删除成功解压后的原压缩包，并在删除前重新检查解压目录与原文件。
function Invoke-ArchiveDeletion {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ArchiveDeletionCandidateList
    )

    if ($ArchiveDeletionCandidateList.Count -eq 0) {
        Write-Host "没有成功解压的原压缩包可删除。" -ForegroundColor Green
        return
    }

    # 删除也按已删、跳过、失败三类统计，便于用户核对危险操作结果。
    $deletedArchiveList = [System.Collections.Generic.List[string]]::new()
    $skippedArchiveList = [System.Collections.Generic.List[object]]::new()
    $failedArchiveList = [System.Collections.Generic.List[object]]::new()
    $processedCount = 0
    $lastPercent = -1

    foreach ($candidate in $ArchiveDeletionCandidateList) {
        $processedCount++
        Write-ProgressBar `
            -Activity '原压缩包删除' `
            -Status '正在检查并删除原压缩包' `
            -ProcessedCount $processedCount `
            -TotalCount $ArchiveDeletionCandidateList.Count `
            -LastPercent ([ref]$lastPercent)

        if (-not (Test-FileSystemDirectory -Path $candidate.TargetPath)) {
            $skippedArchiveList.Add([pscustomobject]@{
                    ArchivePath = $candidate.ArchivePath
                    Message     = '解压目标目录已不存在，为避免误删，保留原压缩包。'
                })
            continue
        }

        if (-not (Test-FileSystemFile -Path $candidate.ArchivePath)) {
            $skippedArchiveList.Add([pscustomobject]@{
                    ArchivePath = $candidate.ArchivePath
                    Message     = '原压缩包已不存在。'
                })
            continue
        }

        try {
            Remove-FileSystemItem -Path $candidate.ArchivePath
            $deletedArchiveList.Add($candidate.ArchivePath)
        }
        catch {
            $failedArchiveList.Add([pscustomobject]@{
                    ArchivePath = $candidate.ArchivePath
                    Message     = $_.Exception.Message
                })
        }
    }

    Write-ProgressBar `
        -Activity '原压缩包删除' `
        -Status '原压缩包删除完成' `
        -ProcessedCount $ArchiveDeletionCandidateList.Count `
        -TotalCount $ArchiveDeletionCandidateList.Count `
        -LastPercent ([ref]$lastPercent) `
        -Force
    Complete-DynamicStatusLine

    if ($deletedArchiveList.Count -gt 0) {
        Write-Host
        Write-Host "已删除原压缩包:" -ForegroundColor Magenta
        foreach ($deletedArchivePath in $deletedArchiveList) {
            Write-Host "  $deletedArchivePath" -ForegroundColor Magenta
        }
    }

    if ($skippedArchiveList.Count -gt 0) {
        Write-Host
        Write-Host "跳过删除:" -ForegroundColor Yellow
        foreach ($skippedArchive in $skippedArchiveList) {
            Write-Host "  $($skippedArchive.ArchivePath)" -ForegroundColor Yellow
            Write-Host "    原因: $($skippedArchive.Message)" -ForegroundColor DarkGray
        }
    }

    if ($failedArchiveList.Count -gt 0) {
        Write-Host
        Write-Host "原压缩包删除失败:" -ForegroundColor Red
        foreach ($failedArchive in $failedArchiveList) {
            Write-Host "  $($failedArchive.ArchivePath)" -ForegroundColor Red
            Write-Host "    原因: $($failedArchive.Message)" -ForegroundColor DarkGray
        }
    }

    Write-Host
    Write-Host -NoNewline "原压缩包删除完成。已删除 " -ForegroundColor White
    Write-Host -NoNewline $deletedArchiveList.Count -ForegroundColor Magenta
    Write-Host -NoNewline "，跳过 " -ForegroundColor White
    Write-Host -NoNewline $skippedArchiveList.Count -ForegroundColor Yellow
    Write-Host -NoNewline "，失败 " -ForegroundColor White
    Write-Host $failedArchiveList.Count -ForegroundColor $(if ($failedArchiveList.Count -gt 0) { 'Red' } else { 'Green' })
}

# 解压完成后处理原压缩包删除流程。
function Invoke-PostExtractionArchiveDeletion {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ExtractionSuccessList
    )

    $archiveDeletionCandidateList = @(Get-ArchiveDeletionCandidateList -ExtractionSuccessList $ExtractionSuccessList)
    if ($archiveDeletionCandidateList.Count -eq 0) {
        Write-Host "没有成功解压的原压缩包可删除。" -ForegroundColor Green
        return
    }

    if ($AssumeYes) {
        Write-Host
        Write-Host "已启用 -yes，将跳过删除确认并默认删除成功解压的原压缩包。" -ForegroundColor Yellow
        if (Wait-DangerousOperationGracePeriod `
                -WarningMessage '危险操作: 将删除成功解压的原压缩包。' `
                -AdditionalWarningMessage "计划删除原压缩包数: $($archiveDeletionCandidateList.Count)。删除操作不会进入回收站。" `
                -CancelledMessage '已取消默认删除，原压缩包已保留。' `
                -CompletedMessage '倒计时结束，开始删除原压缩包。' `
                -CountdownMessageFormat '倒计时 {0} 秒后删除原压缩包，按 Enter 取消...') {
            Invoke-ArchiveDeletion -ArchiveDeletionCandidateList $archiveDeletionCandidateList
        }
        return
    }

    Write-ArchiveDeletionPreview -ArchiveDeletionCandidateList $archiveDeletionCandidateList
    $deletionChoice = Read-ArchiveDeletionChoice
    if ($deletionChoice -eq 'Delete') {
        Invoke-ArchiveDeletion -ArchiveDeletionCandidateList $archiveDeletionCandidateList
        return
    }

    Write-Host "已保留原压缩包。" -ForegroundColor Green
}

# 串联解压与解压后的原压缩包删除流程。
function Invoke-ExtractionWorkflow {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ExtractionPlanList
    )

    $extractionResult = Invoke-ArchiveExtraction -ExtractionPlanList $ExtractionPlanList
    if ($null -eq $extractionResult) {
        return
    }

    Invoke-PostExtractionArchiveDeletion -ExtractionSuccessList @($extractionResult.SuccessList)
}

# ========== 主逻辑 ==========

if ($Help) {
    Show-HelpText
    exit 0
}

if ($null -eq $PathList -or $PathList.Count -eq 0) {
    $ResolvedInputItemList = @(Read-InteractivePathList)
}
else {
    $EffectivePathList = @($PathList | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($EffectivePathList.Count -eq 0) {
        Write-Host "未提供有效路径，已退出。" -ForegroundColor Yellow
        exit 0
    }

    $ResolvedInputResult = Resolve-InputPathList -PathList $EffectivePathList
    if (-not $ResolvedInputResult.Success) {
        Write-Host "$($ResolvedInputResult.Error) 本次输入不保留。" -ForegroundColor Red
        exit 1
    }

    $ResolvedInputItemList = @($ResolvedInputResult.Items)
}

if ($ResolvedInputItemList.Count -eq 0) {
    Write-Host "未提供新的有效路径，已退出。" -ForegroundColor Yellow
    exit 0
}

Write-StageMessage "开始生成解压计划，输入路径数: $($ResolvedInputItemList.Count)"
$ExtractionPlanResult = New-ArchiveExtractionPlan -InputItemList $ResolvedInputItemList
$ExtractionPlanList = @($ExtractionPlanResult.Items)

Write-DeferredScanWarningList -WarningList $ExtractionPlanResult.Warnings -Title '压缩包扫描跳过汇总'

if ($ExtractionPlanList.Count -eq 0) {
    Write-Host
    Write-Host "扫描完成，未发现可解压的 .zip / .rar / .7z 压缩包。" -ForegroundColor Green
    exit 0
}

Write-ArchiveExtractorStatus -ExtractionPlanList $ExtractionPlanList

if ($AssumeYes) {
    Write-Host
    Write-Host "已启用 -yes，将跳过解压预览和删除确认，执行默认解压和默认删除。" -ForegroundColor Yellow
    Write-Host -NoNewline "计划解压压缩包数: " -ForegroundColor White
    Write-Host -NoNewline $ExtractionPlanList.Count -ForegroundColor Magenta
    Write-Host -NoNewline "，已去重: " -ForegroundColor White
    Write-Host $ExtractionPlanResult.DuplicateCount -ForegroundColor Yellow

    Invoke-ExtractionWorkflow -ExtractionPlanList $ExtractionPlanList
    exit 0
}

Write-ExtractionPreview -ExtractionPlanList $ExtractionPlanList -DuplicateCount $ExtractionPlanResult.DuplicateCount
$OperationChoice = Read-OperationChoice

switch ($OperationChoice) {
    'Default' {
        Invoke-ExtractionWorkflow -ExtractionPlanList $ExtractionPlanList
    }
    'Manual' {
        $SelectedExtractionList = @(Read-ManualExtractionSelection -ExtractionPlanList $ExtractionPlanList)
        if ($SelectedExtractionList.Count -gt 0) {
            Invoke-ExtractionWorkflow -ExtractionPlanList $SelectedExtractionList
        }
    }
    'Exit' {
        Write-Host "已退出，未解压任何压缩包。" -ForegroundColor Yellow
    }
}
