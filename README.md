# Scripts

个人跨平台脚本仓库，主要按平台和分享方式组织。

## 目录结构

| 目录 | 说明 |
| --- | --- |
| `Windows` | Windows 平台 PowerShell 7 脚本；脚本共用 `common.psm1`。 |
| `Share` | 可单文件分享给其他电脑使用的脚本；当前包含 PowerShell 5.1 兼容版重复文件清理脚本。 |
| `macOS` | macOS / zsh 自定义函数和公共工具函数。 |

## Windows 脚本

Windows 脚本说明见 [Windows/README.md](Windows/README.md)。

当前包含：

- `Convert-OfficeFiles.ps1`：调用本机 Office 原生应用，把 `.doc`、`.xls`、`.ppt` 转换为 `.docx`、`.xlsx`、`.pptx`。
- `Remove-DuplicateFiles.ps1`：查找并删除重复文件，支持单目录、多目录合并和参考目录模式。
- `Remove-EmptyFolders.ps1`：删除输入目录下的空子文件夹。
- `common.psm1`：Windows 脚本公共模块，提供动态进度条、延迟警告输出、路径输入解析等工具函数。

## 共享脚本

`Share` 目录中的脚本用于复制到其他电脑独立使用，不依赖 `Windows/common.psm1`。

当前包含：

- `Remove-DuplicateFiles.ps1`：PowerShell 5.1 兼容版重复文件清理脚本。

## macOS 脚本

`macOS/function.zsh` 是自定义函数入口，会自动加载同目录的 `common.zsh`。通常在 zsh 配置文件中 source `function.zsh` 即可。
