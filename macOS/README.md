# macOS 脚本

这个文件夹用于存放 macOS 上的 zsh 自定义函数和公共工具函数。

`common.zsh` 提供公共工具，`function.zsh` 提供日常使用函数。通常在 `~/.zshrc` 中按顺序 source 两个文件：

```zsh
source /Users/wozsun/Code/Scripts/macOS/common.zsh
source /Users/wozsun/Code/Scripts/macOS/function.zsh
```

`common.zsh` 只放跨函数复用的输出、扫描、路径、移动和临时文件工具，一般不需要在命令行直接调用。

## function.zsh

`function.zsh` 提供媒体文件时间写入、重命名、归类、校验和图片转换函数。

### 参数说明

| 函数 | 用法 | 说明 |
| --- | --- | --- |
| `hp` | `hp` | 显示可用函数总览。 |
| `ete` | `ete [-t hour] file1 [file2 ...]` | 根据文件名时间 `YYYYMMDD_HHMMSS[NN]` 写入 HEIC / MOV 元数据时间标签。 |
| `rtf` | `rtf file_or_directory1 [file_or_directory2 ...]` | 读取拍摄时间并重命名为 `YYYYMMDD_HHMMSS[NN].ext`。 |
| `mtc` | `mtc file_or_directory1 [file_or_directory2 ...]` | 校验 HEIC / MOV 元数据时间与文件名时间，不一致时调用 `ete` 修复。 |
| `ctw` | `ctw [-q 0-100] file_or_directory1 [file_or_directory2 ...]` | 将图片转换为 webp。未传 `-q` 时自动压到 500KB 内。 |
| `cmv` | `cmv directory` | 按文件名日期归类到 `YYYY/MMDD` 两级目录。 |
| `fmv` | `fmv directory` | 将子目录媒体文件提取到根目录后，依次执行 `rtf`、`cmv` 与 `mtc`。 |

每个函数可通过 `函数名 -h` 查看详细帮助。

### 使用示例

写入 HEIC / MOV 时间标签，默认时区为 `+08:00`：

```zsh
ete 20240601_123456.heic
```

使用指定时区写入时间标签：

```zsh
ete -t -8 20240601_123456.mov
```

读取拍摄时间并重命名当前目录文件：

```zsh
rtf /Users/name/Pictures/Import
```

校验已归类目录中的 HEIC / MOV 时间：

```zsh
mtc /Users/name/Pictures/Import
```

将图片转换为 webp：

```zsh
ctw /Users/name/Pictures/Import
```

手动指定 webp 质量，不限制目标体积：

```zsh
ctw -q 85 /Users/name/Pictures/Import
```

按时间文件名归类根目录文件：

```zsh
cmv /Users/name/Pictures/Import
```

提取子目录媒体文件并完成重命名、归类和时间校验：

```zsh
fmv /Users/name/Pictures/Import
```

### 注意事项

- 脚本专为 macOS zsh 准备，使用 zsh 特性和 macOS 自带工具。
- `ete`、`rtf`、`mtc` 和 `fmv` 需要 `exiftool`；未安装时可执行 `brew install exiftool`。
- `ctw` 需要 `cwebp` 和 `sips`；`cwebp` 可通过 `brew install webp` 安装，`sips` 为 macOS 自带工具。
- `rtf` 支持 `heic/jpg/jpeg/dng/cr3/mov/mp4`，目录输入时只处理当前层文件，不递归。
- `mtc` 只校验 `heic/mov`，目录输入时递归扫描。
- `ctw` 支持 `jpg/jpeg/png/tif/tiff/bmp`，目录输入时递归扫描。
- `cmv` 只处理目标目录当前层级中符合 `YYYYMMDD_HHMMSS[NN].ext` 命名的文件；看起来已经是 `YYYY/MMDD` 的归类目录会跳过。
- `fmv` 只提取子目录层级中的 `heic/jpg/jpeg/dng/cr3/mov/mp4`，不会处理根目录已有文件。
- `fmv` 提取后会删除空子目录，这是当前设计。
- `rtf` 和 `fmv` 遇到目标文件名已存在时，会自动追加 `01-99` 两位序号；超过 99 次冲突时跳过。
- 文件移动默认不覆盖已存在的真实文件；断开的符号链接会被移除并替换为真实文件。
