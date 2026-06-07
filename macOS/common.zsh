# 通用 zsh 工具函数。
# 约定：需要返回数组的内部函数统一写入 zsh 约定变量 reply；
# 调用方如果还会继续调用其他 reply 风格函数，应立即复制到局部数组。
# 所有函数都用 emulate -L zsh 隔离选项，避免用户 zshrc 中的别名、setopt 影响脚本行为。

# ========== 彩色输出 ==========

# 判断当前输出目标是否适合使用颜色。
_common_should_color() {
    emulate -L zsh -o typeset_silent

    # 遵守通用 NO_COLOR 约定；CLICOLOR_FORCE 可在非 tty 场景强制彩色输出。
    [[ -n "${NO_COLOR:-}" ]] && return 1
    [[ "${TERM:-dumb}" == "dumb" ]] && return 1
    [[ -t 1 || ( -n "${CLICOLOR_FORCE:-}" && "${CLICOLOR_FORCE}" != "0" ) ]]
}

# 输出带颜色的标签；正文使用普通 print，避免文件名中的 % 被 prompt expansion 误解析。
msg_label() {
    emulate -L zsh -o typeset_silent

    # 颜色和标签单独用 prompt expansion；正文用普通 print，避免文件名里的 % 被当成格式语义。
    local color_name="$1"
    local label_text="$2"
    shift 2

    if _common_should_color; then
        print -Pn "%B%F{${color_name}}${label_text}:%f%b "
    else
        print -rn -- "${label_text}: "
    fi
    print -r -- "$*"
}

# 错误信息写入 stderr。
msg_error() {
    emulate -L zsh -o typeset_silent

    msg_label red "错误" "$@" >&2
}

# 警告信息写入 stderr。
msg_warn() {
    emulate -L zsh -o typeset_silent

    msg_label yellow "警告" "$@" >&2
}

# 普通提示信息写入 stdout。
msg_info() {
    emulate -L zsh -o typeset_silent

    msg_label blue "信息" "$@"
}

# 阶段进度信息写入 stdout。
msg_progress() {
    emulate -L zsh -o typeset_silent

    msg_label cyan "进度" "$@"
}

# 成功结果写入 stdout。
msg_success() {
    emulate -L zsh -o typeset_silent

    msg_label green "完成" "$@"
}

# ========== 路径与扫描 ==========

# 收集 find -print0 结果到 reply。
# 使用临时文件承接 find 输出，避免管道 while 的子 shell 问题，同时保留 find 的退出状态。
_common_collect_find_results() {
    emulate -L zsh -o typeset_silent

    reply=()
    local temp_file=""
    local -i find_status=0

    if ! temp_file=$(command mktemp "${TMPDIR:-/tmp}/common-find.XXXXXX"); then
        msg_error "创建临时文件失败，无法收集扫描结果。"
        return 1
    fi

    {
        if command find "$@" -print0 > "$temp_file"; then
            local found_path
            # -print0 / read -d '' 能安全处理空格、换行和其他特殊字符路径。
            while IFS= read -r -d '' found_path; do
                reply+=("$found_path")
            done < "$temp_file"
        else
            find_status=1
        fi
    } always {
        if [[ -n "$temp_file" && -e "$temp_file" ]] && ! command rm -f -- "$temp_file"; then
            msg_warn "临时文件清理失败: \"$temp_file\""
        fi
    }

    return $find_status
}

# 将扩展名列表转换为 find 可用的 -iname 条件，结果写入 reply。
_common_build_extension_find_args() {
    emulate -L zsh -o typeset_silent

    reply=()
    local extension
    for extension in "$@"; do
        extension="${extension:l}"
        [[ -z "$extension" ]] && continue
        # 生成形如: -iname '*.jpg' -o -iname '*.png'，由调用方负责包进括号。
        [[ ${#reply[@]} -gt 0 ]] && reply+=("-o")
        reply+=("-iname" "*.${extension}")
    done
}

# 判断路径扩展名是否在允许列表中；未传允许列表时表示不限制扩展名。
_common_path_has_extension() {
    emulate -L zsh -o typeset_silent

    local file_path="$1"
    shift

    # 没传允许列表时表示“不限制扩展名”，便于扫描函数复用同一套逻辑。
    [[ $# -eq 0 ]] && return 0

    local file_ext="$(_common_path_lower_extension "$file_path")"
    [[ -n "$file_ext" ]] || return 1

    local extension
    for extension in "$@"; do
        [[ "$file_ext" == "${extension:l}" ]] && return 0
    done

    return 1
}

# 对路径列表按字典序排序并去重，结果写入 reply。
_common_sort_unique_paths() {
    emulate -L zsh -o typeset_silent

    local -a input_paths=("$@")
    # (@uon)：按字典序排序并去重，保证后续处理顺序稳定。
    reply=("${(@uon)input_paths}")
}

# 扫描目录并按需过滤扩展名，结果写入 reply。
# scan_scope:
# - current: 仅当前层文件。
# - recursive: 当前目录及所有子目录文件。
# - children: 仅子目录层级文件，不包含根目录当前层文件。
_common_collect_directory_files() {
    emulate -L zsh -o typeset_silent

    local directory_path="$1"
    local scan_scope="$2"
    shift 2

    local -a find_args=("$directory_path")
    # scan_scope 集中控制 find 的深度，避免各业务函数手写略有差异的扫描参数。
    case "$scan_scope" in
        current)
            find_args+=("-mindepth" "1" "-maxdepth" "1" "-type" "f")
            ;;
        recursive)
            find_args+=("-type" "f")
            ;;
        children)
            find_args+=("-mindepth" "2" "-type" "f")
            ;;
        *)
            msg_error "未知扫描范围: $scan_scope"
            reply=()
            return 1
            ;;
    esac

    if [[ $# -gt 0 ]]; then
        _common_build_extension_find_args "$@"
        local -a extension_find_args=("${reply[@]}")
        if [[ ${#extension_find_args[@]} -gt 0 ]]; then
            # find "$dir" ... '(' -iname '*.jpg' -o -iname '*.png' ')'。
            find_args+=("(" "${extension_find_args[@]}" ")")
        fi
    fi

    if ! _common_collect_find_results "${find_args[@]}"; then
        return 1
    fi

    local -a found_files=("${reply[@]}")
    _common_sort_unique_paths "${found_files[@]}"
}

# 从输入参数收集文件，结果写入 reply。
# 文件参数会直接校验扩展名；目录参数会按 scan_scope 扫描后再过滤。
_common_collect_input_files() {
    emulate -L zsh -o typeset_silent

    local scan_scope="$1"
    local unsupported_kind="$2"
    shift 2

    # 参数格式: scan_scope unsupported_kind ext... -- input...
    # 用 -- 分隔扩展名和实际路径，避免路径刚好像扩展名时产生歧义。
    local -a allowed_extensions=()
    while [[ $# -gt 0 && "$1" != "--" ]]; do
        allowed_extensions+=("${1:l}")
        shift
    done

    if [[ "${1:-}" != "--" ]]; then
        msg_error "内部错误: 收集输入文件时缺少 -- 分隔符"
        reply=()
        return 1
    fi
    shift

    local -a collected_files=()
    local input_path
    for input_path in "$@"; do
        if [[ -f "$input_path" ]]; then
            # 单文件输入也走扩展名校验，保持和目录扫描结果一致。
            if _common_path_has_extension "$input_path" "${allowed_extensions[@]}"; then
                collected_files+=("$input_path")
            else
                msg_error "不支持的${unsupported_kind} \"$input_path\""
            fi
        elif [[ -d "$input_path" ]]; then
            if _common_collect_directory_files "$input_path" "$scan_scope" "${allowed_extensions[@]}"; then
                # reply 会被后续 helper 覆盖，所以立刻复制到局部数组。
                local -a scanned_files=("${reply[@]}")
                collected_files+=("${scanned_files[@]}")
            else
                msg_error "扫描目录失败: \"$input_path\""
            fi
        else
            msg_error "\"$input_path\" 不是有效的文件或目录，跳过。"
        fi
    done

    _common_sort_unique_paths "${collected_files[@]}"
}

# ========== 命令与文件操作 ==========

# 检查外部命令是否可用。
_common_command_exists() {
    emulate -L zsh -o typeset_silent

    command -v "$1" >/dev/null 2>&1
}

# 获取文件扩展名并转为小写；无扩展名时返回空字符串。
_common_path_lower_extension() {
    emulate -L zsh -o typeset_silent

    # 使用 zsh 的 :e 取最后一段扩展名；例如 .tar.gz 返回 gz。
    print -r -- "${1:t:e:l}"
}

# 生成不冲突文件路径，冲突时在扩展名前追加 01-99；断开的符号链接按空位处理。
_common_next_available_file_path() {
    emulate -L zsh -o typeset_silent

    local directory_path="$1"
    local base_name="$2"
    local file_ext="$3"

    local candidate_path="$directory_path/$base_name"
    [[ -n "$file_ext" ]] && candidate_path="${candidate_path}.${file_ext}"

    # -e 为 false 的断开符号链接会被视为空位；这是为了允许用真实文件替换失效链接。
    if [[ ! -e "$candidate_path" ]]; then
        print -r -- "$candidate_path"
        return 0
    fi

    local -i counter
    local suffix
    for (( counter = 1; counter <= 99; counter++ )); do
        # 固定两位序号，和 rtf/fmv 的文件名规则保持一致。
        suffix=$(printf "%02d" "$counter")
        candidate_path="$directory_path/${base_name}${suffix}"
        [[ -n "$file_ext" ]] && candidate_path="${candidate_path}.${file_ext}"

        if [[ ! -e "$candidate_path" ]]; then
            print -r -- "$candidate_path"
            return 0
        fi
    done

    return 1
}

# 执行文件移动/替换，并统一输出失败原因。
_common_move_file() {
    emulate -L zsh -o typeset_silent

    local source_path="$1"
    local target_path="$2"

    # 允许覆盖目标；只在明确需要替换输出文件的场景使用。
    if command mv -- "$source_path" "$target_path"; then
        return 0
    fi

    msg_error "移动失败: \"$source_path\" -> \"$target_path\""
    return 1
}

# 执行不覆盖现有目标的文件移动；断开的符号链接按空位处理。
_common_move_file_no_clobber() {
    emulate -L zsh -o typeset_silent

    local source_path="$1"
    local target_path="$2"

    # 先用 -e 做业务层检查：真实文件/目录/有效链接都算占用。
    # 断开的符号链接没有实际目标文件，先移除链接本身，再允许真实文件占用这个名字。
    if [[ -L "$target_path" && ! -e "$target_path" ]]; then
        if ! command rm -f -- "$target_path"; then
            msg_error "删除失效符号链接失败: \"$target_path\""
            return 1
        fi
    elif [[ -e "$target_path" ]]; then
        msg_warn "跳过移动: 目标已存在 \"$target_path\""
        return 1
    fi

    # mv -n 是最后一道防线，避免检查和移动之间目标被其他进程创建。
    if ! command mv -n -- "$source_path" "$target_path"; then
        msg_error "移动失败: \"$source_path\" -> \"$target_path\""
        return 1
    fi

    # macOS mv -n 遇到已存在目标时可能返回 0；源文件还在就说明实际未移动。
    if [[ -e "$source_path" ]]; then
        msg_warn "跳过移动: 目标已存在 \"$target_path\""
        return 1
    fi

    return 0
}

# 清理临时文件；失败时只给出警告，不中断主流程。
_common_remove_temp_file() {
    emulate -L zsh -o typeset_silent

    local temp_path="$1"

    # 空路径直接视为已清理，方便 always 块无条件调用。
    [[ -z "$temp_path" || ! -e "$temp_path" ]] && return 0

    if ! command rm -f -- "$temp_path"; then
        msg_warn "临时文件清理失败: \"$temp_path\""
        return 1
    fi

    return 0
}

# 清理临时目录；仅用于 mktemp -d 创建的脚本临时目录。
_common_remove_temp_directory() {
    emulate -L zsh -o typeset_silent

    local temp_dir="$1"

    # 只清理调用方传入的临时目录；不要把业务目录交给这个 helper。
    [[ -z "$temp_dir" || ! -d "$temp_dir" ]] && return 0

    if ! command rm -rf -- "$temp_dir"; then
        msg_warn "临时目录清理失败: \"$temp_dir\""
        return 1
    fi

    return 0
}
