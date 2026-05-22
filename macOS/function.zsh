# macOS 自定义媒体处理函数。

# 自动加载同目录公共工具函数；外部只需要 source 当前文件。
_load_common_functions() {
    emulate -L zsh -o typeset_silent

    # %x 指向当前被 source 的文件，比 %N 更适合在函数内定位脚本路径。
    local script_path="${${(%):-%x}:A}"
    local script_dir="${script_path:h}"
    local common_path="$script_dir/common.zsh"

    if [[ ! -r "$common_path" ]]; then
        print -u2 -- "错误: 未找到公共工具文件: $common_path"
        return 1
    fi

    source "$common_path"
}

if ! _load_common_functions; then
    unfunction _load_common_functions 2>/dev/null
    return 1 2>/dev/null || exit 1
fi
unfunction _load_common_functions 2>/dev/null

# ========== 帮助总览 ==========

hp() {
    emulate -L zsh -o typeset_silent

    cat <<'EOF'

可用函数列表：

- ete [-t hour] file1 [file2 ...]
  根据文件名（YYYYMMDD_HHMMSS）写入媒体时间标签（heic/mov）

- rtf file_or_directory1 [file_or_directory2 ...]
  读取拍摄时间并重命名文件为 YYYYMMDD_HHMMSS[NN].ext

- mtc file_or_directory1 [file_or_directory2 ...]
  同时校验媒体时间标签与文件名时间一致性，不一致时调用 ete 修复（heic/mov）

- ctw [-q 0-100] file_or_directory1 [file_or_directory2 ...]
  将图片转换为 webp；默认自动压到 500KB 内，使用 -q 时不限制体积

- cmv directory
  按文件名日期（YYYYMMDD）归类到对应子目录

- fmv directory
  将子目录图片/视频提取到根目录后，依次执行 rtf、cmv 与 mtc

查看详细帮助:
  函数名 -h

EOF
}

# ========== 媒体时间处理 ==========

# 函数: ete
ete() {
    emulate -L zsh -o typeset_silent

    if [[ "${1:-}" == "-h" ]]; then
        cat <<'EOF'
用法:
  ete [-t hour] file1 [file2 ...]

功能:
  根据文件名中的时间（YYYYMMDD_HHMMSS[NN]）写入媒体元数据时间标签。
  支持 heic / mov。

参数:
  -t hour  写入时区，默认 8；例如 -t 8 表示 +08:00，-t -8 表示 -08:00。

行为说明:
  heic:
    写入 CreateDate / ModifyDate / DateTimeOriginal，并写入 OffsetTime* 时区标签。
  mov:
    以 QuickTime UTC 语义写入 CreateDate / ModifyDate / CreationDate / Media* / Track*。
EOF
        return 0
    fi

    local time_zone_hour="8"
    local -a input_files=()

    # 将 -t 的小时数规范化为 exiftool 需要的 ±HH:00。
    while [[ $# -gt 0 ]]; do
        case "$1" in
            -t)
                if [[ $# -lt 2 ]]; then
                    msg_error "参数 -t 缺少时区小时数，例如: -t 8"
                    return 1
                fi
                time_zone_hour="$2"
                shift 2
                ;;
            -*)
                msg_error "未知参数: $1"
                msg_error "用法: ete [-t hour] file1 [file2 ...]"
                return 1
                ;;
            *)
                input_files+=("$1")
                shift
                ;;
        esac
    done

    if [[ ${#input_files[@]} -lt 1 ]]; then
        msg_error "用法: ete [-t hour] file1 [file2 ...]"
        return 1
    fi

    local time_zone_text="$time_zone_hour"
    local time_zone_sign="+"
    if [[ "$time_zone_text" == -* ]]; then
        time_zone_sign="-"
        time_zone_text="${time_zone_text#-}"
    elif [[ "$time_zone_text" == +* ]]; then
        time_zone_text="${time_zone_text#+}"
    fi

    if [[ -z "$time_zone_text" || "$time_zone_text" == *[!0-9]* ]]; then
        msg_error "时区小时数无效: $time_zone_hour"
        msg_error "示例: -t 8、-t 7、-t -8"
        return 1
    fi

    local -i time_zone_abs=$(( 10#$time_zone_text ))
    if (( time_zone_abs > 14 )); then
        msg_error "时区小时数超出范围: $time_zone_hour"
        return 1
    fi

    local time_zone="${time_zone_sign}$(printf '%02d' "$time_zone_abs"):00"

    if ! _common_command_exists exiftool; then
        msg_error "未检测到 exiftool。请先执行: brew install exiftool"
        return 1
    fi

    local -i failed_count=0

    for file_path in "${input_files[@]}"; do
        if [[ ! -f "$file_path" ]]; then
            msg_error "文件 \"$file_path\" 不存在，跳过。"
            ((failed_count++))
            continue
        fi

        local file_name="${file_path:t}"
        local file_ext="$(_common_path_lower_extension "$file_path")"
        local file_base="${file_path:t:r:l}"

        # ete 允许文件名带 rtf 追加序号（如 YYYYMMDD_HHMMSS01），解析时仅取前 15 位时间主体
        if ! _media_is_timestamp_name "$file_base"; then
            msg_error "文件名 \"$file_name\" 不符合 YYYYMMDD_HHMMSS[NN]，跳过。"
            ((failed_count++))
            continue
        fi

        local base_time="$(_media_timestamp_body "$file_base")"
        local formatted_time=$(date -j -f "%Y%m%d_%H%M%S" "$base_time" +"%Y:%m:%d %H:%M:%S" 2>/dev/null)
        if [[ -z "$formatted_time" ]]; then
            msg_error "无法从文件名 \"$file_name\" 解析时间，跳过。"
            ((failed_count++))
            continue
        fi

        case "$file_ext" in
            heic)
                msg_progress "检测到 HEIC 文件，正在修改对应时间标签..."
                if ! exiftool \
                    -CreateDate="$formatted_time" \
                    -ModifyDate="$formatted_time" \
                    -DateTimeOriginal="$formatted_time" \
                    -OffsetTime="$time_zone" \
                    -OffsetTimeOriginal="$time_zone" \
                    -OffsetTimeDigitized="$time_zone" \
                    -overwrite_original \
                    "$file_path"; then
                    msg_error "修改时间标签失败: \"$file_path\""
                    ((failed_count++))
                    continue
                fi
                msg_info "修改结果如下："
                exiftool \
                    -m \
                    -fast \
                    -OffsetTime \
                    -CreateDate \
                    -ModifyDate \
                    -DateTimeOriginal \
                    "$file_path"
                ;;
            mov)
                local adjusted_time="${formatted_time}${time_zone}"
                msg_progress "检测到 MOV 文件，正在修改对应时间标签..."
                if ! exiftool \
                    -api QuickTimeUTC=1 \
                    -overwrite_original \
                    -CreateDate="$adjusted_time" \
                    -ModifyDate="$adjusted_time" \
                    -CreationDate="$adjusted_time" \
                    -MediaCreateDate="$adjusted_time" \
                    -MediaModifyDate="$adjusted_time" \
                    -TrackCreateDate="$adjusted_time" \
                    -TrackModifyDate="$adjusted_time" \
                    "$file_path"; then
                    msg_error "修改时间标签失败: \"$file_path\""
                    ((failed_count++))
                    continue
                fi
                msg_info "修改结果如下："
                exiftool \
                    -m \
                    -fast \
                    -api QuickTimeUTC=1 \
                    -OffsetTime \
                    -CreateDate \
                    -ModifyDate \
                    -CreationDate \
                    -MediaCreateDate \
                    -MediaModifyDate \
                    -TrackCreateDate \
                    -TrackModifyDate \
                    "$file_path"
                ;;
            *)
                msg_error "不支持的文件类型 \"$file_ext\""
                ((failed_count++))
                ;;
        esac
    done

    (( failed_count == 0 ))
}

# 函数: rtf
rtf() {
    emulate -L zsh -o typeset_silent

    if [[ "${1:-}" == "-h" ]]; then
        cat <<'EOF'
用法:
  rtf file_or_directory1 [file_or_directory2 ...]

功能:
  读取文件拍摄时间并重命名为 YYYYMMDD_HHMMSS[NN].ext。

行为说明:
  输入目录时，仅处理该目录当前层文件（不递归）。
  仅处理常见媒体文件: heic/jpg/jpeg/dng/cr3/mov/mp4。
  时间读取优先级: DateTimeOriginal > CreationDate > CreateDate。
  mov 读取时启用 QuickTimeUTC=1。
  文件名冲突时自动追加两位序号（01, 02...）。
EOF
        return 0
    fi

    if [[ $# -lt 1 ]]; then
        msg_error "用法: rtf file_or_directory1 [file_or_directory2 ...]"
        return 1
    fi

    if ! _common_command_exists exiftool; then
        msg_error "未检测到 exiftool。请先执行: brew install exiftool"
        return 1
    fi

    local -a rtf_extensions=(heic jpg jpeg dng cr3 mov mp4)
    _common_collect_input_files current "媒体类型" "${rtf_extensions[@]}" -- "$@"
    local -a sorted_target_files=("${reply[@]}")

    if [[ ${#sorted_target_files[@]} -eq 0 ]]; then
        msg_error "未找到可处理的文件。"
        return 1
    fi

    local -i failed_count=0

    for file_path in "${sorted_target_files[@]}"; do

        local file_name="${file_path:t}"
        local file_ext="$(_common_path_lower_extension "$file_path")"

        # 使用 exiftool 提取时间（单次调用，按优先级取第一个可用值）
        local -a capture_times
        case "$file_ext" in
            mov)
                # MOV 按 QuickTime UTC 语义读取
                capture_times=("${(@f)$(exiftool -m -f -fast -api QuickTimeUTC=1 -s3 \
                    -DateTimeOriginal \
                    -CreationDate \
                    -CreateDate \
                    "$file_path")}")
                ;;
            *)
                capture_times=("${(@f)$(exiftool -m -f -fast -s3 \
                    -DateTimeOriginal \
                    -CreationDate \
                    -CreateDate \
                    "$file_path")}")
                ;;
        esac

        local original_time=""
        local candidate_time
        for candidate_time in "${capture_times[@]}"; do
            if [[ -n "$candidate_time" && "$candidate_time" != "-" ]]; then
                original_time="$candidate_time"
                break
            fi
        done

        if [[ -z "$original_time" ]]; then
            msg_error "无法读取 \"$file_path\" 的拍摄时间，跳过。"
            ((failed_count++))
            continue
        fi

        # 仅用于命名：统一截取前 19 位 YYYY:MM:DD HH:MM:SS（忽略后续时区）
        local name_time="${original_time:0:19}"

        # 将 Date 转换为 yyyymmdd_hhmmss 格式
        local formatted_time=$(date -j -f "%Y:%m:%d %H:%M:%S" "$name_time" +"%Y%m%d_%H%M%S" 2>/dev/null)
        if [[ -z "$formatted_time" ]]; then
            msg_error "时间格式转换失败，跳过 \"$file_path\"。"
            ((failed_count++))
            continue
        fi

        local dir_path="${file_path:h}"
        local base_name="$formatted_time"
        local current_name_part="${file_path:t:r}"
        local new_file_name=""

        if _media_is_timestamp_name "$current_name_part" && [[ "$(_media_timestamp_body "$current_name_part")" == "$formatted_time" ]]; then
            msg_info "跳过: \"$file_path\" 已符合目标命名。"
            continue
        fi

        if ! new_file_name="$(_common_next_available_file_path "$dir_path" "$base_name" "$file_ext")"; then
            msg_error "\"$file_path\" 的时间冲突超过 99 次，跳过。"
            ((failed_count++))
            continue
        fi

        if _common_move_file_no_clobber "$file_path" "$new_file_name"; then
            msg_success "文件 \"$file_path\" 已重命名为 \"$new_file_name\""
        else
            ((failed_count++))
        fi
    done

    (( failed_count == 0 ))
}

# 函数: mtc
mtc() {
    emulate -L zsh -o typeset_silent

    if [[ "${1:-}" == "-h" ]]; then
        cat <<'EOF'
用法:
  mtc file_or_directory1 [file_or_directory2 ...]

功能:
  校验媒体时间一致性，并校验文件名时间与元数据时间是否一致。
  不一致时（且文件名符合规范）调用 ete 修复。

行为说明:
  目录输入时递归扫描 heic 与 mov。
  heic 校验: CreateDate / ModifyDate / DateTimeOriginal。
  mov 校验: CreateDate / ModifyDate / CreationDate / Media* / Track*（QuickTimeUTC=1）。
  文件名格式要求: YYYYMMDD_HHMMSS[NN]（NN 为 01-99）。
EOF
        return 0
    fi

    if [[ $# -lt 1 ]]; then
        msg_error "用法: mtc file_or_directory1 [file_or_directory2 ...]"
        return 1
    fi

    if ! _common_command_exists exiftool; then
        msg_error "未检测到 exiftool。请先执行: brew install exiftool"
        return 1
    fi

    local -i failed_count=0

    for file_path; do
        local -a files_heic=()
        local -a files_mov=()

        if [[ -d "$file_path" ]]; then
            if ! _common_collect_directory_files "$file_path" recursive heic mov; then
                msg_error "扫描目录失败: \"$file_path\""
                ((failed_count++))
                continue
            fi

            local -a found_media_files=("${reply[@]}")
            local found_media
            for found_media in "${found_media_files[@]}"; do
                if _common_path_has_extension "$found_media" mov; then
                    files_mov+=("$found_media")
                elif _common_path_has_extension "$found_media" heic; then
                    files_heic+=("$found_media")
                fi
            done
        elif [[ -f "$file_path" ]]; then
            # 如果是单文件，按扩展名分流，避免同一文件被图片/视频逻辑重复处理
            if _common_path_has_extension "$file_path" mov; then
                files_mov=("$file_path")
            elif _common_path_has_extension "$file_path" heic; then
                files_heic=("$file_path")
            else
                msg_info "跳过: mtc 仅校验 HEIC/MOV \"$file_path\""
                continue
            fi
        else
            msg_error "\"$file_path\" 不是有效的文件或目录，跳过。"
            ((failed_count++))
            continue
        fi

        msg_progress "扫描结果: \"$file_path\" -> HEIC ${#files_heic[@]} 个, MOV ${#files_mov[@]} 个"

        for file_heic in "${files_heic[@]}"; do
            local -a time_values
            time_values=("${(@f)$(exiftool -m -f -fast -s3 \
                -CreateDate \
                -ModifyDate \
                -DateTimeOriginal \
                "$file_heic")}")

            local create_date="${time_values[1]}"
            local modify_date="${time_values[2]}"
            local original_date="${time_values[3]}"

            local -i metadata_consistent=0
            if [[ "$create_date" == "$modify_date" && "$create_date" == "$original_date" ]]; then
                metadata_consistent=1
            fi

            if ! _mtc_check_and_fix "$file_heic" "图片" "$metadata_consistent" "$create_date" \
                "CreateDate: $create_date" \
                "ModifyDate: $modify_date" \
                "DateTimeOriginal: $original_date"; then
                ((failed_count++))
            fi
        done

        for file_mov in "${files_mov[@]}"; do
            local -a mov_time_values
            mov_time_values=("${(@f)$(exiftool -m -f -fast -api QuickTimeUTC=1 -s3 \
                -CreateDate \
                -ModifyDate \
                -CreationDate \
                -MediaCreateDate \
                -MediaModifyDate \
                -TrackCreateDate \
                -TrackModifyDate \
                "$file_mov")}")

            local create_date="${mov_time_values[1]}"
            local modify_date="${mov_time_values[2]}"
            local creation_date="${mov_time_values[3]}"
            local media_create_date="${mov_time_values[4]}"
            local media_modify_date="${mov_time_values[5]}"
            local track_create_date="${mov_time_values[6]}"
            local track_modify_date="${mov_time_values[7]}"

            local -i metadata_consistent=0
            if [[ "$create_date" == "$modify_date" && \
                  "$create_date" == "$creation_date" && \
                  "$create_date" == "$media_create_date" && \
                  "$create_date" == "$media_modify_date" && \
                  "$create_date" == "$track_create_date" && \
                  "$create_date" == "$track_modify_date" ]]; then
                metadata_consistent=1
            fi

            if ! _mtc_check_and_fix "$file_mov" "视频" "$metadata_consistent" "$create_date" \
                "CreateDate: $create_date" \
                "ModifyDate: $modify_date" \
                "CreationDate: $creation_date" \
                "MediaCreateDate: $media_create_date" \
                "MediaModifyDate: $media_modify_date" \
                "TrackCreateDate: $track_create_date" \
                "TrackModifyDate: $track_modify_date"; then
                ((failed_count++))
            fi
        done
    done

    (( failed_count == 0 ))
}

# ========== 图片转换 ==========

# 函数: ctw
ctw() {
    emulate -L zsh -o typeset_silent

    if [[ "${1:-}" == "-h" ]]; then
        cat <<'EOF'
用法:
  ctw [-q 0-100] file_or_directory1 [file_or_directory2 ...]

功能:
  将输入图片转换为 webp，输出到原目录同名 .webp 文件。

参数:
  -q 0-100
    手动指定质量；使用 -q 时不限制目标体积。

行为说明:
  转换时会按需等比缩放，确保输出图片短边不超过 3072，长边不超过 4500。
  默认质量 80。
  未手动指定 q 时，目标最大体积 500KB，超限自动降质重试。
  手动指定 q 时，不做体积限制，按指定质量直接输出。
  如果目标 webp 已存在，转换成功后会覆盖并在输出中标明。
  输入目录时，递归处理该目录及所有子目录下常见图片文件。
EOF
        return 0
    fi

    local -i quality=80
    local -i min_quality=10
    local -i quality_step=5
    local -i max_size_kb=500
    local -i max_size_bytes=$((max_size_kb * 1024))
    local -i max_short_side=3072
    local -i max_long_side=4500
    local -i enforce_size_limit=1

    if [[ "${1:-}" == "-q" ]]; then
        if [[ -z "${2:-}" ]]; then
            msg_error "用法: ctw [-q 0-100] file_or_directory1 [file_or_directory2 ...]"
            return 1
        fi
        local quality_value="${2:-}"
        if [[ ! "$quality_value" =~ '^[0-9]{1,3}$' || "$quality_value" -lt 0 || "$quality_value" -gt 100 ]]; then
            msg_error "quality 必须是 0-100 的整数。"
            return 1
        fi
        quality="$quality_value"
        enforce_size_limit=0
        shift 2
    fi

    if [[ $# -eq 0 ]]; then
        msg_error "用法: ctw [-q 0-100] file_or_directory1 [file_or_directory2 ...]"
        return 1
    fi

    if ! _common_command_exists cwebp; then
        msg_error "未检测到 cwebp。请先执行: brew install webp"
        return 1
    fi

    if ! _common_command_exists sips; then
        msg_error "未检测到 sips，无法进行尺寸缩放。"
        return 1
    fi

    _common_collect_input_files recursive "图片类型" jpg jpeg png tif tiff bmp -- "$@"
    local -a sorted_target_files=("${reply[@]}")

    if [[ ${#sorted_target_files[@]} -eq 0 ]]; then
        msg_error "未找到可处理的图片文件。"
        return 1
    fi

    local -i failed_count=0

    for input in "${sorted_target_files[@]}"; do
        if [[ ! -f "$input" ]]; then
            msg_error "文件不存在 \"$input\""
            ((failed_count++))
            continue
        fi

        local dir="${input:h}"
        local filename="${input:t}"
        local name="${input:t:r}"
        local output="$dir/$name.webp"
        local input_display="$filename"
        local output_display="$name.webp"
        local overwrite_note=""
        [[ -e "$output" ]] && overwrite_note=", 覆盖已有文件"

        local tmp_dir=""
        if ! tmp_dir=$(command mktemp -d "$dir/.ctw_tmp.XXXXXX"); then
            msg_error "创建临时目录失败，跳过 \"$input_display\""
            ((failed_count++))
            continue
        fi

        {
        local tmp_output="$tmp_dir/output.webp"
        local -i try_quality=$quality
        local -i converted=0
        local final_size=""
        local source_for_convert="$input"
        local tmp_resized=""
        local resize_note=", 未缩放"

        local -a sips_info_lines
        sips_info_lines=("${(@f)$(sips -g pixelWidth -g pixelHeight "$input" 2>/dev/null)}")
        local pixel_width=""
        local pixel_height=""
        local sips_info_line
        for sips_info_line in "${sips_info_lines[@]}"; do
            case "$sips_info_line" in
                *pixelWidth:*)
                    pixel_width="${sips_info_line##*: }"
                    ;;
                *pixelHeight:*)
                    pixel_height="${sips_info_line##*: }"
                    ;;
            esac
        done

        if [[ "$pixel_width" =~ '^[0-9]+$' && "$pixel_height" =~ '^[0-9]+$' ]]; then
            local -i short_side=$(( pixel_width < pixel_height ? pixel_width : pixel_height ))
            local -i long_side=$(( pixel_width > pixel_height ? pixel_width : pixel_height ))

            local -i resize_by_short_num=1
            local -i resize_by_short_den=1
            if (( short_side > max_short_side )); then
                resize_by_short_num=$max_short_side
                resize_by_short_den=$short_side
            fi

            local -i resize_by_long_num=1
            local -i resize_by_long_den=1
            if (( long_side > max_long_side )); then
                resize_by_long_num=$max_long_side
                resize_by_long_den=$long_side
            fi

            local -i scale_num=$resize_by_short_num
            local -i scale_den=$resize_by_short_den
            if (( resize_by_long_num * scale_den < scale_num * resize_by_long_den )); then
                scale_num=$resize_by_long_num
                scale_den=$resize_by_long_den
            fi

            if (( scale_num < scale_den )); then
                local -i new_width=$(( (pixel_width * scale_num + scale_den / 2) / scale_den ))
                local -i new_height=$(( (pixel_height * scale_num + scale_den / 2) / scale_den ))
                local input_ext="$(_common_path_lower_extension "$input")"
                tmp_resized="$tmp_dir/resized.${input_ext}"

                if sips -z "$new_height" "$new_width" "$input" --out "$tmp_resized" >/dev/null 2>&1; then
                    source_for_convert="$tmp_resized"
                    resize_note=", 缩放: ${pixel_width}x${pixel_height}->${new_width}x${new_height}"
                else
                    msg_error "尺寸缩放失败 \"$input_display\""
                    _common_remove_temp_file "$tmp_resized" >/dev/null
                    _common_remove_temp_file "$tmp_output" >/dev/null
                    ((failed_count++))
                    continue
                fi
            fi
        else
            msg_error "无法读取图片尺寸，跳过 \"$input_display\""
            _common_remove_temp_file "$tmp_output" >/dev/null
            ((failed_count++))
            continue
        fi

        local -i conversion_failed=0

        while (( enforce_size_limit == 0 || try_quality >= min_quality )); do
            if ! cwebp -quiet -q "$try_quality" "$source_for_convert" -o "$tmp_output"; then
                msg_error "转换失败 \"$input_display\""
                _common_remove_temp_file "$tmp_output" >/dev/null
                conversion_failed=1
                ((failed_count++))
                break
            fi

            final_size=$(stat -f%z "$tmp_output" 2>/dev/null)
            if (( enforce_size_limit == 0 )); then
                if _common_move_file "$tmp_output" "$output"; then
                    converted=1
                    if [[ -n "$final_size" ]]; then
                        msg_success "完成: \"$input_display\" -> \"$output_display\" (q=$try_quality, size=$((final_size/1024))KB${resize_note}, 手动q不限制体积${overwrite_note})"
                    else
                        msg_success "完成: \"$input_display\" -> \"$output_display\" (q=$try_quality${resize_note}, 手动q不限制体积${overwrite_note})"
                    fi
                else
                    _common_remove_temp_file "$tmp_output" >/dev/null
                    conversion_failed=1
                    ((failed_count++))
                fi
                break
            fi

            if [[ -n "$final_size" ]] && (( final_size <= max_size_bytes )); then
                if _common_move_file "$tmp_output" "$output"; then
                    converted=1
                    msg_success "完成: \"$input_display\" -> \"$output_display\" (q=$try_quality, size=$((final_size/1024))KB${resize_note}${overwrite_note})"
                else
                    _common_remove_temp_file "$tmp_output" >/dev/null
                    conversion_failed=1
                    ((failed_count++))
                fi
                break
            fi

            if (( try_quality == min_quality )); then
                break
            fi

            ((try_quality-=quality_step))
        done

        if (( converted == 0 && conversion_failed == 0 )); then
            if [[ -f "$tmp_output" ]]; then
                if _common_move_file "$tmp_output" "$output"; then
                    converted=1
                    final_size=$(stat -f%z "$output" 2>/dev/null)
                    if [[ -n "$final_size" ]]; then
                        msg_warn "\"$input_display\" 在 q=${min_quality} 时仍大于 ${max_size_kb}KB，已输出 \"$output_display\" (size=$((final_size/1024))KB${resize_note}${overwrite_note})"
                    else
                        msg_warn "\"$input_display\" 在 q=${min_quality} 时仍大于 ${max_size_kb}KB，已输出 \"$output_display\"${resize_note}${overwrite_note}"
                    fi
                else
                    _common_remove_temp_file "$tmp_output" >/dev/null
                    conversion_failed=1
                    ((failed_count++))
                fi
            else
                conversion_failed=1
                ((failed_count++))
            fi
        fi

        } always {
            _common_remove_temp_directory "$tmp_dir" >/dev/null
        }
    done

    (( failed_count == 0 ))
}

# ========== 文件归类 ==========

# 函数: cmv
cmv() {
    emulate -L zsh -o typeset_silent

    if [[ "${1:-}" == "-h" ]]; then
        cat <<'EOF'
用法:
  cmv directory

功能:
  按文件名前 8 位日期归类到两级目录 YYYY/MMDD。

行为说明:
  仅处理目标目录当前层级的文件，不递归子目录。
  仅处理文件名符合 YYYYMMDD_HHMMSS[NN].ext 的文件。
  隐藏文件（如 .DS_Store）和无扩展名文件会跳过。
EOF
        return 0
    fi

    if [[ $# -ne 1 ]]; then
        msg_error "用法: cmv directory"
        return 1
    fi

    local target_dir="$1"
    if [[ ! -d "$target_dir" ]]; then
        msg_error "\"$target_dir\" 不是有效目录。"
        return 1
    fi

    local normalized_target_dir="${target_dir:A}"
    local target_leaf="${normalized_target_dir:t}"
    local target_parent="${normalized_target_dir:h:t}"
    if [[ "$target_parent" =~ '^[0-9]{4}$' && "$target_leaf" =~ '^[0-9]{4}$' ]]; then
        msg_warn "跳过: \"$target_dir\" 看起来已经是 YYYY/MMDD 归类目录。请对根目录执行 cmv。"
        return 0
    fi

    local -i moved_count=0
    local -i skipped_count=0

    # 公共扫描函数会用 NUL 分隔读取路径，并返回排序去重后的当前层文件。
    if ! _common_collect_directory_files "$target_dir" current; then
        msg_error "扫描目录失败: \"$target_dir\""
        return 1
    fi

    local -a current_files=("${reply[@]}")
    for file_path in "${current_files[@]}"; do
        local file_name="${file_path:t}"
        local name_part="${file_path:t:r}"
        local file_ext="$(_common_path_lower_extension "$file_path")"

        # 跳过隐藏文件和无扩展名文件，避免把系统文件或异常命名文件纳入归类。
        [[ "$file_name" == .* ]] && continue
        [[ -z "$file_ext" ]] && continue

        # 仅归类 rtf/ete 使用的时间命名格式: YYYYMMDD_HHMMSS[NN]。
        if ! _media_is_timestamp_name "$name_part"; then
            continue
        fi

        local date_prefix="${name_part:0:8}"
        local year_part="${date_prefix:0:4}"
        local md_part="${date_prefix:4:4}"

        local date_dir="$target_dir/$year_part/$md_part"
        if ! command mkdir -p -- "$date_dir"; then
            msg_error "创建目录失败: \"$date_dir\""
            ((skipped_count++))
            continue
        fi

        local dest_file="$date_dir/$file_name"

        if [[ -e "$dest_file" ]]; then
            msg_warn "跳过归类: 目标已存在 \"$year_part/$md_part/$file_name\""
            ((skipped_count++))
            continue
        fi

        if _common_move_file_no_clobber "$file_path" "$dest_file"; then
            ((moved_count++))
            msg_success "已归类: \"$file_name\" -> \"$year_part/$md_part/$file_name\""
        else
            ((skipped_count++))
        fi
    done

    msg_success "归类完成，已处理 $moved_count 个文件，跳过 $skipped_count 个文件。"
}

# 函数: fmv
fmv() {
    emulate -L zsh -o typeset_silent

    if [[ "${1:-}" == "-h" ]]; then
        cat <<'EOF'
用法:
  fmv directory

功能:
  将子目录中的指定媒体文件提取到根目录，再执行 rtf、cmv 与 mtc。

行为说明:
  仅提取以下扩展名: heic/jpg/jpeg/dng/cr3/mov/mp4。
  仅提取子目录层级文件（mindepth 2），不处理根目录已有文件。
  若根目录存在同名文件，则自动追加两位序号后提取。
  提取后删除空子目录，再执行重命名、归类与时间校验。
EOF
        return 0
    fi

    local -a media_extensions=(heic jpg jpeg dng cr3 mov mp4)

    if [[ $# -ne 1 ]]; then
        msg_error "用法: fmv directory"
        return 1
    fi

    local target_dir="$1"
    if [[ ! -d "$target_dir" ]]; then
        msg_error "\"$target_dir\" 不是有效目录。"
        return 1
    fi

    if ! _common_command_exists exiftool; then
        msg_error "未检测到 exiftool。请先执行: brew install exiftool"
        return 1
    fi

    local -a moved_files=()

    if ! _common_collect_directory_files "$target_dir" children "${media_extensions[@]}"; then
        msg_error "扫描目录失败: \"$target_dir\""
        return 1
    fi

    local -a source_files=("${reply[@]}")
    for src_file in "${source_files[@]}"; do
        local dest_file
        local file_name="${src_file:t}"
        if ! dest_file="$(_common_next_available_file_path "$target_dir" "${file_name:r}" "${file_name:e}")"; then
            msg_error "无法为 \"$src_file\" 生成唯一目标文件名，跳过。"
            continue
        fi

        # 记录移动后的根目录路径，供后续 rtf 精确处理这些文件。
        if _common_move_file_no_clobber "$src_file" "$dest_file"; then
            moved_files+=("$dest_file")
            msg_success "已移动: \"$src_file\" -> \"$dest_file\""
        fi
    done

    # 删除空子目录（不删除目标目录本身）。
    if ! command find "$target_dir" -mindepth 1 -type d -empty -delete; then
        msg_warn "清理空子目录失败: \"$target_dir\""
    fi

    if [[ ${#moved_files[@]} -eq 0 ]]; then
        msg_info "未找到可提取的媒体文件。"
        return 0
    fi

    msg_progress "开始执行 rtf 重命名..."
    if ! rtf "${moved_files[@]}"; then
        msg_warn "rtf 执行失败，已停止后续归类；已移动文件保留在根目录。"
        return 1
    fi

    msg_progress "开始执行 cmv 自动归类..."
    if ! cmv "$target_dir"; then
        msg_warn "cmv 执行失败，请检查根目录中的已移动文件。"
        return 1
    fi

    msg_progress "开始执行 mtc 时间校验..."
    if ! mtc "$target_dir"; then
        msg_warn "mtc 执行失败，请检查已归类文件。"
        return 1
    fi
}

# ========== 媒体内部工具函数 ==========

# 判断文件名主体是否符合 YYYYMMDD_HHMMSS[NN]。
_media_is_timestamp_name() {
    emulate -L zsh -o typeset_silent

    local name_part="$1"
    [[ "$name_part" =~ '^[0-9]{8}_[0-9]{6}((0[1-9])|([1-9][0-9]))?$' ]]
}

# 从标准时间文件名主体中取出 YYYYMMDD_HHMMSS。
_media_timestamp_body() {
    emulate -L zsh -o typeset_silent

    print -r -- "${1:0:15}"
}

# 内部辅助函数：校验文件名与元数据一致性，并在需要时自动修复。
# 参数: file_path label metadata_consistent create_date "field: value" ...
_mtc_check_and_fix() {
    emulate -L zsh -o typeset_silent

    local file_path="$1"
    local label="$2"
    local metadata_consistent="$3"
    local create_date="$4"
    shift 4

    local file_name="${file_path:t}"
    local name_part="${file_path:t:r}"

    local -i filename_valid=0
    local file_name_time=""
    if _media_is_timestamp_name "$name_part"; then
        filename_valid=1
        file_name_time="$(_media_timestamp_body "$name_part")"
    fi

    local create_time_for_name=""
    if [[ "$create_date" != "-" && ${#create_date} -ge 19 ]]; then
        local create_main="${create_date:0:19}"
        create_time_for_name="${create_main//:/}"
        create_time_for_name="${create_time_for_name/ /_}"
    fi

    local -i filename_consistent=0
    if [[ $filename_valid -eq 1 && -n "$create_time_for_name" && "$file_name_time" == "$create_time_for_name" ]]; then
        filename_consistent=1
    fi

    if [[ $metadata_consistent -eq 1 && $filename_consistent -eq 1 ]]; then
        msg_success "${label} \"$file_path\" 的时间数据与文件名一致"
    else
        msg_warn "${label} \"$file_path\" 的一致性检查未通过："
        for field_line; do
            print -r -- "$field_line"
        done
        if [[ $filename_valid -eq 1 ]]; then
            print -r -- "FileNameTime: $file_name_time"
            print -r -- "CreateDateAsName: $create_time_for_name"
        else
            print -r -- "FileNameTime: 文件名不符合 YYYYMMDD_HHMMSS[NN] 格式"
        fi
        msg_progress "执行修改操作..."
        if [[ $filename_valid -eq 1 ]]; then
            ete "$file_path"
        else
            msg_warn "跳过自动修复: 文件名非标准，无法按文件名写入时间"
        fi
    fi
}
