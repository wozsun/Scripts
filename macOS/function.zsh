# 定义 zsh 函数

# 函数: hp
# 功能:
# - 输出所有可用函数及用途说明（命令帮助总览）。
hp() {
    cat <<'EOF'

可用函数列表：

- ete : 根据文件名（YYYYMMDD_HHMMSS）写入/修复媒体时间标签（dng/heic/mov）
- rtf : 读取拍摄时间并重命名文件为 YYYYMMDD_HHMMSS[NN].ext
- etc : 同时校验媒体时间标签与文件名时间一致性，不一致时使用文件名修改（dng/heic/mov）
- ctw : 将图片转换为 webp（默认自动压到 500KB 内；手动 -q 时不限制体积）
- cmv : 按文件名日期（YYYYMMDD）归类到对应子目录
- fmv : 将子目录图片/视频提取到根目录后，依次执行 rtf 与 cmv

EOF
}

# 函数: ete
# 功能:
# - 根据文件名中的时间（YYYYMMDD_HHMMSS[NN]）写入/修复媒体元数据时间标签。
# - 支持 dng / heic / mov。
# 参数:
# - file1 [file2 ...]：一个或多个文件路径。
# 行为说明:
# - dng/heic: 写入 CreateDate / ModifyDate / DateTimeOriginal，并写入 OffsetTime* 时区标签。
# - mov: 以 QuickTime UTC 语义写入 CreateDate / ModifyDate / CreationDate / Media* / Track*。
ete() {
    # 检查参数数量
    if [[ $# -lt 1 ]]; then
        echo "用法: ete file1 [file2 ...]" >&2
        return 1
    fi

    local time_zone="+08:00"

    # 遍历所有输入的文件
    for file_path in "$@"; do
        # 检查文件是否存在
        if [[ ! -f "$file_path" ]]; then
            echo "错误: 文件 \"$file_path\" 不存在，跳过。" >&2
            continue
        fi

        # 提取文件名和后缀
        local file_name="${file_path:t}"
        local file_ext="${file_name##*.}"
        local file_base="${file_name%.*}"
        file_ext="${file_ext:l}"  # zsh 的小写转换
        file_base="${file_base:l}"

        # ete 允许文件名带 rtf 追加序号（如 YYYYMMDD_HHMMSS01），解析时仅取前 15 位时间主体
        if [[ ! "$file_base" =~ '^[0-9]{8}_[0-9]{6}((0[1-9])|([1-9][0-9]))?$' ]]; then
            echo "错误: 文件名 \"$file_name\" 不符合 YYYYMMDD_HHMMSS[NN]，跳过。" >&2
            continue
        fi

        local base_time="${file_base:0:15}"
        local formatted_time=$(date -j -f "%Y%m%d_%H%M%S" "$base_time" +"%Y:%m:%d %H:%M:%S" 2>/dev/null)
        if [[ -z "$formatted_time" ]]; then
            echo "错误: 无法从文件名 \"$file_name\" 解析时间，跳过。" >&2
            continue
        fi

        # 根据扩展名判断如何修改
        if [[ "$file_ext" == "dng" || "$file_ext" == "heic" ]]; then
            local label="HEIF"
            [[ "$file_ext" == "dng" ]] && label="DNG"
            echo "检测到 ${label} 文件，正在修改对应时间标签..."
            exiftool \
                -CreateDate="$formatted_time" \
                -ModifyDate="$formatted_time" \
                -DateTimeOriginal="$formatted_time" \
                -OffsetTime="$time_zone" \
                -OffsetTimeOriginal="$time_zone" \
                -OffsetTimeDigitized="$time_zone" \
                -overwrite_original \
                "$file_path"
            echo "修改结果如下："
            exiftool \
                -m \
                -fast \
                -OffsetTime \
                -CreateDate \
                -ModifyDate \
                -DateTimeOriginal \
                "$file_path"
        elif [[ "$file_ext" == "mov" ]]; then
            local adjusted_time="${formatted_time}${time_zone}"
            echo "检测到 MOV 文件，正在修改对应时间标签..."
            exiftool \
                -api QuickTimeUTC=1 \
                -overwrite_original \
                -CreateDate="$adjusted_time" \
                -ModifyDate="$adjusted_time" \
                -CreationDate="$adjusted_time" \
                -MediaCreateDate="$adjusted_time" \
                -MediaModifyDate="$adjusted_time" \
                -TrackCreateDate="$adjusted_time" \
                -TrackModifyDate="$adjusted_time" \
                "$file_path"
            echo "修改结果如下："
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
        else
            echo "错误: 不支持的文件类型 \"$file_ext\"" >&2
        fi
    done
}

# 函数: rtf
# 功能:
# - 读取文件拍摄时间并重命名为 YYYYMMDD_HHMMSS[NN].ext。
# 参数:
# - [--dry-run|-n]：仅预览重命名结果，不执行 mv。
# - file_or_directory1 [file_or_directory2 ...]：文件或目录。
# 行为说明:
# - 输入目录时，仅处理该目录当前层文件（不递归）。
# - 时间读取优先级: DateTimeOriginal > CreationDate > CreateDate。
# - mov 读取时启用 QuickTimeUTC=1。
# - 文件名冲突时自动追加两位序号（01, 02...）。
rtf() {
    local dry_run=0
    local max_suffix=99  # 同一秒最大冲突序号（01-99）

    # 可选参数：--dry-run / -n（仅预览，不执行重命名）
    if [[ "$1" == "--dry-run" || "$1" == "-n" ]]; then
        dry_run=1
        shift
    fi

    # 检查参数数量
    if [[ $# -lt 1 ]]; then
        echo "用法: rtf [--dry-run|-n] file_or_directory1 [file_or_directory2 ...]" >&2
        return 1
    fi

    local target_files=()
    local input_path

    # 支持输入文件夹：默认处理文件夹当前层的全部文件
    for input_path in "$@"; do
        if [[ -f "$input_path" ]]; then
            target_files+=("$input_path")
        elif [[ -d "$input_path" ]]; then
            while IFS= read -r -d '' found_file; do
                target_files+=("$found_file")
            done < <(find "$input_path" -mindepth 1 -maxdepth 1 -type f -print0)
        else
            echo "错误: \"$input_path\" 不是有效的文件或目录，跳过。" >&2
        fi
    done

    if [[ ${#target_files[@]} -eq 0 ]]; then
        echo "错误: 未找到可处理的文件。" >&2
        return 1
    fi

    # 稳定排序后再处理：基于固定4位编号（0001-9999）按字典序即可等价自然序
    local -a sorted_target_files
    sorted_target_files=("${(@on)target_files}")

    # 遍历所有待处理文件
    for file_path in "${sorted_target_files[@]}"; do

        # 获取文件名与扩展名（用于读取策略与后续命名）
        local file_name="${file_path:t}"
        local file_ext="${file_name##*.}"
        file_ext="${file_ext:l}"  # zsh 的小写转换

        # 使用 exiftool 提取时间（单次调用，按优先级取第一个可用值）
        local -a capture_times
        if [[ "$file_ext" == "mov" ]]; then
            # MOV 按 QuickTime UTC 语义读取
            capture_times=("${(@f)$(exiftool -m -f -fast -api QuickTimeUTC=1 -s3 \
                -DateTimeOriginal \
                -CreationDate \
                -CreateDate \
                "$file_path")}")
        else
            capture_times=("${(@f)$(exiftool -m -f -fast -s3 \
                -DateTimeOriginal \
                -CreationDate \
                -CreateDate \
                "$file_path")}")
        fi

        local original_time=""
        local candidate_time
        for candidate_time in "${capture_times[@]}"; do
            if [[ -n "$candidate_time" && "$candidate_time" != "-" ]]; then
                original_time="$candidate_time"
                break
            fi
        done

        if [[ -z "$original_time" ]]; then
            echo "错误: 无法读取 \"$file_path\" 的拍摄时间，跳过。" >&2
            continue
        fi

        # 仅用于命名：统一截取前 19 位 YYYY:MM:DD HH:MM:SS（忽略后续时区）
        local name_time="${original_time:0:19}"

        # 将 Date 转换为 yyyymmdd_hhmmss 格式
        local formatted_time=$(date -j -f "%Y:%m:%d %H:%M:%S" "$name_time" +"%Y%m%d_%H%M%S")
        if [[ -z "$formatted_time" ]]; then
            echo "错误: 时间格式转换失败，跳过 \"$file_path\"。" >&2
            continue
        fi

        local dir_path="${file_path:h}"  # 文件所在目录
        local base_name="$formatted_time" # 基础文件名
        local new_file_name=""
        local counter=0

        # 循环生成唯一文件名
        while :; do
            if [[ $counter == 0 ]]; then
                new_file_name="${dir_path}/${base_name}.${file_ext}"
            else
                local suffix=$(printf "%02d" "$counter")  # 生成两位数编号
                new_file_name="${dir_path}/${base_name}${suffix}.${file_ext}"
            fi

            if [[ ! -f "$new_file_name" ]]; then
                break
            fi

            if (( counter >= max_suffix )); then
                echo "错误: \"$file_path\" 的时间冲突超过 ${max_suffix} 次，跳过。" >&2
                new_file_name=""
                break
            fi

            ((counter++))
        done

        # 冲突溢出时跳过当前文件
        [[ -z "$new_file_name" ]] && continue

        # 重命名文件（支持仅预览）
        if [[ $dry_run -eq 1 ]]; then
            echo "预览: \"$file_path\" -> \"$new_file_name\""
        else
            mv "$file_path" "$new_file_name"
            echo "文件 \"$file_path\" 已重命名为 \"$new_file_name\""
        fi
    done
}

# 内部辅助函数：校验文件名与元数据一致性，并在需要时自动修复
# 参数: file_path label metadata_consistent create_date "field: value" ...
_etc_check_and_fix() {
    local file_path="$1"
    local label="$2"
    local metadata_consistent="$3"
    local create_date="$4"
    shift 4

    local file_name="${file_path:t}"
    local name_part="${file_name%.*}"

    local filename_valid=0
    local file_name_time=""
    if [[ "$name_part" =~ '^[0-9]{8}_[0-9]{6}((0[1-9])|([1-9][0-9]))?$' ]]; then
        filename_valid=1
        file_name_time="${name_part:0:15}"
    fi

    local create_time_for_name=""
    if [[ "$create_date" != "-" && ${#create_date} -ge 19 ]]; then
        local create_main="${create_date:0:19}"
        create_time_for_name="${create_main//:/}"
        create_time_for_name="${create_time_for_name/ /_}"
    fi

    local filename_consistent=0
    if [[ $filename_valid -eq 1 && -n "$create_time_for_name" && "$file_name_time" == "$create_time_for_name" ]]; then
        filename_consistent=1
    fi

    if [[ $metadata_consistent -eq 1 && $filename_consistent -eq 1 ]]; then
        echo "${label} \"$file_path\" 的时间数据与文件名一致"
    else
        echo "${label} \"$file_path\" 的一致性检查未通过："
        for field_line in "$@"; do
            echo "$field_line"
        done
        if [[ $filename_valid -eq 1 ]]; then
            echo "FileNameTime: $file_name_time"
            echo "CreateDateAsName: $create_time_for_name"
        else
            echo "FileNameTime: 文件名不符合 YYYYMMDD_HHMMSS[NN] 格式"
        fi
        echo "执行修改操作..."
        if [[ $filename_valid -eq 1 ]]; then
            ete "$file_path"
        else
            echo "跳过自动修复: 文件名非标准，无法按文件名写入时间"
        fi
    fi
}

# 函数: etc
# 功能:
# - 校验媒体时间一致性，并校验文件名时间与元数据时间是否一致。
# - 不一致时（且文件名符合规范）调用 ete 修复。
# 参数:
# - file_or_directory [more ...]：文件或目录。
# 行为说明:
# - 目录输入时并发扫描 dng/heic 与 mov。
# - dng/heic 校验: CreateDate / ModifyDate / DateTimeOriginal。
# - mov 校验: CreateDate / ModifyDate / CreationDate / Media* / Track*（QuickTimeUTC=1）。
# - 文件名格式要求: YYYYMMDD_HHMMSS[NN]（NN 为 01-99）。
etc() {
    # 检查参数数量
    if [[ $# -lt 1 ]]; then
        echo "用法: etc file_or_directory" >&2
        return 1
    fi

    # 遍历所有输入的文件或目录
    for file_path in "$@"; do
        local files_image=()
        local files_mov=()

        if [[ -d "$file_path" ]]; then
            # 如果是目录：并发扫描图片与视频，减少大目录等待时间
            # 关闭当前函数作用域内的 job 通知，避免输出 [n] done 提示
            setopt localoptions
            unsetopt monitor

            local image_tmp mov_tmp
            image_tmp=$(mktemp)
            mov_tmp=$(mktemp)

            find "$file_path" -type f \( -iname "*.heic" -o -iname "*.dng" \) -print0 > "$image_tmp" &
            local find_image_pid=$!

            find "$file_path" -type f -iname "*.mov" -print0 > "$mov_tmp" &
            local find_mov_pid=$!

            wait "$find_image_pid"
            wait "$find_mov_pid"

            while IFS= read -r -d '' found_image; do
                files_image+=("$found_image")
            done < "$image_tmp"

            while IFS= read -r -d '' found_mov; do
                files_mov+=("$found_mov")
            done < "$mov_tmp"

            rm -f "$image_tmp" "$mov_tmp"
        elif [[ -f "$file_path" ]]; then
            # 如果是单文件，按扩展名分流，避免同一文件被图片/视频逻辑重复处理
            local file_name="${file_path:t}"
            local file_ext="${file_name##*.}"
            file_ext="${file_ext:l}"

            if [[ "$file_ext" == "mov" ]]; then
                files_mov=("$file_path")
            elif [[ "$file_ext" == "heic" || "$file_ext" == "dng" ]]; then
                files_image=("$file_path")
            else
                echo "错误: 不支持的文件类型 \"$file_path\"，跳过。" >&2
                continue
            fi
        else
            echo "错误: \"$file_path\" 不是有效的文件或目录，跳过。" >&2
            continue
        fi

        echo "扫描结果: \"$file_path\" -> HEIC/DNG ${#files_image[@]} 个, MOV ${#files_mov[@]} 个"

        # 遍历图片文件（HEIC/DNG）
        for file_image in "${files_image[@]}"; do
            local -a time_values
            time_values=("${(@f)$(exiftool -m -f -fast -s3 \
                -CreateDate \
                -ModifyDate \
                -DateTimeOriginal \
                "$file_image")}")

            local create_date="${time_values[1]}"
            local modify_date="${time_values[2]}"
            local original_date="${time_values[3]}"

            local metadata_consistent=0
            if [[ "$create_date" == "$modify_date" && "$create_date" == "$original_date" ]]; then
                metadata_consistent=1
            fi

            _etc_check_and_fix "$file_image" "图片" "$metadata_consistent" "$create_date" \
                "CreateDate: $create_date" \
                "ModifyDate: $modify_date" \
                "DateTimeOriginal: $original_date"
        done

        # 遍历视频文件（MOV）
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

            local metadata_consistent=0
            if [[ "$create_date" == "$modify_date" && \
                  "$create_date" == "$creation_date" && \
                  "$create_date" == "$media_create_date" && \
                  "$create_date" == "$media_modify_date" && \
                  "$create_date" == "$track_create_date" && \
                  "$create_date" == "$track_modify_date" ]]; then
                metadata_consistent=1
            fi

            _etc_check_and_fix "$file_mov" "视频" "$metadata_consistent" "$create_date" \
                "CreateDate: $create_date" \
                "ModifyDate: $modify_date" \
                "CreationDate: $creation_date" \
                "MediaCreateDate: $media_create_date" \
                "MediaModifyDate: $media_modify_date" \
                "TrackCreateDate: $track_create_date" \
                "TrackModifyDate: $track_modify_date"
        done
    done
}

# 函数: ctw
# 功能:
# - 将输入图片转换为 webp，输出到原目录同名 .webp 文件。
# 参数:
# - [-q|--quality 0-100] path1 path2 ...：一个或多个图片文件或目录路径。
# 行为说明:
# - 转换时会按需等比缩放，确保输出图片短边不超过 3072，长边不超过 4500。
# - 默认质量 80；可通过 -q/--quality 手动指定质量。
# - 未手动指定 q 时，目标最大体积 500KB，超限自动降质重试。
# - 手动指定 q 时，不做体积限制，按指定质量直接输出。
# - 输入目录时，递归处理该目录及所有子目录下常见图片文件。
ctw() {
    local quality=80
    local min_quality=10
    local quality_step=5    # 每次降质步长
    local max_size_kb=500
    local max_size_bytes=$((max_size_kb * 1024))
    local max_short_side=3072
    local max_long_side=4500
    local enforce_size_limit=1

    if [[ "$1" == "-q" || "$1" == "--quality" ]]; then
        if [[ -z "$2" ]]; then
            echo "用法: ctw [-q|--quality 0-100] path1 [path2 ...]" >&2
            return 1
        fi
        if [[ ! "$2" =~ '^[0-9]{1,3}$' || "$2" -lt 0 || "$2" -gt 100 ]]; then
            echo "错误: quality 必须是 0-100 的整数。" >&2
            return 1
        fi
        quality="$2"
        enforce_size_limit=0
        shift 2
    fi

    if [[ $# -eq 0 ]]; then
        echo "用法: ctw [-q|--quality 0-100] path1 [path2 ...]" >&2
        return 1
    fi

    if ! command -v cwebp >/dev/null 2>&1; then
        echo "错误: 未检测到 cwebp。请先执行: brew install webp" >&2
        return 1
    fi

    if ! command -v sips >/dev/null 2>&1; then
        echo "错误: 未检测到 sips，无法进行尺寸缩放。" >&2
        return 1
    fi

    local target_files=()
    local input_path

    # 支持输入文件或目录；目录会递归收集常见图片文件
    for input_path in "$@"; do
        if [[ -f "$input_path" ]]; then
            target_files+=("$input_path")
        elif [[ -d "$input_path" ]]; then
            while IFS= read -r -d '' found_file; do
                target_files+=("$found_file")
            done < <(find "$input_path" -type f \
                \( -iname "*.jpg" -o -iname "*.jpeg" -o -iname "*.png" -o -iname "*.tif" -o -iname "*.tiff" -o -iname "*.bmp" \) \
                -print0)
        else
            echo "错误: 路径无效 \"$input_path\"" >&2
        fi
    done

    if [[ ${#target_files[@]} -eq 0 ]]; then
        echo "错误: 未找到可处理的图片文件。" >&2
        return 1
    fi

    local -a sorted_target_files
    sorted_target_files=("${(@on)target_files}")

    for input in "${sorted_target_files[@]}"; do
        if [[ ! -f "$input" ]]; then
            echo "错误: 文件不存在 \"$input\"" >&2
            continue
        fi

        local dir="${input:h}"
        local filename="${input:t}"
        local name="${filename%.*}"
        local output="$dir/$name.webp"
        local input_display="$filename"
        local output_display="$name.webp"

        local tmp_output="$output.tmp.$$"
        local try_quality=$quality
        local converted=0
        local final_size=0
        local source_for_convert="$input"
        local tmp_resized=""
        local resize_note=", 未缩放"

        local sips_info=$(sips -g pixelWidth -g pixelHeight "$input" 2>/dev/null)
        local pixel_width=$(awk '/pixelWidth:/{print $2; exit}' <<< "$sips_info")
        local pixel_height=$(awk '/pixelHeight:/{print $2; exit}' <<< "$sips_info")

        if [[ "$pixel_width" =~ '^[0-9]+$' && "$pixel_height" =~ '^[0-9]+$' ]]; then
            local short_side=$pixel_width
            local long_side=$pixel_width
            if (( pixel_height < short_side )); then
                short_side=$pixel_height
            fi
            if (( pixel_height > long_side )); then
                long_side=$pixel_height
            fi

            local resize_by_short_num=1
            local resize_by_short_den=1
            if (( short_side > max_short_side )); then
                resize_by_short_num=$max_short_side
                resize_by_short_den=$short_side
            fi

            local resize_by_long_num=1
            local resize_by_long_den=1
            if (( long_side > max_long_side )); then
                resize_by_long_num=$max_long_side
                resize_by_long_den=$long_side
            fi

            local scale_num=$resize_by_short_num
            local scale_den=$resize_by_short_den
            if (( resize_by_long_num * scale_den < scale_num * resize_by_long_den )); then
                scale_num=$resize_by_long_num
                scale_den=$resize_by_long_den
            fi

            if (( scale_num < scale_den )); then
                local new_width=$(( (pixel_width * scale_num + scale_den / 2) / scale_den ))
                local new_height=$(( (pixel_height * scale_num + scale_den / 2) / scale_den ))
                local input_ext="${filename##*.}"
                tmp_resized="$dir/.${name}.ctw_resize.$$.${input_ext}"

                if sips -z "$new_height" "$new_width" "$input" --out "$tmp_resized" >/dev/null 2>&1; then
                    source_for_convert="$tmp_resized"
                    resize_note=", 缩放: ${pixel_width}x${pixel_height}->${new_width}x${new_height}"
                else
                    echo "错误: 尺寸缩放失败 \"$input_display\"" >&2
                    rm -f "$tmp_resized" "$tmp_output"
                    continue
                fi
            fi
        else
            echo "错误: 无法读取图片尺寸，跳过 \"$input_display\"" >&2
            rm -f "$tmp_output"
            continue
        fi

        while (( try_quality >= min_quality )); do
            cwebp -quiet -q "$try_quality" "$source_for_convert" -o "$tmp_output"
            if [[ $? -ne 0 ]]; then
                echo "错误: 转换失败 \"$input_display\"" >&2
                rm -f "$tmp_output"
                break
            fi

            final_size=$(stat -f%z "$tmp_output" 2>/dev/null)
            if (( enforce_size_limit == 0 )); then
                mv "$tmp_output" "$output"
                converted=1
                if [[ -n "$final_size" ]]; then
                    echo "完成: \"$input_display\" -> \"$output_display\" (q=$try_quality, size=$((final_size/1024))KB${resize_note}, 手动q不限制体积)"
                else
                    echo "完成: \"$input_display\" -> \"$output_display\" (q=$try_quality${resize_note}, 手动q不限制体积)"
                fi
                break
            fi

            if [[ -n "$final_size" ]] && (( final_size <= max_size_bytes )); then
                mv "$tmp_output" "$output"
                converted=1
                echo "完成: \"$input_display\" -> \"$output_display\" (q=$try_quality, size=$((final_size/1024))KB${resize_note})"
                break
            fi

            if (( try_quality == min_quality )); then
                break
            fi

            ((try_quality-=quality_step))
        done

        if (( converted == 0 )); then
            if [[ -f "$tmp_output" ]]; then
                mv "$tmp_output" "$output"
                final_size=$(stat -f%z "$output" 2>/dev/null)
                if [[ -n "$final_size" ]]; then
                    echo "警告: \"$input_display\" 在 q=${min_quality} 时仍大于 ${max_size_kb}KB，已输出 \"$output_display\" (size=$((final_size/1024))KB${resize_note})"
                else
                    echo "警告: \"$input_display\" 在 q=${min_quality} 时仍大于 ${max_size_kb}KB，已输出 \"$output_display\"${resize_note}"
                fi
            fi
        fi

        if [[ -n "$tmp_resized" ]]; then
            rm -f "$tmp_resized"
        fi
    done
}

# 函数: cmv
# 功能:
# - 按文件名前 8 位日期归类到两级目录 YYYY/MMDD。
# 参数:
# - directory：目标目录（仅处理其当前层文件，不递归）。
# 行为说明:
# - 仅处理文件名符合 YYYYMMDD_HHMMSS[NN].ext 的文件。
# - 隐藏文件（如 .DS_Store）和无扩展名文件会跳过。
cmv() {
    # 参数要求：仅接受 1 个目录参数
    if [[ $# -ne 1 ]]; then
        echo "用法: cmv directory" >&2
        return 1
    fi

    # 目标目录：只处理该目录“当前层级”的文件，不递归子目录
    local target_dir="$1"
    if [[ ! -d "$target_dir" ]]; then
        echo "错误: \"$target_dir\" 不是有效目录。" >&2
        return 1
    fi

    # 统计成功归类（实际发生 mv）的文件数量
    local moved_count=0

    # 使用 find + -print0，确保文件名中即使含空格/特殊字符也能被安全读取
    while IFS= read -r -d '' file_path; do
        # 拆分文件名：
        # - file_name: 含扩展名的完整文件名
        # - name_part: 不含扩展名（用于格式校验）
        local file_name="${file_path:t}"
        local name_part="${file_name%.*}"

        # 快速过滤：
        # 1) 跳过 macOS 隐藏文件（如 .DS_Store）
        # 2) 跳过无扩展名文件（避免把异常命名文件纳入归类）
        [[ "$file_name" == .* ]] && continue
        [[ "$name_part" == "$file_name" ]] && continue

        # 文件名格式校验（仅校验 name_part，不含扩展名）：
        # - YYYYMMDD_HHMMSS
        # - YYYYMMDD_HHMMSSNN（NN 仅允许 01-99，用于兼容 rtf 重名后缀）
        if [[ ! "$name_part" =~ '^[0-9]{8}_[0-9]{6}((0[1-9])|([1-9][0-9]))?$' ]]; then
            continue
        fi

        # 以文件名前 8 位日期构造两级目录：YYYY/MMDD
        local date_prefix="${name_part:0:8}"
        local year_part="${date_prefix:0:4}"
        local md_part="${date_prefix:4:4}"

        # 不存在则创建日期目录
        local date_dir="$target_dir/$year_part/$md_part"
        mkdir -p "$date_dir"

        # 目标路径：日期目录/原文件名
        # 约定前提：文件已由 rtf 统一重命名，归类阶段默认不会重名
        local dest_file="$date_dir/$file_name"

        # 执行移动并记录统计
        mv "$file_path" "$dest_file"
        ((moved_count++))
        echo "已归类: \"$file_name\" -> \"$year_part/$md_part/$file_name\""
    # 仅扫描 target_dir 当前层文件（-mindepth 1 -maxdepth 1 -type f）
    done < <(find "$target_dir" -mindepth 1 -maxdepth 1 -type f -print0)

    # 输出归类统计（仅统计实际移动成功的文件）
    echo "归类完成，共处理 $moved_count 个文件。"
}

# 函数: fmv
# 功能:
# - 将子目录中的指定媒体文件提取到根目录，再执行 rtf 与 cmv。
# 参数:
# - directory：目标根目录。
# 行为说明:
# - 仅提取以下扩展名: heic/jpg/jpeg/dng/cr3/mov/mp4。
# - 仅提取子目录层级文件（mindepth 2），不处理根目录已有文件。
# - 若根目录存在同名文件则跳过该文件。
# - 提取后删除空子目录，再执行重命名与归类。
# 返回:
# - 参数错误或目录无效返回 1；无可提取文件返回 0。
fmv() {
    # 需要提取的媒体文件扩展名
    local media_extensions=("heic" "jpg" "jpeg" "dng" "cr3" "mov" "mp4")

    if [[ $# -ne 1 ]]; then
        echo "用法: fmv directory" >&2
        return 1
    fi

    local target_dir="$1"
    if [[ ! -d "$target_dir" ]]; then
        echo "错误: \"$target_dir\" 不是有效目录。" >&2
        return 1
    fi

    local moved_files=()

    # 从子目录中逐个读取目标文件路径：
    # - IFS= 与 -r：防止 read 对空白和反斜杠做意外处理
    # - -d ''：以 NUL 字符分隔，配合 find -print0 可安全处理包含空格/换行的文件名
    while IFS= read -r -d '' src_file; do
        # 取出文件名（不含目录）
        local base_name="${src_file:t}"
        # 默认目标路径：移动到根目录并保持原文件名
        local dest_file="$target_dir/$base_name"

        # 若目标已存在则跳过
        if [[ -e "$dest_file" ]]; then
            echo "跳过: 目标已存在 \"$dest_file\"" >&2
            continue
        fi

        # 执行移动，并记录新路径，供后续 rtf 批量重命名
        mv "$src_file" "$dest_file"
        moved_files+=("$dest_file")
        echo "已移动: \"$src_file\" -> \"$dest_file\""
    done < <(
        # 仅搜索 target_dir 的“子目录层级”文件（-mindepth 2）
        # 不处理根目录现有文件，仅提取指定媒体格式到根目录
        local find_args=()
        for _ext in "${media_extensions[@]}"; do
            [[ ${#find_args[@]} -gt 0 ]] && find_args+=("-o")
            find_args+=("-iname" "*.${_ext}")
        done
        find "$target_dir" -mindepth 2 -type f \( "${find_args[@]}" \) -print0
    )

    # 删除空子目录（不删除目标目录本身）
    find "$target_dir" -mindepth 1 -type d -empty -delete

    if [[ ${#moved_files[@]} -eq 0 ]]; then
        echo "未找到可提取的媒体文件。"
        return 0
    fi

    echo "开始执行 rtf 重命名..."
    rtf "${moved_files[@]}"

    echo "开始执行 cmv 自动归类..."
    cmv "$target_dir"
}
