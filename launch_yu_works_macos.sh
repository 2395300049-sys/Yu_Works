#!/bin/zsh

set -u

script_dir="$(cd "$(dirname "$0")" && pwd -P)"
python_bin="$script_dir/.venv-macos/bin/python"
log_dir="$HOME/Library/Logs/Yu_Works"
log_file="$log_dir/yu_works.log"

mkdir -p "$log_dir"

if [[ ! -x "$python_bin" ]]; then
  /usr/bin/osascript -e 'display alert "Yu_Works 无法启动" message "macOS 运行环境不存在，请先在 Yu_Works 文件夹中完成安装。" as critical'
  exit 1
fi

cd "$script_dir" || exit 1
export PYTHONUNBUFFERED=1
exec "$python_bin" "$script_dir/gui_main.py" >>"$log_file" 2>&1
