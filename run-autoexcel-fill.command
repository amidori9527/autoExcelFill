#!/bin/zsh
set -e

SCRIPT_DIR="${0:A:h}"
BIN="$SCRIPT_DIR/dist/autoexcel-fill/autoexcel-fill"

cd "$SCRIPT_DIR"
"$BIN"

echo
echo "处理完成。按回车关闭窗口。"
read
