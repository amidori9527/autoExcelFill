#!/bin/zsh
set -e

SCRIPT_DIR="${0:A:h}"
if [[ -x "$SCRIPT_DIR/autoexcel-fill/autoexcel-fill" ]]; then
  BIN="$SCRIPT_DIR/autoexcel-fill/autoexcel-fill"
else
  BIN="$SCRIPT_DIR/dist/autoexcel-fill/autoexcel-fill"
fi

cd "$SCRIPT_DIR"
"$BIN"

echo
echo "处理完成。按回车关闭窗口。"
read
