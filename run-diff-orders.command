#!/bin/zsh
set -e

SCRIPT_DIR="${0:A:h}"
if [[ -x "$SCRIPT_DIR/AutoExcelKit/diff-orders" ]]; then
  BIN="$SCRIPT_DIR/AutoExcelKit/diff-orders"
elif [[ -x "$SCRIPT_DIR/autoexcel/diff-orders" ]]; then
  BIN="$SCRIPT_DIR/autoexcel/diff-orders"
else
  BIN="$SCRIPT_DIR/dist/AutoExcelKit/diff-orders"
fi

cd "$SCRIPT_DIR"
"$BIN"

echo
echo "处理完成。按回车关闭窗口。"
read
