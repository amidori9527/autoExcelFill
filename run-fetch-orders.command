#!/bin/zsh
set -e

SCRIPT_DIR="${0:A:h}"
if [[ -x "$SCRIPT_DIR/AutoExcelKit/fetch-orders" ]]; then
  BIN="$SCRIPT_DIR/AutoExcelKit/fetch-orders"
elif [[ -x "$SCRIPT_DIR/dist/AutoExcelKit/fetch-orders" ]]; then
  BIN="$SCRIPT_DIR/dist/AutoExcelKit/fetch-orders"
else
  echo "找不到 fetch-orders。请确认 AutoExcelKit 文件夹完整。"
  read "?按回车关闭窗口。"
  exit 1
fi

cd "$SCRIPT_DIR"
"$BIN"

echo
echo "处理完成。按回车关闭窗口。"
read
