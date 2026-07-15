#!/bin/zsh
set -e

SCRIPT_DIR="${0:A:h}"
if [[ -x "$SCRIPT_DIR/AutoExcelKit/add-b2b" ]]; then
  BIN="$SCRIPT_DIR/AutoExcelKit/add-b2b"
elif [[ -x "$SCRIPT_DIR/dist/AutoExcelKit/add-b2b" ]]; then
  BIN="$SCRIPT_DIR/dist/AutoExcelKit/add-b2b"
elif [[ -x "$SCRIPT_DIR/.venv/bin/python" && -f "$SCRIPT_DIR/src/autoexcel/add_b2b.py" ]]; then
  cd "$SCRIPT_DIR"
  env PYTHONPATH="$SCRIPT_DIR/src" "$SCRIPT_DIR/.venv/bin/python" -m autoexcel.add_b2b

  echo
  echo "处理完成。按回车关闭窗口。"
  read
  exit 0
else
  echo "找不到 add-b2b。请确认 AutoExcelKit 文件夹完整。"
  read "?按回车关闭窗口。"
  exit 1
fi

cd "$SCRIPT_DIR"
"$BIN"

echo
echo "处理完成。按回车关闭窗口。"
read
