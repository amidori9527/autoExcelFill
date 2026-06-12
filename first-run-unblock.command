#!/bin/zsh
set -e

SCRIPT_DIR="${0:A:h}"

echo "正在解除 macOS 对当前工具文件夹的安全隔离标记..."
xattr -dr com.apple.quarantine "$SCRIPT_DIR" 2>/dev/null || true
chmod +x "$SCRIPT_DIR/run-autoexcel-fill.command" 2>/dev/null || true
chmod +x "$SCRIPT_DIR/autoexcel-fill/autoexcel-fill" 2>/dev/null || true

echo
echo "已完成。现在可以双击 run-autoexcel-fill.command 运行。"
echo "按回车关闭窗口。"
read
